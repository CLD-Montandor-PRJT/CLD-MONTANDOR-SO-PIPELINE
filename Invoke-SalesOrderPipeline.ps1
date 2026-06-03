<#
.SYNOPSIS
    Template-driven sales order pipeline: PDF -> BC (Sandbox-Training or Production).
.DESCRIPTION
    Scans each client Not_Processed folder, extracts order data from PDFs using
    per-client template.json (coordinate-based column extraction via PdfPig),
    then POSTs to BC API v2.0 and OData-PATCHes Your_Reference, Requested_Delivery_Date,
    and Ship_to_Code in a single call.
    Runs in dry-run mode by default. Pass -Execute to write to BC.
.PARAMETER Execute
    Actually POST to BC and move PDFs. Without this flag, only prints what would happen.
.PARAMETER ClientFolder
    Restrict to one client subfolder (e.g. 'RETIF_FITER'). Omit to process all.
#>
param(
    [switch]$Execute,
    [string]$ClientFolder = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$DryRun = -not $Execute
$rootDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$dllDir  = "$rootDir\lib\dlls"

# ---------------------------------------------------------------------------
# Load PdfPig
# ---------------------------------------------------------------------------
Write-Host '[INIT] Loading PdfPig...' -ForegroundColor Cyan
foreach ($dll in @(
    'UglyToad.PdfPig.Core.dll', 'UglyToad.PdfPig.Tokens.dll',
    'UglyToad.PdfPig.Tokenization.dll', 'UglyToad.PdfPig.Fonts.dll',
    'UglyToad.PdfPig.dll'
)) {
    Add-Type -Path "$dllDir\$dll" -ErrorAction SilentlyContinue
}
Write-Host '  PdfPig ready.' -ForegroundColor Green

# ---------------------------------------------------------------------------
# Credentials
# ---------------------------------------------------------------------------
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class WinCredPL {
    [DllImport("advapi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    private static extern bool CredRead(string targetName, uint type, uint flags, out IntPtr credPtr);
    [DllImport("advapi32.dll")]
    private static extern void CredFree(IntPtr credPtr);
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
    private struct CREDENTIAL {
        public uint Flags; public uint Type; public string TargetName;
        public string Comment; public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
        public uint CredentialBlobSize; public IntPtr CredentialBlob;
        public uint Persist; public uint AttributeCount; public IntPtr Attributes;
        public string TargetAlias; public string UserName;
    }
    public static string GetSecret(string target) {
        IntPtr ptr;
        if (!CredRead(target, 1, 0, out ptr)) return null;
        var cred = Marshal.PtrToStructure<CREDENTIAL>(ptr);
        var secret = Marshal.PtrToStringUni(cred.CredentialBlob, (int)cred.CredentialBlobSize / 2);
        CredFree(ptr);
        return secret;
    }
}
"@ -Language CSharp

$tenantId     = [WinCredPL]::GetSecret('Montandor_BC_TenantId')
$clientId     = [WinCredPL]::GetSecret('Montandor_BC_ClientId')
$clientSecret = [WinCredPL]::GetSecret('Montandor_BC_ClientSecret')

# ---------------------------------------------------------------------------
# OAuth token
# ---------------------------------------------------------------------------
Write-Host '[AUTH] Acquiring OAuth tokens...' -ForegroundColor Cyan
$token = Invoke-RestMethod -Method Post `
    -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
    -Body @{
        grant_type    = 'client_credentials'
        client_id     = $clientId
        client_secret = $clientSecret
        scope         = 'https://api.businesscentral.dynamics.com/.default'
    } -ContentType 'application/x-www-form-urlencoded'

$graphToken = Invoke-RestMethod -Method Post `
    -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
    -Body @{
        grant_type    = 'client_credentials'
        client_id     = $clientId
        client_secret = $clientSecret
        scope         = 'https://graph.microsoft.com/.default'
    } -ContentType 'application/x-www-form-urlencoded'

$authHeader  = @{ Authorization = "Bearer $($token.access_token)" }
$jsonHeader  = $authHeader + @{ 'Content-Type' = 'application/json' }
$graphHeader = @{ Authorization = "Bearer $($graphToken.access_token)" }
$companyId   = '4e422ae7-867a-ef11-a671-000d3a45ce6c'
$bcItemCache = @{}
Write-Host "  BC token acquired. Expires in $($token.expires_in)s." -ForegroundColor Green
Write-Host "  Graph token acquired. Expires in $($graphToken.expires_in)s." -ForegroundColor Green

# ---------------------------------------------------------------------------
# Error helper — extracts BC's JSON error message from a failed Invoke-RestMethod
# ---------------------------------------------------------------------------
function Get-ApiError {
    param($Err)
    if ($Err.ErrorDetails.Message) {
        try {
            $body = $Err.ErrorDetails.Message | ConvertFrom-Json -ErrorAction Stop
            if ($body.error.message) { return $body.error.message }
        } catch { }
        return $Err.ErrorDetails.Message   # raw body if not JSON
    }
    return $Err.Exception.Message
}

# ---------------------------------------------------------------------------
# Deduplication — checks open sales orders AND posted invoices
# ---------------------------------------------------------------------------
function Test-BcOrderExists {
    param([string]$CustomerNumber, [string]$OrderRef, [string]$Environment)
    if (-not $OrderRef) { return $false }
    $odata = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$Environment/ODataV4/Company('Montandor_Andorra')"

    # Open + Released orders (Document_Type=Order covers both)
    $f1 = "Sell_to_Customer_No eq '$CustomerNumber' and Your_Reference eq '$OrderRef' and Document_Type eq 'Order'"
    $r1 = Invoke-RestMethod -Uri "$odata/SalesOrder?`$filter=$f1&`$select=No&`$top=1" -Headers $authHeader
    if (@($r1.value).Count -gt 0) { return $true }

    # Posted invoices — only if web service is published in BC
    try {
        $f2 = "Sell_to_Customer_No eq '$CustomerNumber' and Your_Reference eq '$OrderRef'"
        $r2 = Invoke-RestMethod -Uri "$odata/PostedSalesInvoice?`$filter=$f2&`$select=No&`$top=1" -Headers $authHeader
        if (@($r2.value).Count -gt 0) { return $true }
    } catch { }

    return $false
}

# ---------------------------------------------------------------------------
# BC ship-to address lookup (Order Pipeline extension — api/montandor/pipeline/v1.0)
# ---------------------------------------------------------------------------
function Get-BcShipToCode {
    param([string]$CustomerNumber, [string]$PostCode, [string]$Environment)
    $uri = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$Environment/api/montandor/pipeline/v1.0/companies($companyId)/shipToAddresses?`$filter=customerNumber eq '$CustomerNumber' and postCode eq '$PostCode'"
    $resp = Invoke-RestMethod -Uri $uri -Headers $authHeader
    return @($resp.value)
}

# ---------------------------------------------------------------------------
# BC item catalogue (cached per environment, used by text-mode extraction)
# ---------------------------------------------------------------------------
function Get-BcItemNumbers {
    param([string]$Environment)
    if ($bcItemCache.ContainsKey($Environment)) { return $bcItemCache[$Environment] }
    Write-Host "  [ITEMS] Loading BC item list ($Environment)..." -ForegroundColor Cyan
    $uri = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$Environment/api/v2.0/companies($companyId)/items?`$select=number&`$top=5000"
    $resp = Invoke-RestMethod -Uri $uri -Headers $authHeader
    $bcItemCache[$Environment] = @($resp.value | ForEach-Object { $_.number })
    Write-Host "    $($bcItemCache[$Environment].Count) items loaded." -ForegroundColor Green
    return $bcItemCache[$Environment]
}

# ---------------------------------------------------------------------------
# PDF extraction (coordinate-based)
# ---------------------------------------------------------------------------
function Get-PdfOrderData {
    param(
        [string]$PdfPath     = '',
        [byte[]]$PdfBytes    = $null,
        [PSCustomObject]$Template,
        [string[]]$BcItemNumbers = @()
    )

    $pdf = if ($PdfBytes) { [UglyToad.PdfPig.PdfDocument]::Open($PdfBytes) }
           else           { [UglyToad.PdfPig.PdfDocument]::Open($PdfPath)  }
    $allText = ($pdf.GetPages() | ForEach-Object { $_.Text }) -join ' ' -replace '[\x00-\x1F]', ' '

    # Order number — regex on raw text, or coordinate scan if orderRefXRange defined
    $orderRef = ''
    if ($Template.PSObject.Properties['orderNumberRegex'] -and $allText -match $Template.orderNumberRegex) {
        $orderRef = $Matches[1]
    }
    if (-not $orderRef -and $Template.PSObject.Properties['orderRefXRange']) {
        $xOR  = $Template.orderRefXRange
        $yOR  = $Template.orderRefYRange
        $orPat = if ($Template.PSObject.Properties['orderRefPattern']) { $Template.orderRefPattern } else { '^\d{4,6}$' }
        foreach ($w in @($pdf.GetPages())[0].GetWords()) {
            $b = [Math]::Round($w.BoundingBox.Bottom, 1)
            $l = $w.BoundingBox.Left
            if ($b -ge $yOR[0] -and $b -le $yOR[1] -and $l -ge $xOR[0] -and $l -le $xOR[1] -and $w.Text -match $orPat) {
                $orderRef = $w.Text; break
            }
        }
    }

    # Order date (first DD/MM/YYYY match in text = order date, before legal disclaimers)
    $orderDateBC = (Get-Date).ToString('yyyy-MM-dd')
    if ($allText -match $Template.dateRegex) {
        $parsed      = [DateTime]::ParseExact(
            $Matches[1], $Template.dateInputFormat,
            [System.Globalization.CultureInfo]::InvariantCulture)
        $orderDateBC = $parsed.ToString('yyyy-MM-dd')
    }

    # Mapping path — used by both modes
    $mapPath = Join-Path (Split-Path $PdfPath) '..\mapping.json'

    # ---- TEXT mode --------------------------------------------------------
    if ($Template.PSObject.Properties['extractionMode'] -and $Template.extractionMode -eq 'text') {

        # Delivery date
        $deliveryDateBC = ''
        if ($Template.PSObject.Properties['deliveryDateRegex'] -and
            $allText -match $Template.deliveryDateRegex) {
            $parsed = [DateTime]::ParseExact(
                $Matches[1], $Template.dateInputFormat,
                [System.Globalization.CultureInfo]::InvariantCulture)
            $deliveryDateBC = $parsed.ToString('yyyy-MM-dd')
        }

        # Item lines: BC catalogue scan + optional rename map
        $searchList = [System.Collections.Generic.Dictionary[string,string]]::new()
        foreach ($itemNo in $BcItemNumbers) {
            if (-not $searchList.ContainsKey($itemNo)) { $searchList[$itemNo] = $itemNo }
        }
        if (Test-Path $mapPath) {
            $renames = Get-Content $mapPath | ConvertFrom-Json
            foreach ($prop in $renames.PSObject.Properties) {
                $searchList[$prop.Name] = $prop.Value
                if ($prop.Name -ne $prop.Value -and $searchList.ContainsKey($prop.Value)) {
                    $searchList.Remove($prop.Value)
                }
            }
        }
        $lines = @()
        foreach ($searchKey in $searchList.Keys) {
            $pattern = '(?<![A-Z0-9\-])' + [regex]::Escape($searchKey) + '\s*(\d+,\d{2})'
            if ($allText -match $pattern) {
                $lines += [PSCustomObject]@{
                    ItemNumber = $searchList[$searchKey]
                    Quantity   = [double]($Matches[1] -replace ',', '.')
                }
            }
        }

        # Delivery address (template regex fields)
        $shipTo = $null
        if ($Template.PSObject.Properties['deliveryNameRegex'] -and $allText -match $Template.deliveryNameRegex) {
            $addrName     = $Matches[1].Trim()
            $addrLine1    = ''
            $addrPostCode = ''
            $addrCity     = ''
            $addrCountry  = if ($Template.PSObject.Properties['deliveryCountry']) { $Template.deliveryCountry } else { '' }
            if ($Template.PSObject.Properties['deliveryAddress1Regex'] -and $allText -match $Template.deliveryAddress1Regex) {
                $addrLine1 = $Matches[1].Trim()
            }
            if ($Template.PSObject.Properties['deliveryPostCodeRegex'] -and $allText -match $Template.deliveryPostCodeRegex) {
                $addrPostCode = $Matches[1].Trim()
            }
            if ($Template.PSObject.Properties['deliveryCityRegex'] -and $allText -match $Template.deliveryCityRegex) {
                $addrCity = ($Matches[1] -replace '\s+', ' ').Trim()
            }
            if ($addrName) {
                $shipTo = [PSCustomObject]@{
                    name         = $addrName
                    addressLine1 = $addrLine1
                    addressLine2 = ''
                    city         = $addrCity
                    postCode     = $addrPostCode
                    country      = $addrCountry
                }
            }
        }

        # Extract postcode for BC ship-to lookup
        $extractedPostCode = ''
        if ($Template.PSObject.Properties['deliveryPostCodeRegex'] -and $allText -match $Template.deliveryPostCodeRegex) {
            $extractedPostCode = $Matches[1].Trim()
        }

        $pdf.Dispose()
        return [PSCustomObject]@{
            OrderRef       = $orderRef
            OrderDate      = $orderDateBC
            DeliveryDate   = $deliveryDateBC
            ShipTo         = $shipTo
            ShipToPostCode = $extractedPostCode
            Lines          = $lines
        }
    }
    # -----------------------------------------------------------------------

    # Requested delivery date extracted by word Y-coordinate (text ordering is unreliable)
    # Template defines deliveryDateYRange: the Bottom range where the delivery date word sits
    $deliveryDateBC  = ''
    $hasDelivDateCfg = $Template.PSObject.Properties['deliveryDateYRange']

    # Column and address config
    $xRef    = $Template.yourRefColumnXRange
    $xQty    = $Template.qtyColumnXRange
    $yMax    = $Template.itemRowYMax
    $yMin    = $Template.itemRowYMin
    $codePat = $Template.itemCodePattern
    $qtyPat  = if ($Template.PSObject.Properties['qtyPattern']) { $Template.qtyPattern } else { '^\d+,\d{2}$' }

    $hasDeliveryAddrCfg = $Template.PSObject.Properties['deliveryAddressXMin'] -and
                          $Template.PSObject.Properties['deliveryAddressYRange']
    $xAddrMin = if ($hasDeliveryAddrCfg) { $Template.deliveryAddressXMin } else { 9999 }
    $yAddrMin = if ($hasDeliveryAddrCfg) { $Template.deliveryAddressYRange[0] } else { 0 }
    $yAddrMax = if ($hasDeliveryAddrCfg) { $Template.deliveryAddressYRange[1] } else { 0 }

    $yDMin = if ($hasDelivDateCfg) { $Template.deliveryDateYRange[0] } else { 0 }
    $yDMax = if ($hasDelivDateCfg) { $Template.deliveryDateYRange[1] } else { 0 }

    $rowData  = @{}
    $addrRows = @{}  # delivery address words grouped by Bottom row
    $pageNum  = 0

    foreach ($page in $pdf.GetPages()) {
        $pageNum++
        foreach ($word in $page.GetWords()) {
            $b = [Math]::Round($word.BoundingBox.Bottom, 1)
            $l = $word.BoundingBox.Left

            # Line item rows
            if ($b -ge $yMin -and $b -le $yMax) {
                $key = $b.ToString()
                if (-not $rowData.ContainsKey($key)) { $rowData[$key] = @{ ref = ''; qty = '' } }
                if ($l -ge $xRef[0] -and $l -le $xRef[1] -and
                    -not $rowData[$key].ref -and $word.Text -match $codePat) {
                    $rowData[$key].ref = $word.Text
                }
                if ($l -ge $xQty[0] -and $l -le $xQty[1] -and
                    -not $rowData[$key].qty -and $word.Text -match $qtyPat) {
                    $rowData[$key].qty = $word.Text
                }
            }

            # Delivery date: date-format word in the footer Y-band (coordinate-based)
            if ($hasDelivDateCfg -and -not $deliveryDateBC -and
                $b -ge $yDMin -and $b -le $yDMax -and $word.Text -match '^\d{2}/\d{2}/\d{4}$') {
                $parsed = [DateTime]::ParseExact(
                    $word.Text, $Template.dateInputFormat,
                    [System.Globalization.CultureInfo]::InvariantCulture)
                $deliveryDateBC = $parsed.ToString('yyyy-MM-dd')
            }

            # Delivery address: right column, top of PAGE 1 only (later pages are T&Cs)
            if ($pageNum -eq 1 -and $hasDeliveryAddrCfg -and $b -ge $yAddrMin -and $b -le $yAddrMax -and $l -ge $xAddrMin) {
                if (-not $addrRows.ContainsKey($b)) {
                    $addrRows[$b] = [System.Collections.Generic.List[string]]::new()
                }
                $addrRows[$b].Add($word.Text)
            }
        }
    }

    $pdf.Dispose()

    # Parse delivery address rows into BC ship-to fields
    # Rows sorted top-to-bottom (highest Bottom = highest on page = first)
    # Row 1 = name, last row = postcode + city, middle rows = address lines
    $shipTo = $null
    if ($hasDeliveryAddrCfg -and $addrRows.Count -ge 2) {
        $sortedRows = $addrRows.GetEnumerator() |
            Sort-Object { -[double]$_.Key } |
            ForEach-Object { ($_.Value -join ' ').Trim() }

        $name    = $sortedRows[0]
        $lastRow = $sortedRows[-1]
        $midRows = if ($sortedRows.Count -gt 2) { @($sortedRows[1..($sortedRows.Count - 2)]) } else { @() }

        # Last row: "26500 BOURG LES VALENCE" — first token is postcode if numeric
        $lastParts = $lastRow -split '\s+'
        if ($lastParts[0] -match '^\d{4,6}$') {
            $postCode = $lastParts[0]
            $city     = ($lastParts[1..($lastParts.Length - 1)]) -join ' '
        } else {
            $postCode = ''; $city = $lastRow
        }

        $country = if ($Template.PSObject.Properties['shipToCountry']) { $Template.shipToCountry } else { '' }

        $shipTo = [PSCustomObject]@{
            name         = $name
            addressLine1 = if ($midRows.Count -ge 1) { $midRows[0] } else { '' }
            addressLine2 = if ($midRows.Count -ge 2) { $midRows[1] } else { '' }
            city         = $city
            postCode     = $postCode
            country      = $country
        }
    }

    # Apply item mapping if mapping.json exists
    $mapping = $null
    if (Test-Path $mapPath) { $mapping = Get-Content $mapPath | ConvertFrom-Json }

    $lines = @()
    foreach ($key in ($rowData.Keys | Sort-Object { [double]$_ } -Descending)) {
        $row = $rowData[$key]
        if ($row.ref -and $row.qty) {
            $item = $row.ref
            if ($mapping -and $mapping.PSObject.Properties[$item]) {
                $item = $mapping.PSObject.Properties[$item].Value
            }
            $lines += [PSCustomObject]@{
                ItemNumber = $item
                Quantity   = [double]($row.qty -replace ',', '.')
            }
        }
    }

    return [PSCustomObject]@{
        OrderRef       = $orderRef
        OrderDate      = $orderDateBC
        DeliveryDate   = $deliveryDateBC
        ShipTo         = $shipTo
        ShipToPostCode = ''
        Lines          = $lines
    }
}

# ---------------------------------------------------------------------------
# BC write
# ---------------------------------------------------------------------------
function Submit-SalesOrder {
    param(
        [PSCustomObject]$OrderData,
        [PSCustomObject]$Template,
        [string]$ShipToCode = ''
    )

    $env     = $Template.environment
    $apiBase = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$env/api/v2.0/companies($companyId)"
    $odataBase = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$env/ODataV4/Company('Montandor_Andorra')"

    # Step 1: POST sales order header
    $order = $null
    try {
        $order   = Invoke-RestMethod -Method Post -Uri "$apiBase/salesOrders" -Headers $jsonHeader `
            -Body (@{ customerNumber = $Template.customerNumber; orderDate = $OrderData.OrderDate } | ConvertTo-Json)
        Write-Host "    [OK] Order header created: $($order.number)" -ForegroundColor Green
    } catch {
        throw "POST sales order header failed: $(Get-ApiError $_)"
    }
    $orderId = $order.id
    $orderNo = $order.number

    # Step 2: OData PATCH — Your_Reference + Requested_Delivery_Date
    $odataPatch = @{}
    if ($OrderData.OrderRef)     { $odataPatch['Your_Reference']          = $OrderData.OrderRef }
    if ($OrderData.DeliveryDate) { $odataPatch['Requested_Delivery_Date'] = $OrderData.DeliveryDate }

    if ($odataPatch.Count -gt 0) {
        try {
            $odata = Invoke-RestMethod -Method Get `
                -Uri "$odataBase/SalesOrder(Document_Type='Order',No='$orderNo')" -Headers $authHeader
            Invoke-RestMethod -Method Patch `
                -Uri "$odataBase/SalesOrder(Document_Type='Order',No='$orderNo')" `
                -Headers ($authHeader + @{ 'Content-Type' = 'application/json'; 'If-Match' = $odata.'@odata.etag' }) `
                -Body ($odataPatch | ConvertTo-Json) | Out-Null
            if ($OrderData.OrderRef)     { Write-Host "    [OK] Your Reference     : $($OrderData.OrderRef)" -ForegroundColor Green }
            if ($OrderData.DeliveryDate) { Write-Host "    [OK] Requested Delivery : $($OrderData.DeliveryDate)" -ForegroundColor Green }
        } catch {
            Write-Host "    [WARN] OData PATCH failed (Your Reference / Delivery Date): $(Get-ApiError $_)" -ForegroundColor Yellow
        }
    }

    # Step 3: ship-to — code (text-mode BC lookup) or custom address fields (words-mode coordinate extraction)
    if ($ShipToCode) {
        try {
            $reGet  = Invoke-RestMethod -Uri "$apiBase/salesOrders($orderId)" -Headers $authHeader
            $patchH = $authHeader + @{ 'Content-Type' = 'application/json'; 'If-Match' = $reGet.'@odata.etag' }
            Invoke-RestMethod -Method Patch -Uri "$apiBase/salesOrders($orderId)" -Headers $patchH `
                -Body (@{ shipToCode = $ShipToCode } | ConvertTo-Json) | Out-Null
            Write-Host "    [OK] Ship-to code       : $ShipToCode" -ForegroundColor Green
        } catch {
            Write-Host "    [WARN] Ship-to code PATCH failed: $(Get-ApiError $_)" -ForegroundColor Yellow
        }
    } elseif ($OrderData.ShipTo) {
        $st = $OrderData.ShipTo
        try {
            $reGet  = Invoke-RestMethod -Uri "$apiBase/salesOrders($orderId)" -Headers $authHeader
            $patchH = $authHeader + @{ 'Content-Type' = 'application/json'; 'If-Match' = $reGet.'@odata.etag' }
            Invoke-RestMethod -Method Patch -Uri "$apiBase/salesOrders($orderId)" -Headers $patchH `
                -Body (@{
                    shipToName         = $st.name
                    shipToAddressLine1 = $st.addressLine1
                    shipToAddressLine2 = $st.addressLine2
                    shipToCity         = $st.city
                    shipToPostCode     = $st.postCode
                    shipToCountry      = $st.country
                } | ConvertTo-Json) | Out-Null
            Write-Host "    [OK] Ship-to (custom)   : $($st.name) | $($st.addressLine1), $($st.postCode) $($st.city)" -ForegroundColor Green
        } catch {
            Write-Host "    [WARN] Ship-to PATCH failed: $(Get-ApiError $_)" -ForegroundColor Yellow
        }
    }

    # Step 4: POST sales order lines
    $lineErrors = 0
    foreach ($line in $OrderData.Lines) {
        try {
            $result = Invoke-RestMethod -Method Post `
                -Uri "$apiBase/salesOrders($orderId)/salesOrderLines" -Headers $jsonHeader `
                -Body (@{ lineType = 'Item'; lineObjectNumber = $line.ItemNumber; quantity = $line.Quantity } | ConvertTo-Json)
            Write-Host ("    [OK] {0,-22} qty {1,6}  -> EUR {2}" -f $line.ItemNumber, $line.Quantity, $result.unitPrice) -ForegroundColor Green
        } catch {
            Write-Host ("    [ERROR] Line {0} qty {1}: {2}" -f $line.ItemNumber, $line.Quantity, (Get-ApiError $_)) -ForegroundColor Red
            $lineErrors++
        }
    }

    if ($lineErrors -gt 0) {
        Write-Host "    $lineErrors line(s) failed — order $orderNo is open in BC, review manually." -ForegroundColor Yellow
    }

    return $orderNo
}

# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------
if ($DryRun) { Write-Host "`n*** DRY RUN — no BC writes ***" -ForegroundColor Yellow }

$clientDirs = Get-ChildItem $rootDir -Directory |
    Where-Object { -not $ClientFolder -or $_.Name -eq $ClientFolder }

$processed = 0
foreach ($dir in $clientDirs) {
    $tplPath = Join-Path $dir.FullName 'template.json'
    if (-not (Test-Path $tplPath)) { continue }

    $tpl        = Get-Content $tplPath | ConvertFrom-Json
    $bcItems    = @()
    if ($tpl.PSObject.Properties['extractionMode'] -and $tpl.extractionMode -eq 'text') {
        $bcItems = Get-BcItemNumbers -Environment $tpl.environment
    }
    $notProcDir = Join-Path $dir.FullName 'Not_Processed'
    $procDir    = Join-Path $dir.FullName 'Processed'

    $pdfs = @(Get-ChildItem "$notProcDir\*" -Include '*.pdf','*.PDF' -ErrorAction SilentlyContinue)
    if ($pdfs.Count -eq 0) { continue }

    Write-Host ("`n[{0}] {1} — {2} PDF(s)" -f $dir.Name, $tpl.clientName, $pdfs.Count) -ForegroundColor Cyan

    foreach ($pdf in $pdfs) {
        Write-Host "  $($pdf.Name)" -ForegroundColor White
        try {
            $data = Get-PdfOrderData -PdfPath $pdf.FullName -Template $tpl -BcItemNumbers $bcItems
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            # Ship-to BC lookup for text-mode clients
            $isTextMode = $tpl.PSObject.Properties['extractionMode'] -and $tpl.extractionMode -eq 'text'
            $shipToCode = ''
            $skipOrder  = $false
            $stLabel    = if ($data.ShipTo) { "$($data.ShipTo.name) | $($data.ShipTo.addressLine1), $($data.ShipTo.postCode) $($data.ShipTo.city)" } else { '(default)' }

            if ($isTextMode -and $data.ShipToPostCode) {
                $candidates = Get-BcShipToCode -CustomerNumber $tpl.customerNumber -PostCode $data.ShipToPostCode -Environment $tpl.environment
                if ($candidates.Count -eq 0) {
                    $stLabel   = "SKIP — no BC ship-to for postcode $($data.ShipToPostCode)"
                    $skipOrder = $true
                } elseif ($candidates.Count -eq 1) {
                    $shipToCode = $candidates[0].code
                    $stLabel    = "$shipToCode — $($candidates[0].displayName)"
                } else {
                    $codes     = ($candidates | ForEach-Object { $_.code }) -join ', '
                    $stLabel   = "SKIP — ambiguous postcode $($data.ShipToPostCode): $codes"
                    $skipOrder = $true
                }
            }

            # Deduplication check (runs in both dry-run and execute modes)
            $alreadyExists = $false
            if ($data.OrderRef) {
                $alreadyExists = Test-BcOrderExists -CustomerNumber $tpl.customerNumber -OrderRef $data.OrderRef -Environment $tpl.environment
            }

            Write-Host "    Order ref        : $($data.OrderRef)"
            Write-Host "    Order date       : $($data.OrderDate)"
            Write-Host "    Requested deliv. : $(if($data.DeliveryDate){$data.DeliveryDate}else{'(not found)'})"
            Write-Host "    Ship-to          : $stLabel"
            Write-Host "    Duplicate check  : $(if($alreadyExists){'ALREADY IN BC — would skip'}elseif(-not $data.OrderRef){'(no order ref — cannot check)'}else{'OK'})"
            Write-Host "    Lines ($($data.Lines.Count)):"
            $data.Lines | ForEach-Object {
                Write-Host ("      {0,-22} qty {1}" -f $_.ItemNumber, $_.Quantity)
            }

            if (-not $DryRun) {
                if ($alreadyExists) {
                    Write-Host "    [SKIP] Order ref $($data.OrderRef) already exists in BC (open, released, or posted)." -ForegroundColor Yellow
                } elseif ($skipOrder) {
                    Write-Host "    [SKIP] Order not posted — register ship-to postcode $($data.ShipToPostCode) in BC and resubmit." -ForegroundColor Yellow
                } else {
                    $bcNo = Submit-SalesOrder -OrderData $data -Template $tpl -ShipToCode $shipToCode
                    Move-Item -Path $pdf.FullName -Destination (Join-Path $procDir $pdf.Name) -Force
                    Write-Host "    -> BC $bcNo | moved to Processed" -ForegroundColor Green
                    $processed++
                }
            }
        } catch {
            Write-Host "    ERROR: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

$clientSecret = $null; $token = $null

if ($DryRun) {
    Write-Host "`n*** Dry run complete. Use -Execute to post to BC. ***`n" -ForegroundColor Yellow
} else {
    Write-Host "`nDone. $processed order(s) posted to BC.`n" -ForegroundColor Green
}
