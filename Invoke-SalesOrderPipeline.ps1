<#
.SYNOPSIS
    Template-driven sales order pipeline: PDF -> BC (Sandbox-Training or Production).
.DESCRIPTION
    Scans each client Not_Processed folder, extracts order data from PDFs using
    per-client template.json (coordinate-based column extraction via PdfPig),
    then POSTs to the Montandor pipeline API (salesOrderCreations) to create each order
    atomically with yourReference and requestedDeliveryDate set in one call.
    Runs in dry-run mode by default. Pass -Execute to write to BC.
.PARAMETER Execute
    Actually POST to BC and move PDFs. Without this flag, only prints what would happen.
.PARAMETER ClientFolder
    Restrict to one client subfolder (e.g. 'RETIF_FITER'). Omit to process all.
#>
param(
    [switch]$Execute,
    [string]$ClientFolder = '',
    [switch]$FunctionsOnly   # dot-source mode: load functions + tokens, skip main loop
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$DryRun  = -not $Execute
$rootDir = $PSScriptRoot
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
$bcItemCache  = @{}
$custRefCache = @{}
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
    $odata  = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$Environment/ODataV4/Company('Montandor_Andorra')"
    $safeNo = $CustomerNumber -replace "'", "''"
    $safeRef = $OrderRef      -replace "'", "''"

    # Open + Released orders (Document_Type=Order covers both)
    $f1 = "Sell_to_Customer_No eq '$safeNo' and Your_Reference eq '$safeRef' and Document_Type eq 'Order'"
    $r1 = Invoke-RestMethod -Uri "$odata/SalesOrder?`$filter=$f1&`$select=No&`$top=1" -Headers $authHeader
    if (@($r1.value).Count -gt 0) { return $true }

    # Posted invoices — via Order Pipeline AL extension (api/montandor/pipeline/v1.0)
    $pipelineBase = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$Environment/api/montandor/pipeline/v1.0/companies($companyId)"
    $f2 = "sellToCustomerNo eq '$safeNo' and yourReference eq '$safeRef'"
    $r2 = Invoke-RestMethod -Uri "$pipelineBase/postedSalesInvoices?`$filter=$f2&`$select=no&`$top=1" -Headers $authHeader
    if (@($r2.value).Count -gt 0) { return $true }

    return $false
}

# ---------------------------------------------------------------------------
# BC ship-to address lookup (Order Pipeline extension — api/montandor/pipeline/v1.0)
# ---------------------------------------------------------------------------
function Get-BcShipToCode {
    param([string]$CustomerNumber, [string]$PostCode, [string]$Environment)
    $safeNo   = $CustomerNumber -replace "'", "''"
    $safePost = $PostCode       -replace "'", "''"
    $uri = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$Environment/api/montandor/pipeline/v1.0/companies($companyId)/shipToAddresses?`$filter=customerNumber eq '$safeNo' and postCode eq '$safePost'"
    $resp = Invoke-RestMethod -Uri $uri -Headers $authHeader
    return @($resp.value)
}

# ---------------------------------------------------------------------------
# Fetch open sales order lines from BC (for change detection)
# Returns array of {ItemNumber, Quantity} objects, or $null if order is posted/unavailable
# ---------------------------------------------------------------------------
function Get-BcOrderLines {
    param([string]$CustomerNumber, [string]$OrderRef, [string]$Environment)
    $odata   = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$Environment/ODataV4/Company('Montandor_Andorra')"
    $apiBase = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$Environment/api/v2.0/companies($companyId)"
    try {
        # Step 1: use OData to find the BC order No (yourReference not filterable in REST API v2.0)
        $safeNo  = $CustomerNumber -replace "'", "''"
        $safeRef = $OrderRef       -replace "'", "''"
        $f  = "Sell_to_Customer_No eq '$safeNo' and Your_Reference eq '$safeRef' and Document_Type eq 'Order'"
        $r  = Invoke-RestMethod -Uri "$odata/SalesOrder?`$filter=$f&`$select=No&`$top=1" -Headers $authHeader
        if (@($r.value).Count -eq 0) { return $null }  # posted invoice or not found
        $bcNo = $r.value[0].No

        # Step 2: get systemId via REST API v2.0 (needed for lines sub-endpoint)
        $orders = Invoke-RestMethod -Uri "$apiBase/salesOrders?`$filter=number eq '$bcNo'&`$select=id" -Headers $authHeader
        if (@($orders.value).Count -eq 0) { return $null }
        $orderId = $orders.value[0].id

        # Step 3: fetch item lines
        $linesResp = Invoke-RestMethod `
            -Uri "$apiBase/salesOrders($orderId)/salesOrderLines?`$filter=lineType eq 'Item'&`$select=lineObjectNumber,quantity" `
            -Headers $authHeader
        return @($linesResp.value | ForEach-Object {
            [PSCustomObject]@{ ItemNumber = $_.lineObjectNumber; Quantity = $_.quantity }
        })
    } catch {
        return $null  # treat as simple duplicate if lines cannot be fetched
    }
}

# ---------------------------------------------------------------------------
# Fetch current ship-to address from BC for an open sales order
# ---------------------------------------------------------------------------
function Get-BcOrderShipTo {
    param([string]$CustomerNumber, [string]$OrderRef, [string]$Environment)
    $odata   = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$Environment/ODataV4/Company('Montandor_Andorra')"
    $safeNo  = $CustomerNumber -replace "'", "''"
    $safeRef = $OrderRef       -replace "'", "''"
    $f       = "Sell_to_Customer_No eq '$safeNo' and Your_Reference eq '$safeRef' and Document_Type eq 'Order'"
    try {
        $r = Invoke-RestMethod `
            -Uri "$odata/SalesOrder?`$filter=$f&`$select=No,Ship_to_Name,Ship_to_Address,Ship_to_City,Ship_to_Post_Code,Ship_to_Country,Ship_to_Code&`$top=1" `
            -Headers $authHeader
        if (@($r.value).Count -eq 0) { return $null }
        $o = $r.value[0]
        return [PSCustomObject]@{
            name         = $o.Ship_to_Name
            addressLine1 = $o.Ship_to_Address
            city         = $o.Ship_to_City
            postCode     = $o.Ship_to_Post_Code
            country      = $o.Ship_to_Country
            code         = $o.Ship_to_Code
        }
    } catch { return $null }
}

# ---------------------------------------------------------------------------
# Compare extracted PDF lines against existing BC order lines
# Returns array of {Type, ItemNumber, OldQty, NewQty} objects; empty = identical
# ---------------------------------------------------------------------------
function Compare-OrderLines {
    param([PSCustomObject[]]$NewLines, [PSCustomObject[]]$BcLines)
    $newMap = @{}; foreach ($l in $NewLines) { $newMap[$l.ItemNumber] = $l.Quantity }
    $bcMap  = @{}; foreach ($l in $BcLines)  { $bcMap[$l.ItemNumber]  = $l.Quantity }
    $diff = @()
    foreach ($item in ($bcMap.Keys | Sort-Object)) {
        if ($newMap.ContainsKey($item)) {
            if ([Math]::Abs($newMap[$item] - $bcMap[$item]) -gt 0.001) {
                $diff += [PSCustomObject]@{ Type = 'Changed'; ItemNumber = $item; OldQty = $bcMap[$item]; NewQty = $newMap[$item] }
            }
        } else {
            $diff += [PSCustomObject]@{ Type = 'Removed'; ItemNumber = $item; OldQty = $bcMap[$item]; NewQty = $null }
        }
    }
    foreach ($item in ($newMap.Keys | Sort-Object)) {
        if (-not $bcMap.ContainsKey($item)) {
            $diff += [PSCustomObject]@{ Type = 'Added'; ItemNumber = $item; OldQty = $null; NewQty = $newMap[$item] }
        }
    }
    return $diff
}

function Format-Qty { param($Q); if ($null -eq $Q) { return '-' }; if ($Q -eq [Math]::Floor($Q)) { return [string][int]$Q } else { return [string]$Q } }

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

# Normalise a customer reference / Code Article for matching: strip whitespace and leading zeros
# so the PDF's zero-padded code (e.g. "012702") matches BC's stored value (e.g. "12702").
function Normalize-RefNo {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    $t = ($Value -replace '\s', '').TrimStart('0')
    if ($t -eq '') { return '0' }
    return $t
}

# ---------------------------------------------------------------------------
# Customer item references (cached per environment+customer). Maps the customer's own
# article number (the PDF "Code Article" column) to our BC item number, read from BC's
# Item Reference table via the auto-published Item_References_Excel OData page. This is the
# authoritative, business-maintained cross-reference — used to resolve codes that the
# supplier-reference column truncates or renders ambiguous (e.g. In Situ "012702" ->
# "RS-RT-RVS-GY-SET", where the supplier ref is cut to the ambiguous prefix "RS-RT-RVS-GY-").
# Returns @{ normalisedRefNo -> BC item number }. On any failure returns an empty map so the
# caller falls back to the existing supplier-reference resolution (no behaviour change).
# ---------------------------------------------------------------------------
function Get-CustomerItemReferences {
    param([string]$CustomerNumber, [string]$Environment)
    $cacheKey = "$Environment|$CustomerNumber"
    if ($custRefCache.ContainsKey($cacheKey)) { return $custRefCache[$cacheKey] }
    Write-Host "  [REFS] Loading customer item references ($CustomerNumber, $Environment)..." -ForegroundColor Cyan
    $odata = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$Environment/ODataV4/Company('Montandor_Andorra')"
    $safe  = $CustomerNumber -replace "'", "''"
    $flt   = "Reference_Type eq 'Customer' and Reference_Type_No eq '$safe'"
    $map   = @{}
    try {
        $url = "$odata/Item_References_Excel?`$filter=$([uri]::EscapeDataString($flt))&`$top=1000"
        do {
            $r = Invoke-RestMethod -Uri $url -Headers $authHeader
            foreach ($ref in $r.value) {
                $k = Normalize-RefNo $ref.Reference_No
                if ($k -and -not $map.ContainsKey($k)) { $map[$k] = $ref.Item_No }
            }
            $url = if ($r.PSObject.Properties['@odata.nextLink']) { $r.'@odata.nextLink' } else { $null }
        } while ($url)
        Write-Host "    $($map.Count) references loaded." -ForegroundColor Green
    } catch {
        Write-Host "    [WARN] Item reference load failed ($($_.Exception.Message)); falling back to supplier-ref resolution." -ForegroundColor Yellow
    }
    $custRefCache[$cacheKey] = $map
    return $map
}

# ---------------------------------------------------------------------------
# Hyphen/punctuation-insensitive BC item index. Returns @{ normalizedKey -> BC number }
# but ONLY for keys that map to a single BC item; ambiguous stripped forms are omitted
# so they fall through to the unknown-code path rather than guess.
# ---------------------------------------------------------------------------
function Build-NormalizedItemIndex {
    param([string[]]$BcItemNumbers)
    $groups = @{}
    foreach ($n in $BcItemNumbers) {
        if (-not $n) { continue }
        $k = ($n -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
        if (-not $groups.ContainsKey($k)) { $groups[$k] = [System.Collections.Generic.List[string]]::new() }
        $groups[$k].Add($n)
    }
    $index = @{}
    foreach ($kv in $groups.GetEnumerator()) {
        if ($kv.Value.Count -eq 1) { $index[$kv.Key] = $kv.Value[0] }
    }
    return $index
}

# Resolve a raw PDF code to a BC item number.
# Order: explicit mapping.json override -> exact BC match -> unique hyphen-insensitive match.
# Returns $null if unresolved (caller treats it as an unknown code).
function Resolve-ItemCode {
    param([string]$Code, [hashtable]$ExactSet, [hashtable]$NormalizedIndex, $Mapping)
    if (-not $Code) { return $null }
    if ($Mapping -and $Mapping.PSObject.Properties[$Code]) { return $Mapping.PSObject.Properties[$Code].Value }
    if ($ExactSet.ContainsKey($Code)) { return $ExactSet[$Code] }
    $nk = ($Code -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
    if ($NormalizedIndex.ContainsKey($nk)) { return $NormalizedIndex[$nk] }
    return $null
}

# ---------------------------------------------------------------------------
# Suggest likely BC item(s) for an UNRECOGNISED code — used ONLY as a "did you mean"
# hint in alert emails, NEVER to auto-post a line. A truncation drops characters off the
# end, so the real BC code (hyphen-stripped) starts with the unknown code (hyphen-stripped).
# ---------------------------------------------------------------------------
function Get-ItemSuggestion {
    param([string]$Code, [string[]]$BcItemNumbers, [int]$Max = 4)
    if (-not $Code) { return @() }
    $nc = ($Code -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
    if ($nc.Length -lt 4) { return @() }
    $pairs = foreach ($n in $BcItemNumbers) {
        if ($n) { [PSCustomObject]@{ Number = $n; Norm = ($n -replace '[^A-Za-z0-9]', '').ToUpperInvariant() } }
    }
    # 1) BC codes the unknown code is a prefix of (classic end-truncation)
    $hits = @($pairs | Where-Object { $_.Norm -ne $nc -and $_.Norm.StartsWith($nc) } | Select-Object -ExpandProperty Number)
    # 2) fallback: share the first 6 chars (truncation plus a near-end typo)
    if ($hits.Count -eq 0 -and $nc.Length -ge 6) {
        $stem = $nc.Substring(0, 6)
        $hits = @($pairs | Where-Object { $_.Norm.StartsWith($stem) } | Select-Object -ExpandProperty Number)
    }
    return @($hits | Select-Object -Unique -First $Max)
}

# Format one unrecognised code as an HTML-encoded list entry, with a "did you mean" hint if available.
function Format-UnknownCodeHtml {
    param([string]$Code, [string[]]$BcItemNumbers)
    $enc = [System.Net.WebUtility]::HtmlEncode($Code)
    $sug = @(Get-ItemSuggestion -Code $Code -BcItemNumbers $BcItemNumbers)
    if ($sug.Count -gt 0) {
        $sugEnc = ($sug | ForEach-Object { [System.Net.WebUtility]::HtmlEncode($_) }) -join ', '
        return "$enc &nbsp;&rarr;&nbsp; did you mean: <em>$sugEnc</em>?"
    }
    return $enc
}

# ---------------------------------------------------------------------------
# PDF extraction (coordinate-based)
# ---------------------------------------------------------------------------
function Get-PdfOrderData {
    param(
        [string]$PdfPath     = '',
        [byte[]]$PdfBytes    = $null,
        [PSCustomObject]$Template,
        [string[]]$BcItemNumbers = @(),
        [string]$ClientDir   = ''
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
    $mapBase = if ($ClientDir) { $ClientDir } else { Split-Path $PdfPath }
    $mapPath = Join-Path $mapBase 'mapping.json'

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

        # Item lines: single pass over PDF text in document order to preserve PDF line sequence
        $itemLookup = [System.Collections.Generic.Dictionary[string,string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($itemNo in $BcItemNumbers) {
            if (-not $itemLookup.ContainsKey($itemNo)) { $itemLookup[$itemNo] = $itemNo }
        }
        if (Test-Path $mapPath) {
            $renames = Get-Content $mapPath | ConvertFrom-Json
            foreach ($prop in $renames.PSObject.Properties) {
                $itemLookup[$prop.Name] = $prop.Value
                if ($prop.Name -ne $prop.Value -and $itemLookup.ContainsKey($prop.Value)) {
                    $itemLookup.Remove($prop.Value)
                }
            }
        }
        $lines        = @()
        $unknownCodes = @()
        $seenItems    = [System.Collections.Generic.HashSet[string]]::new()

        if ($Template.PSObject.Properties['qtyBeforeCode'] -and $Template.qtyBeforeCode) {
            # Qty-before-code mode (e.g. Aligro-Demaurex: "54x    2PI...ITEM_CODE 2 PI").
            # The ordered quantity (54) precedes the item code in the row as {N}x.
            #
            # Pattern: \s\d{2,3}\d{6}\s*(\d{1,4})x
            # Anchors on the full row structure: space + 2-3 digit Pos + 6-digit NoArt + qty.
            # A naive \d{6}\s*(\d{1,4})x matches at offset inside "200443674100x", yielding
            # qty=4100 instead of 100. Including the Pos prefix pins the match to row starts.
            # Dimensions ("60x115cm") and pack labels ("7x") in descriptions are never
            # preceded by a space+Pos+NoArt sequence — they are never matched.
            #
            # (?<![A-Z\-]) on codes prevents substring matches (SMA510-WT inside BL-SMA510-WT)
            # while allowing codes that follow a digit in a description (e.g. "A4MC-BRA4-WR").
            $qtyPositions = @([regex]::Matches($allText, '\s\d{2,3}\d{6}\s*(\d{1,4})x') | ForEach-Object {
                [PSCustomObject]@{ Idx = $_.Index; Qty = [double]$_.Groups[1].Value }
            })
            $codeHits = @{}
            foreach ($code in $itemLookup.Keys) {
                $m = [regex]::Match($allText, '(?<![A-Z\-])' + [regex]::Escape($code))
                if ($m.Success) {
                    $preceding = @($qtyPositions | Where-Object { $_.Idx -lt $m.Index }) | Select-Object -Last 1
                    if ($preceding) { $codeHits[$code] = @{ Idx = $m.Index; Qty = $preceding.Qty } }
                }
            }
            foreach ($code in ($codeHits.Keys | Sort-Object { $codeHits[$_].Idx })) {
                $mapped = $itemLookup[$code]
                if ($seenItems.Add($mapped)) {
                    $lines += [PSCustomObject]@{
                        ItemNumber = $mapped
                        Quantity   = $codeHits[$code].Qty
                    }
                }
            }
            # Detect codes in Réf.Fourn. column not matched to BC — code followed by pack qty + unit.
            # Two false-positive guards:
            #   - Require hyphen in code (filters header/address text like "OFICINA 21140 AV")
            #   - Require code is not a suffix match of a known BC code (filters "A4MC-CBA4-BR"
            #     which is description "A4" concatenated with the known code "MC-CBA4-BR")
            $seenUnknown = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($m in [regex]::Matches($allText, '(?<![A-Z\-])([A-Z][A-Z0-9\-]{2,})\s+\d+\s+[A-Z]{2,3}')) {
                $code = $m.Groups[1].Value
                if ($code -notmatch '-') { continue }
                if (-not $codeHits.ContainsKey($code) -and $seenUnknown.Add($code)) {
                    $isSuffix = $codeHits.Keys | Where-Object { $code.EndsWith($_, [System.StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
                    if (-not $isSuffix) { $unknownCodes += $code }
                }
            }
        } elseif ($Template.PSObject.Properties['qtyUnit'] -and $Template.qtyUnit) {
            # Integer+unit qty mode (e.g. GAFIC: 8PCE, 18PQT).
            # A general pattern would greedily absorb quantity digits into the code (ELE-BL-LA12PCE
            # becomes code=ELE-BL-LA1, qty=2). Per-code exact search avoids this entirely.
            $unitPat  = $Template.qtyUnit
            $recover  = $Template.PSObject.Properties['recoverFromDesignation'] -and $Template.recoverFromDesignation
            $codeHits = @{}
            foreach ($code in $itemLookup.Keys) {
                $m = [regex]::Match($allText, [regex]::Escape($code) + "(\d+)(?:$unitPat)")
                if ($m.Success) { $codeHits[$code] = @{ Idx = $m.Index; Qty = $m.Groups[1].Value } }
            }

            # Collect line records carrying text position, so any line recovered from the Désignation
            # (below) sorts back into document order alongside the per-code matches.
            $lineRecs = @()
            foreach ($code in $codeHits.Keys) {
                $mapped = $itemLookup[$code]
                if ($seenItems.Add($mapped)) {
                    $lineRecs += [PSCustomObject]@{ Idx = $codeHits[$code].Idx; ItemNumber = $mapped; Quantity = [double]$codeHits[$code].Qty }
                }
            }

            # Detect codes in VERMES# references not matched to BC — unit suffix forces correct backtracking.
            # Position-based dedup: if a known code was already matched at the same text position, skip —
            # the greedy regex may extract a slightly different string (e.g. ELE-M-SM1 vs ELE-M-SM) but
            # the position overlap proves the item was already handled by the per-code exact search above.
            $knownPositions = [System.Collections.Generic.HashSet[int]]::new()
            foreach ($v in $codeHits.Values) { [void]$knownPositions.Add($v.Idx) }
            $seenUnknown = [System.Collections.Generic.HashSet[string]]::new()
            $prevIdx     = 0
            foreach ($m in [regex]::Matches($allText, '#([A-Z][A-Z0-9\-]+)(\d+)(?:' + $unitPat + ')')) {
                $curIdx = $m.Index
                if ($knownPositions.Contains($curIdx + 1)) { $prevIdx = $curIdx; continue }
                $code = $m.Groups[1].Value
                $qty  = $m.Groups[2].Value
                if ($codeHits.ContainsKey($code) -or -not $seenUnknown.Add($code)) { $prevIdx = $curIdx; continue }

                # Désignation recovery (GAFIC, gated by recoverFromDesignation): the Ref Fournisseur is
                # sometimes truncated in its own column (e.g. BL-SMA100-V7) while the full BC item code
                # appears inside the row's Désignation text (e.g. "...COLORE X7 BL-SMA100-V7-AS [14477]").
                # Search this row's window (bounded by the previous and current VERMES# anchors) for the
                # truncated code extended to the next space/bracket; if that fuller string is a BC item
                # (or mapping alternative) use it. Qty comes from the VERMES# anchor. The leading Ref is
                # glued to the description (BL-SMA100-V7FEUTRE...) so it self-rejects — only the real code matches.
                $rec = $null
                if ($recover) {
                    $winStart = [Math]::Max(0, $prevIdx)
                    $window   = $allText.Substring($winStart, $curIdx - $winStart)
                    foreach ($mm in [regex]::Matches($window, [regex]::Escape($code) + '[^\s\[\]]*')) {
                        $cand = $mm.Value
                        if ($cand -eq $code) { continue }   # bare truncated ref — already known not in BC
                        if ($itemLookup.ContainsKey($cand)) { $rec = $itemLookup[$cand]; break }
                    }
                }
                if ($rec) {
                    if ($seenItems.Add($rec)) {
                        $lineRecs += [PSCustomObject]@{ Idx = $curIdx; ItemNumber = $rec; Quantity = [double]$qty }
                        Write-Host "    [RECOVERED] $code -> $rec (qty $qty) from Désignation" -ForegroundColor Green
                    }
                } else {
                    $unknownCodes += $code
                }
                $prevIdx = $curIdx
            }

            foreach ($r in ($lineRecs | Sort-Object Idx)) {
                $lines += [PSCustomObject]@{ ItemNumber = $r.ItemNumber; Quantity = $r.Quantity }
            }
        } else {
            # Comma-decimal qty mode (e.g. CR Distribution: 5,20) — single pass over text.
            $seenUnknown = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($m in [regex]::Matches($allText, '(?<![A-Z0-9\-])([A-Z][A-Z0-9\-]{2,})\s*(\d+,\d{2})')) {
                $code = $m.Groups[1].Value
                if ($itemLookup.ContainsKey($code)) {
                    $mapped = $itemLookup[$code]
                    if ($seenItems.Add($mapped)) {
                        $lines += [PSCustomObject]@{
                            ItemNumber = $mapped
                            Quantity   = [double]($m.Groups[2].Value -replace ',', '.')
                        }
                    }
                } elseif ($seenUnknown.Add($code)) {
                    $unknownCodes += $code
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
            UnknownCodes   = $unknownCodes
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

    # Optional "Code Article" column (customer's own article number). When present, each row's
    # Code Article is resolved via BC customer item references first (authoritative), with the
    # supplier-reference column as fallback. Opt-in per client via articleCodeXRange.
    $xArt   = if ($Template.PSObject.Properties['articleCodeXRange'])  { $Template.articleCodeXRange } else { $null }
    $artPat = if ($Template.PSObject.Properties['articleCodePattern']) { $Template.articleCodePattern } else { '^\d{4,6}$' }

    $hasDeliveryAddrCfg = $Template.PSObject.Properties['deliveryAddressXMin'] -and
                          $Template.PSObject.Properties['deliveryAddressYRange']
    $xAddrMin = if ($hasDeliveryAddrCfg) { $Template.deliveryAddressXMin } else { 9999 }
    $yAddrMin = if ($hasDeliveryAddrCfg) { $Template.deliveryAddressYRange[0] } else { 0 }
    $yAddrMax = if ($hasDeliveryAddrCfg) { $Template.deliveryAddressYRange[1] } else { 0 }

    $yDMin = if ($hasDelivDateCfg) { $Template.deliveryDateYRange[0] } else { 0 }
    $yDMax = if ($hasDelivDateCfg) { $Template.deliveryDateYRange[1] } else { 0 }

    # Collect code words and qty words separately. Codes keep one row per (page, rounded Bottom);
    # qty words are paired to the nearest code-row within a small baseline tolerance below, so a
    # qty nudged 1px off its code by a wrapped description is no longer dropped.
    $refRows  = @{}   # "P{page}_{bottom}" -> [PSCustomObject]@{ Ref; Page; Bottom }
    $qtyList  = [System.Collections.Generic.List[object]]::new()
    $artList  = [System.Collections.Generic.List[object]]::new()   # Code Article words (if configured)
    $addrRows = @{}  # delivery address words grouped by Bottom row
    $pageNum  = 0

    foreach ($page in $pdf.GetPages()) {
        $pageNum++
        foreach ($word in $page.GetWords()) {
            $b = [Math]::Round($word.BoundingBox.Bottom, 0)
            $l = $word.BoundingBox.Left

            # Line item rows
            if ($b -ge $yMin -and $b -le $yMax) {
                if ($l -ge $xRef[0] -and $l -le $xRef[1] -and $word.Text -match $codePat) {
                    $key = "P${pageNum}_${b}"
                    if (-not $refRows.ContainsKey($key)) {
                        $refRows[$key] = [PSCustomObject]@{ Ref = $word.Text; Page = $pageNum; Bottom = $b }
                    }
                }
                if ($l -ge $xQty[0] -and $l -le $xQty[1] -and $word.Text -match $qtyPat) {
                    $qtyList.Add([PSCustomObject]@{ Page = $pageNum; Bottom = $b; Text = $word.Text })
                }
                if ($xArt -and $l -ge $xArt[0] -and $l -le $xArt[1] -and $word.Text -match $artPat) {
                    $artList.Add([PSCustomObject]@{ Page = $pageNum; Bottom = $b; Text = $word.Text })
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
        try {
            $sortedRows = $addrRows.GetEnumerator() |
                Sort-Object { -[double]$_.Key } |
                ForEach-Object { ($_.Value -join ' ').Trim() }

            $name    = $sortedRows[0]
            $lastRow = $sortedRows[-1]
            [array]$midRows = if ($sortedRows.Count -gt 2) { @($sortedRows[1..($sortedRows.Count - 2)]) } else { @() }

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
        } catch {
            $shipTo = $null   # unexpected PDF layout — continue without ship-to; dedup check still runs
        }
    }

    # Pair each qty word, and each Code Article word, to the nearest code-row on the same page.
    $rowTol = 3
    foreach ($q in $qtyList) {
        $cand = $refRows.Values |
            Where-Object { $_.Page -eq $q.Page -and [Math]::Abs($_.Bottom - $q.Bottom) -le $rowTol } |
            Sort-Object { [Math]::Abs($_.Bottom - $q.Bottom) } |
            Select-Object -First 1
        if ($cand -and -not $cand.PSObject.Properties['Qty']) {
            $cand | Add-Member -NotePropertyName Qty -NotePropertyValue $q.Text
        }
    }
    foreach ($a in $artList) {
        $cand = $refRows.Values |
            Where-Object { $_.Page -eq $a.Page -and [Math]::Abs($_.Bottom - $a.Bottom) -le $rowTol } |
            Sort-Object { [Math]::Abs($_.Bottom - $a.Bottom) } |
            Select-Object -First 1
        if ($cand -and -not $cand.PSObject.Properties['Art']) {
            $cand | Add-Member -NotePropertyName Art -NotePropertyValue $a.Text
        }
    }

    # Resolve codes. When a Code Article column is configured, resolve each row by its Code Article
    # via BC customer item references FIRST (authoritative — immune to supplier-ref truncation/
    # ambiguity), then fall back to the supplier-ref column: mapping.json -> exact BC match ->
    # unique hyphen-insensitive match. Unresolved codes are collected so the caller diverts them to
    # the amber 'unrecognised items' alert instead of creating an order whose line POSTs all fail.
    $mapping = $null
    if (Test-Path $mapPath) { $mapping = Get-Content $mapPath | ConvertFrom-Json }
    $exactSet = @{}
    foreach ($n in $BcItemNumbers) { if ($n) { $exactSet[$n] = $n } }
    $normIndex = Build-NormalizedItemIndex -BcItemNumbers $BcItemNumbers

    $custRefMap = @{}
    if ($xArt -and $Template.PSObject.Properties['customerNumber'] -and $Template.customerNumber) {
        $custRefMap = Get-CustomerItemReferences -CustomerNumber $Template.customerNumber -Environment $Template.environment
    }

    $lines        = @()
    $unknownCodes = @()
    $seenUnknown  = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($r in ($refRows.Values | Sort-Object Page, @{ Expression = 'Bottom'; Descending = $true })) {
        if (-not $r.PSObject.Properties['Qty']) { continue }
        $resolved = $null
        $artCode  = if ($r.PSObject.Properties['Art']) { $r.Art } else { '' }
        if ($artCode) {
            $ak = Normalize-RefNo $artCode
            if ($custRefMap.ContainsKey($ak)) { $resolved = $custRefMap[$ak] }
        }
        if (-not $resolved) {
            $resolved = Resolve-ItemCode -Code $r.Ref -ExactSet $exactSet -NormalizedIndex $normIndex -Mapping $mapping
        }
        if ($resolved) {
            $lines += [PSCustomObject]@{ ItemNumber = $resolved; Quantity = [double]($r.Qty -replace ',', '.') }
        } elseif ($seenUnknown.Add($r.Ref)) {
            $unknownCodes += $r.Ref
        }
    }

    return [PSCustomObject]@{
        OrderRef       = $orderRef
        OrderDate      = $orderDateBC
        DeliveryDate   = $deliveryDateBC
        ShipTo         = $shipTo
        ShipToPostCode = ''
        Lines          = $lines
        UnknownCodes   = $unknownCodes
    }
}

# ---------------------------------------------------------------------------
# BC write
# ---------------------------------------------------------------------------
function Submit-SalesOrder {
    param(
        [PSCustomObject]$OrderData,
        [PSCustomObject]$Template,
        [string]$ShipToCode  = '',
        [byte[]]$PdfBytes    = $null,
        [string]$PdfFileName = ''
    )

    $env          = $Template.environment
    $apiBase      = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$env/api/v2.0/companies($companyId)"
    $pipelineBase = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$env/api/montandor/pipeline/v1.0/companies($companyId)"

    # Step 1: POST sales order — atomic: customerNumber + yourReference + requestedDeliveryDate in one call
    $postBody = [ordered]@{ customerNumber = $Template.customerNumber; orderDate = $OrderData.OrderDate }
    if ($OrderData.OrderRef)     { $postBody['yourReference']          = $OrderData.OrderRef }
    if ($OrderData.DeliveryDate) { $postBody['requestedDeliveryDate']  = $OrderData.DeliveryDate }

    $order = $null
    try {
        $order = Invoke-RestMethod -Method Post -Uri "$pipelineBase/salesOrderCreations" -Headers $jsonHeader `
            -Body ($postBody | ConvertTo-Json)
        Write-Host "    [OK] Order created      : $($order.number)" -ForegroundColor Green
        if ($OrderData.OrderRef)     { Write-Host "    [OK] Your Reference     : $($OrderData.OrderRef)" -ForegroundColor Green }
        if ($OrderData.DeliveryDate) { Write-Host "    [OK] Requested Delivery : $($OrderData.DeliveryDate)" -ForegroundColor Green }
    } catch {
        throw "POST sales order creation failed: $(Get-ApiError $_)"
    }
    $orderId = $order.systemId
    $orderNo = $order.number

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

    # Step 4: POST sales order lines. Individual line failures are tolerated — a bad line is
    # skipped, not fatal. The caller is told which lines were skipped so it can notify (amber).
    $lineOk      = 0
    $failedLines = [System.Collections.Generic.List[string]]::new()
    foreach ($line in $OrderData.Lines) {
        try {
            $result = Invoke-RestMethod -Method Post `
                -Uri "$apiBase/salesOrders($orderId)/salesOrderLines" -Headers $jsonHeader `
                -Body (@{ lineType = 'Item'; lineObjectNumber = $line.ItemNumber; quantity = $line.Quantity } | ConvertTo-Json)
            Write-Host ("    [OK] {0,-22} qty {1,6}  -> EUR {2}" -f $line.ItemNumber, $line.Quantity, $result.unitPrice) -ForegroundColor Green
            $lineOk++
        } catch {
            $lineErr = Get-ApiError $_
            Write-Host ("    [ERROR] Line {0} qty {1}: {2}" -f $line.ItemNumber, $line.Quantity, $lineErr) -ForegroundColor Red
            $failedLines.Add("$($line.ItemNumber)  qty $($line.Quantity)  —  $lineErr")
        }
    }

    # No line made it onto the order — roll back the empty header so the order is NOT left in BC,
    # and the next watcher run retries the PDF (e.g. once the missing items are created in BC).
    if ($lineOk -eq 0) {
        try {
            $reGet = Invoke-RestMethod -Uri "$apiBase/salesOrders($orderId)" -Headers $authHeader
            Invoke-RestMethod -Method Delete -Uri "$apiBase/salesOrders($orderId)" `
                -Headers ($authHeader + @{ 'If-Match' = $reGet.'@odata.etag' }) | Out-Null
            Write-Host "    [ROLLBACK] No lines added — empty order $orderNo deleted." -ForegroundColor Yellow
        } catch {
            Write-Host "    [WARN] Rollback delete failed: $(Get-ApiError $_)" -ForegroundColor Yellow
        }
        throw "No lines could be added to order $orderNo — order not posted. The PDF will be retried on the next run."
    }

    # Step 5: POST order comment (if fixedComment is set in template)
    if ($Template.PSObject.Properties['fixedComment'] -and $Template.fixedComment) {
        try {
            Invoke-RestMethod -Method Post -Uri "$pipelineBase/salesOrderComments" -Headers $jsonHeader `
                -Body (@{ documentNo = $orderNo; comment = $Template.fixedComment } | ConvertTo-Json) | Out-Null
            Write-Host "    [OK] Comment            : $($Template.fixedComment)" -ForegroundColor Green
        } catch {
            Write-Host "    [WARN] Comment POST failed: $(Get-ApiError $_)" -ForegroundColor Yellow
        }
    }

    # Step 6: Attach PDF to the BC sales order (Factbox → Attachments → Documents)
    if ($null -ne $PdfBytes -and $PdfBytes.Length -gt 0 -and $PdfFileName) {
        try {
            $attMeta = Invoke-RestMethod -Method Post `
                -Uri "$apiBase/salesOrders($orderId)/documentAttachments" `
                -Headers $jsonHeader `
                -Body (@{ fileName = $PdfFileName } | ConvertTo-Json)
            $attId  = $attMeta.id
            $attGet = Invoke-RestMethod -Uri "$apiBase/salesOrders($orderId)/documentAttachments($attId)" `
                -Headers $authHeader
            Invoke-RestMethod -Method Put `
                -Uri "$apiBase/salesOrders($orderId)/documentAttachments($attId)/attachmentContent" `
                -Headers ($authHeader + @{ 'Content-Type' = 'application/octet-stream'; 'If-Match' = $attGet.'@odata.etag' }) `
                -Body $PdfBytes | Out-Null
            Write-Host "    [OK] PDF attached        : $PdfFileName" -ForegroundColor Green
        } catch {
            $pdfErr = Get-ApiError $_
            Write-Host "    [WARN] PDF attachment failed: $pdfErr" -ForegroundColor Yellow
            # Notify supcom + x.planchette — order is in BC but PDF must be attached manually
            if (Get-Command Send-NotificationEmail -ErrorAction SilentlyContinue) {
                $pdfBody = Build-HtmlShell -Title 'PDF Not Attached' `
                    -Subtitle "Order $orderNo was posted to BC but the source PDF could not be attached automatically." `
                    -Body (
                        (Build-InfoBox ([ordered]@{ 'Order' = $orderNo; 'Client' = $Template.clientName; 'File' = $PdfFileName })) +
                        (Build-AlertBox "Please attach <strong>$PdfFileName</strong> manually in BC: open order <strong>$orderNo</strong> &rarr; Factbox &rarr; Attachments &rarr; Documents.<br><br><strong>Error:</strong> $([System.Net.WebUtility]::HtmlEncode($pdfErr))")
                    )
                Send-NotificationEmail `
                    -Subject     "[Sales Order] PDF not attached — order $orderNo" `
                    -AlsoNotify  @('x.planchette@montandor.com') `
                    -Body        $pdfBody `
                    -ContentType 'HTML'
            }
        }
    }

    return [PSCustomObject]@{ OrderNo = $orderNo; SkippedLines = $failedLines }
}

# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------
if ($FunctionsOnly) { return }   # dot-sourced by Watch-SalesOrderEmail.ps1 — stop here

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
                    $pdfFileBytes = [System.IO.File]::ReadAllBytes($pdf.FullName)
                    $result = Submit-SalesOrder -OrderData $data -Template $tpl -ShipToCode $shipToCode -PdfBytes $pdfFileBytes -PdfFileName $pdf.Name
                    $bcNo   = $result.OrderNo
                    Move-Item -Path $pdf.FullName -Destination (Join-Path $procDir $pdf.Name) -Force
                    Write-Host "    -> BC $bcNo | moved to Processed" -ForegroundColor Green
                    if (@($result.SkippedLines).Count -gt 0) {
                        Write-Host ("    [NOTIFY] {0} line(s) skipped: {1}" -f @($result.SkippedLines).Count, (@($result.SkippedLines) -join '; ')) -ForegroundColor Yellow
                    }
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
