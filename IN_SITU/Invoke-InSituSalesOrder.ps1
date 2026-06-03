<#
.SYNOPSIS
    Creates a BC Production sales order from an In Situ PDF purchase order
    and moves the source file to the Processed folder on success.
.DESCRIPTION
    Reads credentials from Windows Credential Manager, authenticates via OAuth 2.0,
    POSTs the sales order header and 5 lines via BC API v2.0, then PATCHes
    Your_Reference via the OData SalesOrder endpoint (the only surface that exposes it),
    and finally moves the PDF from Not_Processed to Processed.
.NOTES
    Customer  : In Situ (C0000065)
    PO ref    : 93966  -> Your Reference (set via OData PATCH after header creation)
    Order date: 15/05/2026
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# 0. Credentials
# ---------------------------------------------------------------------------
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class WinCredIS {
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

$tenantId     = [WinCredIS]::GetSecret('Montandor_BC_TenantId')
$clientId     = [WinCredIS]::GetSecret('Montandor_BC_ClientId')
$clientSecret = [WinCredIS]::GetSecret('Montandor_BC_ClientSecret')

# ---------------------------------------------------------------------------
# 1. Config
# ---------------------------------------------------------------------------
$environment = 'Production'
$companyId   = '4e422ae7-867a-ef11-a671-000d3a45ce6c'
$apiBase     = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$environment/api/v2.0/companies($companyId)"
$odataBase   = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$environment/ODataV4/Company('Montandor_Andorra')"

$pdfSource   = 'C:\Sales_Order_BC_Input\IN_SITU\Not_Processed\OCDF14A_292755_1.PDF'
$pdfDest     = 'C:\Sales_Order_BC_Input\IN_SITU\Processed\OCDF14A_292755_1.PDF'

# ---------------------------------------------------------------------------
# 2. OAuth token
# ---------------------------------------------------------------------------
Write-Host '[1/5] Acquiring OAuth token...' -ForegroundColor Cyan
$token = Invoke-RestMethod -Method Post `
    -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
    -Body @{
        grant_type    = 'client_credentials'
        client_id     = $clientId
        client_secret = $clientSecret
        scope         = 'https://api.businesscentral.dynamics.com/.default'
    } `
    -ContentType 'application/x-www-form-urlencoded'

$authHeader = @{ Authorization = "Bearer $($token.access_token)" }
$jsonHeader = $authHeader + @{ 'Content-Type' = 'application/json' }
Write-Host "  Token acquired. Expires in $($token.expires_in)s." -ForegroundColor Green

# ---------------------------------------------------------------------------
# 3. POST sales order header
# ---------------------------------------------------------------------------
Write-Host '[2/5] Creating sales order header...' -ForegroundColor Cyan

$order = Invoke-RestMethod -Method Post `
    -Uri "$apiBase/salesOrders" `
    -Headers $jsonHeader `
    -Body (@{ customerNumber = 'C0000065'; orderDate = '2026-05-15' } | ConvertTo-Json)

$orderId     = $order.id
$orderNumber = $order.number
Write-Host "  Sales order created: $orderNumber (id: $orderId)" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 4. PATCH Your_Reference via OData (not exposed in API v2.0)
# ---------------------------------------------------------------------------
Write-Host '[3/5] Setting Your Reference via OData...' -ForegroundColor Cyan

# Fetch the OData record to get its etag
$odataOrder = Invoke-RestMethod -Method Get `
    -Uri "$odataBase/SalesOrder(Document_Type='Order',No='$orderNumber')" `
    -Headers $authHeader

$patchHeader = $authHeader + @{
    'Content-Type' = 'application/json'
    'If-Match'     = $odataOrder.'@odata.etag'
}

Invoke-RestMethod -Method Patch `
    -Uri "$odataBase/SalesOrder(Document_Type='Order',No='$orderNumber')" `
    -Headers $patchHeader `
    -Body (@{ Your_Reference = '93966' } | ConvertTo-Json) | Out-Null

Write-Host "  Your Reference set to: 93966" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 5. POST sales order lines
# ---------------------------------------------------------------------------
Write-Host '[4/5] Adding order lines...' -ForegroundColor Cyan

$lines = @(
    [pscustomobject]@{ itemNumber = 'SMA720-WT';        quantity = 60 }
    [pscustomobject]@{ itemNumber = 'SBD-BL-85';        quantity = 3  }
    [pscustomobject]@{ itemNumber = 'ELE-BL-LA';        quantity = 30 }
    [pscustomobject]@{ itemNumber = 'RS-RT-RVS-RD-SET'; quantity = 12 }
    [pscustomobject]@{ itemNumber = 'EZL-TE-165';       quantity = 2  }
)

foreach ($line in $lines) {
    $result = Invoke-RestMethod -Method Post `
        -Uri "$apiBase/salesOrders($orderId)/salesOrderLines" `
        -Headers $jsonHeader `
        -Body (@{ lineType = 'Item'; lineObjectNumber = $line.itemNumber; quantity = $line.quantity } | ConvertTo-Json)

    Write-Host ("  {0,-20} qty {1,3}  ->  €{2}" -f $line.itemNumber, $line.quantity, $result.unitPrice) -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# 6. Move PDF to Processed
# ---------------------------------------------------------------------------
Write-Host '[5/5] Moving PDF to Processed...' -ForegroundColor Cyan
Move-Item -Path $pdfSource -Destination $pdfDest -Force
Write-Host "  $([System.IO.Path]::GetFileName($pdfSource)) -> Processed" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------
Write-Host "`nDone. BC Sales Order $orderNumber | Your Reference 93966 | In Situ | Open for review." -ForegroundColor Green

$clientSecret = $null
$token        = $null
