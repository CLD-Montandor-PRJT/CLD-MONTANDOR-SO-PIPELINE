<#
.SYNOPSIS
    Offline, credential-free test suite for the non-PDF-attachment-handling feature
    (Sales Order PDF pipeline). Plain-assertion runner (no Pester — only Pester 3.4.0 is
    installed on this machine, which predates the 5.x idioms this pipeline's other test
    suites use; a plain runner avoids syntax-mismatch risk without adding a dependency).
.DESCRIPTION
    Exercises, in complete isolation from BC/Graph/credentials:
      - Get-OfficeDocText (new): .odt/.docx -> plain text
      - Get-TextModeOrderData (extracted from Get-PdfOrderData's text-mode branch):
        the qtyUnit whitespace-tolerance + substring-collision-guard fix
      - End-to-end extraction of the real PO2607-000071.odt fixture
      - Golden-master regression: unmodified (.orig.ps1) vs modified (.ps1) Get-PdfOrderData
        against 3 real GAFIC PDFs — must be byte-for-byte identical
      - Get-NonPdfAttachmentAction (new): the no-PDF routing decision
      - Build-UnsupportedAttachmentHtml + Send-UnsupportedAttachmentNotification (new):
        notification payload, with Send-NotificationEmail stubbed (never sends)
    Exits 0 if all assertions pass, 1 otherwise.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$devRoot = Split-Path $PSScriptRoot -Parent
$script:passed = 0
$script:failed = 0
$script:failMessages = @()

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if ($Condition) {
        $script:passed++
        Write-Output "  [PASS] $Message"
    } else {
        $script:failed++
        $script:failMessages += $Message
        Write-Output "  [FAIL] $Message"
    }
}

function Assert-Equal {
    param($Expected, $Actual, [string]$Message)
    $ok = ($Expected -eq $Actual)
    Assert-True $ok "$Message (expected=[$Expected] actual=[$Actual])"
}

function Assert-Null {
    param($Value, [string]$Message)
    Assert-True ($null -eq $Value) "$Message (actual=[$Value])"
}

function Assert-NotNull {
    param($Value, [string]$Message)
    Assert-True ($null -ne $Value) $Message
}

function ConvertTo-CanonicalJson {
    param($Object)
    return ($Object | ConvertTo-Json -Depth 15 -Compress)
}

function Assert-DeepEqual {
    param($Expected, $Actual, [string]$Message)
    $e = ConvertTo-CanonicalJson $Expected
    $a = ConvertTo-CanonicalJson $Actual
    Assert-True ($e -eq $a) $Message
    if ($e -ne $a) {
        Write-Output "    expected: $e"
        Write-Output "    actual  : $a"
    }
}

# ---------------------------------------------------------------------------
# Harness setup — load PdfPig (local DLLs, offline) and extract pure functions
# from the pipeline scripts via AST (no top-level credential/network code runs).
# ---------------------------------------------------------------------------
Write-Output "`n=== Harness setup ==="

$dllDir = Join-Path $devRoot 'lib\dlls'
foreach ($dll in @(
    'UglyToad.PdfPig.Core.dll', 'UglyToad.PdfPig.Tokens.dll',
    'UglyToad.PdfPig.Tokenization.dll', 'UglyToad.PdfPig.Fonts.dll',
    'UglyToad.PdfPig.dll'
)) {
    Add-Type -Path (Join-Path $dllDir $dll) -ErrorAction SilentlyContinue
}
Write-Output "  PdfPig loaded from $dllDir"

$origPipeline = Join-Path $devRoot 'Invoke-SalesOrderPipeline.orig.ps1'
$newPipeline  = Join-Path $devRoot 'Invoke-SalesOrderPipeline.ps1'
$newWatcher   = Join-Path $devRoot 'Watch-SalesOrderEmail.ps1'

# Load the OLD (pristine, unmodified) Get-PdfOrderData under an aliased name so both old and
# new implementations can be called side-by-side in the same process for the golden-master test.
$oldFuncText = (
    [System.Management.Automation.Language.Parser]::ParseFile($origPipeline, [ref]$null, [ref]$null).FindAll(
        { param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-PdfOrderData' },
        $true
    ) | Select-Object -First 1
).Extent.Text
$oldFuncText = $oldFuncText -replace '^function Get-PdfOrderData', 'function Get-PdfOrderData-OLD'
. ([scriptblock]::Create($oldFuncText))
Write-Output "  Loaded pristine Get-PdfOrderData as Get-PdfOrderData-OLD (from Invoke-SalesOrderPipeline.orig.ps1)"

# Load the NEW (modified) functions under their real names.
. (Join-Path $PSScriptRoot 'Import-PipelineFunctions.ps1') -ScriptPath $newPipeline -FunctionNames @(
    'Get-PdfOrderData', 'Get-TextModeOrderData', 'Get-OfficeDocText'
)
Write-Output "  Loaded modified Get-PdfOrderData, Get-TextModeOrderData, Get-OfficeDocText (from Invoke-SalesOrderPipeline.ps1)"

# Load the NEW routing/notification functions from Watch-SalesOrderEmail.ps1.
. (Join-Path $PSScriptRoot 'Import-PipelineFunctions.ps1') -ScriptPath $newWatcher -FunctionNames @(
    'Get-NonPdfAttachmentAction', 'Build-UnsupportedAttachmentHtml', 'Send-UnsupportedAttachmentNotification',
    'Build-InfoBox', 'Build-AlertBox', 'Build-HtmlShell'
)
Write-Output "  Loaded Get-NonPdfAttachmentAction, Build-UnsupportedAttachmentHtml, Send-UnsupportedAttachmentNotification (from Watch-SalesOrderEmail.ps1)"

$gaficTemplate    = Get-Content (Join-Path $devRoot 'Gafic\template.json') -Raw | ConvertFrom-Json
$gaficKnownItems  = @(Get-Content (Join-Path $PSScriptRoot 'fixtures\gafic-known-items.json') -Raw | ConvertFrom-Json)
$odtFixturePath   = Join-Path $PSScriptRoot 'fixtures\PO2607-000071.odt'
$odtBytes         = [System.IO.File]::ReadAllBytes($odtFixturePath)

# ===========================================================================
Write-Output "`n=== Group 1: Get-OfficeDocText ==="
# ===========================================================================

$r = Get-OfficeDocText -FileBytes ([byte[]]@(1,2,3,4)) -FileName 'invoice.pdf'
Assert-Null $r "Unsupported extension (.pdf) returns `$null"

$r = Get-OfficeDocText -FileBytes ([byte[]]@(1,2,3,4)) -FileName 'invoice.doc'
Assert-Null $r "Unsupported extension (.doc) returns `$null"

$r = Get-OfficeDocText -FileBytes ([byte[]]@(0xDE,0xAD,0xBE,0xEF,1,2,3)) -FileName 'PO123.odt'
Assert-Null $r "Corrupt/non-ZIP .odt bytes return `$null (no throw)"

$r = Get-OfficeDocText -FileBytes $odtBytes -FileName 'PO2607-000071_Commande Fournisseur.odt'
Assert-NotNull $r "Real PO2607-000071.odt: extraction returns non-null text"
Assert-True ($r -match 'PO2607-000071') "Real .odt text contains order ref PO2607-000071"
Assert-True ($r -match 'BL-SMA510-RD') "Real .odt text contains code BL-SMA510-RD"
Assert-True ($r -match 'ELE-B-LA') "Real .odt text contains code ELE-B-LA"

# ===========================================================================
Write-Output "`n=== Group 2: Get-TextModeOrderData — qtyUnit whitespace + boundary-guard fix ==="
# ===========================================================================

$mapPathNone = Join-Path $devRoot 'Gafic\__no_mapping_here__'  # deliberately missing -> Test-Path false

# 2a. Zero-space (PdfPig-glued) pattern — must still match (no regression on live PDF path)
$textZeroSpace = 'VERMES#BL-SMA510-RD1PCE5.95'
$d = Get-TextModeOrderData -AllText $textZeroSpace -Template $gaficTemplate `
        -BcItemNumbers @('BL-SMA510-RD') -MapPath $mapPathNone -OrderRef 'T1' -OrderDateBC '2026-01-01'
Assert-Equal 1 (@($d.Lines).Count) "Zero-space pattern: 1 line extracted"
if (@($d.Lines).Count -eq 1) {
    Assert-Equal 'BL-SMA510-RD' $d.Lines[0].ItemNumber "Zero-space pattern: correct item number"
    Assert-Equal 1 ([int]$d.Lines[0].Quantity) "Zero-space pattern: correct qty"
}

# 2b. Real-space (office-doc) pattern — the whitespace-tolerance half of the fix
$textSpaced = 'VERMES#BL-SMA510-RD 1 PCE 5.95'
$d = Get-TextModeOrderData -AllText $textSpaced -Template $gaficTemplate `
        -BcItemNumbers @('BL-SMA510-RD') -MapPath $mapPathNone -OrderRef 'T2' -OrderDateBC '2026-01-01'
Assert-Equal 1 (@($d.Lines).Count) "Spaced (office-doc-style) pattern: 1 line extracted"
if (@($d.Lines).Count -eq 1) {
    Assert-Equal 'BL-SMA510-RD' $d.Lines[0].ItemNumber "Spaced pattern: correct item number"
    Assert-Equal 1 ([int]$d.Lines[0].Quantity) "Spaced pattern: correct qty"
}

# 2c. Substring collision — the boundary-guard half of the fix. Both 'SMA510-RD' and
# 'BL-SMA510-RD' are real, distinct BC items (confirmed: SMA510-RD appears standalone in
# real GAFIC PDF PO2606-031876; BL-SMA510-RD in PO2607-000071). Only BL-SMA510-RD's own row
# should produce a line — 'SMA510-RD' must NOT phantom-match as a substring of it.
$textCollision = 'VERMES#BL-SMA510-RD 1 PCE 5.95'
$d = Get-TextModeOrderData -AllText $textCollision -Template $gaficTemplate `
        -BcItemNumbers @('BL-SMA510-RD','SMA510-RD') -MapPath $mapPathNone -OrderRef 'T3' -OrderDateBC '2026-01-01'
Assert-Equal 1 (@($d.Lines).Count) "Collision guard: exactly 1 line (no phantom SMA510-RD)"
Assert-True (-not (@($d.Lines) | Where-Object { $_.ItemNumber -eq 'SMA510-RD' })) "Collision guard: no line for bare SMA510-RD"
if (@($d.Lines).Count -ge 1) {
    Assert-True ((@($d.Lines) | Where-Object { $_.ItemNumber -eq 'BL-SMA510-RD' }).Count -eq 1) "Collision guard: BL-SMA510-RD line present exactly once"
}

# 2d. Reverse order in BcItemNumbers must not change the outcome (dictionary iteration order
# should not matter — guards against an implementation that only works by accident of ordering)
$d2 = Get-TextModeOrderData -AllText $textCollision -Template $gaficTemplate `
        -BcItemNumbers @('SMA510-RD','BL-SMA510-RD') -MapPath $mapPathNone -OrderRef 'T3b' -OrderDateBC '2026-01-01'
Assert-Equal 1 (@($d2.Lines).Count) "Collision guard (reversed item order): still exactly 1 line"

# 2e. A legitimately shorter code that is genuinely NOT part of a longer code in the text must
# still match normally (guard must not be over-broad and block real short-code matches)
$textShortOnly = 'VERMES#SMA510-RD 17 PCE 1.77'
$d3 = Get-TextModeOrderData -AllText $textShortOnly -Template $gaficTemplate `
        -BcItemNumbers @('SMA510-RD','BL-SMA510-RD') -MapPath $mapPathNone -OrderRef 'T4' -OrderDateBC '2026-01-01'
Assert-Equal 1 (@($d3.Lines).Count) "Standalone short code (no collision context): 1 line"
if (@($d3.Lines).Count -eq 1) {
    Assert-Equal 'SMA510-RD' $d3.Lines[0].ItemNumber "Standalone short code: correct item number"
    Assert-Equal 17 ([int]$d3.Lines[0].Quantity) "Standalone short code: correct qty"
}

# ===========================================================================
Write-Output "`n=== Group 3: End-to-end ODT extraction (PO2607-000071) ==="
# ===========================================================================

$odtText = Get-OfficeDocText -FileBytes $odtBytes -FileName 'PO2607-000071_Commande Fournisseur.odt'
$mapPathGafic = Join-Path $devRoot 'Gafic\mapping.json'   # does not exist for GAFIC — fine, Test-Path handles it

# OrderRef/OrderDate normally computed by Get-PdfOrderData before the text/coordinate branch;
# replicate the same regex against the office-doc text here to feed Get-TextModeOrderData directly.
$orderRef = ''
if ($odtText -match $gaficTemplate.orderNumberRegex) { $orderRef = $Matches[1] }
$orderDateBC = ''
if ($odtText -match $gaficTemplate.dateRegex) {
    $parsed = [DateTime]::ParseExact($Matches[1], $gaficTemplate.dateInputFormat, [System.Globalization.CultureInfo]::InvariantCulture)
    $orderDateBC = $parsed.ToString('yyyy-MM-dd')
}

$odtData = Get-TextModeOrderData -AllText $odtText -Template $gaficTemplate `
    -BcItemNumbers $gaficKnownItems -MapPath $mapPathGafic -OrderRef $orderRef -OrderDateBC $orderDateBC

Assert-Equal 'PO2607-000071' $odtData.OrderRef "ODT: OrderRef extracted correctly"
Assert-Equal 9 (@($odtData.Lines).Count) "ODT: exactly 9 lines extracted"
Assert-Equal 0 (@($odtData.UnknownCodes).Count) "ODT: 0 unknown codes"
Assert-True (-not (@($odtData.Lines) | Where-Object { $_.ItemNumber -eq 'SMA510-RD' })) "ODT: no phantom SMA510-RD line"

$expectedOdtLines = [ordered]@{
    'BL-SMA510-RD' = 1; 'ELE-B-LA' = 5; 'ELE-B-ME' = 4; 'ELE-TE-LA' = 6
    'MC-CRWC-BL'   = 10; 'MCS-6A4-LSS' = 1; 'SECCLEAN-KL' = 2
    'TT-DB-XL'     = 6; 'WBW-TE-40-60' = 3
}
foreach ($kv in $expectedOdtLines.GetEnumerator()) {
    $line = @($odtData.Lines) | Where-Object { $_.ItemNumber -eq $kv.Key }
    Assert-True ($null -ne $line) "ODT: line present for $($kv.Key)"
    if ($line) { Assert-Equal $kv.Value ([int]$line.Quantity) "ODT: qty correct for $($kv.Key)" }
}

# ===========================================================================
Write-Output "`n=== Group 4: Golden-master regression — real GAFIC PDFs, OLD vs NEW ==="
# ===========================================================================

$gaficDir = Join-Path $devRoot 'Gafic'
# PO2604/PO2605 expected counts are independently confirmed (prior session record: 15 / 13
# lines respectively). PO2606's count (36) is derived by manual cross-reference of this
# session's own VERMES#-anchored discovery against the gafic-known-items.json fixture, not
# from an independent prior record — flagged in the handoff report as a lower-confidence check.
$pdfCases = @(
    @{ File = 'PO2604-027627_Commande Fournisseur.pdf'; ExpectedLines = 15; Confirmed = $true }
    @{ File = 'PO2605-030215_Commande Fournisseur.pdf'; ExpectedLines = 13; Confirmed = $true }
    @{ File = 'PO2606-031876_Commande Fournisseur.pdf'; ExpectedLines = 36; Confirmed = $false }
)

foreach ($case in $pdfCases) {
    $pdfPath = Join-Path $gaficDir $case.File
    $oldData = Get-PdfOrderData-OLD -PdfPath $pdfPath -Template $gaficTemplate -BcItemNumbers $gaficKnownItems -ClientDir $gaficDir
    $newData = Get-PdfOrderData     -PdfPath $pdfPath -Template $gaficTemplate -BcItemNumbers $gaficKnownItems -ClientDir $gaficDir

    $confidence = if ($case.Confirmed) { 'confirmed baseline' } else { 'derived, not independently confirmed' }
    Assert-Equal $case.ExpectedLines (@($oldData.Lines).Count) "$($case.File): OLD code line count matches $confidence"
    Assert-DeepEqual $oldData $newData "$($case.File): NEW code output is byte-for-byte identical to OLD (golden master)"
}

# ===========================================================================
Write-Output "`n=== Group 5: Get-NonPdfAttachmentAction — routing decision ==="
# ===========================================================================

$odtAtt  = [PSCustomObject]@{ name = 'PO123.odt';  contentType = 'application/vnd.oasis.opendocument.text'; isInline = $false }
$docxAtt = [PSCustomObject]@{ name = 'PO123.docx'; contentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'; isInline = $false }
$docAtt  = [PSCustomObject]@{ name = 'PO123.doc';  contentType = 'application/msword'; isInline = $false }
$rtfAtt  = [PSCustomObject]@{ name = 'PO123.rtf';  contentType = 'application/rtf'; isInline = $false }
$imgAtt  = [PSCustomObject]@{ name = 'scan.jpg';   contentType = 'image/jpeg'; isInline = $false }
$inlineOdtAtt = [PSCustomObject]@{ name = 'sig.odt'; contentType = 'application/vnd.oasis.opendocument.text'; isInline = $true }

$route = Get-NonPdfAttachmentAction -IsTextMode $true -Attachments @($odtAtt)
Assert-Equal 'ExtractOfficeDoc' $route.Action "Text-mode + .odt attachment -> ExtractOfficeDoc"
Assert-Equal 'PO123.odt' $route.Attachment.name "Text-mode + .odt attachment -> correct attachment chosen"

$route = Get-NonPdfAttachmentAction -IsTextMode $true -Attachments @($docxAtt)
Assert-Equal 'ExtractOfficeDoc' $route.Action "Text-mode + .docx attachment -> ExtractOfficeDoc"

$route = Get-NonPdfAttachmentAction -IsTextMode $false -Attachments @($odtAtt)
Assert-Equal 'NotifyUnsupported' $route.Action "Coordinate-mode + .odt attachment -> NotifyUnsupported (coordinate mode never attempts office-doc extraction)"

$route = Get-NonPdfAttachmentAction -IsTextMode $true -Attachments @($docAtt, $rtfAtt, $imgAtt)
Assert-Equal 'NotifyUnsupported' $route.Action "Text-mode + only .doc/.rtf/image attachments -> NotifyUnsupported"

$route = Get-NonPdfAttachmentAction -IsTextMode $true -Attachments @()
Assert-Equal 'NotifyUnsupported' $route.Action "Text-mode + zero attachments -> NotifyUnsupported"

$route = Get-NonPdfAttachmentAction -IsTextMode $true -Attachments @($inlineOdtAtt)
Assert-Equal 'NotifyUnsupported' $route.Action "Text-mode + inline-only .odt (e.g. embedded signature image) -> NotifyUnsupported, inline attachments are not candidates"

# ===========================================================================
Write-Output "`n=== Group 6: Notification payload — Build-UnsupportedAttachmentHtml + Send-UnsupportedAttachmentNotification ==="
# ===========================================================================

$html = Build-UnsupportedAttachmentHtml -ClientName 'GAFIC' -SenderEmail 'orders@gafic1965.com' `
    -EmailSubject 'Commande PO2607-000071' -FileNames @('PO2607-000071.odt') -ContentTypes @('application/vnd.oasis.opendocument.text')

Assert-True ($html -match [regex]::Escape('GAFIC')) "Notification HTML contains client/customer name"
Assert-True ($html -match [regex]::Escape('orders@gafic1965.com')) "Notification HTML contains sender email"
Assert-True ($html -match [regex]::Escape('Commande PO2607-000071')) "Notification HTML contains email subject"
Assert-True ($html -match [regex]::Escape('PO2607-000071.odt')) "Notification HTML contains file name"
Assert-True ($html -match [regex]::Escape('application/vnd.oasis.opendocument.text')) "Notification HTML contains content-type"

# Stub Send-NotificationEmail (redefine in this scope — later definition wins for subsequent
# calls). Captures the call instead of hitting Graph. Never sends a real email.
$script:capturedCalls = @()
function Send-NotificationEmail {
    param([string]$Subject, [string]$Body, [string[]]$AlsoNotify = @(), [string]$ContentType = 'Text')
    $script:capturedCalls += [PSCustomObject]@{ Subject = $Subject; Body = $Body; AlsoNotify = $AlsoNotify; ContentType = $ContentType }
}

Send-UnsupportedAttachmentNotification -ClientName 'GAFIC' -SenderEmail 'orders@gafic1965.com' `
    -EmailSubject 'Commande PO2607-000071' -FileNames @('PO2607-000071.odt') -ContentTypes @('application/vnd.oasis.opendocument.text')

Assert-Equal 1 (@($script:capturedCalls).Count) "Send-UnsupportedAttachmentNotification calls Send-NotificationEmail exactly once (stubbed — no real send)"
if (@($script:capturedCalls).Count -eq 1) {
    $call = $script:capturedCalls[0]
    Assert-Equal '[Sales Order] Unsupported attachment — manual entry needed' $call.Subject "Captured call: exact subject line"
    Assert-True (@($call.AlsoNotify) -contains 'x.planchette@montandor.com') "Captured call: AlsoNotify includes x.planchette@montandor.com"
    Assert-Equal 'HTML' $call.ContentType "Captured call: HTML content type"
    Assert-True ($call.Body -match [regex]::Escape('PO2607-000071.odt')) "Captured call: body contains file name"
}

# ===========================================================================
Write-Output "`n=== SUMMARY ==="
Write-Output "  Passed: $script:passed"
Write-Output "  Failed: $script:failed"
if ($script:failed -gt 0) {
    Write-Output "`n  Failing assertions:"
    $script:failMessages | ForEach-Object { Write-Output "    - $_" }
    exit 1
}
exit 0
