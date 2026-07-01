param([string]$PdfPath)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$dllDir = "$PSScriptRoot\..\lib\dlls"
foreach ($dll in @(
    'UglyToad.PdfPig.Core.dll', 'UglyToad.PdfPig.Tokens.dll',
    'UglyToad.PdfPig.Tokenization.dll', 'UglyToad.PdfPig.Fonts.dll',
    'UglyToad.PdfPig.dll'
)) {
    Add-Type -Path "$dllDir\$dll" -ErrorAction SilentlyContinue
}
$pdf = [UglyToad.PdfPig.PdfDocument]::Open($PdfPath)
$allText = ($pdf.GetPages() | ForEach-Object { $_.Text }) -join ' ' -replace '[\x00-\x1F]', ' '
Write-Output "=== allText (first 3000 chars) ==="
Write-Output $allText.Substring(0, [Math]::Min(3000, $allText.Length))
Write-Output "`n=== VERMES# occurrences ==="
[regex]::Matches($allText, 'VERMES#[A-Za-z0-9\-]+\d*(?:PCE|PQT)?') | ForEach-Object { Write-Output $_.Value }
$pdf.Dispose()
