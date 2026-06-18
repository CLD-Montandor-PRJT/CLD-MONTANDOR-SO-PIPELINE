<#
.SYNOPSIS
    Files Vermes (G. Vermes B.V.) invoice PDFs against the matching Business Central
    purchase order — no PO is created or modified, the PDF is only attached.
.DESCRIPTION
    Scans administration@montandor.com for emails from finance.dept@vermes.nl with a PDF
    attachment, received within a rolling lookback window. For each PDF it extracts the
    "Your Ref.:" value (the BC purchase order number) from the delivery-address box, matches
    the BC purchase order by number, and attaches the PDF to that PO's documentAttachments —
    only if an attachment with the same filename is not already present (idempotent).

    Unmatched invoices are retried silently each run; a single amber alert is sent once an
    invoice is still unmatched after $alertAfterDays (tracked in a local ledger so it fires
    exactly once). Unreadable / ambiguous / wrong-vendor cases alert once immediately.

    Reuses the SO pipeline's OAuth tokens, PdfPig and BC helpers via -FunctionsOnly dot-source.
    Reads BC writes ONLY the documentAttachments of an existing PO. Email is never modified.
.PARAMETER DryRun
    Extract + match + decide, but do NOT attach to BC and do NOT send alert emails — just log
    what WOULD happen. Used for verification.
.PARAMETER NoAlert
    Attach matched/unattached PDFs but suppress all alert emails. Used for the one-time backlog
    sweep so old unmatched invoices don't trigger a burst of alerts (and don't poison the ledger).
.PARAMETER LookbackDays
    Override the config lookback window (0 = use config). Used for the backlog sweep.
.NOTES
    Requires Mail.Read + Mail.Send (administration@ mailbox) + BC API write on documentAttachments.
    Credentials in Windows Credential Manager: Montandor_BC_TenantId/ClientId/ClientSecret.
    Schedule: Task Scheduler at 09:00, 12:00, 15:00, 18:00.
#>
[CmdletBinding()]
param(
    [switch]$DryRun,
    [switch]$NoAlert,
    [int]$LookbackDays = 0   # 0 = use config value
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Single-instance guard (named mutex; OS releases on process exit) ---
$mutex = [System.Threading.Mutex]::new($false, 'Montandor-VermesInvoiceFiling')
$haveLock = $false
try { $haveLock = $mutex.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $haveLock = $true }
if (-not $haveLock) { Write-Host '[EXIT] Another instance is already running.' -ForegroundColor Yellow; return }

try {
    # Capture our own switches BEFORE dot-sourcing: the pipeline runs `$DryRun = -not $Execute`
    # at top level and dot-sourcing shares scope, which would otherwise clobber our -DryRun.
    $wantDryRun = [bool]$DryRun
    # Load pipeline functions + DLLs + OAuth tokens ($graphHeader, $authHeader, $jsonHeader, PdfPig)
    . (Join-Path $PSScriptRoot 'Invoke-SalesOrderPipeline.ps1') -FunctionsOnly
    $FunctionsOnly = $false
    $DryRun = $wantDryRun   # restore our value (pipeline overwrote $DryRun during dot-source)

    $cfg        = Get-Content (Join-Path $PSScriptRoot 'VermesInvoiceFiling.config.json') | ConvertFrom-Json
    $graphBase  = "https://graph.microsoft.com/v1.0/users/$($cfg.mailbox)"
    $apiBase    = "https://api.businesscentral.dynamics.com/v2.0/$tenantId/$($cfg.environment)/api/v2.0/companies($companyId)"
    $ledgerPath = Join-Path $PSScriptRoot 'vermes-alerted.json'
    $modeTag    = if ($DryRun) { '[DRY-RUN] ' } else { '' }
    Write-Host "`n${modeTag}[VERMES] Filing invoices from $($cfg.sender) into BC POs ($($cfg.environment))" -ForegroundColor Cyan

    # --- Alert ledger (fire-once) ---
    $ledger = @{}
    if (Test-Path $ledgerPath) {
        try { (Get-Content $ledgerPath -Raw | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $ledger[$_.Name] = $_.Value } } catch { $ledger = @{} }
    }
    function Save-Ledger {
        # prune entries older than 30 days, then persist
        $cut = [DateTime]::UtcNow.AddDays(-30)
        $keep = @{}
        foreach ($k in $ledger.Keys) {
            $ts = $null; try { $ts = [DateTime]::Parse($ledger[$k].alertedAt) } catch {}
            if ($ts -and $ts -ge $cut) { $keep[$k] = $ledger[$k] }
        }
        ($keep | ConvertTo-Json -Depth 4) | Set-Content $ledgerPath -Encoding UTF8
        $script:ledger = $keep
    }

    # --- Helpers ---------------------------------------------------------------
    function Get-FromAddr($m) { try { if ($m.from) { $m.from.emailAddress.address } else { '' } } catch { '' } }

    # Reconstruct visual text lines (group words by rounded Y, sort by X) and pull "Your Ref.:".
    # PdfPig's raw page text interleaves columns, so a regex over it is unreliable; line
    # reconstruction restores the on-page row so the value next to the label is captured.
    function Get-YourRef([byte[]]$Bytes, [string]$Pattern) {
        $doc = [UglyToad.PdfPig.PdfDocument]::Open($Bytes)
        try {
            foreach ($p in $doc.GetPages()) {
                $byY = @{}
                foreach ($w in $p.GetWords()) {
                    $y = [Math]::Round($w.BoundingBox.Bottom, 0)
                    if (-not $byY.ContainsKey($y)) { $byY[$y] = [System.Collections.Generic.List[object]]::new() }
                    $byY[$y].Add([PSCustomObject]@{ X = $w.BoundingBox.Left; T = $w.Text })
                }
                foreach ($y in $byY.Keys) {
                    $line = (($byY[$y] | Sort-Object X) | ForEach-Object { $_.T }) -join ' '
                    if ($line -match $Pattern) { return $Matches[1] }
                }
            }
        } finally { $doc.Dispose() }
        return $null
    }

    # Match a BC purchase order by its number; returns the PO object (with vendorNumber) or $null.
    function Find-PurchaseOrder([string]$Ref) {
        $safe = $Ref -replace "'", "''"
        $r = Invoke-RestMethod -Uri "$apiBase/purchaseOrders?`$filter=number eq '$safe'&`$select=id,number,vendorNumber" -Headers $authHeader
        return @($r.value)
    }

    function Get-PoAttachmentNames([string]$PoId) {
        $r = Invoke-RestMethod -Uri "$apiBase/purchaseOrders($PoId)/documentAttachments?`$select=fileName" -Headers $authHeader
        return @($r.value | ForEach-Object { $_.fileName })
    }

    function Add-PdfAttachment([string]$PoId, [string]$FileName, [byte[]]$Bytes) {
        $meta = Invoke-RestMethod -Method Post -Uri "$apiBase/purchaseOrders($PoId)/documentAttachments" `
            -Headers $jsonHeader -Body (@{ fileName = $FileName } | ConvertTo-Json)
        $get = Invoke-RestMethod -Uri "$apiBase/purchaseOrders($PoId)/documentAttachments($($meta.id))" -Headers $authHeader
        Invoke-RestMethod -Method Put -Uri "$apiBase/purchaseOrders($PoId)/documentAttachments($($meta.id))/attachmentContent" `
            -Headers ($authHeader + @{ 'Content-Type' = 'application/octet-stream'; 'If-Match' = $get.'@odata.etag' }) `
            -Body $Bytes | Out-Null
    }

    function Send-Alert([string]$Subject, [string]$BodyHtml) {
        if ($DryRun) { Write-Host "    ${modeTag}WOULD ALERT: $Subject" -ForegroundColor Magenta; return }
        $recipients = @($cfg.alertRecipients | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
        $payload = @{ message = @{ subject = $Subject; body = @{ contentType = 'HTML'; content = $BodyHtml }; toRecipients = $recipients } } | ConvertTo-Json -Depth 6
        try { Invoke-RestMethod -Method Post -Uri "$graphBase/sendMail" -Headers ($graphHeader + @{ 'Content-Type' = 'application/json' }) -Body $payload | Out-Null }
        catch { Write-Host "    [WARN] alert email failed: $($_.Exception.Message)" -ForegroundColor Yellow }
    }

    function Alert-Once([string]$Key, [string]$Subject, [string]$BodyHtml) {
        if ($NoAlert) { Write-Host "    (alert suppressed by -NoAlert): $Subject" -ForegroundColor DarkGray; return }
        if ($ledger.ContainsKey($Key)) { Write-Host "    (already alerted: $Key)" -ForegroundColor DarkGray; return }
        Send-Alert -Subject $Subject -BodyHtml $BodyHtml
        $ledger[$Key] = [PSCustomObject]@{ alertedAt = [DateTime]::UtcNow.ToString('o') }
        if (-not $DryRun) { Save-Ledger }
    }

    function Alert-Body([string]$Heading, [string]$Detail, [hashtable]$Fields) {
        $rows = ($Fields.GetEnumerator() | ForEach-Object { "<tr><td style='padding:3px 12px 3px 0;font-weight:600;color:#555'>$($_.Key)</td><td style='padding:3px 0'>$([System.Net.WebUtility]::HtmlEncode([string]$_.Value))</td></tr>" }) -join ''
        return "<h2 style='margin:0 0 4px'>$Heading</h2><div style='background:#fff3cd;color:#856404;border-left:4px solid #ffc107;padding:12px 16px;border-radius:0 4px 4px 0;margin:12px 0;font-size:13px'>$Detail</div><table style='font-size:13px'>$rows</table>"
    }

    # --- Scan: Vermes invoices in the lookback window (paged) -------------------
    $lookback = if ($LookbackDays -gt 0) { [double]$LookbackDays } else { [double]$cfg.lookbackDays }
    $sinceUtc = [DateTime]::UtcNow.AddDays(-$lookback)
    $search   = 'from:' + $cfg.sender
    $raw      = @()
    $url      = "$graphBase/messages?`$search=`"$search`"&`$select=id,subject,from,receivedDateTime,hasAttachments&`$top=50"
    $pages    = 0
    do {
        $resp  = Invoke-RestMethod -Uri $url -Headers ($graphHeader + @{ ConsistencyLevel = 'eventual' })
        $raw  += $resp.value
        $url   = if ($resp.PSObject.Properties['@odata.nextLink']) { $resp.'@odata.nextLink' } else { $null }
        $pages++
        # stop paging once results are older than the window (search returns newest-first)
        $oldest = $resp.value | Select-Object -Last 1
        if ($oldest -and [DateTime]$oldest.receivedDateTime -lt $sinceUtc) { break }
    } while ($url -and $pages -lt 20)

    $msgs = @($raw | Where-Object { (Get-FromAddr $_) -eq $cfg.sender -and $_.hasAttachments -and [DateTime]$_.receivedDateTime -ge $sinceUtc })
    Write-Host "  $($msgs.Count) Vermes email(s) with attachments in the last $([int]$lookback) day(s)." -ForegroundColor DarkGray

    $attached = 0; $skipped = 0; $pending = 0; $alerted = 0
    foreach ($m in $msgs) {
        $ageDays = ([DateTime]::UtcNow - [DateTime]$m.receivedDateTime).TotalDays
        $atts = (Invoke-RestMethod -Uri "$graphBase/messages/$($m.id)/attachments?`$select=id,name,contentType,size" -Headers $graphHeader).value
        $allPdfs = @($atts | Where-Object { $_.name -match '\.pdf$' -or $_.contentType -eq 'application/pdf' })
        # Only treat actual Vermes invoice PDFs (e.g. "Verkoopfactuur 26006472.pdf") as fileable;
        # this skips non-invoice PDFs that arrive from the same sender (payment advices, forwarded
        # customer invoices, etc.) silently — no alert noise.
        $invoicePat = if ($cfg.PSObject.Properties['invoiceFilePattern'] -and $cfg.invoiceFilePattern) { $cfg.invoiceFilePattern } else { '^Verkoopfactuur' }
        $pdfs = @($allPdfs | Where-Object { $_.name -match $invoicePat })
        if ($pdfs.Count -eq 0) {
            if ($allPdfs.Count -gt 0) { Write-Host "  [SKIP] '$($m.subject)' — no Vermes invoice PDF (only: $(@($allPdfs | ForEach-Object { $_.name }) -join ', '))." -ForegroundColor DarkGray }
            continue
        }

        foreach ($pdf in $pdfs) {
            Write-Host "  [$($m.receivedDateTime)] '$($m.subject)' :: $($pdf.name)"
            $bytes = [System.Convert]::FromBase64String((Invoke-RestMethod -Uri "$graphBase/messages/$($m.id)/attachments/$($pdf.id)" -Headers $graphHeader).contentBytes)
            $key   = "$($m.id)|$($pdf.name)"

            $ref = Get-YourRef -Bytes $bytes -Pattern $cfg.yourRefRegex
            if (-not $ref) {
                Write-Host "    [ALERT] Could not read 'Your Ref.:' from the PDF." -ForegroundColor Yellow
                Alert-Once -Key $key `
                    -Subject "[Vermes] Could not read PO reference — $($pdf.name)" `
                    -BodyHtml (Alert-Body 'Vermes Invoice — PO reference unreadable' "The <strong>Your Ref.:</strong> value could not be read from this PDF, so it was not filed. Please attach it to the correct purchase order manually in BC." ([ordered]@{ 'File' = $pdf.name; 'Email' = $m.subject; 'From' = (Get-FromAddr $m) }))
                $alerted++; continue
            }

            $pos = @(Find-PurchaseOrder $ref)
            if ($pos.Count -eq 0) {
                if ($ageDays -ge [double]$cfg.alertAfterDays) {
                    Write-Host "    [ALERT] No BC PO '$ref' after $([int]$ageDays)d — alerting once." -ForegroundColor Yellow
                    Alert-Once -Key $key `
                        -Subject "[Vermes] Purchase order $ref not found — $($pdf.name)" `
                        -BodyHtml (Alert-Body 'Vermes Invoice — PO not found' "No Business Central purchase order numbered <strong>$([System.Net.WebUtility]::HtmlEncode($ref))</strong> was found after $([int]$ageDays) day(s). The invoice PDF has not been filed. Check the PO number, then attach manually if needed." ([ordered]@{ 'Your Ref (PO)' = $ref; 'File' = $pdf.name; 'Email' = $m.subject }))
                    $alerted++
                } else {
                    Write-Host "    [WAIT] No BC PO '$ref' yet ($([int]$ageDays)d) — will retry." -ForegroundColor DarkGray
                    $pending++
                }
                continue
            }
            if ($pos.Count -gt 1) {
                Write-Host "    [ALERT] '$ref' matched $($pos.Count) POs — ambiguous." -ForegroundColor Yellow
                Alert-Once -Key $key -Subject "[Vermes] Ambiguous PO $ref — $($pdf.name)" `
                    -BodyHtml (Alert-Body 'Vermes Invoice — ambiguous PO' "Your Ref <strong>$([System.Net.WebUtility]::HtmlEncode($ref))</strong> matched more than one purchase order. Please file manually." ([ordered]@{ 'Your Ref (PO)' = $ref; 'Matches' = $pos.Count; 'File' = $pdf.name }))
                $alerted++; continue
            }

            $po = $pos[0]
            if ($cfg.vendorNumber -and $po.vendorNumber -ne $cfg.vendorNumber) {
                Write-Host "    [ALERT] PO $ref vendor=$($po.vendorNumber) != Vermes $($cfg.vendorNumber)." -ForegroundColor Yellow
                Alert-Once -Key $key -Subject "[Vermes] PO $ref is not a Vermes order — $($pdf.name)" `
                    -BodyHtml (Alert-Body 'Vermes Invoice — vendor mismatch' "Your Ref <strong>$([System.Net.WebUtility]::HtmlEncode($ref))</strong> matched purchase order <strong>$($po.number)</strong>, but its vendor is <strong>$($po.vendorNumber)</strong>, not Vermes ($($cfg.vendorNumber)). Not filed — please check." ([ordered]@{ 'Your Ref (PO)' = $ref; 'PO vendor' = $po.vendorNumber; 'File' = $pdf.name }))
                $alerted++; continue
            }

            $existing = @(Get-PoAttachmentNames $po.id)
            if ($existing -contains $pdf.name) {
                Write-Host "    [SKIP] Already attached to PO $($po.number)." -ForegroundColor DarkGray
                $skipped++; continue
            }

            if ($DryRun) {
                Write-Host "    ${modeTag}WOULD ATTACH '$($pdf.name)' -> PO $($po.number) (vendor $($po.vendorNumber))." -ForegroundColor Green
                $attached++
            } else {
                Add-PdfAttachment -PoId $po.id -FileName $pdf.name -Bytes $bytes
                Write-Host "    [OK] Attached '$($pdf.name)' -> PO $($po.number)." -ForegroundColor Green
                $attached++
            }
        }
    }

    Write-Host "`n${modeTag}Done. attached=$attached skipped=$skipped pending=$pending alerted=$alerted" -ForegroundColor Cyan
}
finally {
    if ($haveLock) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
