<#
.SYNOPSIS
    Watches supcom@montandor.com for purchase order emails and posts them to BC.
.DESCRIPTION
    Polls the inbox for unread emails with PDF attachments received today.
    Routes each email to the correct client using senders.json per client folder.
    Downloads PDF bytes directly into memory — no files written to disk.
    Marks the email as read only after a successful BC post.
    Run on a schedule via Task Scheduler (every N minutes).
.NOTES
    Requires Mail.Read + Mail.Send on the Claude-BC-ReadOnly-Xavier app registration.
    Credentials in Windows Credential Manager: Montandor_BC_TenantId/ClientId/ClientSecret.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Load all pipeline functions, DLLs, and tokens (skips the main folder-scan loop)
. (Join-Path $PSScriptRoot 'Invoke-SalesOrderPipeline.ps1') -FunctionsOnly

$supcom    = 'supcom@montandor.com'
$graphBase = "https://graph.microsoft.com/v1.0/users/$supcom"

# ---------------------------------------------------------------------------
# Route an email to a client folder by matching senders.json rules
# ---------------------------------------------------------------------------
function Get-ClientForEmail {
    param([string]$SenderEmail, [string]$EmailBody)
    foreach ($dir in Get-ChildItem $rootDir -Directory) {
        $sendersPath = Join-Path $dir.FullName 'senders.json'
        $tplPath     = Join-Path $dir.FullName 'template.json'
        if (-not (Test-Path $sendersPath) -or -not (Test-Path $tplPath)) { continue }
        $s = Get-Content $sendersPath | ConvertFrom-Json

        $senderMatch = $false
        if ($s.PSObject.Properties['senderDomain'] -and $s.senderDomain) {
            $senderMatch = $SenderEmail -like "*$($s.senderDomain)"
        }
        if (-not $senderMatch -and $s.PSObject.Properties['senderAddresses']) {
            foreach ($addr in @($s.senderAddresses)) {
                if ($SenderEmail -ieq $addr) { $senderMatch = $true; break }
            }
        }
        if (-not $senderMatch) { continue }

        if ($s.PSObject.Properties['bodyKeyword'] -and $s.bodyKeyword) {
            if ($EmailBody -notmatch [regex]::Escape($s.bodyKeyword)) { continue }
        }

        return $dir.FullName
    }
    return $null
}

# ---------------------------------------------------------------------------
# Download a PDF attachment as raw bytes (no file written to disk)
# ---------------------------------------------------------------------------
function Get-PdfAttachmentBytes {
    param([string]$MessageId, [string]$AttachmentId)
    $att = Invoke-RestMethod `
        -Uri "$graphBase/messages/$MessageId/attachments/$AttachmentId" `
        -Headers $graphHeader
    return [System.Convert]::FromBase64String($att.contentBytes)
}

# ---------------------------------------------------------------------------
# Send a plain-text notification email to supcom@montandor.com via Graph API
# ---------------------------------------------------------------------------
function Send-NotificationEmail {
    param([string]$Subject, [string]$Body, [string[]]$AlsoNotify = @())
    $recipients = @(@{ emailAddress = @{ address = $supcom } })
    foreach ($addr in $AlsoNotify) {
        $recipients += @{ emailAddress = @{ address = $addr } }
    }
    $payload = @{
        message = @{
            subject      = $Subject
            body         = @{ contentType = 'Text'; content = $Body }
            toRecipients = $recipients
        }
    } | ConvertTo-Json -Depth 5
    try {
        Invoke-RestMethod -Method Post -Uri "$graphBase/sendMail" `
            -Headers ($graphHeader + @{ 'Content-Type' = 'application/json' }) `
            -Body $payload | Out-Null
    } catch {
        Write-Host "    [WARN] Notification email failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$today = (Get-Date -Hour 0 -Minute 0 -Second 0).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
Write-Host "`n[WATCH] Scanning $supcom — all emails with attachments since $today" -ForegroundColor Cyan

# No isRead filter — a colleague reading an email before the script runs should not prevent
# the order from being created. Deduplication (Your_Reference check) prevents double-posting.
$filter = "hasAttachments eq true and receivedDateTime ge '$today'"
$msgs   = Invoke-RestMethod `
    -Uri "$graphBase/mailFolders/inbox/messages?`$filter=$filter&`$select=id,subject,from,receivedDateTime,body&`$orderby=receivedDateTime asc&`$top=50" `
    -Headers $graphHeader

Write-Host "Found $($msgs.value.Count) candidate email(s)." -ForegroundColor Cyan

$ordersPosted = 0
$skipped      = 0

foreach ($msg in $msgs.value) {
    $senderEmail = $msg.from.emailAddress.address
    $bodyText    = $msg.body.content -replace '<[^>]+>', ' '   # strip HTML

    Write-Host ("`n  [{0}] {1}" -f $msg.receivedDateTime, $msg.subject)
    Write-Host "  From: $senderEmail"

    $clientDir = Get-ClientForEmail -SenderEmail $senderEmail -EmailBody $bodyText
    if (-not $clientDir) {
        Write-Host "  [SKIP] No routing rule matches this sender." -ForegroundColor Yellow
        $skipped++
        continue
    }

    $tpl = Get-Content (Join-Path $clientDir 'template.json') | ConvertFrom-Json
    Write-Host "  Client: $($tpl.clientName)" -ForegroundColor Cyan

    $atts    = Invoke-RestMethod `
        -Uri "$graphBase/messages/$($msg.id)/attachments?`$select=id,name,contentType,isInline,size" `
        -Headers $graphHeader
    $pdfAtts = @($atts.value | Where-Object { -not $_.isInline -and $_.contentType -match 'pdf' })

    if ($pdfAtts.Count -eq 0) {
        Write-Host "  [SKIP] No PDF attachments in this email." -ForegroundColor Yellow
        $skipped++
        continue
    }

    $isTextMode = $tpl.PSObject.Properties['extractionMode'] -and $tpl.extractionMode -eq 'text'
    $bcItems    = if ($isTextMode) { Get-BcItemNumbers -Environment $tpl.environment } else { @() }
    $emailOk    = $true

    foreach ($att in $pdfAtts) {
        Write-Host "  PDF: $($att.name) ($([Math]::Round($att.size/1KB, 1)) KB)"
        try {
            $pdfBytes = Get-PdfAttachmentBytes -MessageId $msg.id -AttachmentId $att.id
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            $data = Get-PdfOrderData -PdfBytes $pdfBytes -Template $tpl -BcItemNumbers $bcItems

            # Ship-to lookup for text-mode clients
            $shipToCode = ''
            $skipOrder  = $false
            if ($isTextMode -and $data.ShipToPostCode) {
                $candidates = Get-BcShipToCode -CustomerNumber $tpl.customerNumber -PostCode $data.ShipToPostCode -Environment $tpl.environment
                if ($candidates.Count -eq 0) {
                    Write-Host "    [SKIP] No BC ship-to for postcode $($data.ShipToPostCode)" -ForegroundColor Yellow
                    $skipOrder = $true; $emailOk = $false
                    Send-NotificationEmail `
                        -Subject "[Sales Order] Action required — $($tpl.clientName) ref $($data.OrderRef)" `
                        -Body @"
A purchase order could not be posted to BC because the delivery postcode is not registered as a ship-to address.

Client:    $($tpl.clientName)
Order ref: $($data.OrderRef)
From:      $senderEmail
Email:     $($msg.subject)
Postcode:  $($data.ShipToPostCode)

Action: Add postcode $($data.ShipToPostCode) as a ship-to address for $($tpl.clientName) in BC.
The order will be picked up automatically on the next watcher run.
"@
                } elseif ($candidates.Count -eq 1) {
                    $shipToCode = $candidates[0].code
                    Write-Host "    Ship-to: $shipToCode — $($candidates[0].displayName)" -ForegroundColor Green
                } else {
                    $codes = ($candidates | ForEach-Object { $_.code }) -join ', '
                    Write-Host "    [SKIP] Ambiguous postcode $($data.ShipToPostCode): $codes" -ForegroundColor Yellow
                    $skipOrder = $true; $emailOk = $false
                    Send-NotificationEmail `
                        -Subject "[Sales Order] Action required — $($tpl.clientName) ref $($data.OrderRef)" `
                        -Body @"
A purchase order could not be posted to BC because multiple ship-to addresses match the delivery postcode.

Client:    $($tpl.clientName)
Order ref: $($data.OrderRef)
From:      $senderEmail
Email:     $($msg.subject)
Postcode:  $($data.ShipToPostCode)
Matches:   $codes

Action: The postcode is ambiguous — please post this order manually in BC.
"@
                }
            }

            # Deduplication
            $alreadyExists = $false
            if ($data.OrderRef) {
                $alreadyExists = Test-BcOrderExists -CustomerNumber $tpl.customerNumber `
                    -OrderRef $data.OrderRef -Environment $tpl.environment
            }

            Write-Host "    Order ref : $($data.OrderRef)"
            Write-Host "    Lines     : $($data.Lines.Count)"

            if ($alreadyExists) {
                $bcLines  = Get-BcOrderLines -CustomerNumber $tpl.customerNumber -OrderRef $data.OrderRef -Environment $tpl.environment
                $lineDiff = if ($bcLines) { Compare-OrderLines -NewLines $data.Lines -BcLines $bcLines } else { @() }

                if ($lineDiff.Count -gt 0) {
                    # Same order ref, different lines — customer likely modified their PO
                    Write-Host "    [NOTIFY] Modified PO detected — $($lineDiff.Count) line difference(s)." -ForegroundColor Yellow
                    Send-NotificationEmail `
                        -Subject "[Sales Order] Modified PO — $($tpl.clientName) ref $($data.OrderRef)" `
                        -Body @"
A purchase order has been received that matches an existing BC order but with different line items.
The BC order has NOT been updated automatically. Please review and update manually if needed.

Client:    $($tpl.clientName)
Order ref: $($data.OrderRef)
From:      $senderEmail
Email:     $($msg.subject)

Changes vs existing BC order:
$($lineDiff -join "`n")
"@
                } else {
                    # True duplicate (same lines) — silent regardless of email age
                    Write-Host "    [SKIP] Order ref $($data.OrderRef) already in BC, lines unchanged." -ForegroundColor Yellow
                }
            } elseif (-not $skipOrder) {
                $bcNo = Submit-SalesOrder -OrderData $data -Template $tpl -ShipToCode $shipToCode
                Write-Host "    -> BC $bcNo posted to $($tpl.environment)" -ForegroundColor Green
                $ordersPosted++
            }
        } catch {
            $errMsg = $_.Exception.Message
            Write-Host "    [ERROR] $errMsg" -ForegroundColor Red
            $emailOk = $false
            Send-NotificationEmail `
                -Subject "[Sales Order] Error — $($tpl.clientName) $($att.name)" `
                -AlsoNotify @('x.planchette@montandor.com') `
                -Body @"
An error occurred while processing a purchase order PDF. The order has NOT been posted to BC.

Client:    $($tpl.clientName)
Email:     $($msg.subject)
From:      $senderEmail
File:      $($att.name)
Error:     $errMsg

Please review and process this order manually if needed.
"@
        }
    }

    if ($emailOk) {
        Write-Host "  All PDFs processed." -ForegroundColor Green
    } else {
        Write-Host "  One or more PDFs had errors — will retry on next run." -ForegroundColor Yellow
    }
}

# Wipe credentials from memory
$clientSecret = $null; $token = $null; $graphToken = $null

Write-Host ("`n[WATCH] Complete. Orders posted: {0} | Emails skipped: {1}`n" -f $ordersPosted, $skipped) -ForegroundColor Cyan
