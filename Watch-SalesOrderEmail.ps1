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
# Mark an email as read in the inbox
# ---------------------------------------------------------------------------
function Set-EmailRead {
    param([string]$MessageId)
    Invoke-RestMethod -Method Patch `
        -Uri "$graphBase/messages/$MessageId" `
        -Headers ($graphHeader + @{ 'Content-Type' = 'application/json' }) `
        -Body '{"isRead":true}' | Out-Null
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$today = (Get-Date -Hour 0 -Minute 0 -Second 0).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
Write-Host "`n[WATCH] Scanning $supcom — unread emails with attachments since $today" -ForegroundColor Cyan

$filter = "isRead eq false and hasAttachments eq true and receivedDateTime ge '$today'"
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
                } elseif ($candidates.Count -eq 1) {
                    $shipToCode = $candidates[0].code
                    Write-Host "    Ship-to: $shipToCode — $($candidates[0].displayName)" -ForegroundColor Green
                } else {
                    $codes = ($candidates | ForEach-Object { $_.code }) -join ', '
                    Write-Host "    [SKIP] Ambiguous postcode $($data.ShipToPostCode): $codes" -ForegroundColor Yellow
                    $skipOrder = $true; $emailOk = $false
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
                Write-Host "    [SKIP] Order ref $($data.OrderRef) already exists in BC." -ForegroundColor Yellow
            } elseif (-not $skipOrder) {
                $bcNo = Submit-SalesOrder -OrderData $data -Template $tpl -ShipToCode $shipToCode
                Write-Host "    -> BC $bcNo posted to $($tpl.environment)" -ForegroundColor Green
                $ordersPosted++
            }
        } catch {
            Write-Host "    [ERROR] $($_.Exception.Message)" -ForegroundColor Red
            $emailOk = $false
        }
    }

    if ($emailOk) {
        Set-EmailRead -MessageId $msg.id
        Write-Host "  Email marked as read." -ForegroundColor Green
    } else {
        Write-Host "  Email left unread — will retry on next run." -ForegroundColor Yellow
    }
}

# Wipe credentials from memory
$clientSecret = $null; $token = $null; $graphToken = $null

Write-Host ("`n[WATCH] Complete. Orders posted: {0} | Emails skipped: {1}`n" -f $ordersPosted, $skipped) -ForegroundColor Cyan
