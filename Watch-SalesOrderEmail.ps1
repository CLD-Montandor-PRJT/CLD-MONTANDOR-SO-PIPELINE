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
$FunctionsOnly = $false   # reset scope bleed — pipeline switch sets this to $true in dot-source scope

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
    param([string]$Subject, [string]$Body, [string[]]$AlsoNotify = @(), [string]$ContentType = 'Text')
    $recipients = @(@{ emailAddress = @{ address = $supcom } })
    foreach ($addr in $AlsoNotify) {
        $recipients += @{ emailAddress = @{ address = $addr } }
    }
    $payload = @{
        message = @{
            subject      = $Subject
            body         = @{ contentType = $ContentType; content = $Body }
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

function Build-InfoBox {
    param([hashtable]$Fields, [string]$Bg = '#f0f0f0')
    $rows = ($Fields.GetEnumerator() | ForEach-Object {
        "<tr><td style='padding:4px 12px 4px 0;font-weight:600;white-space:nowrap;color:#555'>$($_.Key)</td><td style='padding:4px 0'>$($_.Value)</td></tr>"
    }) -join ''
    return "<div style='background:$Bg;padding:12px 16px;border-radius:4px;margin-bottom:20px'><table style='font-size:13px;border:none'>$rows</table></div>"
}

function Build-AlertBox {
    param([string]$Message, [string]$Bg = '#fff3cd', [string]$Fg = '#856404', [string]$Border = '#ffc107')
    return "<div style='background:$Bg;color:$Fg;border-left:4px solid $Border;padding:12px 16px;border-radius:0 4px 4px 0;margin-bottom:20px;font-size:13px'>$Message</div>"
}

function Build-HtmlShell {
    param([string]$Title, [string]$Subtitle, [string]$Body)
    return "<html><body style='font-family:Arial,Calibri,sans-serif;font-size:13px;color:#333;max-width:700px;margin:0 auto;padding:20px'>
  <h2 style='margin:0 0 4px'>Sales Order Pipeline &mdash; $Title</h2>
  <p style='margin:0 0 16px;color:#666;font-size:12px'>$Subtitle</p>
  $Body
</body></html>"
}

function Build-ShipToNotFoundHtml {
    param([string]$ClientName, [string]$OrderRef, [string]$SenderEmail, [string]$EmailSubject, [string]$PostCode)
    $info   = Build-InfoBox ([ordered]@{ Client=''; 'Order ref'=''; From=''; Email='' })
    # Build ordered hashtable properly
    $fields = [ordered]@{ 'Client' = $ClientName; 'Order ref' = $OrderRef; 'From' = $SenderEmail; 'Email' = $EmailSubject; 'Postcode' = "<strong>$PostCode</strong>" }
    $info   = Build-InfoBox $fields
    $alert  = Build-AlertBox "Add postcode <strong>$PostCode</strong> as a ship-to address for <strong>$ClientName</strong> in BC. The order will be picked up automatically on the next watcher run."
    return Build-HtmlShell -Title 'Delivery postcode not registered in BC' -Subtitle 'This order has not been posted and requires manual action.' -Body "$info$alert"
}

function Build-AmbiguousPostcodeHtml {
    param([string]$ClientName, [string]$OrderRef, [string]$SenderEmail, [string]$EmailSubject, [string]$PostCode, [string]$Matches)
    $fields = [ordered]@{ 'Client' = $ClientName; 'Order ref' = $OrderRef; 'From' = $SenderEmail; 'Email' = $EmailSubject; 'Postcode' = "<strong>$PostCode</strong>"; 'Matching codes' = "<strong>$Matches</strong>" }
    $info   = Build-InfoBox $fields
    $alert  = Build-AlertBox "Postcode <strong>$PostCode</strong> matches multiple ship-to addresses. Please post this order manually in BC."
    return Build-HtmlShell -Title 'Delivery postcode matches multiple ship-to addresses' -Subtitle 'This order has not been posted and requires manual action.' -Body "$info$alert"
}

function Build-ProcessingErrorHtml {
    param([string]$ClientName, [string]$SenderEmail, [string]$EmailSubject, [string]$FileName,
          [string]$ErrorMessage, [string]$ExceptionType = '', [string]$FailingLine = '', [string]$StackTrace = '')
    $fields = [ordered]@{ 'Client' = $ClientName; 'From' = $SenderEmail; 'Email' = $EmailSubject; 'File' = $FileName }
    $info   = Build-InfoBox $fields
    $alert  = Build-AlertBox -Message "<strong>$ExceptionType</strong><br>$([System.Net.WebUtility]::HtmlEncode($ErrorMessage))" -Bg '#f8d7da' -Fg '#721c24' -Border '#f5c6cb'
    $detail = ''
    if ($FailingLine) {
        $detail += "<div style='margin-bottom:12px'><p style='margin:0 0 4px;font-weight:600;font-size:12px;color:#555'>FAILING LINE</p><pre style='background:#f8f8f8;border:1px solid #ddd;border-radius:4px;padding:10px;font-size:12px;margin:0;overflow-x:auto'>$([System.Net.WebUtility]::HtmlEncode($FailingLine.Trim()))</pre></div>"
    }
    if ($StackTrace) {
        $detail += "<div style='margin-bottom:12px'><p style='margin:0 0 4px;font-weight:600;font-size:12px;color:#555'>STACK TRACE</p><pre style='background:#f8f8f8;border:1px solid #ddd;border-radius:4px;padding:10px;font-size:12px;margin:0;overflow-x:auto'>$([System.Net.WebUtility]::HtmlEncode($StackTrace))</pre></div>"
    }
    $note   = "<p style='color:#555;font-size:12px;margin:12px 0 0'>The order has <strong>not</strong> been posted to BC. Please review and process manually if needed.</p>"
    return Build-HtmlShell -Title 'Processing Error' -Subtitle 'An error occurred while processing a purchase order PDF.' -Body "$info$alert$detail$note"
}

function Build-ModifiedOrderHtml {
    param(
        [string]$ClientName,
        [string]$OrderRef,
        [string]$SenderEmail,
        [string]$EmailSubject,
        [PSCustomObject[]]$Diff        = @(),
        [string]$ShipToOld             = '',
        [string]$ShipToNew             = ''
    )

    $body = "<div style='background:#f0f0f0;padding:12px 16px;border-radius:4px;font-size:13px;margin-bottom:20px;line-height:2'>
    <strong>Client:</strong> $ClientName &nbsp;&nbsp;&nbsp;
    <strong>Order ref:</strong> $OrderRef <br>
    <strong>From:</strong> $SenderEmail &nbsp;&nbsp;&nbsp;
    <strong>Email:</strong> $EmailSubject
  </div>"

    if ($ShipToOld -or $ShipToNew) {
        $body += "<div style='margin-bottom:20px'>
    <p style='margin:0 0 8px;font-weight:600;font-size:13px'>Delivery address changed</p>
    <table style='border-collapse:collapse;width:100%;font-size:13px'>
      <thead><tr>
        <th style='padding:8px 12px;border:1px solid #ddd;text-align:left;background:#856404;color:#fff;width:50%'>In BC (current)</th>
        <th style='padding:8px 12px;border:1px solid #ddd;text-align:left;background:#155724;color:#fff;width:50%'>In PDF (new)</th>
      </tr></thead>
      <tbody><tr>
        <td style='padding:10px 12px;border:1px solid #ddd;vertical-align:top;background:#fff3cd;color:#856404'>$([System.Net.WebUtility]::HtmlEncode($ShipToOld) -replace "`n",'<br>')</td>
        <td style='padding:10px 12px;border:1px solid #ddd;vertical-align:top;background:#d4edda;color:#155724'>$([System.Net.WebUtility]::HtmlEncode($ShipToNew) -replace "`n",'<br>')</td>
      </tr></tbody>
    </table></div>"
    }

    if (@($Diff).Count -gt 0) {
        $rowColors = @{
            Changed = @{ bg = '#fff3cd'; fg = '#856404' }
            Added   = @{ bg = '#d4edda'; fg = '#155724' }
            Removed = @{ bg = '#f8d7da'; fg = '#721c24' }
        }
        $rows = ''
        foreach ($d in ($Diff | Sort-Object Type, ItemNumber)) {
            $c      = $rowColors[$d.Type]
            $wasStr = if ($null -eq $d.OldQty) { '&mdash;' } else { if ($d.OldQty -eq [Math]::Floor($d.OldQty)) { [int]$d.OldQty } else { $d.OldQty } }
            $nowStr = if ($null -eq $d.NewQty) { '&mdash;' } else { if ($d.NewQty -eq [Math]::Floor($d.NewQty)) { [int]$d.NewQty } else { $d.NewQty } }
            $rows  += "<tr style='background:$($c.bg);color:$($c.fg)'>
            <td style='padding:7px 12px;border:1px solid #ddd;font-weight:600'>$($d.ItemNumber)</td>
            <td style='padding:7px 12px;border:1px solid #ddd'>$($d.Type)</td>
            <td style='padding:7px 12px;border:1px solid #ddd;text-align:right'>$wasStr</td>
            <td style='padding:7px 12px;border:1px solid #ddd;text-align:right'>$nowStr</td>
        </tr>"
        }
        $body += "<table style='border-collapse:collapse;width:100%;font-size:13px'>
    <thead><tr>
      <th style='padding:8px 12px;border:1px solid #ddd;text-align:left;background:#343a40;color:#fff'>Item</th>
      <th style='padding:8px 12px;border:1px solid #ddd;text-align:left;background:#343a40;color:#fff'>Change</th>
      <th style='padding:8px 12px;border:1px solid #ddd;text-align:right;background:#343a40;color:#fff'>Previous Qty</th>
      <th style='padding:8px 12px;border:1px solid #ddd;text-align:right;background:#343a40;color:#fff'>New Qty</th>
    </tr></thead>
    <tbody>$rows</tbody>
  </table>"
    }

    return "<html><body style='font-family:Arial,Calibri,sans-serif;font-size:13px;color:#333;max-width:700px;margin:0 auto;padding:20px'>
  <h2 style='margin:0 0 4px'>Sales Order Pipeline &mdash; Modified Order</h2>
  <p style='margin:0 0 16px;color:#666;font-size:12px'>The BC order has <strong>not</strong> been updated automatically. Please review and update manually if needed.</p>
  $body
</body></html>"
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
if ($FunctionsOnly) { return }   # dot-source mode — load functions only, skip main block

$today = (Get-Date -Hour 0 -Minute 0 -Second 0).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
Write-Host "`n[WATCH] Scanning $supcom — all emails with attachments since $today" -ForegroundColor Cyan

# No isRead filter — a colleague reading an email before the script runs should not prevent
# the order from being created. Deduplication (Your_Reference check) prevents double-posting.
$filter = "hasAttachments eq true and receivedDateTime ge $today"
try {
    $msgs = Invoke-RestMethod `
        -Uri "$graphBase/mailFolders/inbox/messages?`$filter=$filter&`$select=id,subject,from,receivedDateTime,body&`$top=50" `
        -Headers $graphHeader
} catch {
    $errMsg   = $_.Exception.Message
    $errType  = $_.Exception.GetType().Name
    $errLine  = if ($_.InvocationInfo.Line) { $_.InvocationInfo.Line } else { '' }
    $errStack = if ($_.ScriptStackTrace)    { $_.ScriptStackTrace }    else { '' }
    Write-Host "[ERROR] Failed to fetch inbox: $errMsg" -ForegroundColor Red
    Send-NotificationEmail `
        -Subject     "[Sales Order] Watcher failed — could not read inbox" `
        -AlsoNotify  @('x.planchette@montandor.com') `
        -Body        (Build-ProcessingErrorHtml -ClientName 'Sales Order Watcher' -SenderEmail 'N/A' -EmailSubject 'N/A' -FileName 'Watch-SalesOrderEmail.ps1' -ErrorMessage $errMsg -ExceptionType $errType -FailingLine $errLine -StackTrace $errStack) `
        -ContentType 'HTML'
    exit 1
}

Write-Host "Found $($msgs.value.Count) candidate email(s)." -ForegroundColor Cyan

$ordersPosted = 0
$skipped      = 0
$seenRefs     = [System.Collections.Generic.HashSet[string]]::new()   # in-run dedup guard

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
            $data = Get-PdfOrderData -PdfBytes $pdfBytes -Template $tpl -BcItemNumbers $bcItems -ClientDir $clientDir

            # Ship-to lookup for text-mode clients
            $shipToCode = ''
            $skipOrder  = $false
            if ($isTextMode -and $data.ShipToPostCode) {
                $candidates = Get-BcShipToCode -CustomerNumber $tpl.customerNumber -PostCode $data.ShipToPostCode -Environment $tpl.environment
                if ($candidates.Count -eq 0) {
                    Write-Host "    [SKIP] No BC ship-to for postcode $($data.ShipToPostCode)" -ForegroundColor Yellow
                    $skipOrder = $true; $emailOk = $false
                    Send-NotificationEmail `
                        -Subject      "[Sales Order] Action required — $($tpl.clientName) ref $($data.OrderRef)" `
                        -Body         (Build-ShipToNotFoundHtml -ClientName $tpl.clientName -OrderRef $data.OrderRef -SenderEmail $senderEmail -EmailSubject $msg.subject -PostCode $data.ShipToPostCode) `
                        -ContentType  'HTML'
                } elseif ($candidates.Count -eq 1) {
                    $shipToCode = $candidates[0].code
                    Write-Host "    Ship-to: $shipToCode — $($candidates[0].displayName)" -ForegroundColor Green
                } else {
                    $codes = ($candidates | ForEach-Object { $_.code }) -join ', '
                    Write-Host "    [SKIP] Ambiguous postcode $($data.ShipToPostCode): $codes" -ForegroundColor Yellow
                    $skipOrder = $true; $emailOk = $false
                    Send-NotificationEmail `
                        -Subject     "[Sales Order] Action required — $($tpl.clientName) ref $($data.OrderRef)" `
                        -Body        (Build-AmbiguousPostcodeHtml -ClientName $tpl.clientName -OrderRef $data.OrderRef -SenderEmail $senderEmail -EmailSubject $msg.subject -PostCode $data.ShipToPostCode -Matches $codes) `
                        -ContentType 'HTML'
                }
            }

            # Deduplication — in-memory guard first (same run), then BC query
            $refKey        = "$($tpl.customerNumber)|$($data.OrderRef)"
            $alreadyExists = $false
            if ($data.OrderRef) {
                if ($seenRefs.Contains($refKey)) {
                    $alreadyExists = $true
                    Write-Host "    [SKIP] $($data.OrderRef) already posted this run — skipping." -ForegroundColor Yellow
                } else {
                    $alreadyExists = Test-BcOrderExists -CustomerNumber $tpl.customerNumber `
                        -OrderRef $data.OrderRef -Environment $tpl.environment
                }
            }

            Write-Host "    Order ref : $($data.OrderRef)"
            Write-Host "    Lines     : $($data.Lines.Count)"

            if (-not $data.OrderRef) {
                Write-Host "    [SKIP] No order reference extracted — PDF is not a recognisable purchase order." -ForegroundColor Yellow
                continue
            }

            if ($alreadyExists) {
                $bcLines  = @(Get-BcOrderLines -CustomerNumber $tpl.customerNumber -OrderRef $data.OrderRef -Environment $tpl.environment)
                $lineDiff = @(if ($bcLines) { Compare-OrderLines -NewLines $data.Lines -BcLines $bcLines } else { @() })

                # Ship-to comparison
                $shipToChanged = $false
                $shipToOldStr  = ''
                $shipToNewStr  = ''
                $bcShipTo = Get-BcOrderShipTo -CustomerNumber $tpl.customerNumber -OrderRef $data.OrderRef -Environment $tpl.environment
                if ($bcShipTo -and $data.ShipTo) {
                    # Coordinate-mode: compare full address fields
                    $pdfKey = "$($data.ShipTo.name)|$($data.ShipTo.addressLine1)|$($data.ShipTo.postCode)|$($data.ShipTo.city)"
                    $bcKey  = "$($bcShipTo.name)|$($bcShipTo.addressLine1)|$($bcShipTo.postCode)|$($bcShipTo.city)"
                    if ($pdfKey.Trim() -ne $bcKey.Trim()) {
                        $shipToChanged = $true
                        $shipToOldStr  = "$($bcShipTo.name)`n$($bcShipTo.addressLine1)`n$($bcShipTo.postCode) $($bcShipTo.city)"
                        $shipToNewStr  = "$($data.ShipTo.name)`n$($data.ShipTo.addressLine1)`n$($data.ShipTo.postCode) $($data.ShipTo.city)"
                    }
                } elseif ($bcShipTo -and $shipToCode -and $shipToCode -ne $bcShipTo.code) {
                    # Text-mode: compare resolved ship-to code
                    $shipToChanged = $true
                    $shipToOldStr  = $bcShipTo.code
                    $shipToNewStr  = $shipToCode
                }

                if ($lineDiff.Count -gt 0 -or $shipToChanged) {
                    $changes = @()
                    if ($lineDiff.Count -gt 0)  { $changes += "$($lineDiff.Count) line difference(s)" }
                    if ($shipToChanged)           { $changes += "delivery address changed" }
                    Write-Host "    [NOTIFY] Modified PO detected — $($changes -join ', ')." -ForegroundColor Yellow
                    $html = Build-ModifiedOrderHtml `
                        -ClientName   $tpl.clientName `
                        -OrderRef     $data.OrderRef `
                        -SenderEmail  $senderEmail `
                        -EmailSubject $msg.subject `
                        -Diff         $lineDiff `
                        -ShipToOld    $shipToOldStr `
                        -ShipToNew    $shipToNewStr
                    Send-NotificationEmail `
                        -Subject     "[Sales Order] Modified Order — $($tpl.clientName) ref $($data.OrderRef)" `
                        -Body        $html `
                        -ContentType 'HTML'
                } else {
                    Write-Host "    [SKIP] Order ref $($data.OrderRef) already in BC, lines and address unchanged." -ForegroundColor Yellow
                }
            } elseif (-not $skipOrder) {
                $bcNo = Submit-SalesOrder -OrderData $data -Template $tpl -ShipToCode $shipToCode -PdfBytes $pdfBytes -PdfFileName $att.name
                $seenRefs.Add($refKey) | Out-Null
                Write-Host "    -> BC $bcNo posted to $($tpl.environment)" -ForegroundColor Green
                $ordersPosted++
            }
        } catch {
            $errMsg    = $_.Exception.Message
            $errType   = $_.Exception.GetType().Name
            $errLine   = if ($_.InvocationInfo.Line)        { $_.InvocationInfo.Line }        else { '' }
            $errStack  = if ($_.ScriptStackTrace)           { $_.ScriptStackTrace }           else { '' }
            Write-Host "    [ERROR] $errMsg" -ForegroundColor Red
            $emailOk = $false
            Send-NotificationEmail `
                -Subject     "[Sales Order] Error — $($tpl.clientName) $($att.name)" `
                -AlsoNotify  @('x.planchette@montandor.com') `
                -Body        (Build-ProcessingErrorHtml -ClientName $tpl.clientName -SenderEmail $senderEmail -EmailSubject $msg.subject -FileName $att.name -ErrorMessage $errMsg -ExceptionType $errType -FailingLine $errLine -StackTrace $errStack) `
                -ContentType 'HTML'
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
