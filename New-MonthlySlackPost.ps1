<#
.SYNOPSIS
  Posts the monthly Global Purchase Report - "PO Price Comparison" notice to a Slack
  channel, via the Slack Web API directly (chat.postMessage / files.*ExternalUpload),
  using the Bot Token saved by Save-SlackToken.ps1.

.DESCRIPTION
  Same content/purpose as New-MonthlyEmailDraft.ps1's Gmail draft - this is a SECOND,
  independent distribution channel, not a replacement. Posts to a channel (not
  individual DMs) so colleagues who join Slack later can still see it in the channel's
  history, since not everyone at the company has a Slack account yet.

  Unlike the Gmail draft (which never attaches - the raw report exceeds Gmail's ~25MB
  limit once base64-inflated), this DOES attach a file: the single-file interactive
  HTML report, which at ~28MB comfortably fits Slack's much higher per-file upload
  limit. Uses Slack's modern 3-step external upload flow (files.getUploadURLExternal ->
  raw PUT -> files.completeUploadExternal with initial_comment), since the old
  files.upload endpoint is deprecated.

.PARAMETER ReportMonth
  "yyyy-MM", e.g. "2026-07".

.PARAMETER Channel
  Slack channel name (without '#') or channel ID to post to. Must be a public channel
  if using only chat:write.public, or the bot must already be a member for a private one.

.PARAMETER DeliverablePath
  Path to this month's outside-delivery deliverable, just to confirm it exists as a
  readiness check (referenced by name in the message, never itself attached - still
  too large even for Slack once you account for the xlsx already being the source of
  the HTML report). Defaults to the standard naming convention.

.PARAMETER HtmlAttachmentPath
  Path to the interactive HTML report to actually attach. Defaults to the dated
  snapshot Generate-PurchaseReportHtml.ps1 -Publish produces for this ReportMonth
  ("...Report_html_<YYYYMM>.html"); falls back to the TEMP working file (with a
  warning) if that dated snapshot doesn't exist yet.

.PARAMETER IncludeHtmlIntro
  Adds a one-time bolded "New:" callout introducing the interactive web (HTML) format
  itself - added 2026-08-27 when this channel first started receiving it. Defaults to
  $true; once the channel is used to it, pass -IncludeHtmlIntro:$false to drop it from
  future months' posts without editing this script.

.EXAMPLE
  .\New-MonthlySlackPost.ps1 -ReportMonth "2026-07" -Channel "global-purchase-report"
  .\New-MonthlySlackPost.ps1 -ReportMonth "2026-08" -Channel "C0BSWRJU85T" -IncludeHtmlIntro:$false
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d{4}-\d{2}$')]
    [string]$ReportMonth,

    [Parameter(Mandatory = $true)]
    [string]$Channel,

    [string]$DeliverablePath,

    [string]$HtmlAttachmentPath,

    [bool]$IncludeHtmlIntro = $true
)

$ErrorActionPreference = 'Stop'

$tokenPath = Join-Path $PSScriptRoot 'Credentials\slack_bot_token.enc'
if (-not (Test-Path $tokenPath)) {
    throw "No stored Slack token found at $tokenPath - run .\Save-SlackToken.ps1 first (see SLACK_SETUP.md)."
}
$encrypted = (Get-Content -Path $tokenPath -Raw).Trim()
$secure = ConvertTo-SecureString -String $encrypted
$bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
try {
    $botToken = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
} finally {
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
}

$year = [int]$ReportMonth.Substring(0, 4)
$month = [int]$ReportMonth.Substring(5, 2)
$periodLabel = "(1-$month)$year"   # matches Publish-OutsideReport.ps1 / New-MonthlyEmailDraft.ps1

if (-not $DeliverablePath) {
    $reportDir = Join-Path $PSScriptRoot "Global Purchase Report RPA\REPORT"
    $pattern = "FIN_MHK+STC+MZQ+MIL+RACO_PUR_REPORT-$periodLabel*.xlsx"
    $candidates = Get-ChildItem -Path $reportDir -Filter $pattern -ErrorAction SilentlyContinue
    if ($candidates) {
        $DeliverablePath = ($candidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
    }
}
if (-not $DeliverablePath -or -not (Test-Path $DeliverablePath)) {
    throw "No deliverable found for $periodLabel in 'Global Purchase Report RPA\REPORT' - run Publish-OutsideReport.ps1 -ReportMonth `"$ReportMonth`" first, or pass -DeliverablePath explicitly."
}

$deliverableName = [System.IO.Path]::GetFileName($DeliverablePath)
Write-Host "Deliverable found (readiness check only, NOT attached): $DeliverablePath"

if (-not $HtmlAttachmentPath) {
    $htmlDir = Join-Path $PSScriptRoot "Global Purchase Report RPA\REPORT"
    $yyyymm = "{0:D4}{1:D2}" -f $year, $month
    $datedHtmlPath = Join-Path $htmlDir "FIN_Global_Purchase -Report_html_$yyyymm.html"
    $tempHtmlPath = Join-Path $htmlDir "FIN_Global_Purchase -Report_html_TEMP.html"
    if (Test-Path $datedHtmlPath) {
        $HtmlAttachmentPath = $datedHtmlPath
    } elseif (Test-Path $tempHtmlPath) {
        Write-Host "Warning: dated snapshot not found ($datedHtmlPath) - falling back to the TEMP working file. Run Generate-PurchaseReportHtml.ps1 -Publish to create the dated snapshot for this month." -ForegroundColor Yellow
        $HtmlAttachmentPath = $tempHtmlPath
    } else {
        throw "No HTML report found (checked $datedHtmlPath and $tempHtmlPath) - run Generate-PurchaseReportHtml.ps1 -Publish first, or pass -HtmlAttachmentPath explicitly."
    }
}
if (-not (Test-Path $HtmlAttachmentPath)) {
    throw "HTML report not found: $HtmlAttachmentPath - run Generate-PurchaseReportHtml.ps1 -Publish first, or pass -HtmlAttachmentPath explicitly."
}
$htmlFileName = [System.IO.Path]::GetFileName($HtmlAttachmentPath)
$htmlBytes = [System.IO.File]::ReadAllBytes($HtmlAttachmentPath)
Write-Host "HTML report to attach: $HtmlAttachmentPath ($([Math]::Round($htmlBytes.Length / 1MB, 1)) MB)"
Write-Host "Channel      : #$Channel"
Write-Host "HTML intro   : $(if ($IncludeHtmlIntro) { 'included' } else { 'omitted' }) (-IncludeHtmlIntro)"

$htmlIntroLine = ""
if ($IncludeHtmlIntro) {
    $htmlIntroLine = "`n*New:* this report is now shared as an interactive web (HTML) file - download the attachment below and open it in your web browser (double-click it, or drag it into Chrome/Edge) to sort, filter and search the data interactively, no Excel required.`n"
}

$text = @"
*Global purchase report $periodLabel - PO Price Comparison*

Dear All,

Please find attached the interactive PO Price Comparison report (Fittings, Glass Lids, Packaging and Coating items) - open the attached HTML file in a browser.
$htmlIntroLine
(Source workbook: ``$deliverableName`` - too large to attach directly; ask Anthony Man for a copy or a Google Drive link if you need the raw Excel file.)

If you have any questions, please feel free to contact Anthony Man (anthonym@meyer-msc.com).

Best Regards,
Anthony - Finance Department
"@

$headers = @{ Authorization = "Bearer $botToken" }

# ---- Step 1: get an upload URL + file_id (files.upload is deprecated) ----
$uploadUrlBody = @{ filename = $htmlFileName; length = $htmlBytes.Length }
$uploadUrlResp = Invoke-RestMethod -Uri 'https://slack.com/api/files.getUploadURLExternal' -Method Post `
    -Headers $headers -Body $uploadUrlBody
if (-not $uploadUrlResp.ok) {
    throw "files.getUploadURLExternal failed: $($uploadUrlResp.error)"
}

# ---- Step 2: PUT the raw bytes to that URL ----
Invoke-RestMethod -Uri $uploadUrlResp.upload_url -Method Post -Body $htmlBytes -ContentType 'application/octet-stream' | Out-Null

# ---- Step 3: complete the upload, share into the channel with the message as initial_comment ----
$completePayload = @{
    files           = @(@{ id = $uploadUrlResp.file_id; title = $htmlFileName })
    channel_id      = $Channel
    initial_comment = $text
} | ConvertTo-Json -Depth 5
$completePayloadBytes = [System.Text.Encoding]::UTF8.GetBytes($completePayload)
$response = Invoke-RestMethod -Uri 'https://slack.com/api/files.completeUploadExternal' -Method Post `
    -Headers $headers -ContentType 'application/json; charset=utf-8' -Body $completePayloadBytes

if (-not $response.ok) {
    throw "Slack API returned an error: $($response.error) - common causes: bot not invited to a private channel (invite it, or use a public channel with chat:write.public), or the token/scopes are wrong (see SLACK_SETUP.md)."
}

Write-Host "Posted to #$Channel with HTML report attached." -ForegroundColor Green
Write-Host "This is a live post (Slack has no 'draft' concept), already visible in the channel." -ForegroundColor Yellow
