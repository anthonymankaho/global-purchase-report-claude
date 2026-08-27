<#
.SYNOPSIS
  Creates a Gmail DRAFT (never sends) for the monthly Global Purchase Report -
  "PO Price Comparison" distribution email, via the Gmail REST API directly
  (gmail.compose scope only - this script never calls the send endpoint).

.DESCRIPTION
  Matches the reference format found in the original UiPath RPA's SendMailX activity
  (Main.xaml) and its supporting Project_Notebook.xlsx "Scratchpad" sheet: subject
  "Global purchase report (1-M)YYYY - PO Price Comparison", the same body wording,
  and the same confirmed 13-person distribution list (3 To + 10 Cc).

  Deliberately does NOT auto-attach the report file, matching the original RPA's own
  documented reason: the file is too large for email (the old RPA's own follow-up
  message told the operator to attach manually and share via Google Drive instead -
  Gmail's ~25MB limit, further inflated by base64 encoding, easily exceeded by this
  report). Only verifies the month's deliverable actually exists (as a readiness
  check), then creates the draft - attach it yourself in Gmail before sending.

.PARAMETER ReportMonth
  "yyyy-MM", e.g. "2026-07". Used to build both the "(1-M)YYYY" label in the subject
  and to locate/verify this month's deliverable exists.

.PARAMETER DeliverablePath
  Path to this month's outside-delivery deliverable, just to confirm it exists before
  drafting (never attached - see above). Defaults to the standard naming convention.

.PARAMETER To / -Cc
  Override the default 13-person distribution list, e.g. for a test run to yourself only.

.EXAMPLE
  .\New-MonthlyEmailDraft.ps1 -ReportMonth "2026-07"
  .\New-MonthlyEmailDraft.ps1 -ReportMonth "2026-07" -To "anthonym@meyer-msc.com" -Cc @()
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d{4}-\d{2}$')]
    [string]$ReportMonth,

    [string]$DeliverablePath,

    [string[]]$To = @(
        'joe@meyer-mil.com', 'stevenc@meyer-mhk.com', 'sl@meyer-mmc.com'
    ),
    [string[]]$Cc = @(
        'ksk@meyer-mil.com', 'aprilz@meyer-mmc.com', 'spchoi@meyer-mmc.com', 'hugo@meyer-mil.com',
        'Wscheung@meyer-mil.com', 'alex@meyer-msc.com', 'walker@meyer-mil.com', 'pccken@silverstone-stc.com',
        'michaelt@meyer-msc.com', 'fiona@meyer-msc.com', 'chengpuicheeken@gmail.com', 'adaw@meyer-mhk.com',
        'kenny@meyer-mzq.com'
    )
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'GmailOAuth.Common.ps1')

$year = [int]$ReportMonth.Substring(0, 4)
$month = [int]$ReportMonth.Substring(5, 2)
$periodLabel = "(1-$month)$year"   # matches Publish-OutsideReport.ps1's own $periodLabel format exactly

if (-not $DeliverablePath) {
    # Just a readiness check (latest RV for the month) - never attached, see .DESCRIPTION.
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

Write-Host "Deliverable found (readiness check only, NOT attached): $DeliverablePath"
Write-Host "Recipients   : $($To.Count) in To, $($Cc.Count) in Cc"
Write-Host "Subject      : Global purchase report $periodLabel - PO Price Comparison"

$accessToken = Get-GmailAccessToken

$bodyText = @"
<html>
<body style="font-family: Calibri, sans-serif; font-size: 11pt; color: #000000;">
<p>*** THIS IS AN AUTOMATED EMAIL. PLEASE DO NOT REPLY ***</p>

<p>Dear All,</p>

<p>The attachment is the PO Price Comparison Tables for Fittings, Glass Lids, Packaging and Coating items.</p>

<p>If you have any questions, please feel free to contact Anthony Man (anthonym@meyer-msc.com).</p>

<p>Best Regards,<br><br>
Anthony<br>
Finance Department</p>
</body>
</html>
"@

$boundary = "----=_GlobalPurchaseReport_$([guid]::NewGuid().ToString('N'))"
$toHeader = ($To -join ', ')
$ccHeader = ($Cc -join ', ')
$subject = "Global purchase report $periodLabel - PO Price Comparison"

$mimeLines = New-Object 'System.Collections.Generic.List[string]'
[void]$mimeLines.AddRange([string[]]@(
    "To: $toHeader",
    "Cc: $ccHeader",
    "Subject: $subject",
    "MIME-Version: 1.0",
    "Content-Type: multipart/mixed; boundary=`"$boundary`"",
    "",
    "--$boundary",
    "Content-Type: text/html; charset=`"UTF-8`"",
    "",
    $bodyText,
    "",
    "--$boundary--"
))
$mime = $mimeLines -join "`r`n"

$rawBytes = [System.Text.Encoding]::UTF8.GetBytes($mime)
$rawBase64Url = [Convert]::ToBase64String($rawBytes).Replace('+', '-').Replace('/', '_').Replace('=', '')

$payload = @{ message = @{ raw = $rawBase64Url } } | ConvertTo-Json -Depth 5
$payloadBytes = [System.Text.Encoding]::UTF8.GetBytes($payload)

$headers = @{ Authorization = "Bearer $accessToken" }
$response = Invoke-RestMethod -Uri 'https://gmail.googleapis.com/gmail/v1/users/me/drafts' -Method Post `
    -Headers $headers -ContentType 'application/json; charset=utf-8' -Body $payloadBytes

Write-Host "Draft created - id: $($response.id)" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps (matches the original RPA's own manual step - the report file is too large to email as an attachment):" -ForegroundColor Yellow
Write-Host "  1. Open Gmail, find this draft, manually attach: $DeliverablePath" -ForegroundColor Yellow
Write-Host "  2. Share the report file via Google Drive (Editor access) and reference it in the email if needed." -ForegroundColor Yellow
Write-Host "  3. Review recipients and content, then send yourself." -ForegroundColor Yellow
