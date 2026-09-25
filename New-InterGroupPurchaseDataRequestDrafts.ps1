<#
.SYNOPSIS
  Creates 5 separate Gmail DRAFTS (never sends) requesting monthly purchase data from
  5 inter-group company entities, via the Gmail REST API directly (gmail.compose scope
  only - this script never calls the send endpoint).

.DESCRIPTION
  Each draft is addressed to a single company contact, with subject/body translated
  into that recipient's language (natural business-email phrasing, not a literal
  word-for-word translation), and the same Cc list on all 5.

.PARAMETER ReportMonth
  "yyyy-MM", e.g. "2026-09" - the month of purchase data to request.

.PARAMETER Only
  Company code(s) to limit this run to, e.g. -Only MZQ for a single test draft.
  Defaults to all 5: MZQ, MIL, RACO, TVS, MAZ.

.PARAMETER Cc
  Overrides the default Cc list (Fiona Yip, Alex Wong, Steven Cheng), applied to
  every draft created by this run.

.EXAMPLE
  .\New-InterGroupPurchaseDataRequestDrafts.ps1 -ReportMonth "2026-09" -Only MZQ
  .\New-InterGroupPurchaseDataRequestDrafts.ps1 -ReportMonth "2026-09"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d{4}-(0[1-9]|1[0-2])$')]
    [string]$ReportMonth,

    [ValidateSet('MZQ', 'MIL', 'RACO', 'TVS', 'MAZ')]
    [string[]]$Only,

    [string[]]$Cc = @('fiona@meyer-msc.com', 'alex@meyer-msc.com', 'stevenc@meyer-mhk.com')
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'GmailOAuth.Common.ps1')

$year = [int]$ReportMonth.Substring(0, 4)
$month = [int]$ReportMonth.Substring(5, 2)

$monthNamesEn = @('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December')
$monthNamesIt = @('gennaio', 'febbraio', 'marzo', 'aprile', 'maggio', 'giugno', 'luglio', 'agosto', 'settembre', 'ottobre', 'novembre', 'dicembre')
$monthNamesEs = @('enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre')

$companies = @(
    [pscustomobject]@{ Code = 'MZQ';  To = 'nikijiao@meyer-mzq.com'; Attention = 'Niki Jiao'; Language = 'zh-Hans' }
    [pscustomobject]@{ Code = 'MIL';  To = 'ksk@meyer-mil.com';      Attention = 'KSK';        Language = 'en' }
    [pscustomobject]@{ Code = 'RACO'; To = 'yunyu@raco-ril.com';     Attention = '喻芸';        Language = 'zh-Hans' }
    [pscustomobject]@{ Code = 'TVS';  To = 'o.demarco@tvs-spa.it';   Attention = 'Olimpia';    Language = 'it' }
    [pscustomobject]@{ Code = 'MAZ';  To = 'aconterno@mazalcorp.com'; Attention = 'Alfredo';   Language = 'es' }
)

if ($Only) {
    $companies = $companies | Where-Object { $_.Code -in $Only }
}

function Get-EmailContent {
    param($Code, $Attention, $Language, $Year, $Month)

    switch ($Language) {
        'en' {
            $monthName = $monthNamesEn[$Month - 1]
            $subject = "Request $Code Purchase Data $monthName $Year"
            $body = "<p>Hi $Attention,</p><p>Would you mind sharing the $Code purchase data for the period of $monthName ${Year}?</p><p>Best regards,<br>Anthony</p>"
        }
        'zh-Hans' {
            $subject = "请求提供 $Code ${Year}年${Month}月采购数据"
            $body = "<p>$Attention，你好：</p><p>请问是否方便提供 $Code ${Year}年${Month}月份的采购数据？</p><p>谢谢！</p><p>Anthony</p>"
        }
        'it' {
            $monthName = $monthNamesIt[$Month - 1]
            $subject = "Richiesta dati di acquisto $Code - $monthName $Year"
            $body = "<p>Gentile $Attention,</p><p>potrebbe cortesemente condividere i dati di acquisto di $Code relativi al periodo di $monthName ${Year}?</p><p>Cordiali saluti,<br>Anthony</p>"
        }
        'es' {
            $monthName = $monthNamesEs[$Month - 1]
            $subject = "Solicitud de datos de compras de $Code - $monthName de $Year"
            $body = "<p>Hola $Attention,</p><p>¿Podrías compartir los datos de compras de $Code correspondientes al periodo de $monthName de ${Year}?</p><p>Saludos cordiales,<br>Anthony</p>"
        }
        default { throw "Unsupported language: $Language" }
    }

    return @{ Subject = $subject; Body = $body }
}

function ConvertTo-MimeEncodedHeader {
    param([string]$Text)
    if ($Text -match '[^\x00-\x7F]') {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $b64 = [Convert]::ToBase64String($bytes)
        return "=?UTF-8?B?${b64}?="
    }
    return $Text
}

$accessToken = Get-GmailAccessToken
$ccHeader = ($Cc -join ', ')

foreach ($company in $companies) {
    $content = Get-EmailContent -Code $company.Code -Attention $company.Attention -Language $company.Language -Year $year -Month $month
    $subject = $content.Subject
    $encodedSubject = ConvertTo-MimeEncodedHeader -Text $subject

    $bodyHtml = @"
<html>
<body style="font-family: Calibri, sans-serif; font-size: 11pt; color: #000000;">
$($content.Body)
</body>
</html>
"@

    $boundary = "----=_InterGroupPurchaseDataRequest_$([guid]::NewGuid().ToString('N'))"
    $mimeLines = New-Object 'System.Collections.Generic.List[string]'
    [void]$mimeLines.AddRange([string[]]@(
        "To: $($company.To)",
        "Cc: $ccHeader",
        "Subject: $encodedSubject",
        "MIME-Version: 1.0",
        "Content-Type: multipart/mixed; boundary=`"$boundary`"",
        "",
        "--$boundary",
        "Content-Type: text/html; charset=`"UTF-8`"",
        "",
        $bodyHtml,
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

    Write-Host "[$($company.Code)] Draft created - id: $($response.id) | To: $($company.To) | Subject: $subject" -ForegroundColor Green
}

Write-Host ""
Write-Host "Done. Review each draft in Gmail before sending." -ForegroundColor Yellow
