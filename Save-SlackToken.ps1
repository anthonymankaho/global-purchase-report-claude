<#
.SYNOPSIS
  One-time setup: prompts for the Slack Bot User OAuth Token (from
  https://api.slack.com/apps -> your app -> OAuth & Permissions, see SLACK_SETUP.md)
  and saves it encrypted to Credentials\slack_bot_token.enc.

.DESCRIPTION
  Uses Windows DPAPI (ConvertTo-SecureString / ConvertFrom-SecureString, no explicit
  key) - same technique as the Gmail OAuth token - so the encrypted file is tied to
  this Windows user account and machine, and can't be decrypted if copied elsewhere.

  Unlike Gmail's OAuth flow, Slack bot tokens obtained via "Install to Workspace" don't
  need a refresh-token dance - the token itself is what New-MonthlySlackPost.ps1 uses
  directly against the Slack Web API.
#>

$ErrorActionPreference = 'Stop'

$credDir = Join-Path $PSScriptRoot 'Credentials'
if (-not (Test-Path $credDir)) { New-Item -ItemType Directory -Path $credDir | Out-Null }
$tokenPath = Join-Path $credDir 'slack_bot_token.enc'

$secureInput = Read-Host -Prompt 'Paste the Slack Bot User OAuth Token (starts with xoxb-)' -AsSecureString
$bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureInput)
try {
    $plainToken = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
} finally {
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
}

if ($plainToken -notmatch '^xoxb-') {
    Write-Host "Warning: token doesn't start with 'xoxb-' - double-check you copied the Bot User OAuth Token, not the Signing Secret or App-Level Token." -ForegroundColor Yellow
}

$secure = ConvertTo-SecureString -String $plainToken -AsPlainText -Force
$encrypted = ConvertFrom-SecureString -SecureString $secure
Set-Content -Path $tokenPath -Value $encrypted -Encoding ASCII

Write-Host "Saved (encrypted): $tokenPath" -ForegroundColor Green
Write-Host "You can now run: .\New-MonthlySlackPost.ps1 -ReportMonth `"yyyy-MM`" -Channel `"channel-name`""
