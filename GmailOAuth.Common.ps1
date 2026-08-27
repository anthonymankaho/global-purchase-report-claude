<#
Shared helpers for Gmail draft automation, dot-sourced by New-MonthlyEmailDraft.ps1.

Reuses the same Google Cloud OAuth app / gmail.compose-scope authorization already set
up for the Weekly Exchange Report project (same Google account, same scope) - the token
file was copied from there rather than running a fresh Connect-GmailOAuth.ps1 flow.

Tokens are stored encrypted via Windows DPAPI (ConvertTo-SecureString / ConvertFrom-
SecureString with no explicit key) - this ties the encrypted file to the current Windows
user account and machine, so it can't be decrypted if copied elsewhere. That's exactly
why copying the .enc file between projects on the SAME machine/account is safe.
#>

$script:GmailTokenPath = Join-Path $PSScriptRoot 'Credentials\gmail_oauth_token.enc'

function Get-GmailTokenData {
    if (-not (Test-Path $script:GmailTokenPath)) {
        throw "No stored Gmail authorization found at $script:GmailTokenPath - copy gmail_oauth_token.enc from the Weekly Exchange Report - CLAUDE project's Credentials folder, or run a fresh OAuth flow."
    }
    # .Trim() matters here - Set-Content appends a trailing newline that Get-Content -Raw
    # includes verbatim, and ConvertTo-SecureString rejects the encrypted string if it has
    # any extra trailing whitespace/newline characters.
    $encrypted = (Get-Content -Path $script:GmailTokenPath -Raw).Trim()
    $secure = ConvertTo-SecureString -String $encrypted
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        $json = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    } finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
    return $json | ConvertFrom-Json
}

function Get-GmailAccessToken {
    $data = Get-GmailTokenData
    $body = @{
        client_id     = $data.client_id
        client_secret = $data.client_secret
        refresh_token = $data.refresh_token
        grant_type    = 'refresh_token'
    }
    try {
        $response = Invoke-RestMethod -Uri 'https://oauth2.googleapis.com/token' -Method Post -Body $body
    } catch {
        throw "Failed to refresh the Gmail access token - the stored authorization may have expired or been revoked. Original error: $($_.Exception.Message)"
    }
    return $response.access_token
}
