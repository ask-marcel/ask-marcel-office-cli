# ask-marcel-office device-code login
# Bypasses SentinelOne/EDR + v2rayN proxy. Uses OAuth 2.0 Device Authorization Grant.
# Run: powershell -File poc-device-code.ps1

param()
$ErrorActionPreference = "Stop"

$CLIENT_ID = "27922004-5251-4030-b22d-91ecd9a37ea4"
$SCOPES = "openid profile offline_access https://graph.microsoft.com/.default"
$CACHE_DIR = Join-Path $env:USERPROFILE ".ask-marcel"
$CACHE_PATH = Join-Path $CACHE_DIR "token-cache.json"

# -- Helper: HTTP POST with proxy bypass ------------------------------------
function HttpPost($url, $body) {
    Add-Type -AssemblyName System.Web
    $webReq = [System.Net.WebRequest]::Create($url)
    $webReq.Method = "POST"
    $webReq.ContentType = "application/x-www-form-urlencoded"
    $webReq.Proxy = [System.Net.GlobalProxySelection]::GetEmptyWebProxy()
    
    $bodyStr = ($body.GetEnumerator() | ForEach-Object { 
        "$($_.Key)=$([System.Web.HttpUtility]::UrlEncode($_.Value))" 
    }) -join "&"
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($bodyStr)
    $webReq.ContentLength = $bytes.Length
    $reqStream = $webReq.GetRequestStream()
    $reqStream.Write($bytes, 0, $bytes.Length)
    $reqStream.Close()
    
    try {
        $webResp = $webReq.GetResponse()
        $respStream = $webResp.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($respStream)
        $respBody = $reader.ReadToEnd()
        $reader.Close(); $webResp.Close()
        return $respBody | ConvertFrom-Json
    } catch [System.Net.WebException] {
        if ($_.Exception.Response) {
            $respStream = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($respStream)
            $respBody = $reader.ReadToEnd()
            $reader.Close()
            return $respBody | ConvertFrom-Json
        }
        throw
    }
}

# -- Proxy cleanup -----------------------------------------------------------
Write-Host "=== ask-marcel-office login (device code) ===" -ForegroundColor Cyan
Write-Host ""

$regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
Remove-ItemProperty -Path $regPath -Name ProxyServer -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $regPath -Name AutoConfigURL -ErrorAction SilentlyContinue
Remove-Item Env:\HTTP_PROXY -ErrorAction SilentlyContinue
Remove-Item Env:\HTTPS_PROXY -ErrorAction SilentlyContinue
Remove-Item Env:\http_proxy -ErrorAction SilentlyContinue
Remove-Item Env:\https_proxy -ErrorAction SilentlyContinue

Write-Host "  Proxy settings cleared" -ForegroundColor Green
Write-Host ""

# -- Request device code -----------------------------------------------------
$dc = HttpPost "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode" @{
    client_id = $CLIENT_ID
    scope = $SCOPES
}

if ($dc.error) {
    Write-Host "FAILED: $($dc.error) - $($dc.error_description)" -ForegroundColor Red
    exit 1
}

Write-Host "--------------------------------------------------"
Write-Host "  Open this URL:  $($dc.verification_uri)"
Write-Host "  Enter code:     $($dc.user_code)"
Write-Host "--------------------------------------------------"
Write-Host ""
Write-Host "  Waiting for sign-in (expires in $($dc.expires_in)s)..."

# -- Poll for token -----------------------------------------------------------
$start = Get-Date
$timeout = [int]$dc.expires_in + 10

while (((Get-Date) - $start).TotalSeconds -lt $timeout) {
    $data = HttpPost "https://login.microsoftonline.com/common/oauth2/v2.0/token" @{
        grant_type = "urn:ietf:params:oauth:grant-type:device_code"
        client_id = $CLIENT_ID
        device_code = $dc.device_code
    }

    if ($data.access_token) {
        Write-Host ""
        Write-Host "  Authenticated!" -ForegroundColor Green

        $parts = $data.access_token.Split(".")
        $b64 = $parts[1].Replace("-", "+").Replace("_", "/")
        while ($b64.Length % 4 -ne 0) { $b64 += "=" }
        $payload = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($b64))
        $claims = $payload | ConvertFrom-Json

        Write-Host "  UPN:    $($claims.upn)"
        Write-Host "  Scopes: $($claims.scp)"
        Write-Host "  Expires in: $($data.expires_in)s"

        New-Item -ItemType Directory -Force -Path $CACHE_DIR | Out-Null
        $expiresOn = [int][double]::Parse(
            (Get-Date).AddSeconds([int]$data.expires_in).ToUniversalTime().Subtract(
                (Get-Date "1970-01-01")
            ).TotalSeconds.ToString()
        )
        @{
            access_token = $data.access_token
            refresh_token = $data.refresh_token
            expires_on = $expiresOn
        } | ConvertTo-Json -Depth 2 | Set-Content $CACHE_PATH

        Write-Host ""
        Write-Host "  Token cache saved to $CACHE_PATH" -ForegroundColor Green
        exit 0
    }

    if ($data.error -eq "authorization_pending") {
        Write-Host "." -NoNewline
        Start-Sleep 5
        continue
    }
    if ($data.error -eq "slow_down") {
        Start-Sleep 10
        continue
    }

    Write-Host ""
    Write-Host "  Error: $($data.error) - $($data.error_description)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "  Timed out" -ForegroundColor Red
exit 1