# Simple PowerShell Web Service for ISO Creation
# Requires: PowerShell 5.1+
# Usage: Run as administrator

# Requires powershell-yaml module
Import-Module powershell-yaml -ErrorAction Stop

$configPath = Join-Path $PSScriptRoot 'WebService.config.yaml'
if (-not (Test-Path $configPath)) {
    throw "Config file not found: $configPath"
}
$config = ConvertFrom-Yaml (Get-Content $configPath -Raw)

$AuthHeader = $config.auth_header
$Port = $config.port
$InputBase = $config.input_base
$OutputBase = $config.output_base
$InstallerScript = $config.installer_script
$EdsFolder = $config.eds_folder

# If installer_script is not set, default to same folder as WebService.ps1
if (-not $InstallerScript -or $InstallerScript -eq "") {
    $InstallerScript = Join-Path $PSScriptRoot 'Create-Windows11Installer.ps1'
}

Add-Type -AssemblyName System.Net.HttpListener
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://*:$Port/")
$listener.Start()
Write-Host "Web service started on port $Port" -ForegroundColor Green

function Send-Callback {
    param(
        [string]$CallbackUrl,
        [string]$IsoName
    )
    $body = @{ iso_name = $IsoName } | ConvertTo-Json
    try {
        Invoke-RestMethod -Uri $CallbackUrl -Method Post -Body $body -ContentType 'application/json'
        Write-Host "Callback sent to $CallbackUrl" -ForegroundColor Green
    } catch {
        Write-Host "Failed to send callback: $($_.Exception.Message)" -ForegroundColor Red
    }
}

$global:Processing = $false

while ($listener.IsListening) {
    $context = $listener.GetContext()
    $request = $context.Request
    $response = $context.Response
    $authHeader = $request.Headers["Authorization"]
    if ($authHeader -ne $AuthHeader) {
        $response.StatusCode = 401
        $response.Close()
        continue
    }
    if ($request.HttpMethod -eq "POST") {
        if ($global:Processing) {
            $response.StatusCode = 429
            $msg = "Server is busy processing another image. Please wait and try again."
            $response.OutputStream.Write([Text.Encoding]::UTF8.GetBytes($msg), 0, $msg.Length)
            $response.Close()
            continue
        }
        $global:Processing = $true
        try {
            $body = $null
            try {
                $reader = New-Object IO.StreamReader($request.InputStream)
                $body = $reader.ReadToEnd() | ConvertFrom-Json
            } catch {}
            $isoPath = $body.iso_path
            $callbackUrl = $body.callback_url
            $cuPath = $body.cu_path
            $winEdition = $body.win_edition
            if (-not $isoPath) {
                $isoPath = $request.QueryString["iso_path"]
            }
            if (-not $callbackUrl) {
                $callbackUrl = $request.QueryString["callback_url"]
            }
            if (-not $cuPath) {
                $cuPath = $request.QueryString["cu_path"]
            }
            if (-not $winEdition) {
                $winEdition = $request.QueryString["win_edition"]
            }
            if (-not $winEdition) {
                $winEdition = "Windows 11 Pro"
            }
            $uid = [guid]::NewGuid().ToString()
            $isoName = "winiso_$uid.iso"
            if (-not (Test-Path $OutputBase)) { New-Item -ItemType Directory -Path $OutputBase | Out-Null }
            $isoOutputPath = Join-Path $OutputBase $isoName
            # Build installer command
            $cmd = @()
            $cmd += "& '$InstallerScript'"
            $cmd += "-ISOPath '$isoPath' -SkipDownload -SkipUSBCreation -CreateISO -ISOOutputPath '$isoOutputPath'"
            if ($cuPath) { $cmd += "-CUPath '$cuPath'" }
            if ($winEdition) { $cmd += "-WinEdition '$winEdition'" }
            if ($EdsFolder -and $EdsFolder -ne "") { $cmd += "-EDSFolder '$EdsFolder'" }
            $fullCmd = $cmd -join ' '
            Write-Host "Starting ISO creation: $fullCmd" -ForegroundColor Cyan
            try {
                Invoke-Expression $fullCmd
                Send-Callback -CallbackUrl $callbackUrl -IsoName $isoName
                $response.StatusCode = 200
                $response.OutputStream.Write([Text.Encoding]::UTF8.GetBytes("ISO creation started. Callback will be sent when done."), 0, 62)
            } catch {
                $response.StatusCode = 500
                $response.OutputStream.Write([Text.Encoding]::UTF8.GetBytes("Error: $($_.Exception.Message)"), 0, 7)
            }
            $response.Close()
        } finally {
            $global:Processing = $false
        }
    } else {
        $response.StatusCode = 405
        $response.Close()
    }
}
