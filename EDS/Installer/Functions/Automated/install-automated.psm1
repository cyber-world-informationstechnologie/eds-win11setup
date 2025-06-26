function Install-Automated {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]$EDS_Server = "https://example.com/eds",
        [Parameter(Mandatory=$false)]
        [switch]$DryRun = $false,
        [Parameter(Mandatory=$false)]
        [string]$WinPeDrive = "X:",
        [Parameter(Mandatory=$false)]
        [string]$EDSFolderName = "CWI",
        [switch]$Debugging = $false
    )

    # Check if the EDS server URL is provided
    if (-not $EDS_Server) {
        Write-Host "No EDS server URL provided. Exiting installer."
        return
    }

    # Validate the EDS server URL format
    if (-not $EDS_Server -or $EDS_Server -notmatch '^https?://') {
        Write-Host "Invalid EDS server URL format: $EDS_Server. Exiting installer."
        return
    }   
 
 # Retry logic for EDS server connectivity
    $maxRetries = 6
    $retryDelay = 10 # seconds
    $attempt = 0
    $connected = $false
    while ($attempt -lt $maxRetries -and -not $connected) {
        try {
            $response = Invoke-WebRequest -Uri $EDS_Server -UseBasicParsing -TimeoutSec 5
            if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 400) {
                $connected = $true
                break
            } else {
                Write-Host "EDS server $EDS_Server is not reachable (HTTP $($response.StatusCode)). Retrying in $retryDelay seconds..."
            }
        } catch {
            Write-Host "EDS server $EDS_Server is not reachable. Retrying in $retryDelay seconds..."
        }
        $attempt++
        if ($attempt -lt $maxRetries) {
            Start-Sleep -Seconds $retryDelay
        }
    }
    if (-not $connected) {
        Write-Host "EDS server $EDS_Server is not reachable after $($maxRetries * $retryDelay) seconds. Exiting installer."
        return
    }

    # Generate a random 6-digit number for the device name
    $randomNumber = Get-Random -Minimum 100000 -Maximum 999999
    $deviceName = "EDS-Auto-$randomNumber"

    if ($Debugging) {
        $serialNumber = "SN-$randomNumber"  # Placeholder for serial number
        Write-Host "Debugging mode enabled. Using random serial number: $serialNumber"
    } else {
        $serialNumber = Get-SerialNumber
    }

    $registerInfo = @{
        deviceName = $deviceName
        deviceType = Get-DeviceType
        deviceSerial = $serialNumber
    }

    $guiPath = "$PSScriptRoot\..\..\GUI"

    . "$guiPath\AutomatedForm.ps1"
    
    $automatedForm = Initialize-AutomatedForm -DryRun $DryRun

    # Register device and let Register-Device handle all error dialogs and exceptions
    $registerResponse = $null
    try {
        $registerResponse = Register-Device -Uri $EDS_Server -registerInfo $registerInfo
        if ($null -eq $registerResponse) {
            Set-Infotext "Device registration failed. Exiting..."
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Seconds 2
            $automatedForm.Close()
            Write-Host "Registration failed, exiting installer."
            return
        }
    } catch {
        Set-Infotext "Device registration failed. Exiting..."
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Seconds 2
        $automatedForm.Close()
        Write-Host "Registration failed, exiting installer."
        return
    }

    Set-Infotext "Device registered successfully"
    Set-JobIdText -JobId $registerResponse.jobId
    # Set job name label if available
    Set-JobNameText -JobName $registerResponse.jobName
    
    Start-Sleep -Seconds 2

    $deviceToken = $registerResponse.deviceToken
    $jobId = $registerResponse.jobId

    Set-JobStatus -jobId $jobId -status "waiting_for_instructions" -deviceToken $deviceToken -uri $EDS_Server

    # Update GUI for polling
    Set-Infotext "Waiting for instructions. Rechecking in 15 seconds..."

    $headers = @{
        "Content-Type"    = "application/json"
        "X-Device-Token"  = $deviceToken
    }

    $jobInstructionsUrl = "$EDS_Server/api/jobs/$jobId/instructions"
    $polling = $true
    $jobContext

    while ($polling) {
        Set-ProgressBar 0
        if ($automatedForm.Tag -eq 'skip') { break }
        try {
            $response = Invoke-WebRequest -Uri $jobInstructionsUrl -UseBasicParsing -TimeoutSec 10 -Headers $headers
            if ($response.StatusCode -eq 200 -and $response) {
                $json = $null
                try { $json = $response.Content | ConvertFrom-Json } catch {}
                if ($json -and $($json.action) -and $($json.action) -ne "WAIT_FOR_INSTRUCTIONS") {
                    $automatedForm.Tag = 'instructions'
                    $polling = $false
                    Write-Host "Received action: $($json.action)"
                    $jobContext = $($json.context)
                    break
                }
            }
        } catch {
            Write-Host "Error checking for instructions: $_"
            break
        }
        for ($i=0; $i -lt 15; $i++) {
            Start-Sleep -Seconds 1
            Set-ProgressBar ($i + 1)
            [System.Windows.Forms.Application]::DoEvents()
            if ($automatedForm.Tag -eq 'skip') { $polling = $false; break }
        }
        Set-Infotext "Waiting for instructions. Rechecking in 15 seconds..."
        [System.Windows.Forms.Application]::DoEvents()
    }

    if ($automatedForm.Tag -eq 'skip') {
        Write-Host "User chose to skip auto-configuration. Starting manual configuration."
    } elseif ($automatedForm.Tag -eq 'instructions') {
        Write-Host "Instructions received from server. Proceeding with automated configuration."
        Set-Infotext "Instructions received from server. Proceeding with automated configuration."
        Start-Sleep -Seconds 2

        $automatedForm.Close()

        Set-JobStatus -jobId $jobId -status "installing" -deviceToken $deviceToken -uri $EDS_Server

        # Load unattend.xml from disk and update device name
        $unattendPath = Join-Path $WinPeDrive 'Temp/unattended.xml'
        [xml]$unattendXml = Get-Content -Path $unattendPath

        Write-Host "Updating device name in unattend.xml to $($jobContext.deviceName)"
        Set-UnattendedDeviceName -deviceName $jobContext.deviceName -xmlDoc $unattendXml
        Set-LocalAccount -xmlDoc $unattendXml -UserName 'root' -PasswordBase64 $jobContext.localPassword
        # Convert jobContext to hashtable for Set-UnattendedUserInput
        $jobContextHash = @{
            deviceToken = $deviceToken
            jobId = $jobId
            apiUrl = "$EDS_Server/api"
            edsFolderName = $EDSFolderName
        }
        $jobContext.PSObject.Properties | ForEach-Object { $jobContextHash[$_.Name] = $_.Value }
        $jobContextHash['localPassword'] = ""
        Write-Host "Updating user input in unattend.xml"
        Set-UnattendedUserInput -xmlDoc $unattendXml -UserInput $jobContextHash
    
        Write-Host "Starting installation process for device"
        if ($DryRun -ne $true) {
            $installDrive = Get-InstallationDrive -EDSFolderName $EDSFolderName

            Write-Host "Starting installation with unattended.xml $unattendPath"
            New-Item -Path "$installDrive\Temp" -ItemType Directory -Force | Out-Null
            # Check if autounattend.xml already exists in the root of the install drive
            $autoUnattendPath = Join-Path $installDrive 'autounattend.xml'
            $autoUnattendOldPath = Join-Path $installDrive 'autounattend_old.xml'
            if (Test-Path $autoUnattendPath) {
                Move-Item -Path $autoUnattendPath -Destination $autoUnattendOldPath -Force
            }
            # Copy the new unattend.xml to the root of the install drive as autounattend.xml
            Copy-Item -Path $unattendPath -Destination $autoUnattendPath -Force

            # Also save a backup in the Temp folder
            Copy-Item -Path $unattendPath -Destination "$installDrive\Temp\unattended_save.xml" -Force
            Start-Process -FilePath "$WinPeDrive\setup.exe" -ArgumentList "/unattend:$unattendPath" -NoNewWindow
        } else {
            Write-Host "Skipping actual setup because of Dry-Run"
        }

    } else {
        Write-Host "Form closed without action."
    }

    $automatedForm.Close()
    $automatedForm.Dispose()
}

Export-ModuleMember -Function Install-Automated
