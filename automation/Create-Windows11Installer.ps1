<#
.SYNOPSIS
    Automates the creation of a customized Windows 11 installer with EDS integration.

.DESCRIPTION
    This script downloads the latest Windows 11 ISO, creates a bootable USB drive or ISO file,
    and applies all necessary modifications as described in the project README.
    
.PARAMETER OutputPath
    Path where temporary files and the final installer will be created.
    
.PARAMETER USBDrive
    Optional: Specify the USB drive letter (e.g., 'E'). If not specified, will prompt interactively.
    
.PARAMETER SkipDownload
    If specified, skips the ISO download and uses an existing ISO file.
    
.PARAMETER ISOPath
    Path to an existing Windows 11 ISO file (required if SkipDownload is used).

.PARAMETER SkipUSBCreation
    If specified, only prepares the files but does not create a USB drive.

.PARAMETER CreateISO
    If specified, creates a bootable ISO file from the prepared installer files.
    Requires Windows ADK to be installed.

.PARAMETER ISOOutputPath
    Path where the output ISO file should be created. If not specified, uses the OutputPath directory.

.EXAMPLE
    .\Create-Windows11Installer.ps1
    
.EXAMPLE
    .\Create-Windows11Installer.ps1 -ISOPath "C:\ISOs\Win11.iso" -SkipDownload -USBDrive "E"

.EXAMPLE
    .\Create-Windows11Installer.ps1 -ISOPath "C:\ISOs\Win11.iso" -SkipDownload -SkipUSBCreation -CreateISO -ISOOutputPath "C:\Output\Win11_Custom.iso"
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputPath = "$env:TEMP\WindowsInstallerBuild",
    
    [Parameter()]
    [string]$USBDrive,
    
    [Parameter()]
    [switch]$SkipDownload,
    
    [Parameter()]
    [string]$ISOPath,

    [Parameter()]
    [switch]$SkipUSBCreation,

    [Parameter()]
    [switch]$CreateISO,

    [Parameter()]
    [string]$ISOOutputPath,

    [Parameter()]
    [string]$EDSFolder
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

# Helper functions for colored output
function Write-Step {
    param([string]$Message, [string]$Color = 'Cyan')
    Write-Host "`n==> $Message" -ForegroundColor $Color
}

function Write-Success {
    param([string]$Message)
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function Write-Warn {
    param([string]$Message)
    Write-Host "[WARNING] $Message" -ForegroundColor Yellow
}

function Write-Fail {
    param([string]$Message)
    Write-Host "[ERROR]- $Message" -ForegroundColor Red
}

# Check for admin privileges
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Download Windows 11 ISO using Microsoft's official method
function Get-Windows11ISO {
    param(
        [string]$OutputPath,
        [string]$Language = "en-us",
        [string]$Edition = "Windows 11"
    )
    
    Write-Step "Downloading Windows 11 ISO"
    Write-Host "Language: $Language"
    Write-Host "Edition: $Edition"
    
    $isoFileName = "Win11_$(Get-Date -Format 'yyyyMMdd')_${Language}_x64.iso"
    $isoPath = Join-Path $OutputPath $isoFileName
    
    Write-Warn "Microsoft does not provide direct download links via API."
    Write-Host "Please download the Windows 11 ISO manually from:"
    Write-Host "https://www.microsoft.com/software-download/windows11" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "After downloading, press any key to continue and select the ISO file..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "ISO files (*.iso)|*.iso|All files (*.*)|*.*"
    $openFileDialog.Title = "Select Windows 11 ISO file"
    
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $selectedIso = $openFileDialog.FileName
        Write-Success "Selected ISO: $selectedIso"
        return $selectedIso
    } else {
        throw "No ISO file selected. Cannot continue."
    }
}

# Mount ISO and return drive letter
function Mount-ISOImage {
    param([string]$ISOPath)
    
    Write-Step "Mounting ISO image"
    $mountResult = Mount-DiskImage -ImagePath $ISOPath -PassThru
    $driveLetter = ($mountResult | Get-Volume).DriveLetter
    
    if (-not $driveLetter) {
        throw "Failed to mount ISO or retrieve drive letter"
    }
    
    Write-Success "ISO mounted to ${driveLetter}:\"
    return "${driveLetter}:"
}

# Get available USB drives
function Get-USBDrives {
    $result = @()
    $usbDisks = Get-Disk | Where-Object { $_.BusType -eq 'USB' }
    
    foreach ($disk in $usbDisks) {
        $partitions = Get-Partition -DiskNumber $disk.Number -ErrorAction SilentlyContinue
        $hasValidPartition = $false
        
        foreach ($partition in $partitions) {
            $volume = Get-Volume -Partition $partition -ErrorAction SilentlyContinue
            if ($volume -and $volume.DriveLetter) {
                $result += [PSCustomObject]@{
                    DriveLetter = $volume.DriveLetter
                    DiskNumber = $disk.Number
                    FriendlyName = $disk.FriendlyName
                    Size = [math]::Round($disk.Size / 1GB, 2)
                    VolumeName = $volume.FileSystemLabel
                }
                $hasValidPartition = $true
            }
        }
        
        # If USB disk has no partition with drive letter, still include it
        if (-not $hasValidPartition) {
            $result += [PSCustomObject]@{
                DriveLetter = $null
                DiskNumber = $disk.Number
                FriendlyName = $disk.FriendlyName
                Size = [math]::Round($disk.Size / 1GB, 2)
                VolumeName = "(No formatted partition)"
            }
        }
    }
    
    return $result
}

# Prepare USB drive
function Initialize-USBDrive {
    param(
        [string]$DriveLetter,
        [int]$DiskNumber
    )
    
    Write-Step "Preparing USB drive ${DriveLetter}:"
    Write-Warn "This will ERASE ALL DATA on the drive!"
    
    $confirm = Read-Host "Type 'YES' to confirm"
    if ($confirm -ne 'YES') {
        throw "USB drive preparation cancelled by user"
    }
    
    Write-Host "Cleaning disk..."
    Clear-Disk -Number $DiskNumber -RemoveData -RemoveOEM -Confirm:$false
    
    Write-Host "Creating partition (32GB for FAT32 compatibility)..."
    # Create a 32GB partition to ensure FAT32 compatibility
    # Windows 11 installer typically needs ~6-8GB, so 32GB is more than enough
    $partitionSize = 32GB
    $diskSize = (Get-Disk -Number $DiskNumber).Size
    
    # If disk is smaller than 32GB, use maximum size
    if ($diskSize -lt $partitionSize) {
        $partition = New-Partition -DiskNumber $DiskNumber -UseMaximumSize -AssignDriveLetter
    } else {
        $partition = New-Partition -DiskNumber $DiskNumber -Size $partitionSize -AssignDriveLetter
    }
    
    $newDriveLetter = $partition.DriveLetter
    
    Write-Host "Formatting as FAT32..."
    Format-Volume -DriveLetter $newDriveLetter -FileSystem FAT32 -NewFileSystemLabel "WIN11_USB" -Confirm:$false | Out-Null
    
    Write-Success "USB drive prepared: ${newDriveLetter}:\"
    return "${newDriveLetter}:"
}

# Copy Windows files to USB (excluding install.wim)
function Copy-WindowsFiles {
    param(
        [string]$SourceDrive,
        [string]$TargetDrive,
        [bool]$ExcludeInstallWim = $true
    )
    
    Write-Step "Copying Windows installation files"
    Write-Host "Source: $SourceDrive"
    Write-Host "Target: $TargetDrive"
    
    $excludeFiles = @()
    if ($ExcludeInstallWim) {
        $excludeFiles += "install.wim"
        Write-Host "Excluding: install.wim (will be handled separately)"
    }
    
    $sourcePath = "$SourceDrive\"
    $targetPath = "$TargetDrive\"
    
    # Use robocopy for efficient file copying
    $robocopyArgs = @(
        $sourcePath,
        $targetPath,
        '/E',           # Copy subdirectories, including empty ones
        '/NFL',         # No file list
        '/NDL',         # No directory list
        '/NP',          # No progress
        '/R:3',         # Retry 3 times
        '/W:5'          # Wait 5 seconds between retries
    )
    
    if ($excludeFiles.Count -gt 0) {
        $robocopyArgs += '/XF'
        $robocopyArgs += $excludeFiles
    }
    
    $result = & robocopy @robocopyArgs
    $exitCode = $LASTEXITCODE
    
    # Robocopy exit codes: 0-7 are success, 8+ are errors
    if ($exitCode -ge 8) {
        throw "Failed to copy files. Robocopy exit code: $exitCode"
    }
    
    Write-Success "Files copied successfully"
}

# Get install.wim information
function Get-WimImageInfo {
    param([string]$WimPath)
    
    Write-Step "Analyzing install.wim images"
    Write-Host "Image file: $WimPath"
    
    # Run DISM and capture output
    $dismOutput = & dism /Get-WimInfo /WimFile:"$WimPath" 2>&1 | Out-String
    
    # Check if DISM succeeded
    if ($LASTEXITCODE -ne 0) {
        Write-Host "DISM output:" -ForegroundColor Yellow
        Write-Host $dismOutput
        throw "DISM failed to read the image file. Exit code: $LASTEXITCODE"
    }
    
    Write-Host "DISM output (first 500 chars):" -ForegroundColor Gray
    Write-Host $dismOutput.Substring(0, [Math]::Min(500, $dismOutput.Length)) -ForegroundColor Gray
    
    $imageList = @()
    $currentImage = $null
    
    foreach ($line in $dismOutput -split "`r?`n") {
        # Match Index (handles both quoted and unquoted, English and German)
        if ($line -match '^\s*Index\s*:\s*"?(\d+)"?') {
            if ($currentImage) {
                $imageList += $currentImage
            }
            $currentImage = @{
                Index = [int]$matches[1]
                Name = ''
                Description = ''
                Size = ''
            }
        } 
        # Match Name (English and German)
        elseif ($line -match '^\s*(Name|Namen)\s*:\s*"?(.+?)"?\s*$') {
            if ($currentImage) {
                $currentImage.Name = $matches[2].Trim()
            }
        } 
        # Match Description (English: Description, German: Beschreibung)
        elseif ($line -match '^\s*(Description|Beschreibung)\s*:\s*"?(.+?)"?\s*$') {
            if ($currentImage) {
                $currentImage.Description = $matches[2].Trim()
            }
        } 
        # Match Size (English: Size, German: Größe/Groesse)
        elseif ($line -match '^\s*(Size|Größe|Groesse)\s*:\s*(.+)') {
            if ($currentImage) {
                $currentImage.Size = $matches[2].Trim()
            }
        }
    }
    
    if ($currentImage) {
        $imageList += $currentImage
    }
    
    if ($imageList.Count -eq 0) {
        Write-Host "Full DISM output:" -ForegroundColor Yellow
        Write-Host $dismOutput
    }
    
    return $imageList
}

# Get Windows version information from WIM
function Get-WindowsVersion {
    param([string]$WimPath, [int]$ImageIndex = 1)
    Write-Host "Detecting Windows version..." -ForegroundColor Gray
    $dismOutput = & dism /Get-WimInfo /WimFile:"$WimPath" /Index:$ImageIndex 2>&1 | Out-String
    if ($dismOutput -match '(Version|Build)\s*:\s*(\d+\.\d+\.\d+)') {
        $fullVersion = $matches[2]
        $buildNumber = $fullVersion.Split('.')[2]
        # Load version map from YAML
        $versionMapPath = Join-Path $PSScriptRoot 'WindowsVersionMap.yaml'
        $versionMap = @{
        }
        if (Test-Path $versionMapPath) {
            $versionMap = ConvertFrom-Yaml (Get-Content $versionMapPath -Raw)
        }
        $marketingVersion = $versionMap[$buildNumber]
        if (-not $marketingVersion) {
            $marketingVersion = "Build $buildNumber"
        }
        return @{
            FullVersion = $fullVersion
            BuildNumber = $buildNumber
            MarketingVersion = $marketingVersion
        }
    }
    return $null
}

# Export specific image from install.wim
function Export-WimImage {
    param(
        [string]$SourceWim,
        [int]$ImageIndex,
        [string]$DestinationWim,
        [string]$CompressionType = "max"
    )
    
    Write-Step "Exporting Windows image (Index: $ImageIndex)"
    Write-Host "This may take several minutes..."
    Write-Host "Using compression: $CompressionType" -ForegroundColor Gray
    
    $dismArgs = @(
        '/Export-Image',
        "/SourceImageFile:`"$SourceWim`"",
        "/SourceIndex:$ImageIndex",
        "/DestinationImageFile:`"$DestinationWim`"",
        "/Compress:$CompressionType"
    )
    
    $result = & dism @dismArgs
    
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to export image. DISM exit code: $LASTEXITCODE"
    }
    
    # Show file size
    $fileSize = (Get-Item $DestinationWim).Length
    $fileSizeGB = [math]::Round($fileSize / 1GB, 2)
    Write-Host "Exported WIM size: $fileSizeGB GB" -ForegroundColor Gray
    
    Write-Success "Image exported successfully"
}

# Copy EDS folder to USB
function Copy-EDSFolder {
    param(
        [string]$TargetDrive
    )
    
    Write-Step "Copying EDS folder to target"
    
    # Get the repository root (automation folder -> eds-win11setup)
    $edsSource = $EDSFolder
    if (-not $edsSource -or $edsSource -eq "") {
        $scriptRoot = Split-Path $PSScriptRoot -Parent
        $edsSource = Join-Path $scriptRoot "EDS"
    }
    
    Write-Host "Looking for EDS folder at: $edsSource" -ForegroundColor Gray
    
    if (-not (Test-Path $edsSource)) {
        throw "EDS folder not found at: $edsSource"
    }
    
    $edsTarget = Join-Path $TargetDrive "EDS"
    
    Write-Host "Copying to: $edsTarget" -ForegroundColor Gray
    Copy-Item -Path $edsSource -Destination $edsTarget -Recurse -Force
    
    Write-Success "EDS folder copied successfully"
}

# Find oscdimg.exe from Windows ADK
function Find-OscdImg {
    $possiblePaths = @(
        "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
        "${env:ProgramFiles}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
        "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\x86\Oscdimg\oscdimg.exe"
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            return $path
        }
    }
    
    return $null
}

# Create bootable ISO from folder
function New-BootableISO {
    param(
        [string]$SourceFolder,
        [string]$OutputISO,
        [string]$VolumeLabel = "WIN11_EDS"
    )
    
    Write-Step "Creating bootable ISO"
    
    # Find oscdimg.exe
    $oscdimg = Find-OscdImg
    
    if (-not $oscdimg) {
        Write-Warn "oscdimg.exe not found. Windows ADK is required to create ISO files."
        Write-Host "Download from: https://go.microsoft.com/fwlink/?linkid=2196127" -ForegroundColor Yellow
        Write-Host ""
        $manualChoice = Read-Host "Do you want to continue without creating ISO? (y/N)"
        if ($manualChoice -ne 'y' -and $manualChoice -ne 'Y') {
            throw "Windows ADK not installed"
        }
        return $false
    }
    
    Write-Host "Using oscdimg: $oscdimg" -ForegroundColor Gray
    Write-Host "Source folder: $SourceFolder" -ForegroundColor Gray
    Write-Host "Output ISO: $OutputISO" -ForegroundColor Gray
    
    # Get boot files
    $etfsboot = Join-Path $SourceFolder "boot\etfsboot.com"
    $efisys = Join-Path $SourceFolder "efi\microsoft\boot\efisys.bin"
    
    if (-not (Test-Path $etfsboot)) {
        throw "Boot file not found: $etfsboot"
    }
    
    if (-not (Test-Path $efisys)) {
        throw "EFI boot file not found: $efisys"
    }
    
    Write-Host "Creating ISO (this may take several minutes)..." -ForegroundColor Cyan
    # Ensure output directory exists
    $isoDir = Split-Path $OutputISO -Parent
    if (-not (Test-Path $isoDir)) {
        Write-Host "Creating ISO output directory: $isoDir" -ForegroundColor Gray
        New-Item -ItemType Directory -Path $isoDir -Force | Out-Null
    }
    # oscdimg parameters for UEFI + BIOS bootable ISO
    $oscdimgArgs = @(
        '-m',                           # Ignore maximum image size
        '-o',                           # Optimize storage
        '-u2',                          # UDF filesystem
        '-udfver102',                   # UDF version 1.02
        "-bootdata:2#p0,e,b`"$etfsboot`"#pEF,e,b`"$efisys`"",  # Dual boot (BIOS + UEFI)
        "-l`"$VolumeLabel`"",          # Volume label
        "`"$SourceFolder`"",            # Source folder
        "`"$OutputISO`""                # Output ISO
    )
    Write-Host "Running: oscdimg $($oscdimgArgs -join ' ')" -ForegroundColor Gray
    $process = Start-Process -FilePath $oscdimg -ArgumentList $oscdimgArgs -Wait -NoNewWindow -PassThru
    if ($process.ExitCode -ne 0) {
        throw "oscdimg failed with exit code: $($process.ExitCode)"
    }
    if (Test-Path $OutputISO) {
        $isoSize = (Get-Item $OutputISO).Length
        $isoSizeGB = [math]::Round($isoSize / 1GB, 2)
        Write-Success "ISO created successfully: $OutputISO ($isoSizeGB GB)"
        return $true
    } else {
        throw "ISO file was not created"
    }
}

# Interactive menu for image selection
function Select-WindowsImage {
    param([array]$Images)
    
    Write-Host "`nAvailable Windows images:" -ForegroundColor Cyan
    Write-Host ("=" * 80) -ForegroundColor Cyan
    
    for ($i = 0; $i -lt $Images.Count; $i++) {
        $img = $Images[$i]
        Write-Host "[$($i+1)] " -NoNewline -ForegroundColor Yellow
        Write-Host "Index $($img.Index): " -NoNewline -ForegroundColor White
        Write-Host "$($img.Name)" -ForegroundColor Green
        if ($img.Description) {
            Write-Host "    Description: $($img.Description)" -ForegroundColor Gray
        }
        if ($img.Size) {
            Write-Host "    Size: $($img.Size)" -ForegroundColor Gray
        }
        Write-Host ""
    }
    
    do {
        $selection = Read-Host "Select image number (1-$($Images.Count))"
        $selectedIndex = -1
        if ([int]::TryParse($selection, [ref]$selectedIndex)) {
            $selectedIndex = $selectedIndex - 1
        } else {
            Write-Host "Invalid input. Please enter a number between 1 and $($Images.Count)." -ForegroundColor Red
        }
    } while ($selectedIndex -lt 0 -or $selectedIndex -ge $Images.Count)
    
    return $Images[$selectedIndex]
}

# Main script execution
function Main {
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "                                                                " -ForegroundColor Cyan
    Write-Host "     Windows 11 Installer Creation Tool                        " -ForegroundColor Cyan
    Write-Host "     EDS Integration                                           " -ForegroundColor Cyan
    Write-Host "                                                                " -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""

    # Check for admin privileges
    if (-not (Test-Administrator)) {
        Write-Fail "This script requires administrator privileges!"
        Write-Host "Please run PowerShell as Administrator and try again."
        exit 1
    }
    Write-Success "Running with administrator privileges"
    
    # Create output directory
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    Write-Success "Working directory: $OutputPath"
    
    # Step 1: Get or download ISO
    $isoFile = $null
    if ($SkipDownload -and $ISOPath) {
        if (-not (Test-Path $ISOPath)) {
            Write-Fail "Specified ISO file not found: $ISOPath"
            exit 1
        }
        $isoFile = $ISOPath
        Write-Success "Using existing ISO: $isoFile"
    } else {
        # Interactive language selection
        Write-Host "`nSelect Windows 11 language:" -ForegroundColor Cyan
        $languages = @(
            @{Code="en-us"; Name="English (United States)"}
            @{Code="de-de"; Name="German (Germany)"}
            @{Code="fr-fr"; Name="French (France)"}
            @{Code="es-es"; Name="Spanish (Spain)"}
            @{Code="it-it"; Name="Italian (Italy)"}
            @{Code="ja-jp"; Name="Japanese (Japan)"}
            @{Code="pt-br"; Name="Portuguese (Brazil)"}
        )
        
        for ($i = 0; $i -lt $languages.Count; $i++) {
            Write-Host "[$($i+1)] $($languages[$i].Name)" -ForegroundColor Yellow
        }
        
        do {
            $langSelection = Read-Host "`nSelect language (1-$($languages.Count))"
            $langIndex = -1
            if ([int]::TryParse($langSelection, [ref]$langIndex)) {
                $langIndex = $langIndex - 1
            } else {
                Write-Host "Invalid input. Please enter a number between 1 and $($languages.Count)." -ForegroundColor Red
            }
        } while ($langIndex -lt 0 -or $langIndex -ge $languages.Count)
        
        $selectedLanguage = $languages[$langIndex].Code
        Write-Success "Selected: $($languages[$langIndex].Name)"
        
        $isoFile = Get-Windows11ISO -OutputPath $OutputPath -Language $selectedLanguage
    }
    
    # Step 2: Mount ISO
    $isoDrive = Mount-ISOImage -ISOPath $isoFile
    
    try {
        # Step 3: Check for install.wim or install.esd
        $installWimSource = Join-Path $isoDrive "sources\install.wim"
        $installEsdSource = Join-Path $isoDrive "sources\install.esd"
        
        $imageSource = $null
        $isEsd = $false
        
        if (Test-Path $installWimSource) {
            $imageSource = $installWimSource
            Write-Success "Found install.wim"
        } elseif (Test-Path $installEsdSource) {
            $imageSource = $installEsdSource
            $isEsd = $true
            Write-Success "Found install.esd (will be converted to install.wim)"
        } else {
            throw "Neither install.wim nor install.esd found in ISO at: $isoDrive\sources\"
        }
        
        $wimImages = Get-WimImageInfo -WimPath $imageSource
        if ($wimImages.Count -eq 0) {
            throw "No images found in $([System.IO.Path]::GetFileName($imageSource))"
        }
        
        Write-Host "`nFound $($wimImages.Count) Windows image(s) in install.wim"
        
        # Step 4: Detect Windows version
        $versionInfo = Get-WindowsVersion -WimPath $imageSource -ImageIndex 1
        if ($versionInfo) {
            Write-Success "Detected Windows version: $($versionInfo.MarketingVersion) (Build $($versionInfo.BuildNumber))"
        }
        
        # Step 5: Select image
        $selectedImage = Select-WindowsImage -Images $wimImages
        Write-Success "Selected: $($selectedImage.Name) (Index: $($selectedImage.Index))"
        
        # Ask about cumulative updates
        Write-Host "`n" + ("=" * 80) -ForegroundColor Cyan
        Write-Warn "Cumulative Update Installation"
        Write-Host "The script can help you apply cumulative updates to the image."
        Write-Host "You can download the latest updates from:"
        
        # Construct version-specific URL
        if ($versionInfo -and $versionInfo.MarketingVersion -notlike "Build*") {
            $updateUrl = "https://catalog.update.microsoft.com/Search.aspx?q=Kumulatives%20Update%20f%C3%BCr%20Windows%2011%20Version%20$($versionInfo.MarketingVersion)"
            Write-Host $updateUrl -ForegroundColor Yellow
            Write-Host "(Cumulative Update for Windows 11 Version $($versionInfo.MarketingVersion) - Build $($versionInfo.BuildNumber))" -ForegroundColor Gray
        } else {
            $updateUrl = "https://catalog.update.microsoft.com/Search.aspx?q=Kumulatives%20Update%20f%C3%BCr%20Windows%2011"
            Write-Host $updateUrl -ForegroundColor Yellow
        }
        
        Write-Host ""
        $openBrowser = Read-Host "Open this URL in browser? (Y/n)"
        if ($openBrowser -ne 'n' -and $openBrowser -ne 'N') {
            Start-Process $updateUrl
            Write-Success "Browser opened"
        }
        
        Write-Host ""
        $applyCU = Read-Host "Do you want to apply a cumulative update? (y/N)"
        
        $cuPath = $null
        if ($applyCU -eq 'y' -or $applyCU -eq 'Y') {
            Add-Type -AssemblyName System.Windows.Forms
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.Filter = "Windows Update files (*.msu;*.cab)|*.msu;*.cab|All files (*.*)|*.*"
            $openFileDialog.Title = "Select Cumulative Update file"
            
            if ($openFileDialog.ShowDialog() -eq 'OK') {
                $cuPath = $openFileDialog.FileName
                Write-Success "Selected update: $cuPath"
            }
        }
        
        # Step 5: USB Drive preparation (if not skipped)
        $usbDriveLetter = $null
        if (-not $SkipUSBCreation) {
            Write-Host "`n" + ("=" * 80) -ForegroundColor Cyan
            $usbDrives = @(Get-USBDrives)
            
            if ($usbDrives.Count -eq 0 -or $null -eq $usbDrives) {
                Write-Warn "No USB drives detected."
                $skipUSB = Read-Host "Continue without creating USB drive? (y/N)"
                if ($skipUSB -ne 'y' -and $skipUSB -ne 'Y') {
                    throw "No USB drive available"
                }
                $SkipUSBCreation = $true
            } else {
                Write-Host "Detected $($usbDrives.Count) USB drive(s):" -ForegroundColor Cyan
                for ($i = 0; $i -lt $usbDrives.Count; $i++) {
                    $drive = $usbDrives[$i]
                    Write-Host "[$($i+1)] " -NoNewline -ForegroundColor Yellow
                    if ($drive.DriveLetter) {
                        Write-Host "$($drive.DriveLetter): - $($drive.FriendlyName) - $($drive.Size) GB" -ForegroundColor Green
                    } else {
                        Write-Host "Disk $($drive.DiskNumber) - $($drive.FriendlyName) - $($drive.Size) GB (No drive letter)" -ForegroundColor Green
                    }
                    if ($drive.VolumeName) {
                        Write-Host "    Label: $($drive.VolumeName)" -ForegroundColor Gray
                    }
                }
                
                $usbIndex = -1
                do {
                    $usbSelection = Read-Host "`nSelect USB drive (1-$($usbDrives.Count))"
                    # Try to parse as integer, handle invalid input
                    $parsed = 0
                    if ([int]::TryParse($usbSelection, [ref]$parsed) -and $parsed -ge 1 -and $parsed -le $usbDrives.Count) {
                        $usbIndex = $parsed - 1
                    } else {
                        Write-Host "Invalid input. Please enter a number between 1 and $($usbDrives.Count)." -ForegroundColor Red
                        $usbIndex = -1
                    }
                } while ($usbIndex -lt 0)
                
                $selectedUSB = $usbDrives[$usbIndex]
                
                # Check minimum size (8GB)
                if ($selectedUSB.Size -lt 8) {
                    Write-Warn "USB drive is less than 8 GB. This may not be sufficient."
                    $continue = Read-Host "Continue anyway? (y/N)"
                    if ($continue -ne 'y' -and $continue -ne 'Y') {
                        throw "USB drive too small"
                    }
                }
                
                $usbDriveLetter = Initialize-USBDrive -DriveLetter $selectedUSB.DriveLetter -DiskNumber $selectedUSB.DiskNumber
            }
        }
        
        # Step 6: Copy files or prepare in output directory
        if ($SkipUSBCreation) {
            Write-Step "Preparing installer files in output directory"
            $targetPath = Join-Path $OutputPath "Installer"
            if (-not (Test-Path $targetPath)) {
                New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
            }
            Copy-WindowsFiles -SourceDrive $isoDrive -TargetDrive $targetPath -ExcludeInstallWim $true
        } else {
            Copy-WindowsFiles -SourceDrive $isoDrive -TargetDrive $usbDriveLetter -ExcludeInstallWim $true
            $targetPath = $usbDriveLetter
        }
        
        # Step 7: Export and copy install.wim
        Write-Step "Processing install image"
        $exportedWim = Join-Path $OutputPath "install_exported.wim"
        Export-WimImage -SourceWim $imageSource -ImageIndex $selectedImage.Index -DestinationWim $exportedWim
        
        # Apply cumulative update if selected
        if ($cuPath) {
            Write-Step "Applying Cumulative Update"
            Write-Host "Mounting image..."
            
            $mountDir = Join-Path $OutputPath "mount"
            if (-not (Test-Path $mountDir)) {
                New-Item -ItemType Directory -Path $mountDir -Force | Out-Null
            }
            
            & dism /Mount-Wim /WimFile:"$exportedWim" /Index:1 /MountDir:"$mountDir"
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "Applying update (this may take several minutes)..."
                & dism /Image:"$mountDir" /Add-Package /PackagePath:"$cuPath"
                
                if ($LASTEXITCODE -eq 0) {
                    Write-Success "Update applied successfully"
                } else {
                    Write-Warn "Update application failed or completed with warnings"
                }
                
                Write-Host "Committing changes and unmounting..."
                & dism /Unmount-Wim /MountDir:"$mountDir" /Commit
                
                if ($LASTEXITCODE -eq 0) {
                    Write-Success "Image unmounted successfully"
                } else {
                    Write-Warn "Unmount completed with warnings"
                }
                
                # Cleanup mount directory
                Remove-Item $mountDir -Force -ErrorAction SilentlyContinue
            } else {
                Write-Fail "Failed to mount image"
            }
        }
        
        # Copy the processed install.wim to target
        $targetSourcesDir = Join-Path $targetPath "sources"
        if (-not (Test-Path $targetSourcesDir)) {
            New-Item -ItemType Directory -Path $targetSourcesDir -Force | Out-Null
        }
        
        # Check file size and available space
        $wimSize = (Get-Item $exportedWim).Length
        $wimSizeGB = [math]::Round($wimSize / 1GB, 2)
        Write-Host "Install.wim size: $wimSizeGB GB" -ForegroundColor Gray
        
        if ($usbDriveLetter) {
            $usbDrive = $usbDriveLetter.TrimEnd(':')
            $availableSpace = (Get-Volume -DriveLetter $usbDrive).SizeRemaining
            $availableSpaceGB = [math]::Round($availableSpace / 1GB, 2)
            Write-Host "USB drive free space: $availableSpaceGB GB" -ForegroundColor Gray
            
            # Check if we need to split the WIM file (FAT32 has 4GB file size limit)
            if ($wimSize -gt 4GB) {
                Write-Warn "WIM file is larger than 4GB, splitting is required for FAT32"
                Write-Host "Splitting install.wim into multiple files..."
                
                $targetWim = Join-Path $targetSourcesDir "install.swm"
                & dism /Split-Image /ImageFile:"$exportedWim" /SWMFile:"$targetWim" /FileSize:3800
                
                if ($LASTEXITCODE -eq 0) {
                    Write-Success "WIM file split successfully"
                } else {
                    throw "Failed to split WIM file"
                }
            } else {
                Write-Host "Copying install.wim to target..."
                Copy-Item -Path $exportedWim -Destination (Join-Path $targetSourcesDir "install.wim") -Force
                Write-Success "install.wim copied"
            }
        } else {
            Write-Host "Copying install.wim to target..."
            Copy-Item -Path $exportedWim -Destination (Join-Path $targetSourcesDir "install.wim") -Force
            Write-Success "install.wim copied"
        }
        
        # Step 8: Copy EDS folder
        Copy-EDSFolder -TargetDrive $targetPath
        
        # Step 9: Create ISO file (if requested)
        $isoCreated = $false
        if ($CreateISO) {
            Write-Host "`n" + ("=" * 80) -ForegroundColor Cyan
            
            if (-not $ISOOutputPath) {
                $defaultISOName = "Windows11_$($selectedImage.Name.Replace(' ', '_'))_EDS.iso"
                $ISOOutputPath = Join-Path (Split-Path $OutputPath -Parent) $defaultISOName
            }
            
            # Ensure .iso extension
            if (-not $ISOOutputPath.EndsWith('.iso')) {
                $ISOOutputPath += '.iso'
            }
            
            # Check if file already exists
            if (Test-Path $ISOOutputPath) {
                Write-Warn "ISO file already exists: $ISOOutputPath"
                $overwrite = Read-Host "Overwrite existing file? (y/N)"
                if ($overwrite -eq 'y' -or $overwrite -eq 'Y') {
                    Remove-Item $ISOOutputPath -Force
                } else {
                    Write-Host "ISO creation skipped" -ForegroundColor Yellow
                    $CreateISO = $false
                }
            }
            
            if ($CreateISO) {
                $isoCreated = New-BootableISO -SourceFolder $targetPath -OutputISO $ISOOutputPath -VolumeLabel "WIN11_EDS"
            }
        }
        
        # Success summary
        Write-Host "`n" + ("=" * 80) -ForegroundColor Green
        Write-Host "[OK] Windows 11 Installer created successfully!" -ForegroundColor Green
        Write-Host ("=" * 80) -ForegroundColor Green
        Write-Host "`nSummary:" -ForegroundColor Cyan
        Write-Host "  - Windows Edition: $($selectedImage.Name)"
        Write-Host "  - Image Index: $($selectedImage.Index)"
        if ($cuPath) {
            Write-Host "  - Cumulative Update: Applied"
        }
        if ($isoCreated) {
            Write-Host "  - ISO File: $ISOOutputPath"
        }
        if ($SkipUSBCreation) {
            Write-Host "  - Location: $targetPath"
            if (-not $isoCreated) {
                Write-Host "`nTo create a bootable USB, copy the contents of $targetPath to a FAT32-formatted USB drive."
            }
        } else {
            Write-Host "  - USB Drive: $usbDriveLetter"
            Write-Host "`nThe USB drive is now bootable and ready to use!"
        }
        Write-Host ""
        
    } finally {
        # Cleanup: Unmount ISO
        Write-Step "Cleaning up"
        Dismount-DiskImage -ImagePath $isoFile -ErrorAction SilentlyContinue | Out-Null
        Write-Success "ISO unmounted"
        
        # Ask about cleanup
        $cleanup = Read-Host "`nDelete temporary files? (y/N)"
        if ($cleanup -eq 'y' -or $cleanup -eq 'Y') {
            $tempWim = Join-Path $OutputPath "install_exported.wim"
            if (Test-Path $tempWim) {
                Remove-Item $tempWim -Force
                Write-Success "Temporary files cleaned up"
            }
        }
    }
}

# Run main function
try {
    Main
} catch {
    Write-Fail "An error occurred: $($_.Exception.Message)"
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}
