# Enterprise Deployment Suite (EDS) - Windows 11 Setup

**Related documentation:**

- [Automation/README.md: Windows 11 Installer Automation Guide](Create-Windows11Installer.md)
- [WebService.md: ISO Creation Web Service API](WebService.md)

## Ways to Create Windows Images

- **Interactive PowerShell script:** Use `automation/Create-Windows11Installer.ps1` for step-by-step creation of USB or ISO images with EDS integration. See [automation/README.md](automation/README.md) for details and usage examples.
- **Web Service API:** Use the HTTP API described in [WebService.md](WebService.md) to trigger image creation remotely, with support for callback notifications and custom parameters.

# Overview

This repository is a companion repository for [Enterprise-Deployment-Suite](https://github.com/markush97/Enterprise-Deployment-Suite). This repo contains the necessary files and guides to allow for creation of a highly customizable Windows11 Installer - either via USB-Stick or PXE.

The aim of this repository to create easy to maintain and adopt Windows-Installer that can be modified and customized beyond the abilities of the (Auto)unattendend.xml functionality that windows brings with their installer and winPe. All of this is achieved (except for the inital setup) without complicated tools like dism.exe or OSCDimg.exe.

## Usage with (Offline)USB-Stick

To be able to use these scripts with a default Windows USB-Installer, you have to follow these instructions:

1. Download a Windows 10 or Windows 11 .iso file from Microsoft, such as Win11_24H2_English_x64.iso.
2. Mount that .iso image in File Explorer. This tutorial assumes that the .iso image is mounted to the D:\ drive.
3. Plug in a USB flash drive with a capacity of at least 8 GiB. Using diskpart.exe or diskmgmt.msc, wipe that drive, create a single partition and format it with the FAT32 file system. This tutorial assumes that the flash drive is assigned the E:\ drive letter.
4. Launch a PowerShell session with administrative privileges:
   - Copy the contents of the .iso image to the flash drive, but omit the large install.wim file:
     > Copy-Item -Path 'D:\*' -Destination 'E:\' -Recurse -Force -Exclude 'install.wim';
   - Copy the large install.wim file to the flash drive in chunks of 2 GiB:
     > Dism.exe /Split-Image /ImageFile:"D:\sources\install.wim" /SWMFile:"E:\sources\install.swm" /FileSize:2048;
5. Modify your boot.wim (and optionally install.wim) like instructed later in this document.
6. Unmount the .iso image.
7. Copy the "EDS" folder of this repository to the root of the USB-Stick
8. Boot the USB-Stick

# Install.wim

The install.wim is the image of the to-be-created Windows installation. This will basically be copied into the new Windows-Root (C:). So if you want any tools, drivers or files to exist in the newly created installation, place them there.

## Downloading the Latest CU/SU (Cumulative Update/Security Update)

Find the latest KB (Knowledge Base) update using this link: https://catalog.update.microsoft.com/Search.aspx?q=Kumulatives%20Update%20f%C3%BCr%20Windows%2011

At the time of writing, this is "2025-05 Cumulative Update for Windows 11 Version 24H2 for x64-based systems (KB5058411)".

The file should be approximately 3-5GB in size.

### Installing in the Installation Media

Mount the install.wim and apply the update using dism.exe. Alternatively, a helper tool like DISMTools can be used.

```cmd
dism /Image:C:<install.wim-mountdir> /Add-Package /PackagePath:<path-to-CU.msu>
```

Example:

```cmd
dism /Image:"C:\Temp\Imaging\FullAutomation\mount" /Add-Package /PackagePath:"C:\Users\mhinkel\Downloads\windows11.0-kb5058411-x64_fc93a482441b42bcdbb035f915d4be2047d63de5.msu"
```

**IMPORTANT**: This process can take several minutes.

## Remove unneeded Images for fully automated installation

By default, most Microsoft-Windows installation mediums contain multiple windows version (Pro, Home, Enterprise..). To allow fully-automated windows-installations (and slim the installmedium down) it is required to
remove every Image-Index except the one that is needed, from the installation medium.

This can be achieved with a helper-tool like DISMTools or dism.exe by only exporting the needed image:

```cmd
Dism /Export-Image /SourceImageFile:"D:\sources\install.wim" /SourceIndex:1 /DestinationImageFile:"D:\sources\pro.wim"
```

# Boot.wim

The boot.wim is the "OS" the installer uses. This is being booted when booting from a windows-usbstick or winpe environment (e.g. PXE).

## Overview

While it's not strictly necessary to build a custom boot.wim a custom boot.wim becomes necessary in the following scenarios:

- Special drivers (e.g., for WiFi setup or within certain VMs)
- Requirements beyond unattended.xml capabilities (e.g., loading autounattended.xml from a web server)
- Custom Windows Installer GUI

Since the boot.wim can be reused for different Windows11-Installer (and maybe even for Windows12), it makes sense to modify one boot.wim to your likings once and keep it.

This guide will aim to provide minimal instructions for such modifications.

## Important Notes

In most guides referenced below, the boot.wim file is already mounted (e.g., using the dism.exe tool). The term "commit" is often used in this context - only through committing are changes written to the .wim file.

**Critical**: The standard Windows 11 ISO contains two images in the boot.wim:

1. First Image (WinPE): Really slim WinPE Environment with smaller footprint but more flexible
2. Second Image: Bigger image containing the classic windows-11 Setup-GUI and other stuff, which we will use in our case, since we only want to inject some configuration-files and let the normal windows11-installer handle the rest.

## Loading Drivers or Modules

While drivers and packages can be loaded manually using dism.exe commands ("Dism /Add-Package"), this process is tedious and error-prone. We recommend using [DISMTools](https://github.com/CodingWonders/DISMTools).

Process:

1. Mount the boot.wim image
2. Add packages using the tool's interface
3. Commit changes frequently to prevent corruption

## PowerShell Requirements

For PowerShell functionality in the installer, the following packages must be loaded:

```
C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs\WinPE-WMI.cab
C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs\WinPe-NetFx.cab
C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs\WinPe-Scripting.cab
C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs\WinPe-Powershell.cab
```

## Windows Installer Modification

While not needed, we group all EDS-custom information in a "EDS" folder at the boot.wim root for better traceability and reproducibility.

### Entry Points

#### Standard WinPE (startnet.cmd)

> Better use the method for (winpeshl.ini) below for minimal modifications

The entry point in a boot.wim for standard WinPE is `<boot.wim-root>\Windows\System32\startnet.cmd`. This script runs first at startup. The boot.wim file is always mounted as drive X: in the installer.

Default content with CWI modifications:

```cmd
echo Initializing custom CWI-Installer
X:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -NoExit -File X:\Entrypoint.ps1
```

#### Windows 11 ISO (winpeshl.ini)

For standard Windows 11 ISOs, the entry point is winpeshl.exe, which directly starts setup. To run custom scripts, create `Windows/System32/winpeshl.ini`:

```ini
[LaunchApps]
X:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -NoExit -File X:\Entrypoint.ps1
```

For detailed winpeshl.ini syntax, refer to the [Microsoft Documentation](https://learn.microsoft.com/windows-hardware/manufacture/desktop/winpeshlini-reference-launching-an-app-when-winpe-starts?view=windows-11).

The start script is intentionally minimal to maintain maximum flexibility.

#### Entrypoint.ps1

Boths modifications above do nothing but executing Entrypoint.ps1. This file has to reside inside the root-folder of the boot.wim. Again it makes modifications easier, to keep this file as simple as possible.

For this reason the [Entrypoint.ps1](/bootwim-modifications/Entrypoint.ps1) in this repository only has the task to search every attached drive for a folder called "EDS" (or whatever you define it to be named like) that has an File named "eds.cfg" inside it. This file can be completly empty aswell. Its main reason is to tell the Installer which Medium to use in case there are multiple USB-Sticks attached.

At last the entrypoint.ps1 starts a "Start.ps1" located at the root of the found "EDS"-Folder. Every complicated logic should be handled inside this script.
