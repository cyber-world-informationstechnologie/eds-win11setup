# Automation Scripts

This folder contains automation scripts to simplify the creation and maintenance of Windows 11 installation media with EDS integration.

## Create-Windows11Installer.ps1

Automates the entire process of creating a customized Windows 11 installer as described in the main project README.

### Features

- ✅ Interactive selection of Windows edition and language
- ✅ Automatic image extraction and optimization
- ✅ Optional cumulative update integration
- ✅ USB drive preparation with proper formatting
- ✅ Bootable ISO file creation
- ✅ Automatic EDS folder integration
- ✅ Support for both USB and file-based output
- ✅ Automatic WIM file splitting for FAT32 compatibility

### Prerequisites

- Windows 10/11 with PowerShell 5.1 or later
- Administrator privileges
- At least 15 GB free disk space
- USB drive with at least 8 GB capacity (if creating bootable USB)
- Internet connection (for ISO download)
- Windows ADK (Assessment and Deployment Kit) - Required only for ISO creation
  - Download: https://go.microsoft.com/fwlink/?linkid=2196127

### Usage

#### Basic Usage (Interactive)

```powershell
# Run with all interactive prompts
.\Create-Windows11Installer.ps1
```

This will:

1. Prompt you to download or select a Windows 11 ISO
2. Let you choose the Windows edition (Pro, Home, Enterprise, etc.)
3. Optionally apply cumulative updates
4. Select and prepare a USB drive
5. Copy all necessary files including the EDS folder

#### Advanced Usage

**Use an existing ISO:**

```powershell
.\Create-Windows11Installer.ps1 -ISOPath "C:\ISOs\Win11_24H2_English_x64.iso" -SkipDownload
```

**Skip USB creation (prepare files only):**

```powershell
.\Create-Windows11Installer.ps1 -SkipUSBCreation -OutputPath "D:\Win11Installer"
```

**Create a bootable ISO file:**

```powershell
.\Create-Windows11Installer.ps1 -ISOPath "C:\ISOs\Win11.iso" -SkipDownload -SkipUSBCreation -CreateISO -ISOOutputPath "C:\Output\Win11_Custom.iso"
```

**Specify USB drive directly:**

```powershell
.\Create-Windows11Installer.ps1 -USBDrive "E"
```

**Full automation example:**

```powershell
.\Create-Windows11Installer.ps1 `
    -ISOPath "C:\ISOs\Win11.iso" `
    -SkipDownload `
    -USBDrive "E" `
    -OutputPath "D:\Temp"
```

### Parameters

| Parameter         | Type   | Description                                                               |
| ----------------- | ------ | ------------------------------------------------------------------------- |
| `OutputPath`      | String | Path for temporary files (default: `$env:TEMP\WindowsInstallerBuild`)     |
| `USBDrive`        | String | USB drive letter (e.g., "E"). If not specified, will prompt interactively |
| `SkipDownload`    | Switch | Skip ISO download and use existing file                                   |
| `ISOPath`         | String | Path to existing Windows 11 ISO (required with `-SkipDownload`)           |
| `SkipUSBCreation` | Switch | Only prepare files without creating USB drive                             |
| `CreateISO`       | Switch | Create a bootable ISO file from the prepared installer files              |
| `ISOOutputPath`   | String | Path where the output ISO file should be created                          |

### Interactive Prompts

The script will guide you through:

1. **Language Selection**: Choose from 7 major languages (English, German, French, Spanish, Italian, Japanese, Portuguese)
2. **Windows Edition**: Select from available editions in the ISO (Pro, Home, Enterprise, etc.)
3. **Cumulative Updates**: Option to apply Windows Updates (.msu files) to the image
4. **USB Drive Selection**: Pick from detected USB drives with size and label information
5. **Confirmation**: Safety confirmation before erasing USB drive data

### What Gets Created

#### On USB Drive (or Output Folder):

```
[USB Drive or Output Folder]
├── bootmgr
├── bootmgr.efi
├── boot/
│   └── [Windows boot files]
├── efi/
│   └── [EFI boot files]
├── sources/
│   ├── boot.wim
│   └── install.wim  [Only selected Windows edition]
├── support/
├── autorun.inf
├── setup.exe
└── EDS/             [Copied from project root]
    ├── eds.cfg
    ├── Start.ps1
    ├── Installer/
    ├── Setup/
    └── Windows/
```

### Process Overview

The script performs these steps automatically:

1. **Validation**: Check admin privileges and create working directory
2. **ISO Acquisition**: Download or use existing Windows 11 ISO
3. **ISO Mount**: Mount the ISO to access its contents
4. **Image Analysis**: Parse install.wim to show available Windows editions
5. **Edition Selection**: User selects desired Windows edition
6. **Update Integration** (Optional): Mount image and apply cumulative updates
7. **Image Export**: Extract only the selected edition (reduces size)
8. **USB Preparation**: Format USB drive as FAT32 with proper structure
9. **File Copy**: Copy Windows files (excluding original install.wim)
10. **Image Deployment**: Copy the optimized install.wim
11. **EDS Integration**: Copy EDS folder with all scripts and configurations
12. **Cleanup**: Unmount ISO and optionally remove temporary files

### Cumulative Updates

To apply Windows Updates:

1. Download the latest cumulative update from [Microsoft Update Catalog](https://catalog.update.microsoft.com/)
2. Search for "Cumulative Update for Windows 11 Version [your version]"
3. Download the `.msu` file (usually 3-5 GB)
4. When prompted by the script, select the downloaded `.msu` file
5. The script will mount the image, apply the update, and commit changes

**Note**: Update application can take 10-30 minutes depending on system performance.

### Troubleshooting

**"This script requires administrator privileges"**

- Right-click PowerShell and select "Run as Administrator"

**"No USB drives detected"**

- Ensure USB drive is properly connected
- Try a different USB port
- Check if the drive appears in Disk Management

**"Failed to export image"**

- Ensure you have enough free disk space (at least 15 GB)
- Check that DISM is available (`dism /?` in command prompt)
- Try running the script again

**"Robocopy failed to copy files"**

- Check USB drive is not write-protected
- Ensure USB drive is properly formatted
- Verify the ISO is not corrupted

**Update application fails**

- Ensure the update matches your Windows version (24H2, 23H2, etc.)
- Check that the `.msu` file is not corrupted
- Verify you have enough disk space

### Safety Features

- ⚠️ USB drive formatting requires explicit "YES" confirmation
- ⚠️ Shows drive information before erasing
- ⚠️ Validates minimum USB size (8 GB)
- ⚠️ Uses safe file operations with error handling
- ⚠️ Automatic cleanup of mounted images on failure

### Performance Tips

- Use an SSD for the `OutputPath` to speed up image operations
- Use USB 3.0 or faster for quicker file transfers
- Close other applications during update application
- Expected total time: 20-45 minutes (depending on hardware and updates)

### Example Session

```
==> Downloading Windows 11 ISO

Select Windows 11 language:
[1] English (United States)
[2] German (Germany)
[3] French (France)
...

Select language (1-7): 1
✓ Selected: English (United States)

==> Mounting ISO image
✓ ISO mounted to D:\

==> Analyzing install.wim images
✓ Found 6 Windows image(s)

Available Windows images:
[1] Index 1: Windows 11 Home
[2] Index 2: Windows 11 Home N
[3] Index 3: Windows 11 Home Single Language
[4] Index 4: Windows 11 Education
[5] Index 5: Windows 11 Pro
[6] Index 6: Windows 11 Pro N

Select image number (1-6): 5
✓ Selected: Windows 11 Pro (Index: 5)

Do you want to apply a cumulative update? (y/N): y
✓ Selected update: KB5058411.msu

Detected USB drives:
[1] E: - SanDisk Ultra 3.0 - 32 GB

Select USB drive (1-1): 1
✓ USB drive prepared: E:\

==> Copying Windows installation files
✓ Files copied successfully

==> Processing install.wim
✓ Image exported successfully

==> Applying Cumulative Update
✓ Update applied successfully

==> Copying EDS folder to USB drive
✓ EDS folder copied successfully

✓ Windows 11 Installer created successfully!

Summary:
  • Windows Edition: Windows 11 Pro
  • Image Index: 5
  • Cumulative Update: Applied
  • USB Drive: E:\

The USB drive is now bootable and ready to use!
```

### Known Limitations

- ISO download still requires manual download from Microsoft website (no direct API available)
- Cumulative update integration requires manual download
- FAT32 filesystem limits individual file size to 4 GB
- Some advanced DISM operations may require Windows ADK

### Support

For issues or questions:

1. Check the main project [README.md](../README.md)
2. Open an issue on the GitHub repository
3. Review the [Enterprise-Deployment-Suite](https://github.com/markush97/Enterprise-Deployment-Suite) documentation
