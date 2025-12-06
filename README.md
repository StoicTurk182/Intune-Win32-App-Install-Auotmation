# Intune Win32 Package Bulk Creator - Simplified

**Simple, reliable bulk packaging without wrapper scripts**

## Quick Start

```powershell
# 1. Place your EXE/MSI installers in the Source folder

# 2. Run the script
.\New-IntuneWin32Package-Simple.ps1 -SourceFolder ".\Source" -OutputFolder ".\Output" -GenerateDetectionScripts

# 3. Find your packages in Output\Packages\
```

## What It Does

✅ Packages multiple apps in one command  
✅ Preserves original app names (.intunewin matches installer name)  
✅ Generates detection scripts for Intune  
✅ Creates CSV report with install/uninstall commands  
✅ Direct installer execution (no wrapper scripts = more reliable)  

## Usage

### Basic Packaging
```powershell
.\New-IntuneWin32Package-Simple.ps1 -SourceFolder "C:\Apps" -OutputFolder "C:\Packages"
```

### With Detection Scripts (Recommended)
```powershell
.\New-IntuneWin32Package-Simple.ps1 `
    -SourceFolder "C:\Apps" `
    -OutputFolder "C:\Packages" `
    -GenerateDetectionScripts
```

### With Custom Config File
```powershell
.\New-IntuneWin32Package-Simple.ps1 `
    -SourceFolder "C:\Apps" `
    -OutputFolder "C:\Packages" `
    -ConfigFile ".\AppConfig.json" `
    -GenerateDetectionScripts
```

## Output

After running, you'll find:

```
Output/
├── Packages/                    ← .intunewin files (upload to Intune)
├── DetectionScripts/            ← PowerShell detection scripts
└── Logs/
    └── PackagingReport_*.csv    ← Install/uninstall commands
```

## Intune Configuration

The script provides the exact commands to use in Intune:

**EXE Installers:**
```
Install command:   installer.exe /S
Uninstall command: installer.exe /S
```

**MSI Installers:**
```
Install command:   msiexec /i "installer.msi" /qn
Uninstall command: msiexec /x "installer.msi" /qn
```

## Custom Install Commands (AppConfig.json)

For apps that need specific install switches, create `AppConfig.json`:

```json
{
  "applications": [
    {
      "name": "7-Zip",
      "installCommand": "7z2301-x64.exe /S",
      "uninstallCommand": "C:\\Program Files\\7-Zip\\Uninstall.exe /S"
    },
    {
      "name": "GoogleChrome",
      "installCommand": "GoogleChromeStandaloneEnterprise64.msi /qn",
      "uninstallCommand": "msiexec /x {PRODUCT-GUID} /qn"
    },
    {
      "name": "AdobeReader",
      "installCommand": "AcroRdrDC.exe /sAll /msi EULA_ACCEPT=YES",
      "uninstallCommand": "msiexec /x {AC76BA86-7AD7-1033-7B44-AC0F074E4100} /qn"
    }
  ]
}
```

## Common Silent Install Switches

| Application | Silent Switch |
|-------------|---------------|
| 7-Zip | `/S` |
| Google Chrome | `/silent /install` |
| Adobe Reader | `/sAll /msi EULA_ACCEPT=YES` |
| Notepad++ | `/S` |
| VLC | `/S` |
| Firefox | `-ms` |
| WinRAR | `/S` |
| Zoom | `/silent` |
| VSCode | `/VERYSILENT /MERGETASKS=!runcode` |

## Detection Scripts

Auto-generated detection scripts check for app installation in:
- `C:\Program Files\[AppName]`
- `C:\Program Files (x86)\[AppName]`

You can customize these after generation if needed.

## Folder Structure

```
IntunePackaging/
├── New-IntuneWin32Package-Simple.ps1  ← Main script
├── AppConfig.json                     ← Optional: Custom install commands
├── Tools/
│   └── IntuneWinAppUtil.exe          ← Microsoft's packaging tool
├── Source/                            ← PUT YOUR INSTALLERS HERE
│   ├── 7z.exe
│   ├── Chrome.msi
│   └── Notepad++.exe
└── Output/                            ← Generated packages
    ├── Packages/
    ├── DetectionScripts/
    └── Logs/
```

## Parameters

| Parameter | Required | Description | Default |
|-----------|----------|-------------|---------|
| `-SourceFolder` | Yes | Folder with installers | - |
| `-OutputFolder` | Yes | Where to save packages | - |
| `-IntuneWinAppUtilPath` | No | Path to IntuneWinAppUtil.exe | `.\Tools\IntuneWinAppUtil.exe` |
| `-GenerateDetectionScripts` | No | Create detection scripts | False |
| `-ConfigFile` | No | JSON config for custom commands | - |

## Example Workflow

### Step 1: Prepare Installers
```
Source/
├── 7z2301-x64.exe
├── GoogleChromeStandaloneEnterprise64.msi
└── npp.8.5.8.Installer.x64.exe
```

### Step 2: Run Script
```powershell
.\New-IntuneWin32Package-Simple.ps1 `
    -SourceFolder ".\Source" `
    -OutputFolder ".\Output" `
    -GenerateDetectionScripts
```

### Step 3: Review Output
```
[2024-11-29 14:30:15] [Success] Processing Complete
[2024-11-29 14:30:15] [Info] Total Installers: 3
[2024-11-29 14:30:15] [Success] Successfully Packaged: 3
[2024-11-29 14:30:15] [Info] Failed: 0

Successfully Packaged Applications:

AppName      FileName                              InstallCommand                              UninstallCommand
-------      --------                              --------------                              ----------------
7z2301-x64   7z2301-x64.exe                       7z2301-x64.exe /S                          7z2301-x64.exe /S
GoogleCh...  GoogleChromeStandaloneEnterprise64... msiexec /i "GoogleChromeStandaloneEnt...   msiexec /x "GoogleChromeStandaloneEnt...
npp.8.5.8... npp.8.5.8.Installer.x64.exe          npp.8.5.8.Installer.x64.exe /S             npp.8.5.8.Installer.x64.exe /S
```

### Step 4: Upload to Intune
1. Go to Microsoft Endpoint Manager > Apps > Windows > Add
2. Select: Windows app (Win32)
3. Upload .intunewin from `Output\Packages\`
4. Configure:
   - **Install command:** (from CSV report)
   - **Uninstall command:** (from CSV report)
   - **Detection method:** Upload script from `Output\DetectionScripts\`

## Troubleshooting

### No Installers Found
**Issue:** `No installer files found in: .\Source`

**Solution:** 
- Verify source folder contains .exe or .msi files
- Check folder path is correct

### Packaging Failed
**Issue:** `Packaging failed with exit code X`

**Solution:**
- Run PowerShell as Administrator
- Check IntuneWinAppUtil.exe exists in Tools folder
- Avoid special characters in file paths
- Use short, simple paths

### IntuneWinAppUtil Not Found
**Issue:** `IntuneWinAppUtil.exe not found`

**Solution:**
- Download from: https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool
- Place in `.\Tools\` folder
- Or use `-IntuneWinAppUtilPath` parameter

## Best Practices

✅ **Test install commands locally** before packaging  
✅ **Use AppConfig.json** for apps with non-standard switches  
✅ **Always generate detection scripts** with `-GenerateDetectionScripts`  
✅ **Review CSV report** to verify install/uninstall commands  
✅ **Test on pilot devices** before production deployment  
✅ **Keep installers backed up** separately from Source folder  

## Why This Approach?

**Simple = Reliable**
- Direct installer execution (what Intune was designed for)
- No additional layers of complexity
- Easier to troubleshoot
- Standard Microsoft approach

**When Things Fail:**
- Check Intune deployment status
- Test install command manually on a device
- Verify detection script logic
- Review installer's own logs

## Version History

**v3.0 - Simplified**
- Removed wrapper scripts for reliability
- Direct installer execution
- AppConfig.json for custom commands
- Cleaner, simpler code

**v2.0**
- Added wrapper scripts (removed in v3.0)

**v1.0**
- Initial release

---

**Author:** Andrew J Jones
**Version:** 3.0 (Simplified)  
**Last Updated:** 2024-11-29  
**PowerShell Version:** 5.1+
