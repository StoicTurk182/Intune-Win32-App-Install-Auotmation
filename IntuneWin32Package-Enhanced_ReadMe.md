# Enhanced Intune Win32 Bulk Packager v4.0

Intelligent bulk packaging for Intune Win32 applications with automatic detection rule discovery.

## Features

✅ **Automatic MSI Product Code Extraction** - No manual extraction needed  
✅ **Intelligent Silent Switch Detection** - Tests common switches (Inno Setup, NSIS, InstallShield)  
✅ **Registry Detection Auto-Discovery** - Finds registry keys automatically  
✅ **32-bit vs 64-bit Detection** - Automatically identifies app architecture  
✅ **Comprehensive CSV Report** - Ready-to-use install/uninstall commands and detection rules  
✅ **Bulk Processing** - Package multiple apps at once  

---

## Quick Start
```powershell
# Basic usage
.\New-IntuneWin32Package-Enhanced.ps1 -SourceFolder "C:\Apps" -OutputFolder "C:\Packages"

# With install testing (recommended)
.\New-IntuneWin32Package-Enhanced.ps1 -SourceFolder "C:\Apps" -OutputFolder "C:\Packages" -TestInstalls
```

---

## Requirements

- IntuneWinAppUtil.exe (in `.\Tools\` folder)
- PowerShell 5.1 or later
- Administrator privileges (for `-TestInstalls`)

---

## Output Structure
```
OutputFolder/
├── Packages/                    # .intunewin files
└── Reports/                     # CSV with detection rules & commands
```

---

## CSV Report Columns

| Column | Description |
|--------|-------------|
| **InstallCommand** | Ready-to-use install command |
| **UninstallCommand** | Ready-to-use uninstall command |
| **DetectionType** | MSI / Registry / File |
| **MSIProductCode** | Product GUID for MSI detection |
| **RegistryKeyPath** | Full registry path for detection |
| **RegistryValueName** | Registry value to check |
| **RegistryValue** | Expected value for detection |
| **Is32BitApp** | Yes/No for Intune configuration |

---

## Supported Installers

- **MSI** - Automatic product code extraction
- **EXE** - Inno Setup, NSIS, InstallShield (auto-detected)

---

## Example Workflow

1. **Place installers in source folder**
2. **Run script**
```powershell
   .\New-IntuneWin32Package-Enhanced.ps1 -SourceFolder "C:\Apps" -OutputFolder "C:\Packages"
```
3. **Review CSV report** for install/uninstall commands and detection rules
4. **Upload .intunewin packages** to Intune
5. **Copy/paste** commands and detection settings from CSV into Intune

---

## Detection Rule Examples

### MSI Application
```
Detection Type: MSI
MSI Product Code: {037F9FCB-D6E4-4D2D-A5DD-90EB7AE651F}
```

### Registry Detection
```
Detection Type: Registry
Key Path: HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\CPUID CPU-Z_is1
Value Name: DisplayName
Operator: Equals
Value: CPUID CPU-Z 2.17
32-bit App: No
```

---

## Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-SourceFolder` | Yes | Folder containing installers |
| `-OutputFolder` | Yes | Output location for packages and reports |
| `-IntuneWinAppUtilPath` | No | Path to IntuneWinAppUtil.exe (default: `.\Tools\IntuneWinAppUtil.exe`) |
| `-TestInstalls` | No | Test installations before packaging (requires admin) |
| `-SkipMSIExtraction` | No | Skip MSI product code extraction |

---

## Troubleshooting

**"IntuneWinAppUtil.exe not found"**  
→ Ensure IntuneWinAppUtil.exe is in `.\Tools\` folder

**"Packaging failed"**  
→ Check installer file is not corrupted  
→ Ensure filename has no special characters

**Install command doesn't work**  
→ Run with `-TestInstalls` to auto-detect working switches  
→ Manually test: `.\installer.exe /VERYSILENT /NORESTART`

**Detection rule not found**  
→ Install app manually on test machine  
→ Run registry export script to find keys  
→ Or use file detection as fallback

---

## Known Limitations

- Install testing requires administrator privileges
- Some proprietary installers may not support standard silent switches
- Registry detection requires app to be installed for auto-discovery

---

## Version History

**v4.0** - Enhanced detection, MSI extraction, intelligent switch testing  
**v3.0** - Simplified direct installation  
**v2.0** - Wrapper script generation  
**v1.0** - Basic bulk packaging  

---

## Related Scripts

- **Get-MSIProductCode.ps1** - Extract MSI product codes from folder
- **Export-IntuneAppDetection.ps1** - Export installed app registry details 

# Basic usage
.\New-IntuneWin32Package-Enhanced.ps1 -SourceFolder "C:\Apps" -OutputFolder "C:\Packages"

# With install testing (recommended if you have admin rights)
.\New-IntuneWin32Package-Enhanced.ps1 -SourceFolder "C:\Apps" -OutputFolder "C:\Packages" -TestInstalls

# Skip MSI extraction (faster if not needed)
.\New-IntuneWin32Package-Enhanced.ps1 -SourceFolder "C:\Apps" -OutputFolder "C:\Packages" -SkipMSIExtraction

---

*Last updated: November 2025*