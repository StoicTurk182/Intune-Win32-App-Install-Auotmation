# New-IntuneWin32Package-Enhanced

Bulk packaging tool for Intune Win32 applications with intelligent detection of silent install switches, registry keys, and MSI product codes.

## Requirements

- PowerShell 5.1+
- IntuneWinAppUtil.exe (Microsoft Win32 Content Prep Tool)
- Administrator privileges (if using -TestInstalls)

Download IntuneWinAppUtil.exe from: https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool

## Quick Start

```powershell
.\New-IntuneWin32Package-Enhanced.ps1 -SourceFolder "C:\Apps" -OutputFolder "C:\Packages"
```

## Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| SourceFolder | Yes | Folder containing EXE/MSI installers |
| OutputFolder | Yes | Destination for .intunewin packages |
| IntuneWinAppUtilPath | No | Path to IntuneWinAppUtil.exe (default: .\Tools\IntuneWinAppUtil.exe) |
| TestInstalls | No | Test installations locally before packaging |
| SkipMSIExtraction | No | Skip MSI product code extraction |

## Usage Examples

Basic packaging:

```powershell
.\New-IntuneWin32Package-Enhanced.ps1 -SourceFolder "C:\Installers" -OutputFolder "C:\IntunePackages"
```

With custom tool path:

```powershell
.\New-IntuneWin32Package-Enhanced.ps1 -SourceFolder "C:\Apps" -OutputFolder "C:\Output" -IntuneWinAppUtilPath "C:\Tools\IntuneWinAppUtil.exe"
```

With installation testing (requires admin):

```powershell
.\New-IntuneWin32Package-Enhanced.ps1 -SourceFolder "C:\Apps" -OutputFolder "C:\Output" -TestInstalls
```

## Output Structure

```
OutputFolder/
├── Packages/
│   ├── AppName1.intunewin
│   └── AppName2.intunewin
└── Reports/
    └── IntunePackaging_YYYYMMDD_HHmmss.csv
```

## CSV Report Fields

The generated CSV includes all information needed for Intune deployment:

- AppName, FileName, SourcePath, PackagePath
- InstallCommand, UninstallCommand
- DetectionType (MSI/Registry/File)
- MSIProductCode
- RegistryKeyPath, RegistryValueName, RegistryOperator, RegistryValue
- Is32BitApp, ProductVersion, Publisher

## Supported Installer Types

| Type | Silent Switches Tested |
|------|------------------------|
| MSI | /qn /norestart |
| EXE (Inno Setup) | /VERYSILENT /NORESTART |
| EXE (NSIS) | /S |
| EXE (InstallShield) | /silent /norestart |

## Detection Rules

The script automatically determines the best detection method:

1. MSI installers: Uses ProductCode from MSI metadata
2. EXE installers: Searches registry uninstall keys for matching DisplayName
3. Fallback: File-based detection (manual configuration required)

## References

- Microsoft Win32 Content Prep Tool: https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool
- Intune Win32 App Management: https://learn.microsoft.com/en-us/mem/intune/apps/apps-win32-app-management
- Win32 App Detection Rules: https://learn.microsoft.com/en-us/mem/intune/apps/apps-win32-add#step-4-detection-rules
- MSI Silent Install Switches: https://learn.microsoft.com/en-us/windows/win32/msi/standard-installer-command-line-options
