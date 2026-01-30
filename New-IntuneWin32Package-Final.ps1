<#
.SYNOPSIS
    Enhanced bulk packaging tool for Intune Win32 applications with intelligent detection
.NOTES
    Author: Orion / Refined by Gemini
    Version: 4.1 - Added COM cleanup, Post-test registry capture, and Reboot suppression
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$SourceFolder,

    [Parameter(Mandatory = $true)]
    [string]$OutputFolder,

    [Parameter(Mandatory = $false)]
    [string]$IntuneWinAppUtilPath = ".\Tools\IntuneWinAppUtil.exe",

    [Parameter(Mandatory = $false)]
    [switch]$TestInstalls,

    [Parameter(Mandatory = $false)]
    [switch]$SkipMSIExtraction
)

#Requires -Version 5.1

# Initialize
$ErrorActionPreference = "Stop"
$Script:ProcessedApps = @()
$Script:FailedApps = @()
$Script:StartTime = Get-Date

#region Helper Functions

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = @{ 'Info' = 'Cyan'; 'Success' = 'Green'; 'Warning' = 'Yellow'; 'Error' = 'Red' }[$Level]
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Get-MSIProductInfo {
    param([string]$MsiPath)
    $WindowsInstaller = $null
    $Database = $null
    try {
        $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
        $Database = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($MsiPath, 0))
        
        $properties = @('ProductCode', 'ProductName', 'ProductVersion', 'Manufacturer')
        $result = @{}
        
        foreach ($prop in $properties) {
            $View = $Database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $Database, "SELECT Value FROM Property WHERE Property = '$prop'")
            $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
            $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
            if ($Record) {
                $result[$prop] = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
            }
        }
        
        return [PSCustomObject]@{
            ProductCode    = $result.ProductCode
            ProductName    = $result.ProductName
            ProductVersion = $result.ProductVersion
            Manufacturer   = $result.Manufacturer
            RegistryPath   = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($result.ProductCode)"
        }
    }
    catch {
        Write-Log "Failed to extract MSI info from $MsiPath : $_" -Level Warning
        return $null
    }
    finally {
        # Critical COM cleanup to prevent file locks
        if ($Database) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Database) | Out-Null }
        if ($WindowsInstaller) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Get-SilentInstallSwitches {
    param([string]$Extension)
    $switchSets = @{
        '.exe' = @(
            @('/VERYSILENT', '/NORESTART', '/SUPPRESSMSGBOXES'),
            @('/S'),
            @('/silent', '/norestart'),
            @('/q', '/norestart')
        )
        '.msi' = @(
            @('/i', '/qn', '/norestart', 'REBOOT=ReallySuppress')
        )
    }
    return $switchSets[$Extension.ToLower()]
}

function Find-InstalledAppRegistry {
    param([string]$AppName)
    $UninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($Path in $UninstallPaths) {
        $Apps = Get-ItemProperty $Path -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*$AppName*" -or $_.Publisher -like "*$AppName*" }
        if ($Apps) {
            $App = $Apps | Select-Object -First 1
            $RegPath = $App.PSPath -replace 'Microsoft.PowerShell.Core\\Registry::', ''
            return [PSCustomObject]@{
                DisplayName     = $App.DisplayName
                DisplayVersion  = $App.DisplayVersion
                Publisher       = $App.Publisher
                KeyPath         = $RegPath -replace 'HKEY_LOCAL_MACHINE\\', 'HKLM\'
                ValueName       = 'DisplayName'
                DetectionValue  = $App.DisplayName
                Is32Bit         = $RegPath -like "*WOW6432Node*"
                UninstallString = $App.UninstallString
                QuietUninstallString = $App.QuietUninstallString
            }
        }
    }
    return $null
}

#endregion

# --- Main Script ---

# Resolve Tool Path early
try { $IntuneWinAppUtilPath = (Resolve-Path $IntuneWinAppUtilPath).Path } catch { Write-Log "Tool not found!" -Level Error; exit 1 }

Write-Log "Starting Enhanced Packager v4.1" -Level Info

$installers = Get-ChildItem -Path $SourceFolder -Include @("*.exe", "*.msi") -File -Recurse
$packagesFolder = New-Item -ItemType Directory -Path (Join-Path $OutputFolder "Packages") -Force
$reportsFolder  = New-Item -ItemType Directory -Path (Join-Path $OutputFolder "Reports") -Force

foreach ($installer in $installers) {
    $appName = $installer.BaseName
    $extension = $installer.Extension.ToLower()
    $registryInfo = $null
    $msiInfo = $null
    
    Write-Log "Processing: $appName" -Level Info

    try {
        # 1. MSI Extraction
        if ($extension -eq '.msi' -and -not $SkipMSIExtraction) {
            $msiInfo = Get-MSIProductInfo -MsiPath $installer.FullName
        }

        # 2. Command Discovery & Optional Testing
        $switches = (Get-SilentInstallSwitches -Extension $extension)[0] # Default to first set
        $installCmd = if ($extension -eq '.msi') { "msiexec /i `"$($installer.Name)`" /qn /norestart REBOOT=ReallySuppress" } else { "$($installer.Name) $($switches -join ' ')" }

        if ($TestInstalls) {
            if ($PSCmdlet.ShouldProcess($appName, "Install test on local machine")) {
                Write-Log "Running test installation..." -Level Warning
                $p = Start-Process -FilePath $installer.FullName -ArgumentList $switches -Wait -PassThru -NoNewWindow
                if ($p.ExitCode -in @(0, 3010)) {
                    Write-Log "Test success. Capturing registry..." -Level Success
                    $registryInfo = Find-InstalledAppRegistry -AppName $appName
                }
            }
        }

        # 3. Packaging
        $workingDir = New-Item -ItemType Directory -Path (Join-Path $OutputFolder "Working_$appName") -Force
        $sourceDir = New-Item -ItemType Directory -Path (Join-Path $workingDir "Source") -Force
        Copy-Item $installer.FullName -Destination $sourceDir -Force

        $utilArgs = @("-c", $sourceDir.FullName, "-s", $installer.Name, "-o", $workingDir.FullName, "-q")
        $proc = Start-Process -FilePath $IntuneWinAppUtilPath -ArgumentList $utilArgs -Wait -PassThru -NoNewWindow
        
        if ($proc.ExitCode -eq 0) {
            $outFile = Get-ChildItem $workingDir -Filter "*.intunewin" | Select-Object -First 1
            Move-Item $outFile.FullName -Destination (Join-Path $packagesFolder "$appName.intunewin") -Force
            
            # 4. Data Logging (Detection Rules)
            $Script:ProcessedApps += [PSCustomObject]@{
                AppName           = $appName
                InstallCommand    = $installCmd
                UninstallCommand  = if ($msiInfo) { "msiexec /x $($msiInfo.ProductCode) /qn" } elseif ($registryInfo.QuietUninstallString) { $registryInfo.QuietUninstallString } else { "CHECK_REGISTRY" }
                DetectionType     = if ($msiInfo) { "MSI" } elseif ($registryInfo) { "Registry" } else { "File" }
                RegistryKey       = if ($msiInfo) { $msiInfo.RegistryPath } else { $registryInfo.KeyPath }
            }
            Write-Log "Successfully packaged $appName" -Level Success
        }
    }
    catch {
        Write-Log "Error processing $appName: $_" -Level Error
        $Script:FailedApps += [PSCustomObject]@{ App = $appName; Error = $_.Exception.Message }
    }
    finally {
        if (Test-Path $workingDir) { Remove-Item $workingDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

# Export Report
$reportPath = Join-Path $reportsFolder "Report_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
$Script:ProcessedApps | Export-Csv -Path $reportPath -NoTypeInformation
Write-Log "Done! Report at $reportPath" -Level Success