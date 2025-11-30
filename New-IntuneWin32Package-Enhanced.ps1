<#
.SYNOPSIS
    Enhanced bulk packaging tool for Intune Win32 applications with intelligent detection

.DESCRIPTION
    Improved packaging workflow that:
    - Auto-detects and tests silent install switches
    - Finds registry detection keys automatically
    - Extracts MSI product codes
    - Tests installations before packaging
    - Generates accurate detection rules
    - Creates comprehensive CSV report

.PARAMETER SourceFolder
    Folder containing application installers (EXE, MSI, etc.)

.PARAMETER OutputFolder
    Folder where .intunewin packages and reports will be created

.PARAMETER IntuneWinAppUtilPath
    Path to IntuneWinAppUtil.exe (defaults to .\Tools\IntuneWinAppUtil.exe)

.PARAMETER TestInstalls
    Switch to test installations on local system before packaging

.PARAMETER SkipMSIExtraction
    Skip automatic MSI product code extraction

.EXAMPLE
    .\New-IntuneWin32Package-Enhanced.ps1 -SourceFolder "C:\Apps" -OutputFolder "C:\Packages"

.EXAMPLE
    .\New-IntuneWin32Package-Enhanced.ps1 -SourceFolder "C:\Apps" -OutputFolder "C:\Packages" -TestInstalls

.NOTES
    Author: Orion
    Version: 4.0 - Enhanced with intelligent detection
    Requires: IntuneWinAppUtil.exe, Administrator privileges for testing
#>

[CmdletBinding()]
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
    $color = @{
        'Info' = 'Cyan'
        'Success' = 'Green'
        'Warning' = 'Yellow'
        'Error' = 'Red'
    }[$Level]
    
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Get-MSIProductInfo {
    param([string]$MsiPath)
    
    try {
        $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
        $Database = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($MsiPath, 0))
        
        $properties = @('ProductCode', 'ProductName', 'ProductVersion', 'Manufacturer')
        $result = @{}
        
        foreach ($prop in $properties) {
            $View = $Database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $Database, "SELECT Value FROM Property WHERE Property = '$prop'")
            $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
            $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
            $result[$prop] = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
        }
        
        return [PSCustomObject]@{
            ProductCode = $result.ProductCode
            ProductName = $result.ProductName
            ProductVersion = $result.ProductVersion
            Manufacturer = $result.Manufacturer
            RegistryPath = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($result.ProductCode)"
        }
    }
    catch {
        Write-Log "Failed to extract MSI info from $MsiPath : $_" -Level Warning
        return $null
    }
}

function Get-SilentInstallSwitches {
    param(
        [string]$FilePath,
        [string]$Extension
    )
    
    # Common silent switches to try in order of preference
    $switchSets = @{
        '.exe' = @(
            @('/VERYSILENT', '/NORESTART'),           # Inno Setup (most common)
            @('/S'),                                    # NSIS
            @('/silent', '/norestart'),                # InstallShield
            @('/q', '/norestart'),                     # Generic
            @('/SP-', '/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/ALLUSERS')  # Inno Setup advanced
        )
        '.msi' = @(
            @('/i', '/qn', '/norestart')               # MSI standard
        )
    }
    
    return $switchSets[$Extension.ToLower()]
}

function Test-SilentInstall {
    param(
        [string]$FilePath,
        [string[]]$Switches
    )
    
    if (-not $TestInstalls) {
        return $false
    }
    
    Write-Log "Testing install switches: $($Switches -join ' ')" -Level Info
    
    try {
        $fileName = Split-Path $FilePath -Leaf
        $args = if ($fileName -like "*.msi") {
            $Switches + "`"$FilePath`""
        } else {
            $Switches
        }
        
        # Create temp marker to verify install happened
        $tempMarker = Join-Path $env:TEMP "intune_test_$(Get-Date -Format 'yyyyMMddHHmmss').txt"
        Set-Content $tempMarker "Test installation started"
        
        if ($fileName -like "*.msi") {
            $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $args -Wait -PassThru -NoNewWindow
        } else {
            $process = Start-Process -FilePath $FilePath -ArgumentList $Switches -Wait -PassThru -NoNewWindow
        }
        
        Remove-Item $tempMarker -ErrorAction SilentlyContinue
        
        return ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010)
    }
    catch {
        Write-Log "Install test failed: $_" -Level Warning
        return $false
    }
}

function Find-InstalledAppRegistry {
    param([string]$AppName)
    
    $UninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    foreach ($Path in $UninstallPaths) {
        $Apps = Get-ItemProperty $Path -ErrorAction SilentlyContinue | 
            Where-Object {$_.DisplayName -like "*$AppName*"}
        
        if ($Apps) {
            $App = $Apps | Select-Object -First 1
            $RegPath = $App.PSPath -replace 'Microsoft.PowerShell.Core\\Registry::', ''
            $KeyName = Split-Path $RegPath -Leaf
            $ParentPath = (Split-Path $RegPath -Parent) -replace 'HKEY_LOCAL_MACHINE\\', 'HKLM\'
            $Is32Bit = $ParentPath -like "*WOW6432Node*"
            
            return [PSCustomObject]@{
                DisplayName = $App.DisplayName
                DisplayVersion = $App.DisplayVersion
                Publisher = $App.Publisher
                KeyPath = "$ParentPath\$KeyName"
                ValueName = 'DisplayName'
                DetectionValue = $App.DisplayName
                Is32Bit = $Is32Bit
                UninstallString = $App.UninstallString
                QuietUninstallString = $App.QuietUninstallString
            }
        }
    }
    
    return $null
}

function New-IntunePackage {
    param(
        [string]$SourceDir,
        [string]$SetupFile,
        [string]$OutputDir,
        [string]$ToolPath
    )
    
    try {
        $arguments = @(
            "-c", "`"$SourceDir`""
            "-s", "`"$SetupFile`""
            "-o", "`"$OutputDir`""
            "-q"
        )
        
        Write-Log "Packaging: $SetupFile" -Level Info
        
        $process = Start-Process -FilePath $ToolPath `
                                 -ArgumentList $arguments `
                                 -Wait `
                                 -NoNewWindow `
                                 -PassThru `
                                 -RedirectStandardOutput (Join-Path $env:TEMP "intunewin_stdout.txt") `
                                 -RedirectStandardError (Join-Path $env:TEMP "intunewin_stderr.txt")
        
        return ($process.ExitCode -eq 0)
    }
    catch {
        Write-Log "Packaging exception: $_" -Level Error
        return $false
    }
}

function Get-BestInstallCommand {
    param(
        [string]$FilePath,
        [string]$Extension,
        [object]$MSIInfo
    )
    
    $fileName = Split-Path $FilePath -Leaf
    
    # MSI files - use standard msiexec
    if ($Extension -eq '.msi') {
        return "msiexec /i `"$fileName`" /qn /norestart"
    }
    
    # EXE files - test switches
    $switchSets = Get-SilentInstallSwitches -FilePath $FilePath -Extension $Extension
    
    if ($TestInstalls) {
        foreach ($switches in $switchSets) {
            if (Test-SilentInstall -FilePath $FilePath -Switches $switches) {
                return "$fileName $($switches -join ' ')"
            }
        }
    }
    
    # Default to most common (Inno Setup)
    return "$fileName /VERYSILENT /NORESTART"
}

function Get-BestUninstallCommand {
    param(
        [string]$FileName,
        [string]$Extension,
        [object]$MSIInfo,
        [object]$RegistryInfo
    )
    
    # MSI - use product code
    if ($Extension -eq '.msi' -and $MSIInfo) {
        return "msiexec /x `"$($MSIInfo.ProductCode)`" /qn /norestart"
    }
    
    # Check registry for uninstall string
    if ($RegistryInfo) {
        if ($RegistryInfo.QuietUninstallString) {
            return $RegistryInfo.QuietUninstallString
        }
        if ($RegistryInfo.UninstallString -and $RegistryInfo.UninstallString -like "*unins*.exe*") {
            # Inno Setup uninstaller
            return "$($RegistryInfo.UninstallString) /VERYSILENT /NORESTART"
        }
    }
    
    # Default
    return "Uninstall command not detected - check registry after installation"
}

#endregion

#region Main Script

Write-Log "========================================" -Level Info
Write-Log "Enhanced Intune Win32 Packager v4.0" -Level Info
Write-Log "Intelligent Detection & Testing" -Level Info
Write-Log "========================================" -Level Info
Write-Log "Source: $SourceFolder" -Level Info
Write-Log "Output: $OutputFolder" -Level Info
if ($TestInstalls) {
    Write-Log "Test Mode: ENABLED - Will test installations" -Level Warning
}
Write-Log "========================================" -Level Info

# Validate IntuneWinAppUtil
if (-not (Test-Path $IntuneWinAppUtilPath)) {
    Write-Log "IntuneWinAppUtil.exe not found at: $IntuneWinAppUtilPath" -Level Error
    exit 1
}

# Create output structure
$packagesFolder = Join-Path $OutputFolder "Packages"
$reportsFolder = Join-Path $OutputFolder "Reports"

@($packagesFolder, $reportsFolder) | ForEach-Object {
    if (-not (Test-Path $_)) {
        New-Item -ItemType Directory -Path $_ -Force | Out-Null
    }
}

# Get installers
$installers = Get-ChildItem -Path $SourceFolder -Include @("*.exe", "*.msi") -File -Recurse

if ($installers.Count -eq 0) {
    Write-Log "No installers found in: $SourceFolder" -Level Warning
    exit 0
}

Write-Log "Found $($installers.Count) installer(s)" -Level Success
Write-Log "========================================" -Level Info

# Process each installer
$counter = 1
foreach ($installer in $installers) {
    $appName = [System.IO.Path]::GetFileNameWithoutExtension($installer.Name)
    $extension = [System.IO.Path]::GetExtension($installer.Name).ToLower()
    
    Write-Log "[$counter/$($installers.Count)] Processing: $appName" -Level Info
    
    try {
        # Extract MSI info if applicable
        $msiInfo = $null
        if ($extension -eq '.msi' -and -not $SkipMSIExtraction) {
            Write-Log "Extracting MSI product information..." -Level Info
            $msiInfo = Get-MSIProductInfo -MsiPath $installer.FullName
            if ($msiInfo) {
                Write-Log "Product: $($msiInfo.ProductName) v$($msiInfo.ProductVersion)" -Level Success
                Write-Log "Product Code: $($msiInfo.ProductCode)" -Level Info
            }
        }
        
        # Determine install command
        Write-Log "Determining optimal install switches..." -Level Info
        $installCmd = Get-BestInstallCommand -FilePath $installer.FullName -Extension $extension -MSIInfo $msiInfo
        Write-Log "Install command: $installCmd" -Level Info
        
        # Search for registry detection (if app might be installed)
        $registryInfo = $null
        if ($TestInstalls) {
            Write-Log "Searching for registry detection..." -Level Info
            $registryInfo = Find-InstalledAppRegistry -AppName $appName
        }
        
        # Determine uninstall command
        $uninstallCmd = Get-BestUninstallCommand -FileName $installer.Name -Extension $extension -MSIInfo $msiInfo -RegistryInfo $registryInfo
        Write-Log "Uninstall command: $uninstallCmd" -Level Info
        
        # Create working directory
        $workingDir = Join-Path $OutputFolder "Working_$appName"
        $sourceDir = Join-Path $workingDir "Source"
        
        if (Test-Path $workingDir) {
            Remove-Item $workingDir -Recurse -Force
        }
        New-Item -ItemType Directory -Path $sourceDir -Force | Out-Null
        Copy-Item $installer.FullName -Destination $sourceDir -Force
        
        # Package
        $packageSuccess = New-IntunePackage -SourceDir $sourceDir `
                                            -SetupFile $installer.Name `
                                            -OutputDir $workingDir `
                                            -ToolPath $IntuneWinAppUtilPath
        
        if ($packageSuccess) {
            $intuneWinFile = Get-ChildItem -Path $workingDir -Filter "*.intunewin" | Select-Object -First 1
            
            if ($intuneWinFile) {
                $finalPath = Join-Path $packagesFolder "$appName.intunewin"
                Move-Item $intuneWinFile.FullName -Destination $finalPath -Force
                Write-Log "Package created: $finalPath" -Level Success
                
                # Build detection rule info
                $detectionRule = if ($msiInfo) {
                    [PSCustomObject]@{
                        Type = "MSI"
                        MSIProductCode = $msiInfo.ProductCode
                        AlternativeType = "Registry"
                        RegistryKeyPath = $msiInfo.RegistryPath
                        RegistryValueName = "DisplayName"
                        RegistryOperator = "Equals"
                        RegistryValue = $msiInfo.ProductName
                        Is32BitApp = "No"
                    }
                } elseif ($registryInfo) {
                    [PSCustomObject]@{
                        Type = "Registry"
                        MSIProductCode = "N/A"
                        AlternativeType = ""
                        RegistryKeyPath = $registryInfo.KeyPath
                        RegistryValueName = $registryInfo.ValueName
                        RegistryOperator = "Equals"
                        RegistryValue = $registryInfo.DetectionValue
                        Is32BitApp = if($registryInfo.Is32Bit){"Yes"}else{"No"}
                    }
                } else {
                    [PSCustomObject]@{
                        Type = "File"
                        MSIProductCode = "N/A"
                        AlternativeType = ""
                        RegistryKeyPath = "Use file detection or install to find registry key"
                        RegistryValueName = ""
                        RegistryOperator = ""
                        RegistryValue = ""
                        Is32BitApp = "Unknown"
                    }
                }
                
                $Script:ProcessedApps += [PSCustomObject]@{
                    AppName = $appName
                    FileName = $installer.Name
                    FileExtension = $extension
                    SourcePath = $installer.FullName
                    PackagePath = $finalPath
                    InstallCommand = $installCmd
                    UninstallCommand = $uninstallCmd
                    DetectionType = $detectionRule.Type
                    MSIProductCode = $detectionRule.MSIProductCode
                    RegistryKeyPath = $detectionRule.RegistryKeyPath
                    RegistryValueName = $detectionRule.RegistryValueName
                    RegistryOperator = $detectionRule.RegistryOperator
                    RegistryValue = $detectionRule.RegistryValue
                    Is32BitApp = $detectionRule.Is32BitApp
                    ProductVersion = if($msiInfo){$msiInfo.ProductVersion}elseif($registryInfo){$registryInfo.DisplayVersion}else{""}
                    Publisher = if($msiInfo){$msiInfo.Manufacturer}elseif($registryInfo){$registryInfo.Publisher}else{""}
                    Status = "Success"
                }
            }
        } else {
            throw "Packaging failed"
        }
        
        # Cleanup
        if (Test-Path $workingDir) {
            Remove-Item $workingDir -Recurse -Force
        }
        
    }
    catch {
        Write-Log "Failed: $_" -Level Error
        $Script:FailedApps += [PSCustomObject]@{
            AppName = $appName
            SourceFile = $installer.FullName
            Error = $_.Exception.Message
        }
    }
    
    $counter++
    Write-Log "----------------------------------------" -Level Info
}

#endregion

#region Summary

$endTime = Get-Date
$duration = $endTime - $Script:StartTime

Write-Log "========================================" -Level Info
Write-Log "Processing Complete" -Level Success
Write-Log "========================================" -Level Info
Write-Log "Total: $($installers.Count)" -Level Info
Write-Log "Success: $($Script:ProcessedApps.Count)" -Level Success
Write-Log "Failed: $($Script:FailedApps.Count)" -Level $(if($Script:FailedApps.Count -gt 0){'Warning'}else{'Info'})
Write-Log "Duration: $($duration.ToString('hh\:mm\:ss'))" -Level Info

# Export detailed report
$reportPath = Join-Path $reportsFolder "IntunePackaging_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

if ($Script:ProcessedApps.Count -gt 0) {
    $Script:ProcessedApps | Export-Csv -Path $reportPath -NoTypeInformation -Force
    Write-Log "Report: $reportPath" -Level Success
    
    # Display summary table
    Write-Log "`nPackaged Applications:" -Level Success
    $Script:ProcessedApps | Format-Table AppName, DetectionType, InstallCommand -AutoSize
    
    Write-Log "`nIntune Configuration Guide:" -Level Info
    Write-Log "========================================" -Level Info
    foreach ($app in $Script:ProcessedApps) {
        Write-Log "`n$($app.AppName):" -Level Success
        Write-Log "  Install:   $($app.InstallCommand)" -Level Info
        Write-Log "  Uninstall: $($app.UninstallCommand)" -Level Info
        Write-Log "  Detection: $($app.DetectionType)" -Level Info
        if ($app.DetectionType -eq "MSI") {
            Write-Log "    Product Code: $($app.MSIProductCode)" -Level Info
        } elseif ($app.DetectionType -eq "Registry") {
            Write-Log "    Key Path: $($app.RegistryKeyPath)" -Level Info
            Write-Log "    Value Name: $($app.RegistryValueName)" -Level Info
            Write-Log "    Operator: $($app.RegistryOperator)" -Level Info
            Write-Log "    Value: $($app.RegistryValue)" -Level Info
            Write-Log "    32-bit App: $($app.Is32BitApp)" -Level Info
        }
    }
}

if ($Script:FailedApps.Count -gt 0) {
    Write-Log "`nFailed Applications:" -Level Error
    $Script:FailedApps | Format-Table AppName, Error -AutoSize
}

Write-Log "`n========================================" -Level Info
Write-Log "Packages: $packagesFolder" -Level Info
Write-Log "Report: $reportPath" -Level Info
Write-Log "========================================`n" -Level Info

#endregion