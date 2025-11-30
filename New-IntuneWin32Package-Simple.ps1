<#
.SYNOPSIS
    Bulk packaging tool for creating Intune Win32 application packages (.intunewin files)

.DESCRIPTION
    Simple, reliable bulk packaging for Intune Win32 apps
    - Processes multiple applications at once
    - Preserves original application names
    - Generates detection scripts
    - Direct installer execution (no wrapper scripts)
    - CSV report with install/uninstall commands

.PARAMETER SourceFolder
    Folder containing application installers (EXE, MSI, etc.)

.PARAMETER OutputFolder
    Folder where .intunewin packages will be created

.PARAMETER IntuneWinAppUtilPath
    Path to IntuneWinAppUtil.exe (defaults to .\Tools\IntuneWinAppUtil.exe)

.PARAMETER GenerateDetectionScripts
    Switch to auto-generate PowerShell detection scripts

.PARAMETER ConfigFile
    Optional JSON config file for app-specific install parameters

.EXAMPLE
    .\New-IntuneWin32Package.ps1 -SourceFolder "C:\Apps" -OutputFolder "C:\IntunePackages"

.EXAMPLE
    .\New-IntuneWin32Package.ps1 -SourceFolder "C:\Apps" -OutputFolder "C:\Packages" -GenerateDetectionScripts

.NOTES
    Author: Orion
    Version: 3.0 - Simplified (No wrapper scripts)
    Requires: IntuneWinAppUtil.exe
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
    [switch]$GenerateDetectionScripts,

    [Parameter(Mandatory = $false)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$ConfigFile
)

#Requires -Version 5.1

# Initialize script
$ErrorActionPreference = "Stop"
$ProgressPreference = "Continue"

# Script variables
$Script:ProcessedApps = @()
$Script:FailedApps = @()
$Script:StartTime = Get-Date

#region Functions

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        'Info'    { 'Cyan' }
        'Success' { 'Green' }
        'Warning' { 'Yellow' }
        'Error'   { 'Red' }
    }
    
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Test-IntuneWinAppUtil {
    param([string]$Path)
    
    if (-not (Test-Path $Path)) {
        Write-Log "IntuneWinAppUtil.exe not found at: $Path" -Level Error
        Write-Log "Please ensure IntuneWinAppUtil.exe is available at the specified path." -Level Error
        return $false
    }
    
    Write-Log "Found IntuneWinAppUtil.exe at: $Path" -Level Success
    return $true
}

function Get-AppConfig {
    param([string]$ConfigPath)
    
    if ([string]::IsNullOrEmpty($ConfigPath)) {
        return $null
    }
    
    try {
        $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        Write-Log "Loaded configuration from: $ConfigPath" -Level Success
        return $config
    }
    catch {
        Write-Log "Failed to load config file: $_" -Level Warning
        return $null
    }
}

function New-AppWorkingDirectory {
    param(
        [string]$AppName,
        [string]$SourceFile,
        [string]$BaseOutputPath
    )
    
    # Create unique working directory for this app
    $workingDir = Join-Path $BaseOutputPath "Working_$AppName"
    $sourceDir = Join-Path $workingDir "Source"
    
    # Clean up if exists
    if (Test-Path $workingDir) {
        Remove-Item $workingDir -Recurse -Force
    }
    
    # Create directories
    New-Item -ItemType Directory -Path $sourceDir -Force | Out-Null
    
    # Copy source file to working directory
    Copy-Item $SourceFile -Destination $sourceDir -Force
    
    return @{
        WorkingDir = $workingDir
        SourceDir = $sourceDir
        SetupFile = Split-Path $SourceFile -Leaf
    }
}

function Invoke-IntuneWinPackaging {
    param(
        [string]$SourceDir,
        [string]$SetupFile,
        [string]$OutputDir,
        [string]$IntuneWinAppUtilPath
    )
    
    try {
        # Build IntuneWinAppUtil.exe arguments
        $arguments = @(
            "-c", "`"$SourceDir`""
            "-s", "`"$SetupFile`""
            "-o", "`"$OutputDir`""
            "-q"  # Quiet mode
        )
        
        Write-Log "Packaging: $SetupFile" -Level Info
        Write-Log "Command: $IntuneWinAppUtilPath $($arguments -join ' ')" -Level Info
        
        # Execute IntuneWinAppUtil.exe
        $process = Start-Process -FilePath $IntuneWinAppUtilPath `
                                 -ArgumentList $arguments `
                                 -Wait `
                                 -NoNewWindow `
                                 -PassThru `
                                 -RedirectStandardOutput (Join-Path $env:TEMP "intunewin_stdout.txt") `
                                 -RedirectStandardError (Join-Path $env:TEMP "intunewin_stderr.txt")
        
        if ($process.ExitCode -eq 0) {
            Write-Log "Successfully packaged: $SetupFile" -Level Success
            return $true
        }
        else {
            $stderr = Get-Content (Join-Path $env:TEMP "intunewin_stderr.txt") -Raw -ErrorAction SilentlyContinue
            Write-Log "Packaging failed with exit code $($process.ExitCode): $stderr" -Level Error
            return $false
        }
    }
    catch {
        Write-Log "Exception during packaging: $_" -Level Error
        return $false
    }
}

function New-DetectionScript {
    param(
        [string]$AppName,
        [string]$SetupFile,
        [string]$OutputPath
    )
    
    $fileName = Split-Path $SetupFile -Leaf
    $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    
    # Default detection: Check for file in Program Files
    $detectionScript = @"
# Detection script for $AppName
# Auto-generated by New-IntuneWin32Package.ps1
# Created: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

`$AppName = "$fileNameWithoutExt"
`$ProgramFilesPaths = @(
    "`${env:ProgramFiles}\`$AppName",
    "`${env:ProgramFiles(x86)}\`$AppName"
)

foreach (`$path in `$ProgramFilesPaths) {
    if (Test-Path `$path) {
        Write-Host "Application detected at: `$path"
        exit 0
    }
}

# Application not found
exit 1
"@
    
    $scriptPath = Join-Path $OutputPath "Detect_$AppName.ps1"
    Set-Content -Path $scriptPath -Value $detectionScript -Force
    Write-Log "Generated detection script: $scriptPath" -Level Success
    
    return $scriptPath
}

function Get-DefaultInstallCommand {
    param(
        [string]$FileName,
        [object]$Config,
        [string]$AppName
    )
    
    # Check if config has custom command for this app
    if ($Config -and $Config.applications) {
        $appConfig = $Config.applications | Where-Object { $_.name -eq $AppName }
        if ($appConfig -and $appConfig.installCommand) {
            return $appConfig.installCommand
        }
    }
    
    # Default based on file extension
    $extension = [System.IO.Path]::GetExtension($FileName).ToLower()
    
    switch ($extension) {
        ".exe" { return "$FileName /S" }
        ".msi" { return "msiexec /i `"$FileName`" /qn" }
        default { return $FileName }
    }
}

function Get-DefaultUninstallCommand {
    param(
        [string]$FileName,
        [object]$Config,
        [string]$AppName
    )
    
    # Check if config has custom command for this app
    if ($Config -and $Config.applications) {
        $appConfig = $Config.applications | Where-Object { $_.name -eq $AppName }
        if ($appConfig -and $appConfig.uninstallCommand) {
            return $appConfig.uninstallCommand
        }
    }
    
    # Default based on file extension
    $extension = [System.IO.Path]::GetExtension($FileName).ToLower()
    
    switch ($extension) {
        ".exe" { return "$FileName /S" }
        ".msi" { return "msiexec /x `"$FileName`" /qn" }
        default { return $FileName }
    }
}

#endregion

#region Main Script

Write-Log "========================================" -Level Info
Write-Log "Intune Win32 Package Bulk Creator v3.0" -Level Info
Write-Log "Simple & Reliable - Direct Installation" -Level Info
Write-Log "========================================" -Level Info
Write-Log "Source Folder: $SourceFolder" -Level Info
Write-Log "Output Folder: $OutputFolder" -Level Info
Write-Log "========================================" -Level Info

# Validate IntuneWinAppUtil.exe
if (-not (Test-IntuneWinAppUtil -Path $IntuneWinAppUtilPath)) {
    exit 1
}

# Create output directories
$packagesFolder = Join-Path $OutputFolder "Packages"
$detectionScriptsFolder = Join-Path $OutputFolder "DetectionScripts"
$logsFolder = Join-Path $OutputFolder "Logs"

@($packagesFolder, $detectionScriptsFolder, $logsFolder) | ForEach-Object {
    if (-not (Test-Path $_)) {
        New-Item -ItemType Directory -Path $_ -Force | Out-Null
        Write-Log "Created directory: $_" -Level Info
    }
}

# Load configuration if provided
$config = Get-AppConfig -ConfigPath $ConfigFile

# Get all installer files
$installerExtensions = @("*.exe", "*.msi")
$installers = Get-ChildItem -Path $SourceFolder -Include $installerExtensions -File -Recurse

if ($installers.Count -eq 0) {
    Write-Log "No installer files found in: $SourceFolder" -Level Warning
    exit 0
}

Write-Log "Found $($installers.Count) installer(s) to process" -Level Info
Write-Log "========================================" -Level Info

# Process each installer
$counter = 1
foreach ($installer in $installers) {
    $appName = [System.IO.Path]::GetFileNameWithoutExtension($installer.Name)
    
    Write-Log "[$counter/$($installers.Count)] Processing: $appName" -Level Info
    
    try {
        # Create working directory structure
        $workingDirs = New-AppWorkingDirectory -AppName $appName `
                                               -SourceFile $installer.FullName `
                                               -BaseOutputPath $OutputFolder
        
        # Package the application
        $packageSuccess = Invoke-IntuneWinPackaging -SourceDir $workingDirs.SourceDir `
                                                     -SetupFile $workingDirs.SetupFile `
                                                     -OutputDir $workingDirs.WorkingDir `
                                                     -IntuneWinAppUtilPath $IntuneWinAppUtilPath
        
        if ($packageSuccess) {
            # Find the generated .intunewin file
            $intuneWinFile = Get-ChildItem -Path $workingDirs.WorkingDir -Filter "*.intunewin" -File | Select-Object -First 1
            
            if ($intuneWinFile) {
                # Move to final packages folder with original app name
                $finalPackagePath = Join-Path $packagesFolder "$appName.intunewin"
                Move-Item $intuneWinFile.FullName -Destination $finalPackagePath -Force
                
                Write-Log "Package created: $finalPackagePath" -Level Success
                
                # Generate detection script if requested
                if ($GenerateDetectionScripts) {
                    New-DetectionScript -AppName $appName `
                                       -SetupFile $installer.FullName `
                                       -OutputPath $detectionScriptsFolder | Out-Null
                }
                
                # Get install commands (from config or defaults)
                $installCmd = Get-DefaultInstallCommand -FileName $installer.Name -Config $config -AppName $appName
                $uninstallCmd = Get-DefaultUninstallCommand -FileName $installer.Name -Config $config -AppName $appName
                
                $Script:ProcessedApps += [PSCustomObject]@{
                    AppName = $appName
                    FileName = $installer.Name
                    SourceFile = $installer.FullName
                    PackagePath = $finalPackagePath
                    InstallCommand = $installCmd
                    UninstallCommand = $uninstallCmd
                    Status = "Success"
                }
            }
            else {
                throw "IntuneWin file not found after packaging"
            }
        }
        else {
            throw "Packaging process failed"
        }
        
        # Clean up working directory
        if (Test-Path $workingDirs.WorkingDir) {
            Remove-Item $workingDirs.WorkingDir -Recurse -Force
        }
    }
    catch {
        Write-Log "Failed to process $appName : $_" -Level Error
        
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

#region Summary Report

$endTime = Get-Date
$duration = $endTime - $Script:StartTime

Write-Log "========================================" -Level Info
Write-Log "Processing Complete" -Level Success
Write-Log "========================================" -Level Info
Write-Log "Total Installers: $($installers.Count)" -Level Info
Write-Log "Successfully Packaged: $($Script:ProcessedApps.Count)" -Level Success
Write-Log "Failed: $($Script:FailedApps.Count)" -Level $(if ($Script:FailedApps.Count -gt 0) { 'Warning' } else { 'Info' })
Write-Log "Duration: $($duration.ToString('hh\:mm\:ss'))" -Level Info
Write-Log "========================================" -Level Info

# Export summary report
$reportPath = Join-Path $logsFolder "PackagingReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

if ($Script:ProcessedApps.Count -gt 0) {
    $Script:ProcessedApps | Export-Csv -Path $reportPath -NoTypeInformation -Force
    Write-Log "Report exported: $reportPath" -Level Success
}

# Display processed apps
if ($Script:ProcessedApps.Count -gt 0) {
    Write-Log "`nSuccessfully Packaged Applications:" -Level Success
    $Script:ProcessedApps | Format-Table AppName, FileName, InstallCommand, UninstallCommand -AutoSize
}

# Display failed apps
if ($Script:FailedApps.Count -gt 0) {
    Write-Log "`nFailed Applications:" -Level Error
    $Script:FailedApps | Format-Table AppName, Error -AutoSize
}

Write-Log "`n========================================" -Level Info
Write-Log "Next Steps - Configure in Intune:" -Level Info
Write-Log "========================================" -Level Info
Write-Log "Packages location: $packagesFolder" -Level Info
if ($GenerateDetectionScripts) {
    Write-Log "Detection scripts location: $detectionScriptsFolder" -Level Info
}
Write-Log "`nFor each app, use these commands in Intune:" -Level Info
$Script:ProcessedApps | ForEach-Object {
    Write-Log "`n$($_.AppName):" -Level Success
    Write-Log "  Install:   $($_.InstallCommand)" -Level Info
    Write-Log "  Uninstall: $($_.UninstallCommand)" -Level Info
}

#endregion
