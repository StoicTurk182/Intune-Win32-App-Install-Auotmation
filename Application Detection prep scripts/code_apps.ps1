<#
.SYNOPSIS
    Export all installed applications with Intune detection details to CSV
.PARAMETER OutputPath
    Path where the CSV file will be saved (optional)
#>

param(
    [string]$OutputPath = "C:\Temp\IntuneAppDetection_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Ensure output directory exists
$OutputDir = Split-Path $OutputPath -Parent
if (-not (Test-Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}

# Registry paths to check
$UninstallPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

Write-Host "`nScanning installed applications..." -ForegroundColor Cyan

$AllApps = @()

foreach ($Path in $UninstallPaths) {
    try {
        $Apps = Get-ItemProperty $Path -ErrorAction SilentlyContinue | 
            Where-Object {$_.DisplayName} |
            Sort-Object DisplayName
        
        foreach ($App in $Apps) {
            # Get the registry key name (GUID or identifier)
            $RegPath = $App.PSPath -replace 'Microsoft.PowerShell.Core\\Registry::', ''
            $KeyName = Split-Path $RegPath -Leaf
            $ParentPath = Split-Path $RegPath -Parent
            
            # Determine if it's 32-bit app on 64-bit system
            $Is32BitOn64Bit = $ParentPath -like "*WOW6432Node*"
            
            # Clean up the registry path for Intune
            $IntuneKeyPath = $ParentPath -replace 'HKEY_LOCAL_MACHINE\\', 'HKLM\'
            
            $AllApps += [PSCustomObject]@{
                'Application Name' = $App.DisplayName
                'Version' = $App.DisplayVersion
                'Publisher' = $App.Publisher
                'Install Date' = $App.InstallDate
                'Install Location' = $App.InstallLocation
                'Intune Rule Type' = 'Registry'
                'Intune Key Path' = "$IntuneKeyPath\$KeyName"
                'Intune Value Name' = 'DisplayName'
                'Intune Detection Method' = 'String comparison'
                'Intune Operator' = 'Equals'
                'Intune Value' = $App.DisplayName
                'Is 32-bit App on 64-bit' = if($Is32BitOn64Bit){"Yes"}else{"No"}
                'Registry Key Name' = $KeyName
                'Uninstall String' = $App.UninstallString
                'Quiet Uninstall String' = $App.QuietUninstallString
                'MSI Product Code' = if($KeyName -match '^\{[A-F0-9-]+\}$'){$KeyName}else{"N/A"}
                'Full Registry Path' = $RegPath
            }
        }
    }
    catch {
        Write-Warning "Error accessing $Path : $_"
    }
}

# Remove duplicates (sometimes apps appear in both locations)
$UniqueApps = $AllApps | Sort-Object 'Application Name' -Unique

Write-Host "Found $($UniqueApps.Count) unique installed applications" -ForegroundColor Green

# Export to CSV
$UniqueApps | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "`n✓ Successfully exported to: $OutputPath" -ForegroundColor Green
Write-Host "`nFile contains columns:" -ForegroundColor Yellow
Write-Host "  - Application Name, Version, Publisher" -ForegroundColor White
Write-Host "  - Intune Detection Rule (Key Path, Value Name, Operator, Value)" -ForegroundColor White
Write-Host "  - Uninstall Strings" -ForegroundColor White
Write-Host "  - MSI Product Code (if applicable)" -ForegroundColor White
Write-Host "  - Full Registry Paths`n" -ForegroundColor White

# Display preview
Write-Host "Preview (first 10 apps):" -ForegroundColor Cyan
$UniqueApps | Select-Object 'Application Name', 'Version', 'Publisher', 'Intune Key Path' -First 10 | Format-Table -AutoSize

# Open the CSV file
$OpenFile = Read-Host "`nOpen CSV file now? (Y/N)"
if ($OpenFile -eq 'Y' -or $OpenFile -eq 'y') {
    Start-Process $OutputPath
}