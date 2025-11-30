<#
.SYNOPSIS
    Extract MSI Product Codes from folder or file with interactive selection
.DESCRIPTION
    Scans MSI files and extracts product codes for Intune detection rules
.PARAMETER Path
    Path to MSI file or folder containing MSI files (optional - prompts if not provided)
.PARAMETER Recursive
    Search subfolders for MSI files
.PARAMETER ExportCSV
    Export results to CSV file
#>

param(
    [string]$Path,
    [switch]$Recursive,
    [switch]$ExportCSV
)

function Get-MSIProductCode {
    param([string]$MsiPath)
    
    try {
        $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
        $Database = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($MsiPath, 0))
        
        # Get Product Code
        $View = $Database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $Database, "SELECT Value FROM Property WHERE Property = 'ProductCode'")
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
        $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
        $ProductCode = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
        
        # Get Product Name
        $View = $Database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $Database, "SELECT Value FROM Property WHERE Property = 'ProductName'")
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
        $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
        $ProductName = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
        
        # Get Product Version
        $View = $Database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $Database, "SELECT Value FROM Property WHERE Property = 'ProductVersion'")
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
        $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
        $ProductVersion = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
        
        # Get Manufacturer
        $View = $Database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $Database, "SELECT Value FROM Property WHERE Property = 'Manufacturer'")
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
        $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
        $Manufacturer = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
        
        return [PSCustomObject]@{
            FileName = Split-Path $MsiPath -Leaf
            FilePath = $MsiPath
            ProductName = $ProductName
            ProductVersion = $ProductVersion
            Manufacturer = $Manufacturer
            ProductCode = $ProductCode
            IntuneRegistryPath = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$ProductCode"
        }
    }
    catch {
        Write-Warning "Failed to extract info from $MsiPath : $_"
        return $null
    }
}

function Show-FolderBrowser {
    param([string]$Description = "Select folder containing MSI files")
    
    Add-Type -AssemblyName System.Windows.Forms
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.Description = $Description
    $FolderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
    
    if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $FolderBrowser.SelectedPath
    }
    return $null
}

function Show-FileBrowser {
    param([string]$Title = "Select MSI file")
    
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FileBrowser.Title = $Title
    $FileBrowser.Filter = "MSI Files (*.msi)|*.msi|All Files (*.*)|*.*"
    $FileBrowser.Multiselect = $false
    
    if ($FileBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $FileBrowser.FileName
    }
    return $null
}

# Main Script
Clear-Host
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  MSI Product Code Extractor" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# If no path provided, prompt user
if (-not $Path) {
    Write-Host "Select input type:" -ForegroundColor Yellow
    Write-Host "  1. Single MSI file" -ForegroundColor White
    Write-Host "  2. Folder containing MSI files" -ForegroundColor White
    Write-Host "  3. Enter path manually`n" -ForegroundColor White
    
    $Choice = Read-Host "Enter choice (1-3)"
    
    switch ($Choice) {
        "1" {
            Write-Host "Opening file browser..." -ForegroundColor Cyan
            $Path = Show-FileBrowser
            if (-not $Path) {
                Write-Host "No file selected. Exiting." -ForegroundColor Red
                exit
            }
        }
        "2" {
            Write-Host "Opening folder browser..." -ForegroundColor Cyan
            $Path = Show-FolderBrowser
            if (-not $Path) {
                Write-Host "No folder selected. Exiting." -ForegroundColor Red
                exit
            }
        }
        "3" {
            $Path = Read-Host "Enter full path to MSI file or folder"
        }
        default {
            Write-Host "Invalid choice. Exiting." -ForegroundColor Red
            exit
        }
    }
}

# Validate path exists
if (-not (Test-Path $Path)) {
    Write-Host "ERROR: Path not found: $Path" -ForegroundColor Red
    exit
}

# Determine if path is file or folder
$IsFolder = (Get-Item $Path).PSIsContainer

# Get MSI files
if ($IsFolder) {
    Write-Host "`nScanning folder: $Path" -ForegroundColor Cyan
    
    if ($Recursive) {
        $MsiFiles = Get-ChildItem -Path $Path -Filter "*.msi" -Recurse -File
    } else {
        $MsiFiles = Get-ChildItem -Path $Path -Filter "*.msi" -File
    }
    
    if ($MsiFiles.Count -eq 0) {
        Write-Host "No MSI files found in: $Path" -ForegroundColor Yellow
        exit
    }
    
    Write-Host "Found $($MsiFiles.Count) MSI file(s)`n" -ForegroundColor Green
} else {
    # Single file
    if ($Path -notlike "*.msi") {
        Write-Host "ERROR: File must be an MSI installer" -ForegroundColor Red
        exit
    }
    $MsiFiles = @(Get-Item $Path)
}

# Process MSI files
$Results = @()
$ProcessedCount = 0

foreach ($MsiFile in $MsiFiles) {
    $ProcessedCount++
    Write-Host "[$ProcessedCount/$($MsiFiles.Count)] Processing: " -NoNewline -ForegroundColor Cyan
    Write-Host $MsiFile.Name -ForegroundColor White
    
    $MsiInfo = Get-MSIProductCode -MsiPath $MsiFile.FullName
    
    if ($MsiInfo) {
        $Results += $MsiInfo
        
        Write-Host "  Product Name:    " -NoNewline -ForegroundColor Gray
        Write-Host $MsiInfo.ProductName -ForegroundColor Green
        Write-Host "  Version:         " -NoNewline -ForegroundColor Gray
        Write-Host $MsiInfo.ProductVersion -ForegroundColor Green
        Write-Host "  Manufacturer:    " -NoNewline -ForegroundColor Gray
        Write-Host $MsiInfo.Manufacturer -ForegroundColor Yellow
        Write-Host "  Product Code:    " -NoNewline -ForegroundColor Gray
        Write-Host $MsiInfo.ProductCode -ForegroundColor Magenta
        Write-Host "  Registry Path:   " -NoNewline -ForegroundColor Gray
        Write-Host $MsiInfo.IntuneRegistryPath -ForegroundColor Cyan
        Write-Host ""
    }
}

# Summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Summary" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Total MSI files found:      $($MsiFiles.Count)" -ForegroundColor White
Write-Host "Successfully processed:     $($Results.Count)" -ForegroundColor Green
Write-Host "Failed:                     $($MsiFiles.Count - $Results.Count)" -ForegroundColor $(if($MsiFiles.Count - $Results.Count -gt 0){"Red"}else{"Green"})
Write-Host ""

# Display results table
if ($Results.Count -gt 0) {
    Write-Host "Intune Detection Details:" -ForegroundColor Yellow
    $Results | Format-Table -AutoSize @{
        Label="Product Name"
        Expression={$_.ProductName}
        Width=30
    }, @{
        Label="Version"
        Expression={$_.ProductVersion}
        Width=10
    }, @{
        Label="Product Code"
        Expression={$_.ProductCode}
        Width=38
    }, @{
        Label="File Name"
        Expression={$_.FileName}
        Width=25
    }
}

# Export to CSV
if ($ExportCSV -or $Results.Count -gt 1) {
    $ExportPath = "C:\Temp\MSI_ProductCodes_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    # Ensure directory exists
    $ExportDir = Split-Path $ExportPath -Parent
    if (-not (Test-Path $ExportDir)) {
        New-Item -Path $ExportDir -ItemType Directory -Force | Out-Null
    }
    
    $Results | Select-Object FileName, ProductName, ProductVersion, Manufacturer, ProductCode, IntuneRegistryPath, FilePath | 
        Export-Csv -Path $ExportPath -NoTypeInformation
    
    Write-Host "✓ Results exported to: $ExportPath" -ForegroundColor Green
    
    $OpenFile = Read-Host "`nOpen CSV file? (Y/N)"
    if ($OpenFile -eq 'Y' -or $OpenFile -eq 'y') {
        Start-Process $ExportPath
    }
}

# Copy to clipboard (single result)
if ($Results.Count -eq 1) {
    $ClipboardData = @"
INTUNE MSI DETECTION RULE:
Product Name: $($Results[0].ProductName)
Product Version: $($Results[0].ProductVersion)
Product Code: $($Results[0].ProductCode)

Detection Configuration:
  Rule Type: MSI
  MSI Product Code: $($Results[0].ProductCode)

Alternative Registry Detection:
  Rule Type: Registry
  Key Path: $($Results[0].IntuneRegistryPath)
  Value Name: DisplayName
  Operator: Equals
  Value: $($Results[0].ProductName)
"@
    
    try {
        Set-Clipboard -Value $ClipboardData
        Write-Host "`n✓ Detection details copied to clipboard!" -ForegroundColor Green
    }
    catch {
        Write-Host "`nCould not copy to clipboard (may be running remotely)" -ForegroundColor Yellow
    }
}

Write-Host "`n========================================`n" -ForegroundColor Cyan