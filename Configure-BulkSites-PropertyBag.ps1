<#
.SYNOPSIS
    Bulk Configure SharePoint Site Collection Property Bags for Keepit Backup
    
.DESCRIPTION
    This script configures custom property bags on multiple SharePoint site collections
    to enable Keepit backup. Sites are read from a CSV file with different property values.
    
.PARAMETER CsvPath
    Path to CSV file containing site URLs and property values
    CSV Format: SiteUrl,PropertyValue
    
.PARAMETER PropertyKey
    The property bag key to set (default: "customproperty")
    
.EXAMPLE
    .\Configure-BulkSites-PropertyBag.ps1 -CsvPath "C:\SharePointSites.csv"
    
.NOTES
    Requires: PnP.PowerShell module
    Permissions: SharePoint Administrator
    
    CSV Format Example:
    SiteUrl,PropertyValue
    https://m365x69484988.sharepoint.com/sites/ContosoBrand,CASIB
    https://m365x69484988.sharepoint.com/sites/GlobalMarketing,CASIB
    https://m365x69484988.sharepoint.com/sites/Retail,CASIB
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$false)]
    [string]$PropertyKey = "customproperty"
)

# Function to check if PnP.PowerShell module is installed
function Test-PnPModule {
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Write-Host "PnP.PowerShell module is not installed." -ForegroundColor Red
        $install = Read-Host "Would you like to install it now? (Y/N)"
        if ($install -eq 'Y' -or $install -eq 'y') {
            Write-Host "Installing PnP.PowerShell module..." -ForegroundColor Yellow
            Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
            Write-Host "Module installed successfully!" -ForegroundColor Green
        } else {
            Write-Host "Cannot proceed without PnP.PowerShell module. Exiting." -ForegroundColor Red
            exit
        }
    }
}

# Function to read existing property bag value
function Get-ExistingPropertyBag {
    param(
        [string]$Key
    )
    
    try {
        $existingValue = Get-PnPPropertyBag -Key $Key -ErrorAction SilentlyContinue
        return $existingValue
    }
    catch {
        return $null
    }
}

# Function to configure property bag for a single site
function Set-SitePropertyBag {
    param(
        [string]$SiteUrl,
        [string]$PropertyKey,
        [string]$PropertyValue,
        [ref]$Results
    )
    
    $result = [PSCustomObject]@{
        SiteUrl = $SiteUrl
        PropertyKey = $PropertyKey
        PropertyValue = $PropertyValue
        PreviousValue = $null
        Status = "Processing"
        Message = ""
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    
    try {
        Write-Host "`n  Processing: $SiteUrl" -ForegroundColor Cyan
        
        # Connect to the site
        Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop
        
        # Get site title
        $web = Get-PnPWeb
        Write-Host "    Site Title: $($web.Title)" -ForegroundColor White
        
        # Check for existing property bag
        $existingValue = Get-ExistingPropertyBag -Key $PropertyKey
        $result.PreviousValue = if ($existingValue) { $existingValue } else { "None" }
        
        if ($existingValue) {
            Write-Host "    Previous value: $existingValue" -ForegroundColor Yellow
        } else {
            Write-Host "    No previous value found" -ForegroundColor White
        }
        
        # Set the property bag
        Write-Host "    Setting property bag: $PropertyKey = $PropertyValue" -ForegroundColor White
        Set-PnPPropertyBagValue -Key $PropertyKey -Value $PropertyValue
        
        # Verify
        Start-Sleep -Seconds 1
        $verifyValue = Get-ExistingPropertyBag -Key $PropertyKey
        
        if ($verifyValue -eq $PropertyValue) {
            $result.Status = "Success"
            $result.Message = "Property bag configured successfully"
            Write-Host "    ✓ Success!" -ForegroundColor Green
        } else {
            $result.Status = "Failed"
            $result.Message = "Verification failed - Expected: $PropertyValue, Got: $verifyValue"
            Write-Host "    ✗ Verification failed!" -ForegroundColor Red
        }
        
        # Disconnect
        Disconnect-PnPOnline
    }
    catch {
        $result.Status = "Error"
        $result.Message = $_.Exception.Message
        Write-Host "    ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
        
        # Attempt to disconnect
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore disconnect errors
        }
    }
    
    # Add result to collection
    $Results.Value += $result
}

# Main script execution
try {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Bulk SharePoint Property Bag Configuration" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    # Check for PnP module
    Test-PnPModule
    
    # Import the module
    Import-Module PnP.PowerShell -ErrorAction Stop
    
    # Validate CSV file exists
    if (-not (Test-Path -Path $CsvPath)) {
        Write-Host "✗ CSV file not found: $CsvPath" -ForegroundColor Red
        exit 1
    }
    
    # Import CSV
    Write-Host "Loading sites from CSV: $CsvPath" -ForegroundColor Yellow
    $sites = Import-Csv -Path $CsvPath
    
    if (-not $sites -or $sites.Count -eq 0) {
        Write-Host "✗ No sites found in CSV file" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "✓ Found $($sites.Count) site(s) to process`n" -ForegroundColor Green
    
    # Validate CSV columns
    if (-not ($sites[0].PSObject.Properties.Name -contains "SiteUrl")) {
        Write-Host "✗ CSV must contain 'SiteUrl' column" -ForegroundColor Red
        exit 1
    }
    
    if (-not ($sites[0].PSObject.Properties.Name -contains "PropertyValue")) {
        Write-Host "✗ CSV must contain 'PropertyValue' column" -ForegroundColor Red
        exit 1
    }
    
    # Display sites to be processed
    Write-Host "Sites to be configured:" -ForegroundColor Cyan
    $sites | ForEach-Object { 
        Write-Host "  - $($_.SiteUrl) -> $($_.PropertyValue)" -ForegroundColor White 
    }
    
    $confirm = Read-Host "`nProceed with configuration? (Y/N)"
    if ($confirm -ne 'Y' -and $confirm -ne 'y') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        exit
    }
    
    # Initialize results collection
    $results = @()
    
    # Process each site
    Write-Host "`nStarting bulk configuration..." -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    
    $counter = 1
    foreach ($site in $sites) {
        Write-Host "`n[$counter of $($sites.Count)]" -ForegroundColor Magenta
        Set-SitePropertyBag -SiteUrl $site.SiteUrl -PropertyKey $PropertyKey -PropertyValue $site.PropertyValue -Results ([ref]$results)
        $counter++
    }
    
    # Summary
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Configuration Summary" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    $successCount = ($results | Where-Object { $_.Status -eq "Success" }).Count
    $failedCount = ($results | Where-Object { $_.Status -ne "Success" }).Count
    
    Write-Host "Total Sites: $($results.Count)" -ForegroundColor White
    Write-Host "Successful: $successCount" -ForegroundColor Green
    Write-Host "Failed: $failedCount" -ForegroundColor $(if ($failedCount -gt 0) { "Red" } else { "White" })
    
    # Display failed sites if any
    if ($failedCount -gt 0) {
        Write-Host "`nFailed Sites:" -ForegroundColor Red
        $results | Where-Object { $_.Status -ne "Success" } | ForEach-Object {
            Write-Host "  - $($_.SiteUrl)" -ForegroundColor Yellow
            Write-Host "    Error: $($_.Message)" -ForegroundColor Red
        }
    }
    
    # Export results to CSV
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $resultsPath = Join-Path (Split-Path $CsvPath -Parent) "PropertyBag_Results_$timestamp.csv"
    $results | Export-Csv -Path $resultsPath -NoTypeInformation
    
    Write-Host "`n✓ Results exported to: $resultsPath" -ForegroundColor Green
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Bulk Configuration Complete!" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
}
catch {
    Write-Host "`n✗ CRITICAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    exit 1
}
