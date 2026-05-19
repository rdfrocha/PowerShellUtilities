<#
.SYNOPSIS
    Bulk removes OneDrive sites for shared mailboxes.
    Ensures the site is unlocked before deletion.
    Supports Dry-Run

.DESCRIPTION
    This script reads a CSV containing either UPN or onedriveURL columns.
    For each entry:
      - Resolves the OneDrive site URL if only UPN is provided
      - Unlocks the SPO site if locked
      - Deletes the SPO site
      - Optional: Deletes from Deleted Sites

.PARAMETER csvPath 
    Location of csv file containing UPN or OneDrive URLs to process

.PARAMETER tenant 
    Tenant name to build urls. e.g. contoso for contoso-admin.sharpoint.com

.PARAMETER purgeDeletedSites
    Force purge form Deleted Sites

.PARAMETER dryRun
    No action, test

.NOTES
    Requires: PnP.PowerShell Module
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$csvPath,

    [Parameter(Mandatory)]
    [string]$tenant,

    [switch]$purgeDeletedSites,
    
    [switch]$dryRun
)

# Connect to SharePoint Admin Center
Connect-PnPOnline -Url "https://$tenant-admin.sharepoint.com" -Interactive

# Load CSV
$items = Import-Csv $csvPath

foreach ($item in $items) {

    # Determine OneDrive URL
    $onedriveURL = $null

    if ($item.onedriveURL) {
        $onedriveURL = $item.onedriveURL
    }
    elseif ($item.UPN) {
        # Convert UPN to OneDrive URL format
        $alias = $item.UPN.Replace('@', '_').Replace('.', '_')
        $onedriveURL = "https://$tenant-my.sharepoint.com/personal/$alias"
    }
    else {
        Write-Host 'Skipping entry: no UPN or onedriveURL' -ForegroundColor Yellow
        continue
    }

    Write-Host "Processing: $onedriveURL" -ForegroundColor Cyan

    
    # Retrieve the site
    try {
        $site = Get-PnPTenantSite -Url $onedriveURL -ErrorAction Stop
    }
    catch {
        Write-Host "Site not found: $onedriveURL" -ForegroundColor Yellow
        continue
    }

    # Report lock state
    Write-Host "Current LockState: $($site.LockState)" -ForegroundColor Gray

    # Unlock if needed
    if ($site.LockState -ne 'Unlock') {
        if ($dryRun) {
            Write-Host '[Dry-Run] Would unlock site...' -ForegroundColor DarkYellow
        }
        else {
            Write-Host 'Unlocking site...' -ForegroundColor Green
            Set-PnPTenantSite -Url $onedriveURL -LockState Unlock -ErrorAction Stop
        }
    }

    # Delete the site
    if ($dryRun) {
        Write-Host '[Dry-Run] Would delete site' -ForegroundColor DarkYellow
    }
    else {
        Write-Host 'Deleting site' -ForegroundColor Red
        Remove-PnPTenantSite -Url $onedriveURL -Force -SkipRecycleBin -ErrorAction Stop
    }

    # Purge deleted sites
    if ($purgeDeletedSites) {
        if ($dryRun) {
            Write-Host '[Dry-Run] Would purge deleted site' -ForegroundColor DarkYellow
        }
        else {
            Write-Host 'Purging deleted site' -ForegroundColor DarkRed
            Remove-PnPDeletedSite -Identity $onedriveURL -Confirm:$false -ErrorAction SilentlyContinue
        }
    }

    if ($dryRun) {
        Write-Host "Dry-Run complete for: $onedriveURL" -ForegroundColor Yellow
    }
    else {
        Write-Host "Completed: $onedriveURL" -ForegroundColor Green
    }
}

Write-Host "Finished processing. Was it a Dry-Run? $dryRun" -ForegroundColor Cyan