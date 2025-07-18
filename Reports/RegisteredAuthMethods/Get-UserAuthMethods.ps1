<#
.SYNOPSIS
Generates a report of user authentication methods registered in Microsoft Graph and exports it to a CSV file.

.DESCRIPTION
This script retrieves Multi-Factor Authentication (MFA) registration details for users from Microsoft Graph using the `Get-MgBetaReportAuthenticationMethodUserRegistrationDetail` cmdlet. 
The data is processed and exported to a CSV file for further analysis. The script ensures that the Microsoft.Graph module is installed and that the user is authenticated to Microsoft Graph with the required permissions.

.PARAMETER ExportPath
Specifies the file path where the CSV report will be saved. Defaults to "MFAReport.csv" in the script's directory.

.PARAMETER Overwrite
A switch parameter that, when specified, allows overwriting the existing CSV file at the specified ExportPath.

.EXAMPLE
.\Get-MgBetaReportAuthenticationMethodUser.ps1
Generates the MFA report and saves it to "MFAReport.csv" in the script's directory.

.EXAMPLE
.\Get-MgBetaReportAuthenticationMethodUser.ps1 -ExportPath "C:\Reports\MFAReport.csv"
Generates the MFA report and saves it to "C:\Reports\MFAReport.csv".

.EXAMPLE
.\Get-MgBetaReportAuthenticationMethodUser.ps1 -ExportPath "C:\Reports\MFAReport.csv" -Overwrite
Generates the MFA report and overwrites the existing file at "C:\Reports\MFAReport.csv".

.NOTES
- Requires the Microsoft.Graph PowerShell module to be installed.
- Requires the "Reports.Read.All" permission scope in Microsoft Graph.
- The script processes user data and exports it in a sorted manner by UserPrincipalName.

#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$ExportPath = "$PSScriptRoot\MFAReport.csv",
    [switch]$Overwrite
)

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Error "Microsoft.Graph module is not installed. Please install it first."
    exit 1
}

# no auth automation, user driven
if (-not (Get-MgContext)) {
    Write-Verbose "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "Reports.Read.All"
}

try {
    Write-Verbose "Retrieving MFA registration details..."
    $data = Get-MgBetaReportAuthenticationMethodUserRegistrationDetail -Property * -All

    if (-not $data) {
        Write-Warning "No MFA registration data found."
        exit 1
    }

    $total = $data.Count
    $progress = 0
    $report = foreach ($user in $data) {
        $progress++
        Write-Progress -Activity "Processing users" -Status "$progress of $total" -PercentComplete (($progress / $total) * 100)
        [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            UserPreferredMethodForSecondaryAuthentication = $user.UserPreferredMethodForSecondaryAuthentication
            MethodsRegistered = [string]::join(",", ($user.MethodsRegistered))
        }
    }

    $report = $report | Sort-Object UserPrincipalName

    $csvParams = @{
        Path = $ExportPath
        Force = $true
        NoTypeInformation = $true
    }
    if (-not $Overwrite) { $csvParams.NoClobber = $true }

    $report | Export-Csv @csvParams
    Write-Host "Exported MFA report to $ExportPath"
} catch {
    Write-Error "Failed to generate MFA report: $($_.Exception.Message)"
    exit 1
}