<#
.SYNOPSIS
    Generates and emails a report of Privileged Identity Management (PIM) role activations.

.DESCRIPTION
    This script retrieves PIM role activation audit logs from Microsoft Graph for the last 30 days,
    processes the data, and sends a CSV report via email to specified recipients. It includes details
    such as user information, activated roles, activation duration, and status.

.NOTES
    File Name      : PIM_Activations.ps1
    Prerequisites  : 
    - Microsoft Graph PowerShell SDK
    - Azure AD App Registration with appropriate permissions
    - Stored credentials file with Client ID and Secret
    Required Permissions:
    - AuditLog.Read.All
    - Mail.Send

.INPUTS
    None. This script does not accept pipeline input.

.OUTPUTS
    Sends email(s) with CSV attachment containing PIM activation details.

.PARAMETER Recipients
    Email addresses to send the report to.

.PARAMETER StartDate
    Start date for the report (default: 31 days ago).

.PARAMETER EndDate
    End date for the report (default: today).

.PARAMETER IDsPath
    Path to credentials folder.

.PARAMETER FromAddress
    Sender email address.

.PARAMETER Silent
    Suppress console output except for errors.

.EXAMPLE
    .\PIM_Activations.ps1 -Recipients 'user1@domain.com','user2@domain.com' -Silent

.FUNCTIONALITY
    - Authenticates to Microsoft Graph using stored app credentials
    - Retrieves PIM activation audit logs for the past 31 days
    - Processes and formats activation data
    - Generates and sends CSV report via email to configured recipients
    - Handles errors with appropriate messaging
    - Cleans up temporary files and disconnects from Graph

.LICENSE

    MIT License

    Copyright (c) 2025 Rafael da Rocha

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

.DISCLAIMER
    This script is provided "as-is" without any warranty or guarantee of its functionality or suitability for any purpose. 
    Use at your own risk. The author is not responsible for any issues, damages, or losses that may arise from using this script.
    Ensure you test the script in a controlled environment before deploying it in production.
#>

#Start-Transcript -OutputDirectory $PSScriptRoot #Debug

param (
    [Parameter(Mandatory = $true)]    
    [string[]]$Recipients = @(), # "user1@companydomain.com","User2@companydomain.com"
    [datetime]$StartDate = (Get-Date).AddDays(-31),
    [datetime]$EndDate = (Get-Date),
    [Parameter(Mandatory = $true)]
    [string]$IDsPath, # Folder where stored credentials file is saved
    #sender account - Exchange app restrictions might prevent using accounts not added to policy target
    [Parameter(Mandatory = $true)]
    [string]$FromAddress, # 'Reports@companydomain.com'
    [Parameter(Mandatory = $true)]
    [string]$TenantName, # 'companytenant.onmicrosoft.com'
    [switch]$Silent # Supress console output
)

function Write-Log {
    param([string]$Message)
    if (-not $Silent) { 
        Write-Host $Message 
    }
    else {
        Add-Content -Path "$PSScriptRoot\PIMActivations.log" -Value "$(Get-Date -Format o): $Message"
    }
}

function Get-GraphToken {
    param(
        [PSCredential]$Creds,
        $TenantName
    )
    $connparams = @{
        Uri         = "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token"
        Method      = 'POST'
        ErrorAction = 'Stop'
    }
    $ReqTokenBody = @{
        Grant_Type    = 'client_credentials'
        Scope         = 'https://graph.microsoft.com/.default'
        client_Id     = $Creds.GetNetworkCredential().UserName
        Client_Secret = $Creds.GetNetworkCredential().Password
    }
    try {
        Invoke-RestMethod @connparams -Body $ReqTokenBody
    }
    catch {
        Write-Error "Failed to acquire token: $_"
        exit 1
    }
}

function New-GraphConnection {
    param([string]$AccessToken)
    try {
        # get command parameters to handle different Graph Version requirements for AccessToken
        $CheckTokenType = (Get-Command Connect-MgGraph).Parameters['AccessToken'] 
        if ($CheckTokenType.ParameterType -eq [securestring]) {
            Connect-MgGraph -AccessToken $(ConvertTo-SecureString $AccessToken -AsPlainText -Force) -NoWelcome
        }
        else {
            Connect-MgGraph -AccessToken $AccessToken -NoWelcome
        }
    }
    catch {
        Write-Error "Failed to connect to Graph: $_"
        exit 1
    }
}

function Get-PIMActivations {
    param($StartDate, $EndDate)
    $params = @{
        Filter = @"
activityDateTime ge $($StartDate.ToString('yyyy-MM-ddTHH:mm:ssZ')) and 
activityDateTime le $($EndDate.ToString('yyyy-MM-ddTHH:mm:ssZ')) and 
category eq 'RoleManagement' and 
loggedByService eq 'PIM' and 
ActivityDisplayName eq 'Add member to role completed (PIM activation)'
"@
    }
    try {
        Get-MgAuditLogDirectoryAudit @params -All
    }
    catch {
        Write-Error "Error retrieving audit logs: $_"
        exit 1
    }
}

function Format-ActivationData {
    param($Activations)
    try {
        $Activations | Select-Object `
        @{Name = 'UserDisplayName'; Expression = { $_.InitiatedBy.User.DisplayName } },
        @{Name = 'UserPrincipalName'; Expression = { $_.InitiatedBy.User.UserPrincipalName } },
        @{Name = 'Role'; Expression = { $_.TargetResources[0].DisplayName } },
        @{Name = 'ActivatedDateTime'; Expression = { [datetime]::parse($($_.AdditionalDetails | Where-Object { $_.Key -eq 'StartTime' } | Select-Object -ExpandProperty Value)) } },
        @{Name = 'Duration'; Expression = {
                $ActivationTime = ($_.AdditionalDetails | Where-Object { $_.Key -eq 'StartTime' } | Select-Object -ExpandProperty Value)
                $expirationTime = ($_.AdditionalDetails | Where-Object { $_.Key -eq 'ExpirationTime' } | Select-Object -ExpandProperty Value)
                if ($ActivationTime -and $expirationTime) {
                    $duration = [datetime]::parse($expirationTime) - [datetime]::parse($ActivationTime)
                    "$([math]::Round($duration.TotalHours, 2)) hours"
                }
                else { 'Unknown' }
            }
        },
        @{Name = 'Status'; Expression = { if ($_.Result -eq 'success') { 'Activated' } else { 'Failed' } } },
        @{Name = 'Reason'; Expression = { ($_.ResultReason -as [string]).Trim()} }
    }
    catch {
        Write-Error "Error parsing data for CSV: $_"
        exit 1
    }
}

function Send-PIMReport {
    param (
        [object[]]$CsvData,
        [string]$MailTo,
        [string]$FromAddress,
        [datetime]$StartDate,
        [datetime]$EndDate
    )
    $periodDays = ($EndDate - $StartDate).Days
    $timestamp = Get-Date -Format 'yyyyMMdd'
    $csvFileName = "PIM_Activations_Report_${timestamp}.csv"
    $tempFile = Join-Path -Path $env:TEMP -ChildPath $csvFileName
    try {
        $CsvData | Export-Csv -Path $tempFile -NoTypeInformation -Encoding UTF8
        $contentBytes = [System.IO.File]::ReadAllBytes($tempFile)
        $contentBase64 = [System.Convert]::ToBase64String($contentBytes)
        $messageBody = @{
            message         = @{
                subject      = "PIM Activations Report - Last $periodDays Days"
                body         = @{
                    contentType = 'Text'
                    content     = "Please find attached the PIM activations report for the period $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))."
                }
                toRecipients = @(
                    @{ emailAddress = @{ address = "$MailTo" } }
                )
                attachments  = @(
                    @{
                        '@odata.type' = '#microsoft.graph.fileAttachment'
                        name          = $csvFileName
                        contentType   = 'text/csv'
                        contentBytes  = $contentBase64
                    }
                )
            }
            saveToSentItems = $false
        }
        $sendMailEndpoint = "https://graph.microsoft.com/v1.0/users/$FromAddress/sendMail"
        Invoke-MgGraphRequest -Method POST -Uri $sendMailEndpoint -Body $messageBody -ContentType 'application/json'
        Write-Log "Report sent to $MailTo"
        return @{Recipient=$MailTo;Status='Sent';File=$csvFileName}
    }
    catch {
        Write-Error "Failed to send email to $($MailTo): $_"
        return @{Recipient=$MailTo;Status='Error';File=$csvFileName}
    }
    finally {
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
    }
}

# Main script logic

Write-Log 'Starting PIM Activations report...'

# Check module
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Reports)) {
    Write-Error 'Microsoft.Graph.Reports module is not installed.'
    exit 1
}
Import-Module Microsoft.Graph.Reports

<# 
###  *-Access.xml contains the ClientID/Secret, decryptable by the User/Host pair it was generated for with the following two lines
### $Credentials = Get-Credential -Message 'Provide AppID as User and Secret as Password'
### $Credentials | Export-Clixml -Path "$PSScriptRoot\$env:USERNAME@$env:COMPUTERNAME-Access.xml" -Confirm:$false
#>
$CredsFile = "$IDsPath\$env:USERNAME@$env:computername-Access.xml"
if (-not (Test-Path -Path $CredsFile)) {
    Write-Error "Credentials file not found at path: $CredsFile. Please ensure the file exists and is accessible."
    exit 1
}
$Creds = Import-Clixml -Path $CredsFile

try {
    $TokenResponse = Get-GraphToken -Creds $Creds -TenantName $TenantName
    New-GraphConnection -AccessToken $($TokenResponse.access_token)

    Write-Log "Retrieving PIM activations from $StartDate to $EndDate..."
    $allActivations = Get-PIMActivations -StartDate $StartDate -EndDate $EndDate
    $csvData = Format-ActivationData -Activations $allActivations

    $summary = @()
    foreach ($recipient in $Recipients) {
        Write-Log "Sending report to $recipient..."
        $result = Send-PIMReport -CsvData $csvData -MailTo $recipient -FromAddress $FromAddress -StartDate $StartDate -EndDate $EndDate
        $summary += $result
    }
    Disconnect-Graph | Out-Null
    Write-Log 'Report sent and session disconnected.'
    Write-Log "Summary:`n$(($summary | ConvertTo-Json -Depth 3))"
}
catch {
    Write-Error "Script failed: $_"
    exit 1
}