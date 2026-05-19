<#
.SYNOPSIS
    Generates and sends a privileged account status report for Domain and Enterprise Admins.

.DESCRIPTION
    This script retrieves all members of Domain Admin and Enterprise Admin groups from Active Directory,
    checks if the accounts are enabled or disabled, and sends an HTML-formatted email report via Microsoft Graph API.
    
    Prerequisites:
    - PowerShell 5.1 or higher (NuGet PackageManager required for versions < 7)
    - ActiveDirectory module installed
    - Microsoft.Graph.Authentication module installed
    - Valid Microsoft Graph credentials XML file
    - Appropriate permissions in AD and Microsoft 365

.PARAMETER Recipients
    Comma-separated list of email addresses to receive the report.
    Example: "admin@contoso.com,report@contoso.com"

.PARAMETER TenantName
    The Azure AD tenant name (e.g., "contoso.onmicrosoft.com").

.PARAMETER IDsPath
    Path to the directory containing the credentials XML file.
    File format: "{USERNAME}@{COMPUTERNAME}-Access.xml"

.PARAMETER FromAddress
    The email address used to send the report via Microsoft Graph.
    The app registration needs to be setup to allow access to the sending exchange mailbox.
    https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac#configure-applicationaccesspolicy

.PARAMETER EnableLog
    Turns on transcript logging and additional error messages for debugging

.EXAMPLE
    .\Get-PAccStatus.ps1 -Recipients "admin@contoso.com" -TenantName "contoso.onmicrosoft.com" `
        -IDsPath "C:\Credentials" -FromAddress "reports@contoso.com"

.NOTES
    - The script uses client credentials OAuth2 flow for Graph authentication
    - Microsoft Graph disconnect is performed at the end

.OUTPUTS
    - Report is sent via email only
    - Transcript logs are saved to .\Logs when -EnableLog is used
    - HTML email includes conditional formatting (red for Enabled, green for Disabled)

.CREDENTIALS
    To generate the credentials XML file:
    1. As the user that will run the script (use RunAs or PsExec), execute the following command in PowerShell:
       $creds = Get-Credential -Message "Enter your Microsoft Graph app registration credentials (Client ID as username, Client Secret as password)"
    2. Export to the credentials directory:
       $creds | Export-Clixml -Path "C:\Credentials\$env:USERNAME@$env:computername-Access.xml"
    3. The file will be encrypted using DPAPI and can only be decrypted by the same user on the same computer.

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

param (
    [Parameter(Mandatory = $true)]
    [string]$Recipients,

    [Parameter(Mandatory = $true)]
    [String]$TenantName,

    [Parameter(Mandatory = $true)]
    [String]$IDsPath,

    [Parameter(Mandatory = $true)]
    [String]$FromAddress,

    [Parameter]
    [switch]$EnableLog
)

# Set target folder for transcripts, create if it doesn't exist.
$LogDir = "$PSScriptRoot\logs"
if (!(Test-Path $LogDir)) {
    New-Item -Path $LogDir -ItemType Directory -Force
}

If ($EnableLog) {
Start-Transcript -OutputDirectory $LogDir
$TranscriptEnabled = $true
}

# Check PS Version, install Nuget PackageManager if not installed and not running PoSh 7
if ($PSVersionTable.PSVersion.Major -lt '7') {
    if (-not (Get-PackageProvider -Name 'Nuget')) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false
    }
}

# Check for required Modules, install if missing
$RequiredModules = @(
    'ActiveDirectory',
    'Microsoft.Graph.Authentication')

foreach ($Module in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $Module)) {
        Install-Module -Name $Module -Force -Confirm:$false
    }
}

Import-Module -Name $RequiredModules -ErrorAction Stop -InformationAction SilentlyContinue

function Send-Report {
    param (
        [Parameter(Mandatory = $true)]
        [object[]]$TableData,
        
        [Parameter(Mandatory = $true)]
        [array]$MailTo,

        [Parameter(Mandatory = $true)]
        [string]$MessageSubject,

        [Parameter(Mandatory = $true)]
        [string]$MessageContent,

        [Parameter(Mandatory = $true)]
        [string]$From
    )
    
    try {
        # Convert data to HTML table and add conditional formatting for Enabled column
        $htmlTable = $TableData | ConvertTo-Html -Fragment
        $htmlTable = $htmlTable -replace '<td>True</td>', '<td style="color: red;">Enabled</td>'
        $htmlTable = $htmlTable -replace '<td>False</td>', '<td style="color: green;">Disabled</td>'
        
        # Add some basic CSS styling
        $htmlBody = @"
        <style>
            table { 
                border-collapse: collapse; 
                width: auto; 
                margin: 0 auto;
            }
            th, td { 
                padding: 8px; 
                text-align: left; 
                border: 1px solid #ddd; 
                white-space: nowrap;
            }
            th { 
                background-color: #f2f2f2; 
            }
            tr:nth-child(even) { 
                background-color: #f9f9f9; 
            }
        </style>
        <p>$MessageContent</p>
        $htmlTable
"@
        
        # compose message
        $messageBody = @{
            message         = @{
                subject      = $MessageSubject
                body         = @{
                    contentType = 'HTML'
                    content     = $htmlBody
                }
                toRecipients = $MailTo
            }
            saveToSentItems = $false
        }
        
        # Send message
        $sendMailEndpoint = "https://graph.microsoft.com/v1.0/users/$From/sendMail"
        Invoke-MgGraphRequest -Method POST -Uri $sendMailEndpoint -Body $messageBody -ContentType 'application/json'
    }
    catch {
        if ($TranscriptEnabled) {
            Write-Error "Failed to send email: $_"
            Stop-Transcript
        }
        throw
    }
}

if ($TranscriptEnabled) {
    Write-Output "DEBUG: Recipients input: $Recipients"
    Write-Output "DEBUG: Recipients is Array: $($Recipients.GetType()).IsArray"
}

$ADInfo = Get-ADDomain

$ADFQDN = $ADInfo.DNSRoot

# Use well-known SIDs instead of localized group names (works in any language)
$DomainSID = $ADInfo.DomainSID.Value
$DADSid = "$DomainSID-512"
$EASid = "$DomainSID-519"

if ($TranscriptEnabled) {
    Write-Output "DEBUG: Domain SID: $DomainSID"
    Write-Output "DEBUG: Domain Admins SID: $DADSid"
    Write-Output "DEBUG: Enterprise Admins SID: $EASid"
}

$adminGroups = @($DADSid, $EASid)

$AccData = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($group in $adminGroups) {
    $members = Get-ADGroupMember -Identity $group -Recursive | Where-Object { $_.objectClass -eq 'user' }
    
    # Batch fetch all user details at once instead of individual lookups
    $members | Get-ADUser -Properties Enabled | ForEach-Object {
        $AccData.Add([PSCustomObject]@{
                Name           = $_.Name
                SamAccountName = $_.SamAccountName
                Enabled        = $_.Enabled
            })
    }
}

# Remove duplicate accounts (user may be in both Domain and Enterprise Admins)
$AccData = $AccData | Sort-Object -Property SamAccountName -Unique

# Check for cred file
$CredsFile = "$IDsPath\$env:USERNAME@$env:computername-Access.xml"
if (Test-Path -Path $CredsFile) {
    $Creds = (Import-Clixml -Path $CredsFile)
}
else {
    if ($TranscriptEnabled) {
        Write-Error "Credentials file not found in $IDsPath. Please ensure the file exists and is accessible."
        Stop-Transcript
    }
    exit 1
}

# Read client ID and secret from credentials XML
$ClientID = ($Creds).GetNetworkCredential().UserName
$ClientSecret = ($Creds).GetNetworkCredential().Password

#graph connection parameters 
$connparams = @{
    Uri         = "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token"
    Method      = 'POST'
    ErrorAction = 'Stop'
}

#build token request body
$ReqTokenBody = @{
    grant_type    = 'client_credentials'
    scope         = 'https://graph.microsoft.com/.default'
    client_id     = $ClientID
    client_secret = $ClientSecret
}

# Get access token
try {
    $TokenResponse = Invoke-RestMethod @connparams -Body $ReqTokenBody
}
catch {
    if ($TranscriptEnabled) {
        Write-Error "Failed to acquire token from $TenantName - verify tenant name, client ID, and secret are correct: $_"
        Stop-Transcript
    }
    throw
}

# Connect to graph
try {
    Connect-MgGraph -AccessToken $(ConvertTo-SecureString $TokenResponse.access_token -AsPlainText -Force) -NoWelcome
}
catch {
    if ($TranscriptEnabled) {
        Write-Error "Failed to connect to Graph: $_"
        Stop-Transcript
    }
    throw
}

$MsgSubject = "$ADFQDN - Privileged Account Status"
$MsgBody = "Status report for Domain and Enterprise Admin accounts on domain $ADFQDN."

# Split recipients string into array and create recipient objects
$recipientArray = $Recipients.Split(',', [StringSplitOptions]::RemoveEmptyEntries)

# Validate email addresses
$emailRegex = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'

foreach ($email in $recipientArray) {
    if ($email.Trim() -notmatch $emailRegex) {
        if ($TranscriptEnabled) {
            Write-Error "Invalid email format: $($email.Trim())"
            Stop-Transcript
        }
        exit 1
    }
}

if ($FromAddress -notmatch $emailRegex) {
    if ($TranscriptEnabled) {
        Write-Error "Invalid FromAddress email format: $FromAddress"
        Stop-Transcript
    }
    exit 1
}

$mgRecipients = $recipientArray | ForEach-Object {
    @{
        EmailAddress = @{
            Address = $_.Trim()
        }
    }
}

if ($TranscriptEnabled) {
    Write-Output "DEBUG: Total admin accounts found: $($AccData.Count)"
    Write-Output "DEBUG: Recipients count: $($mgRecipients.Count)"
    Write-Output "DEBUG: Recipients: $($mgRecipients | ConvertTo-Json)"
}

# Validate that we have admin accounts to report
if ($AccData.Count -eq 0) {
    if ($TranscriptEnabled) {
        Write-Warning 'No admin accounts found in Domain or Enterprise Admin groups - Report not sent'
        Stop-Transcript
    }
    exit 0
}

# Call send
Send-Report -TableData $AccData -From $FromAddress -MailTo $mgRecipients -MessageSubject $MsgSubject -MessageContent $MsgBody

# Disconnect graph
Disconnect-MgGraph | Out-Null

if ($TranscriptEnabled) {
    Write-Output "Script completed successfully - Report sent to $($mgRecipients.Count) recipient(s) with $($AccData.Count) admin account(s)"
    Stop-Transcript
}

# End of script