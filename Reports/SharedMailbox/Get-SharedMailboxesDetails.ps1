<#
.SYNOPSIS
Exports details of shared mailboxes in an Exchange Online environment to a CSV file.

.DESCRIPTION
This script connects to Exchange Online, retrieves all shared mailboxes, and gathers detailed information about each mailbox, including:
- Display name
- Primary SMTP address
- Last logon time
- Last message time
- Manager
- Users with FullAccess permissions (up to 3)
- Mailbox size in GB

The collected data is exported to a CSV file, with an option to sort the mailboxes by size.

.PARAMETER ExportPath
Specifies the file path where the CSV file will be saved. Defaults to "SharedMailboxes.csv" in the script's directory.

.PARAMETER SortBySize
If specified, sorts the mailboxes by size in descending order before exporting to the CSV file.

.FUNCTION Get-MailboxDetails
Retrieves detailed information about a single mailbox, including:
- Display name
- Primary SMTP address
- Last logon time
- Last message time
- Manager
- Users with FullAccess permissions (up to 3)
- Mailbox size in GB

.FUNCTION Export-MailboxDetails
Exports the collected mailbox details to a CSV file. Optionally sorts the data by mailbox size.

.EXAMPLE
.\SharedMailboxes.ps1 -ExportPath "C:\Exports\SharedMailboxes.csv" -SortBySize

Connects to Exchange Online, retrieves shared mailboxes, gathers their details, sorts them by size, and exports the data to "C:\Exports\SharedMailboxes.csv".

.NOTES
- Requires the Exchange Online PowerShell module.
- Ensure you have the necessary permissions to access mailbox details.
- Uncomment the Connect-ExchangeOnline and Disconnect-ExchangeOnline lines to enable connection management.
- For automation scenarios, an Auth function can be added.

#>

param(
    [string]$ExportPath = "$PSScriptRoot\SharedMailboxes.csv",
    [switch]$SortBySize = $false
)

function Get-MailboxDetails {
    param([object]$Mailbox)

    Write-Host "Processing $($Mailbox.DisplayName)"

    # Get mailbox statistics
    try {
        $mailboxStats = Get-MailboxStatistics -Identity $Mailbox.UserPrincipalName -ErrorAction Stop
    } catch {
        Write-Host "Error getting mailbox statistics for $($Mailbox.DisplayName)"
        return $null
    }

    # Calculate mailbox size in GB
    $sizeBytes = 0
    if ($mailboxStats.TotalItemSize.Value -and $mailboxStats.TotalItemSize.Value.ToBytes) {
        $sizeBytes = [double]$mailboxStats.TotalItemSize.Value.ToBytes()
    }
    $mailboxSizeGB = [math]::Round($sizeBytes / 1GB, 2)

    # Get last logon time
    $lastLogonTime = if ($null -eq $mailboxStats.LastLogonTime) { 'None' } else { $mailboxStats.LastLogonTime }

    # Get mailbox manager
    try {
        $user = Get-User -Identity $Mailbox.UserPrincipalName -ErrorAction Stop
        $manager = if ($null -eq $user.Manager) { 'None' } else { $user.Manager }
    } catch {
        $manager = 'None'
    }

    # Get mailbox permissions (first 3 users with FullAccess)
    try {
        $permissionsArray = Get-MailboxPermission -Identity $Mailbox.UserPrincipalName |
            Where-Object { $_.User -notlike 'NT AUTHORITY\SELF' -and $_.AccessRights -contains 'FullAccess' } |
            Select-Object -ExpandProperty User -First 3
        $permissions = ($permissionsArray -join ', ')
        if (-not $permissions) { $permissions = 'None' }
    } catch {
        $permissions = 'None|Error'
    }

    # Get last message time
    try {
        $lastMessage = Get-MailboxFolderStatistics -Identity $Mailbox.UserPrincipalName -IncludeOldestAndNewestItems |
            Sort-Object -Property Date -Descending |
            Select-Object -ExpandProperty Date -First 1
        $lastMessageDate = if ($null -eq $lastMessage) { 'None' } else { $lastMessage }
    } catch {
        $lastMessageDate = 'None'
    }

    # Return mailbox details object
    [PSCustomObject]@{
        MailboxName         = $Mailbox.DisplayName
        MailboxEmailAddress = $Mailbox.PrimarySmtpAddress
        LastLogon           = $lastLogonTime
        LastMessageTime     = $lastMessageDate
        Manager             = $manager
        FullAccessFirst3    = $permissions
        MailboxSizeGB       = $mailboxSizeGB
    }
}

function Export-MailboxDetails {
    param(
        [array]$Details,
        [string]$Path,
        [switch]$Sort
    )
    if ($Sort) {
        $Details = $Details | Sort-Object MailboxSizeGB -Descending
    }
    $Details | Export-Csv -Path $Path -NoTypeInformation -Force
}

# Main script logic
$StartTime = Get-Date
# Connect-ExchangeOnline

Write-Host "Getting shared mailboxes..."
$sharedMailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox

Write-Host "Total Mailboxes: $($sharedMailboxes.Count)"
$mailboxDetails = @()
$i = 0

foreach ($mailbox in $sharedMailboxes) {
    $details = Get-MailboxDetails -Mailbox $mailbox
    if ($details) {
        $mailboxDetails += $details
    }
    $i++
    Write-Host "Done $i of $($sharedMailboxes.Count)"
}

Export-MailboxDetails -Details $mailboxDetails -Path $ExportPath -Sort:($SortBySize.IsPresent)
Write-Host "Exported to $ExportPath"

#Disconnect-ExchangeOnline -Confirm:$false

$EndTime = Get-Date
$executionTime = New-TimeSpan -Start $StartTime -End $EndTime
Write-Host ("Script Execution time: {0:hh\:mm\:ss}" -f $executionTime)