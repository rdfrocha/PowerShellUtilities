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

    #Write-Host "Processing $($Mailbox.DisplayName)"

    # Get mailbox statistics
    try {
        $mailboxStats = Get-MailboxStatistics -Identity $Mailbox.UserPrincipalName -ErrorAction Stop
    }
    catch {
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
    }
    catch {
        $manager = 'None'
    }

    # Get mailbox permissions (first 3 users with FullAccess)
    try {
        $permissionsArray = Get-MailboxPermission -Identity $Mailbox.UserPrincipalName |
            Where-Object { $_.User -notlike 'NT AUTHORITY\SELF' -and $_.AccessRights -contains 'FullAccess' } |
                Select-Object -ExpandProperty User -First 3
        $permissions = ($permissionsArray -join ', ')
        if (-not $permissions) { $permissions = 'None' }
    }
    catch {
        $permissions = 'None|Error'
    }

    # Get last message time
    try {
        $lastMessage = Get-MailboxFolderStatistics -Identity $Mailbox.UserPrincipalName -IncludeOldestAndNewestItems |
            Sort-Object -Property Date -Descending |
                Select-Object -ExpandProperty Date -First 1
        $lastMessageDate = if ($null -eq $lastMessage) { 'None' } else { $lastMessage }
    }
    catch {
        $lastMessageDate = 'None|Error'
    }

    # Get forwards on the mailbox
    try {
        $server = $mailbox.ForwardingSMTPAddress
        if ($server) { $server = [string]$server }  # coerce to string safely

        $ruleRecipients = @()

        $rules = Get-InboxRule -Mailbox $Mailbox -ErrorAction SilentlyContinue |
            Where-Object { $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo }

        foreach ($r in $rules) {
            foreach ($t in @('ForwardTo', 'RedirectTo', 'ForwardAsAttachmentTo')) {
                $recips = $r.$t
                if ($recips) {
                    foreach ($recip in $recips) {
                        # Normalize recipient display to something useful (SMTP if available)
                        $asText = if ($recip -is [string]) {
                            $recip
                        }
                        elseif ($recip.PSObject.Properties.Match('Address').Count) {
                            $recip.Address
                        }
                        elseif ($recip.PSObject.Properties.Match('PrimarySmtpAddress').Count) {
                            $recip.PrimarySmtpAddress
                        }
                        elseif ($recip.PSObject.Properties.Match('Name').Count) {
                            $recip.Name
                        }
                        else {
                            [string]$recip
                        }
                        $ruleRecipients += "$t :`t$asText"
                    }
                }
            }
        }

        $forwardSummary =
        if (-not $server -and -not $ruleRecipients) { 'None' }
        elseif ($server -and -not $ruleRecipients) { "Server:`t$server" }
        elseif (-not $server -and $ruleRecipients) { ($ruleRecipients -join '; ') }
        else { "Server:`t$server; " + ($ruleRecipients -join '; ') }

    }
    catch {
        $forwardSummary = 'None|Error'
    }

    [PSCustomObject]@{
        MailboxName         = $Mailbox.DisplayName
        MailboxEmailAddress = $Mailbox.PrimarySmtpAddress
        ForwardingAddress   = $forwardSummary
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
Clear-Host
$StartTime = Get-Date
# Connect-ExchangeOnline

# Get current Progress preferences, so we can reset them at the end
$OriginalProgView = $PSStyle.Progress.View
$OriginalProgPref = $ProgressPreference

# Set Progress preferences for this session
$ProgressPreference = 'Continue'
$PSStyle.Progress.View = 'Classic'

Write-Host 'Getting shared mailboxes...'
$sharedMailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox -Properties DisplayName, UserPrincipalName, PrimarySmtpAddress, ForwardingSMTPAddress

$i = 0
$TotalItems = $sharedMailboxes.Count
$mailboxDetails = @()

foreach ($mailbox in ($sharedMailboxes | Select-Object -First 50)) {
    # Calculate elapsed time
    $ElapsedTime = ((Get-Date) - $StartTime).TotalSeconds
    $StillToDo = $TotalItems - $i
    $avgItemTime = $ElapsedTime / $i

    $i++
    if ($i -eq 1 ) {
        Write-Progress -Activity 'Processing shared mailboxes' -Status "$($Mailbox.DisplayName) - Mailbox $i of $TotalItems" -PercentComplete (($i / $TotalItems) * 100)
    }
    else {
        Write-Progress -Activity 'Processing shared mailboxes' -Status "$($Mailbox.DisplayName) - Mailbox $i of $TotalItems" -PercentComplete (($i / $TotalItems) * 100) -SecondsRemaining ($StillToDo * $avgItemTime)
    }

    $details = Get-MailboxDetails -Mailbox $mailbox
    if ($details) {
        $mailboxDetails += $details
    }
}

# Clear the progress bar when done
Write-Progress -Activity 'Processing shared mailboxes' -Completed

Export-MailboxDetails -Details $mailboxDetails -Path $ExportPath -Sort:($SortBySize.IsPresent)
Write-Host "Exported to $ExportPath"

#Disconnect-ExchangeOnline -Confirm:$false

$EndTime = Get-Date
$executionTime = New-TimeSpan -Start $StartTime -End $EndTime
Write-Host ('Script Execution time: {0:hh\:mm\:ss}' -f $executionTime)

# Reset Progress preferences
$ProgressPreference = $OriginalProgPref
$PSStyle.Progress.View = $OriginalProgView