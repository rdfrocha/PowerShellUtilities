<#
.SYNOPSIS
Retrieves detailed information about shared mailboxes in Exchange Online and exports the data to a CSV file.

.DESCRIPTION
This script connects to Exchange Online and collects comprehensive details about shared mailboxes, including:
- Mailbox size and statistics
- Last logon and message times
- Manager information
- Full access permissions
- Forwarding rules and addresses

The results are exported to a CSV file with optional sorting by mailbox size.

.PARAMETER ExportPath
The file path where the CSV export will be saved. Defaults to "SharedMailboxes.csv" in the script root directory.
Type: [string]
Default: "$PSScriptRoot\SharedMailboxes.csv"

.PARAMETER SortBySize
Switch parameter to sort the exported results by mailbox size in descending order.
Type: [switch]
Default: $false

.PARAMETER Subset
Limits the number of shared mailboxes to process. When set to 0 (default), all mailboxes are processed.
Type: [int]
Default: 0

.EXAMPLE
.\Get-SharedMailboxesDetails.ps1 -ExportPath "C:\Reports\Mailboxes.csv" -SortBySize

.EXAMPLE
.\Get-SharedMailboxesDetails.ps1 -Subset 10

.NOTES
Requires:
- Exchange Online PowerShell module (EXO v2+)
- Connection to Exchange Online before running the script (Connect-ExchangeOnline)
- Appropriate permissions to query mailbox statistics and permissions

Error handling is implemented for each query operation. Failed operations return 'None|Error' or 'None' to allow processing to continue.

.OUTPUTS
[PSCustomObject] with the following properties:
- MailboxName: Display name of the mailbox
- MailboxEmailAddress: Primary SMTP address
- ForwardingAddress: Concatenated string of server forwards and inbox rule redirects
- LastLogon: Last logon timestamp or 'None'
- LastSentMessageTime: Most recent date in Sent Items or 'None'
- LastReceivedMessageTime: Most recent date in Inbox or 'None'
- Manager: Manager distinguished name or 'None'
- FullAccessFirst3: First 3 users with FullAccess permissions or 'None'
- MailboxSizeGB: Total mailbox size rounded to 2 decimal places
#>

param(
    [string]$ExportPath = "$PSScriptRoot\SharedMailboxes.csv",
    [switch]$SortBySize = $false,
    [int]$Subset = 0
)

function Get-MailboxDetails {
    param([object]$Mailbox)

    # Get mailbox statistics
    try {
        $mailboxStats = Get-EXOMailboxStatistics -Identity $Mailbox.UserPrincipalName -ErrorAction Stop
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
        $permissionsArray = Get-EXOMailboxPermission -Identity $Mailbox.UserPrincipalName |
            Where-Object { $_.User -notlike 'NT AUTHORITY\SELF' -and $_.AccessRights -contains 'FullAccess' }
        $permissions = (($permissionsArray.User | Select-Object -First 3) -join ', ')
        if (-not $permissions) { $permissions = 'None' }
    }
    catch {
        $permissions = 'None|Error'
    }

    # Get newest item in sent
    try {
        $lastsentMessage = Get-EXOMailboxFolderStatistics -Identity $Mailbox.UserPrincipalName -IncludeOldestAndNewestItems -Folderscope 'SentItems' |
            Sort-Object -Property NewestItemReceivedDate -Descending |
                Select-Object -ExpandProperty NewestItemReceivedDate -First 1
        $lastsentMessageDate = if ($null -eq $lastsentMessage) { 'None' } else { $lastsentMessage }
    }
    catch {
        $lastsentMessageDate = 'None|Error'
    }

    # Get newest item in inbox
    try {
        $lastInboxMessage = Get-EXOMailboxFolderStatistics -Identity $Mailbox.UserPrincipalName -IncludeOldestAndNewestItems -Folderscope 'Inbox' |
            Sort-Object -Property NewestItemReceivedDate -Descending |
                Select-Object -ExpandProperty NewestItemReceivedDate -First 1 
        $lastInboxMessageDate = if ($null -eq $lastInboxMessage) { 'None' } else { $lastInboxMessage }
    }
    catch {
        $lastInboxMessageDate = 'None|Error'
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
        MailboxName             = $Mailbox.DisplayName
        MailboxEmailAddress     = $Mailbox.PrimarySmtpAddress
        ForwardingAddress       = $forwardSummary
        LastLogon               = $lastLogonTime
        LastSentMessageTime     = $lastsentMessageDate
        LastReceivedMessageTime = $lastInboxMessageDate
        Manager                 = $manager
        FullAccessFirst3        = $permissions
        MailboxSizeGB           = $mailboxSizeGB
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
    $Details | Export-Csv -Path $Path -NoTypeInformation -Force -Encoding utf8BOM
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

# Apply subset filter if specified
if ($Subset -gt 0) {
    $sharedMailboxes = $sharedMailboxes | Select-Object -First $Subset
}

$i = 0
$TotalItems = $sharedMailboxes.Count
$mailboxDetails = @()

foreach ($mailbox in $sharedMailboxes) {
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

# Reset Progress preferences
$ProgressPreference = $OriginalProgPref
$PSStyle.Progress.View = $OriginalProgView

$EndTime = Get-Date
$executionTime = New-TimeSpan -Start $StartTime -End $EndTime
Write-Host ('Script Execution time: {0:hh\:mm\:ss}' -f $executionTime)