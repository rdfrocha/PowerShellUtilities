<#
.SYNOPSIS
    Automates PingCastle security assessment reporting with email notifications.
    Adapted from https://github.com/aidanfora/AutomatePingCastle

.DESCRIPTION
    This script automates the process of running PingCastle security assessments, generating reports,
    and sending email notifications. It includes functionality to:
    - Update PingCastle using its auto-updater
    - Run PingCastle healthcheck assessment
    - Compare current scores with previous assessment
    - Generate and compress HTML reports
    - Send email notifications with the results

.PARAMETER SmtpServer
    The SMTP server address for sending email notifications.

.PARAMETER SmtpPort
    The SMTP server port number. Defaults to 25.

.PARAMETER SenderEmail
    The email address used as the sender for notifications.
    Must be a valid email format.

.PARAMETER RecipientEmails
    An array of email addresses to receive the PingCastle reports.
    Must be valid email formats.

.EXAMPLE
    .\PC-Automate.ps1 -SmtpServer "smtp.contoso.com" -SenderEmail "sender@contoso.com" -RecipientEmails "recipient@contoso.com"

.EXAMPLE
    .\PC-Automate.ps1 -SmtpServer "smtp.contoso.com" -SmtpPort 587 -SenderEmail "sender@contoso.com" -RecipientEmails @("recipient1@contoso.com", "recipient2@contoso.com")

.NOTES
    File Name      : PC-Automate.ps1
    Prerequisite   : PingCastle must be installed in the same directory as the script
    Author         : Unknown
    Version        : 1.0

.OUTPUTS
    - HTML report file in the Reports directory
    - Email notification with compressed report
    - Updated XML file for score comparison
#>

#region parameters

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$SmtpServer,

    [int]$SmtpPort = 25,

    [Parameter(Mandatory=$true)]
    [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
    [string]$SenderEmail,

    [Parameter(Mandatory=$true)]
    [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
    [string[]]$RecipientEmails
)

#endregion

#region variables
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

# Define variables
$ApplicationName = 'PingCastle'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$PingCastlePath = Join-Path $ScriptRoot $ApplicationName
$ExecutablePath = Join-Path $PingCastlePath "$ApplicationName.exe"
$ReportFolder = Join-Path $PingCastlePath 'Reports'
$ReportFileNameFormat = 'ad_hc_{0}_{1}.html' # Domain and Date will be added dynamically

#endregion

#region Functions

# Function to extract score from XML report
function Get-PingCastleScores {
    param (
        [string]$xmlReportPath
    )

    $xmlContent = [xml](Get-Content $xmlReportPath)
    $scores = [ordered]@{
        'Global Score'           = $xmlContent.SelectSingleNode('/HealthcheckData/GlobalScore').InnerText
        'Stale Objects Score'    = $xmlContent.SelectSingleNode('/HealthcheckData/StaleObjectsScore').InnerText
        'Trust Score'            = $xmlContent.SelectSingleNode('/HealthcheckData/TrustScore').InnerText
        'Privileged Group Score' = $xmlContent.SelectSingleNode('/HealthcheckData/PrivilegiedGroupScore').InnerText
        'Anomaly Score'          = $xmlContent.SelectSingleNode('/HealthcheckData/AnomalyScore').InnerText
    }
    return $scores
}

# Function to compare two sets of scores
function Compare-Scores {
    param (
        [hashtable]$CurrentScores,
        [hashtable]$PreviousScores
    )

    foreach ($key in $CurrentScores.Keys) {
        if ($PreviousScores[$key] -ne $CurrentScores[$key]) {
            return $true
        }
    }
    return $false
}

# Function to run PingCastle and generate report
function Invoke-PingCastle {
    if (-not (Test-Path $ExecutablePath)) {
        throw "Executable not found: $ExecutablePath"
    }

    if (-not (Test-Path $ReportFolder)) {
        New-Item -Path $ReportFolder -ItemType Directory | Out-Null
    }

    try {
        Set-Location -Path $PingCastlePath
        Start-Process -FilePath $ExecutablePath -ArgumentList '--healthcheck --level Full' -WindowStyle Hidden -Wait
    }
    catch {
        throw "Failed to execute PingCastle: $_"
    }

    # Adjust the report file name and location
    $DefaultReportName = 'ad_hc_{0}.html' -f $env:USERDNSDOMAIN.ToLower()
    $DefaultReportPath = Join-Path $PingCastlePath $DefaultReportName
    $NewReportFileName = $ReportFileNameFormat -f $env:USERDNSDOMAIN.ToLower(), (Get-Date -UFormat '%d%m%y_%H%M%S')
    $NewReportPath = Join-Path $ReportFolder $NewReportFileName

    if (Test-Path $DefaultReportPath) {
        Move-Item -Path $DefaultReportPath -Destination $NewReportPath
    }
    else {
        throw "Report not generated: $DefaultReportPath"
    }

    return $NewReportPath
}

# Function to create a zip file of the report
function Compress-ZipFile {
    param (
        [string]$FilePath
    )

    $zipFilePath = "$FilePath.zip"
    if (Test-Path $zipFilePath) { Remove-Item $zipFilePath } # Remove existing zip file if any

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.IO.Compression

    $zip = [System.IO.Compression.ZipFile]::Open($zipFilePath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $fileName = [System.IO.Path]::GetFileName($FilePath)
        $fileInfo = New-Object System.IO.FileInfo($FilePath)
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $fileInfo.FullName, $fileName, [System.IO.Compression.CompressionLevel]::Optimal)
    }
    finally {
        $zip.Dispose()
    }

    return $zipFilePath
}

# Function to send email with attachment
function Send-EmailReport {
    param (
        [string]$ZipFilePath,
        [string]$To,
        [System.Collections.Specialized.OrderedDictionary]$ReportScores,
        [System.Collections.Specialized.OrderedDictionary]$OldScores,
        [bool]$IsDataChanged
    )

    # Get the current month
    $Month = (Get-Date).Month

    # Calculate the quarter
    $Quarter = [math]::Ceiling($Month / 3)

    $mailMessage = New-Object System.Net.Mail.MailMessage
    $mailMessage.IsBodyHtml = $true
    $mailMessage.From = $SenderEmail
    $mailMessage.To.Add($RecipientEmail)
    if ($IsDataChanged) {
        $mailMessage.Subject = "Q$Quarter PingCastle Report (Changes present!)"
        # build a sumarized change display
        $ScoreChanges = @()
        foreach ($key in ($OldScores.Keys + $ReportScores.Keys | Select-Object -Unique)) {
            $v1 = $OldScores[$key]
            $v2 = $ReportScores[$key]

            $ScoreChanges += [PSCustomObject]@{
                Metric           = $key
                'Previous Score' = $v1
                'Current Score'  = $v2
                Change           = if ($v1 -and $v2) { $v2 - $v1 } else { $null }
            }
        }

        $mailMessage.Body = "<p>Please find the attached updated PingCastle report. Scores changed from last quarter.</p>
        <p>- Overview:</p>
        $($ScoreChanges | ConvertTo-Html -Fragment -Property Metric,'Previous Score','Current Score',Change)"
    }
    else {
        $mailMessage.Subject = "Q$Quarter PingCastle Report (No changes)"
        $mailMessage.Body = "<p>No changes detected in the latest PingCastle report.</>
        <p>- Overview:</p>
        $($reportscores.GetEnumerator() | ForEach-Object {
    [PSCustomObject]@{
        Metric   = $_.Key   # Rename the key header
        Score =$_.Value # Rename the value header
    }
} | ConvertTo-Html -Fragment -Property Metric,Score)"
    }

    try {
        # Create and add the attachment
        $attachment = New-Object System.Net.Mail.Attachment($ZipFilePath)
        $mailMessage.Attachments.Add($attachment)

        $smtpClient = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
        
        # options not needed for internal relay
        #$smtpClient.EnableSsl = $false # Enable SSL/TLS
        #$smtpClient.Credentials = New-Object System.Net.NetworkCredential($smtpUser, $smtpPassword)

        $smtpClient.Send($mailMessage)
        Write-Host "Email sent with report to $RecipientEmail"
    }
    catch {
        Write-Error "Failed to send email: $_"
    }
    finally {
        $mailMessage.Dispose()
        if (Test-Path $ZipFilePath) { Remove-Item $ZipFilePath -Force } # Delete the zip file after sending
    }
}

#endregion

#region Main Script Execution

# Update Logic
$UpdaterName = 'PingCastleAutoUpdater'
$UpdaterPath = Join-Path $PingCastlePath "$UpdaterName.exe"

if (Test-Path $UpdaterPath) {
    try {
        Set-Location -Path $PingCastlePath
        Start-Process -FilePath $UpdaterPath -ArgumentList '--wait-for-days 30' -WindowStyle Hidden -Wait
        Write-Host 'PingCastle updated successfully.'
        Set-Location -Path $PSScriptRoot
    }
    catch {
        Write-Error "Failed to run the updater: $_"
    }
}
else {
    Write-Host "Updater not found: $UpdaterPath"
}

# Main script execution
try {
    $NewReportPath = Invoke-PingCastle
    Compress-ZipFile -FilePath $NewReportPath
    $ZipFilePath = "$NewReportPath.zip"

    # Dynamically determine the XML file name based on the domain name
    $domainName = $env:USERDNSDOMAIN.ToLower()
    $xmlFilePath = Join-Path $PingCastlePath "ad_hc_$domainName.xml"
    $oldXmlFilePath = Join-Path $PingCastlePath "ad_hc_${domainName}_old.xml"

    # Compare Scores
    $IsDataChanged = $false
    if (Test-Path $oldXmlFilePath) {
        $CurrentScores = Get-PingCastleScores -xmlReportPath $xmlFilePath
        $PreviousScores = Get-PingCastleScores -xmlReportPath $oldXmlFilePath
        $IsDataChanged = Compare-Scores -CurrentScores $CurrentScores -PreviousScores $PreviousScores
    }

    # Send email report (Add in recepient email)
    Send-EmailReport -ZipFilePath $ZipFilePath -To $RecipientEmails -IsDataChanged $IsDataChanged -ReportScores $CurrentScores -OldScores $PreviousScores

    # Rename the new XML file and delete the old one
    if (Test-Path $xmlFilePath) {
        if (Test-Path $oldXmlFilePath) {
            Remove-Item -Path $oldXmlFilePath
        }
        $newName = "ad_hc_${domainName}_old.xml"
        Rename-Item -Path $xmlFilePath -NewName $newName
    }
}
catch {
    Write-Error $_
}

#endregion