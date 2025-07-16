<#
.SYNOPSIS
    Manages the installation, update, and configuration of the Zabbix Agent on Windows systems.
    Optimized for running as a scheduled task.

.DESCRIPTION
    This script automates the process of installing, updating, and configuring the Zabbix Agent 
    on Windows machines. It is optimized for running as a scheduled task with minimal interaction.
    Features include:
    - Fully automated installation and updates
    - Configuration management
    - Comprehensive logging for headless operation
    - Cleanup of temporary files
    - Proper exit codes for task scheduler monitoring
    - Silent operation with no user interaction required
    - Support for both Zabbix Agent and Zabbix Agent 2

.PARAMETER ZabbixServer
    The hostname or IP address of the target Zabbix server to which the agent should report.

.PARAMETER ZabbixRepo
    URL to the Zabbix repository. Defaults to version 7.0.

.PARAMETER Architecture
    Target architecture for the Zabbix agent. Options:
    - windows-amd64-openssl: 64-bit version
    - windows-i386-openssl: 32-bit version
    - auto: Automatically detect based on operating system (default)

.PARAMETER DefaultAgentType
    Type of agent to install when no existing agent is found. Options: Agent, Agent 2. Defaults to Agent.

.PARAMETER LogFilePath
    Path where log files should be written. Defaults to ProgramData\ZabbixAgent directory.

.PARAMETER BackupRetentionDays
    Number of days to retain configuration backups. Defaults to 30.

.PARAMETER Scheduled
    Indicates the script is running as a scheduled task. Suppresses console output.

.EXAMPLE
    # Run as scheduled task (use this in Task Scheduler)
    powershell.exe -ExecutionPolicy Bypass -File "<path_to_script>\ManageZabbixAgent.ps1" -ZabbixServer "zabbix.example.com" -Scheduled

.EXAMPLE
    # Install Zabbix Agent 2 specifically
    powershell.exe -ExecutionPolicy Bypass -File "<path_to_script>\ManageZabbixAgent.ps1" -ZabbixServer "zabbix.example.com" -DefaultAgentType "Agent 2"

.NOTES
    - Requires PowerShell 5.1 or later
    - Must be run as an administrator or SYSTEM account
    - Exit codes:
      0 - Success
      1 - General Error
      2 - Installer Download Failed
      3 - Multiple Services Detected
      4 - Configuration Error
      5 - Missing Administrator Privileges

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

#region parameters

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$ZabbixServer,
    
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ZabbixRepo = 'https://cdn.zabbix.com/zabbix/binaries/stable/7.0/',
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('windows-amd64-openssl', 'windows-i386-openssl', 'auto')]
    [string]$Architecture = 'auto',
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('Agent', 'Agent 2')]
    [string]$DefaultAgentType = 'Agent',
    
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogFilePath = "$env:ProgramData\ZabbixAgent\logs\ManageZabbixAgent.log",

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 365)]
    [int]$BackupRetentionDays = 30,
    
    [Parameter(Mandatory = $false)]
    [switch]$Scheduled
)

#endregion

#region global settings and requirements

# Global error handling settings
$ErrorActionPreference = 'Stop'
$global:InstallerPath = "$env:TEMP\zabbix_agent_installer.msi"
$global:IsScheduledTask = $Scheduled.IsPresent
$global:StartTime = Get-Date
$global:ScriptPath = $MyInvocation.MyCommand.Path
# Load required assemblies
Add-Type -AssemblyName System.ServiceProcess

#endregion

#region Functions

function Write-LogEntry {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )
    
    # Create timestamp
    $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    
    # Format the log entry
    $LogEntry = "$Timestamp [$Level] - $Message"
    
    # Ensure directory exists
    $LogDir = Split-Path -Path $LogFilePath -Parent
    if (-not (Test-Path -Path $LogDir)) {
        try {
            New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        catch {
            # If we can't create the directory, fall back to TEMP
            $LogFilePath = "$env:TEMP\ManageZabbixAgentScript.log"
            $LogDir = Split-Path -Path $LogFilePath -Parent
        }
    }
    
    # Write to log file
    try {
        Add-Content -Path $LogFilePath -Value $LogEntry -ErrorAction Stop
    }
    catch {
        # If we can't write to the log, try creating a new log in TEMP
        $LogFilePath = "$env:TEMP\ManageZabbixAgentScript_Fallback.log"
        try {
            Add-Content -Path $LogFilePath -Value "Failed to write to primary log: $_" -ErrorAction SilentlyContinue
            Add-Content -Path $LogFilePath -Value $LogEntry -ErrorAction SilentlyContinue
        }
        catch {
            # If we still can't write, there's nothing more we can do as a scheduled task
        }
    }
    
    # Output to console if not running as scheduled task
    if (-not $global:IsScheduledTask) {
        $color = switch ($Level) {
            'INFO' { 'White' }
            'WARNING' { 'Yellow' }
            'ERROR' { 'Red' }
            'SUCCESS' { 'Green' }
            default { 'White' }
        }
        Write-Host -ForegroundColor $color "$Level`: $Message"
    }
}

function Initialize-LogFile {
    [CmdletBinding()]
    param ()

    # Always use $script:LogFilePath for consistency
    $script:LogFilePath = $LogFilePath

    # Ensure the directory exists
    $LogDir = Split-Path -Path $script:LogFilePath -Parent
    if (-not (Test-Path -Path $LogDir)) {
        try {
            New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        catch {
            # If we can't create the directory, fall back to TEMP
            $script:LogFilePath = "$env:TEMP\ManageZabbixAgentScript.log"
            $LogDir = Split-Path -Path $script:LogFilePath -Parent
        }
    }

    # Check if the log file exists and rotate if needed
    if (Test-Path -Path $script:LogFilePath) {
        try {
            $LogFile = Get-Item -Path $script:LogFilePath
            # If file is greater than 5MB, archive it
            if ($LogFile.Length -gt 5MB) {
                $Archive = "$($script:LogFilePath).$(Get-Date -Format 'yyyyMMdd_HHmmss').archive"
                Move-Item -Path $script:LogFilePath -Destination $Archive -Force

                # Create a new log file
                New-Item -Path $script:LogFilePath -ItemType File -Force | Out-Null

                # Clean up old archives (keep last 5)
                Get-ChildItem -Path $LogDir -Filter "$(Split-Path -Path $script:LogFilePath -Leaf).*.archive" |
                    Sort-Object LastWriteTime -Descending |
                        Select-Object -Skip 5 |
                            Remove-Item -Force
            }
        }
        catch {
            # If rotation fails, just try to continue with existing file
        }
    }
    else {
        # Create a new log file if it doesn't exist
        try {
            New-Item -Path $script:LogFilePath -ItemType File -Force | Out-Null
        }
        catch {
            # If we can't create the log, try TEMP
            $script:LogFilePath = "$env:TEMP\ManageZabbixAgentScript.log"
            New-Item -Path $script:LogFilePath -ItemType File -Force -ErrorAction SilentlyContinue | Out-Null
        }
    }

    # Write header information
    $Header = @"
#########################
Zabbix Agent Management Script
Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
Parameters: ZabbixServer=$ZabbixServer, Architecture=$Architecture, DefaultAgentType=$DefaultAgentType
Script Path: $global:ScriptPath
Scheduled Mode: $global:IsScheduledTask

Error Codes:
0 - Success
1 - General Error
2 - Installer Download Failed
3 - Multiple Services Detected
4 - Configuration Error
5 - Missing Administrator Privileges
#########################

"@
    Add-Content -Path $script:LogFilePath -Value $Header
}

function Register-CleanupHandler {
    [CmdletBinding()]
    param()
    
    # Set a script-level cleanup action
    $global:Cleanup = {
        param($ExitCode = 0)
        
        # Calculate runtime
        $runtime = (Get-Date) - $global:StartTime
        Write-LogEntry -Message "Script execution completed. Runtime: $($runtime.ToString('hh\:mm\:ss'))" -Level 'INFO'
        
        # Clean up installer file
        if (Test-Path -Path $global:InstallerPath) {
            Remove-Item -Path $global:InstallerPath -Force -ErrorAction SilentlyContinue
            Write-LogEntry -Message "Removed installer file: $global:InstallerPath" -Level 'INFO'
        }
        
        # Report exit status
        Write-LogEntry -Message "Exiting with code: $ExitCode" -Level $(if ($ExitCode -eq 0) { 'SUCCESS' }else { 'ERROR' })
        
        # Create status file for monitoring
        try {
            $statusDir = "$env:ProgramData\ZabbixAgent\status"
            if (-not (Test-Path $statusDir)) {
                New-Item -Path $statusDir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
            }
            
            $statusInfo = @{
                LastRun  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                ExitCode = $ExitCode
                Runtime  = $runtime.ToString('hh\:mm\:ss')
                Server   = $ZabbixServer
            }
            
            $statusInfo | ConvertTo-Json | Set-Content -Path "$statusDir\LastRunStatus.json" -Force -ErrorAction SilentlyContinue
        }
        catch {
            # Status file is non-critical
        }
    }
    
    # Register an event handler for script exit if possible
    try {
        $null = Register-EngineEvent -SourceIdentifier ([System.Management.Automation.PsEngineEvent]::Exiting) -Action {
            & $global:Cleanup
        } -ErrorAction SilentlyContinue
    }
    catch {
        Write-LogEntry -Message "Failed to register cleanup handler: $_" -Level 'WARNING'
    }
}

function Assert-Administrator {
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    
    try {
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        $isSystem = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name -eq 'NT AUTHORITY\SYSTEM'
        
        if (-not ($isAdmin -or $isSystem)) {
            Write-LogEntry -Message 'Script must be run as an Administrator or SYSTEM account.' -Level 'ERROR'
            return $false
        }
        
        Write-LogEntry -Message "Script is running with sufficient privileges: $(if($isAdmin){'Administrator'}else{'SYSTEM'})" -Level 'INFO'
        return $true
    }
    catch {
        Write-LogEntry -Message "Failed to check administrator privileges: $_" -Level 'ERROR'
        return $false
    }
}

function Get-SystemArchitecture {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    
    try {
        if ([Environment]::Is64BitOperatingSystem) {
            Write-LogEntry -Message 'Detected 64-bit operating system' -Level 'INFO'
            return 'windows-amd64-openssl'
        }
        else {
            Write-LogEntry -Message 'Detected 32-bit operating system' -Level 'INFO'
            return 'windows-i386-openssl'
        }
    }
    catch {
        Write-LogEntry -Message "Error detecting system architecture: $_. Defaulting to 64-bit." -Level 'WARNING'
        return 'windows-amd64-openssl'
    }
}

function Invoke-WebRequestWithRetry {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Url,
        
        [Parameter(Mandatory = $false)]
        [string]$OutFile = $null,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryCount = 5,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryDelaySeconds = 10
    )
    
    $attempt = 0
    $success = $false
    $result = $null
    
    while ($attempt -lt $RetryCount -and -not $success) {
        $attempt++
        try {
            if ($PSCmdlet.ShouldProcess($Url, 'Download content')) {
                if (-not [string]::IsNullOrWhiteSpace($OutFile)) {
                    # Download file mode
                    Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -ErrorAction Stop
                    $success = $true
                    Write-LogEntry -Message "Successfully downloaded $Url to $OutFile" -Level 'INFO'
                }
                else {
                    # Get content mode
                    $result = Invoke-WebRequest -Uri $Url -UseBasicParsing -ErrorAction Stop
                    $success = $true
                    Write-LogEntry -Message "Successfully fetched content from $Url" -Level 'INFO'
                }
            }
            else {
                # WhatIf mode - simulate success
                $success = $true
                Write-LogEntry -Message "WhatIf: Would download from $Url" -Level 'INFO'
                if (-not [string]::IsNullOrWhiteSpace($OutFile)) {
                    Write-LogEntry -Message "WhatIf: Would save to $OutFile" -Level 'INFO'
                }
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-LogEntry -Message "Attempt $attempt of $RetryCount failed: $errorMessage" -Level 'WARNING'
            
            if ($attempt -lt $RetryCount) {
                Write-LogEntry -Message "Retrying in $RetryDelaySeconds seconds..." -Level 'INFO'
                Start-Sleep -Seconds $RetryDelaySeconds
            }
            else {
                Write-LogEntry -Message "All $RetryCount download attempts failed for $Url" -Level 'ERROR'
                throw "Failed to download from $Url after $RetryCount attempts: $errorMessage"
            }
        }
    }
    
    return $result
}

function Get-ZabbixAgentType {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServiceName
    )
    
    if ($ServiceName -like '*Agent 2*') {
        return 'Agent 2'
    }
    else {
        return 'Agent'
    }
}

function Get-LatestVersion {
    [CmdletBinding()]
    [OutputType([Version])]
    param (
        [Parameter(Mandatory = $true)]
        [string]$RepoUrl
    )
    
    try {
        Write-LogEntry -Message 'Checking for the latest Zabbix Agent version...' -Level 'INFO'
        
        $HtmlContent = Invoke-WebRequestWithRetry -Url $RepoUrl
        
        # Extract version directories
        $VersionLinks = $HtmlContent.Links | 
            Where-Object { $_.href -match '^\d+\.\d+\.\d+\/$' } | 
                ForEach-Object { $_.href.TrimEnd('/') }
        
        if ($VersionLinks -and $VersionLinks.Count -gt 0) {
            # Convert to version objects and find highest
            $HighestVersion = ($VersionLinks | 
                    ForEach-Object { [Version]$_ } | 
                        Sort-Object -Descending)[0]
                
            Write-LogEntry -Message "Latest version found: $HighestVersion" -Level 'SUCCESS'
            return $HighestVersion
        }
        else {
            Write-LogEntry -Message "No version links found in repository. Check the URL: $RepoUrl" -Level 'ERROR'
            return $null
        }
    }
    catch {
        Write-LogEntry -Message "Error retrieving latest version: $_" -Level 'ERROR'
        return $null
    }
}

function Get-InstalledVersion {
    [CmdletBinding()]
    [OutputType([Version])]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ExecutablePath
    )
    
    try {
        Write-LogEntry -Message 'Checking installed Zabbix Agent version...' -Level 'INFO'
        
        if (Test-Path -Path $ExecutablePath) {
            $versionInfo = (Get-Item -Path $ExecutablePath).VersionInfo.ProductVersion
            
            if ([string]::IsNullOrWhiteSpace($versionInfo)) {
                Write-LogEntry -Message "Unable to extract version information from $ExecutablePath" -Level 'WARNING'
                return $null
            }
            
            Write-LogEntry -Message "Installed version: $versionInfo" -Level 'INFO'
            return [Version]$versionInfo
        }
        else {
            Write-LogEntry -Message "Executable does not exist at $ExecutablePath" -Level 'WARNING'
            return $null
        }
    }
    catch {
        Write-LogEntry -Message "Error retrieving installed version: $_" -Level 'ERROR'
        return $null
    }
}

function Test-AgentNeedsUpdate {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExecutablePath,
        
        [Parameter(Mandatory = $true)]
        [Version]$LatestVersion
    )
    
    try {
        $installedVersion = Get-InstalledVersion -ExecutablePath $ExecutablePath
        
        if ($null -eq $installedVersion) {
            Write-LogEntry -Message 'Could not determine installed version, update may be required' -Level 'WARNING'
            return $true
        }
        
        if ($installedVersion -lt $LatestVersion) {
            Write-LogEntry -Message "Update needed: Installed=$installedVersion, Latest=$LatestVersion" -Level 'INFO'
            return $true
        }
        else {
            Write-LogEntry -Message "Agent is up to date: $installedVersion" -Level 'SUCCESS'
            return $false
        }
    }
    catch {
        Write-LogEntry -Message "Error comparing versions: $_" -Level 'ERROR'
        return $false
    }
}

function Get-ZabbixInstaller {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$InstallerUrl,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$InstallerPath
    )
    
    try {
        Write-LogEntry -Message 'Downloading Zabbix Agent installer...' -Level 'INFO'
        
        if ($PSCmdlet.ShouldProcess($InstallerUrl, 'Download installer')) {
            Invoke-WebRequestWithRetry -Url $InstallerUrl -OutFile $InstallerPath
            
            # Verify file was downloaded successfully
            if (Test-Path -Path $InstallerPath) {
                $fileSize = (Get-Item -Path $InstallerPath).Length
                if ($fileSize -gt 0) {
                    Write-LogEntry -Message "Installer downloaded successfully ($([Math]::Round($fileSize/1MB, 2)) MB)" -Level 'SUCCESS'
                    return $true
                }
                else {
                    Write-LogEntry -Message 'Downloaded installer file is empty' -Level 'ERROR'
                    return $false
                }
            }
            else {
                Write-LogEntry -Message 'Installer file not found after download' -Level 'ERROR'
                return $false
            }
        }
        else {
            Write-LogEntry -Message "WhatIf: Would download installer from $InstallerUrl to $InstallerPath" -Level 'INFO'
            return $true
        }
    }
    catch {
        Write-LogEntry -Message "Failed to download installer: $_" -Level 'ERROR'
        return $false
    }
}

function Install-ZabbixAgent {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$InstallerPath,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Agent', 'Agent 2')]
        [string]$AgentType = 'Agent'
    )
    
    try {
        Write-LogEntry -Message "Installing Zabbix $AgentType..." -Level 'INFO'
        
        # MSI arguments
        $arguments = "/L*v `"$env:TEMP\zabbix_${AgentType}_install.log`" /i `"$InstallerPath`" /qn " + 
        "SERVER=`"$ZabbixServer`" SERVERACTIVE=`"$ZabbixServer`" ENABLEPATH=1"
        
        if ($PSCmdlet.ShouldProcess("msiexec.exe $arguments", "Install Zabbix $AgentType")) {
            # Start installation process
            Write-LogEntry -Message "Running command: msiexec.exe $arguments" -Level 'INFO'
            $process = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
            
            # Check if installation was successful
            if ($process.ExitCode -eq 0) {
                Write-LogEntry -Message "Installation of Zabbix $AgentType completed successfully" -Level 'SUCCESS'
                return $true
            }
            else {
                Write-LogEntry -Message "Installation of Zabbix $AgentType failed with exit code: $($process.ExitCode)" -Level 'ERROR'
                return $false
            }
        }
        else {
            Write-LogEntry -Message "WhatIf: Would install Zabbix $AgentType with arguments: $arguments" -Level 'INFO'
            return $true
        }
    }
    catch {
        Write-LogEntry -Message "Error during installation: $_" -Level 'ERROR'
        return $false
    }
    finally {
        # Always attempt to remove the installer file
        if (Test-Path -Path $InstallerPath) {
            Remove-Item -Path $InstallerPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Update-ZabbixAgent {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$InstallerPath,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ConfigFilePath,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Agent', 'Agent 2')]
        [string]$AgentType = 'Agent'
    )
    
    try {
        Write-LogEntry -Message "Updating Zabbix $AgentType..." -Level 'INFO'
        
        # MSI arguments - use existing config file
        $arguments = "/L*v `"$env:TEMP\zabbix_${AgentType}_update.log`" /i `"$InstallerPath`" /qn ENABLEPATH=1 NONMSICONFNAME=`"$ConfigFilePath`""
        
        if ($PSCmdlet.ShouldProcess("msiexec.exe $arguments", "Update Zabbix $AgentType")) {
            # Start update process
            Write-LogEntry -Message "Running command: msiexec.exe $arguments" -Level 'INFO'
            $process = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
            
            # Check if update was successful
            if ($process.ExitCode -eq 0) {
                Write-LogEntry -Message "Update of Zabbix $AgentType completed successfully" -Level 'SUCCESS'
                return $true
            }
            else {
                Write-LogEntry -Message "Update of Zabbix $AgentType failed with exit code: $($process.ExitCode)" -Level 'ERROR'
                return $false
            }
        }
        else {
            Write-LogEntry -Message "WhatIf: Would update Zabbix $AgentType with arguments: $arguments" -Level 'INFO'
            return $true
        }
    }
    catch {
        Write-LogEntry -Message "Error during update: $_" -Level 'ERROR'
        return $false
    }
    finally {
        # Always attempt to remove the installer file
        if (Test-Path -Path $InstallerPath) {
            Remove-Item -Path $InstallerPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Update-ZabbixConfiguration {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ConfigFilePath,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ServiceName,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$Replacements
    )
    
    try {
        Write-LogEntry -Message "Updating Zabbix Agent configuration at $ConfigFilePath" -Level 'INFO'
        
        if (-not (Test-Path -Path $ConfigFilePath)) {
            Write-LogEntry -Message "Configuration file not found: $ConfigFilePath" -Level 'ERROR'
            return $false
        }
        
        # Create backup
        $backupTimestamp = (Get-Date -Format 'yyyyMMdd_HHmmss')
        $backupFile = "$ConfigFilePath.backup.$backupTimestamp"
        
        if ($PSCmdlet.ShouldProcess($ConfigFilePath, 'Update configuration')) {
            # Backup the file
            Copy-Item -Path $ConfigFilePath -Destination $backupFile -Force
            Write-LogEntry -Message "Configuration backup created: $backupFile" -Level 'INFO'
            
            # Read content
            $configContent = Get-Content -Path $ConfigFilePath -Raw
            
            # Apply replacements
            foreach ($pattern in $Replacements.Keys) {
                $replacement = $Replacements[$pattern]
                $configContent = $configContent -replace $pattern, $replacement
            }
            
            # Write updated content
            Set-Content -Path $ConfigFilePath -Value $configContent -Force
            Write-LogEntry -Message 'Configuration updated successfully' -Level 'SUCCESS'
            
            # Clean up old backups
            Get-ChildItem -Path (Split-Path -Path $ConfigFilePath -Parent) |
                Where-Object { $_.Name -like "$(Split-Path -Path $ConfigFilePath -Leaf).backup.*" -and 
                    $_.LastWriteTime -lt (Get-Date).AddDays(-$BackupRetentionDays) } |
                    ForEach-Object {
                        Remove-Item -Path $_.FullName -Force
                        Write-LogEntry -Message "Removed old backup: $($_.Name)" -Level 'INFO'
                    }
            
            # Restart the service
            if (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue) {
                # Try to restart service with retry logic
                $serviceRestarted = $false
                $retryCount = 0
                $maxRetries = 3
                
                while (-not $serviceRestarted -and $retryCount -lt $maxRetries) {
                    $retryCount++
                    try {
                        Write-LogEntry -Message "Attempting to restart service $ServiceName (attempt $retryCount of $maxRetries)..." -Level 'INFO'
                        Restart-Service -Name $ServiceName -Force -ErrorAction Stop
                        $serviceRestarted = $true
                        Write-LogEntry -Message "Service $ServiceName restarted successfully" -Level 'SUCCESS'
                    }
                    catch {
                        Write-LogEntry -Message "Failed to restart service on attempt $($retryCount): $_" -Level 'WARNING'
                        if ($retryCount -lt $maxRetries) {
                            Write-LogEntry -Message 'Waiting 10 seconds before retry...' -Level 'INFO'
                            Start-Sleep -Seconds 10
                        }
                        else {
                            Write-LogEntry -Message 'All restart attempts failed. Configuration was updated but service could not be restarted.' -Level 'ERROR'
                        }
                    }
                }
            }
            else {
                Write-LogEntry -Message "Service $ServiceName not found, skipping restart" -Level 'WARNING'
            }
            
            return $true
        }
        else {
            Write-LogEntry -Message "WhatIf: Would update configuration at $ConfigFilePath" -Level 'INFO'
            return $true
        }
    }
    catch {
        Write-LogEntry -Message "Error updating configuration: $_" -Level 'ERROR'
        return $false
    }
}

function Get-ZabbixService {
    [CmdletBinding()]
    [OutputType([System.ServiceProcess.ServiceController[]])]
    param()
    
    try {
        Write-LogEntry -Message 'Checking for Zabbix Agent services...' -Level 'INFO'
        $services = Get-Service -ErrorAction SilentlyContinue | 
            Where-Object { $_.Name -like '*Zabbix Agent*' }
        
        if ($null -eq $services -or $services.Count -eq 0) {
            Write-LogEntry -Message 'No Zabbix Agent services found' -Level 'INFO'
            return @()
        }
        elseif ($services.Count -eq 1) {
            Write-LogEntry -Message "Found Zabbix Agent service: $($services.Name)" -Level 'INFO'
            return $services
        }
        else {
            Write-LogEntry -Message "Multiple Zabbix Agent services detected: $($services.Name -join ', ')" -Level 'WARNING'
            return $services
        }
    }
    catch {
        Write-LogEntry -Message "Error checking for Zabbix services: $_" -Level 'ERROR'
        return @()
    }
}

function Get-ZabbixServiceDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.ServiceProcess.ServiceController]$Service
    )
    
    try {
        $serviceDetails = Get-CimInstance -ClassName Win32_Service -Filter "Name = '$($Service.Name)'"
        
        # Get service binary path
        $serviceDetails = Get-WmiObject -Class Win32_Service -Filter "Name = '$($Service.Name)'"
        
        if ($null -eq $serviceDetails) {
            Write-LogEntry -Message "Failed to get WMI details for service $($Service.Name)" -Level 'ERROR'
            return $null
        }
        
        $binaryPathName = $serviceDetails.PathName
        $agentType = Get-ZabbixAgentType -ServiceName $Service.Name
        Write-LogEntry -Message "Detected agent type: $agentType" -Level 'INFO'
        
        # Extract executable and config paths - pattern to handle both Agent and Agent 2 formats
        if ($binaryPathName -match '^"([^"]+)"\s+(?:--config|-c)\s+"([^"]+)"') {
            $executablePath = $Matches[1]
            $configFilePath = $Matches[2]
            
            Write-LogEntry -Message "Service: $($Service.Name)" -Level 'INFO'
            Write-LogEntry -Message "Executable: $executablePath" -Level 'INFO'
            Write-LogEntry -Message "Config file: $configFilePath" -Level 'INFO'
            
            return @{
                Service        = $Service
                ExecutablePath = $executablePath
                ConfigFilePath = $configFilePath
                BinaryPathName = $binaryPathName
                AgentType      = $agentType
            }
        }
        else {
            Write-LogEntry -Message "Invalid binary path format for service $($Service.Name): $binaryPathName" -Level 'WARNING'
            return $null
        }
    }
    catch {
        Write-LogEntry -Message "Error getting service details: $_" -Level 'ERROR'
        return $null
    }
}

function Get-ZabbixInstallerUrl {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory = $true)]
        [string]$RepoUrl,
        
        [Parameter(Mandatory = $true)]
        [Version]$Version,
        
        [Parameter(Mandatory = $true)]
        [string]$Architecture,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Agent', 'Agent 2')]
        [string]$AgentType
    )
    
    $agentString = if ($AgentType -eq 'Agent 2') { 'zabbix_agent2' } else { 'zabbix_agent' }
    $installerUrl = "${RepoUrl}${Version}/${agentString}-${Version}-${Architecture}.msi"
    
    Write-LogEntry -Message "Generated installer URL for $($AgentType): $installerUrl" -Level 'INFO'
    return $installerUrl
}

function Test-ServerConfigured {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ConfigFilePath,
        
        [Parameter(Mandatory = $true)]
        [string]$ZabbixServer
    )
    
    try {
        if (-not (Test-Path -Path $ConfigFilePath)) {
            Write-LogEntry -Message "Configuration file not found: $ConfigFilePath" -Level 'WARNING'
            return $false
        }
        
        $configContent = Get-Content -Path $ConfigFilePath -Raw
        
        # Check if the configuration contains the correct server
        if ($configContent -match [regex]::Escape($ZabbixServer)) {
            Write-LogEntry -Message "Configuration references correct Zabbix server: $ZabbixServer" -Level 'SUCCESS'
            return $true
        }
        else {
            Write-LogEntry -Message "Configuration does not reference Zabbix server: $ZabbixServer" -Level 'WARNING'
            return $false
        }
    }
    catch {
        Write-LogEntry -Message "Error checking server configuration: $_" -Level 'ERROR'
        return $false
    }
}

function Get-ConfigurationReplacements {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ZabbixServer,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Agent', 'Agent 2')]
        [string]$AgentType
    )
    
    if ($AgentType -eq 'Agent 2') {
        # Zabbix Agent 2 configuration patterns
        return @{
            '^\s*Server=.*'       = "Server=$ZabbixServer"
            '^\s*ServerActive=.*' = "ServerActive=$ZabbixServer"
            # Agent 2 might use slightly different hostname format or other configs
            '^\s*Hostname=.*'     = '# Hostname='
        }
    }
    else {
        # Standard Zabbix Agent configuration patterns
        return @{
            '^\s*Server=.*'       = "Server=$ZabbixServer"
            '^\s*ServerActive=.*' = "ServerActive=$ZabbixServer"
            '^\s*HOSTNAME=.*'     = '# HOSTNAME='
        }
    }
}

function Wait-ForServiceState {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServiceName,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Running', 'Stopped')]
        [string]$DesiredState,
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 60
    )
    
    $startTime = Get-Date
    $timeoutTime = $startTime.AddSeconds($TimeoutSeconds)
    
    Write-LogEntry -Message "Waiting for service $ServiceName to be in state: $DesiredState (timeout: ${TimeoutSeconds}s)" -Level 'INFO'
    
    while ((Get-Date) -lt $timeoutTime) {
        try {
            $service = Get-Service -Name $ServiceName -ErrorAction Stop
            if ($service.Status -eq $DesiredState) {
                Write-LogEntry -Message "Service $ServiceName is now $DesiredState" -Level 'SUCCESS'
                return $true
            }
            
            Start-Sleep -Seconds 2
        }
        catch {
            Write-LogEntry -Message "Error checking service status: $_" -Level 'WARNING'
            Start-Sleep -Seconds 2
        }
    }
    
    Write-LogEntry -Message "Timed out waiting for service $ServiceName to be $DesiredState" -Level 'ERROR'
    return $false
}

#endregion

#region Main Script Execution
try {
    # Initialize logging
    Initialize-LogFile
    Write-LogEntry -Message "Script execution started in $(if($Scheduled){'Scheduled Task'}else{'Interactive'}) mode" -Level 'INFO'
    
    # Auto-detect architecture if set to "auto"
    if ($Architecture -eq 'auto') {
        $Architecture = Get-SystemArchitecture
        Write-LogEntry -Message "Auto-detected architecture: $Architecture" -Level 'INFO'
    }
    # Register cleanup handler
    Register-CleanupHandler
    
    # Verify administrator privileges
    if (-not (Assert-Administrator)) {
        & $global:Cleanup 5
        exit 5
    }
    
    # Get latest version
    $latestVersion = Get-LatestVersion -RepoUrl $ZabbixRepo
    if ($null -eq $latestVersion) {
        Write-LogEntry -Message 'Could not determine latest version. Exiting.' -Level 'ERROR'
        & $global:Cleanup 2
        exit 2
    }
    
    # Main processing loop
    $runChecks = $true
    $iterationCount = 0
    $maxIterations = 3  # Prevent infinite loops
    
    while ($runChecks -and $iterationCount -lt $maxIterations) {
        $iterationCount++
        Write-LogEntry -Message "Starting iteration $iterationCount of $maxIterations" -Level 'INFO'
        $runChecks = $false  # Reset flag
        
        # Get Zabbix services
        $zabbixServices = Get-ZabbixService
        
        # Handle multiple services
        if ($zabbixServices.Count -gt 1) {
            Write-LogEntry -Message 'Multiple Zabbix services detected. Manual intervention required.' -Level 'ERROR'
            & $global:Cleanup 3
            exit 3
        }
        
        # Install Zabbix if no service exists
        if ($zabbixServices.Count -eq 0) {
            Write-LogEntry -Message "No Zabbix services found. Installing Zabbix $DefaultAgentType..." -Level 'INFO'
            
            # Generate the appropriate installer URL for the default agent type
            $installerUrl = Get-ZabbixInstallerUrl -RepoUrl $ZabbixRepo -Version $latestVersion -Architecture $Architecture -AgentType $DefaultAgentType
            
            if (Get-ZabbixInstaller -InstallerUrl $installerUrl -InstallerPath $global:InstallerPath) {
                if (Install-ZabbixAgent -InstallerPath $global:InstallerPath -AgentType $DefaultAgentType) {
                    Write-LogEntry -Message "Zabbix $DefaultAgentType installed successfully" -Level 'SUCCESS'
                    $runChecks = $true  # Re-run checks after installation
                }
                else {
                    Write-LogEntry -Message "Failed to install Zabbix $DefaultAgentType" -Level 'ERROR'
                    & $global:Cleanup 1
                    exit 1
                }
            }
            else {
                Write-LogEntry -Message "Failed to download Zabbix $DefaultAgentType installer" -Level 'ERROR'
                & $global:Cleanup 2
                exit 2
            }
        }
        # Service exists, check configuration
        else {
            foreach ($service in $zabbixServices) {
                Write-LogEntry -Message "Processing service: $($service.Name)" -Level 'INFO'
                
                # Get service details including agent type
                $serviceDetails = Get-ZabbixServiceDetails -Service $service
                
                if ($null -eq $serviceDetails) {
                    Write-LogEntry -Message "Could not get service details for $($service.Name). Skipping." -Level 'WARNING'
                    continue
                }
                
                # Get the appropriate configuration replacements for this agent type
                $replacements = Get-ConfigurationReplacements -ZabbixServer $ZabbixServer -AgentType $serviceDetails.AgentType
                
                # Generate the correct installer URL for this agent type
                $installerUrl = Get-ZabbixInstallerUrl -RepoUrl $ZabbixRepo -Version $latestVersion -Architecture $Architecture -AgentType $serviceDetails.AgentType
                
                # Check if executable exists
                if (-not (Test-Path -Path $serviceDetails.ExecutablePath)) {
                    Write-LogEntry -Message "Executable file not found at $($serviceDetails.ExecutablePath)" -Level 'ERROR'
                    
                    # Reinstall agent if executable is missing
                    if (Get-ZabbixInstaller -InstallerUrl $installerUrl -InstallerPath $global:InstallerPath) {
                        if (Install-ZabbixAgent -InstallerPath $global:InstallerPath -AgentType $serviceDetails.AgentType) {
                            Write-LogEntry -Message "Zabbix $($serviceDetails.AgentType) reinstalled due to missing executable" -Level 'SUCCESS'
                            $runChecks = $true  # Re-run checks after installation
                        }
                        else {
                            Write-LogEntry -Message "Failed to reinstall Zabbix $($serviceDetails.AgentType)" -Level 'ERROR'
                            & $global:Cleanup 1
                            exit 1
                        }
                    }
                    else {
                        Write-LogEntry -Message "Failed to download Zabbix $($serviceDetails.AgentType) installer" -Level 'ERROR'
                        & $global:Cleanup 2
                        exit 2
                    }
                    
                    continue
                }
                
                # Check if configuration file exists
                if (-not (Test-Path -Path $serviceDetails.ConfigFilePath)) {
                    Write-LogEntry -Message "Configuration file not found at $($serviceDetails.ConfigFilePath)" -Level 'WARNING'
                    
                    # Reinstall agent if config is missing
                    if (Get-ZabbixInstaller -InstallerUrl $installerUrl -InstallerPath $global:InstallerPath) {
                        if (Install-ZabbixAgent -InstallerPath $global:InstallerPath -AgentType $serviceDetails.AgentType) {
                            Write-LogEntry -Message "Zabbix $($serviceDetails.AgentType) reinstalled due to missing config" -Level 'SUCCESS'
                            $runChecks = $true  # Re-run checks after installation
                        }
                        else {
                            Write-LogEntry -Message "Failed to reinstall Zabbix $($serviceDetails.AgentType)" -Level 'ERROR'
                            & $global:Cleanup 1
                            exit 1
                        }
                    }
                    else {
                        Write-LogEntry -Message "Failed to download Zabbix $($serviceDetails.AgentType) installer" -Level 'ERROR'
                        & $global:Cleanup 2
                        exit 2
                    }
                    
                    continue
                }
                
                # Check if server is configured correctly
                if (-not (Test-ServerConfigured -ConfigFilePath $serviceDetails.ConfigFilePath -ZabbixServer $ZabbixServer)) {
                    Write-LogEntry -Message 'Server configuration needs update' -Level 'INFO'
                    
                    if (Update-ZabbixConfiguration -ConfigFilePath $serviceDetails.ConfigFilePath -ServiceName $service.Name -Replacements $replacements) {
                        Write-LogEntry -Message 'Configuration updated successfully' -Level 'SUCCESS'
                        $runChecks = $true  # Re-run checks after configuration update
                    }
                    else {
                        Write-LogEntry -Message 'Failed to update configuration' -Level 'ERROR'
                        & $global:Cleanup 4
                        exit 4
                    }
                    
                    continue
                }
                
                # Check for agent updates
                if (Test-AgentNeedsUpdate -ExecutablePath $serviceDetails.ExecutablePath -LatestVersion $latestVersion) {
                    Write-LogEntry -Message 'Agent update available' -Level 'INFO'
                    
                    if (Get-ZabbixInstaller -InstallerUrl $installerUrl -InstallerPath $global:InstallerPath) {
                        try {
                            if (Update-ZabbixAgent -InstallerPath $global:InstallerPath -ConfigFilePath $serviceDetails.ConfigFilePath -AgentType $serviceDetails.AgentType) {
                                Write-LogEntry -Message 'Agent updated successfully' -Level 'SUCCESS'
                                
                                $runChecks = $true  # Re-run checks after update
                            }
                            else {
                                Write-LogEntry -Message 'Failed to update agent' -Level 'ERROR'
                                
                                # Attempt to start service even after failed update
                                if ($PSCmdlet.ShouldProcess($service.Name, 'Start service after failed update')) {
                                    Start-Service -Name $service.Name -ErrorAction SilentlyContinue
                                }
                                
                                & $global:Cleanup 1
                                exit 1
                            }
                        }
                        catch {
                            Write-LogEntry -Message "Error during update process: $_" -Level 'ERROR'
                            
                            # Attempt to start service after error
                            if ($PSCmdlet.ShouldProcess($service.Name, 'Start service after error')) {
                                Start-Service -Name $service.Name -ErrorAction SilentlyContinue
                            }
                            
                            & $global:Cleanup 1
                            exit 1
                        }
                    }
                    else {
                        Write-LogEntry -Message "Failed to download Zabbix $($serviceDetails.AgentType) installer for update" -Level 'ERROR'
                        & $global:Cleanup 2
                        exit 2
                    }
                }
                else {
                    Write-LogEntry -Message "Zabbix $($serviceDetails.AgentType) is up to date and properly configured" -Level 'SUCCESS'
                }
            }
        }
        
        if ($runChecks) {
            Write-LogEntry -Message 'Rerunning checks...' -Level 'INFO'
        }
    }
    
    # Call cleanup with success code
    & $global:Cleanup 0
    exit 0
}
catch {
    # Catch any unhandled exceptions
    Write-LogEntry -Message "Unhandled exception: $_" -Level 'ERROR'
    Write-LogEntry -Message "Stack trace: $($_.ScriptStackTrace)" -Level 'ERROR'
    
    # Call cleanup with failure code
    & $global:Cleanup 1
    exit 1
}
#endregion