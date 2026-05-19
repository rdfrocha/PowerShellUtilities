<#
.SYNOPSIS
    Try to backup the BitLocker key protector to Azure AD, force BitLocker recovery 
    on the OS drive and restart the computer.
    Workaround for Intune's "Remote Lock" not supporting Windows devices.

.DESCRIPTION
    This script checks for the OS drive protected by BitLocker, attempts to backup the key 
    protector to Azure AD, and then forces BitLocker recovery using manage-bde.exe.
    Restarts the computer if force recovery is successful.
    Transcript log is available and can be collected using Intune 'Collect Diagnostics'.

.EXAMPLE
    Use as Detection script for a remediation in Intune, run on demand from device details.

.OUTPUTS
    Writes log output to transcript file. Exits with code 0 on success, 1 on failure.

.NOTES
    - BitLocker must be enabled on the OS drive.
    - The device must be Azure AD joined and have the Intune Management Extension installed.
    - The user must have permissions to run BitLocker and backup keys to Azure AD.

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

# transcript for some debug capability
Start-Transcript -Path 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneWinRemoteLock.log' -Append

try {
    Write-Output 'Starting force BitLocker Remote Lock remediation script.'

    $osDrive = Get-BitLockerVolume | Where-Object { $_.VolumeType -eq 'OperatingSystem' }
    Write-Output "Checked for OS drive. Result: $($osDrive.MountPoint)"

    if ($osDrive) {
        # List all available key protectors
        Write-Output "Found key protectors:"
        $osDrive.KeyProtector | ForEach-Object { 
            Write-Output "  Type: $($_.KeyProtectorType), ID: $($_.KeyProtectorId)"
        }

        # Find a recovery password key protector (most suitable for backup)
        $keyProtector = $osDrive.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } | Select-Object -First 1

        if (-not $keyProtector) {
            # Fallback to any available key protector
            $keyProtector = $osDrive.KeyProtector | Select-Object -First 1
            Write-Output "No recovery password found, using $($keyProtector.KeyProtectorType) instead."
        }

        # Try to backup the key protector. No error handling, if it fails computer should still be locked.
        Write-Output "Backing up BitLocker key protector ($($keyProtector.KeyProtectorType))..."
        BackupToAAD-BitLockerKeyProtector -MountPoint $osDrive.MountPoint -KeyProtectorId $keyProtector.KeyProtectorId
        Write-Output 'Backup complete.'
        
        Write-Output 'Forcing BitLocker recovery...'
        & manage-bde.exe -forcerecovery $osDrive.MountPoint
        if ($LASTEXITCODE -ne 0) {
            throw "manage-bde.exe failed with exit code $LASTEXITCODE"
        }
        Write-Output 'Force recovery command issued.'
        
        Write-Output 'Remediation succeeded.'
        Write-Output 'Restarting computer...'
        Stop-Transcript
        Restart-Computer -Force
        # exit code might not register if computer restarts before
        exit 0
    }
    else {
        Write-Error 'No OS drive found.'
        Stop-Transcript
        exit 1
    }
}
catch {
    Write-Error "Remediation failed: $_"
    Stop-Transcript
    exit 1
}