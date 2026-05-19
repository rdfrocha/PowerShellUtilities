# SPO Preservation Hold Cleanup
<#
.SYNOPSIS
Hard-deletes items from a OneDrive `Preservation Hold Library` in fixed-size batches using PnP.PowerShell.

.DESCRIPTION
This script is designed for controlled bulk deletion of items from a Preservation Hold Library (PHL),
typically after retention/eDiscovery conditions have been removed. It processes items in batches, uses
retry logic for transient failures (including throttling), and reports throughput/ETA as it runs.

Key behaviors:
- Deletes are issued with -Recycle:$false (hard-delete request, bypassing normal recycle flow for list-item delete).
- Batch execution uses Invoke-PnPBatch -StopOnException so failures are detected clearly.
- Retries are exponential (or based on server-supplied delay when available).
- Optional recycle-bin purge can be enabled after list-item deletion.
- Script stops on unhandled errors ($ErrorActionPreference = Stop).

.PREREQUISITES
- PowerShell 7+ or Windows PowerShell 5.1 with compatible PnP.PowerShell module.
- PnP.PowerShell installed and imported.
- You must connect to the target OneDrive site before deletion:
  Connect-PnPOnline -Url <OneDriveUrl> -ClientId <AppId>.
- Permissions must allow listing and deleting items from the target library.

.IMPORTANT
- If retention/eDiscovery holds are still active, deleted content can reappear.
- Validate hold state and legal/compliance requirements before running.
- Test first with a small `BatchSize` in a non-production or controlled scenario.

.CONFIGURATION
- `$ListTitle`: Target list/library title (default: Preservation Hold Library).
- `$BatchSize`: Items requested and queued per iteration (default: 200).
- `$PurgeRecycleBin`: When $true, clears first-stage and second-stage recycle bin at the end.
- `$DelayBetweenBatchesSec`: Optional cooldown between batches.
- `$MaxRetries`: Retry attempts per batch invocation.
- `$ErrorActionPreference`: Set to Stop for strict error handling.

.OUTPUT
- Console telemetry:
  - Batch size and duration
  - Batch rate (items/sec)
  - Rolling average rate (last 10 batches)
  - Overall rate
  - ETA based on rolling rate
- Final summary:
  - Total deleted
  - Total elapsed time
  - Average overall throughput

.EXAMPLE
# 1) Connect first (interactive):
Connect-PnPOnline -Url "https://<tenant>-my.sharepoint.com/personal/<user>" -ClientId "<app-id>"

# 2) Run script with defaults:
.\Purge-PHL.ps1

.EXAMPLE
# Run with larger batches and recycle-bin purge enabled:
$BatchSize = 500
$PurgeRecycleBin = $true
.\Purge-PHL.ps1
#>

# Optional references (left commented intentionally).
# $adminUrl    = "https://-admin.sharepoint.com/"
# $oneDriveUrl = "https://-my.sharepoint.com/personal/"

# Target library/list name to purge.
$ListTitle = 'Preservation Hold Library'

# Number of items to fetch and queue per loop iteration.
$BatchSize = 200

# If enabled, purge first-stage and second-stage recycle bins after item deletion completes.
$PurgeRecycleBin = $false

# Optional pause between batches to reduce service pressure.
$DelayBetweenBatchesSec = 0

# Maximum retry attempts for transient failures on batch execution.
$MaxRetries = 6

# Fail fast on non-terminating errors from cmdlets unless locally handled.
$ErrorActionPreference = 'Stop'

# Connect intentionally left commented so execution is explicit and operator-controlled.
# Connect-PnPOnline -Url $OneDriveUrl -ClientId 

function Get-RetryDelaySeconds {
    <#
    .SYNOPSIS
    Determines retry delay from exception details or exponential fallback.

    .DESCRIPTION
    Attempts to parse a server-supplied wait duration from known patterns
    (e.g., Retry-After headers/messages). If no explicit delay is found,
    returns exponential backoff with a capped maximum.

    .PARAMETER Attempt
    Current 1-based retry attempt number.

    .PARAMETER Exception
    Exception captured from the failed operation.

    .OUTPUTS
    [int] Number of seconds to wait before next retry.
    #>
    param(
        [int]$Attempt,
        [System.Exception]$Exception
    )

    # Normalize exception to searchable text.
    $msg = $Exception.ToString()

    # Pattern 1: Retry-After style messages.
    if ($msg -match 'Retry-After[^0-9]*([0-9]+)') { return [int]$Matches[1] }

    # Pattern 2: Generic "... X second(s) ..." message text.
    if ($msg -match '([0-9]+)\s*second') { return [int]$Matches[1] }

    # Fallback: exponential backoff (2^attempt), capped at 60 seconds.
    return [Math]::Min(60, [Math]::Pow(2, $Attempt))
}

function Invoke-WithRetry {
    <#
    .SYNOPSIS
    Executes an action with bounded retry behavior.

    .DESCRIPTION
    Runs the provided script block up to `$MaxRetries` times. On failure:
    - Logs warning with attempt counter and error message.
    - Waits using delay from `Get-RetryDelaySeconds`.
    - Rethrows on final failure.

    .PARAMETER Action
    Script block to execute.

    .PARAMETER ActionName
    Friendly name for logging context.
    #>
    param(
        [scriptblock]$Action,
        [string]$ActionName
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            & $Action
            return
        }
        catch {
            $isLast = $attempt -eq $MaxRetries
            $delay = Get-RetryDelaySeconds -Attempt $attempt -Exception $_.Exception

            Write-Warning "$ActionName failed (attempt $attempt/$MaxRetries): $($_.Exception.Message)"

            if ($isLast) { throw }

            Write-Host "Waiting $delay second(s) before retry..."
            Start-Sleep -Seconds $delay
        }
    }
}

try {
    # Validate target list exists before entering delete loop.
    $null = Get-PnPList -Identity $ListTitle
    Write-Host "Target list found: $ListTitle"

    # Capture initial count for progress reporting and ETA approximation.
    $listInfo = Get-PnPList -Identity $ListTitle -Includes ItemCount
    $initialCount = [int]$listInfo.ItemCount

    # Running totals/state.
    $totalDeleted = 0
    $loop = 0

    # Global stopwatch for end-to-end throughput.
    $swTotal = [System.Diagnostics.Stopwatch]::StartNew()

    # Rolling rates list (last 10 batches) for smoother ETA.
    $rates = New-Object System.Collections.Generic.List[double]

    while ($true) {
        $loop++

        # CAML query:
        # - RecursiveAll: include all folders/subfolders
        # - ViewFields: only fetch ID for minimal payload
        # - RowLimit: batch size per loop
        $query = "<View Scope='RecursiveAll'><ViewFields><FieldRef Name='ID'/></ViewFields><RowLimit Paged='FALSE'>$BatchSize</RowLimit></View>"
        $items = Get-PnPListItem -List $ListTitle -Query $query

        # Exit condition: no more items to delete.
        if (-not $items -or $items.Count -eq 0) {
            Write-Host 'No more items found.'
            break
        }

        $swBatch = [System.Diagnostics.Stopwatch]::StartNew()
        Write-Host ('Batch {0}: queueing {1} item(s)...' -f $loop, $items.Count)

        # Build PnP batch to reduce round-trips and improve throughput.
        $batch = New-PnPBatch
        foreach ($item in $items) {
            # Hard-delete request for each list item (no recycle on item delete call).
            Remove-PnPListItem -List $ListTitle -Identity $item.Id -Recycle:$false -Batch $batch
        }

        # Execute queued batch with retry protection.
        Invoke-WithRetry -Action { Invoke-PnPBatch -Batch $batch -StopOnException } -ActionName "Invoke-PnPBatch (batch $loop)"

        $swBatch.Stop()
        $totalDeleted += $items.Count

        # Per-batch telemetry.
        $batchSec = [math]::Round($swBatch.Elapsed.TotalSeconds, 2)
        $batchRate = if ($swBatch.Elapsed.TotalSeconds -gt 0) { $items.Count / $swBatch.Elapsed.TotalSeconds } else { 0.0 }

        # Maintain rolling rate window (last 10 batches).
        $rates.Add($batchRate)
        if ($rates.Count -gt 10) { $rates.RemoveAt(0) }

        $rollingRate = if ($rates.Count -gt 0) { ($rates | Measure-Object -Average).Average } else { 0.0 }
        $overallRate = if ($swTotal.Elapsed.TotalSeconds -gt 0) { $totalDeleted / $swTotal.Elapsed.TotalSeconds } else { 0.0 }

        # ETA uses initial count and rolling rate for a more stable estimate.
        $remaining = [math]::Max(0, $initialCount - $totalDeleted)
        $etaSec = if ($rollingRate -gt 0) { [int]($remaining / $rollingRate) } else { 0 }
        $etaTs = [timespan]::FromSeconds($etaSec)

        Write-Host ('Batch {0}: {1} items in {2}s ({3} items/s). Total: {4}/{5}. Rolling: {6} items/s. Overall: {7} items/s. ETA: {8:hh\:mm\:ss}' -f `
                $loop,
            $items.Count,
            $batchSec,
            [math]::Round($batchRate, 2),
            $totalDeleted,
            $initialCount,
            [math]::Round($rollingRate, 2),
            [math]::Round($overallRate, 2),
            $etaTs
        )

        # Optional pacing between batches.
        if ($DelayBetweenBatchesSec -gt 0) {
            Start-Sleep -Seconds $DelayBetweenBatchesSec
        }
    }

    if ($PurgeRecycleBin) {
        # First-stage recycle bin purge.
        Write-Host 'Purging first-stage recycle bin...'
        $first = Get-PnPRecycleBinItem -FirstStage
        foreach ($rb in $first) {
            Clear-PnPRecycleBinItem -Identity $rb.Id -Force
        }

        # Second-stage recycle bin purge.
        Write-Host 'Purging second-stage recycle bin...'
        $second = Get-PnPRecycleBinItem -SecondStage
        foreach ($rb in $second) {
            Clear-PnPRecycleBinItem -Identity $rb.Id -Force
        }
    }

    $swTotal.Stop()

    # Final summary with average rate across total elapsed time.
    Write-Host ('Completed. Total deleted: {0}. Total time: {1:hh\:mm\:ss}. Avg rate: {2} items/s' -f `
            $totalDeleted,
        $swTotal.Elapsed,
        [math]::Round(($totalDeleted / [math]::Max(1, $swTotal.Elapsed.TotalSeconds)), 2)
    )
}
catch {
    # Emit the primary exception message, then rethrow for caller/automation visibility.
    Write-Error $_.Exception.Message
    throw
}
finally {
    # Intentionally commented to avoid disconnecting an externally managed session.
    # Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
