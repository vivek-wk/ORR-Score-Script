[CmdletBinding(SupportsShouldProcess)]
param()

$ErrorActionPreference = 'Stop'
$resourceUrl = 'api://518a59cf-dba0-4cb9-8392-8a328a4831bf'
$apiBaseUrl = 'https://conformance-p3.pegasus.wkcloud.io/api/v4.0/orr'
$reportScript = Join-Path $PSScriptRoot 'Get-OrrScore.ps1'

function Read-MenuSelection {
    param(
        [Parameter(Mandatory)] [string]$Prompt,
        [Parameter(Mandatory)] [string[]]$Options,
        [string[]]$Default = @(),
        [switch]$MultiSelect
    )

    $selected = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $Default) { [void]$selected.Add($item) }
    $index = 0

    try {
        while ($true) {
            Clear-Host
            Write-Host ("{0}  (Up/Down move, {1}Enter accept)" -f $Prompt, $(if ($MultiSelect) { 'Space toggle, ' } else { '' })) -ForegroundColor Yellow
            foreach ($option in $Options) {
                $marker = if ($MultiSelect -and $selected.Contains($option)) { '[x]' } elseif ($MultiSelect) { '[ ]' } else { '' }
                $pointer = if ($Options[$index] -eq $option) { '>' } else { ' ' }
                Write-Host ("$pointer $marker $option").TrimEnd() -ForegroundColor $(if ($Options[$index] -eq $option) { 'White' } else { 'Gray' }) -BackgroundColor $(if ($Options[$index] -eq $option) { 'DarkBlue' } else { 'Black' })
            }

            $key = [Console]::ReadKey($true)
            switch ($key.Key) {
                'UpArrow' { $index = ($index - 1 + $Options.Count) % $Options.Count }
                'DownArrow' { $index = ($index + 1) % $Options.Count }
                'Spacebar' {
                    if ($MultiSelect) {
                        if ($selected.Contains($Options[$index])) { [void]$selected.Remove($Options[$index]) } else { [void]$selected.Add($Options[$index]) }
                    }
                }
                'Enter' {
                    Clear-Host
                    if ($MultiSelect) {
                        if ($selected.Count -eq 0) { [void]$selected.Add($Options[$index]) }
                        return @($Options | Where-Object { $selected.Contains($_) })
                    }
                    return $Options[$index]
                }
                'Escape' { Clear-Host; return $Default }
            }
        }
    }
    catch {
        $value = (Read-Host "$Prompt (enter a number 1-$($Options.Count))").Trim()
        if ($value -notmatch '^[1-9][0-9]*$' -or [int]$value -gt $Options.Count) { throw "Invalid selection '$value'." }
        return $Options[[int]$value - 1]
    }
}

function Read-OptionalBoolean {
    param([Parameter(Mandatory)] [string]$Prompt, [Parameter(Mandatory)] [bool]$Default)
    $default = if ($Default) { 'Yes' } else { 'No' }
    return (Read-MenuSelection -Prompt $Prompt -Options @('Yes', 'No') -Default @($default)) -eq 'Yes'
}

function Read-InlineConfirmation {
    param([Parameter(Mandatory)] [string]$Prompt)

    $options = @('Yes', 'No')
    $index = 0
    Write-Host $Prompt -ForegroundColor Yellow
    $startLine = [Console]::CursorTop
    try {
        while ($true) {
            for ($optionIndex = 0; $optionIndex -lt $options.Count; $optionIndex++) {
                [Console]::SetCursorPosition(0, $startLine + $optionIndex)
                $pointer = if ($optionIndex -eq $index) { '>' } else { ' ' }
                $line = "$pointer $($options[$optionIndex])"
                Write-Host $line.PadRight([Math]::Max(1, [Console]::WindowWidth - 1)) -ForegroundColor $(if ($optionIndex -eq $index) { 'White' } else { 'Gray' }) -BackgroundColor $(if ($optionIndex -eq $index) { 'DarkBlue' } else { 'Black' }) -NoNewline
            }
            [Console]::SetCursorPosition(0, $startLine + $options.Count)
            $key = [Console]::ReadKey($true)
            switch ($key.Key) {
                'UpArrow' { $index = ($index - 1 + $options.Count) % $options.Count }
                'DownArrow' { $index = ($index + 1) % $options.Count }
                'Enter' {
                    Write-Host ''
                    return $options[$index] -eq 'Yes'
                }
            }
        }
    }
    catch {
        return (Read-Host "$Prompt (Yes/No)").Trim() -match '^(?i:yes)$'
    }
}

function ConvertTo-PlainTextToken {
    param([Parameter(Mandatory)]$Token)
    if ($Token -isnot [securestring]) { return [string]$Token }
    $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Token)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer) }
}

function Get-OrrHeaders {
    if (-not (Get-Module -ListAvailable -Name Az.Accounts)) { throw 'Az.Accounts is not installed. Install it with: Install-Module Az.Accounts -Scope CurrentUser' }
    Import-Module Az.Accounts
    $accessToken = Get-AzAccessToken -ResourceUrl $resourceUrl
    $plainToken = ConvertTo-PlainTextToken -Token $accessToken.Token
    return @{ Authorization = "Bearer $plainToken"; Accept = 'application/json' }
}

function Start-OrrJob {
    param(
        [Parameter(Mandatory)] [string]$BitId,
        [Parameter(Mandatory)] [string]$Environment,
        [Parameter(Mandatory)] [hashtable]$Headers
    )

    $runUri = "$apiBaseUrl/jobs/$([uri]::EscapeDataString($BitId))"
    if (-not $PSCmdlet.ShouldProcess($runUri, "POST ORR evaluation for $BitId")) { return $null }
    try {
        $body = @{ environments = @($Environment) } | ConvertTo-Json -Compress
        return Invoke-RestMethod -Uri $runUri -Method Post -Headers $Headers -ContentType 'application/json' -Body $body -TimeoutSec 60
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode) {
            $responseBody = $null
            try {
                $response = $_.Exception.Response
                if ($response -is [System.Net.Http.HttpResponseMessage]) {
                    $responseBody = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                } else {
                    $stream = $response.GetResponseStream()
                    $reader = [System.IO.StreamReader]::new($stream)
                    $responseBody = $reader.ReadToEnd()
                    $reader.Dispose()
                }
            }
            catch { }
            $detail = if ([string]::IsNullOrWhiteSpace($responseBody)) { $_.Exception.Message } else { $responseBody.Trim() }
            throw "ORR evaluation request for $BitId failed with HTTP $statusCode. $detail"
        }
        throw "Unable to trigger ORR evaluation for $BitId. $($_.Exception.Message)"
    }
}

function Get-OrrJobStatus {
    param(
        [Parameter(Mandatory)]$QueueItem,
        [Parameter(Mandatory)] [hashtable]$Headers
    )

    $statusUrl = $QueueItem.StatusUrl
    if ([string]::IsNullOrWhiteSpace($statusUrl)) {
        if ($QueueItem.JobId) {
            $statusUrl = "$apiBaseUrl/jobs/$([uri]::EscapeDataString([string]$QueueItem.JobId))/status"
        } else {
            $QueueItem.Status = 'Unknown'
            $QueueItem.Message = 'No status URL or job id available for this job.'
            $QueueItem.Done = $true
            $QueueItem.UpdatedUtc = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            return $true
        }
    }

    try {
        $status = Invoke-RestMethod -Uri $statusUrl -Method Get -Headers $Headers -TimeoutSec 60 -ErrorAction Stop
        $QueueItem.Status = if ([string]::IsNullOrWhiteSpace([string]$status.status)) { 'Running' } else { [string]$status.status }
        $QueueItem.Message = if ([string]::IsNullOrWhiteSpace([string]$status.message)) { '' } else { [string]$status.message }
        $QueueItem.UpdatedUtc = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        if ($QueueItem.Status -in @('Completed', 'Success', 'Succeeded', 'Failed', 'Error', 'Cancelled', 'Canceled')) {
            $QueueItem.Done = $true
        }
        return $QueueItem.Done
    }
    catch {
        $QueueItem.UpdatedUtc = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $statusCode = $_.Exception.Response.StatusCode.value__
        switch ($statusCode) {
            422 {
                # This API transitions a finished job to a silent 422 terminal state.
                $QueueItem.Status = 'Completed'
                $QueueItem.Message = 'Job reached its terminal state. Confirm the result in the report.'
                $QueueItem.Done = $true
            }
            404 {
                $QueueItem.Status = 'Unavailable'
                $QueueItem.Message = 'Job status returned 404 (job not found).'
                $QueueItem.Done = $true
            }
            default {
                $QueueItem.FailedPolls++
                if ($QueueItem.FailedPolls -ge 5) {
                    $QueueItem.Status = 'Unavailable'
                    $QueueItem.Message = "Status endpoint unreachable after 5 attempts: $($_.Exception.Message)"
                    $QueueItem.Done = $true
                } else {
                    $QueueItem.Message = "Status poll error ($statusCode); will retry."
                }
            }
        }
        return $QueueItem.Done
    }
}

function Wait-OrrJob {
    param(
        [Parameter(Mandatory)] [object]$Job,
        [Parameter(Mandatory)] [string]$Environment,
        [Parameter(Mandatory)] [hashtable]$Headers,
        [int]$TimeoutSeconds = 120,
        [int]$PollIntervalSeconds = 3
    )

    if ($null -eq $Job) { return }

    $statusUrl = if ($Job.statusUrl) {
        [string]$Job.statusUrl
    } elseif ($Job.url) {
        [string]$Job.url
    } elseif ($Job.jobId) {
        "$apiBaseUrl/jobs/$([uri]::EscapeDataString([string]$Job.jobId))"
    } else {
        $null
    }

    $queueItem = [pscustomobject]@{
        StatusUrl   = $statusUrl
        Environment = $Environment
        Status      = if ([string]::IsNullOrWhiteSpace([string]$Job.status)) { 'Pending' } else { [string]$Job.status }
        Message     = ''
        JobId       = [string]$Job.jobId
        Application = '(batch job)'
        BitId       = $Job.bitId
        Number      = 0
        UpdatedUtc  = ''
        Done        = $false
        FailedPolls = 0
    }

    $startTime = [datetime]::UtcNow
    while ($queueItem.Status -in @('Pending', 'Queued', 'Running', 'InProgress', 'Processing')) {
        if (([datetime]::UtcNow - $startTime).TotalSeconds -ge $TimeoutSeconds) {
            Write-Host "    Timeout waiting for job $($Job.jobId) (status: $($queueItem.Status))" -ForegroundColor Yellow
            break
        }
        Start-Sleep -Seconds $PollIntervalSeconds
        Get-OrrJobStatus -QueueItem $queueItem -Headers $Headers
    }

    if ($queueItem.Status -in @('Completed', 'Success', 'Succeeded')) {
        Write-Host "    Job $($Job.jobId) completed." -ForegroundColor Green
    }
    elseif ($queueItem.Status -in @('Failed', 'Error')) {
        Write-Host "    Job $($Job.jobId) failed: $($queueItem.Message)" -ForegroundColor Red
    }
    else {
        Write-Host "    Job $($Job.jobId) status: $($queueItem.Status)" -ForegroundColor Gray
    }
}

function Show-OrrJobTable {
    param([Parameter(Mandatory)] [object[]]$QueueItems)

    Clear-Host
    Write-Host ('ORR evaluation jobs ({0})  -  live status refreshed {1} UTC' -f $QueueItems.Count, (Get-Date -Format 'HH:mm:ss')) -ForegroundColor Cyan
    Write-Host ('=' * 100) -ForegroundColor DarkCyan
    $QueueItems |
        Select-Object Number, Application, BitId, Environment, Status, Message, UpdatedUtc, JobId |
        Format-Table -AutoSize |
        Out-String -Width 260 |
        Write-Host
}

function Watch-OrrJobs {
    param(
        [Parameter(Mandatory)] [object[]]$QueueItems,
        [Parameter(Mandatory)] [hashtable]$Headers,
        [int]$TimeoutSeconds = 3600,
        [int]$PollIntervalSeconds = 5
    )

    if (-not $QueueItems -or $QueueItems.Count -eq 0) { return }

    $startTime = [datetime]::UtcNow
    do {
        Start-Sleep -Seconds $PollIntervalSeconds
        foreach ($item in $QueueItems) {
            if (-not $item.Done) { [void](Get-OrrJobStatus -QueueItem $item -Headers $Headers) }
        }
        Show-OrrJobTable -QueueItems $QueueItems
        if (([datetime]::UtcNow - $startTime).TotalSeconds -ge $TimeoutSeconds) {
            Write-Host "Timeout after $TimeoutSeconds seconds; statuses may be incomplete." -ForegroundColor Yellow
            foreach ($item in $QueueItems) {
                if (-not $item.Done) { $item.Done = $true; $item.Status = 'Timed out' }
            }
        }
    } while (@($QueueItems | Where-Object { -not $_.Done }).Count -gt 0)

    Show-OrrJobTable -QueueItems $QueueItems
    $completed = @($QueueItems | Where-Object { $_.Status -in @('Completed', 'Success', 'Succeeded') })
    $attention = @($QueueItems | Where-Object { $_.Status -in @('Failed', 'Error', 'Unavailable', 'Timed out', 'Not queued', 'Unknown') })
    Write-Host ''
    Write-Host ("Result: {0} of {1} job(s) finished." -f $completed.Count, $QueueItems.Count) -ForegroundColor $(if ($attention.Count -eq 0) { 'Green' } else { 'Yellow' })
    if ($attention.Count -gt 0) {
        Write-Host 'Jobs needing attention:' -ForegroundColor Yellow
        foreach ($item in $attention) {
            Write-Host ("  {0} [{1}] - {2}: {3}" -f $item.Application, $item.BitId, $item.Status, $item.Message)
        }
    }
    Write-Host 'Confirm the resulting scores by choosing "Review recent scores".' -ForegroundColor Gray
}

function Get-OrrLatestEvaluationId {
    param(
        [Parameter(Mandatory)] [string]$BitId,
        [Parameter(Mandatory)] [string]$Environment,
        [Parameter(Mandatory)] [hashtable]$Headers
    )

    $uri = "$apiBaseUrl/evaluations?bitId=$([uri]::EscapeDataString($BitId))&environment=$([uri]::EscapeDataString($Environment))&latestOnly=true"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $Headers -TimeoutSec 30 -ErrorAction Stop
        if (@($response.items).Count -gt 0) { return [string]$response.items[0].id }
    }
    catch { }
    return ''
}

function Confirm-OrrEvaluations {
    param(
        [Parameter(Mandatory)] [object[]]$QueueItems,
        [Parameter(Mandatory)] [hashtable]$Headers
    )

    foreach ($item in $QueueItems) {
        if ($item.Status -in @('Not queued', 'Unavailable', 'Timed out', 'Unknown', 'Unverifiable')) { continue }
        if ([string]::IsNullOrWhiteSpace([string]$item.PreEvaluationId)) {
            $item.Status = 'Unverifiable'
            $item.Message = 'Pre-run evaluation baseline was unavailable; check the portal to confirm.'
            continue
        }
        $newId = Get-OrrLatestEvaluationId -BitId $item.BitId -Environment $item.Environment -Headers $Headers
        if ([string]::IsNullOrWhiteSpace($newId)) {
            $item.Status = 'No evaluation'
            $item.Message = 'No evaluation is on record for this environment.'
        }
        elseif ($newId -cne [string]$item.PreEvaluationId) {
            $item.Status = 'New evaluation'
            $item.Message = "Evaluation written: $newId"
        }
        else {
            $item.Status = 'Cached (no change)'
            $item.Message = 'No new evaluation written; the API reused the cached result because the underlying data did not change. The portal still shows the last evaluation.'
        }
    }

    $verified = @($QueueItems | Where-Object { $_.Status -in @('New evaluation', 'Cached (no change)', 'No evaluation', 'Unverifiable') })
    if ($verified.Count -gt 0) {
        Write-Host ''
        Write-Host 'Verification - did each finished job write a new evaluation for the portal?' -ForegroundColor Cyan
        $verified |
            Select-Object Application, BitId, Environment, Status, Message |
            Format-Table -AutoSize -Wrap |
            Out-String -Width 200 |
            Write-Host
    }

    $newCount = @($QueueItems | Where-Object { $_.Status -eq 'New evaluation' }).Count
    $cachedCount = @($QueueItems | Where-Object { $_.Status -eq 'Cached (no change)' }).Count
    if ($newCount -gt 0) {
        Write-Host ("{0} job(s) produced a NEW evaluation - visible in the portal and in 'Review recent scores'." -f $newCount) -ForegroundColor Green
    }
    if ($cachedCount -gt 0) {
        Write-Host ("{0} job(s) finished but wrote NO new evaluation - cached results were reused (underlying data unchanged). The portal will still show the last evaluation; there is nothing new to review yet." -f $cachedCount) -ForegroundColor Yellow
    }
}

$action = Read-MenuSelection -Prompt 'ORR action' -Options @('Run evaluations', 'Review recent scores') -Default @('Review recent scores')
if ($action -eq 'Review recent scores') {
    & $reportScript
    exit $LASTEXITCODE
}

$selection = Read-MenuSelection -Prompt 'Evaluate' -Options @('One BIT ID', 'Multiple BIT IDs') -Default @('One BIT ID')
$bitId = $null
$division = $null
if ($selection -eq 'One BIT ID') {
    $bitId = (Read-Host 'Enter BIT ID').Trim()
    if ($bitId -notmatch '^[A-Za-z0-9]+$') { throw "Invalid BIT ID '$bitId'. Enter only letters and numbers." }
}
else {
    $division = Read-MenuSelection -Prompt 'Department / Division' -Options @('ALL', 'TAA', 'HLT', 'FCC', 'CORP', 'CPESG', 'DXG', 'LNR', 'GBS', 'GGM', 'UNKNOWN') -Default @('ALL')
}

$validEnvironments = @('Production', 'DisasterRecovery', 'Development', 'Integration', 'LoadTesting', 'NonProduction', 'QualityAssurance', 'Staging', 'Sandbox', 'UserAcceptanceTesting', 'All')
$environments = @(Read-MenuSelection -Prompt 'Environment' -Options $validEnvironments -Default @('Production') -MultiSelect)

$ipmOnly = $false
$activeOnly = $true
if (-not $bitId) {
    $ipmOnly = Read-OptionalBoolean -Prompt 'Show IPM-managed applications only?' -Default $true
    $activeOnly = Read-OptionalBoolean -Prompt 'Show active applications only?' -Default $true
}

$headers = Get-OrrHeaders

$applications = if ($bitId) {
    @([pscustomobject]@{ BitId = $bitId; Application = '(single BIT ID)' })
} else {
    $matchingBitIds = [System.Collections.Generic.List[string]]::new()
    $matchingApplications = [System.Collections.Generic.List[object]]::new()
    $continuationToken = $null
    do {
        $query = "environments=$([uri]::EscapeDataString(($environments -join ',')))&ipmOnly=$($ipmOnly.ToString().ToLowerInvariant())&activeOnly=$($activeOnly.ToString().ToLowerInvariant())&latestOnly=true&limit=200&includeCount=true"
        if ($division -and $division -ne 'ALL') { $query += "&division=$([uri]::EscapeDataString($division))" }
        if ($continuationToken) { $query += "&continuationToken=$([uri]::EscapeDataString($continuationToken))" }
        $response = Invoke-RestMethod -Uri "$apiBaseUrl/evaluations?$query" -Method Get -Headers $headers -TimeoutSec 60
        foreach ($item in @($response.items)) {
            [void]$matchingBitIds.Add([string]$item.bitId)
            [void]$matchingApplications.Add([pscustomobject]@{ BitId = [string]$item.bitId; Application = [string]$item.applicationName })
        }
        $continuationToken = $response.continuationToken
    } while (-not [string]::IsNullOrWhiteSpace($continuationToken))
    @($matchingApplications | Sort-Object BitId -Unique)
}

$queueItems = @(
    foreach ($application in $applications) {
        foreach ($environment in $environments) {
            [pscustomobject]@{
                BitId = $application.BitId
                Application = $application.Application
                Environment = $environment
            }
        }
    }
)

if ($queueItems.Count -eq 0) { Write-Host 'No applications matched the selected filters.' -ForegroundColor Yellow; exit 0 }
Write-Host "Applications selected for evaluation ($($queueItems.Count) job(s)):" -ForegroundColor Cyan
$applicationNumber = 0
foreach ($queueItem in $queueItems) {
    $applicationNumber++
    Write-Host ("  {0,3}. {1} [{2}] - {3}" -f $applicationNumber, $queueItem.Application, $queueItem.BitId, $queueItem.Environment) -ForegroundColor Gray
}
if (-not (Read-InlineConfirmation -Prompt 'Queue evaluations for all listed applications?')) {
    Write-Host 'Evaluation run cancelled.' -ForegroundColor Yellow
    exit 0
}

Write-Host "Queueing evaluations for $($queueItems.Count) job(s)..." -ForegroundColor Cyan
Write-Host 'For each app, the script captures the latest evaluation id before and after the run, so you will know whether the portal got a NEW evaluation or the API reused the cached one.' -ForegroundColor Gray
try {
    $jobItems = [System.Collections.Generic.List[object]]::new()
    $queueFailed = [System.Collections.Generic.List[string]]::new()
    $counter = 0
    foreach ($queueItem in $queueItems) {
        $counter++
        $targetBitId = $queueItem.BitId
        Write-Progress -Activity 'Queueing ORR evaluations' -Status "Submitting $targetBitId in $($queueItem.Environment) ($($counter) of $($queueItems.Count))" -PercentComplete ([int](($counter / $queueItems.Count) * 100))
        $preEvalId = Get-OrrLatestEvaluationId -BitId $targetBitId -Environment $queueItem.Environment -Headers $headers
        $job = $null
        try {
            $job = Start-OrrJob -BitId $targetBitId -Environment $queueItem.Environment -Headers $headers
        }
        catch {
            $queueFailed.Add("$targetBitId [$($queueItem.Environment)]")
            Write-Host "  $targetBitId in $($queueItem.Environment) failed to queue: $($_.Exception.Message)" -ForegroundColor Red
        }

        if ($null -eq $job) {
            $jobItems.Add([pscustomobject]@{
                Number      = $counter
                Application = $queueItem.Application
                BitId       = $targetBitId
                Environment = $queueItem.Environment
                Status      = 'Not queued'
                Message     = 'Job was not submitted.'
                JobId       = ''
                StatusUrl   = $null
                UpdatedUtc  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                Done        = $true
                FailedPolls = 0
            })
            continue
        }

        $jobStatusUrl = if ($job.statusUrl) {
            [string]$job.statusUrl
        } elseif ($job.jobId) {
            "$apiBaseUrl/jobs/$([uri]::EscapeDataString([string]$job.jobId))/status"
        } else {
            $null
        }
        $initialStatus = if ([string]::IsNullOrWhiteSpace([string]$job.status)) { 'Pending' } else { [string]$job.status }
        $jobItems.Add([pscustomobject]@{
Number      = $counter
                Application = $queueItem.Application
                BitId       = $targetBitId
                Environment = $queueItem.Environment
                Status      = $initialStatus
                Message     = if ($job.message) { [string]$job.message } else { '' }
                JobId       = [string]$job.jobId
                StatusUrl   = $jobStatusUrl
                UpdatedUtc  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                Done        = $initialStatus -in @('Completed', 'Success', 'Succeeded', 'Failed', 'Error', 'Cancelled', 'Canceled')
                FailedPolls = 0
                PreEvaluationId = $preEvalId
            })
        Write-Host ("  {0}: {1} (job {2})" -f $targetBitId, $initialStatus, $job.jobId) -ForegroundColor Green
    }
}
finally {
    Write-Progress -Activity 'Queueing ORR evaluations' -Completed
}

Show-OrrJobTable -QueueItems @($jobItems)
$pendingItems = @($jobItems | Where-Object { -not $_.Done })
if ($pendingItems.Count -gt 0) {
    Write-Host "Waiting on $($pendingItems.Count) job(s) - Ctrl+C to stop polling." -ForegroundColor Yellow
    Watch-OrrJobs -QueueItems $pendingItems -Headers $headers
}
else {
    Write-Host 'All jobs are already finished; see the table above.' -ForegroundColor Gray
}
Confirm-OrrEvaluations -QueueItems @($jobItems) -Headers $headers
if ($queueFailed.Count -gt 0) {
    Write-Host "Failed to queue ($($queueFailed.Count)): $($queueFailed -join ', ')" -ForegroundColor Red
}
