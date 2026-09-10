[CmdletBinding()]
param(
    [Parameter()]
    [ValidateSet('Table', 'Detailed', 'Summary')]
    [string]$OutputMode,

    [Parameter()]
    [ValidatePattern('^[A-Za-z0-9]+$')]
    [string]$BitId,

    [Parameter()]
    [string]$Environment,

    [Parameter()]
    [string]$Division,

    [Parameter()]
    [Nullable[bool]]$IpmOnly,

    [Parameter()]
    [Nullable[bool]]$EvaluatedOnly,

    [Parameter()]
    [Nullable[bool]]$ActiveOnly,

    [Parameter()]
    [Nullable[bool]]$Live,

    [Parameter()]
    [Nullable[bool]]$Preview,

    [Parameter()]
    [Nullable[bool]]$Exception,

    [Parameter()]
    [Nullable[bool]]$Maintenance,

    [Parameter()]
    [Nullable[bool]]$Failing,

    [Parameter()]
    [switch]$ExportPdf,

    [Parameter()]
    [string]$PdfPath,

    [Parameter()]
    [ValidateSet('HTML', 'PDF')]
    [string]$ExportFormat,

    [Parameter()]
    [switch]$AllMatching
)

$ErrorActionPreference = 'Stop'
$scriptVersion = '1.1'
$resourceUrl = 'api://518a59cf-dba0-4cb9-8392-8a328a4831bf'
$apiBaseUrl = 'https://conformance-p3.pegasus.wkcloud.io/api/v4.0/orr'
$reportsDirectory = Join-Path $PSScriptRoot 'reports'

function Write-Section {
    param(
        [Parameter(Mandatory)]
        [string]$Title,

        [ConsoleColor]$Color = [ConsoleColor]::White
    )

    Write-Host ''
    Write-Host $Title -ForegroundColor $Color
    Write-Host ('=' * $Title.Length) -ForegroundColor $Color
}

function Write-Field {
    param(
        [Parameter(Mandatory)]
        [string]$Label,

        $Value,

        [ConsoleColor]$ValueColor = [ConsoleColor]::Gray
    )

    $displayValue = if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        '(none)'
    }
    else {
        [string]$Value
    }

    Write-Host ('  {0,-21}: ' -f $Label) -NoNewline -ForegroundColor DarkGray
    Write-Host $displayValue -ForegroundColor $ValueColor
}

function Get-DisplayRemark {
    param(
        [AllowNull()]
        [string]$Remark,

        [Parameter(Mandatory)]
        [string]$BitId
    )

    if ($null -eq $Remark -or [string]::IsNullOrWhiteSpace($Remark)) { return '' }
    return ($Remark -replace '(?i)<BIT_ID>', $BitId)
}

function ConvertTo-HtmlDisplayValue {
    param(
        [AllowNull()]
        $Value
    )

    $displayValue = if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { '(none)' } else { [string]$Value }
    return [System.Net.WebUtility]::HtmlEncode($displayValue) -replace "`r?`n", '<br>'
}

function Format-Score {
    param(
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return '(none)'
    }
    $scoreText = [string]$Value
    if ($scoreText.EndsWith('%')) {
        return $scoreText
    }
    return "$scoreText%"
}

function Read-MenuSelection {
    param(
        [Parameter(Mandatory)]
        [string]$Prompt,

        [Parameter(Mandatory)]
        [string[]]$Options,

        [string[]]$Default = @(),

        [switch]$MultiSelect
    )

    $selected = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $Default) {
        [void]$selected.Add($item)
    }
    $index = 0

    try {
        while ($true) {
            Clear-Host
            Write-Host ("{0}  (Up/Down move, {1}Enter accept)" -f $Prompt, $(if ($MultiSelect) { 'Space toggle, ' } else { '' })) -ForegroundColor Yellow
            foreach ($option in $Options) {
                $marker = if ($selected.Contains($option)) { '[x]' } else { '[ ]' }
                $pointer = if ($Options[$index] -eq $option) { '>' } else { ' ' }
                $line = if ($MultiSelect) { "$pointer $marker $option" } else { "$pointer $option" }
                $isCurrent = $Options[$index] -eq $option
                $isSelected = $selected.Contains($option)
                $foreground = if ($isCurrent) { 'White' } else { 'Gray' }
                $background = if ($isCurrent) { 'DarkBlue' } else { 'Black' }
                Write-Host ($line.PadRight([Math]::Max(1, [Console]::WindowWidth - 1))) -ForegroundColor $foreground -BackgroundColor $background
            }
            Write-Host (' ' * [Math]::Max(1, [Console]::WindowWidth - 1))

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
                    if ($MultiSelect) {
                        if ($selected.Count -eq 0) { [void]$selected.Add($Options[$index]) }
                        Clear-Host
                        return @($Options | Where-Object { $selected.Contains($_) })
                    }
                    Clear-Host
                    return $Options[$index]
                }
                'Escape' {
                    Clear-Host
                    return $Default
                }
            }
        }
    }
    catch {
        $value = (Read-Host "$Prompt (enter a number 1-$($Options.Count))").Trim()
        if ($value -notmatch '^[1-9][0-9]*$' -or [int]$value -gt $Options.Count) {
            throw "Invalid selection '$value'. Choose a number from 1 to $($Options.Count)."
        }
        return $Options[[int]$value - 1]
    }
}

function Write-CheckpointTableLine {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [AllowNull()]
        [string[]]$Values,

        [Parameter(Mandatory)]
        [int[]]$Widths
    )

    Write-Host '| ' -NoNewline -ForegroundColor Gray
    for ($cellIndex = 0; $cellIndex -lt $Values.Count; $cellIndex++) {
        if ($cellIndex -eq $Values.Count - 1 -and $Values[$cellIndex] -match '^(.*?)(\[(Preview|Exception|Live|Maintenance)\])(.*)$') {
            Write-Host $Matches[1] -NoNewline -ForegroundColor Gray
            $chipColor = switch ($Matches[3]) {
                'Preview' { 'Gray' }
                'Exception' { 'Red' }
                'Live' { 'Blue' }
                'Maintenance' { 'Yellow' }
            }
            Write-Host $Matches[2] -NoNewline -ForegroundColor $chipColor
            Write-Host $Matches[4] -NoNewline -ForegroundColor Gray
            $padding = $Widths[$cellIndex] - $Values[$cellIndex].Length
            if ($padding -gt 0) { Write-Host (' ' * $padding) -NoNewline }
        }
        else {
            Write-Host $Values[$cellIndex].PadRight($Widths[$cellIndex]) -NoNewline -ForegroundColor Gray
        }
        Write-Host ' |' -NoNewline -ForegroundColor Gray
        if ($cellIndex -lt $Values.Count - 1) { Write-Host ' ' -NoNewline }
    }
    Write-Host ''
}

function Write-CheckpointTable {
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows
    )

    $headers = @('No.', 'BIT ID', 'App Name', 'Last Evaluation (UTC)', 'Score', 'Failed Checkpoints')
    $groups = @($Rows | Group-Object BitId, Application)
    $tableRows = for ($groupIndex = 0; $groupIndex -lt $groups.Count; $groupIndex++) {
        $group = $groups[$groupIndex]
        $first = $group.Group[0]
        $uniqueCheckpoints = @($group.Group | Sort-Object Checkpoint, CheckpointVersion, Result, Status, NormalizedRemark -Unique)
        $checkpointText = ($uniqueCheckpoints | ForEach-Object {
                "{0} ({1}) - {2}`n[{3}]`n{4}" -f $_.Checkpoint, $_.CheckpointVersion, $_.Name, $_.Status, (Get-DisplayRemark -Remark $_.NormalizedRemark -BitId $_.BitId)
            }) -join "`n--------------------------------------------------------------------------------`n"
        if ([string]::IsNullOrWhiteSpace($checkpointText)) {
            $checkpointText = '(none)'
        }
        [pscustomobject]@{
            Number = $groupIndex + 1
            BitId = $first.BitId
            Application = $first.Application
            GeneratedAtUtc = $first.GeneratedAtUtc
            Score = Format-Score $first.Score
            FailedCheckpoints = $checkpointText
        }
    }
    $widths = @(4, 14, 24, 19, 8, 80)
    $border = '+' + (($widths | ForEach-Object { '-' * ($_ + 2) }) -join '+') + '+'
    Write-Host $border -ForegroundColor Gray
    Write-Host (('| {0,-' + $widths[0] + '} | {1,-' + $widths[1] + '} | {2,-' + $widths[2] + '} | {3,-' + $widths[3] + '} | {4,-' + $widths[4] + '} | {5,-' + $widths[5] + '} |') -f $headers[0], $headers[1], $headers[2], $headers[3], $headers[4], $headers[5]) -ForegroundColor Gray
    Write-Host $border -ForegroundColor Gray
    foreach ($row in $tableRows) {
        $checkpointLines = @($row.FailedCheckpoints -split "`r?`n" | ForEach-Object {
                $remaining = $_.Trim()
                while ($remaining.Length -gt $widths[5]) {
                    $breakAt = $remaining.LastIndexOf(' ', $widths[5])
                    if ($breakAt -lt 1) { $breakAt = $widths[5] }
                    $remaining.Substring(0, $breakAt).Trim()
                    $remaining = $remaining.Substring($breakAt).Trim()
                }
                if ($remaining) { $remaining }
            })
        for ($lineIndex = 0; $lineIndex -lt $checkpointLines.Count; $lineIndex++) {
            $values = if ($lineIndex -eq 0) {
                @($row.Number, $row.BitId, $row.Application, $row.GeneratedAtUtc, $row.Score, $checkpointLines[$lineIndex])
            }
            else {
                @('', '', '', '', '', $checkpointLines[$lineIndex])
            }
            Write-CheckpointTableLine -Values $values -Widths $widths
        }
        if ($row -ne $tableRows[-1]) {
            Write-Host ('.' * ($border.Length)) -ForegroundColor DarkGray
        }
    }
    Write-Host $border -ForegroundColor Gray
}

function New-CheckpointHtml {
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows,

        [ValidateSet('Table', 'Detailed', 'Summary')]
        [string]$ReportMode = 'Table'
    )

    $groups = @($Rows | Group-Object BitId, Application)
    if ($ReportMode -ne 'Table') {
        $sections = for ($groupIndex = 0; $groupIndex -lt $groups.Count; $groupIndex++) {
            $group = $groups[$groupIndex]
            $first = $group.Group[0]
            if ($ReportMode -eq 'Detailed') {
                $checkpointHtml = ($group.Group | Sort-Object Checkpoint, CheckpointVersion, Result, Status, NormalizedRemark -Unique | ForEach-Object {
                    $lifecycleColor = switch -Regex ([string]$_.Status) {
                        'Maintenance' { '#b36b00'; break }
                        'Exception' { '#c00000'; break }
                        'Preview' { '#666666'; break }
                        default { '#4da3ff' }
                    }
                    $fields = @(
                        ('<div class="checkpoint-title"><a href="{0}">{1} ({2}) - {3}</a></div>' -f $_.CheckpointUrl, (ConvertTo-HtmlDisplayValue $_.Checkpoint), (ConvertTo-HtmlDisplayValue $_.CheckpointVersion), (ConvertTo-HtmlDisplayValue $_.Name)),
                        ('<div><b>Version:</b> {0}</div>' -f (ConvertTo-HtmlDisplayValue $_.CheckpointVersion)),
                        ('<div><b>Status:</b> {0}</div>' -f (ConvertTo-HtmlDisplayValue $_.Result)),
                        ('<div><b>Lifecycle:</b> <font color="{0}"><b><i>[ {1} ]</i></b></font></div>' -f $lifecycleColor, (ConvertTo-HtmlDisplayValue $_.Status)),
                        ('<div><b>Is failing:</b> {0}</div>' -f (ConvertTo-HtmlDisplayValue $_.IsFailing)),
                        ('<div><b>Is preview:</b> {0}</div>' -f (ConvertTo-HtmlDisplayValue $_.IsPreview)),
                        ('<div><b>Has exception:</b> {0}</div>' -f (ConvertTo-HtmlDisplayValue $_.HasException)),
                        ('<div><b>Is scored:</b> {0}</div>' -f (ConvertTo-HtmlDisplayValue $_.IsScored)),
                        ('<div><b>Is maintenance:</b> {0}</div>' -f (ConvertTo-HtmlDisplayValue $_.IsMaintenance)),
                        ('<div><b>Exception number:</b> {0}</div>' -f (ConvertTo-HtmlDisplayValue $_.ExceptionNumber)),
                        ('<div><b>State reason:</b> {0}</div>' -f (ConvertTo-HtmlDisplayValue $_.StateReason)),
                        ('<div><b>State end date:</b> {0}</div>' -f (ConvertTo-HtmlDisplayValue $_.StateEndDate)),
                        ('<div><b>Accountable:</b> {0}</div>' -f (ConvertTo-HtmlDisplayValue $_.Accountable)),
                        ('<div><b>System:</b> {0}</div>' -f (ConvertTo-HtmlDisplayValue $_.Systems)),
                        ('<div><b>Remarks:</b> {0}</div>' -f (ConvertTo-HtmlDisplayValue $_.OriginalRemarks)),
                        ('<div><b>References:</b> {0}</div>' -f (ConvertTo-HtmlDisplayValue $_.References))
                    )
                    '<article class="checkpoint-detail">{0}</article>' -f ($fields -join "`n")
                }) -join "`n"
                '<section><h2>{0}. {1} [{2}]</h2><p><b>ORR portal:</b> <a href="{3}">{3}</a><br><b>Evaluation:</b> <a href="{4}">{4}</a><br><b>Matching:</b> {5}</p>{6}</section>' -f ($groupIndex + 1), ([System.Net.WebUtility]::HtmlEncode([string]$first.Application)), ([System.Net.WebUtility]::HtmlEncode([string]$first.BitId)), $first.ApplicationUrl, $first.EvaluationUrl, $group.Count, $checkpointHtml
                continue
            }
            $validCheckpoints = @($group.Group | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Checkpoint) } | Sort-Object Checkpoint, CheckpointVersion, Result, Status, NormalizedRemark -Unique)
            $checkpointBlocks = @(for ($cIdx = 0; $cIdx -lt $validCheckpoints.Count; $cIdx++) {
                $_ = $validCheckpoints[$cIdx]
                $url = [System.Net.WebUtility]::HtmlEncode(([string]$_.CheckpointUrl))
                $id = [System.Net.WebUtility]::HtmlEncode(([string]$_.Checkpoint))
                $name = [System.Net.WebUtility]::HtmlEncode(([string]$_.Name))
                $remark = [System.Net.WebUtility]::HtmlEncode((Get-DisplayRemark -Remark $_.NormalizedRemark -BitId $_.BitId))
                $version = [System.Net.WebUtility]::HtmlEncode(([string]$_.CheckpointVersion))
                $result = [System.Net.WebUtility]::HtmlEncode(([string]$_.Result))
                $chipClass = switch -Regex ([string]$_.Status) {
                    'Maintenance' { 'maintenance'; break }
                    'Exception' { 'exception'; break }
                    'Preview' { 'preview'; break }
                    default { 'live' }
                }
                $lifecycleColor = switch ($chipClass) { 'preview' { '#666666c9' }; 'exception' { '#b80808b4' }; 'maintenance' { '#b36b00b9' }; default { '#4da3ffc2' } }
                $titleText = "$id ($version) - $name"
                $remarkHtml = if ([string]::IsNullOrWhiteSpace($remark)) { '' } else { "<br>$remark" }
                $exNum = [string]$_.ExceptionNumber
                $statusLabel = if (-not [string]::IsNullOrWhiteSpace($exNum) -and ([string]$_.Status) -match 'Exception') { "Exception - $([System.Net.WebUtility]::HtmlEncode($exNum))" } else { [System.Net.WebUtility]::HtmlEncode([string]$_.Status) }
                "<font color=`"$lifecycleColor`"><b><i>[ $statusLabel ]</i></b></font><br><a href=`"$url`" target=`"_blank`"><b>$titleText</b></a>$remarkHtml"
            })
            $rowCount = $checkpointBlocks.Count
            $cells = @(
                ($groupIndex + 1),
                $first.BitId,
                $first.Application,
                $first.GeneratedAtUtc,
                $first.Score
            )
            $encodedCells = $cells | ForEach-Object { [System.Net.WebUtility]::HtmlEncode(([string]$_)) }

            if ($rowCount -eq 0) {
                '<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>(none)</td></tr>' -f $encodedCells[0], $encodedCells[1], $encodedCells[2], $encodedCells[3], $encodedCells[4]
            }
            elseif ($rowCount -eq 1) {
                '<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td></tr>' -f $encodedCells[0], $encodedCells[1], $encodedCells[2], $encodedCells[3], $encodedCells[4], $checkpointBlocks[0]
            }
            else {
                $rowList = [System.Collections.Generic.List[string]]::new()
                $rowList.Add(('<tr><td rowspan="{0}">{1}</td><td rowspan="{0}">{2}</td><td rowspan="{0}">{3}</td><td rowspan="{0}">{4}</td><td rowspan="{0}">{5}</td><td>{6}</td></tr>' -f $rowCount, $encodedCells[0], $encodedCells[1], $encodedCells[2], $encodedCells[3], $encodedCells[4], $checkpointBlocks[0]))
                for ($cIdx = 1; $cIdx -lt $rowCount; $cIdx++) {
                    $rowList.Add(('<tr><td>{0}</td></tr>' -f $checkpointBlocks[$cIdx]))
                }
                $rowList -join "`n"
            }
        }
        return ($sections -join "`n")
    }

    $htmlRows = foreach ($groupIndex in 0..($groups.Count - 1)) {
        $group = $groups[$groupIndex]
        $first = $group.Group[0]
        $uniqueCheckpoints = @($group.Group | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Checkpoint) } | Sort-Object Checkpoint, CheckpointVersion, Result, Status, NormalizedRemark -Unique)
        $rowCount = $uniqueCheckpoints.Count

        $checkpointBlocks = @(for ($cIdx = 0; $cIdx -lt $rowCount; $cIdx++) {
            $_ = $uniqueCheckpoints[$cIdx]
            $url = [System.Net.WebUtility]::HtmlEncode(([string]$_.CheckpointUrl))
            $id = [System.Net.WebUtility]::HtmlEncode(([string]$_.Checkpoint))
            $name = [System.Net.WebUtility]::HtmlEncode(([string]$_.Name))
            $remark = [System.Net.WebUtility]::HtmlEncode((Get-DisplayRemark -Remark $_.NormalizedRemark -BitId $_.BitId))
            $version = [System.Net.WebUtility]::HtmlEncode(([string]$_.CheckpointVersion))
            $result = [System.Net.WebUtility]::HtmlEncode(([string]$_.Result))
            $chipClass = switch -Regex ([string]$_.Status) {
                'Maintenance' { 'maintenance'; break }
                'Exception' { 'exception'; break }
                'Preview' { 'preview'; break }
                default { 'live' }
            }
            $lifecycleColor = switch ($chipClass) { 'preview' { '#666666' }; 'exception' { '#c00000' }; 'maintenance' { '#b36b00' }; default { '#4da3ff' } }
            $titleText = "$id ($version) - $name"
            $remarkHtml = if ([string]::IsNullOrWhiteSpace($remark)) { '' } else { "<br>$remark" }
            $exNum = [string]$_.ExceptionNumber
            $statusLabel = if (-not [string]::IsNullOrWhiteSpace($exNum) -and ([string]$_.Status) -match 'Exception') { "Exception - $([System.Net.WebUtility]::HtmlEncode($exNum))" } else { [System.Net.WebUtility]::HtmlEncode([string]$_.Status) }
            "<font color=`"$lifecycleColor`"><b><i>[ $statusLabel ]</i></b></font><br><a href=`"$url`" target=`"_blank`"><b>$titleText</b></a>$remarkHtml"
        })

        $cells = @(
            ($groupIndex + 1),
            $first.BitId,
            $first.Application,
            $first.GeneratedAtUtc,
            $first.Score
        )
        $encodedCells = $cells | ForEach-Object { [System.Net.WebUtility]::HtmlEncode(([string]$_)) }

        if ($rowCount -eq 0) {
            # No renderable checkpoints after de-duplication — skip this app entirely
            continue
        }
        elseif ($rowCount -eq 1) {
            '<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td></tr>' -f $encodedCells[0], $encodedCells[1], $encodedCells[2], $encodedCells[3], $encodedCells[4], $checkpointBlocks[0]
        }
        else {
            $rowList = [System.Collections.Generic.List[string]]::new()
            $rowList.Add(('<tr><td rowspan="{0}">{1}</td><td rowspan="{0}">{2}</td><td rowspan="{0}">{3}</td><td rowspan="{0}">{4}</td><td rowspan="{0}">{5}</td><td>{6}</td></tr>' -f $rowCount, $encodedCells[0], $encodedCells[1], $encodedCells[2], $encodedCells[3], $encodedCells[4], $checkpointBlocks[0]))
            for ($cIdx = 1; $cIdx -lt $rowCount; $cIdx++) {
                $rowList.Add(('<tr><td>{0}</td></tr>' -f $checkpointBlocks[$cIdx]))
            }
            $rowList -join "`n"
        }
    }
    $htmlRows -join "`n"
}

function Export-CheckpointPdf {
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows,

        [Parameter(Mandatory)]
        [string]$Path,

        [ValidateSet('Table', 'Detailed', 'Summary')]
        [string]$ReportMode = 'Table'
    )

    $browser = @(
        'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
        'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
        'C:\Program Files\Google\Chrome\Application\chrome.exe',
        'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $browser) {
        throw 'PDF export requires Microsoft Edge or Google Chrome.'
    }

    $resolvedPath = [System.IO.Path]::GetFullPath($Path)
    $parentDirectory = Split-Path -Parent $resolvedPath
    if ($parentDirectory) {
        New-Item -ItemType Directory -Path $parentDirectory -Force | Out-Null
    }
    $htmlPath = Join-Path $env:TEMP "orr-report-$([guid]::NewGuid()).html"
    $stdoutPath = Join-Path $env:TEMP "orr-report-$([guid]::NewGuid()).stdout.log"
    $stderrPath = Join-Path $env:TEMP "orr-report-$([guid]::NewGuid()).stderr.log"
    $htmlRows = New-CheckpointHtml -Rows $Rows -ReportMode $ReportMode
    $content = if ($ReportMode -in @('Table', 'Summary')) {
        '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%;"><thead><tr><th>No.</th><th>BIT ID</th><th>App Name</th><th>Last Evaluation (UTC)</th><th>Score</th><th>Failed Checkpoints</th></tr></thead><tbody>' + $htmlRows + '</tbody></table>'
    } else { $htmlRows }
    $html = @"
<!doctype html>
<html><head><meta charset="utf-8"><title>ORR Checkpoint Report</title>
<style>body{font-family:Arial,sans-serif;font-size:10px;color:#000}h1{color:#000}table{border-collapse:collapse;width:100%;table-layout:fixed;background:transparent}th,td{border:1px solid #777;padding:6px;vertical-align:top;color:#000;overflow-wrap:anywhere}th{font-weight:bold;text-align:left}a{color:#0563c1;text-decoration:underline;font-weight:bold}.checkpoint-item{margin-bottom:6px}.lifecycle{font-weight:bold;font-style:italic;margin-bottom:2px}.link{margin-bottom:2px}.remark{color:#333;margin-top:2px}.detail{margin-bottom:12px;padding:8px;background:#f9f9f9;border-left:3px solid #0563c1}.checkpoint-divider{border:0;border-top:1px solid #bbb;margin:6px 0}.chip{display:inline-block;border-radius:10px;padding:2px 7px;font-size:9px;font-weight:bold;color:#000}.preview{background:#d9d9d9}.exception{background:#f4aaaa}.live{background:#a8c7e8}.maintenance{background:#f6b33f}@page{size:landscape;margin:12mm}</style>
</head><body><h1>ORR $ReportMode Report</h1>$content</body></html>
"@

    try {
        Set-Content -LiteralPath $htmlPath -Value $html -Encoding utf8
        $arguments = "--headless --disable-gpu --no-first-run --print-to-pdf=`"$resolvedPath`" `"$htmlPath`""
        $process = Start-Process -FilePath $browser -ArgumentList $arguments -Wait -PassThru -NoNewWindow -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        if ($process.ExitCode -ne 0 -or -not (Test-Path $resolvedPath)) {
            throw "Browser PDF export failed with exit code $($process.ExitCode)."
        }
        Write-Host "PDF exported to $resolvedPath" -ForegroundColor Cyan
    }
    finally {
        Remove-Item -LiteralPath $htmlPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
    }
}

function Export-CheckpointHtml {
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows,

        [Parameter(Mandatory)]
        [string]$Path,

        [ValidateSet('Table', 'Detailed', 'Summary')]
        [string]$ReportMode = 'Table'
    )

    $resolvedPath = [System.IO.Path]::GetFullPath($Path)
    $parentDirectory = Split-Path -Parent $resolvedPath
    if ($parentDirectory) {
        New-Item -ItemType Directory -Path $parentDirectory -Force | Out-Null
    }
    $htmlRows = New-CheckpointHtml -Rows $Rows -ReportMode $ReportMode
    $content = if ($ReportMode -in @('Table', 'Summary')) {
        '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%;"><thead><tr><th>No.</th><th>BIT ID</th><th>App Name</th><th>Last Evaluation (UTC)</th><th>Score</th><th>Failed Checkpoints</th></tr></thead><tbody>' + $htmlRows + '</tbody></table>'
    } else { $htmlRows }
    $html = @"
<!doctype html>
<html><head><meta charset="utf-8"><title>ORR Checkpoint Report</title>
<style>body{font-family:Arial,sans-serif;font-size:10px;color:#000}h1{color:#000}table{border-collapse:collapse;width:100%;table-layout:fixed;background:transparent}th,td{border:1px solid #777;padding:6px;vertical-align:top;color:#000;overflow-wrap:anywhere}th{font-weight:bold;text-align:left}a{color:#0563c1;text-decoration:underline;font-weight:bold}.checkpoint-item{margin-bottom:6px}.lifecycle{font-weight:bold;font-style:italic;margin-bottom:2px}.link{margin-bottom:2px}.remark{color:#333;margin-top:2px}.detail{margin-bottom:12px;padding:8px;background:#f9f9f9;border-left:3px solid #0563c1}.checkpoint-divider{border:0;border-top:1px solid #bbb;margin:6px 0}.chip{display:inline-block;border-radius:10px;padding:2px 7px;font-size:9px;font-weight:bold;color:#000}.preview{background:#d9d9d9}.exception{background:#f4aaaa}.live{background:#a8c7e8}.maintenance{background:#f6b33f}@page{size:landscape;margin:12mm}</style>
</head><body><h1>ORR $ReportMode Report</h1>$content</body></html>
"@

    Set-Content -LiteralPath $resolvedPath -Value $html -Encoding utf8
    Write-Host "HTML report exported to $resolvedPath" -ForegroundColor Cyan
    Start-Process -FilePath $resolvedPath
}

function Read-OptionalBoolean {
    param(
        [Parameter(Mandatory)]
        [string]$Prompt,

        [Parameter(Mandatory)]
        [bool]$Default
    )

    $choice = Read-MenuSelection -Prompt $Prompt -Options @('Yes', 'No') -Default @($(if ($Default) { 'Yes' } else { 'No' }))
    return [bool]($choice -eq 'Yes')
}

if ([string]::IsNullOrWhiteSpace($OutputMode)) {
    $OutputMode = Read-MenuSelection -Prompt 'Output format' -Options @('Table', 'Detailed', 'Summary') -Default @('Table')
}

if ($PSBoundParameters.ContainsKey('ExportFormat')) {
    $ExportPdf = $true
}
if (-not $PSBoundParameters.ContainsKey('ExportPdf') -and -not $PSBoundParameters.ContainsKey('ExportFormat')) {
    $ExportPdf = Read-OptionalBoolean -Prompt 'Export the report?' -Default $false
}
if ($ExportPdf -and [string]::IsNullOrWhiteSpace($ExportFormat)) {
    $exportFormatSelection = Read-MenuSelection -Prompt 'Export format' -Options @('HTML page', 'PDF') -Default @('HTML page')
    $ExportFormat = if ($exportFormatSelection -eq 'HTML page') { 'HTML' } else { 'PDF' }
}
if ($ExportPdf -and [string]::IsNullOrWhiteSpace($PdfPath)) {
    $extension = if ($ExportFormat -eq 'HTML') { 'html' } else { 'pdf' }
    $PdfPath = Join-Path $reportsDirectory ("ORR-Checkpoint-Report-{0}.{1}" -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $extension)
}
if ([string]::IsNullOrWhiteSpace($BitId) -and -not $AllMatching) {
    $selection = Read-MenuSelection -Prompt 'Retrieve evaluations for' -Options @('One BIT ID', 'Multiple BIT IDs') -Default @('One BIT ID')
    if ($selection -eq 'One BIT ID') {
        $BitId = (Read-Host 'Enter BIT ID').Trim()
    }
    else {
        $AllMatching = $true
    }
}

if ($AllMatching -and -not [string]::IsNullOrWhiteSpace($BitId)) {
    throw 'Use either -BitId for one application or -AllMatching for filter-based retrieval, not both.'
}

if (-not $AllMatching -and $BitId -notmatch '^[A-Za-z0-9]+$') {
    throw "Invalid BIT ID '$BitId'. Enter only letters and numbers."
}

$validEnvironments = @(
    'Production',
    'DisasterRecovery',
    'Development',
    'Integration',
    'LoadTesting',
    'NonProduction',
    'QualityAssurance',
    'Staging',
    'Sandbox',
    'UserAcceptanceTesting',
    'All'
)
if ([string]::IsNullOrWhiteSpace($Environment)) {
    $environmentDefault = if ($AllMatching) { @('Production') } else { @('Production') }
    $Environment = @(Read-MenuSelection -Prompt 'Environment' -Options $validEnvironments -Default $environmentDefault -MultiSelect) -join ','
}

$selectedEnvironments = @($Environment -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
if ($selectedEnvironments.Count -eq 0 -or @($selectedEnvironments | Where-Object { $_ -notin $validEnvironments }).Count -gt 0) {
    throw "Invalid environment '$Environment'. Valid values: $($validEnvironments -join ', ')."
}

if (-not $PSBoundParameters.ContainsKey('Division')) {
    $Division = Read-MenuSelection -Prompt 'Division' -Options @('ALL', 'TAA', 'HLT', 'FCC', 'CORP', 'CPESG', 'DXG', 'LNR', 'GBS', 'GGM', 'UNKNOWN') -Default @('ALL')
}
if ($null -ne $Division -and ([string]$Division).Trim().ToUpperInvariant() -eq 'ALL') {
    $Division = ''
}

if ($AllMatching) {
    if (-not $PSBoundParameters.ContainsKey('IpmOnly')) {
        $IpmOnly = Read-OptionalBoolean -Prompt 'Show IPM managed applications only?' -Default $false
    }
    if (-not $PSBoundParameters.ContainsKey('EvaluatedOnly')) {
        $EvaluatedOnly = Read-OptionalBoolean -Prompt 'Show evaluated applications only?' -Default $false
    }
    if (-not $PSBoundParameters.ContainsKey('ActiveOnly')) {
        $ActiveOnly = Read-OptionalBoolean -Prompt 'Show active applications only?' -Default $true
    }
}

if (-not $PSBoundParameters.ContainsKey('Live')) {
    $Live = Read-OptionalBoolean -Prompt 'Show live checkpoints?' -Default $true
}
if (-not $PSBoundParameters.ContainsKey('Preview')) {
    $Preview = Read-OptionalBoolean -Prompt 'Show preview checkpoints?' -Default $true
}
if (-not $PSBoundParameters.ContainsKey('Exception')) {
    $Exception = Read-OptionalBoolean -Prompt 'Show exception checkpoints?' -Default $true
}
if (-not $PSBoundParameters.ContainsKey('Maintenance')) {
    $Maintenance = Read-OptionalBoolean -Prompt 'Show maintenance checkpoints?' -Default $true
}
if (-not $PSBoundParameters.ContainsKey('Failing')) {
    $failingChoice = Read-MenuSelection -Prompt 'Checkpoints to include' -Options @('Failing checkpoints', 'Non-failing checkpoints', 'All checkpoints') -Default @('Failing checkpoints')
    $Failing = switch ($failingChoice) {
        'Failing checkpoints' { $true }
        'Non-failing checkpoints' { $false }
        'All checkpoints' { $null }
    }
}

Write-Host "ORR Score & Checkpoint Report v$scriptVersion" -ForegroundColor Cyan
Write-Host '----------------------------------' -ForegroundColor DarkCyan
Write-Field -Label 'Output' -Value $OutputMode
Write-Field -Label 'Mode' -Value $(if ($AllMatching) { 'Multiple BIT IDs' } else { "Single BIT ID ($BitId)" })
Write-Field -Label 'Environment' -Value ($selectedEnvironments -join ', ')
Write-Field -Label 'Division' -Value $(if ($Division) { $Division } else { 'All' })
if ($AllMatching) {
    Write-Field -Label 'IPM only' -Value $IpmOnly
    Write-Field -Label 'Evaluated only' -Value $EvaluatedOnly
    Write-Field -Label 'Active only' -Value $ActiveOnly
}
Write-Field -Label 'Live' -Value $Live
Write-Field -Label 'Preview' -Value $Preview
Write-Field -Label 'Exception' -Value $Exception
Write-Field -Label 'Maintenance' -Value $Maintenance
Write-Field -Label 'Failing' -Value $(if ($null -eq $Failing) { 'All' } else { $Failing })

if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
    throw 'Az.Accounts is not installed. Install it with: Install-Module Az.Accounts -Scope CurrentUser'
}

Import-Module Az.Accounts

function ConvertTo-PlainTextToken {
    param([Parameter(Mandatory)]$Token)

    if ($Token -isnot [securestring]) {
        return [string]$Token
    }

    $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Token)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
    }
}

try {
    Write-Host "Getting ORR access token..." -ForegroundColor Cyan
    $accessToken = Get-AzAccessToken -ResourceUrl $resourceUrl
    $plainToken = ConvertTo-PlainTextToken -Token $accessToken.Token
    $headers = @{
        Authorization = "Bearer $plainToken"
        Accept        = 'application/json'
    }

    $queryParameters = [ordered]@{
        environments = $selectedEnvironments -join ','
        latestOnly   = 'true'
        limit        = '100'
        includeCount = 'true'
    }
    if (-not $AllMatching) {
        $queryParameters.bitIds = $BitId
    }
    if (-not [string]::IsNullOrWhiteSpace($Division)) {
        $queryParameters.division = $Division
    }
    if ($null -ne $IpmOnly) {
        $queryParameters.ipmOnly = $IpmOnly.ToString().ToLowerInvariant()
    }
    if ($null -ne $EvaluatedOnly) {
        $queryParameters.evaluatedOnly = $EvaluatedOnly.ToString().ToLowerInvariant()
    }
    if ($null -ne $ActiveOnly) {
        $queryParameters.activeOnly = $ActiveOnly.ToString().ToLowerInvariant()
    }

    $summaries = [System.Collections.Generic.List[object]]::new()
    $continuationToken = $null
    do {
        if ($continuationToken) {
            $queryParameters.continuationToken = $continuationToken
        }
        elseif ($queryParameters.Contains('continuationToken')) {
            $queryParameters.Remove('continuationToken')
        }

        $queryString = ($queryParameters.GetEnumerator() | ForEach-Object {
            '{0}={1}' -f [uri]::EscapeDataString($_.Key), [uri]::EscapeDataString([string]$_.Value)
        }) -join '&'
        $summaryUrl = "$apiBaseUrl/evaluations?$queryString"

        $targetLabel = if ($AllMatching) { 'all matching BIT IDs' } else { $BitId }
        Write-Host "Loading latest $Environment evaluations for $targetLabel..." -ForegroundColor Cyan
        $summary = Invoke-RestMethod -Uri $summaryUrl -Headers $headers -TimeoutSec 60
        foreach ($item in @($summary.items)) {
            $summaries.Add($item)
        }
        $continuationToken = $summary.continuationToken
    } while (-not [string]::IsNullOrWhiteSpace($continuationToken))

    $latestEvaluations = @(
        $summaries |
            Where-Object { ($selectedEnvironments -contains 'All') -or @($_.environmentNames | Where-Object { $_ -in $selectedEnvironments }).Count -gt 0 } |
            Group-Object { ([string]$_.id -split '!', 2)[0] } |
            ForEach-Object {
                $_.Group | Sort-Object { [datetime]$_.generatedAt } -Descending | Select-Object -First 1
            } |
            Sort-Object applicationName
    )

    if ($latestEvaluations.Count -eq 0) {
        $description = if ($AllMatching) { 'the selected filters' } else { "BIT ID $BitId" }
        Write-Warning "No $Environment evaluation was found for $description."
        exit 2
    }

    $results = foreach ($latest in $latestEvaluations) {
        $currentBitId = ([string]$latest.id -split '!', 2)[0]
        $encodedEvaluationId = [uri]::EscapeDataString([string]$latest.id).Replace('!', '%21')
        $detailUrl = "$apiBaseUrl/evaluations/$encodedEvaluationId"
        $detail = Invoke-RestMethod -Uri $detailUrl -Headers $headers -TimeoutSec 60

        if (-not $detail.checkpoints) {
            throw "Evaluation '$($latest.id)' returned no checkpoint data."
        }

        $failedCheckpoints = @($detail.checkpoints | Where-Object { $_.status -eq 'FAIL' })
        $maintenanceCheckpoints = @($detail.checkpoints | Where-Object { $_.metadata.isMaintenance })
        $scoredCheckpoints = @($detail.checkpoints | Where-Object { $_.metadata.isScored })
        $matchingCheckpoints = @(
            $detail.checkpoints | Where-Object {
                $isPreview = $_.metadata.unscoredStatusDetails.unscoredType -eq 'preview'
                $hasException = [bool]$_.hasException
                $isMaintenance = [bool]$_.metadata.isMaintenance
                $isFailing = $_.status -eq 'FAIL'
                $isLive = -not ($isPreview -or $hasException -or $isMaintenance)
                $hasSelectedLifecycleFilter = $Live -or $Preview -or $Exception -or $Maintenance
                $matchesLifecycleFilter = if ($hasSelectedLifecycleFilter) {
                    ($Live -and $isLive) -or
                    ($Preview -and $isPreview) -or
                        ($Exception -and $hasException) -or
                        ($Maintenance -and $isMaintenance)
                }
                else {
                    $false
                }

                $matchesLifecycleFilter -and ($null -eq $Failing -or $isFailing -eq $Failing)
            }
        )

        [pscustomobject]@{
            BitId                  = $currentBitId
            Application            = $latest.applicationName
            Environment            = ($latest.environmentNames -join ', ')
            EvaluationId           = $latest.id
            GeneratedAtUtc         = ([datetime]$latest.generatedAt).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
            OverallScore           = $detail.score
            InclusiveScore         = $latest.inclusiveScore
            TotalCheckpoints       = $detail.checkpoints.Count
            ScoredCheckpoints      = $scoredCheckpoints.Count
            FailedCheckpointCount  = $failedCheckpoints.Count
            MatchingCheckpointCount = $matchingCheckpoints.Count
            Maintenance            = $maintenanceCheckpoints.Count
            MatchingCheckpointData = $matchingCheckpoints
        }
    }

    $matchedResults = @($results | Where-Object { $_.MatchingCheckpointCount -gt 0 })

    $detailedExportRows = @()
    if ($OutputMode -eq 'Detailed' -and $ExportPdf) {
        $detailedRowNumber = 0
        $detailedExportRows = @(
            foreach ($result in $matchedResults) {
                foreach ($checkpoint in $result.MatchingCheckpointData) {
                    $detailedRowNumber++
                    $states = [System.Collections.Generic.List[string]]::new()
                    if ($checkpoint.metadata.isMaintenance) { $states.Add('Maintenance') }
                    if ($checkpoint.hasException) { $states.Add('Exception') }
                    if ($checkpoint.metadata.unscoredStatusDetails.unscoredType -eq 'preview') { $states.Add('Preview') }
                    if ($states.Count -eq 0) { $states.Add('Live') }
                    $normalizedRemark = ([string]$checkpoint.remarks).Trim() -replace '(?i)\b041\s*800[A-Z0-9]{6}\b', '<BIT_ID>' -replace '\s+', ' '

                    [pscustomobject]@{
                        Number = $detailedRowNumber
                        BitId = $result.BitId
                        Application = $result.Application
                        ApplicationUrl = "https://orr.pegasus.wkcloud.io/applications/$($result.BitId)?env=$($selectedEnvironments[0])"
                        EvaluationUrl = "https://orr.pegasus.wkcloud.io/evaluations/$([uri]::EscapeDataString([string]$result.EvaluationId))"
                        GeneratedAtUtc = $result.GeneratedAtUtc
                        Score = Format-Score $result.OverallScore
                        Checkpoint = $checkpoint.metadata.id
                        CheckpointVersion = $checkpoint.metadata.checkpointVersion
                        CheckpointUrl = "https://orr.pegasus.wkcloud.io/evaluations/$([uri]::EscapeDataString([string]$result.EvaluationId))?checkpoint=$([uri]::EscapeDataString([string]$checkpoint.metadata.id))&checkpointVersion=$([uri]::EscapeDataString([string]$checkpoint.metadata.checkpointVersion))"
                        Name = $checkpoint.metadata.displayName
                        Result = $checkpoint.status
                        Status = $states -join ' + '
                        NormalizedRemark = $normalizedRemark
                        IsFailing = ($checkpoint.status -eq 'FAIL')
                        IsPreview = ($checkpoint.metadata.unscoredStatusDetails.unscoredType -eq 'preview')
                        HasException = [bool]$checkpoint.hasException
                        IsScored = $checkpoint.metadata.isScored
                        IsMaintenance = $checkpoint.metadata.isMaintenance
                        ExceptionNumber = $checkpoint.exceptionDetails.number
                        StateReason = $checkpoint.metadata.unscoredStatusDetails.reason
                        StateEndDate = $checkpoint.metadata.unscoredStatusDetails.endDate
                        Accountable = $checkpoint.metadata.accountable
                        Systems = (($checkpoint.items | Where-Object { $_.system } | Select-Object -ExpandProperty system -Unique) -join ', ')
                        OriginalRemarks = $checkpoint.remarks
                        References = (($checkpoint.items | Where-Object { $_.reference } | Select-Object -ExpandProperty reference -Unique) -join [Environment]::NewLine)
                    }
                }
            }
        )
    }

    if ($OutputMode -in @('Table', 'Summary')) {
        $summaryRowNumber = 0
        $summaryRows = @(
            foreach ($result in $matchedResults) {
                foreach ($checkpoint in $result.MatchingCheckpointData) {
                    $summaryRowNumber++
                    $states = [System.Collections.Generic.List[string]]::new()
                    if ($checkpoint.metadata.isMaintenance) {
                        $states.Add('Maintenance')
                    }
                    if ($checkpoint.hasException) {
                        $states.Add('Exception')
                    }
                    if ($checkpoint.metadata.unscoredStatusDetails.unscoredType -eq 'preview') {
                        $states.Add('Preview')
                    }
                    if ($states.Count -eq 0) {
                        $states.Add('Live')
                    }
                    $normalizedRemark = ([string]$checkpoint.remarks).Trim()
                    $normalizedRemark = $normalizedRemark -replace '(?i)\b041\s*800[A-Z0-9]{6}\b', '<BIT_ID>'
                    $normalizedRemark = $normalizedRemark -replace '\s+', ' '

                    [pscustomobject]@{
                        Number      = $summaryRowNumber
                        BitId       = $result.BitId
                        Application = $result.Application
                        ApplicationUrl = "https://orr.pegasus.wkcloud.io/applications/$($result.BitId)?env=$($selectedEnvironments[0])"
                        GeneratedAtUtc = $result.GeneratedAtUtc
                        Score       = Format-Score $result.OverallScore
                        Checkpoint  = $checkpoint.metadata.id
                            CheckpointUrl = "https://orr.pegasus.wkcloud.io/evaluations/$([uri]::EscapeDataString([string]$result.EvaluationId))?checkpoint=$([uri]::EscapeDataString([string]$checkpoint.metadata.id))&checkpointVersion=$([uri]::EscapeDataString([string]$checkpoint.metadata.checkpointVersion))"
                        CheckpointVersion = $checkpoint.metadata.checkpointVersion
                        Name        = $checkpoint.metadata.displayName
                        Result      = $checkpoint.status
                        Status      = $states -join ' + '
                        ExceptionNumber = [string]$checkpoint.exceptionDetails.number
                        Remarks     = $checkpoint.remarks
                        NormalizedRemark = $normalizedRemark
                        GroupKey    = "$($checkpoint.metadata.id)|$($checkpoint.status)|$($states -join ' + ')|$normalizedRemark"
                    }
                }
            }
        )

        if ($summaryRows.Count -eq 0) {
            Write-Section -Title 'RESULT' -Color Yellow
            Write-Host 'No checkpoints matched the selected filters.' -ForegroundColor Yellow
            exit 0
        }

        if ($OutputMode -eq 'Table') {
            Write-Section -Title "CHECKPOINT TABLE ($($summaryRows.Count) row(s))" -Color $(if ($Failing -eq $true) { 'Red' } else { 'Cyan' })
            Write-CheckpointTable -Rows $summaryRows
            if ($ExportPdf) {
                if ($ExportFormat -eq 'HTML') {
                    Export-CheckpointHtml -Rows $summaryRows -Path $PdfPath -ReportMode 'Table'
                }
                else {
                    Export-CheckpointPdf -Rows $summaryRows -Path $PdfPath -ReportMode 'Table'
                }
            }
            exit 0
        }

        $resultLabel = if ($Failing -eq $true) { 'FAILING' } elseif ($Failing -eq $false) { 'NON-FAILING' } else { 'ALL' }
        $summaryGroups = @(
            $summaryRows |
                Group-Object GroupKey |
                Sort-Object @{ Expression = { $_.Group[0].Checkpoint }; Ascending = $true },
                    @{ Expression = { $_.Count }; Descending = $true }
        )
        Write-Section -Title "$resultLabel CHECKPOINT SUMMARY ($($summaryRows.Count) occurrence(s), $($summaryGroups.Count) group(s), $((@($summaryRows.BitId | Sort-Object -Unique)).Count) BIT ID(s))" -Color $(if ($Failing) { 'Red' } else { 'Cyan' })

        $groupNumber = 0
        foreach ($group in $summaryGroups) {
            $groupNumber++
            $first = $group.Group[0]
            Write-Host ("[{0}] {1} - {2}" -f $groupNumber, $first.Checkpoint, $first.Name) -ForegroundColor Yellow
            Write-Field -Label 'Result' -Value $first.Result -ValueColor $(if ($first.Result -eq 'FAIL') { 'Red' } else { 'Cyan' })
            Write-Field -Label 'Lifecycle' -Value $first.Status -ValueColor Cyan
            Write-Field -Label 'Occurrences' -Value $group.Count
            Write-Field -Label 'Remarks' -Value (Get-DisplayRemark -Remark $first.NormalizedRemark -BitId $first.BitId)
            Write-Host ''
            Write-Host '  Affected applications' -ForegroundColor DarkCyan
            Write-Host ('  {0,-4}{1,-14}{2}' -f '#', 'BIT ID', 'Application') -ForegroundColor DarkGray
            Write-Host ('  {0,-4}{1,-14}{2}' -f '-', '------', '-----------') -ForegroundColor DarkGray

            $affectedNumber = 0
            $group.Group |
                Sort-Object BitId, Application -Unique |
                ForEach-Object {
                    $affectedNumber++
                    Write-Host ('  {0,-4}{1,-14}{2}' -f $affectedNumber, $_.BitId, $_.Application)
                }
            Write-Host ''
        }
        if ($ExportPdf) {
            if ($ExportFormat -eq 'HTML') {
                Export-CheckpointHtml -Rows $summaryRows -Path $PdfPath -ReportMode 'Summary'
            }
            else {
                Export-CheckpointPdf -Rows $summaryRows -Path $PdfPath -ReportMode 'Summary'
            }
        }
        exit 0
    }

    foreach ($result in $matchedResults) {
        Write-Section -Title "$($result.Application) [$($result.BitId)]" -Color Cyan
        Write-Field -Label 'ORR portal' -Value "https://orr.pegasus.wkcloud.io/applications/$($result.BitId)`?env=$Environment" -ValueColor Blue
        Write-Field -Label 'Evaluation' -Value "https://orr.pegasus.wkcloud.io/evaluations/$($result.EvaluationId)" -ValueColor Blue
        Write-Field -Label 'Matching' -Value $result.MatchingCheckpointCount -ValueColor Yellow

        $checkpointNumber = 0
        foreach ($checkpoint in $result.MatchingCheckpointData) {
            $checkpointNumber++
            $references = @(
                $checkpoint.items |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_.reference) } |
                    Select-Object -ExpandProperty reference -Unique
            )
            $systems = @(
                $checkpoint.items |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_.system) } |
                    Select-Object -ExpandProperty system -Unique
            )

            Write-Host ''
            Write-Host ("  [{0}] {1} - {2}" -f $checkpointNumber, $checkpoint.metadata.id, $checkpoint.metadata.displayName) -ForegroundColor Yellow
            Write-Field -Label 'Version' -Value $checkpoint.metadata.checkpointVersion
            Write-Field -Label 'Status' -Value $checkpoint.status -ValueColor $(if ($checkpoint.status -eq 'FAIL') { 'Red' } else { 'Cyan' })
            Write-Field -Label 'Is failing' -Value ($checkpoint.status -eq 'FAIL')
            Write-Field -Label 'Is preview' -Value ($checkpoint.metadata.unscoredStatusDetails.unscoredType -eq 'preview')
            Write-Field -Label 'Has exception' -Value ([bool]$checkpoint.hasException)
            Write-Field -Label 'Is scored' -Value $checkpoint.metadata.isScored
            Write-Field -Label 'Is maintenance' -Value $checkpoint.metadata.isMaintenance
            Write-Field -Label 'Exception number' -Value $checkpoint.exceptionDetails.number
            Write-Field -Label 'State reason' -Value $checkpoint.metadata.unscoredStatusDetails.reason
            Write-Field -Label 'State end date' -Value $checkpoint.metadata.unscoredStatusDetails.endDate
            Write-Field -Label 'Accountable' -Value $checkpoint.metadata.accountable
            Write-Field -Label 'System' -Value ($systems -join ', ')
            Write-Field -Label 'Remarks' -Value $checkpoint.remarks
            if ($references.Count -eq 0) {
                Write-Field -Label 'References' -Value $null
            }
            else {
                Write-Field -Label 'References' -Value ($references -join [Environment]::NewLine) -ValueColor Blue
            }
        }
    }

    if ($detailedExportRows.Count -gt 0) {
        if ($ExportFormat -eq 'HTML') {
            Export-CheckpointHtml -Rows $detailedExportRows -Path $PdfPath -ReportMode 'Detailed'
        }
        else {
            Export-CheckpointPdf -Rows $detailedExportRows -Path $PdfPath -ReportMode 'Detailed'
        }
    }

    if ($matchedResults.Count -eq 0) {
        Write-Section -Title 'RESULT' -Color Yellow
        Write-Host 'No checkpoints matched the selected filters.' -ForegroundColor Yellow
        exit 0
    }

    Write-Section -Title "CONSOLIDATED SUMMARY ($($matchedResults.Count) BIT ID(s))" -Color Cyan
    $summaryNumber = 0
    foreach ($result in $matchedResults) {
        $summaryNumber++
        Write-Host ("[{0}] {1} [{2}]" -f $summaryNumber, $result.Application, $result.BitId) -ForegroundColor Cyan
        Write-Field -Label 'Environment' -Value $result.Environment
        Write-Field -Label 'Overall score' -Value (Format-Score $result.OverallScore)
        Write-Field -Label 'Inclusive score' -Value $result.InclusiveScore
        Write-Field -Label 'Total checkpoints' -Value $result.TotalCheckpoints
        Write-Field -Label 'Scored checkpoints' -Value $result.ScoredCheckpoints
        Write-Field -Label 'Matched checkpoints' -Value $result.MatchingCheckpointCount
        Write-Field -Label 'Total failures' -Value $result.FailedCheckpointCount
        Write-Field -Label 'Maintenance' -Value $result.Maintenance
        Write-Field -Label 'Generated UTC' -Value $result.GeneratedAtUtc
        Write-Host ''
    }

}
catch {
    Write-Error "Unable to retrieve ORR score: $($_.Exception.Message)"
    exit 1
}
