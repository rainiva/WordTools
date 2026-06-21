# Compares wordtools-benchmark.csv rows against a locked perf baseline.
# Exit 0 = within thresholds; exit 1 = regression or missing data.

[CmdletBinding(DefaultParameterSetName = 'Compare')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Compare')]
    [string]$BaselineCsv,

    [Parameter(Mandatory = $true, ParameterSetName = 'Compare')]
    [string]$CurrentCsv,

    [Parameter(ParameterSetName = 'Compare')]
    [ValidateSet('PERF-01', 'PERF-02', 'PERF-03', 'All')]
    [string]$ScenarioId = 'All',

    [Parameter(ParameterSetName = 'Compare')]
    [double]$MaxTotalSecondsRatio = 1.05,

    [Parameter(ParameterSetName = 'Compare')]
    [double]$MaxInsertImagesMsRatio = 1.05,

    [Parameter(ParameterSetName = 'Compare')]
    [double]$MaxAddPictureMsRatio = 1.05,

    [Parameter(ParameterSetName = 'Compare')]
    [double]$MaxClearNumberingMsRatio = 1.05,

    [Parameter(ParameterSetName = 'Compare')]
    [double]$MaxCellAvailabilityMsRatio = 1.08,

    [Parameter(ParameterSetName = 'Compare')]
    [double]$MaxProgressUiMsRatio = 1.10,

    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')]
    [switch]$SelfTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir 'lib\BenchmarkCsv.ps1')

function Get-MetricRatioLimit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MetricName
    )

    switch ($MetricName) {
        'total_seconds' { return $MaxTotalSecondsRatio }
        'insert_images_ms' { return $MaxInsertImagesMsRatio }
        'add_picture_ms' { return $MaxAddPictureMsRatio }
        'clear_numbering_ms' { return $MaxClearNumberingMsRatio }
        'cell_availability_ms' { return $MaxCellAvailabilityMsRatio }
        'progress_ui_ms' { return $MaxProgressUiMsRatio }
        default { return $null }
    }
}

function Compare-BenchmarkRowPair {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Label,
        [Parameter(Mandatory = $true)]
        [hashtable]$BaselineRow,
        [Parameter(Mandatory = $true)]
        [hashtable]$CurrentRow
    )

    $failures = New-Object System.Collections.Generic.List[string]
    $reports = New-Object System.Collections.Generic.List[string]

    foreach ($metric in @('total_seconds', 'insert_images_ms', 'add_picture_ms', 'clear_numbering_ms', 'cell_availability_ms', 'progress_ui_ms')) {
        $baseValue = Get-NumericOrNull -Value ([string]$BaselineRow[$metric])
        $currentValue = Get-NumericOrNull -Value ([string]$CurrentRow[$metric])
        $ratioLimit = Get-MetricRatioLimit -MetricName $metric

        if ($null -eq $baseValue -or $null -eq $currentValue) {
            if ($null -ne $baseValue -or $null -ne $currentValue) {
                $failures.Add("$Label $metric missing (baseline=$baseValue current=$currentValue)") | Out-Null
            }
            continue
        }

        if ($baseValue -le 0) {
            $reports.Add("$Label $metric baseline=$baseValue current=$currentValue (skip ratio, baseline non-positive)") | Out-Null
            continue
        }

        $ratio = $currentValue / $baseValue
        $reports.Add(('{0} {1}: baseline={2} current={3} ratio={4:F3} limit={5:F3}' -f $Label, $metric, $baseValue, $currentValue, $ratio, $ratioLimit)) | Out-Null
        if ($ratio -gt $ratioLimit) {
            $failures.Add("$Label $metric ratio $([math]::Round($ratio, 3)) > limit $ratioLimit") | Out-Null
        }
    }

    foreach ($countMetric in $script:BenchmarkExactCountMetrics) {
        $baseCount = Get-IntOrNull -Value ([string]$BaselineRow[$countMetric])
        $currentCount = Get-IntOrNull -Value ([string]$CurrentRow[$countMetric])

        if ($null -eq $baseCount -and $null -eq $currentCount) {
            continue
        }

        $reports.Add(('{0} {1}: baseline={2} current={3}' -f $Label, $countMetric, $baseCount, $currentCount)) | Out-Null
        if ($baseCount -ne $currentCount) {
            $failures.Add("$Label $countMetric changed ($baseCount -> $currentCount)") | Out-Null
        }
    }

    return [PSCustomObject]@{
        Label = $Label
        Reports = $reports.ToArray()
        Failures = $failures.ToArray()
        Passed = ($failures.Count -eq 0)
    }
}

function Invoke-BenchmarkCompare {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaselinePath,
        [Parameter(Mandatory = $true)]
        [string]$CurrentPath,
        [string]$ScenarioFilter = 'All'
    )

    $baseline = Read-BenchmarkCsvFile -Path $BaselinePath
    $current = Read-BenchmarkCsvFile -Path $CurrentPath

    if ($baseline.Rows.Length -eq 0) {
        throw "Baseline CSV has no data rows: $BaselinePath"
    }

    $scenarioRows = @($baseline.Rows)
    if ($ScenarioFilter -ne 'All') {
        $scenarioRows = @($baseline.Rows | Where-Object {
            if ($_.ContainsKey('scenario_id') -and [string]$_.scenario_id -eq $ScenarioFilter) {
                return $true
            }

            $map = $script:BenchmarkScenarioMap[$ScenarioFilter]
            if ($null -eq $map) { return $false }
            return ([string]$_.run_mode -eq $map.RunMode -and (Get-IntOrNull $_.total_files) -eq $map.TotalFiles)
        })
    }

    if ($scenarioRows.Length -eq 0) {
        throw "No baseline rows matched scenario '$ScenarioFilter'."
    }

    $allResults = New-Object System.Collections.Generic.List[object]
    foreach ($baselineRow in $scenarioRows) {
        $label = if ($baselineRow.ContainsKey('scenario_id') -and -not [string]::IsNullOrWhiteSpace([string]$baselineRow.scenario_id)) {
            [string]$baselineRow.scenario_id
        }
        else {
            '{0} (total_files={1})' -f [string]$baselineRow.run_mode, [string]$baselineRow.total_files
        }

        $totalFiles = Get-IntOrNull -Value ([string]$baselineRow.total_files)
        if ($null -eq $totalFiles) { $totalFiles = 0 }

        $currentRow = Get-LatestBenchmarkRow -Rows $current.Rows -RunMode ([string]$baselineRow.run_mode) -TotalFiles $totalFiles -ScenarioId $label
        if ($null -eq $currentRow) {
            $allResults.Add([PSCustomObject]@{
                Label = $label
                Reports = @()
                Failures = @("No matching current row for run_mode=$($baselineRow.run_mode) total_files=$totalFiles")
                Passed = $false
            }) | Out-Null
            continue
        }

        $allResults.Add((Compare-BenchmarkRowPair -Label $label -BaselineRow $baselineRow -CurrentRow $currentRow)) | Out-Null
    }

    return $allResults.ToArray()
}

function Invoke-CompareBenchmarkSelfTest {
    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('wordtools-benchmark-selftest-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

    try {
        $baselinePath = Join-Path $tempRoot 'baseline.csv'
        $passPath = Join-Path $tempRoot 'current-pass.csv'
        $failPath = Join-Path $tempRoot 'current-fail.csv'

        $baseRow = @{
            scenario_id = 'PERF-01'
            run_mode = 'SelectedFiles'
            status = 'Completed'
            total_files = '4'
            cancelled = 'False'
            total_seconds = '10.000'
            insert_images_ms = '8000'
            add_picture_ms = '6000'
            clear_numbering_ms = '100'
            cell_availability_ms = '500'
            progress_ui_ms = '200'
            cell_availability_count = '12'
            add_picture_count = '4'
        }

        $passRow = @{} + $baseRow
        $passRow.total_seconds = '10.200'
        $passRow.insert_images_ms = '8200'
        $passRow.add_picture_ms = '6100'

        $failRow = @{} + $baseRow
        $failRow.total_seconds = '11.000'
        $failRow.add_picture_count = '5'

        Write-BenchmarkCsvFile -Path $baselinePath -Rows @($baseRow) -IncludeScenarioId
        Write-BenchmarkCsvFile -Path $passPath -Rows @($passRow)
        Write-BenchmarkCsvFile -Path $failPath -Rows @($failRow)

        $passResults = Invoke-BenchmarkCompare -BaselinePath $baselinePath -CurrentPath $passPath -ScenarioFilter 'PERF-01'
        if (-not $passResults[0].Passed) {
            throw ('SelfTest pass case failed: ' + ($passResults[0].Failures -join '; '))
        }

        $failResults = Invoke-BenchmarkCompare -BaselinePath $baselinePath -CurrentPath $failPath -ScenarioFilter 'PERF-01'
        if ($failResults[0].Passed) {
            throw 'SelfTest fail case unexpectedly passed.'
        }

        Write-Host 'compare-benchmark.ps1 self-test passed.'
        return 0
    }
    finally {
        if (Test-Path -LiteralPath $tempRoot) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
}

if ($SelfTest) {
    exit (Invoke-CompareBenchmarkSelfTest)
}

if (-not (Test-Path -LiteralPath $BaselineCsv)) {
    Write-Error "Baseline CSV not found: $BaselineCsv"
    exit 2
}

if (-not (Test-Path -LiteralPath $CurrentCsv)) {
    Write-Error "Current CSV not found: $CurrentCsv"
    exit 2
}

$results = @(Invoke-BenchmarkCompare -BaselinePath $BaselineCsv -CurrentPath $CurrentCsv -ScenarioFilter $ScenarioId)
$failed = @($results | Where-Object { -not $_.Passed })

foreach ($result in $results) {
    Write-Host "== $($result.Label) =="
    foreach ($line in $result.Reports) {
        Write-Host "  OK  $line"
    }
    foreach ($line in $result.Failures) {
        Write-Host "  FAIL  $line"
    }
}

if ($failed.Length -gt 0) {
    Write-Host ""
    Write-Host "Benchmark compare FAILED ($($failed.Length) scenario(s))."
    exit 1
}

Write-Host ""
Write-Host "Benchmark compare PASSED ($($results.Length) scenario(s))."
exit 0
