# Aggregates recent wordtools-benchmark.csv runs into a median perf baseline file.

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceCsv,

    [Parameter(Mandatory = $true)]
    [string]$OutputCsv,

    [ValidateSet('PERF-01', 'PERF-02', 'PERF-03', 'All')]
    [string]$ScenarioId = 'All',

    [int]$RunsPerScenario = 3,

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir 'lib\BenchmarkCsv.ps1')

function Get-ScenarioIdsToCapture {
    param([string]$Filter)

    if ($Filter -ne 'All') {
        return ,@($Filter)
    }

    return ,@('PERF-01', 'PERF-02')
}

function Build-MedianBenchmarkRow {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Scenario,
        [Parameter(Mandatory = $true)]
        [array]$Rows
    )

    if ($Rows.Length -eq 0) {
        throw "No eligible rows for scenario $Scenario"
    }

    $template = @{} + $Rows[-1]
    $template['scenario_id'] = $Scenario
    $template['status'] = 'Completed'
    $template['cancelled'] = 'False'
    $template['timestamp_utc'] = (Get-Date).ToUniversalTime().ToString('O', [System.Globalization.CultureInfo]::InvariantCulture)

    $ratioColumns = @('total_seconds', 'initialize_ms', 'clear_numbering_ms', 'calculate_start_number_ms', 'preallocate_rows_ms', 'insert_images_ms', 'wrap_up_ms', 'cell_availability_ms', 'floating_shape_lookup_ms', 'overwrite_clear_ms', 'add_picture_ms', 'cell_validation_ms', 'picture_sizing_ms', 'progress_ui_ms', 'description_write_ms')
    foreach ($column in $ratioColumns) {
        $values = @(
            foreach ($row in $Rows) {
                Get-NumericOrNull -Value ([string]$row[$column])
            }
        ) | Where-Object { $null -ne $_ }
        if ($null -eq $values) { continue }
        $median = Get-Median -Values @([double[]]$values)
        if ($null -ne $median) {
            $template[$column] = Format-BenchmarkValue -Column $column -Value $median
        }
    }

    $countColumns = @('total_files', 'processed_count', 'success_count', 'fail_count', 'merged_cell_count', 'cell_availability_count', 'floating_shape_lookup_count', 'overwrite_clear_count', 'add_picture_count', 'cell_validation_count', 'picture_sizing_count', 'progress_ui_count', 'description_write_count', 'number_alignment', 'number_position')
    foreach ($column in $countColumns) {
        $values = @(
            foreach ($row in $Rows) {
                Get-IntOrNull -Value ([string]$row[$column])
            }
        ) | Where-Object { $null -ne $_ }
        if ($null -eq $values) { continue }
        $median = Get-MedianInt -Values @([int[]]$values)
        if ($null -ne $median) {
            $template[$column] = Format-BenchmarkValue -Column $column -Value $median
        }
    }

    return $template
}

if (-not (Test-Path -LiteralPath $SourceCsv)) {
    Write-Error "Source benchmark CSV not found: $SourceCsv"
    exit 2
}

if ((Test-Path -LiteralPath $OutputCsv) -and -not $Force) {
    Write-Error "Output already exists: $OutputCsv. Use -Force to overwrite."
    exit 2
}

$source = Read-BenchmarkCsvFile -Path $SourceCsv
$scenarioIds = Get-ScenarioIdsToCapture -Filter $ScenarioId
$outputRows = New-Object System.Collections.Generic.List[hashtable]

foreach ($scenario in $scenarioIds) {
    $map = $script:BenchmarkScenarioMap[$scenario]
    if ($null -eq $map) {
        throw "Unknown scenario mapping: $scenario"
    }

    $matches = Select-BenchmarkRows -Rows $source.Rows -RunMode $map.RunMode -TotalFiles $map.TotalFiles -ScenarioId $scenario
    if ($matches.Length -lt $RunsPerScenario) {
        Write-Warning ("{0}: only {1} eligible row(s), need {2}. Using all available." -f $scenario, $matches.Length, $RunsPerScenario)
    }

    $selected = @($matches | Select-Object -Last $RunsPerScenario)

    if ($selected.Length -eq 0) {
        Write-Error "No eligible Completed rows for $scenario (run_mode=$($map.RunMode), total_files=$($map.TotalFiles))."
        exit 1
    }

    $outputRows.Add((Build-MedianBenchmarkRow -Scenario $scenario -Rows $selected)) | Out-Null
    Write-Host ("Captured {0}: {1} run(s) -> median row (run_mode={2}, total_files={3})" -f $scenario, $selected.Length, $map.RunMode, $map.TotalFiles)
}

Write-BenchmarkCsvFile -Path $OutputCsv -Rows $outputRows.ToArray() -IncludeScenarioId
Write-Host "Baseline written to $OutputCsv"
exit 0
