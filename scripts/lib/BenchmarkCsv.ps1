# Shared CSV helpers for wordtools-benchmark.csv (matches BenchmarkLogService header).

Set-StrictMode -Version Latest

$script:BenchmarkCsvHeader = @(
    'timestamp_utc', 'run_mode', 'status', 'document_path', 'log_path', 'source_path',
    'total_files', 'processed_count', 'success_count', 'fail_count', 'merged_cell_count',
    'cancelled', 'need_description', 'use_filename_as_description', 'use_foldername_as_description',
    'auto_numbering', 'number_alignment', 'number_position', 'total_seconds',
    'initialize_ms', 'clear_numbering_ms', 'calculate_start_number_ms', 'preallocate_rows_ms',
    'insert_images_ms', 'wrap_up_ms', 'cell_availability_ms', 'cell_availability_count',
    'floating_shape_lookup_ms', 'floating_shape_lookup_count', 'overwrite_clear_ms',
    'overwrite_clear_count', 'add_picture_ms', 'add_picture_count', 'cell_validation_ms',
    'cell_validation_count', 'picture_sizing_ms', 'picture_sizing_count', 'progress_ui_ms',
    'progress_ui_count', 'description_write_ms', 'description_write_count',
    'skipped_clear_numbering', 'error_message'
)

$script:BenchmarkScenarioMap = @{
    'PERF-01' = @{ RunMode = 'SelectedFiles'; TotalFiles = 4 }
    'PERF-02' = @{ RunMode = 'Folder'; TotalFiles = 5 }
    'PERF-03' = @{ RunMode = 'SelectedFiles'; TotalFiles = 50 }
}

$script:BenchmarkRatioMetrics = @(
    @{ Name = 'total_seconds'; MaxRatio = 1.05 }
    @{ Name = 'insert_images_ms'; MaxRatio = 1.05 }
    @{ Name = 'add_picture_ms'; MaxRatio = 1.05 }
    @{ Name = 'clear_numbering_ms'; MaxRatio = 1.05 }
    @{ Name = 'cell_availability_ms'; MaxRatio = 1.08 }
    @{ Name = 'progress_ui_ms'; MaxRatio = 1.10 }
)

$script:BenchmarkExactCountMetrics = @(
    'cell_availability_count'
    'add_picture_count'
    'floating_shape_lookup_count'
    'overwrite_clear_count'
    'cell_validation_count'
    'picture_sizing_count'
    'progress_ui_count'
    'description_write_count'
)

function Split-CsvLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Line
    )

    $fields = New-Object System.Collections.Generic.List[string]
    $current = New-Object System.Text.StringBuilder
    $inQuotes = $false

    for ($i = 0; $i -lt $Line.Length; $i++) {
        $ch = $Line[$i]

        if ($inQuotes) {
            if ($ch -eq '"') {
                if ($i + 1 -lt $Line.Length -and $Line[$i + 1] -eq '"') {
                    [void]$current.Append('"')
                    $i++
                }
                else {
                    $inQuotes = $false
                }
            }
            else {
                [void]$current.Append($ch)
            }
            continue
        }

        if ($ch -eq '"') {
            $inQuotes = $true
            continue
        }

        if ($ch -eq ',') {
            $fields.Add($current.ToString())
            [void]$current.Clear()
            continue
        }

        [void]$current.Append($ch)
    }

    $fields.Add($current.ToString())
    return ,@($fields.ToArray())
}

function Read-BenchmarkCsvFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Benchmark CSV not found: $Path"
    }

    $lines = [System.IO.File]::ReadAllLines($Path)
    if ($lines.Length -lt 1) {
        throw "Benchmark CSV is empty: $Path"
    }

    $headerFields = Split-CsvLine -Line $lines[0]
    $hasScenarioId = ($headerFields[0] -eq 'scenario_id')
    $benchmarkHeaderStart = if ($hasScenarioId) { 1 } else { 0 }

    $rows = New-Object System.Collections.Generic.List[hashtable]
    for ($lineIndex = 1; $lineIndex -lt $lines.Length; $lineIndex++) {
        $line = $lines[$lineIndex]
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        $fields = Split-CsvLine -Line $line
        if ($fields.Length -lt $headerFields.Length) {
            continue
        }

        $row = @{}
        for ($i = 0; $i -lt $headerFields.Length; $i++) {
            $row[$headerFields[$i]] = $fields[$i]
        }

        if ($hasScenarioId) {
            $row['scenario_id'] = $fields[0]
        }

        $rows.Add($row) | Out-Null
    }

    return [PSCustomObject]@{
        Path = $Path
        HasScenarioId = $hasScenarioId
        Header = $headerFields
        Rows = $rows.ToArray()
    }
}

function Get-NumericOrNull {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $number = 0.0
    if ([double]::TryParse($Value, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }

    return $null
}

function Get-IntOrNull {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $number = 0
    if ([int]::TryParse($Value, [ref]$number)) {
        return $number
    }

    return $null
}

function Test-BenchmarkRowEligible {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Row
    )

    $status = [string]$Row['status']
    $cancelled = [string]$Row['cancelled']
    return ($status -eq 'Completed' -and $cancelled -ne 'True')
}

function Select-BenchmarkRows {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Rows,
        [string]$RunMode,
        [int]$TotalFiles = 0,
        [string]$ScenarioId
    )

    $filtered = @($Rows | Where-Object {
        if (-not (Test-BenchmarkRowEligible -Row $_)) { return $false }

        if (-not [string]::IsNullOrWhiteSpace($ScenarioId)) {
            if ($_.ContainsKey('scenario_id') -and [string]$_.scenario_id -eq $ScenarioId) {
                return $true
            }

            $map = $script:BenchmarkScenarioMap[$ScenarioId]
            if ($null -eq $map) { return $false }
            return ([string]$_.run_mode -eq $map.RunMode -and (Get-IntOrNull $_.total_files) -eq $map.TotalFiles)
        }

        if (-not [string]::IsNullOrWhiteSpace($RunMode) -and [string]$_.run_mode -ne $RunMode) {
            return $false
        }

        if ($TotalFiles -gt 0 -and (Get-IntOrNull $_.total_files) -ne $TotalFiles) {
            return $false
        }

        return $true
    })

    return ,$filtered
}

function Get-LatestBenchmarkRow {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Rows,
        [string]$RunMode,
        [int]$TotalFiles = 0,
        [string]$ScenarioId
    )

    $matches = Select-BenchmarkRows -Rows $Rows -RunMode $RunMode -TotalFiles $TotalFiles -ScenarioId $ScenarioId
    if ($matches.Length -eq 0) {
        return $null
    }

    return $matches[-1]
}

function Get-Median {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [double[]]$Values
    )

    if ($null -eq $Values -or $Values.Length -eq 0) {
        return $null
    }

    $sorted = @($Values | Sort-Object)
    $mid = [int][math]::Floor(($sorted.Length - 1) / 2)
    if ($sorted.Length % 2 -eq 1) {
        return $sorted[$mid]
    }

    return ($sorted[$mid] + $sorted[$mid + 1]) / 2.0
}

function Get-MedianInt {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [int[]]$Values
    )

    if ($null -eq $Values -or $Values.Length -eq 0) {
        return $null
    }

    $sorted = @($Values | Sort-Object)
    $mid = [int][math]::Floor(($sorted.Length - 1) / 2)
    if ($sorted.Length % 2 -eq 1) {
        return $sorted[$mid]
    }

    return [int][math]::Round(($sorted[$mid] + $sorted[$mid + 1]) / 2.0)
}

function Format-BenchmarkValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Column,
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return ''
    }

    if ($Column -eq 'total_seconds') {
        return ([double]$Value).ToString('F3', [System.Globalization.CultureInfo]::InvariantCulture)
    }

    if ($Column -match '_count$' -or $Column -match '_ms$' -or $Column -in @('total_files', 'processed_count', 'success_count', 'fail_count', 'merged_cell_count', 'number_alignment', 'number_position')) {
        return ([int][math]::Round([double]$Value)).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    }

    return [string]$Value
}

function Escape-BenchmarkCsvField {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrEmpty($Value)) {
        return ''
    }

    $normalized = $Value.Replace("`r`n", ' ').Replace("`r", ' ').Replace("`n", ' ')
    while ($normalized.Contains('  ')) {
        $normalized = $normalized.Replace('  ', ' ')
    }

    if ($normalized.IndexOfAny(@(',', '"')) -ge 0) {
        return '"' + $normalized.Replace('"', '""') + '"'
    }

    return $normalized
}

function Write-BenchmarkCsvFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [array]$Rows,
        [switch]$IncludeScenarioId
    )

    $header = if ($IncludeScenarioId) {
        @('scenario_id') + $script:BenchmarkCsvHeader
    }
    else {
        $script:BenchmarkCsvHeader
    }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(($header -join ',')) | Out-Null

    foreach ($row in $Rows) {
        $fields = New-Object System.Collections.Generic.List[string]
        if ($IncludeScenarioId) {
            $fields.Add((Escape-BenchmarkCsvField -Value ([string]$row.scenario_id))) | Out-Null
        }

        foreach ($column in $script:BenchmarkCsvHeader) {
            $value = if ($row.ContainsKey($column)) { [string]$row[$column] } else { '' }
            $fields.Add((Escape-BenchmarkCsvField -Value $value)) | Out-Null
        }

        $lines.Add(($fields.ToArray() -join ',')) | Out-Null
    }

    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    [System.IO.File]::WriteAllLines($Path, $lines.ToArray(), (New-Object System.Text.UTF8Encoding $false))
}
