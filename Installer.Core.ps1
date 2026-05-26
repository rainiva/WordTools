[CmdletBinding()]
param(
    [ValidateSet("Probe", "Plan", "Register", "Unregister", "WpsAddinsWlExperiment")]
    [string]$Mode = "Probe",

    [ValidateSet("PreviewOnly", "Live")]
    [string]$ExecutionIntent = "PreviewOnly",

    [ValidateSet("Auto", "x86", "x64")]
    [string]$Architecture = "Auto",

    [ValidateSet("Debug", "Release", "Debug_verify")]
    [string]$Configuration = "Debug",

    [Alias("Host")]
    [ValidateSet("Word", "WPS", "Both")]
    [string]$RequestedHost = "Word",

    [string]$OutputPath,

    [string]$SummaryTextPath,

    [string]$DllPathOverride,

    [string]$EvidenceLabel,

    [switch]$AllowSelfElevation,

    [switch]$LiveElevatedRelaunch,

    [switch]$AppendEvidenceMarkdown,

    [string]$EvidenceMarkdownPath,

    [ValidateSet("backup", "write", "verify", "restore")]
    [string]$Action,

    [AllowEmptyString()]
    [string]$ProgId = "",

    [AllowEmptyString()]
    [string]$ValuePayload = "",

    [string]$ExperimentId = "wps-addinswl-experiment-1",

    [string]$EvidenceDir
)

# Usage example for the probe gate:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\Installer.Core.ps1 -Mode Probe

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Get-SupportMatrixPath {
    $scriptRoot = Split-Path -Parent $PSCommandPath
    return Join-Path $scriptRoot "Installer.SupportMatrix.json"
}

function Get-DefaultEvidenceMarkdownPath {
    $scriptRoot = Split-Path -Parent $PSCommandPath
    return Join-Path $scriptRoot "docs\installer\host-detection-matrix.md"
}

function Get-SupportMatrix {
    $matrixPath = Get-SupportMatrixPath

    if (-not (Test-Path -LiteralPath $matrixPath)) {
        throw "Support matrix file not found: $matrixPath"
    }

    return Get-Content -LiteralPath $matrixPath -Raw | ConvertFrom-Json
}

function Save-ProbeOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$JsonText,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$TargetPath
    )

    if ([string]::IsNullOrWhiteSpace($TargetPath)) {
        return
    }

    $fullPath = [System.IO.Path]::GetFullPath($TargetPath)
    $parentDirectory = [System.IO.Path]::GetDirectoryName($fullPath)
    if (-not [string]::IsNullOrWhiteSpace($parentDirectory)) {
        [System.IO.Directory]::CreateDirectory($parentDirectory) | Out-Null
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($fullPath, $JsonText, $utf8NoBom)
}

function Save-TextOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$TargetPath
    )

    if ([string]::IsNullOrWhiteSpace($TargetPath)) {
        return
    }

    $fullPath = [System.IO.Path]::GetFullPath($TargetPath)
    $parentDirectory = [System.IO.Path]::GetDirectoryName($fullPath)
    if (-not [string]::IsNullOrWhiteSpace($parentDirectory)) {
        [System.IO.Directory]::CreateDirectory($parentDirectory) | Out-Null
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($fullPath, $Text, $utf8NoBom)
}

function New-LiveFailurePayload {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Operation,

        [Parameter(Mandatory = $true)]
        [object]$ProbeResult,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    return [pscustomobject]@{
        ProbeMode      = $Operation
        EvidenceLabel  = $EvidenceLabel
        ProbedAtUtc    = $ProbeResult.ProbedAtUtc
        SupportState   = $ProbeResult.SupportState
        SupportSummary = $ProbeResult.SupportSummary
        Hosts          = $ProbeResult.Hosts
        LiveFailure    = [pscustomobject]@{
            Operation        = $Operation
            ExecutionMode    = "Live"
            Succeeded        = $false
            ErrorType        = $ErrorRecord.Exception.GetType().FullName
            ErrorMessage     = $ErrorRecord.Exception.Message
            ScriptStackTrace = $ErrorRecord.ScriptStackTrace
        }
    }
}

function Get-LiveResultSummaryText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Operation,

        [Parameter(Mandatory = $true)]
        [bool]$Succeeded,

        [Parameter(Mandatory = $true)]
        [string]$DetailMessage
    )

    if ($Succeeded) {
        return "Shared installer core completed live $Operation successfully.`r`n$DetailMessage"
    }

    return "Shared installer core failed during live $Operation.`r`n$DetailMessage"
}

function Get-DetectedHostDetailsText {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProbeResult
    )

    $details = @($ProbeResult.Hosts | ForEach-Object {
            $hostLabel = "{0} {1}" -f $_.HostName, $_.HostBitness
            $installState = if ($_.PSObject.Properties.Name -contains "InstallState" -and -not [string]::IsNullOrWhiteSpace([string]$_.InstallState)) {
                [string]$_.InstallState
            }
            else {
                "unknown"
            }
            $versionLine = if ($_.PSObject.Properties.Name -contains "VersionLine" -and -not [string]::IsNullOrWhiteSpace([string]$_.VersionLine)) {
                [string]$_.VersionLine
            }
            else {
                "unknown"
            }
            $validationStage = if ($_.PSObject.Properties.Name -contains "ValidationStage" -and -not [string]::IsNullOrWhiteSpace([string]$_.ValidationStage)) {
                [string]$_.ValidationStage
            }
            else {
                "unknown"
            }
            $diagnosticsBundleId = if ($_.PSObject.Properties.Name -contains "DiagnosticsBundleId" -and -not [string]::IsNullOrWhiteSpace([string]$_.DiagnosticsBundleId)) {
                [string]$_.DiagnosticsBundleId
            }
            else {
                "unknown"
            }

            "{0} [InstallState={1}; VersionLine={2}; ValidationStage={3}; DiagnosticsBundleId={4}]" -f $hostLabel, $installState, $versionLine, $validationStage, $diagnosticsBundleId
        })

    if ($details.Count -eq 0) {
        return "(none)"
    }

    return ($details -join ", ")
}

function Format-MarkdownCell {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return "(none)"
    }

    return $Value.Replace("|", "\|").Replace("`r", " ").Replace("`n", " ").Trim()
}

function Ensure-EvidenceMarkdownTable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    $fullPath = [System.IO.Path]::GetFullPath($TargetPath)
    $parentDirectory = [System.IO.Path]::GetDirectoryName($fullPath)
    if (-not [string]::IsNullOrWhiteSpace($parentDirectory)) {
        [System.IO.Directory]::CreateDirectory($parentDirectory) | Out-Null
    }

    if (-not (Test-Path -LiteralPath $fullPath)) {
        $initialText = @(
            "# Host Detection Matrix",
            "",
            "This file stores probe evidence for future installer support expansion.",
            "",
            "## Rules",
            "",
            "- Do not mark a host and bitness combination as supported without real probe evidence.",
            "- Do not change the current Word-only registration path during the probe phase.",
            "- Do not route WPS detection or registration through the live Word flow before validation is complete.",
            "",
            "## Validation Stages",
            "",
            "Use the following stage labels when recording probe or acceptance evidence:",
            "",
            "- `probe only`: host detection or reconnaissance exists, but no real UI proof exists yet",
            "- `ui failed`: a real UI validation attempt happened and did not surface the `WordTools` entry",
            "- `experimental ui passed`: the host surfaced a real UI entry, but `P0` is not yet a formal supported path",
            "- `formal p0 passed`: real UI and formal `P0` acceptance both passed for the tested version line",
            "",
            "## Pending Validation Matrix",
            "",
            "| Host | Bitness | Support state | Validation stage | Evidence |",
            "| --- | --- | --- | --- | --- |",
            "| Word | x86 | Planned | probe only | Not yet collected |",
            "| Word | x64 | Supported | formal p0 passed | Current machine validated in existing evidence set |",
            "| WPS | x86 | Planned | ui failed | Current machine UI failure recorded on 2026-05-24 |",
            "| WPS | x64 | Planned | probe only | Not yet collected |",
            "",
            "## Probe Evidence Log",
            "",
            "| Evidence Label | Probed At (UTC) | Validation Stage | Detected Hosts | Supported Hosts | Planned Hosts | Missing Expected Hosts | Ambiguous Hosts |",
            "| --- | --- | --- | --- | --- | --- | --- | --- |"
        ) -join [Environment]::NewLine

        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($fullPath, $initialText + [Environment]::NewLine, $utf8NoBom)
        return
    }

    $existingText = [System.IO.File]::ReadAllText($fullPath)
    $normalizedText = $existingText
    $legacyHeader = "| Evidence Label | Probed At (UTC) | Detected Hosts | Planned Hosts | Missing Expected Hosts | Ambiguous Hosts |"
    $legacySeparator = "| --- | --- | --- | --- | --- | --- |"
    $newHeader = "| Evidence Label | Probed At (UTC) | Detected Hosts | Supported Hosts | Planned Hosts | Missing Expected Hosts | Ambiguous Hosts |"
    $newSeparator = "| --- | --- | --- | --- | --- | --- | --- |"
    $phasedHeader = "| Evidence Label | Probed At (UTC) | Validation Stage | Detected Hosts | Supported Hosts | Planned Hosts | Missing Expected Hosts | Ambiguous Hosts |"
    $phasedSeparator = "| --- | --- | --- | --- | --- | --- | --- | --- |"
    $legacyPendingHeader = "| Host | Bitness | Probe status | Evidence |"
    $legacyPendingSeparator = "| --- | --- | --- | --- |"
    $phasedPendingHeader = "| Host | Bitness | Support state | Validation stage | Evidence |"
    $phasedPendingSeparator = "| --- | --- | --- | --- | --- |"

    if ($normalizedText.Contains($legacyHeader)) {
        $normalizedText = $normalizedText.Replace($legacyHeader, $newHeader)
    }

    if ($normalizedText.Contains($legacySeparator) -and -not $normalizedText.Contains($newSeparator)) {
        $normalizedText = $normalizedText.Replace($legacySeparator, $newSeparator)
    }

    if ($normalizedText.Contains($newHeader)) {
        $normalizedText = $normalizedText.Replace($newHeader, $phasedHeader)
    }

    if ($normalizedText.Contains($newSeparator)) {
        $normalizedText = $normalizedText.Replace($newSeparator, $phasedSeparator)
    }

    if ($normalizedText.Contains($legacyPendingHeader)) {
        $normalizedText = $normalizedText.Replace($legacyPendingHeader, $phasedPendingHeader)
    }

    if ($normalizedText.Contains($legacyPendingSeparator) -and -not $normalizedText.Contains($phasedPendingSeparator)) {
        $normalizedText = $normalizedText.Replace($legacyPendingSeparator, $phasedPendingSeparator)
    }

    if (-not $normalizedText.Contains("## Validation Stages")) {
        $validationStageSection = @(
            "",
            "## Validation Stages",
            "",
            "Use the following stage labels when recording probe or acceptance evidence:",
            "",
            "- `probe only`: host detection or reconnaissance exists, but no real UI proof exists yet",
            "- `ui failed`: a real UI validation attempt happened and did not surface the `WordTools` entry",
            "- `experimental ui passed`: the host surfaced a real UI entry, but `P0` is not yet a formal supported path",
            "- `formal p0 passed`: real UI and formal `P0` acceptance both passed for the tested version line"
        ) -join [Environment]::NewLine

        if ($normalizedText.Contains("## Pending Validation Matrix")) {
            $pendingIndex = $normalizedText.IndexOf("## Pending Validation Matrix", [System.StringComparison]::Ordinal)
            $prefix = $normalizedText.Substring(0, $pendingIndex).TrimEnd()
            $suffix = $normalizedText.Substring($pendingIndex)
            $normalizedText = $prefix + $validationStageSection + [Environment]::NewLine + $suffix
        }
        else {
            $normalizedText = $normalizedText.TrimEnd() + $validationStageSection + [Environment]::NewLine
        }
    }

    if (-not $normalizedText.Contains("## Pending Validation Matrix")) {
        if ($normalizedText.Contains("## Pending combinations")) {
            $normalizedText = $normalizedText.Replace("## Pending combinations", "## Pending Validation Matrix")
        }
        else {
            $pendingSection = @(
                "",
                "## Pending Validation Matrix",
                "",
                "| Host | Bitness | Support state | Validation stage | Evidence |",
                "| --- | --- | --- | --- | --- |",
                "| Word | x86 | Planned | probe only | Not yet collected |",
                "| Word | x64 | Supported | formal p0 passed | Current machine validated in existing evidence set |",
                "| WPS | x86 | Planned | ui failed | Current machine UI failure recorded on 2026-05-24 |",
                "| WPS | x64 | Planned | probe only | Not yet collected |"
            ) -join [Environment]::NewLine

            if ($normalizedText.Contains("## Probe Evidence Log")) {
                $probeIndex = $normalizedText.IndexOf("## Probe Evidence Log", [System.StringComparison]::Ordinal)
                $prefix = $normalizedText.Substring(0, $probeIndex).TrimEnd()
                $suffix = $normalizedText.Substring($probeIndex)
                $normalizedText = $prefix + $pendingSection + [Environment]::NewLine + $suffix
            }
            else {
                $normalizedText = $normalizedText.TrimEnd() + $pendingSection + [Environment]::NewLine
            }
        }
    }

    if (-not $normalizedText.Contains($phasedHeader)) {
        $tableText = @(
            "",
            "## Probe Evidence Log",
            "",
            $phasedHeader,
            $phasedSeparator
        ) -join [Environment]::NewLine

        $normalizedText = $normalizedText.TrimEnd() + [Environment]::NewLine + $tableText + [Environment]::NewLine
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($fullPath, $normalizedText, $utf8NoBom)
}

function Append-ProbeEvidenceMarkdown {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProbeResult,

        [AllowEmptyString()]
        [string]$TargetPath
    )

    $resolvedPath = $TargetPath
    if ([string]::IsNullOrWhiteSpace($resolvedPath)) {
        $resolvedPath = Get-DefaultEvidenceMarkdownPath
    }

    Ensure-EvidenceMarkdownTable -TargetPath $resolvedPath

    $fullPath = [System.IO.Path]::GetFullPath($resolvedPath)
    $label = Format-MarkdownCell $ProbeResult.EvidenceLabel
    $probedAt = Format-MarkdownCell $ProbeResult.ProbedAtUtc
    $validationStage = Format-MarkdownCell (Get-ValidationStageForProbeResult -ProbeResult $ProbeResult)
    $detected = Format-MarkdownCell (($ProbeResult.SupportSummary.DetectedHosts -join ", "))
    $supported = Format-MarkdownCell (($ProbeResult.SupportSummary.SupportedHosts -join ", "))
    $planned = Format-MarkdownCell (($ProbeResult.SupportSummary.PlannedHosts -join ", "))
    $missing = Format-MarkdownCell (($ProbeResult.SupportSummary.MissingExpectedHosts -join ", "))
    $ambiguous = Format-MarkdownCell (($ProbeResult.SupportSummary.AmbiguousHosts -join ", "))
    $row = "| $label | $probedAt | $validationStage | $detected | $supported | $planned | $missing | $ambiguous |"

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $currentText = [System.IO.File]::ReadAllText($fullPath)
    [System.IO.File]::WriteAllText($fullPath, $currentText.TrimEnd() + [Environment]::NewLine + $row + [Environment]::NewLine, $utf8NoBom)
}

function Get-ValidationStageForProbeResult {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProbeResult
    )

    $evidenceLabel = [string]$ProbeResult.EvidenceLabel
    $supportedHosts = @($ProbeResult.SupportSummary.SupportedHosts)
    $plannedHosts = @($ProbeResult.SupportSummary.PlannedHosts)

    if ($evidenceLabel.IndexOf("UiLoadFailure", [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
        return "ui failed"
    }

    if ($plannedHosts.Count -gt 0) {
        return "probe only"
    }

    if ($supportedHosts.Count -gt 0) {
        return "formal p0 passed"
    }

    return "probe only"
}

function Get-RegistryStringValue {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.Win32.RegistryHive]$Hive,

        [Parameter(Mandatory = $true)]
        [Microsoft.Win32.RegistryView]$View,

        [Parameter(Mandatory = $true)]
        [string]$SubKey,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$ValueName
    )

    try {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, $View)
        $key = $baseKey.OpenSubKey($SubKey)
        if ($null -eq $key) {
            return $null
        }

        $value = $key.GetValue($ValueName, $null)
        if ($value -is [string] -and -not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }

        return $null
    }
    catch {
        return $null
    }
}

function Add-HostCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Candidates,

        [Parameter(Mandatory = $true)]
        [string]$HostName,

        [Parameter(Mandatory = $true)]
        [string]$ProbeSource,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$ExecutablePath
    )

    if ([string]::IsNullOrWhiteSpace($ExecutablePath)) {
        return
    }

    $normalizedPath = $ExecutablePath.Trim('"')
    if (-not (Test-Path -LiteralPath $normalizedPath)) {
        return
    }

    foreach ($existing in $Candidates) {
        if ($existing.ExecutablePath -ieq $normalizedPath) {
            return
        }
    }

    $Candidates.Add([pscustomobject]@{
            HostName       = $HostName
            ExecutablePath = $normalizedPath
            ProbeSource    = $ProbeSource
        })
}

function Get-InstalledWordCandidates {
    $candidates = [System.Collections.Generic.List[object]]::new()
    $appPathSubKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE"

    foreach ($view in @([Microsoft.Win32.RegistryView]::Registry32, [Microsoft.Win32.RegistryView]::Registry64)) {
        $path = Get-RegistryStringValue -Hive LocalMachine -View $view -SubKey $appPathSubKey -ValueName ""
        Add-HostCandidate -Candidates $candidates -HostName "Word" -ProbeSource "AppPaths:$view" -ExecutablePath $path

        $path = Get-RegistryStringValue -Hive CurrentUser -View $view -SubKey $appPathSubKey -ValueName ""
        Add-HostCandidate -Candidates $candidates -HostName "Word" -ProbeSource "AppPathsUser:$view" -ExecutablePath $path
    }

    return $candidates
}

function Get-InstalledWpsCandidates {
    $candidates = [System.Collections.Generic.List[object]]::new()
    $appPathKeys = @(
        "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\wps.exe",
        "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WPS.EXE"
    )

    foreach ($view in @([Microsoft.Win32.RegistryView]::Registry32, [Microsoft.Win32.RegistryView]::Registry64)) {
        foreach ($subKey in $appPathKeys) {
            $path = Get-RegistryStringValue -Hive LocalMachine -View $view -SubKey $subKey -ValueName ""
            Add-HostCandidate -Candidates $candidates -HostName "WPS" -ProbeSource "AppPaths:$view" -ExecutablePath $path

            $path = Get-RegistryStringValue -Hive CurrentUser -View $view -SubKey $subKey -ValueName ""
            Add-HostCandidate -Candidates $candidates -HostName "WPS" -ProbeSource "AppPathsUser:$view" -ExecutablePath $path
        }
    }

    return $candidates
}

function Get-WpsReconData {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExecutablePath,

        [string]$HostBitness
    )

    function Test-AsciiLiteralInFile {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path,

            [Parameter(Mandatory = $true)]
            [string]$Literal
        )

        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
            return $false
        }

        try {
            $text = [System.Text.Encoding]::ASCII.GetString([System.IO.File]::ReadAllBytes($Path))
            return $text.Contains($Literal)
        }
        catch {
            return $false
        }
    }

    function Test-Utf16LeLiteralInFile {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path,

            [Parameter(Mandatory = $true)]
            [string]$Literal
        )

        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
            return $false
        }

        try {
            $bytes = [System.IO.File]::ReadAllBytes($Path)
            if ($bytes.Length -le 1) {
                return $false
            }

            $literalBytes = [System.Text.Encoding]::Unicode.GetBytes($Literal)
            if ($literalBytes.Length -le 0 -or $literalBytes.Length -gt $bytes.Length) {
                return $false
            }

            for ($offset = 0; $offset -le ($bytes.Length - $literalBytes.Length); $offset++) {
                $matched = $true
                for ($index = 0; $index -lt $literalBytes.Length; $index++) {
                    if ($bytes[$offset + $index] -ne $literalBytes[$index]) {
                        $matched = $false
                        break
                    }
                }

                if ($matched) {
                    return $true
                }
            }

            return $false
        }
        catch {
            return $false
        }
    }

    function Test-UInt32LittleEndianPatternInFile {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path,

            [Parameter(Mandatory = $true)]
            [UInt32]$Value
        )

        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
            return $false
        }

        try {
            $bytes = [System.IO.File]::ReadAllBytes($Path)
            $needle = [BitConverter]::GetBytes($Value)
            if ($needle.Length -le 0 -or $needle.Length -gt $bytes.Length) {
                return $false
            }

            for ($offset = 0; $offset -le ($bytes.Length - $needle.Length); $offset++) {
                $matched = $true
                for ($index = 0; $index -lt $needle.Length; $index++) {
                    if ($bytes[$offset + $index] -ne $needle[$index]) {
                        $matched = $false
                        break
                    }
                }

                if ($matched) {
                    return $true
                }
            }

            return $false
        }
        catch {
            return $false
        }
    }

    function Get-NamedDllExports {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path
        )

        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
            return @()
        }

        try {
            $bytes = [System.IO.File]::ReadAllBytes($Path)
            if ([System.Text.Encoding]::ASCII.GetString($bytes, 0, 2) -ne "MZ") {
                return @()
            }

            $peOffset = [BitConverter]::ToUInt32($bytes, 0x3C)
            if ([System.Text.Encoding]::ASCII.GetString($bytes, $peOffset, 4) -ne ("PE" + [char]0 + [char]0)) {
                return @()
            }

            $coffOffset = $peOffset + 4
            $numberOfSections = [BitConverter]::ToUInt16($bytes, $coffOffset + 2)
            $optionalHeaderOffset = $coffOffset + 20
            $optionalHeaderMagic = [BitConverter]::ToUInt16($bytes, $optionalHeaderOffset)
            $dataDirectoryOffset = $optionalHeaderOffset + $(if ($optionalHeaderMagic -eq 0x20B) { 112 } else { 96 })
            $exportTableRva = [BitConverter]::ToUInt32($bytes, $dataDirectoryOffset)
            if ($exportTableRva -eq 0) {
                return @()
            }

            $sectionTableOffset = $optionalHeaderOffset + [BitConverter]::ToUInt16($bytes, $coffOffset + 16)
            $sections = [System.Collections.Generic.List[object]]::new()
            for ($sectionIndex = 0; $sectionIndex -lt $numberOfSections; $sectionIndex++) {
                $offset = $sectionTableOffset + ($sectionIndex * 40)
                $sections.Add([pscustomobject]@{
                        VirtualAddress = [BitConverter]::ToUInt32($bytes, $offset + 12)
                        VirtualSize    = [BitConverter]::ToUInt32($bytes, $offset + 8)
                        SizeOfRawData  = [BitConverter]::ToUInt32($bytes, $offset + 16)
                        PointerToRawData = [BitConverter]::ToUInt32($bytes, $offset + 20)
                    })
            }

            $convertRvaToOffset = {
                param([UInt32]$Rva)

                foreach ($section in $sections) {
                    $mappedSize = [Math]::Max($section.VirtualSize, $section.SizeOfRawData)
                    if ($Rva -ge $section.VirtualAddress -and $Rva -lt ($section.VirtualAddress + $mappedSize)) {
                        return [int]($section.PointerToRawData + ($Rva - $section.VirtualAddress))
                    }
                }

                return $null
            }

            $exportTableOffset = & $convertRvaToOffset $exportTableRva
            if ($null -eq $exportTableOffset) {
                return @()
            }

            $namePointerTableRva = [BitConverter]::ToUInt32($bytes, $exportTableOffset + 32)
            $ordinalTableRva = [BitConverter]::ToUInt32($bytes, $exportTableOffset + 36)
            $namedExportCount = [BitConverter]::ToUInt32($bytes, $exportTableOffset + 24)
            if ($namedExportCount -le 0) {
                return @()
            }

            $exports = [System.Collections.Generic.List[string]]::new()
            for ($exportIndex = 0; $exportIndex -lt $namedExportCount; $exportIndex++) {
                $namePointerOffset = (& $convertRvaToOffset $namePointerTableRva) + ($exportIndex * 4)
                if ($null -eq $namePointerOffset) {
                    continue
                }

                $nameRva = [BitConverter]::ToUInt32($bytes, $namePointerOffset)
                $nameOffset = & $convertRvaToOffset $nameRva
                if ($null -eq $nameOffset) {
                    continue
                }

                $nameBuilder = New-Object System.Text.StringBuilder
                for ($nameIndex = $nameOffset; $nameIndex -lt $bytes.Length -and $bytes[$nameIndex] -ne 0; $nameIndex++) {
                    [void]$nameBuilder.Append([char]$bytes[$nameIndex])
                }

                $exportName = $nameBuilder.ToString()
                if (-not [string]::IsNullOrWhiteSpace($exportName)) {
                    $exports.Add($exportName)
                }
            }

            return @($exports.ToArray())
        }
        catch {
            return @()
        }
    }

    $officeRoot = Split-Path -Path $ExecutablePath -Parent
    $installRoot = Split-Path -Path $officeRoot -Parent
    $addonUserRoot = Join-Path $env:APPDATA "Kingsoft\wps\addons"
    $resolvedHostBitness = if ([string]::IsNullOrWhiteSpace($HostBitness)) { Get-ExecutableBitness -ExecutablePath $ExecutablePath } else { $HostBitness }
    $addonStorageArchitectureSegment = if ($resolvedHostBitness -eq "x64") { "win-x64" } else { "win-i386" }
    $addonPoolRoot = Join-Path $addonUserRoot ("pool\" + $addonStorageArchitectureSegment)
    $addonListRoot = Join-Path $addonUserRoot ("list\" + $addonStorageArchitectureSegment)
    $addonListV3Root = Join-Path $addonUserRoot ("listV3\" + $addonStorageArchitectureSegment)
    $addonStorageCandidates = [System.Collections.Generic.List[object]]::new()
    foreach ($candidateArchitectureSegment in @("win-i386", "win-x64")) {
        $candidatePoolRoot = Join-Path $addonUserRoot ("pool\" + $candidateArchitectureSegment)
        $candidateListRoot = Join-Path $addonUserRoot ("list\" + $candidateArchitectureSegment)
        $candidateListV3Root = Join-Path $addonUserRoot ("listV3\" + $candidateArchitectureSegment)
        $addonStorageCandidates.Add([pscustomobject]@{
            ArchitectureSegment = $candidateArchitectureSegment
            PoolRoot            = $candidatePoolRoot
            PoolRootPresent     = Test-Path -LiteralPath $candidatePoolRoot -PathType Container
            ListRoot            = $candidateListRoot
            ListRootPresent     = Test-Path -LiteralPath $candidateListRoot -PathType Container
            ListV3Root          = $candidateListV3Root
            ListV3RootPresent   = Test-Path -LiteralPath $candidateListV3Root -PathType Container
        })
    }
    $knownDirectories = [System.Collections.Generic.List[string]]::new()

    foreach ($relativePath in @("addons", "setupplugincfg", "startup")) {
        $candidatePath = Join-Path $officeRoot $relativePath
        if (Test-Path -LiteralPath $candidatePath -PathType Container) {
            $knownDirectories.Add($candidatePath)
        }
    }

    $registryPresence = [ordered]@{
        "HKCU:\Software\Kingsoft\Office\6.0\wps"            = Test-Path -LiteralPath "HKCU:\Software\Kingsoft\Office\6.0\wps"
        "HKLM:\Software\Kingsoft\Office\6.0\Common"         = Test-Path -LiteralPath "HKLM:\Software\Kingsoft\Office\6.0\Common"
        "HKCU:\Software\Kingsoft\Office\6.0\wps\Addins"     = Test-Path -LiteralPath "HKCU:\Software\Kingsoft\Office\6.0\wps\Addins"
        "HKCU:\Software\Kingsoft\Office\6.0\wps\AddinsWl"   = Test-Path -LiteralPath "HKCU:\Software\Kingsoft\Office\6.0\wps\AddinsWl"
        "HKCU:\Software\Kingsoft\Office\WPS\AddinsWl"       = Test-Path -LiteralPath "HKCU:\Software\Kingsoft\Office\WPS\AddinsWl"
        "HKLM:\Software\Kingsoft\Office\6.0\wps"            = Test-Path -LiteralPath "HKLM:\Software\Kingsoft\Office\6.0\wps"
    }

    $likelyExternalAddinRegistryRoot = "HKCU:\Software\Kingsoft\Office\WPS\AddinsWl"
    $existingExternalAddinEntryNames = [System.Collections.Generic.List[string]]::new()
    $nonEmptyExternalAddinEntries = [System.Collections.Generic.List[object]]::new()
    if (Test-Path -LiteralPath $likelyExternalAddinRegistryRoot) {
        try {
            $item = Get-ItemProperty -LiteralPath $likelyExternalAddinRegistryRoot
            foreach ($property in $item.PSObject.Properties) {
                if ($property.Name -like "PS*") {
                    continue
                }

                if (-not [string]::IsNullOrWhiteSpace($property.Name)) {
                    $existingExternalAddinEntryNames.Add($property.Name)
                    $propertyValue = [string]$property.Value
                    if (-not [string]::IsNullOrWhiteSpace($propertyValue)) {
                        $nonEmptyExternalAddinEntries.Add([pscustomobject]@{
                                Name  = $property.Name
                                Value = $propertyValue
                            })
                    }
                }
            }
        }
        catch {
        }
    }

    $externalAddinEntryResolutionSamples = [System.Collections.Generic.List[object]]::new()
    foreach ($entryName in ($existingExternalAddinEntryNames | Sort-Object | Select-Object -First 5)) {
        $externalAddinEntryResolutionSamples.Add([pscustomobject]@{
            Name                    = $entryName
            HkcrProgIdPresent       = Test-Path -LiteralPath ("Registry::HKEY_CLASSES_ROOT\" + $entryName)
            HkcuClassesProgIdPresent = Test-Path -LiteralPath ("Registry::HKEY_CURRENT_USER\Software\Classes\" + $entryName)
            HklmClassesProgIdPresent = Test-Path -LiteralPath ("Registry::HKEY_LOCAL_MACHINE\Software\Classes\" + $entryName)
        })
    }

    $resolvedExternalAddinProgIdSampleCount = @(
        $externalAddinEntryResolutionSamples |
        Where-Object { $_.HkcrProgIdPresent -or $_.HkcuClassesProgIdPresent -or $_.HklmClassesProgIdPresent }
    ).Count
    $resolvedExternalAddinProgIdTotalCount = @(
        $existingExternalAddinEntryNames |
        Where-Object {
            $entryName = $_
            (Test-Path -LiteralPath ("Registry::HKEY_CLASSES_ROOT\" + $entryName)) -or
            (Test-Path -LiteralPath ("Registry::HKEY_CURRENT_USER\Software\Classes\" + $entryName)) -or
            (Test-Path -LiteralPath ("Registry::HKEY_LOCAL_MACHINE\Software\Classes\" + $entryName))
        }
    ).Count

    $clues = [System.Collections.Generic.List[string]]::new()
    $setupPluginPath = Join-Path $officeRoot "setupplugincfg\setupplugin.plg"
    if (Test-Path -LiteralPath $setupPluginPath -PathType Leaf) {
        $clues.Add($setupPluginPath)
    }

    $getHeaderAsciiPrefix = {
        param(
            [byte[]]$Bytes,
            [int]$Count = 8
        )

        if ($null -eq $Bytes -or $Bytes.Length -le 0) {
            return $null
        }

        $builder = New-Object System.Text.StringBuilder
        $limit = [Math]::Min($Count, $Bytes.Length)
        for ($index = 0; $index -lt $limit; $index++) {
            $currentByte = $Bytes[$index]
            if ($currentByte -ge 32 -and $currentByte -le 126) {
                [void]$builder.Append([char]$currentByte)
            }
            else {
                [void]$builder.Append('.')
            }
        }

        return $builder.ToString()
    }

    $setupPluginManifestSample = [System.Collections.Generic.List[object]]::new()
    if (Test-Path -LiteralPath $setupPluginPath -PathType Leaf) {
        try {
            $setupPluginText = [System.IO.File]::ReadAllText($setupPluginPath)
            foreach ($pluginMatch in [regex]::Matches($setupPluginText, '<plugin\s+([^>]+)/?>')) {
                if ($setupPluginManifestSample.Count -ge 12) {
                    break
                }

                $attributeMap = @{}
                foreach ($attributeMatch in [regex]::Matches($pluginMatch.Groups[1].Value, '([\w_]+)\s*=\s*"([^"]*)"')) {
                    $attributeMap[$attributeMatch.Groups[1].Value] = $attributeMatch.Groups[2].Value
                }

                if ($attributeMap.Count -le 0) {
                    continue
                }

                $apiNamespace = $null
                if ($attributeMap.ContainsKey("api_namespace")) {
                    $apiNamespace = [string]$attributeMap["api_namespace"]
                }
                elseif ($attributeMap.ContainsKey("api_namespace_list")) {
                    $apiNamespace = [string]$attributeMap["api_namespace_list"]
                }

                $autorunDelaySecs = $null
                if ($attributeMap.ContainsKey("autorun_delaysecs")) {
                    $autorunDelaySecs = [string]$attributeMap["autorun_delaysecs"]
                }

                $pluginName = if ($attributeMap.ContainsKey("name")) { [string]$attributeMap["name"] } else { $null }
                $poolPackageDirectory = $null
                $poolPackageHasPluginProviderJson = $false
                $poolPackageHasRunInfoJson = $false
                $poolPackageHasConfigJson = $false
                $poolPackageHasRunIni = $false
                $poolPackageHasAttrPlg = $false
                $poolPackageRunInfoAppId = $null
                $poolPackageEntryDll = $null
                $poolPackageEntryPoint = $null
                $poolPackageLauncherType = $null
                $declaredHosts = [System.Collections.Generic.List[string]]::new()
                $poolPackageRuntimeShape = "Unknown"
                $indexSampleFileCount = 0
                $indexPluginNameLiteralDetected = $false
                $indexPluginNameUtf16LiteralDetected = $false
                $indexPoolPackageRunInfoAppIdLiteralDetected = $false
                $indexPoolPackageRunInfoAppIdUtf16LiteralDetected = $false
                $indexPoolPackageEntryDllLiteralDetected = $false
                $indexPoolPackageEntryDllUtf16LiteralDetected = $false
                $indexPoolPackageEntryPointLiteralDetected = $false
                $indexPoolPackageEntryPointUtf16LiteralDetected = $false
                $indexPoolPackageLauncherTypeLiteralDetected = $false
                $indexPoolPackageLauncherTypeUtf16LiteralDetected = $false
                if ($attributeMap.ContainsKey("host")) {
                    foreach ($declaredHost in ([string]$attributeMap["host"] -split ";" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
                        $normalizedDeclaredHost = $declaredHost.Trim()
                        if (-not [string]::IsNullOrWhiteSpace($normalizedDeclaredHost)) {
                            $declaredHosts.Add($normalizedDeclaredHost)
                        }
                    }
                }
                if ((-not [string]::IsNullOrWhiteSpace($pluginName)) -and (Test-Path -LiteralPath $addonPoolRoot -PathType Container)) {
                    $poolPackageMatch = Get-ChildItem -LiteralPath $addonPoolRoot -Directory -Filter ($pluginName + "_*") -ErrorAction SilentlyContinue |
                        Sort-Object Name -Descending |
                        Select-Object -First 1
                    if ($null -ne $poolPackageMatch) {
                        $poolPackageDirectory = $poolPackageMatch.FullName
                        $poolPackageHasPluginProviderJson = Test-Path -LiteralPath (Join-Path $poolPackageDirectory "plugin-provider.json") -PathType Leaf
                        $poolPackageHasRunInfoJson = Test-Path -LiteralPath (Join-Path $poolPackageDirectory "runinfo.json") -PathType Leaf
                        $poolPackageHasConfigJson = Test-Path -LiteralPath (Join-Path $poolPackageDirectory "config.json") -PathType Leaf
                        $poolPackageHasRunIni = Test-Path -LiteralPath (Join-Path $poolPackageDirectory "run.ini") -PathType Leaf
                        $poolPackageHasAttrPlg = Test-Path -LiteralPath (Join-Path $poolPackageDirectory "__attr.plg") -PathType Leaf
                        if ($poolPackageHasRunInfoJson) {
                            try {
                                $mappedRunInfoRoot = ConvertFrom-Json ([System.IO.File]::ReadAllText((Join-Path $poolPackageDirectory "runinfo.json")))
                                $mappedFirstApp = $mappedRunInfoRoot.PSObject.Properties | Select-Object -First 1
                                if ($null -ne $mappedFirstApp) {
                                    $poolPackageRunInfoAppId = [string]$mappedFirstApp.Name
                                    $poolPackageEntryDll = [string]$mappedFirstApp.Value.entryDll
                                    $poolPackageEntryPoint = [string]$mappedFirstApp.Value.entryPoint
                                    $poolPackageLauncherType = [string]$mappedFirstApp.Value.launcherType
                                }
                            }
                            catch {
                            }
                        }

                        if ($poolPackageHasRunInfoJson) {
                            $poolPackageRuntimeShape = "RunInfoDrivenWebOrJsApi"
                        }
                        elseif ($poolPackageHasAttrPlg -and (-not $poolPackageHasRunInfoJson) -and (-not $poolPackageHasRunIni)) {
                            $poolPackageRuntimeShape = "DllAttrNativeModule"
                        }
                    }
                }

                $setupPluginManifestSample.Add([pscustomobject]@{
                    Host                          = if ($attributeMap.ContainsKey("host")) { [string]$attributeMap["host"] } else { $null }
                    DeclaredHosts                 = @($declaredHosts.ToArray())
                    Mode                          = if ($attributeMap.ContainsKey("mode")) { [string]$attributeMap["mode"] } else { $null }
                    Name                          = $pluginName
                    Type                          = if ($attributeMap.ContainsKey("type")) { [string]$attributeMap["type"] } else { $null }
                    ApiNamespace                  = $apiNamespace
                    AutorunDelaySecs              = $autorunDelaySecs
                    PoolPackageDirectory          = $poolPackageDirectory
                    PoolPackageDirectoryPresent = ($null -ne $poolPackageDirectory)
                    PoolPackageHasPluginProviderJson = $poolPackageHasPluginProviderJson
                    PoolPackageHasRunInfoJson    = $poolPackageHasRunInfoJson
                    PoolPackageHasConfigJson     = $poolPackageHasConfigJson
                    PoolPackageHasRunIni         = $poolPackageHasRunIni
                    PoolPackageHasAttrPlg        = $poolPackageHasAttrPlg
                    PoolPackageRunInfoAppId      = $poolPackageRunInfoAppId
                    PoolPackageEntryDll          = $poolPackageEntryDll
                    PoolPackageEntryPoint        = $poolPackageEntryPoint
                    PoolPackageLauncherType      = $poolPackageLauncherType
                    PoolPackageRuntimeShape      = $poolPackageRuntimeShape
                    IndexSampleFileCount         = $indexSampleFileCount
                    IndexPluginNameLiteralDetected = $indexPluginNameLiteralDetected
                    IndexPluginNameUtf16LiteralDetected = $indexPluginNameUtf16LiteralDetected
                    IndexPoolPackageRunInfoAppIdLiteralDetected = $indexPoolPackageRunInfoAppIdLiteralDetected
                    IndexPoolPackageRunInfoAppIdUtf16LiteralDetected = $indexPoolPackageRunInfoAppIdUtf16LiteralDetected
                    IndexPoolPackageEntryDllLiteralDetected = $indexPoolPackageEntryDllLiteralDetected
                    IndexPoolPackageEntryDllUtf16LiteralDetected = $indexPoolPackageEntryDllUtf16LiteralDetected
                    IndexPoolPackageEntryPointLiteralDetected = $indexPoolPackageEntryPointLiteralDetected
                    IndexPoolPackageEntryPointUtf16LiteralDetected = $indexPoolPackageEntryPointUtf16LiteralDetected
                    IndexPoolPackageLauncherTypeLiteralDetected = $indexPoolPackageLauncherTypeLiteralDetected
                    IndexPoolPackageLauncherTypeUtf16LiteralDetected = $indexPoolPackageLauncherTypeUtf16LiteralDetected
                })
            }
        }
        catch {
        }
    }

    $setupPluginAuthInfoPath = Join-Path $officeRoot "setupplugincfg\setuppluginauthinfo.json"
    $setupPluginAuthInfoBinaryWrapped = $null
    $setupPluginAuthInfoHeaderAsciiPrefix = $null
    if (Test-Path -LiteralPath $setupPluginAuthInfoPath -PathType Leaf) {
        try {
            $setupPluginAuthInfoBytes = [System.IO.File]::ReadAllBytes($setupPluginAuthInfoPath)
            $setupPluginAuthInfoHeaderAsciiPrefix = & $getHeaderAsciiPrefix $setupPluginAuthInfoBytes
            if ($setupPluginAuthInfoBytes.Length -gt 0) {
                $setupPluginAuthInfoBinaryWrapped = ($setupPluginAuthInfoBytes[0] -ne 0x7B) -and ($setupPluginAuthInfoBytes[0] -ne 0x5B)
            }
        }
        catch {
        }
    }

    $comAddinsDialogPath = Join-Path $officeRoot "cfgs\winclassname\wps.ini"
    if (Test-Path -LiteralPath $comAddinsDialogPath -PathType Leaf) {
        $clues.Add($comAddinsDialogPath)
    }

    $configuredComAddinsDialogHosts = [System.Collections.Generic.List[string]]::new()
    $comAddinsDialogConfigFiles = [System.Collections.Generic.List[string]]::new()
    $configuredComAddinsDialogClass = $null
    $dialogConfigCandidates = @(
        @{ HostName = "WPS"; Path = (Join-Path $officeRoot "cfgs\winclassname\wps.ini") },
        @{ HostName = "WPP"; Path = (Join-Path $officeRoot "cfgs\winclassname\wpp.ini") },
        @{ HostName = "ET"; Path = (Join-Path $officeRoot "cfgs\winclassname\et.ini") }
    )

    foreach ($dialogConfig in $dialogConfigCandidates) {
        $dialogConfigPath = [string]$dialogConfig.Path
        if (-not (Test-Path -LiteralPath $dialogConfigPath -PathType Leaf)) {
            continue
        }

        $comAddinsDialogConfigFiles.Add($dialogConfigPath)

        try {
            $dialogConfigText = [System.IO.File]::ReadAllText($dialogConfigPath)
            $dialogConfigMatch = [regex]::Match($dialogConfigText, "(?m)^\s*KxCOMAddinsDlg\s*=\s*(.+?)\s*$")
            if ($dialogConfigMatch.Success) {
                $configuredValue = $dialogConfigMatch.Groups[1].Value.Trim()
                if (-not [string]::IsNullOrWhiteSpace($configuredValue)) {
                    $configuredComAddinsDialogHosts.Add([string]$dialogConfig.HostName)
                    if ([string]::IsNullOrWhiteSpace($configuredComAddinsDialogClass)) {
                        $configuredComAddinsDialogClass = $configuredValue
                    }
                }
            }
        }
        catch {
        }
    }

    $kshellPath = Join-Path $officeRoot "kshell.dll"
    $comAddinsUiStringsDetected = [System.Collections.Generic.List[string]]::new()
    foreach ($literal in @(
        "COM Add-Ins",
        "Load behavior",
        "Load at Startup",
        "Load on Demand",
        "Failed to add Add-In",
        "Failed to modify Add-In settings",
        "KxCOMAddinsDlg"
    )) {
        if (Test-AsciiLiteralInFile -Path $kshellPath -Literal $literal) {
            $comAddinsUiStringsDetected.Add($literal)
        }
    }

    $wpsApiPath = Join-Path $officeRoot "wpsapi.dll"
    $wppApiPath = Join-Path $officeRoot "wppapi.dll"
    $etApiPath = Join-Path $officeRoot "etapi.dll"
    $ksoApiPath = Join-Path $officeRoot "ksoapi.dll"
    $applicationApiSurface = [ordered]@{
        WpsHasAddIns                  = Test-AsciiLiteralInFile -Path $wpsApiPath -Literal "KyWpsApplication::get_AddIns"
        WpsHasComAddIns               = Test-AsciiLiteralInFile -Path $wpsApiPath -Literal "KyWpsApplication::get_COMAddIns"
        WppHasAddIns                  = Test-AsciiLiteralInFile -Path $wppApiPath -Literal "KyWppApplication::get_AddIns"
        WppHasComAddIns               = Test-AsciiLiteralInFile -Path $wppApiPath -Literal "KyWppApplication::get_COMAddIns"
        EtHasAddIns                   = Test-AsciiLiteralInFile -Path $etApiPath -Literal "KyEtApplication::get_AddIns"
        EtHasComAddIns                = Test-AsciiLiteralInFile -Path $etApiPath -Literal "KyEtApplication::get_COMAddIns"
        SharedTypeLibraryHasComAddIns = (Test-AsciiLiteralInFile -Path $ksoApiPath -Literal "MsoCOMAddIns") -or
                                        (Test-AsciiLiteralInFile -Path $ksoApiPath -Literal "COMAddIns")
    }

    $comAddinsCommandDatabaseEvidence = [System.Collections.Generic.List[string]]::new()
    $comAddinsCommandDatabaseSamples = [System.Collections.Generic.List[object]]::new()
    if (Test-Path -LiteralPath $addonPoolRoot -PathType Container) {
        foreach ($databaseFile in (Get-ChildItem -LiteralPath $addonPoolRoot -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match "\\kcmddb_[^\\]+\\db\\personal_cn\\(wps|wpp|et|pdf)\.db$" })) {
            $containsComAddIns = Test-AsciiLiteralInFile -Path $databaseFile.FullName -Literal "COMAddIns"
            if ($containsComAddIns) {
                $comAddinsCommandDatabaseEvidence.Add($databaseFile.FullName)
                if ($comAddinsCommandDatabaseSamples.Count -lt 8) {
                    $packageVersion = $null
                    $hostDatabaseName = [System.IO.Path]::GetFileNameWithoutExtension($databaseFile.Name)
                    $packageVersionMatch = [regex]::Match($databaseFile.FullName, "\\kcmddb_([^\\]+)\\db\\personal_cn\\")
                    if ($packageVersionMatch.Success) {
                        $packageVersion = $packageVersionMatch.Groups[1].Value
                    }

                    $comAddinsCommandDatabaseSamples.Add([pscustomobject]@{
                        Path                               = $databaseFile.FullName
                        PackageVersion                     = $packageVersion
                        HostDatabaseName                   = $hostDatabaseName
                        ContainsComAddIns                  = $containsComAddIns
                        ContainsKdocerjsapi20             = Test-AsciiLiteralInFile -Path $databaseFile.FullName -Literal "kdocerjsapi20"
                        ContainsKwpsaiwordtool            = Test-AsciiLiteralInFile -Path $databaseFile.FullName -Literal "kwpsaiwordtool"
                        ContainsPictureResourceshopSplit  = Test-AsciiLiteralInFile -Path $databaseFile.FullName -Literal "picture_resourceshop_split"
                        ContainsKdocerjsapi20EntryDll     = Test-AsciiLiteralInFile -Path $databaseFile.FullName -Literal "kdocerjsapi20.dll"
                        ContainsKdocerjsapi20EntryPoint   = Test-AsciiLiteralInFile -Path $databaseFile.FullName -Literal "CreateSplitAppWidget"
                        ContainsKdocerjsapi20LauncherType = Test-AsciiLiteralInFile -Path $databaseFile.FullName -Literal "proxyFrame"
                    })
                }
            }
        }
    }

    $suspiciousModuleFiles = [System.Collections.Generic.List[string]]::new()
    $suspiciousModuleSamples = [System.Collections.Generic.List[object]]::new()
    if (Test-Path -LiteralPath $addonPoolRoot -PathType Container) {
        $modulePatterns = @(
            "kpluginmanager_*\kpluginmanager.dll",
            "kvbarunner_*\kvbarunner.dll",
            "kwpsremixsdksrv_*\kwpsremixsdksrv.dll",
            "wpsioplugin_*\wpsr.dll",
            "wpsioplugin_*\wppr.dll",
            "wpsioplugin_*\et10rw.dll"
        )

        foreach ($relativePattern in $modulePatterns) {
            $match = Get-ChildItem -LiteralPath $addonPoolRoot -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object { $_.FullName -like (Join-Path $addonPoolRoot $relativePattern) } |
                Select-Object -First 1
            if ($null -ne $match) {
                $suspiciousModuleFiles.Add($match.FullName)
                $namedExportSample = @((Get-NamedDllExports -Path $match.FullName) | Select-Object -First 12)
                $suspiciousModuleSamples.Add([pscustomobject]@{
                    Path                   = $match.FullName
                    NamedExportCount       = @($namedExportSample).Count
                    NamedExportSample      = @($namedExportSample)
                    HasDllRegisterServer   = Test-AsciiLiteralInFile -Path $match.FullName -Literal "DllRegisterServer"
                    HasDllGetClassObject   = Test-AsciiLiteralInFile -Path $match.FullName -Literal "DllGetClassObject"
                    HasProgIdLiteral       = Test-AsciiLiteralInFile -Path $match.FullName -Literal "ProgID"
                    HasClsidLiteral        = Test-AsciiLiteralInFile -Path $match.FullName -Literal "CLSID"
                    HasIDispatchLiteral    = Test-AsciiLiteralInFile -Path $match.FullName -Literal "IDispatch"
                    HasLoadBehaviorLiteral = Test-AsciiLiteralInFile -Path $match.FullName -Literal "LoadBehavior"
                })
            }
        }
    }

    $listV3VersionDirectories = [System.Collections.Generic.List[string]]::new()
    $pluginListFiles = [System.Collections.Generic.List[string]]::new()
    if (Test-Path -LiteralPath $addonListV3Root -PathType Container) {
        foreach ($versionDirectory in (Get-ChildItem -LiteralPath $addonListV3Root -Directory -ErrorAction SilentlyContinue)) {
            $listV3VersionDirectories.Add($versionDirectory.FullName)
            $pluginListPath = Join-Path $versionDirectory.FullName "pluginlist.plg"
            if (Test-Path -LiteralPath $pluginListPath -PathType Leaf) {
                $pluginListFiles.Add($pluginListPath)
            }
        }
    }

    $nativePluginPackageMetadataFiles = [System.Collections.Generic.List[string]]::new()
    $nativeMetadataSearchFiles = [System.Collections.Generic.List[string]]::new()
    if (Test-Path -LiteralPath $addonPoolRoot -PathType Container) {
        $nativeMetadataFileNames = @("plugin-provider.json", "runinfo.json", "config.json", "run.ini")
        foreach ($metadataFile in (Get-ChildItem -LiteralPath $addonPoolRoot -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $nativeMetadataFileNames -contains $_.Name })) {
            if (-not $nativeMetadataSearchFiles.Contains($metadataFile.FullName)) {
                $nativeMetadataSearchFiles.Add($metadataFile.FullName)
            }

            if ($nativePluginPackageMetadataFiles.Count -ge 12) {
                continue
            }

            if (-not $nativePluginPackageMetadataFiles.Contains($metadataFile.FullName)) {
                $nativePluginPackageMetadataFiles.Add($metadataFile.FullName)
            }
        }
    }

    $runInfoManifestSamples = [System.Collections.Generic.List[object]]::new()
    $configManifestSamples = [System.Collections.Generic.List[object]]::new()
    $runIniManifestSamples = [System.Collections.Generic.List[object]]::new()
    $preloadFileSamples = [System.Collections.Generic.List[object]]::new()
    $binaryWrappedMetadataSamples = [System.Collections.Generic.List[object]]::new()
    $nativeDllPackageSamples = [System.Collections.Generic.List[object]]::new()
    if (Test-Path -LiteralPath $addonPoolRoot -PathType Container) {
        foreach ($runInfoFile in (Get-ChildItem -LiteralPath $addonPoolRoot -Recurse -File -Filter "runinfo.json" -ErrorAction SilentlyContinue | Select-Object -First 4)) {
            try {
                $runInfoRoot = ConvertFrom-Json ([System.IO.File]::ReadAllText($runInfoFile.FullName))
                $firstApp = $runInfoRoot.PSObject.Properties | Select-Object -First 1
                if ($null -ne $firstApp) {
                    $runInfoManifestSamples.Add([pscustomobject]@{
                        Path         = $runInfoFile.FullName
                        AppId        = [string]$firstApp.Name
                        EntryDll     = [string]$firstApp.Value.entryDll
                        EntryPoint   = [string]$firstApp.Value.entryPoint
                        LauncherType = [string]$firstApp.Value.launcherType
                    })
                }
            }
            catch {
            }
        }

        foreach ($configFile in (Get-ChildItem -LiteralPath $addonPoolRoot -Recurse -File -Filter "config.json" -ErrorAction SilentlyContinue | Select-Object -First 4)) {
            try {
                $configRoot = ConvertFrom-Json ([System.IO.File]::ReadAllText($configFile.FullName))
                $configManifestSamples.Add([pscustomobject]@{
                    Path                = $configFile.FullName
                    OfficeType          = [string]$configRoot.office_type
                    FrontEndVersion     = [string]$configRoot.front_ver
                    SupportsPrefetch    = [bool]$configRoot.isSupportPrefetch
                })
            }
            catch {
            }
        }

        foreach ($runIniFile in (Get-ChildItem -LiteralPath $addonPoolRoot -Recurse -File -Filter "run.ini" -ErrorAction SilentlyContinue | Select-Object -First 4)) {
            try {
                $runIniText = [System.IO.File]::ReadAllText($runIniFile.FullName)
                $entryMatch = [regex]::Match($runIniText, "(?m)^\s*entry\s*=\s*(.+?)\s*$")
                $loadOnlineMatch = [regex]::Match($runIniText, "(?m)^\s*isLoadOnline\s*=\s*(.+?)\s*$")
                $runWhenExitMatch = [regex]::Match($runIniText, "(?m)^\s*isRunWhenExit\s*=\s*(.+?)\s*$")
                $runIniManifestSamples.Add([pscustomobject]@{
                    Path          = $runIniFile.FullName
                    Entry         = if ($entryMatch.Success) { $entryMatch.Groups[1].Value.Trim() } else { $null }
                    IsLoadOnline  = if ($loadOnlineMatch.Success) { $loadOnlineMatch.Groups[1].Value.Trim() } else { $null }
                    IsRunWhenExit = if ($runWhenExitMatch.Success) { $runWhenExitMatch.Groups[1].Value.Trim() } else { $null }
                })
            }
            catch {
            }
        }

        foreach ($preloadFile in (Get-ChildItem -LiteralPath $addonPoolRoot -Recurse -File -Filter "preload-file.json" -ErrorAction SilentlyContinue | Select-Object -First 4)) {
            try {
                $preloadRoot = ConvertFrom-Json ([System.IO.File]::ReadAllText($preloadFile.FullName))
                $pcBundles = @($preloadRoot.pc)
                $mobileBundles = @($preloadRoot.mobile)
                $containsWriterBundle = @($pcBundles + $mobileBundles) | Where-Object { $_ -like "*writer/*" } | Select-Object -First 1
                $preloadFileSamples.Add([pscustomobject]@{
                    Path                  = $preloadFile.FullName
                    PcBundleCount         = $pcBundles.Count
                    MobileBundleCount     = $mobileBundles.Count
                    ContainsWriterBundles = ($null -ne $containsWriterBundle)
                })
            }
            catch {
            }
        }

        foreach ($binaryMetadataFile in (Get-ChildItem -LiteralPath $addonPoolRoot -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -eq "plugin-provider.json" -or $_.Name -eq "__attr.plg" } |
            Select-Object -First 8)) {
            try {
                $binaryMetadataBytes = [System.IO.File]::ReadAllBytes($binaryMetadataFile.FullName)
                if ($binaryMetadataBytes.Length -le 0) {
                    continue
                }

                $looksLikeTextJson = ($binaryMetadataBytes[0] -eq 0x7B) -or ($binaryMetadataBytes[0] -eq 0x5B)
                if ((-not $looksLikeTextJson) -or $binaryMetadataFile.Name -eq "__attr.plg") {
                    $binaryWrappedMetadataSamples.Add([pscustomobject]@{
                        Path              = $binaryMetadataFile.FullName
                        HeaderAsciiPrefix = & $getHeaderAsciiPrefix $binaryMetadataBytes
                    })
                }
            }
            catch {
            }
        }

        foreach ($nativeDllPrefix in @("kdocerjsapi20", "kwpsaiwordtool")) {
            $nativeDllPackageDirectory = Get-ChildItem -LiteralPath $addonPoolRoot -Directory -Filter ($nativeDllPrefix + "_*") -ErrorAction SilentlyContinue |
                Sort-Object Name -Descending |
                Select-Object -First 1
            if ($null -eq $nativeDllPackageDirectory) {
                continue
            }

            $nativeDllPath = Join-Path $nativeDllPackageDirectory.FullName ($nativeDllPrefix + ".dll")
            if (-not (Test-Path -LiteralPath $nativeDllPath -PathType Leaf)) {
                continue
            }

            $nativeDllPackageSamples.Add([pscustomobject]@{
                    PackageDirectory          = $nativeDllPackageDirectory.FullName
                    DllPath                   = $nativeDllPath
                    HasRunInfoJson            = Test-Path -LiteralPath (Join-Path $nativeDllPackageDirectory.FullName "runinfo.json") -PathType Leaf
                    HasConfigJson             = Test-Path -LiteralPath (Join-Path $nativeDllPackageDirectory.FullName "config.json") -PathType Leaf
                    HasRunIni                 = Test-Path -LiteralPath (Join-Path $nativeDllPackageDirectory.FullName "run.ini") -PathType Leaf
                    HasAttrPlg                = Test-Path -LiteralPath (Join-Path $nativeDllPackageDirectory.FullName "__attr.plg") -PathType Leaf
                    NamedExportCount          = @((Get-NamedDllExports -Path $nativeDllPath)).Count
                    NamedExportSample         = @((Get-NamedDllExports -Path $nativeDllPath) | Select-Object -First 12)
                    HasDllRegisterServer      = Test-AsciiLiteralInFile -Path $nativeDllPath -Literal "DllRegisterServer"
                    HasDllUnregisterServer    = Test-AsciiLiteralInFile -Path $nativeDllPath -Literal "DllUnregisterServer"
                    HasDllGetClassObject      = Test-AsciiLiteralInFile -Path $nativeDllPath -Literal "DllGetClassObject"
                    HasProgIdLiteral          = Test-AsciiLiteralInFile -Path $nativeDllPath -Literal "ProgID"
                    HasClsidLiteral           = Test-AsciiLiteralInFile -Path $nativeDllPath -Literal "CLSID"
                    HasIDispatchLiteral       = Test-AsciiLiteralInFile -Path $nativeDllPath -Literal "IDispatch"
                    HasLoadBehaviorLiteral    = Test-AsciiLiteralInFile -Path $nativeDllPath -Literal "LoadBehavior"
                })
        }
    }

    $indexedPluginBinaryHeaderSamples = [System.Collections.Generic.List[object]]::new()
    $indexedPluginBinarySamplePaths = [System.Collections.Generic.List[string]]::new()
    $listV3ShardDirectorySamples = [System.Collections.Generic.List[object]]::new()
    $addIndexedPluginHeaderSample = {
        param([string]$SamplePath)

        if ([string]::IsNullOrWhiteSpace($SamplePath) -or (-not (Test-Path -LiteralPath $SamplePath -PathType Leaf))) {
            return
        }

        if ($indexedPluginBinarySamplePaths.Contains($SamplePath)) {
            return
        }

        try {
            $sampleBytes = [System.IO.File]::ReadAllBytes($SamplePath)
            $indexedPluginBinarySamplePaths.Add($SamplePath)
            $indexedPluginBinaryHeaderSamples.Add([pscustomobject]@{
                Path              = $SamplePath
                HeaderAsciiPrefix = & $getHeaderAsciiPrefix $sampleBytes
            })
        }
        catch {
        }
    }

    foreach ($pluginListPath in ($pluginListFiles | Select-Object -First 2)) {
        & $addIndexedPluginHeaderSample $pluginListPath
    }

    if (Test-Path -LiteralPath $addonListV3Root -PathType Container) {
        foreach ($versionDirectory in (Get-ChildItem -LiteralPath $addonListV3Root -Directory -ErrorAction SilentlyContinue | Select-Object -First 2)) {
            $numericShardDirectories = @(
                Get-ChildItem -LiteralPath $versionDirectory.FullName -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match '^\d+$' }
            )
            $numericShardDirectory = $numericShardDirectories | Select-Object -First 1
            $pluginListPath = Join-Path $versionDirectory.FullName "pluginlist.plg"
            $sampleShardIdsPath = $null
            $sampleShardIdsLength = $null
            $sampleShardDataPath = $null
            $sampleShardDataLength = $null
            if ($null -eq $numericShardDirectory) {
                $listV3ShardDirectorySamples.Add([pscustomobject]@{
                    VersionDirectory          = $versionDirectory.FullName
                    VersionLabel              = $versionDirectory.Name
                    PluginListPath            = if (Test-Path -LiteralPath $pluginListPath -PathType Leaf) { $pluginListPath } else { $null }
                    PluginListLength          = if (Test-Path -LiteralPath $pluginListPath -PathType Leaf) { [long](Get-Item -LiteralPath $pluginListPath).Length } else { $null }
                    NumericShardDirectoryCount = $numericShardDirectories.Count
                    SampleShardDirectoryName  = $null
                    SampleShardIdsPath        = $null
                    SampleShardIdsLength      = $null
                    SampleShardDataPath       = $null
                    SampleShardDataLength     = $null
                    SampleShardDirectoryNameLiteralDetectedInPluginList = $false
                    SampleShardDirectoryNameUtf16LiteralDetectedInPluginList = $false
                    SampleShardDirectoryNameUInt32LittleEndianDetectedInPluginList = $false
                })
                continue
            }

            foreach ($fileName in @("ids", "data")) {
                $samplePath = Join-Path $numericShardDirectory.FullName $fileName
                & $addIndexedPluginHeaderSample $samplePath
                if ($fileName -eq "ids") {
                    $sampleShardIdsPath = $samplePath
                    if (Test-Path -LiteralPath $samplePath -PathType Leaf) {
                        $sampleShardIdsLength = [long](Get-Item -LiteralPath $samplePath).Length
                    }
                }
                elseif ($fileName -eq "data") {
                    $sampleShardDataPath = $samplePath
                    if (Test-Path -LiteralPath $samplePath -PathType Leaf) {
                        $sampleShardDataLength = [long](Get-Item -LiteralPath $samplePath).Length
                    }
                }
            }

            $listV3ShardDirectorySamples.Add([pscustomobject]@{
                VersionDirectory           = $versionDirectory.FullName
                VersionLabel               = $versionDirectory.Name
                PluginListPath             = if (Test-Path -LiteralPath $pluginListPath -PathType Leaf) { $pluginListPath } else { $null }
                PluginListLength           = if (Test-Path -LiteralPath $pluginListPath -PathType Leaf) { [long](Get-Item -LiteralPath $pluginListPath).Length } else { $null }
                NumericShardDirectoryCount = $numericShardDirectories.Count
                SampleShardDirectoryName   = $numericShardDirectory.Name
                SampleShardIdsPath         = $sampleShardIdsPath
                SampleShardIdsLength       = $sampleShardIdsLength
                SampleShardDataPath        = $sampleShardDataPath
                SampleShardDataLength      = $sampleShardDataLength
                SampleShardDirectoryNameLiteralDetectedInPluginList = if (Test-Path -LiteralPath $pluginListPath -PathType Leaf) { Test-AsciiLiteralInFile -Path $pluginListPath -Literal $numericShardDirectory.Name } else { $false }
                SampleShardDirectoryNameUtf16LiteralDetectedInPluginList = if (Test-Path -LiteralPath $pluginListPath -PathType Leaf) { Test-Utf16LeLiteralInFile -Path $pluginListPath -Literal $numericShardDirectory.Name } else { $false }
                SampleShardDirectoryNameUInt32LittleEndianDetectedInPluginList = if ((Test-Path -LiteralPath $pluginListPath -PathType Leaf) -and ($numericShardDirectory.Name -match '^\d+$')) { Test-UInt32LittleEndianPatternInFile -Path $pluginListPath -Value ([UInt32]$numericShardDirectory.Name) } else { $false }
                SampleShardIdLiteralDetectedInPoolMetadata = @(
                    $nativeMetadataSearchFiles |
                        Where-Object { Test-AsciiLiteralInFile -Path $_ -Literal $numericShardDirectory.Name }
                ).Count -gt 0
            })
        }
    }

    foreach ($setupPluginEntry in $setupPluginManifestSample) {
        $pluginName = if ($setupPluginEntry.PSObject.Properties.Name -contains "Name") { [string]$setupPluginEntry.Name } else { $null }
        if ([string]::IsNullOrWhiteSpace($pluginName)) {
            continue
        }

        $setupPluginEntry.IndexSampleFileCount = $indexedPluginBinarySamplePaths.Count
        $setupPluginEntry.IndexPluginNameLiteralDetected = $false
        $setupPluginEntry.IndexPluginNameUtf16LiteralDetected = $false
        $setupPluginEntry.IndexPoolPackageRunInfoAppIdLiteralDetected = $false
        $setupPluginEntry.IndexPoolPackageRunInfoAppIdUtf16LiteralDetected = $false
        $setupPluginEntry.IndexPoolPackageEntryDllLiteralDetected = $false
        $setupPluginEntry.IndexPoolPackageEntryDllUtf16LiteralDetected = $false
        $setupPluginEntry.IndexPoolPackageEntryPointLiteralDetected = $false
        $setupPluginEntry.IndexPoolPackageEntryPointUtf16LiteralDetected = $false
        $setupPluginEntry.IndexPoolPackageLauncherTypeLiteralDetected = $false
        $setupPluginEntry.IndexPoolPackageLauncherTypeUtf16LiteralDetected = $false
        foreach ($samplePath in $indexedPluginBinarySamplePaths) {
            if (Test-AsciiLiteralInFile -Path $samplePath -Literal $pluginName) {
                $setupPluginEntry.IndexPluginNameLiteralDetected = $true
            }

            if (Test-Utf16LeLiteralInFile -Path $samplePath -Literal $pluginName) {
                $setupPluginEntry.IndexPluginNameUtf16LiteralDetected = $true
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$setupPluginEntry.PoolPackageRunInfoAppId)) {
                if (Test-AsciiLiteralInFile -Path $samplePath -Literal ([string]$setupPluginEntry.PoolPackageRunInfoAppId)) {
                    $setupPluginEntry.IndexPoolPackageRunInfoAppIdLiteralDetected = $true
                }

                if (Test-Utf16LeLiteralInFile -Path $samplePath -Literal ([string]$setupPluginEntry.PoolPackageRunInfoAppId)) {
                    $setupPluginEntry.IndexPoolPackageRunInfoAppIdUtf16LiteralDetected = $true
                }
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$setupPluginEntry.PoolPackageEntryDll)) {
                if (Test-AsciiLiteralInFile -Path $samplePath -Literal ([string]$setupPluginEntry.PoolPackageEntryDll)) {
                    $setupPluginEntry.IndexPoolPackageEntryDllLiteralDetected = $true
                }

                if (Test-Utf16LeLiteralInFile -Path $samplePath -Literal ([string]$setupPluginEntry.PoolPackageEntryDll)) {
                    $setupPluginEntry.IndexPoolPackageEntryDllUtf16LiteralDetected = $true
                }
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$setupPluginEntry.PoolPackageEntryPoint)) {
                if (Test-AsciiLiteralInFile -Path $samplePath -Literal ([string]$setupPluginEntry.PoolPackageEntryPoint)) {
                    $setupPluginEntry.IndexPoolPackageEntryPointLiteralDetected = $true
                }

                if (Test-Utf16LeLiteralInFile -Path $samplePath -Literal ([string]$setupPluginEntry.PoolPackageEntryPoint)) {
                    $setupPluginEntry.IndexPoolPackageEntryPointUtf16LiteralDetected = $true
                }
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$setupPluginEntry.PoolPackageLauncherType)) {
                if (Test-AsciiLiteralInFile -Path $samplePath -Literal ([string]$setupPluginEntry.PoolPackageLauncherType)) {
                    $setupPluginEntry.IndexPoolPackageLauncherTypeLiteralDetected = $true
                }

                if (Test-Utf16LeLiteralInFile -Path $samplePath -Literal ([string]$setupPluginEntry.PoolPackageLauncherType)) {
                    $setupPluginEntry.IndexPoolPackageLauncherTypeUtf16LiteralDetected = $true
                }
            }

            if ($setupPluginEntry.IndexPluginNameLiteralDetected -or
                $setupPluginEntry.IndexPluginNameUtf16LiteralDetected -or
                $setupPluginEntry.IndexPoolPackageRunInfoAppIdLiteralDetected -or
                $setupPluginEntry.IndexPoolPackageRunInfoAppIdUtf16LiteralDetected -or
                $setupPluginEntry.IndexPoolPackageEntryDllLiteralDetected -or
                $setupPluginEntry.IndexPoolPackageEntryDllUtf16LiteralDetected -or
                $setupPluginEntry.IndexPoolPackageEntryPointLiteralDetected -or
                $setupPluginEntry.IndexPoolPackageEntryPointUtf16LiteralDetected -or
                $setupPluginEntry.IndexPoolPackageLauncherTypeLiteralDetected -or
                $setupPluginEntry.IndexPoolPackageLauncherTypeUtf16LiteralDetected) {
                break
            }
        }
    }

    $vbaRuntimeArtifacts = [System.Collections.Generic.List[string]]::new()
    $kvbaInstallRoot = Join-Path $addonPoolRoot "kvbarunner_3.1.0.17941\install"
    foreach ($relativePath in @("md5.ini", "zipmd5.ini", "vba7.zip")) {
        $artifactPath = Join-Path $kvbaInstallRoot $relativePath
        if (Test-Path -LiteralPath $artifactPath -PathType Leaf) {
            $vbaRuntimeArtifacts.Add($artifactPath)
        }
    }

    $internalComTypeLibEvidence = [System.Collections.Generic.List[string]]::new()
    $knownTypeLibKeys = @(
        "HKCU:\Software\Classes\TypeLib\{83A4B852-ABDC-482E-91D3-D69C360C6E45}\1.0\0\win32",
        "HKCU:\Software\Classes\TypeLib\{83A4B852-ABDC-482E-91D3-D69C360C6E45}\1.0\HELPDIR",
        "HKLM:\Software\Classes\TypeLib\{83A4B852-ABDC-482E-91D3-D69C360C6E45}\1.0\0\win32",
        "HKLM:\Software\Classes\TypeLib\{83A4B852-ABDC-482E-91D3-D69C360C6E45}\1.0\HELPDIR",
        "HKLM:\Software\Classes\WOW6432Node\TypeLib\{83A4B852-ABDC-482E-91D3-D69C360C6E45}\1.0\0\win32",
        "HKLM:\Software\Classes\WOW6432Node\TypeLib\{83A4B852-ABDC-482E-91D3-D69C360C6E45}\1.0\HELPDIR"
    )

    foreach ($registryPath in $knownTypeLibKeys) {
        if (Test-Path -LiteralPath $registryPath) {
            $internalComTypeLibEvidence.Add($registryPath)
        }
    }

    return [pscustomobject]@{
        InstallRoot            = $installRoot
        OfficeRoot             = $officeRoot
        AddonUserRoot          = $addonUserRoot
        AddonStorageArchitectureSegment = $addonStorageArchitectureSegment
        AddonStorageRoots      = [pscustomobject]@{
            PoolRoot   = $addonPoolRoot
            ListRoot   = $addonListRoot
            ListV3Root = $addonListV3Root
        }
        AddonStorageCandidates = @($addonStorageCandidates.ToArray())
        KnownPluginDirectories = @($knownDirectories.ToArray())
        RegistryKeyPresence    = [pscustomobject]$registryPresence
        ConfigClues            = @($clues.ToArray())
        SetupPluginManifestSample = @($setupPluginManifestSample.ToArray())
        SetupPluginAuthInfoPath = if (Test-Path -LiteralPath $setupPluginAuthInfoPath -PathType Leaf) { $setupPluginAuthInfoPath } else { $null }
        SetupPluginAuthInfoBinaryWrapped = $setupPluginAuthInfoBinaryWrapped
        SetupPluginAuthInfoHeaderAsciiPrefix = $setupPluginAuthInfoHeaderAsciiPrefix
        ConfiguredComAddinsDialogClass = $configuredComAddinsDialogClass
        ConfiguredComAddinsDialogHosts = @($configuredComAddinsDialogHosts.ToArray())
        ComAddinsDialogConfigFiles = @($comAddinsDialogConfigFiles.ToArray())
        ComAddinsDialogHostModule = if ((Test-Path -LiteralPath $kshellPath -PathType Leaf) -and $comAddinsUiStringsDetected.Count -gt 0) { $kshellPath } else { $null }
        ComAddinsUiStringsDetected = @($comAddinsUiStringsDetected.ToArray())
        ApplicationApiSurface = [pscustomobject]$applicationApiSurface
        ComAddinsCommandDatabaseEvidence = @($comAddinsCommandDatabaseEvidence.ToArray())
        ComAddinsCommandDatabaseSamples = @($comAddinsCommandDatabaseSamples.ToArray())
        LikelyExternalAddinRegistryRoot = $likelyExternalAddinRegistryRoot
        ExistingExternalAddinEntryCount = $existingExternalAddinEntryNames.Count
        ExistingExternalAddinEntrySample = @($existingExternalAddinEntryNames | Sort-Object | Select-Object -First 20)
        NonEmptyExternalAddinEntryCount = $nonEmptyExternalAddinEntries.Count
        NonEmptyExternalAddinEntries = @($nonEmptyExternalAddinEntries.ToArray())
        ExternalAddinEntryResolutionSamples = @($externalAddinEntryResolutionSamples.ToArray())
        ResolvedExternalAddinProgIdSampleCount = $resolvedExternalAddinProgIdSampleCount
        ResolvedExternalAddinProgIdTotalCount = $resolvedExternalAddinProgIdTotalCount
        SuspiciousModuleFiles  = @($suspiciousModuleFiles.ToArray())
        SuspiciousModuleSamples = @($suspiciousModuleSamples.ToArray())
        IndexedPluginStores    = [pscustomobject]@{
            ListRootPresent           = Test-Path -LiteralPath $addonListRoot -PathType Container
            ListV3RootPresent         = Test-Path -LiteralPath $addonListV3Root -PathType Container
            ListV3VersionDirectories  = @($listV3VersionDirectories.ToArray())
            PluginListFiles           = @($pluginListFiles.ToArray())
            ListV3ShardDirectorySamples = @($listV3ShardDirectorySamples.ToArray())
            BinaryHeaderSamples       = @($indexedPluginBinaryHeaderSamples.ToArray())
        }
        NativePluginPackageMetadataFiles = @($nativePluginPackageMetadataFiles.ToArray())
        NativePluginMetadataSignals = [pscustomobject]@{
            RunInfoManifestSamples     = @($runInfoManifestSamples.ToArray())
            ConfigManifestSamples      = @($configManifestSamples.ToArray())
            RunIniManifestSamples      = @($runIniManifestSamples.ToArray())
            PreloadFileSamples         = @($preloadFileSamples.ToArray())
            BinaryWrappedMetadataSamples = @($binaryWrappedMetadataSamples.ToArray())
            NativeDllPackageSamples    = @($nativeDllPackageSamples.ToArray())
        }
        VbaRuntimeInstallArtifacts = @($vbaRuntimeArtifacts.ToArray())
        InternalComTypeLibEvidence = @($internalComTypeLibEvidence.ToArray())
        Conclusion             = "WPS host detected with plugin framework directories, setupplugin.plg, binary listV3 plugin indexes, native pool package metadata files such as plugin-provider.json and runinfo.json, internal COM typelib traces, and a populated external add-in candidate root at HKCU:\\Software\\Kingsoft\\Office\\WPS\\AddinsWl. The shared COM add-ins dialog appears to be routed through kshell.dll with KxCOMAddinsDlg mapped to TCOMAddinsDlg.UnicodeClass across WPS, WPP, and ET. Current binary evidence shows sampled WPS, WPP, and ET application surfaces all exposing Application.COMAddIns alongside Application.AddIns. setupplugin.plg currently exposes host/mode/name/type plugin declarations, sampled runinfo.json files expose entryDll, entryPoint, and launcherType fields, sampled config.json files expose office_type routing, sampled run.ini files point to %workingroot%/index.html, and pluginlist.plg plus shard ids/data files all currently present as kplugin-headed binary payloads. Sampled native DLL packages currently do not expose DllRegisterServer, DllGetClassObject, ProgID, CLSID, IDispatch, or LoadBehavior literals, and the sampled kdocerjsapi20 package instead exports JS/API-style entry names such as GetExtensionJsApiObj and loadJsapiService. On this machine, 78 AddinsWl entry names were observed, zero resolved to live ProgID roots, and only two carried non-empty value payloads that currently look version-gate-like rather than LoadBehavior-style flags. The stronger native package evidence indicates that WPS activation is mediated by its own plugin/index/runtime stack, so WordTools write semantics and activation behavior for AddinsWl therefore remain unvalidated."
    }
}

function Invoke-WpsAddinsWlExperiment {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProgId,

        [Parameter(Mandatory = $true)]
        [string]$ValuePayload,

        [Parameter(Mandatory = $true)]
        [string]$ExperimentId,

        [Parameter(Mandatory = $true)]
        [string]$EvidenceDir,

        [switch]$Restore
    )

    $registryPath = "HKCU:\Software\Kingsoft\Office\WPS\AddinsWl"
    $timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz")
    $dateStamp = (Get-Date).ToString("yyyyMMdd")
    $backupFileName = "CurrentMachine-WpsX86-AddinsWl-backup-$dateStamp.reg"
    $backupPath = Join-Path $EvidenceDir $backupFileName

    if (-not (Test-Path -LiteralPath $EvidenceDir -PathType Container)) {
        New-Item -Path $EvidenceDir -ItemType Directory -Force | Out-Null
    }

    # --- Restore mode ---
    if ($Restore) {
        if (-not (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
            $errorResult = [pscustomobject]@{
                ExperimentId   = $ExperimentId
                Timestamp      = $timestamp
                RestoreSucceeded = $false
                Error          = "Backup file not found at $backupPath"
            }
            return $errorResult | ConvertTo-Json -Depth 4
        }

        reg import `"$backupPath`" *>$null
        if ($LASTEXITCODE -ne 0) {
            $errorResult = [pscustomobject]@{
                ExperimentId    = $ExperimentId
                Timestamp       = $timestamp
                RestoreSucceeded = $false
                Error           = "reg import failed with exit code $LASTEXITCODE"
            }
            return $errorResult | ConvertTo-Json -Depth 4
        }

        try {
            $postRestoreCount = (Get-ItemProperty -Path $registryPath -ErrorAction Stop).PSObject.Properties.Count
        }
        catch {
            $errorResult = [pscustomobject]@{
                ExperimentId    = $ExperimentId
                Timestamp       = $timestamp
                RestoreSucceeded = $false
                Error           = "Cannot read AddinsWl after restore: $_"
            }
            return $errorResult | ConvertTo-Json -Depth 4
        }
        $successResult = [pscustomobject]@{
            ExperimentId      = $ExperimentId
            Timestamp         = $timestamp
            RestoreSucceeded  = $true
            BackupPath        = $backupPath
            PostRestoreEntryCount = $postRestoreCount
        }
        return $successResult | ConvertTo-Json -Depth 4
    }

    # --- Pre-existing check ---
    $preExisting = $false
    $preTotal = 0
    try {
        $existingProps = Get-ItemProperty -Path $registryPath -ErrorAction Stop
        $preTotal = $existingProps.PSObject.Properties.Count
        $preExisting = ($existingProps.PSObject.Properties.Name -contains $ProgId)
    }
    catch {
        $errorResult = [pscustomobject]@{
            ExperimentId    = $ExperimentId
            Timestamp       = $timestamp
            WriteSucceeded  = $false
            Error           = "Cannot read AddinsWl registry key: $_"
        }
        return $errorResult | ConvertTo-Json -Depth 4
    }

    if ($preExisting) {
        $result = [pscustomobject]@{
            ExperimentId        = $ExperimentId
            Timestamp           = $timestamp
            WriteSucceeded      = $false
            PreExisting         = $true
            AddinsWlPreTotal    = $preTotal
            Error               = "ProgId '$ProgId' already exists in AddinsWl"
        }
        return $result | ConvertTo-Json -Depth 4
    }

    # --- Backup ---
    reg export "HKCU\Software\Kingsoft\Office\WPS\AddinsWl" `"$backupPath`" *>$null
    if ($LASTEXITCODE -ne 0) {
        $errorResult = [pscustomobject]@{
            ExperimentId    = $ExperimentId
            Timestamp       = $timestamp
            WriteSucceeded  = $false
            Error           = "Backup failed: reg export exited with code $LASTEXITCODE"
        }
        return $errorResult | ConvertTo-Json -Depth 4
    }
    if (-not (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
        $errorResult = [pscustomobject]@{
            ExperimentId    = $ExperimentId
            Timestamp       = $timestamp
            WriteSucceeded  = $false
            Error           = "Backup failed: reg export produced no output at $backupPath"
        }
        return $errorResult | ConvertTo-Json -Depth 4
    }

    $backupContent = Get-Content -LiteralPath $backupPath -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($backupContent)) {
        $errorResult = [pscustomobject]@{
            ExperimentId    = $ExperimentId
            Timestamp       = $timestamp
            WriteSucceeded  = $false
            Error           = "Backup failed: backup.reg is empty"
        }
        return $errorResult | ConvertTo-Json -Depth 4
    }

    # --- Write ---
    try {
        Set-ItemProperty -Path $registryPath -Name $ProgId -Value $ValuePayload -ErrorAction Stop
    }
    catch {
        $errorResult = [pscustomobject]@{
            ExperimentId    = $ExperimentId
            Timestamp       = $timestamp
            WriteSucceeded  = $false
            BackupPath      = $backupPath
            Error           = "Set-ItemProperty failed: $_"
        }
        return $errorResult | ConvertTo-Json -Depth 4
    }

    # --- Verify write ---
    $postTotal = 0
    $entryPresent = $false
    try {
        $postProps = Get-ItemProperty -Path $registryPath -ErrorAction Stop
        $postTotal = $postProps.PSObject.Properties.Count
        $entryPresent = ($postProps.PSObject.Properties.Name -contains $ProgId)
    }
    catch {
        $errorResult = [pscustomobject]@{
            ExperimentId    = $ExperimentId
            Timestamp       = $timestamp
            WriteSucceeded  = $false
            BackupPath      = $backupPath
            Error           = "Post-write verification failed: $_"
        }
        return $errorResult | ConvertTo-Json -Depth 4
    }

    $result = [pscustomobject]@{
        ExperimentId        = $ExperimentId
        Timestamp           = $timestamp
        WriteSucceeded      = $entryPresent
        BackupPath          = $backupPath
        WrittenProgId       = $ProgId
        WrittenPayload      = $ValuePayload
        AddinsWlPreTotal    = $preTotal
        AddinsWlPostTotal   = $postTotal
        PreExisting         = $false
    }
    return $result | ConvertTo-Json -Depth 4
}

function Get-ExecutableBitness {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExecutablePath
    )

    $stream = $null
    $reader = $null

    try {
        $stream = [System.IO.File]::Open($ExecutablePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $reader = New-Object System.IO.BinaryReader($stream)

        $stream.Seek(0x3C, [System.IO.SeekOrigin]::Begin) | Out-Null
        $peHeaderOffset = $reader.ReadInt32()

        $stream.Seek($peHeaderOffset + 4, [System.IO.SeekOrigin]::Begin) | Out-Null
        $machine = $reader.ReadUInt16()

        switch ($machine) {
            0x014c { return "x86" }
            0x8664 { return "x64" }
            0x01c4 { return "arm" }
            0xAA64 { return "arm64" }
            default { return "unknown" }
        }
    }
    catch {
        return "unknown"
    }
    finally {
        if ($reader -ne $null) {
            $reader.Dispose()
        }

        if ($stream -ne $null) {
            $stream.Dispose()
        }
    }
}

function Get-ExecutableVersionLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExecutablePath
    )

    try {
        $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($ExecutablePath)
        $versionValue = if (-not [string]::IsNullOrWhiteSpace($versionInfo.FileVersion)) {
            $versionInfo.FileVersion
        }
        elseif (-not [string]::IsNullOrWhiteSpace($versionInfo.ProductVersion)) {
            $versionInfo.ProductVersion
        }
        else {
            $null
        }

        if (-not [string]::IsNullOrWhiteSpace($versionValue)) {
            return $versionValue
        }
    }
    catch {
    }

    return "unknown"
}

function Get-UiEvidenceState {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ValidationStage
    )

    switch ($ValidationStage) {
        "formal p0 passed" { return "passed" }
        "experimental ui passed" { return "passed" }
        "ui failed" { return "failed" }
        "probe only" { return "pending" }
        default { return "unknown" }
    }
}

function Get-P0EvidenceState {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ValidationStage
    )

    switch ($ValidationStage) {
        "formal p0 passed" { return "passed" }
        "experimental ui passed" { return "pending" }
        "ui failed" { return "blocked" }
        "probe only" { return "pending" }
        default { return "unknown" }
    }
}

function Get-DiagnosticsBundleId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HostName,

        [Parameter(Mandatory = $true)]
        [string]$HostBitness,

        [string]$EvidenceLabel
    )

    $label = if ([string]::IsNullOrWhiteSpace($EvidenceLabel)) { "probe" } else { $EvidenceLabel }
    $safeLabel = ($label -replace '[^A-Za-z0-9._-]', '-').Trim('-')

    if ([string]::IsNullOrWhiteSpace($safeLabel)) {
        $safeLabel = "probe"
    }

    return "WordTools-$safeLabel-$HostName-$HostBitness"
}

function Get-PlannedStatus {
    param(
        [Parameter(Mandatory = $true)]
        [object]$SupportMatrix,

        [Parameter(Mandatory = $true)]
        [string]$HostName,

        [Parameter(Mandatory = $true)]
        [string]$Bitness
    )

    foreach ($target in $SupportMatrix.targets) {
        if ($target.host -eq $HostName -and $target.bitness -eq $Bitness) {
            return $target.status
        }
    }

    return "unknown"
}

function Get-RegistrationView {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HostBitness
    )

    switch ($HostBitness) {
        "x86" { return "Registry32" }
        "x64" { return "Registry64" }
        default { return "Unknown" }
    }
}

function Get-AmbiguityReason {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HostBitness,

        [Parameter(Mandatory = $true)]
        [string]$RegistrationView
    )

    if ($HostBitness -eq "unknown") {
        return "Bitness could not be determined from the detected executable image."
    }

    if ($RegistrationView -eq "Unknown") {
        return "Registration view could not be determined for the detected host bitness."
    }

    return $null
}

function Get-SupportDecision {
    param(
        [Parameter(Mandatory = $true)]
        [object]$SupportMatrix,

        [Parameter(Mandatory = $true)]
        [string]$HostName,

        [Parameter(Mandatory = $true)]
        [string]$HostBitness
    )

    $registrationView = Get-RegistrationView -HostBitness $HostBitness
    $ambiguityReason = Get-AmbiguityReason -HostBitness $HostBitness -RegistrationView $registrationView

    foreach ($target in $SupportMatrix.targets) {
        if ($target.host -eq $HostName -and $target.bitness -eq $HostBitness) {
            return [pscustomobject]@{
                SupportStatus    = $target.status
                ValidationStage  = $target.validationStage
                ActivationRoute  = $target.activationRoute
                SupportReason    = $target.note
                RegistrationView = $registrationView
                AmbiguityReason  = $ambiguityReason
            }
        }
    }

    return [pscustomobject]@{
        SupportStatus    = "unknown"
        ValidationStage  = "unknown"
        ActivationRoute  = "unknown"
        SupportReason    = "No support-matrix entry matched the detected host and bitness."
        RegistrationView = $registrationView
        AmbiguityReason  = $ambiguityReason
    }
}

function Get-RegAsmPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HostBitness
    )

    switch ($HostBitness) {
        "x86" { return "C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" }
        "x64" { return "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" }
        default { return $null }
    }
}

function Get-AddInRegistryPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HostName
    )

    switch ($HostName) {
        "Word" { return "HKLM:\Software\Microsoft\Office\Word\Addins\WordTools.ThisAddIn" }
        "WPS" { return "HKCU:\Software\Kingsoft\Office\WPS\AddinsWl\WordTools.ThisAddIn" }
        default { return $null }
    }
}

function Get-NgenPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HostBitness
    )

    switch ($HostBitness) {
        "x86" { return "C:\Windows\Microsoft.NET\Framework\v4.0.30319\ngen.exe" }
        "x64" { return "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\ngen.exe" }
        default { return $null }
    }
}

function Resolve-RequestedArchitecture {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RequestedArchitecture
    )

    return $RequestedArchitecture
}

function Test-RequestedArchitectureMatchesHost {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RequestedArchitecture,

        [Parameter(Mandatory = $true)]
        [object]$HostTarget
    )

    if ($RequestedArchitecture -eq "Auto") {
        return $true
    }

    $resolvedArchitecture = Resolve-RequestedArchitecture -RequestedArchitecture $RequestedArchitecture
    return $HostTarget.HostBitness -eq $resolvedArchitecture
}

function Get-RequestedHostNames {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RequestedHost
    )

    switch ($RequestedHost) {
        "Both" { return @("Word", "WPS") }
        default { return @($RequestedHost) }
    }
}

function Get-RequestedProbeHosts {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProbeResult,

        [Parameter(Mandatory = $true)]
        [string]$RequestedArchitecture,

        [Parameter(Mandatory = $true)]
        [string]$RequestedHost
    )

    $requestedHosts = Get-RequestedHostNames -RequestedHost $RequestedHost
    $matchingHosts = [System.Collections.Generic.List[object]]::new()

    foreach ($detectedHost in $ProbeResult.Hosts) {
        if (-not $requestedHosts.Contains($detectedHost.HostName)) {
            continue
        }

        if (-not (Test-RequestedArchitectureMatchesHost -RequestedArchitecture $RequestedArchitecture -HostTarget $detectedHost)) {
            continue
        }

        $matchingHosts.Add($detectedHost)
    }

    return [pscustomobject]@{
        RequestedHosts = $requestedHosts
        Hosts          = @($matchingHosts.ToArray())
    }
}

function Get-RequestedProbeResult {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProbeResult,

        [Parameter(Mandatory = $true)]
        [string]$RequestedArchitecture,

        [Parameter(Mandatory = $true)]
        [string]$RequestedHost
    )

    $requestedSelection = Get-RequestedProbeHosts -ProbeResult $ProbeResult -RequestedArchitecture $RequestedArchitecture -RequestedHost $RequestedHost
    $scopedHosts = [System.Collections.Generic.List[object]]::new()

    foreach ($hostTarget in $requestedSelection.Hosts) {
        $scopedHosts.Add($hostTarget)
    }

    return [pscustomobject]@{
        ProbeMode      = $ProbeResult.ProbeMode
        EvidenceLabel  = $ProbeResult.EvidenceLabel
        ProbedAtUtc    = $ProbeResult.ProbedAtUtc
        SupportState   = $ProbeResult.SupportState
        SupportSummary = Get-SupportSummary -Hosts $scopedHosts -SupportMatrix (Get-SupportMatrix)
        Hosts          = @($scopedHosts.ToArray())
    }
}

function Test-IsAdministrator {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-WpsLiveExecutionFeatureEnabled {
    return $false
}

function Assert-LiveExecutionAdministrator {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Operation
    )

    if (-not (Test-IsAdministrator)) {
        throw ("Live {0} requires an elevated administrator PowerShell session because RegAsm /codebase writes machine-level COM registration." -f $Operation.ToLowerInvariant())
    }
}

function Get-LiveSelfElevationArgumentList {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Operation
    )

    $argumentList = [System.Collections.Generic.List[string]]::new()
    $argumentList.Add("-NoProfile")
    $argumentList.Add("-ExecutionPolicy")
    $argumentList.Add("Bypass")
    $argumentList.Add("-File")
    $argumentList.Add($PSCommandPath)
    $argumentList.Add("-Mode")
    $argumentList.Add($Operation)
    $argumentList.Add("-ExecutionIntent")
    $argumentList.Add("Live")
    $argumentList.Add("-RequestedHost")
    $argumentList.Add($RequestedHost)
    $argumentList.Add("-Architecture")
    $argumentList.Add($Architecture)
    $argumentList.Add("-Configuration")
    $argumentList.Add($Configuration)
    $argumentList.Add("-AllowSelfElevation")
    $argumentList.Add("-LiveElevatedRelaunch")

    if (-not [string]::IsNullOrWhiteSpace($DllPathOverride)) {
        $argumentList.Add("-DllPathOverride")
        $argumentList.Add($DllPathOverride)
    }

    if (-not [string]::IsNullOrWhiteSpace($EvidenceLabel)) {
        $argumentList.Add("-EvidenceLabel")
        $argumentList.Add($EvidenceLabel)
    }

    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
        $argumentList.Add("-OutputPath")
        $argumentList.Add($OutputPath)
    }

    if (-not [string]::IsNullOrWhiteSpace($SummaryTextPath)) {
        $argumentList.Add("-SummaryTextPath")
        $argumentList.Add($SummaryTextPath)
    }

    return @($argumentList.ToArray())
}

function Invoke-SelfElevatedLiveExecution {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Operation
    )

    if ($LiveElevatedRelaunch) {
        throw "Live $Operation re-entered the self-elevation path unexpectedly."
    }

    $elevatedArguments = Get-LiveSelfElevationArgumentList -Operation $Operation

    try {
        $process = Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList $elevatedArguments -PassThru -Wait -WindowStyle Hidden
    }
    catch {
        throw "Unable to relaunch the shared installer core with administrator privileges. Please approve the UAC prompt and try again. $($_.Exception.Message)"
    }

    if ($process.ExitCode -ne 0) {
        if (-not [string]::IsNullOrWhiteSpace($SummaryTextPath) -and (Test-Path -LiteralPath $SummaryTextPath)) {
            throw ([System.IO.File]::ReadAllText([System.IO.Path]::GetFullPath($SummaryTextPath))).Trim()
        }

        throw "The elevated live $Operation session exited with code $($process.ExitCode) before the shared installer core confirmed completion."
    }

    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        return [pscustomobject]@{
            Operation      = $Operation
            ExecutionMode  = "Live"
            RelaunchMode   = "ElevatedChildProcess"
            Completed      = $true
            OutputCaptured = $false
        }
    }

    $resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
    if (-not (Test-Path -LiteralPath $resolvedOutputPath)) {
        throw "The elevated live $Operation session completed without writing the expected result file: $resolvedOutputPath"
    }

    $savedJson = [System.IO.File]::ReadAllText($resolvedOutputPath)
    return $savedJson | ConvertFrom-Json
}

function Get-PluginDllPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RequestedConfiguration
    )

    $scriptRoot = Split-Path -Parent $PSCommandPath
    return Join-Path $scriptRoot "WordTools\bin\$RequestedConfiguration\WordTools.dll"
}

function Get-EffectiveDllPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RequestedConfiguration,

        [AllowEmptyString()]
        [string]$RequestedOverride
    )

    if (-not [string]::IsNullOrWhiteSpace($RequestedOverride)) {
        return $RequestedOverride
    }

    return Get-PluginDllPath -RequestedConfiguration $RequestedConfiguration
}

function Invoke-NativeToolCapture {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolPath,

        [Parameter(Mandatory = $true)]
        [string[]]$ArgumentList
    )

    $stdoutPath = [System.IO.Path]::GetTempFileName()
    $stderrPath = [System.IO.Path]::GetTempFileName()

    try {
        $process = Start-Process -FilePath $ToolPath -ArgumentList $ArgumentList -PassThru -Wait -WindowStyle Hidden -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        $stdoutText = if (Test-Path -LiteralPath $stdoutPath) { [System.IO.File]::ReadAllText($stdoutPath) } else { "" }
        $stderrText = if (Test-Path -LiteralPath $stderrPath) { [System.IO.File]::ReadAllText($stderrPath) } else { "" }
        $combinedOutput = @($stdoutText.Trim(), $stderrText.Trim()) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        return [pscustomobject]@{
            ExitCode       = $process.ExitCode
            StandardOutput = $stdoutText.Trim()
            StandardError  = $stderrText.Trim()
            CombinedOutput = ($combinedOutput -join [Environment]::NewLine).Trim()
        }
    }
    finally {
        if (Test-Path -LiteralPath $stdoutPath) {
            Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue
        }

        if (Test-Path -LiteralPath $stderrPath) {
            Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Test-LiveExecutionAllowedForHost {
    param(
        [Parameter(Mandatory = $true)]
        [object]$HostTarget
    )

    $regAsmPath = Get-RegAsmPath -HostBitness $HostTarget.HostBitness
    $registryPath = Get-AddInRegistryPath -HostName $HostTarget.HostName

    switch ($HostTarget.HostName) {
        "Word" {
            return $HostTarget.SupportStatus -eq "supported" -and -not [string]::IsNullOrWhiteSpace($regAsmPath) -and -not [string]::IsNullOrWhiteSpace($registryPath)
        }
        "WPS" {
            return $HostTarget.SupportStatus -eq "supported" `
                -and $HostTarget.HostBitness -eq "x86" `
                -and -not [string]::IsNullOrWhiteSpace($regAsmPath) `
                -and -not [string]::IsNullOrWhiteSpace($registryPath) `
                -and (Test-WpsLiveExecutionFeatureEnabled)
        }
        default {
            return $false
        }
    }
}

function Get-DryRunRegistrationPlan {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProbeResult
    )

    $requestedSelection = Get-RequestedProbeHosts -ProbeResult $ProbeResult -RequestedArchitecture $Architecture -RequestedHost $RequestedHost
    $targets = [System.Collections.Generic.List[object]]::new()

    foreach ($detectedHost in $requestedSelection.Hosts) {
        $progId = "WordTools.ThisAddIn"
        $regAsmPath = Get-RegAsmPath -HostBitness $detectedHost.HostBitness
        $registryPath = Get-AddInRegistryPath -HostName $detectedHost.HostName
        $wouldRegister = Test-LiveExecutionAllowedForHost -HostTarget $detectedHost

        $targets.Add([pscustomobject]@{
                HostName         = $detectedHost.HostName
                HostBitness      = $detectedHost.HostBitness
                SupportStatus    = $detectedHost.SupportStatus
                ExecutionMode    = "DryRun"
                WouldRegister    = $wouldRegister
                ProgId           = $progId
                RegAsmPath       = $regAsmPath
                RegistryPath     = $registryPath
                PlannedActions   = if ($wouldRegister) {
                    @("RegisterComCodebase", "WriteWordAddInRegistry", "OptionallyInstallNgen")
                }
                else {
                    @("SkipRegistration")
                }
                DecisionReason   = if ($wouldRegister) {
                    "Supported host detected. Dry-run shows the future Word registration path without executing it."
                }
                else {
                    "Registration would be skipped because this host is not yet supported by the live registration entrypoint."
                }
            })
    }

    $supportedTargetLabels = [System.Collections.Generic.List[string]]::new()
    $plannedButNotRegistrableLabels = [System.Collections.Generic.List[string]]::new()
    $skippedTargetLabels = [System.Collections.Generic.List[string]]::new()
    $requiredRegAsmModes = [System.Collections.Generic.List[string]]::new()
    $registryWrites = [System.Collections.Generic.List[string]]::new()
    $plannedActionUnion = [System.Collections.Generic.List[string]]::new()
    $registrableTargetCount = 0
    $skippedTargetCount = 0

    foreach ($target in $targets) {
        $targetLabel = "{0} {1}" -f $target.HostName, $target.HostBitness

        if ($target.SupportStatus -eq "supported" -and -not $supportedTargetLabels.Contains($targetLabel)) {
            $supportedTargetLabels.Add($targetLabel)
        }

        if ($target.WouldRegister) {
            $registrableTargetCount++

            if (($target.HostBitness -eq "x86" -or $target.HostBitness -eq "x64") -and -not $requiredRegAsmModes.Contains($target.HostBitness)) {
                $requiredRegAsmModes.Add($target.HostBitness)
            }

            if (-not [string]::IsNullOrWhiteSpace($target.RegistryPath) -and -not $registryWrites.Contains($target.RegistryPath)) {
                $registryWrites.Add($target.RegistryPath)
            }

            foreach ($action in $target.PlannedActions) {
                if (-not [string]::IsNullOrWhiteSpace($action) -and -not $plannedActionUnion.Contains($action)) {
                    $plannedActionUnion.Add($action)
                }
            }
        }
        else {
            $skippedTargetCount++

            if (-not $skippedTargetLabels.Contains($targetLabel)) {
                $skippedTargetLabels.Add($targetLabel)
            }

            if ($target.SupportStatus -eq "planned" -and -not $plannedButNotRegistrableLabels.Contains($targetLabel)) {
                $plannedButNotRegistrableLabels.Add($targetLabel)
            }
        }
    }

    $overallDecision = if ($registrableTargetCount -gt 0) {
        "Would register $registrableTargetCount supported host in dry-run mode without executing any live registration."
    }
    else {
        "Would not register any detected host in dry-run mode because no live-registration-eligible targets were found."
    }

    return [pscustomobject]@{
        ExecutionMode = "DryRun"
        PlanSummary   = [pscustomobject]@{
            ExecutionMode                = "DryRun"
            DetectedTargetCount          = $targets.Count
            RegistrableTargetCount       = $registrableTargetCount
            SkippedTargetCount           = $skippedTargetCount
            SupportedTargetLabels        = $supportedTargetLabels
            PlannedButNotRegistrableLabels = $plannedButNotRegistrableLabels
            SkippedTargetLabels          = $skippedTargetLabels
            RequiredRegAsmModes          = $requiredRegAsmModes
            RegistryWrites               = $registryWrites
            PlannedActionUnion           = $plannedActionUnion
            OverallDecision              = $overallDecision
        }
        Targets       = $targets
    }
}

function Get-RegisterPreviewPlan {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProbeResult
    )

    $scopedProbeResult = Get-RequestedProbeResult -ProbeResult $ProbeResult -RequestedArchitecture $Architecture -RequestedHost $RequestedHost
    $dryRunPlan = Get-DryRunRegistrationPlan -ProbeResult $ProbeResult
    $targets = [System.Collections.Generic.List[object]]::new()
    $previewableTargetCount = 0

    foreach ($target in $dryRunPlan.Targets) {
        $previewTarget = Get-RegisterPreviewTarget -DryRunTarget $target
        $wouldExecute = [bool]$previewTarget.WouldExecute
        if ($wouldExecute) {
            $previewableTargetCount++
        }

        $targets.Add($previewTarget)
    }

    [object[]]$hostRuleSummaries = @(Get-PreviewHostRuleSummaries -Targets $targets)
    $liveReadinessSummary = Get-LiveReadinessSummary -HostRuleSummaries $hostRuleSummaries
    [object[]]$operationManifest = @(Get-PreviewOperationManifest -Targets $targets)
    $installerHandoffSummary = Get-InstallerHandoffSummary -ProbeResult $scopedProbeResult -OperationManifest $operationManifest
    $liveEntrypointStatus = Get-LiveEntrypointStatus
    $migrationChecklist = Get-MigrationChecklist -ProbeResult $scopedProbeResult -LiveReadinessSummary $liveReadinessSummary -LiveEntrypointStatus $liveEntrypointStatus
    $installerPreviewReport = Get-InstallerPreviewReport -Operation "Register" -ProbeResult $scopedProbeResult -OperationManifest $operationManifest -InstallerHandoffSummary $installerHandoffSummary -MigrationChecklist $migrationChecklist

    return [pscustomobject]@{
        Operation              = "Register"
        ExecutionMode          = "PreviewOnly"
        PreviewableCount       = $previewableTargetCount
        RegisterPreviewSummary = Get-PreviewSummary -Targets $targets -ExecutionVerb "register" -NoEligibleDecision "Preview-only register mode found no live-registration-eligible targets."
        HostRuleSummaries      = $hostRuleSummaries
        LiveReadinessSummary   = $liveReadinessSummary
        OperationManifest      = $operationManifest
        InstallerHandoffSummary = $installerHandoffSummary
        LiveEntrypointStatus   = $liveEntrypointStatus
        MigrationChecklist     = $migrationChecklist
        InstallerPreviewReport = $installerPreviewReport
        Targets                = $targets
    }
}

function Get-UnregisterPreviewPlan {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProbeResult
    )

    $scopedProbeResult = Get-RequestedProbeResult -ProbeResult $ProbeResult -RequestedArchitecture $Architecture -RequestedHost $RequestedHost
    $requestedSelection = Get-RequestedProbeHosts -ProbeResult $ProbeResult -RequestedArchitecture $Architecture -RequestedHost $RequestedHost
    $targets = [System.Collections.Generic.List[object]]::new()
    $previewableTargetCount = 0

    foreach ($detectedHost in $requestedSelection.Hosts) {
        $previewTarget = Get-UnregisterPreviewTarget -DetectedHost $detectedHost
        $wouldExecute = [bool]$previewTarget.WouldExecute

        if ($wouldExecute) {
            $previewableTargetCount++
        }

        $targets.Add($previewTarget)
    }

    [object[]]$hostRuleSummaries = @(Get-PreviewHostRuleSummaries -Targets $targets)
    $liveReadinessSummary = Get-LiveReadinessSummary -HostRuleSummaries $hostRuleSummaries
    [object[]]$operationManifest = @(Get-PreviewOperationManifest -Targets $targets)
    $installerHandoffSummary = Get-InstallerHandoffSummary -ProbeResult $scopedProbeResult -OperationManifest $operationManifest
    $liveEntrypointStatus = Get-LiveEntrypointStatus
    $migrationChecklist = Get-MigrationChecklist -ProbeResult $scopedProbeResult -LiveReadinessSummary $liveReadinessSummary -LiveEntrypointStatus $liveEntrypointStatus
    $installerPreviewReport = Get-InstallerPreviewReport -Operation "Unregister" -ProbeResult $scopedProbeResult -OperationManifest $operationManifest -InstallerHandoffSummary $installerHandoffSummary -MigrationChecklist $migrationChecklist

    return [pscustomobject]@{
        Operation                = "Unregister"
        ExecutionMode            = "PreviewOnly"
        PreviewableCount         = $previewableTargetCount
        UnregisterPreviewSummary = Get-PreviewSummary -Targets $targets -ExecutionVerb "unregister" -NoEligibleDecision "Preview-only unregister mode found no live-unregistration-eligible targets."
        HostRuleSummaries        = $hostRuleSummaries
        LiveReadinessSummary     = $liveReadinessSummary
        OperationManifest        = $operationManifest
        InstallerHandoffSummary  = $installerHandoffSummary
        LiveEntrypointStatus     = $liveEntrypointStatus
        MigrationChecklist       = $migrationChecklist
        InstallerPreviewReport   = $installerPreviewReport
        Targets                  = $targets
    }
}

function Get-PreviewHostRuleSummaries {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Targets
    )

    $summaries = [System.Collections.Generic.List[object]]::new()
    $seenKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)

    foreach ($target in $Targets) {
        if (-not ($target.PSObject.Properties.Name -contains "HandlerPreview")) {
            continue
        }

        $handlerPreview = $target.HandlerPreview
        if ($null -eq $handlerPreview -or -not ($handlerPreview.PSObject.Properties.Name -contains "HostRuleSummary")) {
            continue
        }

        $summary = $handlerPreview.HostRuleSummary
        $summaryKey = "{0}|{1}|{2}|{3}" -f $summary.HostName, $summary.HostBitness, $summary.Operation, $summary.HandlerName
        if ($seenKeys.Add($summaryKey)) {
            $summaries.Add($summary)
        }
    }

    return @($summaries.ToArray())
}

function Get-LiveReadinessSummary {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$HostRuleSummaries
    )

    $liveReadyHostLabels = [System.Collections.Generic.List[string]]::new()
    $previewOnlyHostLabels = [System.Collections.Generic.List[string]]::new()
    $probePendingHostLabels = [System.Collections.Generic.List[string]]::new()
    $blockedHostLabels = [System.Collections.Generic.List[string]]::new()

    foreach ($summary in $HostRuleSummaries) {
        $hostLabel = "{0} {1}" -f $summary.HostName, $summary.HostBitness

        if ($summary.LiveExecutionAllowed) {
            $liveReadyHostLabels.Add($hostLabel)
            continue
        }

        switch ($summary.CurrentSupportStatus) {
            "supported" { $previewOnlyHostLabels.Add($hostLabel) }
            "planned" { $probePendingHostLabels.Add($hostLabel) }
            default { $blockedHostLabels.Add($hostLabel) }
        }
    }

    $overallDecision = if ($liveReadyHostLabels.Count -gt 0) {
        "Live readiness has $($liveReadyHostLabels.Count) host ready for live execution and $($previewOnlyHostLabels.Count) additional host still gated behind preview-only enforcement."
    }
    elseif ($previewOnlyHostLabels.Count -gt 0) {
        "All currently supported detected hosts remain preview-only; live execution is still disabled in the shared installer core."
    }
    elseif ($probePendingHostLabels.Count -gt 0) {
        "Detected hosts remain probe-pending and cannot advance to live execution until support validation completes."
    }
    else {
        "No detected hosts are currently ready for live execution."
    }

    return [pscustomobject]@{
        ExecutionMode          = "PreviewOnly"
        DetectedHostCount      = $HostRuleSummaries.Count
        LiveReadyHostCount     = $liveReadyHostLabels.Count
        PreviewOnlyHostCount   = $previewOnlyHostLabels.Count
        ProbePendingHostCount  = $probePendingHostLabels.Count
        BlockedHostCount       = $blockedHostLabels.Count
        LiveReadyHostLabels    = @($liveReadyHostLabels.ToArray())
        PreviewOnlyHostLabels  = @($previewOnlyHostLabels.ToArray())
        ProbePendingHostLabels = @($probePendingHostLabels.ToArray())
        BlockedHostLabels      = @($blockedHostLabels.ToArray())
        OverallDecision        = $overallDecision
    }
}

function Get-PreviewOperationManifest {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Targets
    )

    $manifest = [System.Collections.Generic.List[object]]::new()

    foreach ($target in $Targets) {
        if (-not ($target.PSObject.Properties.Name -contains "HandlerPreview")) {
            continue
        }

        $handlerPreview = $target.HandlerPreview
        if ($null -eq $handlerPreview) {
            continue
        }

        $registryTargets = @()
        if ($handlerPreview.PSObject.Properties.Name -contains "HostRuleSummary") {
            $registryTargets = @($handlerPreview.HostRuleSummary.RegistryTargets)
        }
        elseif ($handlerPreview.PSObject.Properties.Name -contains "RegistryTarget" -and -not [string]::IsNullOrWhiteSpace($handlerPreview.RegistryTarget)) {
            $registryTargets = @([string]$handlerPreview.RegistryTarget)
        }

        $manifest.Add([pscustomobject]@{
            HostLabel            = ("{0} {1}" -f $target.HostName, $target.HostBitness)
            HostName             = $target.HostName
            HostBitness          = $target.HostBitness
            Operation            = $target.Operation
            ExecutionMode        = "PreviewOnly"
            DispatchHandler      = $target.DispatchHandler
            RegAsmInvoker        = $handlerPreview.RegAsmInvoker
            RegAsmToolPath       = if ($handlerPreview.PSObject.Properties.Name -contains "RegAsmPreview") { $handlerPreview.RegAsmPreview.ToolPath } else { $null }
            RegistryTargets      = @($registryTargets)
            PlannedActions       = @($target.PlannedActions)
            LiveExecutionAllowed = if ($handlerPreview.PSObject.Properties.Name -contains "HostRuleSummary") { [bool]$handlerPreview.HostRuleSummary.LiveExecutionAllowed } else { $false }
        })
    }

    return @($manifest.ToArray())
}

function Get-InstallerHandoffSummary {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProbeResult,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$OperationManifest
    )

    $supportedHosts = [System.Collections.Generic.List[string]]::new()
    $unsupportedHosts = [System.Collections.Generic.List[string]]::new()

    foreach ($detectedHost in $ProbeResult.Hosts) {
        $hostLabel = "{0} {1}" -f $detectedHost.HostName, $detectedHost.HostBitness
        if ($detectedHost.SupportStatus -eq "supported") {
            $supportedHosts.Add($hostLabel)
        }
        else {
            $unsupportedHosts.Add($hostLabel)
        }
    }

    $liveCapableHosts = @(
        $OperationManifest |
            Where-Object { $_.LiveExecutionAllowed } |
            ForEach-Object { $_.HostLabel }
    )

    $decision = if ($liveCapableHosts.Count -gt 0) {
        "Supported hosts are available for shared-core live registration through the rerouted installer and script entrypoints."
    }
    elseif ($unsupportedHosts.Count -gt 0) {
        "Detected hosts are not currently supported for live registration and remain preview-only."
    }
    else {
        "No supported hosts were detected, so the installer handoff remains a preview-only no-op."
    }

    return [pscustomobject]@{
        ExecutionMode       = "PreviewOnly"
        DetectedHostCount   = $ProbeResult.Hosts.Count
        SupportedHostCount  = $supportedHosts.Count
        UnsupportedHostCount = $unsupportedHosts.Count
        PreviewActionCount  = $OperationManifest.Count
        SupportedHosts      = @($supportedHosts.ToArray())
        UnsupportedHosts    = @($unsupportedHosts.ToArray())
        LiveCapableHosts    = @($liveCapableHosts)
        UserFacingDecision  = $decision
    }
}

function Get-LiveEntrypointStatus {
    return [pscustomobject]@{
        ExecutionMode                  = "PreviewOnly"
        SharedCoreOwnsLiveRegistration = $true
        ReroutedEntrypoints            = @("RegisterPlugin.ps1", "RegisterPlugin.bat", "Setup.iss")
        PendingEntrypoints             = @()
        CurrentEntrypoints             = @("RegisterPlugin.ps1", "RegisterPlugin.bat", "Setup.iss")
        MigrationDecision              = "RegisterPlugin.ps1, RegisterPlugin.bat, and Setup.iss now delegate live registration to Installer.Core.ps1 for currently supported hosts."
    }
}

function Get-MigrationChecklist {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProbeResult,

        [Parameter(Mandatory = $true)]
        [object]$LiveReadinessSummary,

        [Parameter(Mandatory = $true)]
        [object]$LiveEntrypointStatus
    )

    $blockingItems = [System.Collections.Generic.List[string]]::new()

    $deferredSupportTargets = [System.Collections.Generic.List[string]]::new()

    foreach ($pendingEntrypoint in $LiveEntrypointStatus.PendingEntrypoints) {
        $blockingItems.Add("$pendingEntrypoint has not been rerouted to Installer.Core.ps1 yet.")
    }

    if ($LiveReadinessSummary.PreviewOnlyHostCount -gt 0) {
        $blockingItems.Add("Detected supported hosts are still gated behind preview-only enforcement.")
    }

    foreach ($target in $ProbeResult.SupportState.targets) {
        if ($target.status -eq "planned") {
            $deferredSupportTargets.Add(("{0} {1}" -f $target.host, $target.bitness))
        }
    }

    $readyToReroute = ($blockingItems.Count -eq 0)
    $overallDecision = if ($readyToReroute) {
        "Live entrypoint reroute is complete for currently supported hosts. Remaining planned targets still require probe evidence before support expansion."
    } else {
        "The unified installer core now owns the live script path for validated hosts, but the full installer cutover is not ready yet."
    }

    return [pscustomobject]@{
        ExecutionMode                  = "PreviewOnly"
        ReadyToRerouteLiveEntrypoints = $readyToReroute
        BlockingItems                  = @($blockingItems.ToArray())
        DeferredSupportTargets         = @($deferredSupportTargets.ToArray())
        OverallDecision                = $overallDecision
    }
}

function Get-InstallerPreviewReport {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Operation,

        [Parameter(Mandatory = $true)]
        [object]$ProbeResult,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$OperationManifest,

        [Parameter(Mandatory = $true)]
        [object]$InstallerHandoffSummary,

        [Parameter(Mandatory = $true)]
        [object]$MigrationChecklist
    )

    $supportedHosts = @($InstallerHandoffSummary.SupportedHosts)
    $supportedHostText = if ($supportedHosts.Count -gt 0) { $supportedHosts -join ", " } else { "none" }

    return [pscustomobject]@{
        ExecutionMode                = "PreviewOnly"
        Operation                    = $Operation
        DetectedHostCount            = $ProbeResult.Hosts.Count
        PreviewActionCount           = $OperationManifest.Count
        SupportedHostCount           = $InstallerHandoffSummary.SupportedHostCount
        ReadyToRerouteLiveEntrypoints = $MigrationChecklist.ReadyToRerouteLiveEntrypoints
        SummaryText                  = "Detected hosts: $($ProbeResult.Hosts.Count). Supported hosts: $supportedHostText. Preview actions: $($OperationManifest.Count)."
    }
}

function Get-RegisterPreviewTarget {
    param(
        [Parameter(Mandatory = $true)]
        [object]$DryRunTarget
    )

    switch ($DryRunTarget.HostName) {
        "Word" { return Get-RegisterPreviewTargetForWordHost -DryRunTarget $DryRunTarget }
        "WPS" { return Get-RegisterPreviewTargetForWpsHost -DryRunTarget $DryRunTarget }
        default { return Get-RegisterPreviewTargetForWpsHost -DryRunTarget $DryRunTarget }
    }
}

function Get-RegisterPreviewTargetForWordHost {
    param(
        [Parameter(Mandatory = $true)]
        [object]$DryRunTarget
    )

    $wouldExecute = [bool]$DryRunTarget.WouldRegister

    return [pscustomobject]@{
        HostName        = $DryRunTarget.HostName
        HostBitness     = $DryRunTarget.HostBitness
        SupportStatus   = $DryRunTarget.SupportStatus
        Operation       = "Register"
        ExecutionMode   = "PreviewOnly"
        DispatchHandler = "Register-WordHost"
        HandlerPreview  = Register-WordHost -HostTarget $DryRunTarget -PreviewOnly
        WouldExecute    = $wouldExecute
        ProgId          = $DryRunTarget.ProgId
        RegAsmPath      = $DryRunTarget.RegAsmPath
        RegistryPath    = $DryRunTarget.RegistryPath
        PlannedActions  = $DryRunTarget.PlannedActions
        DecisionReason  = if ($wouldExecute) {
            "Preview-only register mode would execute the supported registration path, but live registration remains disabled in the shared core."
        }
        else {
            "Preview-only register mode would skip this Word target because it is not live-registration-eligible yet."
        }
    }
}

function Get-RegisterPreviewTargetForWpsHost {
    param(
        [Parameter(Mandatory = $true)]
        [object]$DryRunTarget
    )

    return [pscustomobject]@{
        HostName        = $DryRunTarget.HostName
        HostBitness     = $DryRunTarget.HostBitness
        SupportStatus   = $DryRunTarget.SupportStatus
        Operation       = "Register"
        ExecutionMode   = "PreviewOnly"
        DispatchHandler = "Register-WpsHost"
        HandlerPreview  = Register-WpsHost -HostTarget $DryRunTarget -PreviewOnly
        WouldExecute    = $false
        ProgId          = $DryRunTarget.ProgId
        RegAsmPath      = $DryRunTarget.RegAsmPath
        RegistryPath    = $DryRunTarget.RegistryPath
        PlannedActions  = @("SkipRegistration")
        DecisionReason  = "Preview-only register mode would route WPS through a dedicated WPS handler in the future, but live WPS registration is not enabled yet."
    }
}

function Get-UnregisterPreviewTarget {
    param(
        [Parameter(Mandatory = $true)]
        [object]$DetectedHost
    )

    switch ($DetectedHost.HostName) {
        "Word" { return Get-UnregisterPreviewTargetForWordHost -DetectedHost $DetectedHost }
        "WPS" { return Get-UnregisterPreviewTargetForWpsHost -DetectedHost $DetectedHost }
        default { return Get-UnregisterPreviewTargetForWpsHost -DetectedHost $DetectedHost }
    }
}

function Get-UnregisterPreviewTargetForWordHost {
    param(
        [Parameter(Mandatory = $true)]
        [object]$DetectedHost
    )

    $progId = "WordTools.ThisAddIn"
    $regAsmPath = Get-RegAsmPath -HostBitness $DetectedHost.HostBitness
    $registryPath = Get-AddInRegistryPath -HostName $DetectedHost.HostName
    $wouldExecute = $DetectedHost.SupportStatus -eq "supported" -and -not [string]::IsNullOrWhiteSpace($regAsmPath) -and -not [string]::IsNullOrWhiteSpace($registryPath)

    return [pscustomobject]@{
        HostName        = $DetectedHost.HostName
        HostBitness     = $DetectedHost.HostBitness
        SupportStatus   = $DetectedHost.SupportStatus
        Operation       = "Unregister"
        ExecutionMode   = "PreviewOnly"
        DispatchHandler = "Unregister-WordHost"
        HandlerPreview  = Unregister-WordHost -HostTarget $DetectedHost -PreviewOnly
        WouldExecute    = $wouldExecute
        ProgId          = $progId
        RegAsmPath      = $regAsmPath
        RegistryPath    = $registryPath
        PlannedActions  = if ($wouldExecute) {
            @("UnregisterComCodebase", "RemoveWordAddInRegistry", "OptionallyUninstallNgen")
        }
        else {
            @("SkipUnregister")
        }
        DecisionReason  = if ($wouldExecute) {
            "Preview-only unregister mode would remove the supported Word registration path, but live unregistration remains disabled in the shared core."
        }
        else {
            "Preview-only unregister mode would skip this Word target because it is not live-unregistration-eligible yet."
        }
    }
}

function Get-UnregisterPreviewTargetForWpsHost {
    param(
        [Parameter(Mandatory = $true)]
        [object]$DetectedHost
    )

    return [pscustomobject]@{
        HostName        = $DetectedHost.HostName
        HostBitness     = $DetectedHost.HostBitness
        SupportStatus   = $DetectedHost.SupportStatus
        Operation       = "Unregister"
        ExecutionMode   = "PreviewOnly"
        DispatchHandler = "Unregister-WpsHost"
        HandlerPreview  = Unregister-WpsHost -HostTarget $DetectedHost -PreviewOnly
        WouldExecute    = $false
        ProgId          = "WordTools.ThisAddIn"
        RegAsmPath      = Get-RegAsmPath -HostBitness $DetectedHost.HostBitness
        RegistryPath    = Get-AddInRegistryPath -HostName $DetectedHost.HostName
        PlannedActions  = @("SkipUnregister")
        DecisionReason  = "Preview-only unregister mode would route WPS through a dedicated WPS handler in the future, but live WPS unregistration is not enabled yet."
    }
}

function Invoke-RegAsm32 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DllPath,

        [switch]$Unregister,

        [switch]$PreviewOnly
    )

    $toolPath = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
    $arguments = if ($Unregister) { "/unregister `"$DllPath`"" } else { "/codebase `"$DllPath`"" }

    if ($PreviewOnly) {
        return [pscustomobject]@{
            InvokerName   = "Invoke-RegAsm32"
            ExecutionMode = "PreviewOnly"
            ToolPath      = $toolPath
            WouldRun      = $false
            DllPath       = $DllPath
            Arguments     = $arguments
        }
    }

    if (-not (Test-Path -LiteralPath $toolPath)) {
        throw "RegAsm tool not found: $toolPath"
    }

    if (-not (Test-Path -LiteralPath $DllPath)) {
        throw "Plugin DLL not found: $DllPath"
    }

    $toolInvocation = if ($Unregister) {
        Invoke-NativeToolCapture -ToolPath $toolPath -ArgumentList @("/unregister", $DllPath)
    }
    else {
        Invoke-NativeToolCapture -ToolPath $toolPath -ArgumentList @("/codebase", $DllPath)
    }

    $exitCode = [int]$toolInvocation.ExitCode
    if ($exitCode -ne 0) {
        throw "RegAsm32 failed with exit code $exitCode. $($toolInvocation.CombinedOutput)".Trim()
    }

    return [pscustomobject]@{
        InvokerName   = "Invoke-RegAsm32"
        ExecutionMode = "Live"
        ToolPath      = $toolPath
        WouldRun      = $true
        DllPath       = $DllPath
        Arguments     = $arguments
        ExitCode      = $exitCode
        ToolOutput    = $toolInvocation.CombinedOutput
    }
}

function Invoke-RegAsm64 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DllPath,

        [switch]$Unregister,

        [switch]$PreviewOnly
    )

    $toolPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
    $arguments = if ($Unregister) { "/unregister `"$DllPath`"" } else { "/codebase `"$DllPath`"" }

    if ($PreviewOnly) {
        return [pscustomobject]@{
            InvokerName   = "Invoke-RegAsm64"
            ExecutionMode = "PreviewOnly"
            ToolPath      = $toolPath
            WouldRun      = $false
            DllPath       = $DllPath
            Arguments     = $arguments
        }
    }

    if (-not (Test-Path -LiteralPath $toolPath)) {
        throw "RegAsm tool not found: $toolPath"
    }

    if (-not (Test-Path -LiteralPath $DllPath)) {
        throw "Plugin DLL not found: $DllPath"
    }

    $toolInvocation = if ($Unregister) {
        Invoke-NativeToolCapture -ToolPath $toolPath -ArgumentList @("/unregister", $DllPath)
    }
    else {
        Invoke-NativeToolCapture -ToolPath $toolPath -ArgumentList @("/codebase", $DllPath)
    }

    $exitCode = [int]$toolInvocation.ExitCode
    if ($exitCode -ne 0) {
        throw "RegAsm64 failed with exit code $exitCode. $($toolInvocation.CombinedOutput)".Trim()
    }

    return [pscustomobject]@{
        InvokerName   = "Invoke-RegAsm64"
        ExecutionMode = "Live"
        ToolPath      = $toolPath
        WouldRun      = $true
        DllPath       = $DllPath
        Arguments     = $arguments
        ExitCode      = $exitCode
        ToolOutput    = $toolInvocation.CombinedOutput
    }
}

function Set-WordAddInRegistry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RegistryPath
    )

    if (-not (Test-Path -LiteralPath $RegistryPath)) {
        New-Item -Path $RegistryPath -Force | Out-Null
    }

    New-ItemProperty -Path $RegistryPath -Name "FriendlyName" -Value "WordTools" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $RegistryPath -Name "Description" -Value "WordTools COM Add-in for Microsoft Word" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $RegistryPath -Name "LoadBehavior" -Value 3 -PropertyType DWORD -Force | Out-Null
    New-ItemProperty -Path $RegistryPath -Name "CommandLineSafe" -Value 0 -PropertyType DWORD -Force | Out-Null

    return [pscustomobject]@{
        RegistryPath   = $RegistryPath
        ExecutionMode  = "Live"
        ValuesWritten  = @("FriendlyName", "Description", "LoadBehavior", "CommandLineSafe")
    }
}

function Remove-WordAddInRegistry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RegistryPath
    )

    $removed = $false
    if (Test-Path -LiteralPath $RegistryPath) {
        Remove-Item -LiteralPath $RegistryPath -Recurse -Force
        $removed = $true
    }

    return [pscustomobject]@{
        RegistryPath  = $RegistryPath
        ExecutionMode = "Live"
        Removed       = $removed
    }
}

function Set-WpsAddInRegistry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RegistryPath
    )

    if (-not (Test-Path -LiteralPath $RegistryPath)) {
        New-Item -Path $RegistryPath -Force | Out-Null
    }

    New-ItemProperty -Path $RegistryPath -Name "FriendlyName" -Value "WordTools" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $RegistryPath -Name "Description" -Value "WordTools COM Add-in for WPS" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $RegistryPath -Name "LoadBehavior" -Value 3 -PropertyType DWORD -Force | Out-Null
    New-ItemProperty -Path $RegistryPath -Name "CommandLineSafe" -Value 0 -PropertyType DWORD -Force | Out-Null

    return [pscustomobject]@{
        RegistryPath  = $RegistryPath
        ExecutionMode = "Live"
        ValuesWritten = @("FriendlyName", "Description", "LoadBehavior", "CommandLineSafe")
    }
}

function Remove-WpsAddInRegistry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RegistryPath
    )

    $removed = $false
    if (Test-Path -LiteralPath $RegistryPath) {
        Remove-Item -LiteralPath $RegistryPath -Recurse -Force
        $removed = $true
    }

    return [pscustomobject]@{
        RegistryPath  = $RegistryPath
        ExecutionMode = "Live"
        Removed       = $removed
    }
}

function Invoke-NgenForHost {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DllPath,

        [Parameter(Mandatory = $true)]
        [string]$HostBitness,

        [switch]$Uninstall
    )

    $ngenPath = Get-NgenPath -HostBitness $HostBitness
    if ([string]::IsNullOrWhiteSpace($ngenPath) -or -not (Test-Path -LiteralPath $ngenPath)) {
        return [pscustomobject]@{
            ExecutionMode = "Live"
            ToolPath      = $ngenPath
            Attempted     = $false
            Succeeded     = $false
            SkippedReason = "NGen tool not found."
        }
    }

    $toolInvocation = if ($Uninstall) {
        Invoke-NativeToolCapture -ToolPath $ngenPath -ArgumentList @("uninstall", $DllPath)
    }
    else {
        Invoke-NativeToolCapture -ToolPath $ngenPath -ArgumentList @("install", $DllPath)
    }

    $exitCode = [int]$toolInvocation.ExitCode

    return [pscustomobject]@{
        ExecutionMode = "Live"
        ToolPath      = $ngenPath
        Attempted     = $true
        Succeeded     = ($exitCode -eq 0)
        ExitCode      = $exitCode
        Action        = if ($Uninstall) { "Uninstall" } else { "Install" }
        ToolOutput    = $toolInvocation.CombinedOutput
    }
}

function Get-HostRuleEnablementCondition {
    param(
        [Parameter(Mandatory = $true)]
        [object]$HostTarget,

        [Parameter(Mandatory = $true)]
        [string]$Operation,

        [Parameter(Mandatory = $true)]
        [bool]$LiveExecutionAllowed
    )

    $supportStatus = if ($HostTarget.PSObject.Properties.Name -contains "SupportStatus") {
        [string]$HostTarget.SupportStatus
    }
    else {
        "unknown"
    }

    if ($HostTarget.HostName -eq "Word") {
        if ($LiveExecutionAllowed) {
            return "Live $Operation for Word is enabled when the host remains supported and both the RegAsm path and Word add-in registry target resolve successfully. The shared core can execute live Word registration for the validated supported Word path when invoked through a live entrypoint."
        }

        return "Live $Operation for Word is only eligible when the host remains supported and both the RegAsm path and Word add-in registry target resolve successfully. This host is not currently eligible for live execution through the shared core."
    }

    if ($HostTarget.HostBitness -eq "x64") {
        return "Live $Operation for WPS x64 remains disabled until separate x64 probe and live-validation evidence is collected."
    }

    return "Live $Operation for WPS remains disabled until the real WPS add-in write contract and UI-load behavior are validated on-machine. This target is not currently eligible for live execution through the shared core."
}

function New-HostRuleSummary {
    param(
        [Parameter(Mandatory = $true)]
        [object]$HostTarget,

        [Parameter(Mandatory = $true)]
        [string]$Operation,

        [Parameter(Mandatory = $true)]
        [string]$HandlerName,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$PreferredRegAsmInvoker,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$RegistryTargets,

        [Parameter(Mandatory = $true)]
        [bool]$LiveExecutionAllowed
    )

    $supportStatus = if ($HostTarget.PSObject.Properties.Name -contains "SupportStatus") {
        [string]$HostTarget.SupportStatus
    }
    else {
        "unknown"
    }

    $normalizedRegistryTargets = @(
        $RegistryTargets | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    return [pscustomobject]@{
        ExecutionMode         = "PreviewOnly"
        HostName              = $HostTarget.HostName
        HostBitness           = $HostTarget.HostBitness
        VersionLine           = if ($HostTarget.PSObject.Properties.Name -contains "VersionLine") { [string]$HostTarget.VersionLine } else { "unknown" }
        Operation             = $Operation
        HandlerName           = $HandlerName
        CurrentSupportStatus  = $supportStatus
        EnablementCondition   = Get-HostRuleEnablementCondition -HostTarget $HostTarget -Operation $Operation -LiveExecutionAllowed $LiveExecutionAllowed
        PreferredRegAsmInvoker = $PreferredRegAsmInvoker
        RegistryTargets       = $normalizedRegistryTargets
        LiveExecutionAllowed  = $LiveExecutionAllowed
    }
}

function Register-WordHost {
    param(
        [Parameter(Mandatory = $true)]
        [object]$HostTarget,

        [string]$DllPath = "<WordTools.dll>",

        [switch]$PreviewOnly
    )

    $liveExecutionAllowed = Test-LiveExecutionAllowedForHost -HostTarget $HostTarget
    $regAsmPreview = if ($HostTarget.HostBitness -eq "x86") {
        Invoke-RegAsm32 -DllPath $DllPath -PreviewOnly
    }
    else {
        Invoke-RegAsm64 -DllPath $DllPath -PreviewOnly
    }
    $registryTarget = "HKLM:\Software\Microsoft\Office\Word\Addins\WordTools.ThisAddIn"

    if ($PreviewOnly) {
        return [pscustomobject]@{
            HandlerName     = "Register-WordHost"
            ExecutionMode   = "PreviewOnly"
            RegAsmInvoker   = $regAsmPreview.InvokerName
            RegAsmPreview   = $regAsmPreview
            RegistryTarget  = $registryTarget
            HostRuleSummary = New-HostRuleSummary -HostTarget $HostTarget -Operation "Register" -HandlerName "Register-WordHost" -PreferredRegAsmInvoker $regAsmPreview.InvokerName -RegistryTargets @($registryTarget) -LiveExecutionAllowed $liveExecutionAllowed
            WouldRun        = $false
        }
    }

    if (-not $liveExecutionAllowed) {
        throw "Live Word registration is not allowed for $($HostTarget.HostName) $($HostTarget.HostBitness)."
    }

    $regAsmResult = if ($HostTarget.HostBitness -eq "x86") {
        Invoke-RegAsm32 -DllPath $DllPath
    }
    else {
        Invoke-RegAsm64 -DllPath $DllPath
    }

    $registryResult = Set-WordAddInRegistry -RegistryPath $registryTarget
    $ngenResult = Invoke-NgenForHost -DllPath $DllPath -HostBitness $HostTarget.HostBitness

    return [pscustomobject]@{
        HandlerName     = "Register-WordHost"
        ExecutionMode   = "Live"
        RegAsmInvoker   = $regAsmResult.InvokerName
        RegAsmResult    = $regAsmResult
        RegistryTarget  = $registryTarget
        RegistryResult  = $registryResult
        NgenResult      = $ngenResult
        HostRuleSummary = New-HostRuleSummary -HostTarget $HostTarget -Operation "Register" -HandlerName "Register-WordHost" -PreferredRegAsmInvoker $regAsmResult.InvokerName -RegistryTargets @($registryTarget) -LiveExecutionAllowed $liveExecutionAllowed
        WouldRun        = $true
    }
}

function Register-WpsHost {
    param(
        [Parameter(Mandatory = $true)]
        [object]$HostTarget,

        [string]$DllPath = "<WordTools.dll>",

        [switch]$PreviewOnly
    )

    $liveExecutionAllowed = Test-LiveExecutionAllowedForHost -HostTarget $HostTarget
    $regAsmPreview = if ($HostTarget.HostBitness -eq "x86") {
        Invoke-RegAsm32 -DllPath $DllPath -PreviewOnly
    }
    else {
        Invoke-RegAsm64 -DllPath $DllPath -PreviewOnly
    }
    $registryTarget = Get-AddInRegistryPath -HostName "WPS"

    if ($PreviewOnly) {
        return [pscustomobject]@{
            HandlerName     = "Register-WpsHost"
            ExecutionMode   = "PreviewOnly"
            RegAsmInvoker   = $regAsmPreview.InvokerName
            RegAsmPreview   = $regAsmPreview
            RegistryTarget  = $registryTarget
            HostRuleSummary = New-HostRuleSummary -HostTarget $HostTarget -Operation "Register" -HandlerName "Register-WpsHost" -PreferredRegAsmInvoker $regAsmPreview.InvokerName -RegistryTargets @($registryTarget) -LiveExecutionAllowed $liveExecutionAllowed
            WouldRun        = $false
        }
    }

    if (-not $liveExecutionAllowed) {
        throw "Live WPS registration is not allowed for $($HostTarget.HostName) $($HostTarget.HostBitness)."
    }

    $regAsmResult = if ($HostTarget.HostBitness -eq "x86") {
        Invoke-RegAsm32 -DllPath $DllPath
    }
    else {
        Invoke-RegAsm64 -DllPath $DllPath
    }

    $registryResult = Set-WpsAddInRegistry -RegistryPath $registryTarget
    $ngenResult = Invoke-NgenForHost -DllPath $DllPath -HostBitness $HostTarget.HostBitness

    return [pscustomobject]@{
        HandlerName     = "Register-WpsHost"
        ExecutionMode   = "Live"
        RegAsmInvoker   = $regAsmResult.InvokerName
        RegAsmResult    = $regAsmResult
        RegistryTarget  = $registryTarget
        RegistryResult  = $registryResult
        NgenResult      = $ngenResult
        HostRuleSummary = New-HostRuleSummary -HostTarget $HostTarget -Operation "Register" -HandlerName "Register-WpsHost" -PreferredRegAsmInvoker $regAsmResult.InvokerName -RegistryTargets @($registryTarget) -LiveExecutionAllowed $liveExecutionAllowed
        WouldRun        = $true
    }
}

function Unregister-WordHost {
    param(
        [Parameter(Mandatory = $true)]
        [object]$HostTarget,

        [string]$DllPath = "<WordTools.dll>",

        [switch]$PreviewOnly
    )

    $liveExecutionAllowed = Test-LiveExecutionAllowedForHost -HostTarget $HostTarget
    $regAsmPreview = if ($HostTarget.HostBitness -eq "x86") {
        Invoke-RegAsm32 -DllPath $DllPath -Unregister -PreviewOnly
    }
    else {
        Invoke-RegAsm64 -DllPath $DllPath -Unregister -PreviewOnly
    }
    $registryTarget = "HKLM:\Software\Microsoft\Office\Word\Addins\WordTools.ThisAddIn"

    if ($PreviewOnly) {
        return [pscustomobject]@{
            HandlerName     = "Unregister-WordHost"
            ExecutionMode   = "PreviewOnly"
            RegAsmInvoker   = $regAsmPreview.InvokerName
            RegAsmPreview   = $regAsmPreview
            RegistryTarget  = $registryTarget
            HostRuleSummary = New-HostRuleSummary -HostTarget $HostTarget -Operation "Unregister" -HandlerName "Unregister-WordHost" -PreferredRegAsmInvoker $regAsmPreview.InvokerName -RegistryTargets @($registryTarget) -LiveExecutionAllowed $liveExecutionAllowed
            WouldRun        = $false
        }
    }

    if (-not $liveExecutionAllowed) {
        throw "Live Word unregistration is not allowed for $($HostTarget.HostName) $($HostTarget.HostBitness)."
    }

    $regAsmResult = if ($HostTarget.HostBitness -eq "x86") {
        Invoke-RegAsm32 -DllPath $DllPath -Unregister
    }
    else {
        Invoke-RegAsm64 -DllPath $DllPath -Unregister
    }

    $registryResult = Remove-WordAddInRegistry -RegistryPath $registryTarget
    $ngenResult = Invoke-NgenForHost -DllPath $DllPath -HostBitness $HostTarget.HostBitness -Uninstall

    return [pscustomobject]@{
        HandlerName     = "Unregister-WordHost"
        ExecutionMode   = "Live"
        RegAsmInvoker   = $regAsmResult.InvokerName
        RegAsmResult    = $regAsmResult
        RegistryTarget  = $registryTarget
        RegistryResult  = $registryResult
        NgenResult      = $ngenResult
        HostRuleSummary = New-HostRuleSummary -HostTarget $HostTarget -Operation "Unregister" -HandlerName "Unregister-WordHost" -PreferredRegAsmInvoker $regAsmResult.InvokerName -RegistryTargets @($registryTarget) -LiveExecutionAllowed $liveExecutionAllowed
        WouldRun        = $true
    }
}

function Unregister-WpsHost {
    param(
        [Parameter(Mandatory = $true)]
        [object]$HostTarget,

        [string]$DllPath = "<WordTools.dll>",

        [switch]$PreviewOnly
    )

    $liveExecutionAllowed = Test-LiveExecutionAllowedForHost -HostTarget $HostTarget
    $regAsmPreview = if ($HostTarget.HostBitness -eq "x86") {
        Invoke-RegAsm32 -DllPath $DllPath -Unregister -PreviewOnly
    }
    else {
        Invoke-RegAsm64 -DllPath $DllPath -Unregister -PreviewOnly
    }
    $registryTarget = Get-AddInRegistryPath -HostName "WPS"

    if ($PreviewOnly) {
        return [pscustomobject]@{
            HandlerName     = "Unregister-WpsHost"
            ExecutionMode   = "PreviewOnly"
            RegAsmInvoker   = $regAsmPreview.InvokerName
            RegAsmPreview   = $regAsmPreview
            RegistryTarget  = $registryTarget
            HostRuleSummary = New-HostRuleSummary -HostTarget $HostTarget -Operation "Unregister" -HandlerName "Unregister-WpsHost" -PreferredRegAsmInvoker $regAsmPreview.InvokerName -RegistryTargets @($registryTarget) -LiveExecutionAllowed $liveExecutionAllowed
            WouldRun        = $false
        }
    }

    if (-not $liveExecutionAllowed) {
        throw "Live WPS unregistration is not allowed for $($HostTarget.HostName) $($HostTarget.HostBitness)."
    }

    $regAsmResult = if ($HostTarget.HostBitness -eq "x86") {
        Invoke-RegAsm32 -DllPath $DllPath -Unregister
    }
    else {
        Invoke-RegAsm64 -DllPath $DllPath -Unregister
    }

    $registryResult = Remove-WpsAddInRegistry -RegistryPath $registryTarget
    $ngenResult = Invoke-NgenForHost -DllPath $DllPath -HostBitness $HostTarget.HostBitness -Uninstall

    return [pscustomobject]@{
        HandlerName     = "Unregister-WpsHost"
        ExecutionMode   = "Live"
        RegAsmInvoker   = $regAsmResult.InvokerName
        RegAsmResult    = $regAsmResult
        RegistryTarget  = $registryTarget
        RegistryResult  = $registryResult
        NgenResult      = $ngenResult
        HostRuleSummary = New-HostRuleSummary -HostTarget $HostTarget -Operation "Unregister" -HandlerName "Unregister-WpsHost" -PreferredRegAsmInvoker $regAsmResult.InvokerName -RegistryTargets @($registryTarget) -LiveExecutionAllowed $liveExecutionAllowed
        WouldRun        = $true
    }
}

function Get-PreviewSummary {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Targets,

        [Parameter(Mandatory = $true)]
        [string]$ExecutionVerb,

        [Parameter(Mandatory = $true)]
        [string]$NoEligibleDecision
    )

    $requiredRegAsmModes = [System.Collections.Generic.List[string]]::new()
    $registryWrites = [System.Collections.Generic.List[string]]::new()
    $plannedActionUnion = [System.Collections.Generic.List[string]]::new()
    $previewableTargetCount = 0
    $skippedTargetCount = 0

    foreach ($target in $Targets) {
        if ($target.WouldExecute) {
            $previewableTargetCount++

            if (($target.HostBitness -eq "x86" -or $target.HostBitness -eq "x64") -and -not $requiredRegAsmModes.Contains($target.HostBitness)) {
                $requiredRegAsmModes.Add($target.HostBitness)
            }

            if (-not [string]::IsNullOrWhiteSpace($target.RegistryPath) -and -not $registryWrites.Contains($target.RegistryPath)) {
                $registryWrites.Add($target.RegistryPath)
            }

            foreach ($action in $target.PlannedActions) {
                if (-not [string]::IsNullOrWhiteSpace($action) -and -not $plannedActionUnion.Contains($action)) {
                    $plannedActionUnion.Add($action)
                }
            }
        }
        else {
            $skippedTargetCount++
        }
    }

    $overallDecision = if ($previewableTargetCount -gt 0) {
        "Preview-only $ExecutionVerb mode found $previewableTargetCount eligible target without executing any live $ExecutionVerb action."
    }
    else {
        $NoEligibleDecision
    }

    return [pscustomobject]@{
        ExecutionMode        = "PreviewOnly"
        DetectedTargetCount  = $Targets.Count
        PreviewableTargetCount = $previewableTargetCount
        SkippedTargetCount   = $skippedTargetCount
        RequiredRegAsmModes  = $requiredRegAsmModes
        RegistryWrites       = $registryWrites
        PlannedActionUnion   = $plannedActionUnion
        OverallDecision      = $overallDecision
    }
}

function Get-SupportSummary {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Hosts,

        [Parameter(Mandatory = $true)]
        [object]$SupportMatrix
    )

    $detectedHosts = [System.Collections.Generic.List[string]]::new()
    $supportedHosts = [System.Collections.Generic.List[string]]::new()
    $plannedHosts = [System.Collections.Generic.List[string]]::new()
    $unmappedHosts = [System.Collections.Generic.List[string]]::new()
    $missingExpectedHosts = [System.Collections.Generic.List[string]]::new()
    $ambiguousHosts = [System.Collections.Generic.List[string]]::new()

    foreach ($detectedHost in $Hosts) {
        $hostLabel = "{0} {1}" -f $detectedHost.HostName, $detectedHost.HostBitness
        $detectedHosts.Add($hostLabel)

        switch ($detectedHost.SupportStatus) {
            "supported" { $supportedHosts.Add($hostLabel) }
            "planned" { $plannedHosts.Add($hostLabel) }
            "unknown" { $unmappedHosts.Add($hostLabel) }
        }

        if ($detectedHost.HostBitness -eq "unknown" -or $detectedHost.RegistrationView -eq "Unknown") {
            $ambiguousHosts.Add($hostLabel)
        }
    }

    foreach ($target in $SupportMatrix.targets) {
        $expectedLabel = "{0} {1}" -f $target.host, $target.bitness
        if (-not $detectedHosts.Contains($expectedLabel)) {
            $missingExpectedHosts.Add($expectedLabel)
        }
    }

    return [pscustomobject]@{
        DetectedHosts        = $detectedHosts
        SupportedHosts       = $supportedHosts
        PlannedHosts         = $plannedHosts
        UnmappedHosts        = $unmappedHosts
        MissingExpectedHosts = $missingExpectedHosts
        AmbiguousHosts       = $ambiguousHosts
    }
}

function Get-HostInventory {
    $supportMatrix = Get-SupportMatrix
    $inventory = [System.Collections.Generic.List[object]]::new()
    $allCandidates = @()
    $allCandidates += Get-InstalledWordCandidates
    $allCandidates += Get-InstalledWpsCandidates

    foreach ($candidate in $allCandidates) {
        $bitness = Get-ExecutableBitness -ExecutablePath $candidate.ExecutablePath
        $supportDecision = Get-SupportDecision -SupportMatrix $supportMatrix -HostName $candidate.HostName -HostBitness $bitness

        $inventory.Add([pscustomobject]@{
                HostName         = $candidate.HostName
                ExecutablePath   = $candidate.ExecutablePath
                ProbeSource      = $candidate.ProbeSource
                DetectionReason  = "Detected from " + $candidate.ProbeSource
                HostBitness      = $bitness
                VersionLine      = Get-ExecutableVersionLine -ExecutablePath $candidate.ExecutablePath
                SupportState     = $supportDecision.SupportStatus
                RegistrationView = $supportDecision.RegistrationView
                SupportStatus    = $supportDecision.SupportStatus
                ValidationStage  = $supportDecision.ValidationStage
                ActivationRoute  = $supportDecision.ActivationRoute
                InstallState     = "detected"
                UiEvidenceState  = Get-UiEvidenceState -ValidationStage $supportDecision.ValidationStage
                P0EvidenceState  = Get-P0EvidenceState -ValidationStage $supportDecision.ValidationStage
                DiagnosticsBundleId = Get-DiagnosticsBundleId -HostName $candidate.HostName -HostBitness $bitness -EvidenceLabel $EvidenceLabel
                SupportReason    = $supportDecision.SupportReason
                AmbiguityReason  = $supportDecision.AmbiguityReason
                WpsRecon         = if ($candidate.HostName -eq "WPS") { Get-WpsReconData -ExecutablePath $candidate.ExecutablePath -HostBitness $bitness } else { $null }
            })
    }

    return [pscustomobject]@{
        ProbeMode      = $Mode
        EvidenceLabel  = $EvidenceLabel
        ProbedAtUtc    = [DateTime]::UtcNow.ToString("o")
        SupportState   = $supportMatrix
        SupportSummary = Get-SupportSummary -Hosts $inventory -SupportMatrix $supportMatrix
        Hosts          = $inventory
    }
}

function Get-LiveEligibleHosts {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProbeResult,

        [Parameter(Mandatory = $true)]
        [string]$RequestedArchitecture,

        [Parameter(Mandatory = $true)]
        [string]$RequestedHost
    )

    $resolvedArchitecture = Resolve-RequestedArchitecture -RequestedArchitecture $RequestedArchitecture
    $requestedHosts = Get-RequestedHostNames -RequestedHost $RequestedHost
    $eligibleHosts = [System.Collections.Generic.List[object]]::new()

    foreach ($detectedHost in $ProbeResult.Hosts) {
        if (-not $requestedHosts.Contains($detectedHost.HostName)) {
            continue
        }

        if (-not (Test-RequestedArchitectureMatchesHost -RequestedArchitecture $RequestedArchitecture -HostTarget $detectedHost)) {
            continue
        }

        if (-not (Test-LiveExecutionAllowedForHost -HostTarget $detectedHost)) {
            continue
        }

        $eligibleHosts.Add($detectedHost)
    }

    return [pscustomobject]@{
        ResolvedArchitecture = $resolvedArchitecture
        RequestedHosts       = $requestedHosts
        EligibleHosts        = @($eligibleHosts.ToArray())
    }
}

function Get-InstallerStateRegistryPath {
    return "HKLM:\Software\WordTools\InstallerState"
}

function Save-InstallerState {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ExecutionResult
    )

    $registryPath = Get-InstallerStateRegistryPath
    if (-not (Test-Path -LiteralPath $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }

    $hostLabels = @($ExecutionResult.Targets | ForEach-Object { "{0} {1}" -f $_.HostRuleSummary.HostName, $_.HostRuleSummary.HostBitness })
    $regAsmModes = @($ExecutionResult.Targets | ForEach-Object {
            if ($_.RegAsmInvoker -eq "Invoke-RegAsm32") { "x86" }
            elseif ($_.RegAsmInvoker -eq "Invoke-RegAsm64") { "x64" }
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    $registryTargets = @($ExecutionResult.Targets | ForEach-Object { $_.RegistryTarget } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    $recordedVersionLines = @($ExecutionResult.Targets | ForEach-Object {
            if ($_.HostRuleSummary.PSObject.Properties.Name -contains "VersionLine" -and -not [string]::IsNullOrWhiteSpace([string]$_.HostRuleSummary.VersionLine)) {
                "{0} {1}={2}" -f $_.HostRuleSummary.HostName, $_.HostRuleSummary.HostBitness, [string]$_.HostRuleSummary.VersionLine
            }
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)

    New-ItemProperty -Path $registryPath -Name "RecordedTargets" -Value ($hostLabels -join ";") -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name "RegAsmModes" -Value ($regAsmModes -join ";") -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name "RegistryTargets" -Value ($registryTargets -join ";") -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name "RecordedVersionLines" -Value ($recordedVersionLines -join ";") -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name "DllPath" -Value ([string]$ExecutionResult.DllPath) -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name "Configuration" -Value ([string]$ExecutionResult.Configuration) -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name "ExecutionTimestampUtc" -Value ([DateTime]::UtcNow.ToString("o")) -PropertyType String -Force | Out-Null
}

function Get-InstallerState {
    $registryPath = Get-InstallerStateRegistryPath
    if (-not (Test-Path -LiteralPath $registryPath)) {
        return $null
    }

    return [pscustomobject]@{
        RegistryPath     = $registryPath
        RecordedTargets  = [string](Get-ItemPropertyValue -Path $registryPath -Name "RecordedTargets" -ErrorAction SilentlyContinue)
        RegAsmModes      = [string](Get-ItemPropertyValue -Path $registryPath -Name "RegAsmModes" -ErrorAction SilentlyContinue)
        RegistryTargets  = [string](Get-ItemPropertyValue -Path $registryPath -Name "RegistryTargets" -ErrorAction SilentlyContinue)
        RecordedVersionLines = [string](Get-ItemPropertyValue -Path $registryPath -Name "RecordedVersionLines" -ErrorAction SilentlyContinue)
        DllPath          = [string](Get-ItemPropertyValue -Path $registryPath -Name "DllPath" -ErrorAction SilentlyContinue)
        Configuration    = [string](Get-ItemPropertyValue -Path $registryPath -Name "Configuration" -ErrorAction SilentlyContinue)
        ExecutionTimestampUtc = [string](Get-ItemPropertyValue -Path $registryPath -Name "ExecutionTimestampUtc" -ErrorAction SilentlyContinue)
    }
}

function Get-LiveEligibleHostsFromInstallerState {
    $state = Get-InstallerState
    $resolvedArchitecture = Resolve-RequestedArchitecture -RequestedArchitecture $Architecture
    $requestedHosts = Get-RequestedHostNames -RequestedHost $RequestedHost
    $supportMatrix = Get-SupportMatrix
    $eligibleHosts = [System.Collections.Generic.List[object]]::new()
    $recordedVersionLinesByLabel = @{}

    if ($null -eq $state -or [string]::IsNullOrWhiteSpace($state.RecordedTargets)) {
        return [pscustomobject]@{
            ResolvedArchitecture = $resolvedArchitecture
            RequestedHosts       = $requestedHosts
            EligibleHosts        = @()
        }
    }

    foreach ($versionEntry in ([string]$state.RecordedVersionLines -split ";" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $versionSeparatorIndex = $versionEntry.IndexOf("=")
        if ($versionSeparatorIndex -le 0 -or $versionSeparatorIndex -ge ($versionEntry.Length - 1)) {
            continue
        }

        $versionLabel = $versionEntry.Substring(0, $versionSeparatorIndex).Trim()
        $versionLine = $versionEntry.Substring($versionSeparatorIndex + 1).Trim()
        if ([string]::IsNullOrWhiteSpace($versionLabel) -or [string]::IsNullOrWhiteSpace($versionLine)) {
            continue
        }

        $recordedVersionLinesByLabel[$versionLabel] = $versionLine
    }

    foreach ($label in ($state.RecordedTargets -split ";" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $hostTarget = $null
        $hostName = $null
        $hostBitness = $null
        $separatorIndex = $label.LastIndexOf(" ")

        if ($separatorIndex -gt 0 -and $separatorIndex -lt ($label.Length - 1)) {
            $hostName = $label.Substring(0, $separatorIndex).Trim()
            $hostBitness = $label.Substring($separatorIndex + 1).Trim()
        }

        if ([string]::IsNullOrWhiteSpace($hostName) -or [string]::IsNullOrWhiteSpace($hostBitness)) {
            continue
        }

        $supportDecision = Get-SupportDecision -SupportMatrix $supportMatrix -HostName $hostName -HostBitness $hostBitness
        $diagnosticsEvidenceLabel = if (-not [string]::IsNullOrWhiteSpace([string]$state.ExecutionTimestampUtc)) {
            [string]$state.ExecutionTimestampUtc
        }
        else {
            [string]$state.Configuration
        }
        $hostTarget = [pscustomobject]@{
            HostName            = $hostName
            HostBitness         = $hostBitness
            VersionLine         = if ($recordedVersionLinesByLabel.ContainsKey($label)) { [string]$recordedVersionLinesByLabel[$label] } else { "unknown" }
            SupportState        = $supportDecision.SupportStatus
            SupportStatus       = $supportDecision.SupportStatus
            ValidationStage     = $supportDecision.ValidationStage
            ActivationRoute     = $supportDecision.ActivationRoute
            InstallState        = "recorded"
            UiEvidenceState     = Get-UiEvidenceState -ValidationStage $supportDecision.ValidationStage
            P0EvidenceState     = Get-P0EvidenceState -ValidationStage $supportDecision.ValidationStage
            DiagnosticsBundleId = Get-DiagnosticsBundleId -HostName $hostName -HostBitness $hostBitness -EvidenceLabel $diagnosticsEvidenceLabel
            SupportReason       = $supportDecision.SupportReason
            RegistrationView    = $supportDecision.RegistrationView
            AmbiguityReason     = $supportDecision.AmbiguityReason
        }

        if (-not $requestedHosts.Contains($hostTarget.HostName)) {
            continue
        }

        if (-not (Test-RequestedArchitectureMatchesHost -RequestedArchitecture $Architecture -HostTarget $hostTarget)) {
            continue
        }

        if (-not (Test-LiveExecutionAllowedForHost -HostTarget $hostTarget)) {
            continue
        }

        $eligibleHosts.Add($hostTarget)
    }

    return [pscustomobject]@{
        ResolvedArchitecture = $resolvedArchitecture
        RequestedHosts       = $requestedHosts
        EligibleHosts        = @($eligibleHosts.ToArray())
    }
}

function Clear-InstallerState {
    $registryPath = Get-InstallerStateRegistryPath
    if (Test-Path -LiteralPath $registryPath) {
        Remove-Item -LiteralPath $registryPath -Recurse -Force
    }
}

function Get-InstallerDecisionSummaryText {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProbeResult,

        [Parameter(Mandatory = $true)]
        [object]$PreviewPlan
    )

    $detectedHosts = @($ProbeResult.SupportSummary.DetectedHosts)
    $supportedHosts = @($ProbeResult.SupportSummary.SupportedHosts)
    $liveCapableTargets = @($PreviewPlan.OperationManifest | Where-Object { $_.LiveExecutionAllowed } | ForEach-Object { $_.HostLabel })
    $unsupportedDetectedTargets = @($PreviewPlan.InstallerHandoffSummary.UnsupportedHosts)
    $pendingTargets = @($ProbeResult.SupportSummary.MissingExpectedHosts)
    $wouldContinue = if ($liveCapableTargets.Count -gt 0) { "Yes" } else { "No" }
    $detectedHostsText = if ($detectedHosts.Count -gt 0) { $detectedHosts -join ", " } else { "(none)" }
    $detectedHostDetailsText = Get-DetectedHostDetailsText -ProbeResult $ProbeResult
    $supportedHostsText = if ($supportedHosts.Count -gt 0) { $supportedHosts -join ", " } else { "(none)" }
    $liveCapableTargetsText = if ($liveCapableTargets.Count -gt 0) { $liveCapableTargets -join ", " } else { "(none)" }
    $unsupportedDetectedTargetsText = if ($unsupportedDetectedTargets.Count -gt 0) { $unsupportedDetectedTargets -join ", " } else { "(none)" }
    $pendingTargetsText = if ($pendingTargets.Count -gt 0) { $pendingTargets -join ", " } else { "(none)" }

    $lines = @(
        ("DetectedHosts: {0}" -f $detectedHostsText),
        ("DetectedHostDetails: {0}" -f $detectedHostDetailsText),
        ("SupportedHosts: {0}" -f $supportedHostsText),
        ("InstallableTargets: {0}" -f $liveCapableTargetsText),
        ("SkippedDetectedTargets: {0}" -f $unsupportedDetectedTargetsText),
        ("PendingValidationTargets: {0}" -f $pendingTargetsText),
        ("WouldContinue: {0}" -f $wouldContinue),
        ("Summary: {0}" -f $PreviewPlan.InstallerHandoffSummary.UserFacingDecision)
    )

    return ($lines -join [Environment]::NewLine) + [Environment]::NewLine
}

function Get-DryRunDecisionSummaryText {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProbeResult,

        [Parameter(Mandatory = $true)]
        [object]$DryRunPlan
    )

    $detectedHosts = @($ProbeResult.SupportSummary.DetectedHosts)
    $supportedTargets = @($DryRunPlan.PlanSummary.SupportedTargetLabels)
    $registrableTargets = @(
        $DryRunPlan.Targets |
            Where-Object { $_.WouldRegister } |
            ForEach-Object { "{0} {1}" -f $_.HostName, $_.HostBitness }
    )
    $skippedTargets = @($DryRunPlan.PlanSummary.SkippedTargetLabels)
    $wouldContinue = if ($DryRunPlan.PlanSummary.RegistrableTargetCount -gt 0) { "Yes" } else { "No" }
    $detectedHostsText = if ($detectedHosts.Count -gt 0) { $detectedHosts -join ", " } else { "(none)" }
    $detectedHostDetailsText = Get-DetectedHostDetailsText -ProbeResult $ProbeResult
    $supportedTargetsText = if ($supportedTargets.Count -gt 0) { $supportedTargets -join ", " } else { "(none)" }
    $registrableTargetsText = if ($registrableTargets.Count -gt 0) { $registrableTargets -join ", " } else { "(none)" }
    $skippedTargetsText = if ($skippedTargets.Count -gt 0) { $skippedTargets -join ", " } else { "(none)" }

    $lines = @(
        ("DetectedHosts: {0}" -f $detectedHostsText),
        ("DetectedHostDetails: {0}" -f $detectedHostDetailsText),
        ("SupportedTargets: {0}" -f $supportedTargetsText),
        ("RegistrableTargets: {0}" -f $registrableTargetsText),
        ("SkippedTargets: {0}" -f $skippedTargetsText),
        ("WouldContinue: {0}" -f $wouldContinue),
        ("Summary: {0}" -f $DryRunPlan.PlanSummary.OverallDecision)
    )

    return ($lines -join [Environment]::NewLine) + [Environment]::NewLine
}

function Get-ProbeDecisionSummaryText {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProbeResult
    )

    $detectedHosts = @($ProbeResult.SupportSummary.DetectedHosts)
    $supportedHosts = @($ProbeResult.SupportSummary.SupportedHosts)
    $plannedHosts = @($ProbeResult.SupportSummary.PlannedHosts)
    $missingHosts = @($ProbeResult.SupportSummary.MissingExpectedHosts)
    $ambiguousHosts = @($ProbeResult.SupportSummary.AmbiguousHosts)
    $detectedHostsText = if ($detectedHosts.Count -gt 0) { $detectedHosts -join ", " } else { "(none)" }
    $detectedHostDetailsText = Get-DetectedHostDetailsText -ProbeResult $ProbeResult
    $supportedHostsText = if ($supportedHosts.Count -gt 0) { $supportedHosts -join ", " } else { "(none)" }
    $plannedHostsText = if ($plannedHosts.Count -gt 0) { $plannedHosts -join ", " } else { "(none)" }
    $missingHostsText = if ($missingHosts.Count -gt 0) { $missingHosts -join ", " } else { "(none)" }
    $ambiguousHostsText = if ($ambiguousHosts.Count -gt 0) { $ambiguousHosts -join ", " } else { "(none)" }

    $lines = @(
        ("DetectedHosts: {0}" -f $detectedHostsText),
        ("DetectedHostDetails: {0}" -f $detectedHostDetailsText),
        ("SupportedHosts: {0}" -f $supportedHostsText),
        ("PlannedHosts: {0}" -f $plannedHostsText),
        ("MissingExpectedHosts: {0}" -f $missingHostsText),
        ("AmbiguousHosts: {0}" -f $ambiguousHostsText),
        ("Summary: Probe completed. Review host-state details before changing support boundaries.")
    )

    return ($lines -join [Environment]::NewLine) + [Environment]::NewLine
}

function Invoke-LiveRegisterExecution {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProbeResult
    )

    $selection = Get-LiveEligibleHosts -ProbeResult $ProbeResult -RequestedArchitecture $Architecture -RequestedHost $RequestedHost
    $dllPath = Get-EffectiveDllPath -RequestedConfiguration $Configuration -RequestedOverride $DllPathOverride

    if (-not (Test-Path -LiteralPath $dllPath)) {
        throw "Plugin DLL not found: $dllPath"
    }

    if ($selection.EligibleHosts.Count -eq 0) {
        throw "No live-registration-eligible host matched the current request. Only supported Word hosts are currently allowed."
    }

    Assert-LiveExecutionAdministrator -Operation "Register"

    $targetResults = [System.Collections.Generic.List[object]]::new()
    foreach ($eligibleHost in $selection.EligibleHosts) {
        switch ($eligibleHost.HostName) {
            "Word" {
                $targetResults.Add((Register-WordHost -HostTarget $eligibleHost -DllPath $dllPath))
            }
            "WPS" {
                $targetResults.Add((Register-WpsHost -HostTarget $eligibleHost -DllPath $dllPath))
            }
            default {
                throw "Unsupported live register host: $($eligibleHost.HostName)"
            }
        }
    }

    $executionResult = [pscustomobject]@{
        Operation             = "Register"
        ExecutionMode         = "Live"
        RequestedArchitecture = $selection.ResolvedArchitecture
        RequestedHost         = $RequestedHost
        Configuration         = $Configuration
        DllPath               = $dllPath
        AppliedTargetCount    = $targetResults.Count
        Targets               = @($targetResults.ToArray())
        OverallDecision       = "Live registration completed for $($targetResults.Count) host."
    }

    Save-InstallerState -ExecutionResult $executionResult
    return $executionResult
}

function Invoke-LiveUnregisterExecution {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ProbeResult
    )

    $selection = Get-LiveEligibleHosts -ProbeResult $ProbeResult -RequestedArchitecture $Architecture -RequestedHost $RequestedHost
    $dllPath = Get-EffectiveDllPath -RequestedConfiguration $Configuration -RequestedOverride $DllPathOverride

    if (-not (Test-Path -LiteralPath $dllPath)) {
        throw "Plugin DLL not found: $dllPath"
    }

    if ($selection.EligibleHosts.Count -eq 0) {
        $selection = Get-LiveEligibleHostsFromInstallerState
    }

    if ($selection.EligibleHosts.Count -eq 0) {
        throw "No live-unregistration-eligible host matched the current request. Only supported Word hosts are currently allowed."
    }

    Assert-LiveExecutionAdministrator -Operation "Unregister"

    $targetResults = [System.Collections.Generic.List[object]]::new()
    foreach ($eligibleHost in $selection.EligibleHosts) {
        switch ($eligibleHost.HostName) {
            "Word" {
                $targetResults.Add((Unregister-WordHost -HostTarget $eligibleHost -DllPath $dllPath))
            }
            "WPS" {
                $targetResults.Add((Unregister-WpsHost -HostTarget $eligibleHost -DllPath $dllPath))
            }
            default {
                throw "Unsupported live unregister host: $($eligibleHost.HostName)"
            }
        }
    }

    $executionResult = [pscustomobject]@{
        Operation             = "Unregister"
        ExecutionMode         = "Live"
        RequestedArchitecture = $selection.ResolvedArchitecture
        RequestedHost         = $RequestedHost
        Configuration         = $Configuration
        DllPath               = $dllPath
        AppliedTargetCount    = $targetResults.Count
        Targets               = @($targetResults.ToArray())
        OverallDecision       = "Live unregistration completed for $($targetResults.Count) host."
    }

    Clear-InstallerState
    return $executionResult
}

switch ($Mode) {
    "Probe" {
        $probeResult = Get-HostInventory
        if ($AppendEvidenceMarkdown) {
            Append-ProbeEvidenceMarkdown -ProbeResult $probeResult -TargetPath $EvidenceMarkdownPath
        }

        $json = $probeResult | ConvertTo-Json -Depth 8
        Save-ProbeOutput -JsonText $json -TargetPath $OutputPath
        Save-TextOutput -Text (Get-ProbeDecisionSummaryText -ProbeResult $probeResult) -TargetPath $SummaryTextPath
        $json
        break
    }
    "Plan" {
        $probeResult = Get-HostInventory
        $requestedProbeResult = Get-RequestedProbeResult -ProbeResult $probeResult -RequestedArchitecture $Architecture -RequestedHost $RequestedHost
        $planResult = [pscustomobject]@{
            ProbeMode        = $Mode
            EvidenceLabel    = $EvidenceLabel
            ProbedAtUtc      = $probeResult.ProbedAtUtc
            SupportState     = $probeResult.SupportState
            SupportSummary   = $probeResult.SupportSummary
            Hosts            = $probeResult.Hosts
            RegistrationPlan = Get-DryRunRegistrationPlan -ProbeResult $probeResult
        }

        $json = $planResult | ConvertTo-Json -Depth 8
        Save-ProbeOutput -JsonText $json -TargetPath $OutputPath
        Save-TextOutput -Text (Get-DryRunDecisionSummaryText -ProbeResult $requestedProbeResult -DryRunPlan $planResult.RegistrationPlan) -TargetPath $SummaryTextPath
        $json
        break
    }
    "Register" {
        $probeResult = Get-HostInventory
        $requestedProbeResult = Get-RequestedProbeResult -ProbeResult $probeResult -RequestedArchitecture $Architecture -RequestedHost $RequestedHost
        $registerPlan = Get-RegisterPreviewPlan -ProbeResult $probeResult
        $registerResult = [pscustomobject]@{
            ProbeMode      = $Mode
            EvidenceLabel  = $EvidenceLabel
            ProbedAtUtc    = $probeResult.ProbedAtUtc
            SupportState   = $probeResult.SupportState
            SupportSummary = $probeResult.SupportSummary
            Hosts          = $probeResult.Hosts
            RegisterPlan   = $registerPlan
        }

        if ($ExecutionIntent -eq "Live") {
            if ($AllowSelfElevation -and -not $LiveElevatedRelaunch -and -not (Test-IsAdministrator)) {
                $elevatedRegisterResult = Invoke-SelfElevatedLiveExecution -Operation "Register"
                $elevatedJson = $elevatedRegisterResult | ConvertTo-Json -Depth 8
                Save-ProbeOutput -JsonText $elevatedJson -TargetPath $OutputPath
                $elevatedJson
                break
            }

            try {
                $registerExecution = Invoke-LiveRegisterExecution -ProbeResult $probeResult
                $registerResult | Add-Member -NotePropertyName RegisterExecution -NotePropertyValue $registerExecution

                $json = $registerResult | ConvertTo-Json -Depth 8
                Save-ProbeOutput -JsonText $json -TargetPath $OutputPath
                Save-TextOutput -Text (Get-LiveResultSummaryText -Operation "Register" -Succeeded $true -DetailMessage $registerExecution.OverallDecision) -TargetPath $SummaryTextPath

                $registerResult
            }
            catch {
                $failurePayload = New-LiveFailurePayload -Operation "Register" -ProbeResult $probeResult -ErrorRecord $_
                $json = $failurePayload | ConvertTo-Json -Depth 8
                Save-ProbeOutput -JsonText $json -TargetPath $OutputPath
                Save-TextOutput -Text (Get-LiveResultSummaryText -Operation "Register" -Succeeded $false -DetailMessage $_.Exception.Message) -TargetPath $SummaryTextPath
                throw
            }
            break
        }

        $json = $registerResult | ConvertTo-Json -Depth 8
        Save-ProbeOutput -JsonText $json -TargetPath $OutputPath
        Save-TextOutput -Text (Get-InstallerDecisionSummaryText -ProbeResult $requestedProbeResult -PreviewPlan $registerPlan) -TargetPath $SummaryTextPath
        $json
        break
    }
    "Unregister" {
        $probeResult = Get-HostInventory
        $requestedProbeResult = Get-RequestedProbeResult -ProbeResult $probeResult -RequestedArchitecture $Architecture -RequestedHost $RequestedHost
        $unregisterPlan = Get-UnregisterPreviewPlan -ProbeResult $probeResult
        $unregisterResult = [pscustomobject]@{
            ProbeMode       = $Mode
            EvidenceLabel   = $EvidenceLabel
            ProbedAtUtc     = $probeResult.ProbedAtUtc
            SupportState    = $probeResult.SupportState
            SupportSummary  = $probeResult.SupportSummary
            Hosts           = $probeResult.Hosts
            UnregisterPlan  = $unregisterPlan
        }

        if ($ExecutionIntent -eq "Live") {
            if ($AllowSelfElevation -and -not $LiveElevatedRelaunch -and -not (Test-IsAdministrator)) {
                $elevatedUnregisterResult = Invoke-SelfElevatedLiveExecution -Operation "Unregister"
                $elevatedJson = $elevatedUnregisterResult | ConvertTo-Json -Depth 8
                Save-ProbeOutput -JsonText $elevatedJson -TargetPath $OutputPath
                $elevatedJson
                break
            }

            try {
                $unregisterExecution = Invoke-LiveUnregisterExecution -ProbeResult $probeResult
                $unregisterResult | Add-Member -NotePropertyName UnregisterExecution -NotePropertyValue $unregisterExecution

                $json = $unregisterResult | ConvertTo-Json -Depth 8
                Save-ProbeOutput -JsonText $json -TargetPath $OutputPath
                Save-TextOutput -Text (Get-LiveResultSummaryText -Operation "Unregister" -Succeeded $true -DetailMessage $unregisterExecution.OverallDecision) -TargetPath $SummaryTextPath

                $unregisterResult
            }
            catch {
                $failurePayload = New-LiveFailurePayload -Operation "Unregister" -ProbeResult $probeResult -ErrorRecord $_
                $json = $failurePayload | ConvertTo-Json -Depth 8
                Save-ProbeOutput -JsonText $json -TargetPath $OutputPath
                Save-TextOutput -Text (Get-LiveResultSummaryText -Operation "Unregister" -Succeeded $false -DetailMessage $_.Exception.Message) -TargetPath $SummaryTextPath
                throw
            }
            break
        }

        $json = $unregisterResult | ConvertTo-Json -Depth 8
        Save-ProbeOutput -JsonText $json -TargetPath $OutputPath
        Save-TextOutput -Text (Get-InstallerDecisionSummaryText -ProbeResult $requestedProbeResult -PreviewPlan $unregisterPlan) -TargetPath $SummaryTextPath
        $json
        break
    }
    "WpsAddinsWlExperiment" {
        $registryPath = "HKCU:\Software\Kingsoft\Office\WPS\AddinsWl"
        $dateStamp = (Get-Date).ToString("yyyyMMdd")
        $timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz")

        if (-not $EvidenceDir) { $EvidenceDir = Join-Path $PSScriptRoot "..\docs\installer\evidence" }
        if (-not (Test-Path -LiteralPath $EvidenceDir -PathType Container)) {
            New-Item -Path $EvidenceDir -ItemType Directory -Force | Out-Null
        }

        $backupPath = Join-Path $EvidenceDir "CurrentMachine-WpsX86-AddinsWl-backup-$dateStamp.reg"

        switch ($Action) {
            "backup" {
                reg export "HKCU\Software\Kingsoft\Office\WPS\AddinsWl" `"$backupPath`" *>$null
                if ($LASTEXITCODE -ne 0) {
                    Write-Error "Backup failed: reg export exit code $LASTEXITCODE"
                    return
                }
                $result = [pscustomobject]@{
                    Action = "backup"
                    ExperimentId = $ExperimentId
                    Timestamp = $timestamp
                    BackupSucceeded = $true
                    BackupPath = $backupPath
                }
                $result | ConvertTo-Json -Depth 4
                break
            }
            "write" {
                if (-not $ProgId) { Write-Error "-ProgId is required for write action"; return }
                Set-ItemProperty -Path $registryPath -Name $ProgId -Value $ValuePayload -Type String -Force
                $result = [pscustomobject]@{
                    Action = "write"
                    ExperimentId = $ExperimentId
                    Timestamp = $timestamp
                    WriteSucceeded = $true
                    ProgId = $ProgId
                    ValuePayload = $ValuePayload
                }
                $result | ConvertTo-Json -Depth 4
                break
            }
            "verify" {
                if (-not $ProgId) { Write-Error "-ProgId is required for verify action"; return }
                try {
                    $actual = (Get-ItemProperty -Path $registryPath -ErrorAction Stop).$ProgId
                    $match = ($actual -eq $ValuePayload)
                    $result = [pscustomobject]@{
                        Action = "verify"
                        ExperimentId = $ExperimentId
                        Timestamp = $timestamp
                        VerifySucceeded = $match
                        Expected = $ValuePayload
                        Actual = $actual
                    }
                    $result | ConvertTo-Json -Depth 4
                }
                catch {
                    $result = [pscustomobject]@{
                        Action = "verify"
                        ExperimentId = $ExperimentId
                        Timestamp = $timestamp
                        VerifySucceeded = $false
                        Error = "Cannot read AddinsWl: $_"
                    }
                    $result | ConvertTo-Json -Depth 4
                }
                break
            }
            "restore" {
                if (-not (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
                    Write-Error "Backup file not found: $backupPath"
                    return
                }
                reg import `"$backupPath`" *>$null
                if ($LASTEXITCODE -ne 0) {
                    $result = [pscustomobject]@{
                        Action = "restore"
                        ExperimentId = $ExperimentId
                        Timestamp = $timestamp
                        RestoreSucceeded = $false
                        Error = "reg import failed with exit code $LASTEXITCODE"
                    }
                    $result | ConvertTo-Json -Depth 4
                    return
                }
                $result = [pscustomobject]@{
                    Action = "restore"
                    ExperimentId = $ExperimentId
                    Timestamp = $timestamp
                    RestoreSucceeded = $true
                    BackupPath = $backupPath
                }
                $result | ConvertTo-Json -Depth 4
                break
            }
        }
        break
    }
}
