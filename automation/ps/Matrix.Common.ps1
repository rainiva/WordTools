# Shared helpers for WordTools multi-host matrix automation.

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-MatrixRepoRoot {
    param(
        [string]$StartPath = $PSScriptRoot
    )

    $current = Resolve-Path -LiteralPath $StartPath
    while ($null -ne $current) {
        if (Test-Path -LiteralPath (Join-Path $current.Path "WordTools.sln")) {
            return $current.Path
        }

        $parent = Split-Path -Parent $current.Path
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current.Path) {
            break
        }

        $current = Get-Item -LiteralPath $parent
    }

    throw "Unable to locate WordTools repository root from '$StartPath'."
}

function Write-MatrixJsonResult {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Payload
    )

    $json = $Payload | ConvertTo-Json -Depth 20 -Compress:$false
    Write-Output $json
}

function Get-ExecutableBitness {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    $resolved = (Resolve-Path -LiteralPath $Path).Path
    $stream = [System.IO.File]::OpenRead($resolved)
    try {
        $reader = New-Object System.IO.BinaryReader($stream)
        $stream.Seek(0x3C, [System.IO.SeekOrigin]::Begin) | Out-Null
        $peOffset = $reader.ReadInt32()
        $stream.Seek($peOffset + 4, [System.IO.SeekOrigin]::Begin) | Out-Null
        $machine = $reader.ReadUInt16()
    }
    finally {
        $stream.Dispose()
    }

    switch ($machine) {
        0x014c { return "32" }
        0x8664 { return "64" }
        default { return $null }
    }
}

function Get-RegistryExePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppName
    )

    $views = @(
        @{ Hive = [Microsoft.Win32.RegistryHive]::LocalMachine; View = [Microsoft.Win32.RegistryView]::Registry64; Label = "AppPaths:Registry64" },
        @{ Hive = [Microsoft.Win32.RegistryHive]::LocalMachine; View = [Microsoft.Win32.RegistryView]::Registry32; Label = "AppPaths:Registry32" },
        @{ Hive = [Microsoft.Win32.RegistryHive]::CurrentUser; View = [Microsoft.Win32.RegistryView]::Registry64; Label = "AppPathsUser:Registry64" },
        @{ Hive = [Microsoft.Win32.RegistryHive]::CurrentUser; View = [Microsoft.Win32.RegistryView]::Registry32; Label = "AppPathsUser:Registry32" }
    )

    foreach ($entry in $views) {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($entry.Hive, $entry.View)
        $subKey = $baseKey.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\App Paths\$AppName")
        if ($null -eq $subKey) {
            $baseKey.Dispose()
            continue
        }

        try {
            $path = [string]$subKey.GetValue("")
            if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
                return [ordered]@{
                    path = $path
                    source = $entry.Label
                }
            }
        }
        finally {
            $subKey.Dispose()
            $baseKey.Dispose()
        }
    }

    return $null
}

function Find-WpsExecutable {
    $candidates = New-Object System.Collections.Generic.List[string]

    $appPath = Get-RegistryExePath -AppName "wps.exe"
    if ($null -ne $appPath) {
        $candidates.Add([string]$appPath.path) | Out-Null
    }

    $roots = @(
        ${env:ProgramFiles},
        ${env:ProgramFiles(x86)},
        (Join-Path $env:LOCALAPPDATA "Kingsoft\WPS Office")
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($root in $roots) {
        if (-not (Test-Path -LiteralPath $root)) {
            continue
        }

        Get-ChildItem -LiteralPath $root -Filter "wps.exe" -Recurse -ErrorAction SilentlyContinue |
            ForEach-Object { $candidates.Add($_.FullName) | Out-Null }
    }

    return @($candidates | Select-Object -Unique)
}

function Get-OfficeVersionLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExecutablePath
    )

    if (-not (Test-Path -LiteralPath $ExecutablePath)) {
        return $null
    }

    $versionInfo = (Get-Item -LiteralPath $ExecutablePath).VersionInfo
    if ($null -eq $versionInfo) {
        return $null
    }

    return [string]$versionInfo.ProductVersion
}

function Invoke-InstallerCoreIfAvailable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Probe", "Plan", "Register", "Unregister")]
        [string]$Mode,

        [string]$Architecture = "Auto",
        [string]$Configuration = "Release",
        [string]$RequestedHost = "Both",
        [string]$OutputPath,
        [ValidateSet("PreviewOnly", "Live")]
        [string]$ExecutionIntent = "PreviewOnly"
    )

    $corePath = Join-Path $RepoRoot "Installer.Core.ps1"
    if (-not (Test-Path -LiteralPath $corePath)) {
        return $null
    }

    $arguments = @{
        Mode = $Mode
        Architecture = $Architecture
        Configuration = $Configuration
        RequestedHost = $RequestedHost
        ExecutionIntent = $ExecutionIntent
    }

    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
        $arguments.OutputPath = $OutputPath
    }

    $raw = & $corePath @arguments
    if ($null -eq $raw) {
        return $null
    }

    $text = if ($raw -is [System.Array]) {
        ($raw | ForEach-Object { [string]$_ }) -join "`n"
    }
    else {
        [string]$raw
    }

    $text = $text.Trim()
    if ($text.StartsWith("{")) {
        return $text | ConvertFrom-Json
    }

    return $null
}

function Get-WordAddInRegistryCandidates {
    param(
        [string]$ProgId = "WordTools.ThisAddIn",
        [ValidateSet("32", "64", "")]
        [string]$ExpectedBitness = "64"
    )

    $paths = New-Object System.Collections.Generic.List[string]
    $paths.Add("HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId") | Out-Null

    if ($ExpectedBitness -eq "32") {
        $paths.Add("HKLM:\Software\WOW6432Node\Microsoft\Office\Word\Addins\$ProgId") | Out-Null
    }
    else {
        $paths.Add("HKLM:\Software\Microsoft\Office\Word\Addins\$ProgId") | Out-Null
    }

    return @($paths | Select-Object -Unique)
}

function Test-WordAddInRegistryPresent {
    param(
        [string]$ProgId = "WordTools.ThisAddIn",
        [ValidateSet("32", "64", "")]
        [string]$ExpectedBitness = "64"
    )

    foreach ($path in (Get-WordAddInRegistryCandidates -ProgId $ProgId -ExpectedBitness $ExpectedBitness)) {
        if (Test-Path -LiteralPath $path) {
            return [ordered]@{
                present = $true
                registry_path = $path
                load_behavior = (Get-ItemProperty -LiteralPath $path -ErrorAction SilentlyContinue).LoadBehavior
            }
        }
    }

    return [ordered]@{
        present = $false
        registry_path = $null
        load_behavior = $null
    }
}

function Get-PluginPaths {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,
        [string]$Configuration = "Release"
    )

    return [ordered]@{
        x86 = Join-Path $RepoRoot "WordTools\bin\$Configuration\WordTools.dll"
        x64 = Join-Path $RepoRoot "WordTools\bin\$Configuration\WordTools.dll"
        prog_id = "WordTools.ThisAddIn"
        word_addin_key = "HKCU:\Software\Microsoft\Office\Word\Addins\WordTools.ThisAddIn"
    }
}
