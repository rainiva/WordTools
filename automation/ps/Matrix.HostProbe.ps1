# Layer 1: detect Word/WPS hosts without launching Office.

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$OutputPath,
    [string]$EvidenceLabel
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "Matrix.Common.ps1")

function Write-MatrixUtf8File {
    param(
        [string]$Path,
        [string]$Content
    )

    $encoding = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($Path, $Content, $encoding)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Write verification failed: file not found after write: $Path"
    }

    $readBack = [System.IO.File]::ReadAllText($Path, $encoding)
    if ($readBack -ne $Content) {
        throw "Write verification failed: content mismatch for $Path"
    }
}

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Get-MatrixRepoRoot
}

$coreResult = Invoke-InstallerCoreIfAvailable -RepoRoot $RepoRoot -Mode Probe -ExecutionIntent PreviewOnly -OutputPath $OutputPath -RequestedHost Both
if ($null -ne $coreResult) {
    if ($coreResult -is [System.Array] -and $coreResult.Count -eq 1) {
        $coreResult = $coreResult[0]
    }

    $hosts = @()
    if ($null -ne $coreResult.Hosts) {
        foreach ($hostEntry in @($coreResult.Hosts)) {
        $bitness = ([string]$hostEntry.HostBitness).Replace("x86", "32").Replace("x64", "64").Replace("x", "")
        $hosts += [ordered]@{
            host = [string]$hostEntry.HostName
            installed = $true
            bitness = $bitness
            version = Get-OfficeVersionLine -ExecutablePath ([string]$hostEntry.ExecutablePath)
            path = [string]$hostEntry.ExecutablePath
            probe_source = [string]$hostEntry.ProbeSource
            support_status = [string]$hostEntry.SupportStatus
        }
    }
    }

    $word = $hosts | Where-Object { $_.host -eq "Word" } | Select-Object -First 1
    $wps = $hosts | Where-Object { $_.host -eq "WPS" } | Select-Object -First 1

    $payload = [ordered]@{
        layer = "host_probe"
        source = "Installer.Core.ps1"
        probed_at_utc = (Get-Date).ToUniversalTime().ToString("o")
        evidence_label = if ([string]::IsNullOrWhiteSpace($EvidenceLabel)) { [string]$coreResult.EvidenceLabel } else { $EvidenceLabel }
        word = if ($null -eq $word) { @{ installed = $false } } else { $word }
        wps = if ($null -eq $wps) { @{ installed = $false } } else { $wps }
        hosts = $hosts
        support_summary = $coreResult.SupportSummary
        pass = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
        $directory = Split-Path -Parent $OutputPath
        if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }
        Write-MatrixUtf8File -Path $OutputPath -Content ($payload | ConvertTo-Json -Depth 10)
    }

    Write-MatrixJsonResult -Payload $payload
    exit 0
}

function Build-FallbackHostRecord {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HostName,
        [string]$ExecutableName
    )

    $appPath = Get-RegistryExePath -AppName $ExecutableName
    if ($null -eq $appPath) {
        return [ordered]@{
            host = $HostName
            installed = $false
            bitness = $null
            version = $null
            path = $null
            probe_source = $null
        }
    }

    $path = [string]$appPath.path
    $bitness = Get-ExecutableBitness -Path $path
    return [ordered]@{
        host = $HostName
        installed = $true
        bitness = $bitness
        version = Get-OfficeVersionLine -ExecutablePath $path
        path = $path
        probe_source = [string]$appPath.source
    }
}

$wordRecord = Build-FallbackHostRecord -HostName "Word" -ExecutableName "WINWORD.EXE"
$wpsCandidates = @(Find-WpsExecutable)
$wpsRecord = if ($wpsCandidates.Count -gt 0) {
    $path = $wpsCandidates[0]
    [ordered]@{
        host = "WPS"
        installed = $true
        bitness = Get-ExecutableBitness -Path $path
        version = Get-OfficeVersionLine -ExecutablePath $path
        path = $path
        probe_source = "filesystem_scan"
    }
}
else {
    [ordered]@{
        host = "WPS"
        installed = $false
        bitness = $null
        version = $null
        path = $null
        probe_source = $null
    }
}

$payload = [ordered]@{
    layer = "host_probe"
    source = "Matrix.HostProbe.ps1"
    probed_at_utc = (Get-Date).ToUniversalTime().ToString("o")
    evidence_label = $EvidenceLabel
    word = $wordRecord
    wps = $wpsRecord
    hosts = @($wordRecord, $wpsRecord)
    pass = $true
}

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $directory = Split-Path -Parent $OutputPath
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    Write-MatrixUtf8File -Path $OutputPath -Content ($payload | ConvertTo-Json -Depth 10)
}

Write-MatrixJsonResult -Payload $payload
