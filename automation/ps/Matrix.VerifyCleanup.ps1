# Layer 2: verify plugin registry entries were removed.

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$ProgId = "WordTools.ThisAddIn",
    [ValidateSet("32", "64", "")]
    [string]$ExpectedWordBitness = "64"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "Matrix.Common.ps1")

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Get-MatrixRepoRoot
}

$wordClean = $true
foreach ($path in (Get-WordAddInRegistryCandidates -ProgId $ProgId -ExpectedBitness $ExpectedWordBitness)) {
    if (Test-Path -LiteralPath $path) {
        $wordClean = $false
        break
    }
}

$wpsAddinsWlRoots = @(
    "HKCU:\Software\Kingsoft\Office\6.0\wps\AddinsWl",
    "HKCU:\Software\Kingsoft\Office\WPS\AddinsWl"
)
$wpsClean = $true
foreach ($root in $wpsAddinsWlRoots) {
    if (-not (Test-Path -LiteralPath $root)) {
        continue
    }

    $entry = Get-ItemProperty -LiteralPath $root -ErrorAction SilentlyContinue
    if ($null -ne $entry -and ($entry.PSObject.Properties.Name -contains $ProgId)) {
        $wpsClean = $false
        break
    }
}

Write-MatrixJsonResult -Payload ([ordered]@{
    layer = "verify_cleanup"
    word_clean = [bool]$wordClean
    wps_clean = [bool]$wpsClean
    pass = [bool]($wordClean -and $wpsClean)
})
