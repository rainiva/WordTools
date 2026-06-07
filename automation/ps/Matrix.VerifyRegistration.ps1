# Layer 2: verify registration state in registry and plugin files.

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$Configuration = "Release",
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

$pluginPaths = Get-PluginPaths -RepoRoot $RepoRoot -Configuration $Configuration
$wordRegistry = Test-WordAddInRegistryPresent -ProgId $ProgId -ExpectedBitness $ExpectedWordBitness
$wordRegistryOk = [bool]$wordRegistry.present
$wordLoadBehavior = $wordRegistry.load_behavior
$wordRegistryPath = $wordRegistry.registry_path

$wpsAddinsWlRoots = @(
    "HKCU:\Software\Kingsoft\Office\6.0\wps\AddinsWl",
    "HKCU:\Software\Kingsoft\Office\WPS\AddinsWl"
)
$wpsRegistryOk = $false
foreach ($root in $wpsAddinsWlRoots) {
    if (-not (Test-Path -LiteralPath $root)) {
        continue
    }

    $entry = Get-ItemProperty -LiteralPath $root -ErrorAction SilentlyContinue
    if ($null -ne $entry -and ($entry.PSObject.Properties.Name -contains $ProgId)) {
        $wpsRegistryOk = $true
        break
    }
}

$dllExists = Test-Path -LiteralPath $pluginPaths.x64
$loadBehaviorOk = ($null -eq $wordLoadBehavior) -or ([int]$wordLoadBehavior -eq 3)
$pass = $wordRegistryOk -and $dllExists -and $loadBehaviorOk

Write-MatrixJsonResult -Payload ([ordered]@{
    layer = "verify_registration"
    word_registry_ok = [bool]$wordRegistryOk
    wps_registry_ok = [bool]$wpsRegistryOk
    word_registry_path = $wordRegistryPath
    word_load_behavior = $wordLoadBehavior
    load_behavior_ok = [bool]$loadBehaviorOk
    dll_exists = [bool]$dllExists
    dll_path = [string]$pluginPaths.x64
    pass = [bool]$pass
})
