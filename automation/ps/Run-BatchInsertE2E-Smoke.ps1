#Requires -Version 5.1
<#
.SYNOPSIS
  Run tier-1 (smoke) batch-insert E2E with minimal images.

.DESCRIPTION
  UI smoke: AC-UI-B05 (1) + B07 (root only) via COM direct => ~5 images
  Headless smoke: AC-B01, B02, B03, B05 (fixtures, no real folder)

  With --e2e-tier, pytest runs one Word session per leg (batch mode).
  Use --e2e-per-case to force legacy one-Word-per-case runs.

  Prerequisites: Word closed, Release build, registered add-in for UI leg.
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = "",
    [string]$ImageRoot = $env:WORDTOOLS_UI_IMAGE_ROOT,
    [switch]$UiOnly,
    [switch]$HeadlessOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not $RepoRoot) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
}

if (-not $ImageRoot) {
    $ImageRoot = "C:\Users\coxte\Desktop\test2"
}

$automationRoot = Join-Path $RepoRoot "automation"
$venvPython = Join-Path $automationRoot ".venv\Scripts\python.exe"
if (Test-Path $venvPython) {
    $python = $venvPython
}
else {
    $python = "py"
}

function Invoke-SmokePytest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TestPath,

        [Parameter(Mandatory = $true)]
        [string]$MarkerExpr
    )

    $args = @(
        "-m", "pytest",
        $TestPath,
        "--e2e-tier", "smoke",
        "-m", $MarkerExpr,
        "-v",
        "--tb=short"
    )
    & $python @args
    return $LASTEXITCODE
}

if (-not $UiOnly) {
    Write-Host "== Headless smoke (fixtures) ==" -ForegroundColor Cyan
    $code = Invoke-SmokePytest -TestPath (Join-Path $automationRoot "tests\test_batch_insert_e2e.py") -MarkerExpr "integration and smoke"
    if ($code -ne 0) { exit $code }
}

if (-not $HeadlessOnly) {
    Write-Host "== UI smoke (real images: $ImageRoot) ==" -ForegroundColor Cyan
    $env:WORDTOOLS_UI_IMAGE_ROOT = $ImageRoot
    $code = Invoke-SmokePytest -TestPath (Join-Path $automationRoot "tests\test_batch_insert_ui_e2e.py") -MarkerExpr "ui_integration and smoke"
    if ($code -ne 0) { exit $code }
}

Write-Host "Smoke E2E passed." -ForegroundColor Green
