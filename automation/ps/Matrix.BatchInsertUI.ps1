# Layer: Phase A UI E2E — FlaUI drives InsertPhotosForm + ProgressForm via COM automation entry.

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$CaseId = "AC-UI-B03",
    [string]$CaseIds = "",
    [string]$ImageRoot = "",
    [ValidateSet("true", "false")]
    [string]$Direct = "true",
    [string]$Configuration = "Release"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "Matrix.Common.ps1")

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Get-MatrixRepoRoot -StartPath $PSScriptRoot
}

$repoRootPath = (Resolve-Path -LiteralPath $RepoRoot).Path
$wordToolsDll = Join-Path $repoRootPath ("WordTools\bin\{0}\WordTools.dll" -f $Configuration)
$uiProject = Join-Path $repoRootPath "automation\dotnet\BatchInsertUIE2E\BatchInsertUIE2E.csproj"
$uiExe = Join-Path $repoRootPath ("automation\dotnet\BatchInsertUIE2E\bin\{0}\net48\BatchInsertUIE2E.exe" -f $Configuration)

if (-not (Test-Path -LiteralPath $wordToolsDll)) {
    $msbuild = Find-MsBuild
    if ($null -eq $msbuild) {
        $msbuild = "D:\Apps\Visual Studio\MSBuild\Current\Bin\MSBuild.exe"
    }
    if (-not (Test-Path -LiteralPath $msbuild)) {
        Write-MatrixJsonResult -Payload @{
            case_id = $CaseId
            pass = $false
            error = "WordTools.dll missing and MSBuild not found."
        }
        exit 1
    }
    & $msbuild (Join-Path $repoRootPath "WordTools\WordTools.csproj") /p:Configuration=$Configuration /p:Platform=AnyCPU /v:minimal | Out-Null
}

if (-not (Test-Path -LiteralPath $uiExe)) {
    & dotnet build $uiProject -c $Configuration | Out-Null
}

if ([string]::IsNullOrWhiteSpace($ImageRoot)) {
    $ImageRoot = [string]$env:WORDTOOLS_UI_IMAGE_ROOT
}
if ([string]::IsNullOrWhiteSpace($ImageRoot)) {
    $ImageRoot = "C:\Users\coxte\Desktop\test2"
}
if (-not (Test-Path -LiteralPath $ImageRoot)) {
    Write-MatrixJsonResult -Payload @{
        case_id = $CaseId
        pass = $false
        error = "ImageRoot not found: $ImageRoot"
    }
    exit 1
}

$existingWord = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
if ($existingWord) {
    Write-MatrixJsonResult -Payload @{
        case_id = $CaseId
        pass = $false
        error = "Word is already running; close all Word instances before UI E2E."
    }
    exit 1
}

$exeArgs = @(
    "-RepoRoot", $repoRootPath,
    "-ImageRoot", $ImageRoot,
    "-Direct", $Direct
)
if (-not [string]::IsNullOrWhiteSpace($CaseIds)) {
    $exeArgs += @("-CaseIds", $CaseIds)
}
else {
    $exeArgs += @("-CaseId", $CaseId)
}

try {
    $output = & $uiExe @exeArgs 2>&1
    $text = ($output | Out-String).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        Write-MatrixJsonResult -Payload @{
            case_id = $CaseId
            pass = $false
            error = "BatchInsertUIE2E produced no output"
        }
        exit 1
    }

    $payload = $text | ConvertFrom-Json
    Write-MatrixJsonResult -Payload $payload
    if (-not $payload.pass) {
        exit 1
    }
    exit 0
}
catch {
    Write-MatrixJsonResult -Payload @{
        case_id = $CaseId
        pass = $false
        error = $_.Exception.Message
    }
    exit 1
}
finally {
    Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
