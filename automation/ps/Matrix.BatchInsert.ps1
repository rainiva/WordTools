# Layer: batch insert orchestration E2E via BatchInsertE2E.exe + real Word COM.

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$CaseId = "AC-B03",
    [string]$CaseIds = "",
    [ValidateSet("true", "false")]
    [string]$Visible = "false",
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
$e2eProject = Join-Path $repoRootPath "automation\dotnet\BatchInsertE2E\BatchInsertE2E.csproj"
$e2eExe = Join-Path $repoRootPath ("automation\dotnet\BatchInsertE2E\bin\{0}\net48\BatchInsertE2E.exe" -f $Configuration)

if (-not (Test-Path -LiteralPath $wordToolsDll)) {
    $msbuild = Find-MsBuild
    if ($null -eq $msbuild) {
        $msbuild = "D:\Apps\Visual Studio\MSBuild\Current\Bin\MSBuild.exe"
    }
    if (-not (Test-Path -LiteralPath $msbuild)) {
        Write-MatrixJsonResult -Payload @{
            case_id = $CaseId
            pass = $false
            error = "WordTools.dll missing and MSBuild not found. Build WordTools Release first."
        }
        exit 1
    }

    $buildArgs = @(
        (Join-Path $repoRootPath "WordTools\WordTools.csproj"),
        "/p:Configuration=$Configuration",
        "/p:Platform=AnyCPU",
        "/v:minimal"
    )
    & $msbuild @buildArgs | Out-Null
}

if (-not (Test-Path -LiteralPath $e2eExe)) {
    & dotnet build $e2eProject -c $Configuration | Out-Null
    if (-not (Test-Path -LiteralPath $e2eExe)) {
        $e2eExe = Join-Path $repoRootPath ("automation\dotnet\BatchInsertE2E\bin\{0}\net48\BatchInsertE2E.exe" -f $Configuration)
    }
}

$fixturesScript = Join-Path $repoRootPath "automation\scripts\generate-fixtures.ps1"
$templatePath = Join-Path $repoRootPath "automation\assets\table-template.docx"
if (-not (Test-Path -LiteralPath $templatePath)) {
    & $fixturesScript | Out-Null
}

$existingWord = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
if ($existingWord) {
    Write-MatrixJsonResult -Payload @{
        case_id = $CaseId
        pass = $false
        error = "Word is already running; close all Word instances before batch-insert E2E."
    }
    exit 1
}

try {
    $exeArgs = @("-RepoRoot", $repoRootPath, "-Visible", $Visible)
    if (-not [string]::IsNullOrWhiteSpace($CaseIds)) {
        $exeArgs += @("-CaseIds", $CaseIds)
    }
    else {
        $exeArgs += @("-CaseId", $CaseId)
    }

    $output = & $e2eExe @exeArgs 2>&1
    $text = ($output | Out-String).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        Write-MatrixJsonResult -Payload @{
            case_id = $CaseId
            pass = $false
            error = "BatchInsertE2E produced no output"
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
