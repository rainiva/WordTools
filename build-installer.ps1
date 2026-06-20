param(
    [string]$Configuration = "Release",
    [ValidateSet('None', 'Patch', 'Minor', 'Major')]
    [string]$Bump = 'None'
)

$ErrorActionPreference = "Stop"

$repoRoot = $PSScriptRoot
$versionMeta = & (Join-Path $repoRoot "sync-version.ps1") -Bump $Bump
$appVersion = [string]$versionMeta.Version

function Find-IsccPath {
    $candidates = @(
        "D:\Apps\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles}\Inno Setup 6\ISCC.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    $command = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    throw "ISCC.exe was not found. Please install Inno Setup 6 first."
}

function Build-Installer([string]$IsccPath, [string]$ArchitectureSwitch) {
    $arguments = @(
        "/D$ArchitectureSwitch",
        "/DSourceConfiguration=$Configuration",
        "/DMyAppVersion=$appVersion",
        "Setup.iss"
    )

    & $IsccPath @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to build installer for $ArchitectureSwitch."
    }
}

$isccPath = Find-IsccPath
Write-Host "Using ISCC: $isccPath" -ForegroundColor Cyan

Build-Installer -IsccPath $isccPath -ArchitectureSwitch "ARCH_X86"
Build-Installer -IsccPath $isccPath -ArchitectureSwitch "ARCH_X64"

Write-Host ""
Write-Host "Installer build completed:" -ForegroundColor Green
Write-Host "  dist\\WordToolbox_Setup_${appVersion}_x86.exe"
Write-Host "  dist\\WordToolbox_Setup_${appVersion}_x64.exe"
