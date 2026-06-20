# Non-interactive COM registration for automation (requires admin).
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
$dllPath = Join-Path $repoRoot "WordTools\bin\$Configuration\WordTools.dll"
$regAsmPath = Join-Path $env:SystemRoot "Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$ngenPath = Join-Path $env:SystemRoot "Microsoft.NET\Framework64\v4.0.30319\ngen.exe"
$progId = "WordTools.ThisAddIn"
$registryPath = "HKLM:\Software\Microsoft\Office\Word\Addins\$progId"

if (-not (Test-Path $dllPath)) {
    throw "Missing DLL: $dllPath"
}

Write-Output "RegAsm: $dllPath"
& $regAsmPath /codebase $dllPath
if ($LASTEXITCODE -ne 0) {
    throw "RegAsm failed with exit code $LASTEXITCODE"
}

if (-not (Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force | Out-Null
}

New-ItemProperty -Path $registryPath -Name "FriendlyName" -Value "Word工具箱" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $registryPath -Name "Description" -Value "Word工具箱插件" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $registryPath -Name "LoadBehavior" -Value 3 -PropertyType DWORD -Force | Out-Null
New-ItemProperty -Path $registryPath -Name "CommandLineSafe" -Value 0 -PropertyType DWORD -Force | Out-Null

if (Test-Path $ngenPath) {
    & $ngenPath install $dllPath | Out-Null
}

Write-Output "Registered $dllPath"
