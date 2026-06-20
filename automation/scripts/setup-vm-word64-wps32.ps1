# VM-03 环境搭建脚本：Word 64 + WPS 32
# 用法（管理员 PowerShell）：
#   Set-ExecutionPolicy Bypass -Scope Process
#   cd D:\Project\WordTools\automation\scripts
#   .\setup-vm-word64-wps32.ps1

[CmdletBinding()]
param(
    [string]$RepoRoot = "",
    [switch]$SkipBuild,
    [switch]$SkipPythonDeps
)

$ErrorActionPreference = "Stop"

function Write-Step([string]$Message) {
    Write-Host ("[setup] " + $Message) -ForegroundColor Cyan
}

function Test-Administrator {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Find-RepoRoot([string]$StartPath) {
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
    return $null
}

function Find-MsBuild {
    $candidates = @(
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    )
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }
    $cmd = Get-Command msbuild.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

if (-not (Test-Administrator)) {
    Write-Host "[setup] 请用管理员身份运行 PowerShell。" -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Find-RepoRoot -StartPath $PSScriptRoot
}

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    Write-Host "[setup] 找不到 WordTools.sln。请先把仓库放到 VM，例如 D:\Project\WordTools" -ForegroundColor Red
    exit 1
}

Write-Step "仓库路径: $RepoRoot"

$issues = New-Object System.Collections.Generic.List[string]

# .NET Framework 4.8
$net48Release = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
if (Test-Path -LiteralPath $net48Release) {
    $release = (Get-ItemProperty -LiteralPath $net48Release).Release
    if ($release -lt 528040) {
        $issues.Add(".NET Framework 4.8 未安装或版本过低。请安装 .NET Framework 4.8 运行时/开发者包。") | Out-Null
    }
}
else {
    $issues.Add(".NET Framework 4.8 未检测到。") | Out-Null
}

# Python
$python = Get-Command python.exe -ErrorAction SilentlyContinue
if (-not $python) {
    $issues.Add("Python 未安装。请安装 Python 3.10+ 并勾选 Add to PATH。") | Out-Null
}

# MSBuild
$msbuild = Find-MsBuild
if (-not $msbuild) {
    $issues.Add("MSBuild 未找到。请安装 Visual Studio Build Tools（.NET 桌面开发 + MSBuild）。") | Out-Null
}

# Git（可选）
$git = Get-Command git.exe -ErrorAction SilentlyContinue
if (-not $git) {
    Write-Host "[setup] 提示: 未检测到 Git。若代码通过共享文件夹拷贝，可忽略。" -ForegroundColor Yellow
}

if ($issues.Count -gt 0) {
    Write-Host "[setup] 前置依赖未满足：" -ForegroundColor Red
    foreach ($issue in $issues) {
        Write-Host ("  - " + $issue) -ForegroundColor Red
    }
    exit 2
}

Write-Step "依赖检查通过。"

# Python 自动化依赖
if (-not $SkipPythonDeps) {
    Write-Step "安装 automation Python 依赖..."
    Push-Location (Join-Path $RepoRoot "automation")
    try {
        & python.exe -m pip install -r requirements.txt
        if ($LASTEXITCODE -ne 0) { throw "pip install failed" }
    }
    finally {
        Pop-Location
    }
}

# 编译插件
if (-not $SkipBuild) {
    Write-Step "编译 WordTools Release..."
    & $msbuild (Join-Path $RepoRoot "WordTools.sln") /p:Configuration=Release /restore
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[setup] MSBuild 编译失败。" -ForegroundColor Red
        exit 3
    }

    $dllPath = Join-Path $RepoRoot "WordTools\bin\Release\WordTools.dll"
    if (-not (Test-Path -LiteralPath $dllPath)) {
        Write-Host "[setup] 编译完成但未找到 $dllPath" -ForegroundColor Red
        exit 3
    }
    Write-Step "DLL 已生成: $dllPath"
}

# 宿主探测
Write-Step "运行宿主探测 (VM-03)..."
$probeScript = Join-Path $RepoRoot "automation\ps\Matrix.HostProbe.ps1"
$probeOut = Join-Path $RepoRoot "automation\reports\vm03-setup-probe.json"
New-Item -ItemType Directory -Path (Split-Path -Parent $probeOut) -Force | Out-Null
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $probeScript -RepoRoot $RepoRoot -OutputPath $probeOut
if ($LASTEXITCODE -ne 0) {
    Write-Host "[setup] 宿主探测脚本失败。" -ForegroundColor Red
    exit 4
}

$probe = Get-Content -LiteralPath $probeOut -Raw -Encoding UTF8 | ConvertFrom-Json
$wordOk = ($probe.word.installed -and $probe.word.bitness -eq "64")
$wpsOk = ($probe.wps.installed -and $probe.wps.bitness -eq "32")

Write-Host ""
Write-Host "========== VM-03 搭建结果 ==========" -ForegroundColor Green
Write-Host ("Word: installed={0} bitness={1} path={2}" -f $probe.word.installed, $probe.word.bitness, $probe.word.path)
Write-Host ("WPS:  installed={0} bitness={1} path={2}" -f $probe.wps.installed, $probe.wps.bitness, $probe.wps.path)

if (-not $wordOk -or -not $wpsOk) {
    Write-Host "[setup] 宿主组合不符合 VM-03 (Word64+WPS32)。请检查虚拟机安装。" -ForegroundColor Yellow
    exit 5
}

Write-Host "[setup] VM-03 项目环境就绪。" -ForegroundColor Green
Write-Host ""
Write-Host "下一步：" -ForegroundColor Cyan
Write-Host "  cd $RepoRoot\automation"
Write-Host "  python -m pytest tests -q"
Write-Host "  python run_matrix_test.py --env VM-03"
Write-Host ""
Write-Host "Live 测试（先打快照）："
Write-Host "  python run_matrix_test.py --config configs/word64_wps32_live_register.json --skip-phases unregister,verify_cleanup"
