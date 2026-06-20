# WordTools COM 加载项注册脚本
# 当前版本仅正式支持 64 位 Microsoft Word。

[CmdletBinding()]
param(
    [ValidateSet("Auto", "x86", "x64")]
    [string]$Architecture = "Auto",

    [ValidateSet("Debug", "Release", "Debug_verify")]
    [string]$Configuration = "Debug",

    [ValidateSet("Word", "WPS", "Both")]
    [string]$Host = "Word"
)

$ErrorActionPreference = "Stop"

function Test-Administrator {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Show-UnsupportedMessage([string]$Reason) {
    Write-Host "[错误] 当前版本仅支持 64 位 Microsoft Word。" -ForegroundColor Red
    Write-Host "[错误] 暂不支持 32 位 Word、32 位 WPS、64 位 WPS。" -ForegroundColor Red
    Write-Host "[说明] $Reason" -ForegroundColor Yellow
    Read-Host "按回车退出"
    exit 1
}

function Resolve-Architecture([string]$RequestedArchitecture) {
    if ($RequestedArchitecture -eq "Auto") {
        return "x64"
    }

    return $RequestedArchitecture
}

if (-not (Test-Administrator)) {
    Write-Host "[错误] 请以管理员身份运行此脚本！" -ForegroundColor Red
    Write-Host "右键点击此文件，选择“以管理员身份运行”" -ForegroundColor Yellow
    Read-Host "按回车退出"
    exit 1
}

$resolvedArchitecture = Resolve-Architecture $Architecture

if ($resolvedArchitecture -ne "x64") {
    Show-UnsupportedMessage "当前脚本不会为 x86 环境执行注册。"
}

if ($Host -ne "Word") {
    Show-UnsupportedMessage "当前脚本不会为 WPS 或混合宿主执行注册。"
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$dllPath = Join-Path $scriptDir "WordTools\bin\$Configuration\WordTools.dll"
$regAsmPath = Join-Path $env:SystemRoot "Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$ngenPath = Join-Path $env:SystemRoot "Microsoft.NET\Framework64\v4.0.30319\ngen.exe"
$progId = "WordTools.ThisAddIn"
$registryPath = "HKLM:\Software\Microsoft\Office\Word\Addins\$progId"

if (-not (Test-Path $dllPath)) {
    Write-Host "[错误] 找不到 DLL 文件: $dllPath" -ForegroundColor Red
    Write-Host "请先编译 WordTools 项目，或调整 -Configuration 参数。" -ForegroundColor Yellow
    Read-Host "按回车退出"
    exit 1
}

if (-not (Test-Path $regAsmPath)) {
    Write-Host "[错误] 找不到 regasm.exe: $regAsmPath" -ForegroundColor Red
    Read-Host "按回车退出"
    exit 1
}

Write-Host "========================================" -ForegroundColor Green
Write-Host "WordTools 插件注册脚本 (PowerShell)" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "当前版本仅支持 64 位 Microsoft Word。" -ForegroundColor Cyan
Write-Host "DLL 路径: $dllPath" -ForegroundColor Cyan
Write-Host "使用 regasm: $regAsmPath" -ForegroundColor Cyan
Write-Host ""

Write-Host "正在注册 COM 加载项..." -ForegroundColor Yellow
& $regAsmPath /codebase $dllPath

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "[错误] COM 注册失败！" -ForegroundColor Red
    Read-Host "按回车退出"
    exit 1
}

if (-not (Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force | Out-Null
}

New-ItemProperty -Path $registryPath -Name "FriendlyName" -Value "Word工具箱" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $registryPath -Name "Description" -Value "Word工具箱插件" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $registryPath -Name "LoadBehavior" -Value 3 -PropertyType DWORD -Force | Out-Null
New-ItemProperty -Path $registryPath -Name "CommandLineSafe" -Value 0 -PropertyType DWORD -Force | Out-Null

Write-Host ""
Write-Host "正在执行 NGen 预编译..." -ForegroundColor Yellow
if (Test-Path $ngenPath) {
    & $ngenPath install $dllPath
    if ($LASTEXITCODE -eq 0) {
        Write-Host "NGen 预编译完成" -ForegroundColor Green
    } else {
        Write-Host "警告: NGen 预编译失败，插件仍可正常工作" -ForegroundColor Yellow
    }
} else {
    Write-Host "警告: 找不到 ngen.exe，跳过预编译" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "注册成功！" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "下一步操作：" -ForegroundColor Cyan
Write-Host "1. 完全关闭 Microsoft Word（包括后台进程）" -ForegroundColor White
Write-Host "2. 重新打开 64 位 Microsoft Word" -ForegroundColor White
Write-Host "3. 在“文件 -> 选项 -> 加载项 -> COM 加载项”中检查插件" -ForegroundColor White
Write-Host ""
Read-Host "按回车退出"
