# Word COM 加载项注册脚本
# 必须以管理员身份运行

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Green
Write-Host "Word 插件注册脚本 (PowerShell)" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# 检查管理员权限
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "[错误] 请以管理员身份运行此脚本！" -ForegroundColor Red
    Write-Host "右键点击此文件，选择'以管理员身份运行'" -ForegroundColor Yellow
    Read-Host "按任意键退出"
    exit 1
}

# 设置路径
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$DLL_PATH = Join-Path $ScriptDir "WordTools\bin\Debug\WordTools.dll"
$REGASM_PATH = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

if (-not (Test-Path $DLL_PATH)) {
    Write-Host "[错误] 找不到 DLL 文件: $DLL_PATH" -ForegroundColor Red
    Write-Host "请先在 Visual Studio 中编译项目" -ForegroundColor Yellow
    Read-Host "按任意键退出"
    exit 1
}

if (-not (Test-Path $REGASM_PATH)) {
    Write-Host "[错误] 找不到 regasm.exe: $REGASM_PATH" -ForegroundColor Red
    Read-Host "按任意键退出"
    exit 1
}

Write-Host "DLL 路径: $DLL_PATH" -ForegroundColor Cyan
Write-Host "使用 regasm: $REGASM_PATH" -ForegroundColor Cyan
Write-Host ""

# 执行注册
Write-Host "正在注册 COM 加载项..." -ForegroundColor Yellow
& $REGASM_PATH /codebase $DLL_PATH

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "[错误] 注册失败！" -ForegroundColor Red
    Read-Host "按任意键退出"
    exit 1
}

Write-Host ""
Write-Host "正在添加注册表项..." -ForegroundColor Yellow

# 添加 Word Addins 注册表项
$ProgId = "WordTools.ThisAddIn"
$RegistryPath = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"

# 创建注册表项
if (-not (Test-Path $RegistryPath)) {
    New-Item -Path $RegistryPath -Force | Out-Null
}

# 设置注册表值
New-ItemProperty -Path $RegistryPath -Name "FriendlyName" -Value "Word工具箱" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $RegistryPath -Name "Description" -Value "Word工具箱插件" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $RegistryPath -Name "LoadBehavior" -Value 3 -PropertyType DWORD -Force | Out-Null
New-ItemProperty -Path $RegistryPath -Name "CommandLineSafe" -Value 0 -PropertyType DWORD -Force | Out-Null

Write-Host ""
Write-Host "正在执行 NGen 预编译..." -ForegroundColor Yellow
$NGEN_PATH = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\ngen.exe"
if (Test-Path $NGEN_PATH) {
    & $NGEN_PATH install $DLL_PATH
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
Write-Host "2. 重新打开 Microsoft Word" -ForegroundColor White
Write-Host "3. 点击'文件' -> '选项' -> '加载项'" -ForegroundColor White
Write-Host "4. 在底部'管理'下拉菜单选择'COM 加载项'，点击'转到...'" -ForegroundColor White
Write-Host "5. 勾选'Word工具箱'" -ForegroundColor White
Write-Host "6. 点击'确定'" -ForegroundColor White
Write-Host ""
Write-Host "启用后，Word 顶部会出现'Word工具箱'选项卡" -ForegroundColor Yellow
Write-Host ""
Read-Host "按任意键退出"
