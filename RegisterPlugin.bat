@echo off
setlocal
chcp 65001 >nul

set "ARCH=%~1"
set "CONFIGURATION=%~2"
set "HOST=%~3"

if "%ARCH%"=="" set "ARCH=Auto"
if "%CONFIGURATION%"=="" set "CONFIGURATION=Debug"
if "%HOST%"=="" set "HOST=Word"

echo ========================================
echo WordTools 插件注册脚本
echo ========================================
echo.

net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [错误] 请以管理员身份运行此脚本！
    echo 右键点击此文件，选择"以管理员身份运行"
    pause
    exit /b 1
)

if /I "%ARCH%"=="Auto" set "ARCH=x64"

if /I not "%ARCH%"=="x64" (
    echo [错误] 当前版本仅支持 64 位 Microsoft Word。
    echo [错误] 暂不支持 32 位 Word、32 位 WPS、64 位 WPS。
    echo [说明] 当前脚本不会为 x86 环境执行注册。
    pause
    exit /b 1
)

if /I not "%HOST%"=="Word" (
    echo [错误] 当前版本仅支持 64 位 Microsoft Word。
    echo [错误] 暂不支持 32 位 Word、32 位 WPS、64 位 WPS。
    echo [说明] 当前脚本不会为 WPS 或混合宿主执行注册。
    pause
    exit /b 1
)

set "REGASM_PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
set "NGEN_PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\ngen.exe"
set "DLL_PATH=%~dp0WordTools\bin\%CONFIGURATION%\WordTools.dll"

if not exist "%REGASM_PATH%" (
    echo [错误] 找不到 regasm.exe: %REGASM_PATH%
    pause
    exit /b 1
)

if not exist "%DLL_PATH%" (
    echo [错误] 找不到 DLL 文件: %DLL_PATH%
    echo 请先编译 WordTools 项目，或调整配置参数。
    pause
    exit /b 1
)

echo 当前版本仅支持 64 位 Microsoft Word。
echo DLL 路径: %DLL_PATH%
echo 使用 regasm: %REGASM_PATH%
echo.

echo [1/3] 正在注册 COM 组件...
"%REGASM_PATH%" /codebase "%DLL_PATH%"
if %errorLevel% neq 0 (
    echo.
    echo [错误] COM 注册失败！
    pause
    exit /b 1
)

echo.
echo [2/3] 正在写入 Word Addins 注册项...
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordTools.ThisAddIn" /v FriendlyName /t REG_SZ /d "Word工具箱" /f
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordTools.ThisAddIn" /v Description /t REG_SZ /d "Word工具箱插件" /f
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordTools.ThisAddIn" /v LoadBehavior /t REG_DWORD /d 3 /f
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordTools.ThisAddIn" /v CommandLineSafe /t REG_DWORD /d 0 /f

if %errorLevel% neq 0 (
    echo.
    echo [错误] 注册表项写入失败！
    pause
    exit /b 1
)

echo.
echo [3/3] 正在执行 NGen 预编译...
if exist "%NGEN_PATH%" (
    "%NGEN_PATH%" install "%DLL_PATH%"
    if %errorlevel% equ 0 (
        echo NGen 预编译完成
    ) else (
        echo 警告: NGen 预编译失败，插件仍可正常工作
    )
) else (
    echo 警告: 找不到 ngen.exe，跳过预编译
)

echo.
echo ========================================
echo 注册成功！
echo ========================================
echo.
echo 下一步操作：
echo 1. 完全关闭 Microsoft Word（包括后台进程）
echo 2. 重新打开 64 位 Microsoft Word
echo 3. 在“文件 -^> 选项 -^> 加载项 -^> COM 加载项”中检查插件
echo.
pause
endlocal
