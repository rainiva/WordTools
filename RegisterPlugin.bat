@echo off
chcp 65001 >nul
echo ========================================
echo Word 插件注册脚本
echo ========================================
echo.

REM 检查是否以管理员身份运行
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [错误] 请以管理员身份运行此脚本！
    echo 右键点击此文件，选择"以管理员身份运行"
    pause
    exit /b 1
)

REM 设置 regasm 路径（使用 .NET Framework 4.0）
set REGASM_PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe
set DLL_PATH=%~dp0WordTools\bin\Debug\WordTools.dll

if not exist "%REGASM_PATH%" (
    echo [错误] 找不到 regasm.exe: %REGASM_PATH%
    pause
    exit /b 1
)

if not exist "%DLL_PATH%" (
    echo [错误] 找不到 DLL 文件: %DLL_PATH%
    echo 请先在 Visual Studio 中编译项目
    pause
    exit /b 1
)

echo DLL 路径: %DLL_PATH%
echo 使用 regasm: %REGASM_PATH%
echo.

REM 执行 COM 注册
echo [1/2] 正在注册 COM 组件...
"%REGASM_PATH%" /codebase "%DLL_PATH%"
if %errorLevel% neq 0 (
    echo.
    echo [错误] COM 注册失败！
    pause
    exit /b 1
)

echo.
echo [2/2] 正在添加 Word Addins 注册表项...

REM 添加注册表项
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordTools.ThisAddIn" /v FriendlyName /t REG_SZ /d "Word工具箱" /f
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordTools.ThisAddIn" /v Description /t REG_SZ /d "Word工具箱插件" /f
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordTools.ThisAddIn" /v LoadBehavior /t REG_DWORD /d 3 /f
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordTools.ThisAddIn" /v CommandLineSafe /t REG_DWORD /d 0 /f

if %errorLevel% neq 0 (
    echo.
    echo [错误] 注册表项添加失败！
    pause
    exit /b 1
)

echo.
echo ========================================
echo 注册成功！
echo ========================================
echo.
echo 下一步操作：
echo 1. 完全关闭 Microsoft Word（包括后台进程）
echo 2. 重新打开 Microsoft Word
echo 3. 点击"文件" -^> "选项" -^> "加载项"
echo 4. 在底部"管理"下拉菜单选择"COM 加载项"，点击"转到..."
echo 5. 勾选"Word工具箱"
echo 6. 点击"确定"
echo.
echo 启用后，Word 顶部会出现"Word工具箱"选项卡
echo.
pause
