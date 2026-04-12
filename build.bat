@echo off
chcp 65001 >nul
echo ==========================================
echo Word工具箱 - 构建脚本
echo ==========================================
echo.

:: 尝试查找 MSBuild
set "MSBUILD_PATH="

:: 方法1: 使用 vswhere 自动检测（支持任意安装路径）
set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if exist "%VSWHERE%" (
    for /f "delims=" %%i in ('"%VSWHERE%" -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe') do (
        set "MSBUILD_PATH=%%i"
        goto :found
    )
)

:: 方法2: 检查常见位置的 MSBuild
if exist "%ProgramFiles%\Microsoft Visual Studio\2026\Community\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD_PATH=%ProgramFiles%\Microsoft Visual Studio\2026\Community\MSBuild\Current\Bin\MSBuild.exe"
) else if exist "%ProgramFiles%\Microsoft Visual Studio\2026\Professional\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD_PATH=%ProgramFiles%\Microsoft Visual Studio\2026\Professional\MSBuild\Current\Bin\MSBuild.exe"
) else if exist "%ProgramFiles%\Microsoft Visual Studio\2026\Enterprise\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD_PATH=%ProgramFiles%\Microsoft Visual Studio\2026\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
) else if exist "%ProgramFiles%\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD_PATH=%ProgramFiles%\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
) else if exist "%ProgramFiles%\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD_PATH=%ProgramFiles%\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
) else if exist "%ProgramFiles%\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD_PATH=%ProgramFiles%\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
) else if exist "%ProgramFiles%\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD_PATH=%ProgramFiles%\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
)

:: 方法3: 尝试 where 命令
if "%MSBUILD_PATH%"=="" (
    where msbuild >nul 2>nul
    if %errorlevel% equ 0 (
        for /f "delims=" %%i in ('where msbuild') do (
            set "MSBUILD_PATH=%%i"
            goto :found
        )
    )
)

:found
if "%MSBUILD_PATH%"=="" (
    echo 错误: 未找到 MSBuild。请安装 Visual Studio 或 Build Tools。
    echo.
    echo 推荐安装选项:
    echo   1. Visual Studio 2022 Community (免费)
    echo      https://visualstudio.microsoft.com/downloads/
    echo      安装时勾选: .NET 桌面开发 + Office/SharePoint 开发
    echo.
    echo   2. 仅安装 Build Tools for Visual Studio 2022
    echo      https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022
    echo.
    pause
    exit /b 1
)

echo 找到 MSBuild: %MSBUILD_PATH%

:: 生成强名称密钥
echo [1/4] 生成强名称密钥...
if not exist "WordTools\Key.snk" (
    "%MSBUILD_PATH:\MSBuild.exe=\..\..\sn.exe%" -k "WordTools\Key.snk" 2>nul
    if %errorlevel% neq 0 (
        echo 尝试使用 sn.exe 生成密钥...
        where sn >nul 2>nul
        if %errorlevel% equ 0 (
            sn -k "WordTools\Key.snk"
        ) else (
            echo 警告: 无法生成强名称密钥，将使用延迟签名
            echo 请在 Visual Studio 中生成项目以自动创建密钥
        )
    ) else (
        echo 强名称密钥已生成
    )
) else (
    echo 强名称密钥已存在
)

echo.
echo [2/4] 还原 NuGet 包...
nuget restore WordTools.sln 2>nul
if %errorlevel% neq 0 (
    echo 警告: NuGet 还原失败，尝试继续构建...
)

echo.
echo [3/4] 构建解决方案 (Release)...
"%MSBUILD_PATH%" WordTools.sln /p:Configuration=Release /p:Platform="Any CPU" /verbosity:minimal
if %errorlevel% neq 0 (
    echo.
    echo 构建失败！请检查错误信息。
    pause
    exit /b 1
)

echo.
echo [4/4] 复制部署文件...
if not exist "WordTools\bin\Release\publish" mkdir "WordTools\bin\Release\publish"
copy "WordTools\bin\Release\WordTools.dll" "WordTools\bin\Release\publish\" >nul
copy "WordTools\bin\Release\WordTools.pdb" "WordTools\bin\Release\publish\" >nul
copy "WordTools\bin\Release\WordTools.dll.manifest" "WordTools\bin\Release\publish\" >nul
copy "WordTools\WordTools.vsto" "WordTools\bin\Release\publish\" >nul
echo 部署文件已复制到 publish 目录

echo.
echo ==========================================
echo 构建完成！
echo ==========================================
echo.
echo 输出文件位置:
echo   - 插件文件: WordTools\bin\Release\
echo   - 部署文件: WordTools\bin\Release\publish\
echo.
echo 安装方法:
echo   方法1: 直接运行 WordTools.vsto
echo   方法2: 在 Word 中手动添加 COM 加载项
echo.
pause
