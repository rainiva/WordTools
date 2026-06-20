@echo off
setlocal

set "INTERACTIVE_FLAGS="
set "INTERACTIVE_MODE="

if "%~1"=="" if "%~2"=="" if "%~3"=="" if "%~4"=="" if "%~5"=="" (
  set "INTERACTIVE_MODE=1"
)

set "ARCH=%~1"
set "CONFIGURATION=%~2"
set "HOST=%~3"
set "OPTION4=%~4"
set "OPTION5=%~5"
set "PAUSE_MODE="
set "SHOW_HELP="

if /I "%~1"=="/?" set "SHOW_HELP=1"
if /I "%~2"=="/?" set "SHOW_HELP=1"
if /I "%~3"=="/?" set "SHOW_HELP=1"
if /I "%~4"=="/?" set "SHOW_HELP=1"
if /I "%~5"=="/?" set "SHOW_HELP=1"

if defined SHOW_HELP (
  if /I "%~1"=="/nopause" set "PAUSE_MODE=/nopause"
  if /I "%~2"=="/nopause" set "PAUSE_MODE=/nopause"
  if /I "%~3"=="/nopause" set "PAUSE_MODE=/nopause"
  if /I "%~4"=="/nopause" set "PAUSE_MODE=/nopause"
  if /I "%~5"=="/nopause" set "PAUSE_MODE=/nopause"
)

if "%ARCH%"=="" set "ARCH=Auto"
if "%CONFIGURATION%"=="" set "CONFIGURATION=Debug"
if "%HOST%"=="" set "HOST=Word"

if defined INTERACTIVE_MODE (
  call :print_utf8 "44CQ5o+Q56S644CR5Y2z5bCG5omn6KGM5Y+N5rOo5YaM44CC"
  call :print_utf8 "44CQ5o+Q56S644CR5Y+v55u05o6l5Zue6L2m5L2/55So6buY6K6k5Y+C5pWw77ya"
  echo          Auto Debug Word
  call :print_utf8 "44CQ5o+Q56S644CR6ZmE5Yqg5Y+v6YCJ5Y+C5pWw77yaL25vcGF1c2UgIC8/"
  set /p "INTERACTIVE_FLAGS=> "
)

if /I "%OPTION4%"=="/nopause" set "PAUSE_MODE=/nopause"
if /I "%OPTION5%"=="/nopause" set "PAUSE_MODE=/nopause"

echo(%INTERACTIVE_FLAGS%| findstr /I /C:"/?" >nul
if not errorlevel 1 set "SHOW_HELP=1"
echo(%INTERACTIVE_FLAGS%| findstr /I /C:"/nopause" >nul
if not errorlevel 1 set "PAUSE_MODE=/nopause"

if defined SHOW_HELP goto :show_help

call :print_utf8 "44CQ5o+Q56S644CR5Y+v6YCJ5Y+C5pWw77yaL25vcGF1c2UgPSDnu5PmnZ/ml7bkuI3lho3mjInku7vmhI/plK7nu6fnu60="
call :print_utf8 "44CQ5o+Q56S644CR5pu05aSa5biu5Yqp77yaVW5yZWdpc3RlclBsdWdpbi5iYXQgLz8="
echo.

rem Use an explicit UTF-8 read path because Windows PowerShell mis-parses UTF-8 scripts without BOM via -File.
powershell -NoProfile -ExecutionPolicy Bypass -Command "& { $scriptPath = '%~dp0UnregisterPlugin.ps1'; $script:WordToolsUnregisterScriptPath = $scriptPath; $scriptText = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8); $scriptBlock = [ScriptBlock]::Create($scriptText); & $scriptBlock -Architecture '%ARCH%' -Configuration '%CONFIGURATION%' -RequestedHost '%HOST%' }"
set "EXITCODE=%ERRORLEVEL%"

if "%EXITCODE%"=="200" exit /b 0

call :maybe_pause
exit /b %EXITCODE%

:show_help
call :print_utf8 "55So5rOV77ya"
echo   UnregisterPlugin.bat [Architecture] [Configuration] [Host] [/nopause]
echo.
call :print_utf8 "5Y+C5pWw77ya"
echo   Architecture   Auto^|x86^|x64, default Auto
echo   Configuration  Debug^|Release^|Debug_verify, default Debug
echo   Host           Word^|WPS^|Both, default Word
call :print_utf8 "5q2j5byP5pSv5oyB6L6555WM5LuN5LulIDY0IOS9jSBNaWNyb3NvZnQgV29yZCDkuLrlh4bjgII="
call :print_help_option "/nopause" "57uT5p2f5pe25LiN5YaN5pi+56S64oCc6K+35oyJ5Lu75oSP6ZSu57un57ut4oCd"
echo.
call :print_utf8 "56S65L6L77ya"
echo   UnregisterPlugin.bat
echo   UnregisterPlugin.bat Auto Debug Word /nopause
call :maybe_pause
exit /b 0

:print_utf8
powershell -NoProfile -ExecutionPolicy Bypass -Command "$text = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('%~1')); Write-Host $text"
exit /b 0

:print_help_option
set "HELP_OPTION_NAME=%~1"
set "HELP_OPTION_TEXT=%~2"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$text = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('%HELP_OPTION_TEXT%')); Write-Host ('  ' + '%HELP_OPTION_NAME%' + '          ' + $text)"
set "HELP_OPTION_NAME="
set "HELP_OPTION_TEXT="
exit /b 0

:maybe_pause
if /I not "%PAUSE_MODE%"=="/nopause" (
  echo.
  pause
)
exit /b 0
