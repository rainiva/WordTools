# WordTools COM 加载项反注册脚本
# 当前版本正式支持 64 位 Microsoft Word。

[CmdletBinding()]
param(
    [ValidateSet("Auto", "x86", "x64")]
    [string]$Architecture = "Auto",

    [ValidateSet("Debug", "Release", "Debug_verify")]
    [string]$Configuration = "Debug",

    [Alias("Host")]
    [ValidateSet("Word", "WPS", "Both")]
    [string]$RequestedHost = "Word",

    [switch]$ElevatedRetry
)

$ErrorActionPreference = "Stop"

function Test-Administrator {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Show-UnsupportedMessage([string]$Reason) {
    Write-Host "[错误] 当前版本仅支持 64 位 Microsoft Word。" -ForegroundColor Red
    Write-Host "[错误] 暂不支持 32 位 Word、32 位 WPS、64 位 WPS，以及混合宿主组合。" -ForegroundColor Red
    Write-Host "[说明] $Reason" -ForegroundColor Yellow
    Read-Host "按回车退出"
    exit 1
}

function Resolve-Architecture([string]$RequestedArchitecture) {
    return $RequestedArchitecture
}

function Escape-PowerShellSingleQuotedString([string]$Value) {
    if ($null -eq $Value) {
        return [string]::Empty
    }

    return $Value.Replace("'", "''")
}

function Get-WindowsPowerShellPath {
    return Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
}

function Test-UacCancellationException([System.Exception]$Exception) {
    if ($null -eq $Exception) {
        return $false
    }

    if ($Exception.HResult -eq -2147023673) {
        return $true
    }

    $cancelMessage = (New-Object System.ComponentModel.Win32Exception 1223).Message
    return ([string]$Exception.Message).IndexOf($cancelMessage, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
}

function Convert-CommandLineArgument([string]$Value) {
    if ($null -eq $Value -or $Value.Length -eq 0) {
        return '""'
    }

    if ($Value.IndexOfAny([char[]]@(' ', "`t", '"')) -lt 0) {
        return $Value
    }

    return '"' + $Value.Replace('"', '\"') + '"'
}

function Join-CommandLineArguments([string[]]$ArgumentList) {
    return (($ArgumentList | ForEach-Object { Convert-CommandLineArgument -Value $_ }) -join " ")
}

function Start-ElevatedWindowsPowerShell {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ArgumentList,

        [Parameter(Mandatory = $true)]
        [string]$WorkingDirectory
    )

    $powerShellPath = Get-WindowsPowerShellPath
    if (-not (Test-Path -LiteralPath $powerShellPath)) {
        throw "找不到 Windows PowerShell: $powerShellPath"
    }

    $startProcessException = $null

    try {
        Start-Process -FilePath $powerShellPath `
            -ArgumentList $ArgumentList `
            -WorkingDirectory $WorkingDirectory `
            -Verb RunAs `
            -ErrorAction Stop | Out-Null

        return "StartProcess"
    }
    catch {
        if (Test-UacCancellationException -Exception $_.Exception) {
            throw
        }

        $startProcessException = $_.Exception
    }

    try {
        $shellApplication = New-Object -ComObject Shell.Application
        $argumentText = Join-CommandLineArguments -ArgumentList $ArgumentList
        $shellApplication.ShellExecute($powerShellPath, $argumentText, $WorkingDirectory, "runas", 1)
        return "ShellExecute"
    }
    catch {
        $fallbackMessage = if ($null -ne $startProcessException) {
            "Start-Process: $($startProcessException.Message)`r`nShellExecute: $($_.Exception.Message)"
        }
        else {
            $_.Exception.Message
        }

        throw "无法发起管理员提权。$fallbackMessage"
    }
}

function Write-UnregisterStatus([string]$Message, [ConsoleColor]$Color = [ConsoleColor]::Cyan) {
    Write-Host ("[反注册] " + $Message) -ForegroundColor $Color
}

$scriptPath = $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($scriptPath) -and (Get-Variable -Name WordToolsUnregisterScriptPath -Scope Script -ErrorAction SilentlyContinue)) {
    $scriptPath = $script:WordToolsUnregisterScriptPath
}

$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDir)) {
    if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        throw "无法解析 UnregisterPlugin.ps1 所在目录。"
    }

    $scriptDir = Split-Path -Parent $scriptPath
}

Write-UnregisterStatus "WordTools 反注册脚本已启动。"
Write-UnregisterStatus "正在检查管理员权限与反注册参数..."

if (-not (Test-Administrator)) {
    if (-not $ElevatedRetry) {
        try {
            Write-UnregisterStatus "正在请求管理员权限..."
            $elevatedCommandTemplate = @'
& {{
    $scriptPath = '{0}'
    $script:WordToolsUnregisterScriptPath = $scriptPath
    $scriptText = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)
    $scriptBlock = [ScriptBlock]::Create($scriptText)
    & $scriptBlock -Architecture '{1}' -Configuration '{2}' -RequestedHost '{3}' -ElevatedRetry
}}
'@

            $elevatedCommand = $elevatedCommandTemplate -f `
                (Escape-PowerShellSingleQuotedString $scriptPath), `
                (Escape-PowerShellSingleQuotedString $Architecture), `
                (Escape-PowerShellSingleQuotedString $Configuration), `
                (Escape-PowerShellSingleQuotedString $RequestedHost)

            $encodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($elevatedCommand))
            Start-ElevatedWindowsPowerShell `
                -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", $encodedCommand) `
                -WorkingDirectory $scriptDir | Out-Null

            exit 200
        }
        catch {
            if (Test-UacCancellationException -Exception $_.Exception) {
                Write-Host "[错误] 已取消提权，未执行反注册。" -ForegroundColor Red
            }
            else {
                Write-Host ("[错误] 无法发起管理员提权：" + $_.Exception.Message) -ForegroundColor Red
                Write-Host "请先尝试右键点击此文件，选择“以管理员身份运行”。" -ForegroundColor Yellow
            }
            exit 1223
        }
    }

    Write-Host "[错误] 请以管理员身份运行此脚本！" -ForegroundColor Red
    Write-Host "右键点击此文件，选择“以管理员身份运行”" -ForegroundColor Yellow
    Read-Host "按回车退出"
    exit 1
}

Write-UnregisterStatus "已获得管理员权限，正在准备共享安装核心..."

$resolvedArchitecture = Resolve-Architecture $Architecture

if ($resolvedArchitecture -notin @("Auto", "x86", "x64")) {
    Show-UnsupportedMessage "当前脚本不会为非 64 位 Word 宿主执行反注册。"
}

$corePath = Join-Path $scriptDir "Installer.Core.ps1"

if (-not (Test-Path -LiteralPath $corePath)) {
    Write-Host "[错误] 找不到共享安装核心脚本: $corePath" -ForegroundColor Red
    Read-Host "按回车退出"
    exit 1
}

try {
    Write-UnregisterStatus "正在分析宿主与插件 DLL，这一步可能需要几秒..."
    Write-UnregisterStatus "若当前未检测到宿主，将尝试读取上次注册时写入的安装状态记录。"

    $result = & $corePath `
        -Mode Unregister `
        -ExecutionIntent Live `
        -Architecture $resolvedArchitecture `
        -Configuration $Configuration `
        -RequestedHost $RequestedHost

    $execution = $result.UnregisterExecution
    $dllPath = [string]$execution.DllPath
    $target = $execution.Targets[0]
    $hostName = [string]$target.HostRuleSummary.HostName
    $hostBitness = [string]$target.HostRuleSummary.HostBitness
    $hostLabel = if (-not [string]::IsNullOrWhiteSpace($hostName) -and -not [string]::IsNullOrWhiteSpace($hostBitness)) { "$hostName $hostBitness" } else { "当前目标宿主" }
    $regAsmPath = [string]$target.RegAsmResult.ToolPath
    $ngenResult = $target.NgenResult

    Write-Host "========================================" -ForegroundColor Green
    Write-Host "WordTools 插件反注册脚本 (PowerShell)" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "当前正式支持边界仍以 64 位 Microsoft Word 为准。" -ForegroundColor Cyan
    Write-Host "本次目标宿主: $hostLabel" -ForegroundColor Cyan
    Write-Host "DLL 路径: $dllPath" -ForegroundColor Cyan
    Write-Host "使用 regasm: $regAsmPath" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "正在通过共享安装核心反注册 COM 加载项..." -ForegroundColor Yellow
    Write-Host "已通过共享安装核心完成目标宿主的 COM 反注册与加载项注册项清理。" -ForegroundColor Green
    Write-Host ""
    Write-Host "NGen 处理结果..." -ForegroundColor Yellow

    if ($ngenResult.Attempted -and $ngenResult.Succeeded) {
        Write-Host "NGen 卸载完成" -ForegroundColor Green
    }
    elseif ($ngenResult.Attempted) {
        Write-Host "警告: NGen 卸载失败，COM 反注册仍已完成" -ForegroundColor Yellow
    }
    elseif ($ngenResult.PSObject.Properties.Name -contains "SkippedReason" -and -not [string]::IsNullOrWhiteSpace([string]$ngenResult.SkippedReason)) {
        Write-Host ("说明: " + [string]$ngenResult.SkippedReason) -ForegroundColor Cyan
    }
    else {
        Write-Host "说明: 未执行 NGen 卸载或找不到 ngen.exe" -ForegroundColor Cyan
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "反注册成功！" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "下一步操作：" -ForegroundColor Cyan
    Write-Host "1. 完全关闭 Microsoft Word（包括后台进程）" -ForegroundColor White
    Write-Host "2. 重新打开 Word，在 COM 加载项列表中确认 Word工具箱 已消失" -ForegroundColor White
    Write-Host "3. 若曾通过安装包装载，还需在「设置 → 应用」中卸载 Word工具箱 以删除程序文件" -ForegroundColor White
    Write-Host ""
}
catch {
    Write-Host "[错误] 反注册失败：$($_.Exception.Message)" -ForegroundColor Red
}

Read-Host "按回车退出"
