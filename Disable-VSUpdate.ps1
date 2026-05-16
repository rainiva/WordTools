#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Disable Visual Studio updates comprehensively.
.DESCRIPTION
    Targets Visual Studio Installer, the core update mechanism.
    Blocks updates via executable lock, manifest lock, registry,
    scheduled tasks, firewall, and hosts file.
    Supports -Restore to undo changes and -WhatIf to preview.
.PARAMETER Restore
    Restore update functionality
.PARAMETER FixPermissions
    Fix Installer directory permissions only, keep executables blocked
.PARAMETER WhatIf
    Preview only, do not modify
.EXAMPLE
    .\Disable-VSUpdate.ps1
    .\Disable-VSUpdate.ps1 -Restore
    .\Disable-VSUpdate.ps1 -FixPermissions
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$Restore,
    [switch]$FixPermissions
)

$ErrorActionPreference = 'Stop'

#region Paths
$InstallerDir = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer"
if (-not (Test-Path $InstallerDir)) {
    $InstallerDir = "$env:ProgramFiles\Microsoft Visual Studio\Installer"
}

$SetupExe       = Join-Path $InstallerDir 'setup.exe'
$VsInstallerExe = Join-Path $InstallerDir 'vs_installer.exe'
$VsInstallerSvc = Join-Path $InstallerDir 'vs_installerservice.exe'
$ChannelDir     = Join-Path $InstallerDir '_channels'

$VsRootDirs = @(
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022",
    "$env:ProgramFiles\Microsoft Visual Studio\2022"
)

$InstallerTasks = @(
    'Microsoft\VisualStudio\Updates\BackgroundDownload',
    'Microsoft\VisualStudio\Updates\UpdateConfiguration',
    'Microsoft\VisualStudio\Installer\BootstrapperUpdate'
)

$ChannelUrls = @(
    'aka.ms',
    'download.visualstudio.microsoft.com',
    'visualstudio.microsoft.com'
)
#endregion

#region Helper Functions
function Write-Step($msg) { Write-Host "  $msg" -ForegroundColor Green }
function Write-Skip($msg) { Write-Host "  $msg" -ForegroundColor DarkGray }
function Write-Warn($msg) { Write-Warning "  $msg" }

function Stop-VSInstallerProcesses {
    $procs = Get-Process -ErrorAction SilentlyContinue | Where-Object {
        $_.ProcessName -in @('setup','vs_installer','vs_installerservice','vs_installershell','VSIXInstaller')
    }
    foreach ($p in $procs) {
        if ($PSCmdlet.ShouldProcess("$($p.ProcessName) (PID:$($p.Id))", 'Terminate')) {
            Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
            Write-Step "Terminated: $($p.ProcessName)"
        }
    }
}

function Set-ChannelManifestLock {
    param([bool]$Lock)

    if (-not (Test-Path $ChannelDir)) {
        Write-Skip "Channel dir not found: $ChannelDir"
        return
    }

    if ($Lock) {
        if ($PSCmdlet.ShouldProcess($ChannelDir, 'Lock channel manifest dir')) {
            Remove-Item "$ChannelDir\*.json" -Force -ErrorAction SilentlyContinue
            Remove-Item "$ChannelDir\*.cat"  -Force -ErrorAction SilentlyContinue
            attrib +R "$ChannelDir\*.*" /S /D 2>$null
            icacls $ChannelDir /inheritance:r /T /C /Q 2>$null | Out-Null
            icacls $ChannelDir /grant "Administrators:(OI)(CI)RX" /T /C /Q 2>$null | Out-Null
            icacls $ChannelDir /grant "SYSTEM:(OI)(CI)RX" /T /C /Q 2>$null | Out-Null
            icacls $ChannelDir /grant "Users:(OI)(CI)RX" /T /C /Q 2>$null | Out-Null
            Write-Step "Locked channel manifest dir"
        }
    }
    else {
        if ($PSCmdlet.ShouldProcess($ChannelDir, 'Unlock channel manifest dir')) {
            attrib -R "$ChannelDir\*.*" /S /D 2>$null
            icacls $ChannelDir /reset /T /C /Q 2>$null | Out-Null
            Write-Step "Unlocked channel manifest dir"
        }
    }
}

function Set-InstallerExecutableLock {
    param([bool]$Lock)

    if (-not (Test-Path $InstallerDir)) {
        Write-Skip "Installer dir not found"
        return
    }

    $executables = @($SetupExe, $VsInstallerExe, $VsInstallerSvc) | Where-Object { Test-Path $_ }

    if ($Lock) {
        Stop-VSInstallerProcesses
        foreach ($exe in $executables) {
            $blocked = "$exe.blocked"
            if (-not (Test-Path $blocked)) {
                if ($PSCmdlet.ShouldProcess($exe, 'Rename and create placeholder')) {
                    Rename-Item -Path $exe -NewName $blocked -Force
                    '' | Set-Content -Path $exe -Force -Encoding ASCII
                    Write-Step "Blocked: $(Split-Path $exe -Leaf)"
                }
            }
            else {
                Write-Skip "Already blocked: $(Split-Path $exe -Leaf)"
            }
        }
    }
    else {
        foreach ($exe in $executables) {
            $blocked = "$exe.blocked"
            if (Test-Path $blocked) {
                if ($PSCmdlet.ShouldProcess($exe, 'Restore executable')) {
                    Remove-Item $exe -Force -ErrorAction SilentlyContinue
                    Rename-Item -Path $blocked -NewName $exe -Force
                    Write-Step "Restored: $(Split-Path $exe -Leaf)"
                }
            }
        }
    }
}

function Set-InstallerDirectoryDeny {
    param([bool]$Deny)

    if (-not (Test-Path $InstallerDir)) {
        Write-Skip "Installer dir not found"
        return
    }

    if ($Deny) {
        if ($PSCmdlet.ShouldProcess($InstallerDir, 'Apply deny-write')) {
            try {
                icacls $InstallerDir /inheritance:r /T /C /Q 2>$null | Out-Null
                icacls $InstallerDir /grant "Administrators:(OI)(CI)RX" /T /C /Q 2>$null | Out-Null
                icacls $InstallerDir /grant "SYSTEM:(OI)(CI)RX" /T /C /Q 2>$null | Out-Null
                icacls $InstallerDir /grant "Users:(OI)(CI)RX" /T /C /Q 2>$null | Out-Null
                Write-Step "Applied deny-write to Installer dir"
            }
            catch {
                Write-Warn "Could not modify Installer dir ACLs (files may be in use). Continuing..."
            }
        }
    }
    else {
        if ($PSCmdlet.ShouldProcess($InstallerDir, 'Restore Installer dir permissions')) {
            try {
                icacls $InstallerDir /reset /T /C /Q 2>$null | Out-Null
                Write-Step "Restored Installer dir permissions"
            }
            catch {
                Write-Warn "Could not reset Installer dir ACLs (files may be in use). Continuing..."
            }
        }
    }
}

function Set-InstallerRegistry {
    param([bool]$Disable)

    if ($Disable) {
        $keysToSet = @{
            'HKLM:\SOFTWARE\Policies\Microsoft\VisualStudio\Setup' = @{
                'DisableAutoUpdate'      = 1
                'DisableInstallUpdates'  = 1
                'SharedInstallationPath' = ''
            }
            'HKLM:\SOFTWARE\Microsoft\VisualStudio\Setup' = @{
                'UpdateConfiguration' = 'None'
            }
            'HKLM:\SOFTWARE\Policies\Microsoft\VisualStudio\Feedback' = @{
                'DisableFeedbackDialog' = 1
            }
        }

        foreach ($path in $keysToSet.Keys) {
            if (-not (Test-Path $path)) {
                New-Item -Path $path -Force | Out-Null
            }
            $values = $keysToSet[$path]
            foreach ($name in $values.Keys) {
                $value = $values[$name]
                if ($PSCmdlet.ShouldProcess("$path\$name", "Set to $value")) {
                    Set-ItemProperty -Path $path -Name $name -Value $value -Force
                    Write-Step "Registry: $name = $value"
                }
            }
        }
    }
    else {
        $pathsToRemove = @(
            'HKLM:\SOFTWARE\Policies\Microsoft\VisualStudio\Setup'
            'HKLM:\SOFTWARE\Policies\Microsoft\VisualStudio\Feedback'
        )
        foreach ($path in $pathsToRemove) {
            if (Test-Path $path) {
                if ($PSCmdlet.ShouldProcess($path, 'Remove policy key')) {
                    Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Step "Removed: $path"
                }
            }
        }

        $setupPath = 'HKLM:\SOFTWARE\Microsoft\VisualStudio\Setup'
        if (Test-Path $setupPath) {
            $prop = Get-ItemProperty -Path $setupPath -Name 'UpdateConfiguration' -ErrorAction SilentlyContinue
            if ($prop) {
                if ($PSCmdlet.ShouldProcess("$setupPath\UpdateConfiguration", 'Remove property')) {
                    Remove-ItemProperty -Path $setupPath -Name 'UpdateConfiguration' -Force -ErrorAction SilentlyContinue
                    Write-Step "Removed: $setupPath\UpdateConfiguration"
                }
            }
        }
    }
}

function Set-InstallerTasks {
    param([bool]$Disable)

    foreach ($taskPath in $InstallerTasks) {
        $taskName = Split-Path $taskPath -Leaf
        $taskParent = Split-Path $taskPath -Parent
        $fullPath = "\$taskPath"

        $tasks = Get-ScheduledTask -TaskPath "$fullPath\*" -ErrorAction SilentlyContinue
        if (-not $tasks) {
            $tasks = Get-ScheduledTask -TaskName $taskName -TaskPath "\$taskParent\" -ErrorAction SilentlyContinue
        }
        foreach ($task in $tasks) {
            if ($Disable) {
                if ($PSCmdlet.ShouldProcess($task.TaskName, 'Disable scheduled task')) {
                    Disable-ScheduledTask -InputObject $task -ErrorAction SilentlyContinue | Out-Null
                    Write-Step "Disabled task: $($task.TaskName)"
                }
            }
            else {
                if ($PSCmdlet.ShouldProcess($task.TaskName, 'Enable scheduled task')) {
                    Enable-ScheduledTask -InputObject $task -ErrorAction SilentlyContinue | Out-Null
                    Write-Step "Enabled task: $($task.TaskName)"
                }
            }
        }
    }
}

function Set-NetworkBlock {
    param([bool]$Block)

    if ($Block) {
        foreach ($url in $ChannelUrls) {
            $ruleName = "Block VS Installer - $url"
            $existing = Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue
            if ($existing) { Write-Skip "Firewall rule exists: $url"; continue }

            try {
                $ips = [System.Net.Dns]::GetHostAddresses($url) |
                    Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
                    Select-Object -ExpandProperty IPAddressToString -Unique
                foreach ($ip in $ips) {
                    if ($PSCmdlet.ShouldProcess("$url ($ip)", 'Add firewall block rule')) {
                        New-NetFirewallRule `
                            -DisplayName $ruleName `
                            -Direction Outbound `
                            -RemoteAddress $ip `
                            -Action Block `
                            -Profile Any `
                            -ErrorAction SilentlyContinue | Out-Null
                    }
                }
                Write-Step "Firewall blocked: $url"
            }
            catch {
                Write-Warn "Cannot resolve $url"
            }
        }
    }
    else {
        $rules = Get-NetFirewallRule -DisplayName 'Block VS Installer*' -ErrorAction SilentlyContinue
        foreach ($rule in $rules) {
            if ($PSCmdlet.ShouldProcess($rule.DisplayName, 'Remove firewall rule')) {
                Remove-NetFirewallRule -Name $rule.Name -ErrorAction SilentlyContinue
                Write-Step "Removed rule: $($rule.DisplayName)"
            }
        }
    }
}

function Set-HostsBlock {
    param([bool]$Block)

    $hostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
    $content = Get-Content $hostsPath -Raw -ErrorAction SilentlyContinue

    if ($Block) {
        foreach ($url in $ChannelUrls) {
            $entry = "127.0.0.1 $url"
            if ($content -notmatch [regex]::Escape($entry)) {
                if ($PSCmdlet.ShouldProcess("hosts -> $url", 'Add block')) {
                    Add-Content -Path $hostsPath -Value $entry -Force -Encoding UTF8
                    Write-Step "Hosts blocked: $url"
                }
            }
            else {
                Write-Skip "Hosts already has: $url"
            }
        }
    }
    else {
        $newContent = $content
        foreach ($url in $ChannelUrls) {
            $escapedUrl = [regex]::Escape($url)
            $newContent = $newContent -replace "(?m)^\s*127\.0\.0\.1\s+$escapedUrl\s*\r?\n", ''
        }
        if ($content -ne $newContent -and $PSCmdlet.ShouldProcess('hosts', 'Remove VS Installer domain blocks')) {
            Set-Content -Path $hostsPath -Value $newContent -Force -Encoding UTF8
            Write-Step "Cleaned VS Installer domains from hosts"
        }
    }
}
#endregion

#region Main
if ($Restore) {
    Write-Host "`n=== Restore Visual Studio Updates ===" -ForegroundColor Cyan

    Write-Host "`n[1/6] Restore Installer dir permissions (first, to avoid lockout)" -ForegroundColor Yellow
    Set-InstallerDirectoryDeny -Deny $false

    Write-Host "`n[2/6] Restore Installer executables" -ForegroundColor Yellow
    Set-InstallerExecutableLock -Lock $false

    Write-Host "`n[3/6] Unlock channel manifest dir" -ForegroundColor Yellow
    Set-ChannelManifestLock -Lock $false

    Write-Host "`n[4/6] Restore registry policies" -ForegroundColor Yellow
    Set-InstallerRegistry -Disable $false

    Write-Host "`n[5/6] Enable scheduled tasks" -ForegroundColor Yellow
    Set-InstallerTasks -Disable $false

    Write-Host "`n[6/6] Remove firewall rules and clean hosts" -ForegroundColor Yellow
    Set-NetworkBlock -Block $false
    Set-HostsBlock -Block $false

    Write-Host "`nVisual Studio updates restored." -ForegroundColor Green
    Write-Host "Run VS Installer manually to verify update channel." -ForegroundColor Yellow
}
elseif ($FixPermissions) {
    Write-Host "`n=== Fix Permissions Only (Keep Update Block) ===" -ForegroundColor Cyan
    Write-Host "Restores Installer dir access while keeping executables blocked." -ForegroundColor Yellow

    Write-Host "`n[1/2] Restore Installer dir permissions" -ForegroundColor Yellow
    Set-InstallerDirectoryDeny -Deny $false

    Write-Host "`n[2/2] Unlock channel manifest dir" -ForegroundColor Yellow
    Set-ChannelManifestLock -Lock $false

    Write-Host "`nPermissions fixed. Installer shortcuts should work now." -ForegroundColor Green
    Write-Host "Executables remain blocked; updates are still disabled." -ForegroundColor Yellow
    Write-Host "Full restore: .\Disable-VSUpdate.ps1 -Restore" -ForegroundColor Yellow
}
else {
    Write-Host "`n=== Disable Visual Studio Updates (Installer-centric) ===" -ForegroundColor Cyan

    Write-Host "`n[1/6] Terminate Installer processes" -ForegroundColor Yellow
    Stop-VSInstallerProcesses

    Write-Host "`n[2/6] Rename Installer executables (core)" -ForegroundColor Yellow
    Set-InstallerExecutableLock -Lock $true

    Write-Host "`n[3/6] Lock channel manifest dir" -ForegroundColor Yellow
    Set-ChannelManifestLock -Lock $true

    Write-Host "`n[4/6] Deny-write on Installer dir" -ForegroundColor Yellow
    Set-InstallerDirectoryDeny -Deny $true

    Write-Host "`n[5/6] Write registry disable policies" -ForegroundColor Yellow
    Set-InstallerRegistry -Disable $true

    Write-Host "`n[6/6] Disable scheduled tasks and block network" -ForegroundColor Yellow
    Set-InstallerTasks -Disable $true
    Set-NetworkBlock -Block $true
    Set-HostsBlock -Block $true

    Write-Host "`nVisual Studio updates disabled." -ForegroundColor Green
    Write-Host "Core measure: Installer executables renamed and replaced with placeholders." -ForegroundColor Yellow
    Write-Host "Restore: .\Disable-VSUpdate.ps1 -Restore" -ForegroundColor Yellow
    Write-Host "Fix permissions only: .\Disable-VSUpdate.ps1 -FixPermissions" -ForegroundColor Yellow
}
#endregion
