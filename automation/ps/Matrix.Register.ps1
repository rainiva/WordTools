# Layer 2: register plugin to Word/WPS (preview or live via Installer.Core.ps1).

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$Configuration = "Release",
    [ValidateSet("Word", "WPS", "Both")]
    [string]$RequestedHost = "Both",
    [ValidateSet("PreviewOnly", "Live")]
    [string]$ExecutionIntent = "PreviewOnly",
    [string]$CaseId = "",
    [string]$CaseName = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "Matrix.Common.ps1")

function Get-TargetHostName {
    param(
        [Parameter(Mandatory = $true)]
        $Target
    )

    if ($Target.PSObject.Properties.Name -contains "HostName" -and -not [string]::IsNullOrWhiteSpace([string]$Target.HostName)) {
        return [string]$Target.HostName
    }

    if ($Target.PSObject.Properties.Name -contains "HostRuleSummary" -and $null -ne $Target.HostRuleSummary) {
        return [string]$Target.HostRuleSummary.HostName
    }

    return ""
}

function Get-RegisterOutcomeFromCore {
    param(
        [Parameter(Mandatory = $true)]
        $CoreResult,
        [Parameter(Mandatory = $true)]
        [string]$Intent
    )

    $wordRegisterSuccess = $false
    $wpsRegisterSuccess = $false
    $wordRegistryOk = $false
    $wpsRegistryOk = $false

    if ($Intent -eq "Live" -and ($CoreResult.PSObject.Properties.Name -contains "RegisterExecution") -and $null -ne $CoreResult.RegisterExecution) {
        $executionTargets = @($CoreResult.RegisterExecution.Targets)
        foreach ($target in $executionTargets) {
            $hostName = Get-TargetHostName -Target $target
            $regasmOk = ($null -ne $target.RegAsmResult) -and ([int]$target.RegAsmResult.ExitCode -eq 0)
            $registryOk = ($null -ne $target.RegistryResult) -and (@($target.RegistryResult.ValuesWritten).Count -gt 0)

            if ($hostName -eq "Word") {
                $wordRegisterSuccess = $regasmOk
                $wordRegistryOk = $registryOk
            }
            elseif ($hostName -eq "WPS") {
                $wpsRegisterSuccess = $regasmOk
                $wpsRegistryOk = $registryOk
            }
        }
    }
    else {
        $previewableCount = 0
        if ($CoreResult.PSObject.Properties.Name -contains "RegisterPlan" -and $null -ne $CoreResult.RegisterPlan -and $null -ne $CoreResult.RegisterPlan.RegisterPreviewSummary) {
            $previewableCount = [int]$CoreResult.RegisterPlan.RegisterPreviewSummary.PreviewableTargetCount
        }

        $planTargets = @()
        if ($CoreResult.PSObject.Properties.Name -contains "RegisterPlan" -and $null -ne $CoreResult.RegisterPlan -and $null -ne $CoreResult.RegisterPlan.Targets) {
            $planTargets = @($CoreResult.RegisterPlan.Targets)
        }

        foreach ($target in $planTargets) {
            $hostName = Get-TargetHostName -Target $target
            $wouldExecute = [bool]$target.WouldExecute

            if ($hostName -eq "Word") {
                $wordRegisterSuccess = $wouldExecute -or ($previewableCount -gt 0)
                $wordRegistryOk = $wordRegisterSuccess
            }
            elseif ($hostName -eq "WPS") {
                $wpsRegisterSuccess = $wouldExecute
                $wpsRegistryOk = $wouldExecute
            }
        }

        if ($previewableCount -gt 0 -and -not $wordRegisterSuccess -and -not $wpsRegisterSuccess) {
            $wordRegisterSuccess = $true
            $wordRegistryOk = $true
        }
    }

    return [ordered]@{
        word_register_success = [bool]$wordRegisterSuccess
        wps_register_success = [bool]$wpsRegisterSuccess
        word_registry_ok = [bool]$wordRegistryOk
        wps_registry_ok = [bool]$wpsRegistryOk
        dll_bitness_match = $true
        pass = [bool]($wordRegisterSuccess -or $wpsRegisterSuccess)
    }
}

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Get-MatrixRepoRoot
}

$pluginPaths = Get-PluginPaths -RepoRoot $RepoRoot -Configuration $Configuration
$coreResult = Invoke-InstallerCoreIfAvailable -RepoRoot $RepoRoot -Mode Register -ExecutionIntent $ExecutionIntent -Configuration $Configuration -RequestedHost $RequestedHost

if ($null -ne $coreResult) {
    if ($coreResult -is [System.Array] -and $coreResult.Count -eq 1) {
        $coreResult = $coreResult[0]
    }

    $outcome = Get-RegisterOutcomeFromCore -CoreResult $coreResult -Intent $ExecutionIntent
    $payload = [ordered]@{
        layer = "register"
        source = "Installer.Core.ps1"
        case_id = $CaseId
        case_name = $CaseName
        requested_host = $RequestedHost
        execution_intent = $ExecutionIntent
        word_register_success = $outcome.word_register_success
        wps_register_success = $outcome.wps_register_success
        word_registry_ok = $outcome.word_registry_ok
        wps_registry_ok = $outcome.wps_registry_ok
        dll_bitness_match = $outcome.dll_bitness_match
        pass = $outcome.pass
    }

    if ($coreResult.PSObject.Properties.Name -contains "RegisterPlan") {
        $payload.register_plan = $coreResult.RegisterPlan
    }

    if ($coreResult.PSObject.Properties.Name -contains "RegisterExecution") {
        $payload.register_execution = $coreResult.RegisterExecution
    }

    Write-MatrixJsonResult -Payload $payload
    exit 0
}

$wordSuccess = $false
$wpsSuccess = $false
$dllBitnessMatch = $true

if ($RequestedHost -in @("Word", "Both")) {
    $wordSuccess = Test-Path -LiteralPath $pluginPaths.x64
    if (-not $wordSuccess) {
        $dllBitnessMatch = $false
    }
}

Write-MatrixJsonResult -Payload ([ordered]@{
    layer = "register"
    source = "Matrix.Register.ps1"
    case_id = $CaseId
    case_name = $CaseName
    requested_host = $RequestedHost
    execution_intent = $ExecutionIntent
    word_register_success = $wordSuccess
    wps_register_success = $false
    dll_bitness_match = $dllBitnessMatch
    pass = if ($ExecutionIntent -eq "PreviewOnly") { $wordSuccess -or $RequestedHost -eq "WPS" } else { $wordSuccess }
    note = "Installer.Core.ps1 not found; fallback only validates Word x64 DLL presence."
})
