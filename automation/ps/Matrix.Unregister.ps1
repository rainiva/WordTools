# Layer 2: unregister plugin from Word/WPS.

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$Configuration = "Release",
    [ValidateSet("Word", "WPS", "Both")]
    [string]$RequestedHost = "Both",
    [ValidateSet("PreviewOnly", "Live")]
    [string]$ExecutionIntent = "PreviewOnly"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "Matrix.Common.ps1")

function Get-UnregisterOutcomeFromCore {
    param(
        [Parameter(Mandatory = $true)]
        $CoreResult,
        [Parameter(Mandatory = $true)]
        [string]$Intent
    )

    if ($Intent -ne "Live") {
        return [ordered]@{ pass = $true; note = "preview_only" }
    }

    if ($CoreResult.PSObject.Properties.Name -notcontains "UnregisterExecution" -or $null -eq $CoreResult.UnregisterExecution) {
        return [ordered]@{ pass = $false; error = "missing_unregister_execution" }
    }

    $targets = @($CoreResult.UnregisterExecution.Targets)
    if ($targets.Count -eq 0) {
        return [ordered]@{ pass = $false; error = "no_unregister_targets" }
    }

    $allOk = $true
    foreach ($target in $targets) {
        $regasmOk = ($null -ne $target.RegAsmResult) -and ([int]$target.RegAsmResult.ExitCode -eq 0)
        if (-not $regasmOk) {
            $allOk = $false
        }
    }

    return [ordered]@{
        pass = [bool]$allOk
        applied_target_count = [int]$CoreResult.UnregisterExecution.AppliedTargetCount
    }
}

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Get-MatrixRepoRoot
}

$coreResult = Invoke-InstallerCoreIfAvailable -RepoRoot $RepoRoot -Mode Unregister -ExecutionIntent $ExecutionIntent -Configuration $Configuration -RequestedHost $RequestedHost
if ($null -ne $coreResult) {
    if ($coreResult -is [System.Array] -and $coreResult.Count -eq 1) {
        $coreResult = $coreResult[0]
    }

    $outcome = Get-UnregisterOutcomeFromCore -CoreResult $coreResult -Intent $ExecutionIntent
    Write-MatrixJsonResult -Payload ([ordered]@{
        layer = "unregister"
        source = "Installer.Core.ps1"
        requested_host = $RequestedHost
        execution_intent = $ExecutionIntent
        pass = [bool]$outcome.pass
        applied_target_count = $outcome.applied_target_count
        error = $outcome.error
    })
    exit 0
}

Write-MatrixJsonResult -Payload ([ordered]@{
    layer = "unregister"
    source = "Matrix.Unregister.ps1"
    requested_host = $RequestedHost
    execution_intent = $ExecutionIntent
    pass = ($ExecutionIntent -eq "PreviewOnly")
    note = "Installer.Core.ps1 not found; preview-only fallback."
})
