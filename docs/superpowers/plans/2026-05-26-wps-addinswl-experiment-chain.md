# WPS AddinsWl 实验链实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在 `Installer.Core.ps1` 中新增 `Invoke-WpsAddinsWlExperiment` 函数，更新支持矩阵新增 `ui-experiment` 状态，并补充离线测试。

**Architecture:** 纯脚本层变更——PowerShell 函数通过 `reg export`/`reg import` 操作注册表备份恢复，通过 `Set-ItemProperty` 写入 AddinsWl。测试通过 `ReadProjectSource` 做源码级断言。无新依赖、无 COM interop、无 UI 交互。

**Tech Stack:** PowerShell 5.1, .NET Framework 4.8 (测试), JSON

**Spec:** [2026-05-26-wps-addinswl-experiment-chain-design.md](../specs/2026-05-26-wps-addinswl-experiment-chain-design.md)

---

## 文件结构

| 文件 | 职责 | 改动 |
|------|------|------|
| `Installer.Core.ps1` | 新增 `Invoke-WpsAddinsWlExperiment` 函数 | 在文件末尾 `switch` 块之前插入新函数 |
| `Installer.SupportMatrix.json` | WPS x86 条目状态变更为 `ui-experiment` | 修改现有 WPS x86 target 的 `status` 和 `validationStage` 字段 |
| `WordTools.Tests/Program.cs` | 3 个新测试方法 | 新增测试方法 + 注册到 `Main()` |

---

### Task 1: `Installer.Core.ps1` — 新增 `Invoke-WpsAddinsWlExperiment` 函数

**Files:**
- Modify: `Installer.Core.ps1` — 在 `Get-WpsReconData` 函数之后（约 L1588）、`Get-ExecutableBitness` 之前插入

- [ ] **Step 1: 在 `Get-WpsReconData` 闭合后插入新函数**

在 `Get-WpsReconData` 函数闭合大括号之后、`Get-ExecutableBitness` 函数定义之前，插入以下完整函数：

```powershell
function Invoke-WpsAddinsWlExperiment {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProgId,

        [Parameter(Mandatory = $true)]
        [string]$ValuePayload,

        [Parameter(Mandatory = $true)]
        [string]$ExperimentId,

        [Parameter(Mandatory = $true)]
        [string]$EvidenceDir,

        [switch]$Restore
    )

    $registryPath = "HKCU:\Software\Kingsoft\Office\WPS\AddinsWl"
    $timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz")
    $dateStamp = (Get-Date).ToString("yyyyMMdd")
    $backupFileName = "CurrentMachine-WpsX86-AddinsWl-backup-$dateStamp.reg"
    $backupPath = Join-Path $EvidenceDir $backupFileName

    if (-not (Test-Path -LiteralPath $EvidenceDir -PathType Container)) {
        New-Item -Path $EvidenceDir -ItemType Directory -Force | Out-Null
    }

    # --- Restore mode ---
    if ($Restore) {
        if (-not (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
            $errorResult = [pscustomobject]@{
                ExperimentId   = $ExperimentId
                Timestamp      = $timestamp
                RestoreSucceeded = $false
                Error          = "Backup file not found at $backupPath"
            }
            return $errorResult | ConvertTo-Json -Depth 4
        }

        reg import $backupPath *>$null
        if ($LASTEXITCODE -ne 0) {
            $errorResult = [pscustomobject]@{
                ExperimentId    = $ExperimentId
                Timestamp       = $timestamp
                RestoreSucceeded = $false
                Error           = "reg import failed with exit code $LASTEXITCODE"
            }
            return $errorResult | ConvertTo-Json -Depth 4
        }

        $postRestoreCount = (Get-ItemProperty -Path $registryPath).PSObject.Properties.Count
        $successResult = [pscustomobject]@{
            ExperimentId      = $ExperimentId
            Timestamp         = $timestamp
            RestoreSucceeded  = $true
            BackupPath        = $backupPath
            PostRestoreEntryCount = $postRestoreCount
        }
        return $successResult | ConvertTo-Json -Depth 4
    }

    # --- Pre-existing check ---
    $preExisting = $false
    $preTotal = 0
    try {
        $existingProps = Get-ItemProperty -Path $registryPath -ErrorAction Stop
        $preTotal = $existingProps.PSObject.Properties.Count
        $preExisting = ($existingProps.PSObject.Properties.Name -contains $ProgId)
    }
    catch {
        $errorResult = [pscustomobject]@{
            ExperimentId    = $ExperimentId
            Timestamp       = $timestamp
            WriteSucceeded  = $false
            Error           = "Cannot read AddinsWl registry key: $_"
        }
        return $errorResult | ConvertTo-Json -Depth 4
    }

    if ($preExisting) {
        $result = [pscustomobject]@{
            ExperimentId        = $ExperimentId
            Timestamp           = $timestamp
            WriteSucceeded      = $false
            PreExisting         = $true
            AddinsWlPreTotal    = $preTotal
            Error               = "ProgId '$ProgId' already exists in AddinsWl"
        }
        return $result | ConvertTo-Json -Depth 4
    }

    # --- Backup ---
    reg export "HKCU\Software\Kingsoft\Office\WPS\AddinsWl" $backupPath *>$null
    if (-not (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
        $errorResult = [pscustomobject]@{
            ExperimentId    = $ExperimentId
            Timestamp       = $timestamp
            WriteSucceeded  = $false
            Error           = "Backup failed: reg export produced no output at $backupPath"
        }
        return $errorResult | ConvertTo-Json -Depth 4
    }

    $backupContent = Get-Content -LiteralPath $backupPath -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($backupContent)) {
        $errorResult = [pscustomobject]@{
            ExperimentId    = $ExperimentId
            Timestamp       = $timestamp
            WriteSucceeded  = $false
            Error           = "Backup failed: backup.reg is empty"
        }
        return $errorResult | ConvertTo-Json -Depth 4
    }

    # --- Write ---
    try {
        Set-ItemProperty -Path $registryPath -Name $ProgId -Value $ValuePayload -ErrorAction Stop
    }
    catch {
        $errorResult = [pscustomobject]@{
            ExperimentId    = $ExperimentId
            Timestamp       = $timestamp
            WriteSucceeded  = $false
            BackupPath      = $backupPath
            Error           = "Set-ItemProperty failed: $_"
        }
        return $errorResult | ConvertTo-Json -Depth 4
    }

    # --- Verify write ---
    $postTotal = 0
    $entryPresent = $false
    try {
        $postProps = Get-ItemProperty -Path $registryPath -ErrorAction Stop
        $postTotal = $postProps.PSObject.Properties.Count
        $entryPresent = ($postProps.PSObject.Properties.Name -contains $ProgId)
    }
    catch {
        $errorResult = [pscustomobject]@{
            ExperimentId    = $ExperimentId
            Timestamp       = $timestamp
            WriteSucceeded  = $false
            BackupPath      = $backupPath
            Error           = "Post-write verification failed: $_"
        }
        return $errorResult | ConvertTo-Json -Depth 4
    }

    $result = [pscustomobject]@{
        ExperimentId        = $ExperimentId
        Timestamp           = $timestamp
        WriteSucceeded      = $entryPresent
        BackupPath          = $backupPath
        WrittenProgId       = $ProgId
        WrittenPayload      = $ValuePayload
        AddinsWlPreTotal    = $preTotal
        AddinsWlPostTotal   = $postTotal
        PreExisting         = $false
    }
    return $result | ConvertTo-Json -Depth 4
}
```

- [ ] **Step 2: 验证函数语法正确**

```powershell
Get-Command Invoke-WpsAddinsWlExperiment -ErrorAction Stop; Write-Host "PASS"
```

### Task 2: `Installer.SupportMatrix.json` — WPS x86 状态更新

**Files:**
- Modify: `Installer.SupportMatrix.json` — WPS x86 target

- [ ] **Step 1: 将 WPS x86 的 status 改为 `ui-experiment`，validationStage 改为 `experimenting`**

当前内容（第 21-27 行）：
```json
    {
      "host": "WPS",
      "bitness": "x86",
      "status": "planned",
      "validationStage": "ui failed",
      "activationRoute": "WpsNativePlugin",
      "note": "Probe and shared-core live registry evidence were collected on the current machine, but WPS UI validation failed on 2026-05-24: the Tools ribbon did not surface a WordTools entry after registration. Keep WPS x86 blocked until the real WPS add-in write contract is understood and verified."
    },
```

替换为：
```json
    {
      "host": "WPS",
      "bitness": "x86",
      "status": "ui-experiment",
      "validationStage": "experimenting",
      "activationRoute": "WpsNativePlugin",
      "note": "Experimenting AddinsWl write contract via Invoke-WpsAddinsWlExperiment incremental chain starting 2026-05-26. Probe and shared-core live registry evidence previously collected, but WPS UI validation failed on 2026-05-24: the Tools ribbon did not surface a WordTools entry. Keep WPS x86 in experiment until the real WPS add-in write contract is understood and verified through the experiment chain."
    },
```

### Task 3: `WordTools.Tests/Program.cs` — 新增 3 个离线测试

**Files:**
- Modify: `WordTools.Tests/Program.cs` — 新增方法 + Main() 注册

- [ ] **Step 1: 在 `TestSupportMatrixDeclaresValidationStagesAndActivationRoutes` 方法之后插入测试方法**

在 `TestSupportMatrixDeclaresValidationStagesAndActivationRoutes` 闭合 `}` 之后、`TestProbeCoreScaffoldExistsWithoutChangingRegistrationPath` 之前插入：

```csharp
    private static void TestInvokeWpsAddinsWlExperimentDefined()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("function Invoke-WpsAddinsWlExperiment"),
            "Installer.Core.ps1 should define Invoke-WpsAddinsWlExperiment function");
        AssertTrue(core.Contains("-ProgId")
            && core.Contains("-ValuePayload")
            && core.Contains("-ExperimentId")
            && core.Contains("-EvidenceDir"),
            "Invoke-WpsAddinsWlExperiment should expose the four mandatory parameters: ProgId, ValuePayload, ExperimentId, EvidenceDir");
        AssertTrue(core.Contains("[switch]$Restore"),
            "Invoke-WpsAddinsWlExperiment should expose a Restore switch parameter");
        AssertTrue(core.Contains("reg export")
            && core.Contains("reg import"),
            "Invoke-WpsAddinsWlExperiment should use reg export/import for backup and restore");
        AssertTrue(core.Contains("Set-ItemProperty"),
            "Invoke-WpsAddinsWlExperiment should use Set-ItemProperty to write into AddinsWl");
        AssertTrue(core.Contains("PreExisting"),
            "Invoke-WpsAddinsWlExperiment should detect pre-existing AddinsWl entries before writing");
    }

    private static void TestInvokeWpsAddinsWlExperimentBackupFlow()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("Test-Path -LiteralPath $backupPath -PathType Leaf"),
            "Invoke-WpsAddinsWlExperiment should verify the backup .reg file exists before proceeding");
        AssertTrue(core.Contains("Backup failed"),
            "Invoke-WpsAddinsWlExperiment should report a backup-failure error when reg export produces no output");
        AssertTrue(core.Contains("AddinsWlPreTotal")
            && core.Contains("AddinsWlPostTotal"),
            "Invoke-WpsAddinsWlExperiment should track entry counts before and after the write");
        AssertTrue(core.Contains("PostRestoreEntryCount"),
            "Invoke-WpsAddinsWlExperiment restore mode should report the entry count after restore");
    }

    private static void TestSupportMatrixHasUiExperimentState()
    {
        string matrixPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.SupportMatrix.json"));
        string matrixText = File.ReadAllText(matrixPath);
        using JsonDocument matrix = JsonDocument.Parse(matrixText);

        AssertTrue(matrixText.Contains("\"ui-experiment\""),
            "support matrix should declare the new ui-experiment status for hosts currently undergoing incremental experiment-chain validation");

        JsonElement targets = matrix.RootElement.GetProperty("targets");
        AssertTrue(targets.EnumerateArray().Any(entry =>
            string.Equals(entry.GetProperty("host").GetString(), "WPS", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("bitness").GetString(), "x86", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("status").GetString(), "ui-experiment", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("validationStage").GetString(), "experimenting", StringComparison.Ordinal)),
            "support matrix should mark WPS x86 as ui-experiment with an experimenting validation stage");
    }
```

- [ ] **Step 2: 在 `Main()` 中注册这 3 个测试**

在 `Main()` 中（`TestSupportMatrixDeclaresValidationStagesAndActivationRoutes` 注册行之后）插入：

```csharp
        Run("Invoke-WpsAddinsWlExperiment is defined with mandatory parameters and Restore switch", TestInvokeWpsAddinsWlExperimentDefined);
        Run("Invoke-WpsAddinsWlExperiment backup flow verifies .reg existence and tracks entry counts", TestInvokeWpsAddinsWlExperimentBackupFlow);
        Run("Support matrix has ui-experiment state for WPS x86", TestSupportMatrixHasUiExperimentState);
```

- [ ] **Step 3: 构建并运行测试**

```powershell
dotnet build WordTools.Tests\WordTools.Tests.csproj
```

```powershell
dotnet exec WordTools.Tests\bin\Debug\WordTools.Tests.dll
```

Expected: 所有测试 PASS，exit code 0。

### Task 4: 验证支持矩阵 JSON 格式

**Files:**
- Verify: `Installer.SupportMatrix.json`

- [ ] **Step 1: 验证 JSON 仍可正确解析**

```powershell
$matrix = Get-Content -Raw -Path "Installer.SupportMatrix.json" | ConvertFrom-Json
$wpsX86 = $matrix.targets | Where-Object { $_.host -eq "WPS" -and $_.bitness -eq "x86" }
Write-Host "WPS x86 status: $($wpsX86.status)"
Write-Host "WPS x86 validationStage: $($wpsX86.validationStage)"
```

Expected: `status = ui-experiment`, `validationStage = experimenting`。
