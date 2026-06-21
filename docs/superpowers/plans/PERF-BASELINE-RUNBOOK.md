# 插图性能基准采集与对比（PERF Gate）

本文档说明如何锁定 `perf-baseline-2026-06-21.csv`，以及在重构拆分后如何用 `scripts/compare-benchmark.ps1` 验证**插图性能无回归**。

## 前置条件

1. **Release 构建**（与日常性能测试一致）：

```powershell
cd D:\Project\WordTools
msbuild WordTools.sln /p:Configuration=Release
```

2. **注册加载项**（若尚未注册）：按 `RegisterPlugin.ps1` / 安装包流程注册 Release 版 WordTools。

3. **生成 E2E 图片夹具**（PERF-01/02 复用 automation 资源）：

```powershell
.\automation\scripts\generate-fixtures.ps1
```

4. **Word 处于关闭状态**，然后开始下面步骤（避免 COM 冲突）。

5. **Ribbon 日志开关**（性能 CSV 写入门禁）：
   - 开启 **详细日志**
   - 开启 **性能基准 CSV**

   > 若未开启，`InsertionResultPresenter.TryWriteBenchmarkLog` 不会写入 CSV。

---

## 基准场景定义

| 场景 ID | 等价 E2E | run_mode | 文件数 | 操作摘要 |
|---------|----------|----------|--------|----------|
| **PERF-01** | AC-B03 | `SelectedFiles` | 4 | 选中 4 张图 + 描述行 + 自动编号 |
| **PERF-02** | AC-B04 | `Folder` | 5 | 文件夹（根 3 + 子目录 2）+ 子目录标题行 + 编号 |
| **PERF-03** | （大批量） | `SelectedFiles` | 50 | Phase 4 门禁；可选，需先生成 50 张夹具 |

### 夹具路径（Repo 内）

| 用途 | 路径 |
|------|------|
| 表格模板 | `automation/assets/table-template.docx` |
| PERF-01 四图 | `automation/assets/images/selected-4/*.jpg` |
| PERF-02 文件夹 | `automation/assets/images/folder-root/`（含 `sub-a/`） |

建议将文档**另存到固定本地路径**（如 `%USERPROFILE%\Documents\WordToolsPerf\perf-template.docx`），多次运行时使用同一文件，减少路径差异。

---

## 步骤 1：手动跑 PERF-01 / PERF-02（各 3 次）

对每个场景重复 **3 次**（共 6 次插图），每次使用**全新表格区域**或**新文档**，避免旧编号/旧图片干扰。

### PERF-01（选中 4 文件）

1. 打开 `automation/assets/table-template.docx`（或你的固定 perf 文档）
2. 光标放在表格 **第 1 列** 某空行
3. Ribbon → 批量插图
4. 模式：**选择文件**
5. 选中 `automation/assets/images/selected-4/` 下 4 张 jpg
6. 选项：**描述行 = 开**，**自动编号 = 开**（与 AC-B03 一致）
7. 确认插入，等待完成
8. 重复 3 次（可新建文档或移到表格新区域）

### PERF-02（文件夹 5 图）

1. 光标放在表格第 1 列
2. Ribbon → 批量插图
3. 模式：**文件夹**
4. 选择 `automation/assets/images/folder-root/`
5. **包含根目录图片 = 开**，**包含子文件夹图片 = 开**
6. 描述行 + 自动编号 = 开
7. 确认插入，等待完成
8. 重复 3 次

---

## 步骤 2：定位 wordtools-benchmark.csv

默认路径规则（见 `BenchmarkLogService.GetDefaultLogPath`）：

- 优先：当前 Word 文档所在目录下的 `wordtools-benchmark.csv`
- 回退：`%USERPROFILE%\Documents\WordTools\wordtools-benchmark.csv`

插入完成后，找到**最后追加**的 Completed 行（`run_mode=SelectedFiles` / `Folder`）。

---

## 步骤 3：生成 median 基线文件

将上一步找到的 CSV 作为 `-SourceCsv`，写入 repo 基线：

```powershell
cd D:\Project\WordTools

$source = "$env:USERPROFILE\Documents\WordTools\wordtools-benchmark.csv"
# 若 CSV 在文档目录，改为实际路径

.\scripts\capture-benchmark-baseline.ps1 `
  -SourceCsv $source `
  -OutputCsv "docs\superpowers\plans\perf-baseline-2026-06-21.csv" `
  -ScenarioId All `
  -RunsPerScenario 3 `
  -Force
```

成功输出应包含两行 median：`PERF-01`、`PERF-02`。

验证文件非空：

```powershell
Get-Content "docs\superpowers\plans\perf-baseline-2026-06-21.csv" | Measure-Object -Line
# 期望：3 行（header + 2 scenarios）
```

---

## 步骤 4：自测对比脚本

```powershell
.\scripts\compare-benchmark.ps1 -SelfTest
# 期望：compare-benchmark.ps1 self-test passed.  exit 0
```

用基线自比（应 PASS）：

```powershell
.\scripts\compare-benchmark.ps1 `
  -BaselineCsv "docs\superpowers\plans\perf-baseline-2026-06-21.csv" `
  -CurrentCsv "docs\superpowers\plans\perf-baseline-2026-06-21.csv"
```

---

## 重构 Task 后的对比用法

完成涉及热路径的改动后，再跑 PERF 场景 1 次（或 3 次取最新），与基线对比：

```powershell
$current = "$env:USERPROFILE\Documents\WordTools\wordtools-benchmark.csv"

.\scripts\compare-benchmark.ps1 `
  -BaselineCsv "docs\superpowers\plans\perf-baseline-2026-06-21.csv" `
  -CurrentCsv $current `
  -ScenarioId All
```

单场景：

```powershell
.\scripts\compare-benchmark.ps1 `
  -BaselineCsv "docs\superpowers\plans\perf-baseline-2026-06-21.csv" `
  -CurrentCsv $current `
  -ScenarioId PERF-01
```

**exit 0** = 通过；**exit 1** = 性能回归或 count 列变化。

### 默认阈值

| 指标 | 上限 |
|------|------|
| `total_seconds` | 基线 × 1.05 |
| `insert_images_ms` | 基线 × 1.05 |
| `add_picture_ms` | 基线 × 1.05 |
| `clear_numbering_ms` | 基线 × 1.05 |
| `cell_availability_ms` | 基线 × 1.08 |
| `progress_ui_ms` | 基线 × 1.10 |
| `*_count` 列 | **必须与基线完全相等** |

自定义阈值示例：

```powershell
.\scripts\compare-benchmark.ps1 `
  -BaselineCsv "docs\superpowers\plans\perf-baseline-2026-06-21.csv" `
  -CurrentCsv $current `
  -MaxTotalSecondsRatio 1.03
```

---

## PERF-03（50 张，Phase 4 可选）

1. 生成 50 张小图：

```powershell
.\scripts\generate-perf-50-images.ps1
# 输出：automation\assets\images\selected-50\01.jpg … 50.jpg
```

2. 手动选中 50 张插入 1–3 次
3. `capture-benchmark-baseline.ps1 -ScenarioId PERF-03` 追加到基线（需 `-Force` 合并时先备份）

---

## 故障排查

| 现象 | 原因 | 处理 |
|------|------|------|
| CSV 未生成 | 未开详细日志/基准 CSV | Ribbon 两项都开 |
| capture 报 No eligible rows | status≠Completed 或 cancelled=True | 重新跑完整个插入流程 |
| compare 报 count 变化 | 拆分改变了 COM 调用次数 | **必须 revert**，属性能回归 |
| ratio 超阈但 count 正常 | 机器负载 / 首次冷启动 | 同场景再跑 3 次，取 median 对比 |

---

## 相关文件

| 文件 | 用途 |
|------|------|
| `scripts/compare-benchmark.ps1` | 基线 vs 当前对比 |
| `scripts/capture-benchmark-baseline.ps1` | 从原始 CSV 生成 median 基线 |
| `scripts/lib/BenchmarkCsv.ps1` | CSV 解析/写入共享库 |
| `docs/superpowers/plans/perf-baseline-2026-06-21.csv` | 锁定基线（header + median 行） |
| `docs/superpowers/plans/2026-06-21-god-module-split-phase2.md` | 完整拆分方案与 PERF 门禁 |
