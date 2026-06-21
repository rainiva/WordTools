# WordTools 上帝模块拆分重构方案（Phase 2）

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在**不改变用户可感知行为、不降低插图性能**的前提下，将仍超过 400 行的上帝模块拆分为职责单一、可测、可维护的小模块；以自动化回归 + 性能基准对比为安全网，确保稳定可靠。

**Architecture:** 采用 **Strangler Fig（绞杀者）** 模式——先锁行为基线（特征测试 + E2E），再逐块提取；旧类保留为 **薄门面（facade）** 维持调用方不变，内部委托新模块；纯逻辑优先单元测试，COM/Word 行为由 `automation/` E2E 守护。

**Tech Stack:** .NET Framework 4.8（主项目）, xUnit + .NET 8.0-windows（测试）, pytest + PowerShell（多宿主 E2E）, Microsoft.Office.Interop.Word PIA

## Global Constraints

- 目标框架 .NET Framework 4.8，不可升级
- 测试项目 .NET 8.0-windows + xUnit，通过 `<Compile Include="..">` 链接源文件
- 新增接口放在 `Services/Abstractions/`，适配器放在 `Services/Adapters/`
- 新增文件必须在 `WordTools.csproj` 和 `WordTools.Tests.csproj` 中注册
- 中文注释/文案 UTF-8 编码安全：只用 `StrReplace`/`Write` 精确补丁，禁止 PowerShell 整文件覆盖
- **每个 Task 结束必须：** `msbuild WordTools.sln /p:Configuration=Release` 0 errors + `dotnet test` 全绿
- **涉及插图/编号行为的 Task 还必须：** 跑 Phase 0 定义的 E2E 子集（见下文）
- **涉及 `ProgressService` / `BatchInsert*` / `ImageService` / `TableService` 热路径的 Task 还必须：** 跑 Phase 0 性能基准对比（PERF 门禁，见「插图性能零回归策略」）
- 禁止在同一 commit 中混合「行为修复」与「结构拆分」
- 禁止在拆分 commit 中引入任何可能增加 COM 调用次数、UI 刷新次数、GC 频率的改动
- 拆分 commit 消息前缀：`refactor:`；测试 commit 前缀：`test:`；基准 commit 前缀：`perf:`

---

## 现状基线（2026-06-21）

### 已完成（Phase 1–6，见 `2026-06-20-refactor-split-and-interfaces.md`）

| 项 | 状态 |
|----|------|
| 10 个 Abstractions 接口 | ✅ |
| 8 个 Adapters | ✅ |
| TableService ↔ TableNumberingService 文件级拆分 | ✅ |
| ProgressService 部分提取（HP/Escape/Window/Error/Result） | ✅ |
| RibbonController、UiToolkit、InsertPhotosOrchestrator | ✅ |
| 单元测试 | ✅ 74 通过 |

### 仍待处理（上帝模块）

| 文件 | 行数 | 主要问题 |
|------|------|---------|
| `TableNumberingService.cs` | 1178 | `RefreshTableNumbering` ~600 行；`ClearTableNumbering` ~240 行 |
| `ProgressService.cs` | 1039 | `ProcessFileBatch` ~330 行；仍直接编排全部静态服务 |
| `TableService.cs` | 762 | 验证 + 合并检测 + 行操作 + 内容行 4 类职责 |
| `ImageService.cs` | 702 | COM 重试 + 插入 + 批量操作混合 |
| `InsertPhotosForm.cs` | 629 | Layout + 配置持久化 + 事件处理混合 |
| `ConfigService.cs` | 481 | 34 个 public 方法，接口未接入运行时 |

### 目标文件结构（Phase 2 完成后）

```
WordTools/Services/
├── BatchInsert/
│   ├── BatchInsertExecutor.cs          (原 ProcessFileBatch 核心循环, ≤350行)
│   ├── BatchInsertContext.cs           (批次状态 DTO, ≤80行)
│   ├── BatchInsertMemoryManager.cs     (GC/DoEvents 策略, ≤60行)
│   └── BatchInsertStatusReporter.cs    (状态栏/进度更新, ≤80行)
├── Numbering/
│   ├── TableNumberingFacade.cs         (对外 API，薄门面, ≤80行)
│   ├── NumberingScanService.cs         (扫描/配对/快速路径, ≤350行)
│   ├── NumberingClearService.cs        (清除编号, ≤280行)
│   ├── NumberingWriteService.cs        (写入文本/SEQ 域, ≤150行)
│   ├── NumberingSequenceCalculator.cs  (CalculateNextSequenceNumber 等, ≤120行)
│   └── NumberingTextParser.cs          (纯文本/正则解析, ≤100行)
├── Table/
│   ├── TableSelectionHelper.cs         (Selection/当前表, ≤80行)
│   ├── TableCellAvailabilityService.cs (GetCellAvailability 族, ≤350行)
│   ├── TableStructureService.cs        (EnsureRow/AdjustColumns, ≤200行)
│   └── TableContentRowService.cs         (标题/描述/NA 行, ≤180行)
├── Image/
│   ├── ComRetryExecutor.cs             (COM 重试, ≤250行)
│   ├── ImageInsertService.cs           (InsertImage*, ≤200行)
│   └── ImageBatchOperations.cs         (BatchResize/PreAllocate, ≤150行)
├── Config/
│   ├── ConfigServiceAdapter.cs         (IConfigService 实现)
│   ├── RegistryConfigStore.cs          (注册表读写)
│   └── DocumentConfigStore.cs          (文档属性读写)
├── ProgressService.cs                  (薄编排, ≤250行)
├── TableService.cs                     (薄门面, ≤120行)
├── TableNumberingService.cs            (deprecated 别名 → Facade, ≤50行)
└── ImageService.cs                     (薄门面, ≤120行)
```

---

## 稳定性策略（全局，所有 Phase 必须遵守）

### 1. 行为锁定优先于拆分

任何提取前，必须先有**可失败的特征测试**或 **E2E 用例**覆盖该行为。未见 RED 不得动实现。

### 2. 三层回归金字塔

| 层级 | 工具 | 何时跑 | 覆盖 |
|------|------|--------|------|
| L1 纯逻辑单元测试 | `dotnet test` | 每个 Task | 文本解析、错误分类、间隔计算、配置规范化 |
| L2 无 UI 编排测试 | `dotnet test` + Headless adapters | 每个涉及插图的 Task | `InsertPhotosOrchestrator.Execute` + `CapturingNotificationService` |
| L3 Word COM E2E | `automation/` pytest | 每个 Phase 结束 + 发布前 | AC-B01–B06 + 编号刷新 |
| **L4 性能基准** | `scripts/compare-benchmark.ps1` | Phase 1–4 结束 + 涉及热路径 Task | PERF-01–03 vs baseline CSV |

### 3. Strangler 门面规则

- 旧 public API（如 `TableNumberingService.RefreshTableNumbering`）**签名不变**
- 实现改为单行委托：`NumberingScanService.Refresh(...)` 或 `TableNumberingFacade.Refresh(...)`
- 调用方（`ProgressService`、`NumberingRefreshService`、`RibbonController`）在 Phase 2 内**不强制改**为接口注入

### 4. InternalsVisibleTo 策略

将以下 private 纯逻辑方法改为 `internal`，供测试链接：

- `NumberingTextParser.*`
- `TableService.ShouldTreatCellCountMismatchAsMerged`（已是 public，保持）
- `ConfigService` 中无 COM 的 registry 读写 helper（如后续提取）

在 `WordTools/Properties/AssemblyInfo.cs` 添加：

```csharp
[assembly: InternalsVisibleTo("WordTools.Tests")]
```

### 5. 回滚策略

- 每 Task 独立 commit，消息含模块名
- Phase 验收失败时：`git revert` 单个 Task commit，不 revert 整 Phase
- E2E 失败时停止后续 Phase，先修回归再继续
- **性能门禁失败时：** 同 E2E——停止后续 Phase；若单 Task 引入回归，revert 该 Task 后再跑 PERF 对比

---

## 插图性能零回归策略（硬性约束）

> **原则：拆分 = 搬家，不是重写。** 性能敏感代码只允许「剪切-粘贴-委托」，不允许「顺手优化」。

### 现有性能机制（拆分后必须完整保留）

| 机制 | 位置 | 作用 | 拆分要求 |
|------|------|------|----------|
| 高性能模式 | `HighPerformanceModeController` | 关闭 `ScreenUpdating` / `DisplayAlerts` | **整类不动**；`BatchInsertExecutor` 仍通过注入调用 |
| 批量上下文缓存 | `ImageInsertionBatchContext` | 行可用性缓存、浮动形状索引 | **同一实例**贯穿单次插入；`ClearRowAvailability` 触发时机不变 |
| 性能分段计时 | `InsertionPerformanceDiagnostics` | AddPicture / CellAvailability 等 ms 统计 | 字段与 `Record*` 调用点 **1:1 保留** |
| 基准 CSV | `BenchmarkLogService` + `BenchmarkLogEntry` | 写入 `wordtools-benchmark.csv` | 列名与写入时机不变 |
| 动态间隔 | `HighPerformanceModeController` | `_refreshInterval` / `_statusBarUpdateInterval` 按文件数缩放 | 计算公式 **不得改** |
| 内存/GC 策略 | `ProgressService.CleanupMemory` | 分级 GC + 500MB 水位 | 提取到 `BatchInsertMemoryManager` 时 **逻辑字节级相同** |
| 行数缓存 | `ProcessFileBatch` 内 `cachedRowCount` | 减少 `tbl.Rows.Count` COM 调用 | 提取时作为 `BatchInsertMutableState` 字段，更新时机不变 |
| 文件预检 | `FileService.BatchValidateImageFiles` | 循环外批量 IO | 仍在循环 **之前** 调用，不得改为逐张检 |
| 预分配行 | `ImageService.PreAllocateRows` | 大批量前扩表 | 调用条件与参数不变 |
| COM 重试 | `ImageService` `#region COM 重试` | Word 忙时退避 | 重试次数、间隔、HRESULT 判断 **不得改** |

### 性能热点不可改清单（diff 审查必查）

以下若在拆分 commit 中出现 **调用次数增加、条件变化、或调用顺序变化**，该 commit **必须 reject**：

1. `_perfController.Enter()` / `Exit()` — 仍各 **1 次** / 插入 run（成功/失败/finally 路径与现网相同）
2. `new ImageInsertionBatchContext(...)` — 每次 `InsertPhotosWithProgress` / `InsertSelectedPhotosWithProgress` 仍 **1 个**
3. `batchContext.ClearRowAvailability()` — 仅在 **新增行后** 调用（与现 `ProcessFileBatch` 相同行逻辑）
4. `TableService.AdjustTableColumns(tbl, 2)` — 循环前 **1 次** + 列不足时按需（不新增额外全表扫描）
5. `TableService.FindNextSuitableImageRow(..., batchContext)` — 每张图路径不变；不得降级为无 context 重载
6. `ImageService.InsertImageFast(..., batchContext)` — 热路径不变
7. `_statusBarUpdateInterval` / `_memoryCleanInterval` / `_fullGcInterval` / `_saveInterval` — 赋值公式不变
8. `CleanupMemory(processedCount)` — 触发间隔不变
9. `TableNumberingService.ClearTableNumbering` — 仍仅在 `needClearNumbering==true` 时调用；`skippedClear` 语义不变
10. 循环内 `DoEvents()` — 不得新增；仅保留现有 Progress UI / 取消检测路径

### 允许的拆分方式 vs 禁止的拆分方式

| ✅ 允许 | ❌ 禁止 |
|---------|---------|
| 方法体原样移到新类，门面一行委托 | 借机改算法（如合并两次扫描为一次但改变顺序） |
| 提取 DTO 传递现有字段 | 新增 LINQ/反射/字符串拼接在热循环内 |
| `internal` 可见性调整 | 将 static 改 instance 导致重复分配重量级对象 |
| 提取 helper 但保持 call site 次数 | 「更清晰」而在每张图多调一次 COM |
| 把 `Stopwatch` 计时块一起移动 | 删除或合并计时段导致无法对比 |

### PERF 基准场景（Phase 0 锁定，Phase 1–4 每 Phase 结束对比）

| ID | 场景 | 文件数 | 编号 | 用途 |
|----|------|--------|------|------|
| PERF-01 | 选中文件插入（AC-B03 等价） | 4 | 开 | 主回归：插图 + 编号 + 描述 |
| PERF-02 | 文件夹插入（AC-B04 等价） | 5 | 开 | 含子目录标题行 |
| PERF-03 | 大批量模拟 | 50* | 开 | 间隔/GC/预分配路径 |

\*PERF-03：用 `automation/assets/images/` 复制或脚本生成 50 张同尺寸小图；若 CI 超时可在本地/发布前跑，但 **Phase 2 完成前必须跑至少 1 次**。

### PERF 采集方式（复用现有 Benchmark CSV，不新造框架）

1. Release 构建并注册加载项
2. Ribbon 开启 **详细日志 + 性能基准 CSV**
3. 在空白表格模板执行 PERF-01 / PERF-02（PERF-03 可选）
4. 从文档目录或 `%USERPROFILE%\Documents\wordtools-benchmark.csv` 取最后一行
5. 保存到 `docs/superpowers/plans/perf-baseline-2026-06-21.csv`（Phase 0 新建）

**对比脚本（Phase 0 写入 repo）：** `scripts/compare-benchmark.ps1`

```powershell
param(
    [Parameter(Mandatory)] [string]$BaselineCsv,
    [Parameter(Mandatory)] [string]$CurrentCsv,
    [double]$MaxTotalSecondsRatio = 1.05,
    [double]$MaxInsertImagesMsRatio = 1.05,
    [double]$MaxAddPictureMsRatio = 1.05
)
# 取两文件末行同 run_mode，对比 total_seconds / insert_images_ms / add_picture_ms
# 任一项超过 ratio → exit 1
```

### PERF 验收阈值（比「总时间 ±10%」更严格）

| 指标 | 阈值 | 说明 |
|------|------|------|
| `total_seconds` | ≤ 基线 × **1.05** | 端到端用户感知 |
| `insert_images_ms` | ≤ 基线 × **1.05** | 核心插图阶段 |
| `add_picture_ms` | ≤ 基线 × **1.05** | COM AddPicture 累计 |
| `cell_availability_ms` | ≤ 基线 × **1.08** | 允许极小浮动（缓存命中路径敏感） |
| `progress_ui_ms` | ≤ 基线 × **1.10** | UI 更新非核心 |
| `clear_numbering_ms` | ≤ 基线 × **1.05** | Phase 1 编号拆分门禁 |
| `CellAvailabilityCount` / `AddPictureCount` | **与基线相等** | 调用次数不得变 |
| E2E AC-B03 InlineShapes 数 | **= 4** | 结果正确性 |

> 若总时间持平但 `AddPictureCount` 增加——视为 **性能回归**，即使 total_seconds 未超阈值。

### 涉及插图的 Phase 额外门禁

| Phase | PERF 要求 |
|-------|-----------|
| 1（编号） | PERF-01：`clear_numbering_ms` + `total_seconds` 不超阈 |
| 2（BatchInsert） | PERF-01 + PERF-02：**全部核心指标**不超阈 |
| 3（Table） | PERF-01 + PERF-02；重点看 `cell_availability_ms` |
| 4（Image） | PERF-01 + PERF-02 + **PERF-03**；重点看 `add_picture_ms` |
| 5（Config/Form） | PERF-01 抽检（Form 不应进热路径） |

---

## Phase 0：行为基线锁定（必须先完成）

> **Gate：** Phase 1 开始前，Phase 0 全部 checkbox 必须为 `[x]`。

### Task 0.1：记录基线指标

- [ ] 运行并记录：

```powershell
# 行数基线
Get-ChildItem WordTools -Recurse -Filter *.cs |
  Where-Object { $_.FullName -notmatch '\\(obj|bin)\\' } |
  ForEach-Object { [PSCustomObject]@{ Lines=(Get-Content $_.FullName | Measure-Object -Line).Lines; File=$_.FullName.Replace('...\WordTools\','') } } |
  Sort-Object Lines -Descending | Select-Object -First 15

dotnet test WordTools.Tests/WordTools.Tests.csproj --verbosity minimal
msbuild WordTools.sln /p:Configuration=Release
```

- [ ] 将输出粘贴到 `docs/superpowers/plans/2026-06-21-baseline-metrics.txt`（新建）

**验收：** 文件存在；74 测试通过；Release 0 errors。

---

### Task 0.2：编号纯逻辑特征测试（RED → GREEN）

**Files:**
- Create: `WordTools/Services/Numbering/NumberingTextParser.cs`
- Create: `WordTools.Tests/NumberingTextParserTests.cs`
- Modify: `WordTools/Services/TableNumberingService.cs`（委托给 Parser，行为不变）
- Modify: `WordTools/WordTools.csproj`, `WordTools.Tests/WordTools.Tests.csproj`
- Modify: `WordTools/Properties/AssemblyInfo.cs`

**Interfaces:**
- Produces: `NumberingTextParser.TryParseLeadingNumber(string text, out int number)` — 从单元格文本提取 `1.`、`1-1` 等

- [ ] **Step 1: 写失败测试**

```csharp
// WordTools.Tests/NumberingTextParserTests.cs
using WordTools.Services.Numbering;
using Xunit;

namespace WordTools.Tests
{
    public class NumberingTextParserTests
    {
        [Theory]
        [InlineData("1. 说明", 1)]
        [InlineData("12.图", 12)]
        [InlineData("图 1-1", 0)] // 非 leading digit 格式 — 按现有 CalculateNextSequenceNumber 行为
        [InlineData("", 0)]
        public void TryParseLeadingNumber_MatchesExistingBehavior(string text, int expected)
        {
            Assert.Equal(expected > 0, NumberingTextParser.TryParseLeadingNumber(text, out int n));
            if (expected > 0) Assert.Equal(expected, n);
        }
    }
}
```

- [ ] **Step 2:** 确认 FAIL（类不存在）
- [ ] **Step 3:** 从 `TableNumberingService.CalculateNextSequenceNumber` 提取 regex 逻辑到 `NumberingTextParser`
- [ ] **Step 4:** `TableNumberingService` 改为调用 Parser（行为不变）
- [ ] **Step 5:** 全绿 + commit `test: add NumberingTextParser characterization tests`

**验收：** 新测试 ≥3 个；`TableNumberingService` 行数减少 ≥20 行；E2E 子集暂不要求（纯逻辑）。

---

### Task 0.3：定义 E2E 必跑子集（Phase 2 门禁）

**Files:**
- Modify: `automation/docs/BATCH-INSERT-ACCEPTANCE.md`（追加 Phase 2 门禁章节）

- [ ] 在文档中追加：

```markdown
## Phase 2 Refactor Gate（每个涉及 COM 插图的 Task 后必跑）

Prerequisites: Release build, Word 关闭

dotnet test                                    # L1
cd automation && python -m pytest tests/test_batch_insert_e2e.py -m integration -v   # L3: AC-B01–B06
```

- [ ] 本地跑通一次并记录耗时到 baseline 文件

**验收：** AC-B01–B06 全绿；文档已更新。

---

### Task 0.4：Headless 编排冒烟测试补强

**Files:**
- Modify: `WordTools.Tests/Tests.cs` 或 Create: `WordTools.Tests/BatchInsertOrchestratorTests.cs`

- [ ] 新增测试：用 fake/minimal 方式验证 `InsertPhotosOrchestrator` 在 `InsertPhotosAutomationGate.IsEnabled=true` 时，`request==null` 不抛异常（已有则跳过）
- [ ] 新增测试：`InsertPhotosExecutionServices.CreateHeadless(4)` 返回非 null 的三依赖

**验收：** `dotnet test` 测试数 ≥76；0 失败。

---

### Task 0.5：插图性能基准锁定（PERF 门禁前置）

**Files:**
- Create: `scripts/compare-benchmark.ps1`
- Create: `docs/superpowers/plans/perf-baseline-2026-06-21.csv`
- Modify: `automation/docs/BATCH-INSERT-ACCEPTANCE.md`（追加 PERF 章节）

- [ ] **Step 1:** 编写 `scripts/compare-benchmark.ps1`（见上文脚本骨架，补全 CSV 解析与 exit code）
- [ ] **Step 2:** Release 构建 + 注册加载项；开启 Ribbon「详细日志 + 性能基准 CSV」
- [ ] **Step 3:** 手动执行 PERF-01、PERF-02，各跑 **3 次**，取 `total_seconds` / `insert_images_ms` / `add_picture_ms` **中位数**写入 baseline CSV
- [ ] **Step 4:** 记录 `CellAvailabilityCount`、`AddPictureCount` 等 count 列（作为次数基线）
- [ ] **Step 5:** 在 `BATCH-INSERT-ACCEPTANCE.md` 追加 PERF-01–03 操作步骤与阈值表

**验收：**
- `perf-baseline-2026-06-21.csv` 至少 2 行（PERF-01、PERF-02）
- `compare-benchmark.ps1 -BaselineCsv ... -CurrentCsv ...` 对同文件自比 exit 0
- baseline 含：`run_mode,total_seconds,insert_images_ms,add_picture_ms,cell_availability_ms,clear_numbering_ms,cell_availability_count,add_picture_count`

---

## Phase 1：TableNumberingService 内部拆分（最高优先级）

> **风险最高、体积最大。** 严格按 Scan → Clear → Write → Facade 顺序提取。

### 方法归属映射

| 方法 | 新归属 | 约行数 |
|------|--------|--------|
| `RefreshTableNumbering` | `NumberingScanService.Refresh` | ~600 |
| `ClearTableNumbering` | `NumberingClearService.Clear` | ~240 |
| `AddNumberingToDescriptionRows` | `NumberingWriteService.AddToDescriptionRows` | ~120 |
| `CalculateNextSequenceNumber` | `NumberingSequenceCalculator` | ~150 |
| `InsertNumberText`, `InsertSeqField` | `NumberingWriteService` | ~80 |
| 文本/regex 辅助 | `NumberingTextParser` | ~100 |

### Task 1.1：创建 Numbering 目录与 NumberingClearService

**Files:**
- Create: `WordTools/Services/Numbering/NumberingClearService.cs`
- Modify: `WordTools/Services/TableNumberingService.cs`
- Modify: `WordTools/WordTools.csproj`, `WordTools.Tests/WordTools.Tests.csproj`

- [ ] 复制 `ClearTableNumbering` 实现到 `NumberingClearService.Clear`（**字节级相同逻辑**）
- [ ] `TableNumberingService.ClearTableNumbering` → 委托 `NumberingClearService.Clear`
- [ ] 构建 + `dotnet test`
- [ ] **E2E 子集**（编号相关手动：Ribbon 刷新编号 — 见 Phase 1 验收）
- [ ] Commit: `refactor: extract NumberingClearService`

---

### Task 1.2：提取 NumberingWriteService + NumberingSequenceCalculator

- [ ] 移动 `InsertNumberText`, `InsertSeqField`, `AddNumberingToDescriptionRows`, `CalculateNextSequenceNumber`
- [ ] 门面委托
- [ ] 构建 + 测试 + commit

---

### Task 1.3：提取 NumberingScanService（RefreshTableNumbering）

- [ ] 移动 `RefreshTableNumbering` 主体到 `NumberingScanService.Refresh`
- [ ] 注入 `NumberingWriteService` / `NumberingSequenceCalculator` 作为 static 或参数
- [ ] `TableNumberingService.RefreshTableNumbering` 一行委托
- [ ] 构建 + 测试 + **E2E 子集** + commit

---

### Task 1.4：TableNumberingFacade + 瘦身验证

- [ ] 创建 `TableNumberingFacade.cs` 统一对外；`TableNumberingService` 标记 `[Obsolete]` 或保持别名
- [ ] 行数验收：

| 文件 | 目标 |
|------|------|
| `TableNumberingService.cs` | ≤ 50（仅门面） |
| `NumberingScanService.cs` | ≤ 400 |
| `NumberingClearService.cs` | ≤ 300 |
| 任一 Numbering/*.cs | ≤ 400 |

- [ ] Phase 1 验收（见文末验收标准表）

---

## Phase 2：ProgressService → BatchInsert 模块

### Task 2.1：BatchInsertContext（状态 DTO）

**Files:**
- Create: `WordTools/Services/BatchInsert/BatchInsertContext.cs`

- [ ] 从 `ProcessFileBatch` 参数列表提取不可变上下文 record/class：

```csharp
internal sealed class BatchInsertContext
{
    public Table Table { get; }
    public float MinHeightPoints { get; }
    public bool NeedDescription { get; }
    public bool NeedAutoNumbering { get; }
    public WdParagraphAlignment NumberAlignment { get; }
    public int NumberPosition { get; }
    // rowIndex, counters, lists — 可变状态单独 BatchInsertMutableState
}
```

- [ ] 仅创建类型 + 测试构造，**尚未移动逻辑**
- [ ] Commit: `refactor: add BatchInsertContext DTO`

---

### Task 2.2：BatchInsertMemoryManager + BatchInsertStatusReporter

- [ ] 提取 `CleanupMemory` → `BatchInsertMemoryManager.MaybeCollect`
- [ ] 提取 `UpdateStatusBar` → `BatchInsertStatusReporter.Update`
- [ ] `ProgressService` 委托；测试 + commit

---

### Task 2.3：BatchInsertExecutor（核心）

- [ ] 移动 `ProcessFileBatch` → `BatchInsertExecutor.ProcessBatch`
- [ ] 构造函数注入：`IWordApplicationContext`, `EscapeKeyMonitor`, `BatchInsertStatusReporter`, `BatchInsertMemoryManager`
- [ ] **`ImageInsertionBatchContext` 由调用方创建并传入，Executor 不得内部 new 第二个实例**
- [ ] **`InsertionPerformanceDiagnostics` 引用从 ProgressService 传入 Executor，`_activeDiagnostics` 赋值链不断**
- [ ] 移动后用 `git diff` 核对「性能热点不可改清单」10 条
- [ ] `ProgressService` 两个 public 入口改为组装 Executor 并调用
- [ ] **必须跑 E2E 子集 AC-B01–B06**
- [ ] **必须跑 PERF-01 + PERF-02，`compare-benchmark.ps1` exit 0**
- [ ] Commit: `refactor: extract BatchInsertExecutor from ProgressService`

---

### Task 2.4：ProgressService 瘦身验证

| 文件 | 目标 |
|------|------|
| `ProgressService.cs` | ≤ 250 |
| `BatchInsertExecutor.cs` | ≤ 380 |

- [ ] `ProgressService` 仅保留：构造、两个 public 入口、关闭进度窗、组装子组件
- [ ] Phase 2 验收

---

## Phase 3：TableService 职责拆分

### Task 3.1：TableSelectionHelper

- [ ] 移动 `IsSelectionInTable`, `IsSelectionInFirstColumn`, `GetCurrentTable`
- [ ] `TableService` 门面委托
- [ ] 更新 `NumberingRefreshService` 无需改（仍调 TableService）

### Task 3.2：TableCellAvailabilityService

- [ ] 移动 `#region 表格验证` + `#region 合并单元格检测` 全部方法
- [ ] 保留 `ImageInsertionBatchContext` 参数签名

### Task 3.3：TableStructureService + TableContentRowService

- [ ] 移动 `#region 表格操作` → Structure
- [ ] 移动 `#region 标题行和描述行` → ContentRow
- [ ] E2E 子集 + Phase 3 验收

---

## Phase 4：ImageService 拆分

### Task 4.1：ComRetryExecutor

- [ ] 提取 `#region COM 重试`（~240 行）到独立类
- [ ] `ImageService` 门面委托

### Task 4.2：ImageInsertService + ImageBatchOperations

- [ ] 插入逻辑 / 批量操作分离
- [ ] **`ComRetryExecutor` 与 `InsertImageFast` 同文件或同程序集紧邻，避免跨层额外 indirection**
- [ ] E2E 子集 + **PERF-03（50 张）** + Phase 4 验收

---

## Phase 5：Config 与 Form 瘦身（低 COM 风险）

### Task 5.1：ConfigServiceAdapter

- [ ] 实现 `IConfigService` 包装现有 static `ConfigService`
- [ ] `InsertPhotosForm` **暂不强制**改用接口（可选 Task 5.2）

### Task 5.2：InsertPhotosForm — 提取 InsertPhotosFormController

- [ ] 将 `#region Configuration` 读写 Config 逻辑移到 `InsertPhotosFormController.cs`
- [ ] Form 仅保留 Layout + 事件绑定
- [ ] 目标：`InsertPhotosForm.cs` ≤ 400 行

### Task 5.3：FileService UI 清理（若未完成）

- [ ] 确认 `FileService` 无 `FolderBrowserDialog` / `OpenFileDialog`
- [ ] 若有，移至 Form/Controller

---

## Phase 6：适配器接入与债务清理（可选，发布前）

### Task 6.1：ProgressService 依赖静态服务 → 可测试注入（长期）

- [ ] 新增 `IBatchInsertExecutor` 仅用于测试替身
- [ ] 生产仍用默认实现

---

## 验收标准（完整）

### A. 硬性指标（Phase 2 全部完成后）

| 指标 | 当前 (2026-06-21) | 目标 | 验证命令 |
|------|-------------------|------|----------|
| 最大单文件行数 | 1178 | **≤ 400** | PowerShell 行数统计 |
| `ProgressService.cs` | 1039 | **≤ 250** | 同上 |
| `TableNumberingService.cs` | 1178 | **≤ 50**（门面） | 同上 |
| `TableService.cs` | 762 | **≤ 120**（门面） | 同上 |
| `ImageService.cs` | 702 | **≤ 120**（门面） | 同上 |
| 单元测试数 | 74 | **≥ 95** | `dotnet test` |
| Release 构建 | 0 errors | **0 errors** | `msbuild` |
| L3 E2E AC-B01–B06 | 通过 | **通过** | pytest integration |
| **PERF-01/02 对比** | 未锁定 | **≤ 基线 ×1.05** | `compare-benchmark.ps1` |
| **AddPictureCount** | — | **= 基线** | benchmark CSV count 列 |

### B. 架构指标

| 指标 | 目标 |
|------|------|
| 单个类 `#region` 数 | ≤ 2 |
| 单个 public 方法体 | ≤ 80 行（COM 循环除外，循环体提取 helper） |
| 服务层 WinForms 引用 | 仅 Form/Adapter；`Services/BatchInsert/*`、`Services/Numbering/*` **无** `System.Windows.Forms` |
| 门面类 | 每个原上帝模块保留 1 个，仅委托 |
| 新增模块测试覆盖 | 每个 `*Parser`、`*Calculator`、`*Classifier` 有 Theory 测试 |

### C. 行为不变（必须手动 + 自动双重验证）

| 场景 ID | 场景 | 自动 | 手动抽检 |
|---------|------|------|----------|
| AC-B01 | 光标不在表格 | ✅ pytest | — |
| AC-B02 | 光标在第 2 列 | ✅ pytest | — |
| AC-B03 | 选 4 文件 + 自动编号 | ✅ pytest | 每 Phase 结束抽检 1 次 |
| AC-B04 | 文件夹根+子目录 | ✅ pytest | — |
| AC-B05 | 单图 + N/A | ✅ pytest | — |
| AC-B06 | 中途取消 | ✅ pytest | — |
| AC-N01 | Ribbon 刷新编号 | 待补 pytest | **每 Phase 1 Task 后必做** |
| AC-N02 | 清除后重插编号 | 手动 | Phase 1 结束 |
| AC-R01 | Ribbon 日志开关 | 手动 | Phase 5 后 |

### D. 每 Phase 门禁（缺一不可）

| Phase | 合并前必须满足 |
|-------|----------------|
| 0 | 基线文件 + NumberingTextParser 测试 + E2E 文档 + **perf-baseline CSV + compare 脚本** + ≥76 单元测试 |
| 1 | `TableNumberingService` ≤50 行 + AC-N01 + dotnet + E2E + **PERF-01 不超阈** |
| 2 | `ProgressService` ≤250 行 + E2E AC-B01–B06 + **PERF-01/02 不超阈 + count 列相等** |
| 3 | `TableService` ≤120 行 + E2E + **PERF-01/02 不超阈** |
| 4 | `ImageService` ≤120 行 + E2E + **PERF-01/02/03 不超阈** |
| 5 | `InsertPhotosForm` ≤400 行 + 构建测试全绿 + **PERF-01 抽检** |
| 6 | 可选；不阻塞 Phase 2 发布 |

### E. 可靠性专项标准

1. **零行为 commit 原则：** 拆分 commit 的 diff 中不得出现业务逻辑条件变更（`if` 分支增减、常量值变化、异常吞掉方式变化）
2. **COM 异常处理不变：** 提取时保持原 `catch (COMException)` / `SafeIgnore` 位置与消息
3. **编码完整性：** 每个 Phase 结束后 spot-check：`InsertPhotosForm.cs`、`TableNumberingService.cs` 中文注释无乱码
4. **可回滚：** 每个 Task 单 commit；Phase 验收失败可 revert 最后 1–3 个 commit 恢复绿 build

### F. 插图性能专项标准（与 E 同级硬性）

1. **拆分 ≠ 优化 ≠ 重构算法：** 性能相关 PR 标题含 `[perf-neutral]` 或说明「仅移动代码」
2. **Benchmark CSV 对比优先于主观感受：** 不以「感觉差不多」合并 Phase 2+
3. **核心指标阈值：** 见「PERF 验收阈值」表（total/insert_images/add_picture ≤ **105%**）
4. **调用次数零增长：** `AddPictureCount`、`CellAvailabilityCount` 等与基线 **精确相等**
5. **Phase 2 完成定义：** 在**同一台机器、同一 Word 版本、Release 构建**下，PERF-01/02 各 3 次中位数均过线
6. **性能回归处置：** revert → 定位多出的 COM/DoEvents → 修复后再跑 PERF；禁止「后续 Phase 再补性能」

---

## 风险与缓解

| 风险 | 影响 | 缓解 |
|------|------|------|
| `RefreshTableNumbering` 提取引入编号错乱 | 高 | Phase 0 Parser 测试 + Phase 1 每 Task 后 AC-N01 手动 |
| `ProcessFileBatch` 状态变量遗漏 | 高 | 先 DTO 化再移动；E2E AC-B03/B04/B05 |
| COM 引用计数/RCW 生命周期 | 中 | 不改 `try/finally` 结构；不新增 `Marshal.ReleaseComObject` |
| 中文编码损坏 | 高 | AGENTS.md 规则；禁止 shell 改源码 |
| 测试链接 csproj 遗漏 | 中 | 每个新文件同步两个 csproj |
| 拆分范围膨胀 | 中 | Phase 5/6 可延期；**Phase 1–2 为 MVP** |
| **拆分引入额外 COM/UI 调用** | **高** | 性能热点清单 + count 列对比 + revert |
| **BatchInsertExecutor 丢失 batchContext** | **高** | context 由 ProgressService 创建并传入，禁止 Executor 内 new |
| **Extract 时「顺手」改 GC/间隔** | **高** | `BatchInsertMemoryManager` 字节级复制；单元测试已有间隔 Theory |

---

## 推荐实施顺序与工期估算

| 顺序 | Phase | 预估 | 说明 |
|------|-------|------|------|
| 1 | Phase 0 | 0.5–1 天 | 不可跳过 |
| 2 | Phase 1 | 2–3 天 | 编号核心 |
| 3 | Phase 2 | 2–3 天 | 插图核心 |
| 4 | Phase 3 | 1–2 天 | 表格辅助 |
| 5 | Phase 4 | 1–2 天 | 图片/COM |
| 6 | Phase 5 | 1 天 | 配置/UI |
| 7 | Phase 6 | 可选 | 债务 |

**MVP 发布线：** Phase 0 + 1 + 2 完成即可显著降低上帝模块风险；Phase 3–5 可迭代发布。

---

## 执行选项

Plan complete and saved to `docs/superpowers/plans/2026-06-21-god-module-split-phase2.md`.

**1. Subagent-Driven（推荐）** — 每个 Task 派生子 agent，Task 间人工/自动 review

**2. Inline Execution** — 本会话按 Phase 0 → 1 → 2 连续执行，Phase 门禁处暂停确认

**Which approach?**
