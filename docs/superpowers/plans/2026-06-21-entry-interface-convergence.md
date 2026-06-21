# 入口与接口收敛 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在不引入 Config/File/Table 全量依赖注入的前提下，统一 Ribbon 入口分层、收敛用户可见通知出口，并明确 Abstractions 接口的落地状态，消除「看起来可注入、实际全 static」的误导。

**Architecture:** `ThisAddIn` 退化为 COM/Ribbon 薄壳，所有 Ribbon 按钮回调经 `RibbonController` 转发到既有 Orchestrator/Service；通知统一经 `INotificationService`（生产默认 `MessageBoxNotificationService`）；`IConfigService` 等 static 接口仅加文档标注，Phase 2 再落地 Adapter。

**Tech Stack:** .NET Framework 4.8 COM 加载项、WinForms Ribbon、xUnit 源码约束测试、`dotnet test`

## Global Constraints

- 遵循 `AGENTS.md`：含中文文件用 `StrReplace`/`Write` 精确补丁，禁止 PowerShell 整文件读写。
- 最小 diff：不拆 `TableService`/`ConfigService`，不新增 DI 容器，不改 Ribbon.xml，不改变用户可见行为与文案。
- TDD：每 Task 先写/改 failing 源码约束测试 → 实现 → `dotnet test` 全绿。
- COM 自动化入口（`Automation_*`）保留在 `ThisAddIn`，仅转发到与 Ribbon 相同的 Orchestrator 方法。

---

## 范围外（本计划 deliberately 不做）

| 项 | 原因 |
|----|------|
| `ConfigService` → `IConfigService` Adapter + 注入 | 触达 `InsertPhotosForm`/`ProgressService` 全链，改动面过大 |
| `FileService`/`TableService`/`TableNumberingService` 接口化 | Phase 2 god-module-split 范畴 |
| 刷新编号 COM 自动化入口 | 无 E2E 需求，非入口唯一性问题 |
| `InsertPhotosOrchestrator` 自动化 vs UI 的 `Execute`/`ExecuteDeferred` 分支合并 | 行为变更风险 |
| 重构 `ProgressForm`/`FailureDetailsForm` 内 MessageBox | 属窗体内部交互，非服务层通知 |

---

## 目标结构（完成后）

```
ThisAddIn
  ├─ Ribbon_Load / Get*Pressed          → RibbonController
  ├─ OnInsertPhotosClick                → RibbonController → InsertPhotosOrchestrator
  ├─ OnRefreshNumberingClick            → RibbonController → NumberingRefreshService
  ├─ OnAbout / OnToggle* / OnShow*      → RibbonController（已有）
  ├─ Automation_ShowInsertPhotosForm    → InsertPhotosOrchestrator（不变，Gate 仍在 ThisAddIn 或下沉到 Orchestrator 入口）
  └─ Automation_ExecuteFromConfig       → InsertPhotosOrchestrator（不变）

InsertPhotosOrchestrator / NumberingRefreshService / RibbonController
  └─ 用户可见错误与结果提示 → INotificationService（无裸 MessageBox.Show）
```

---

## Task 1：Ribbon 入口统一（优先级 1）

**Files:**
- Modify: `WordTools/ThisAddIn.cs`
- Modify: `WordTools/RibbonController.cs`
- Modify: `WordTools.Tests/Tests.cs`

### Step 1: RED — 约束 ThisAddIn 不再直接 new Orchestrator/Service

- [ ] 在 `Tests.cs` 新增 `ThisAddIn_ribbon_callbacks_delegate_to_ribbon_controller`：
  - 读 `ThisAddIn.cs` 源码
  - `Assert.Contains("InsertPhotosOrchestrator", ribbonControllerSource)` **或** assert ThisAddIn 调 `_ribbonController.OnInsertPhotosClick`
  - `Assert.DoesNotContain("InsertPhotosOrchestrator", addInSource)`（Automation 方法除外，允许 Automation 直接调 Orchestrator）
  - `Assert.DoesNotContain("NumberingRefreshService", addInSource)`
  - `Assert.DoesNotContain("new InsertPhotosOrchestrator", addInSource)`（Automation 块除外：可用正则或分段断言 Automation 区域仍含 Orchestrator）

- [ ] 新增 `Numbering_refresh_entry_owned_by_ribbon_controller`：
  - `RibbonController.cs` 含 `NumberingRefreshService` 与 `RefreshFromCurrentSelection`
  - `ThisAddIn.OnRefreshNumberingClick` 仅一行委托 `_ribbonController`

- [ ] 运行 `dotnet test --filter "ThisAddIn_ribbon|Numbering_refresh"` → 预期 FAIL

### Step 2: GREEN — 实现 RibbonController 转发

- [ ] `RibbonController` 新增：

```csharp
public void OnInsertPhotosClick(Word.Application application)
{
    new InsertPhotosOrchestrator(application).ShowFormAndExecuteIfConfirmed();
}

public void OnRefreshNumberingClick(Word.Application application)
{
    var appContext = new WordApplicationContext(application);
    var notification = new MessageBoxNotificationService();
    new NumberingRefreshService(appContext, notification).RefreshFromCurrentSelection();
}
```

- [ ] `ThisAddIn.OnInsertPhotosClick` → `_ribbonController.OnInsertPhotosClick(Globals.Application)`
- [ ] `ThisAddIn.OnRefreshNumberingClick` → `_ribbonController.OnRefreshNumberingClick(Globals.Application)`
- [ ] Automation 方法保持现状（或可选：抽 `RibbonController` 的 package-private 等价方法供 Automation 复用 — **本 Task 可不做**，避免 Automation 路径改动）

- [ ] `dotnet test` 全绿

### Step 3: Commit

- [ ] `git commit -m "refactor(ribbon): route insert and refresh through RibbonController"`

---

## Task 2：通知出口统一 — Orchestrator + RibbonController（优先级 2a）

**Files:**
- Modify: `WordTools/Services/InsertPhotosOrchestrator.cs`
- Modify: `WordTools/RibbonController.cs`
- Modify: `WordTools.Tests/Tests.cs`

### Step 1: RED

- [ ] 新增 `InsertPhotosOrchestrator_does_not_call_message_box_directly`：
  - 读 `InsertPhotosOrchestrator.cs`
  - `Assert.DoesNotContain("MessageBox.Show", source)`

- [ ] 新增 `RibbonController_does_not_call_message_box_directly`：
  - 读 `RibbonController.cs`
  - `Assert.DoesNotContain("MessageBox.Show", source)`

- [ ] 运行 → FAIL

### Step 2: GREEN — InsertPhotosOrchestrator

- [ ] 增加可选字段/构造：`INotificationService _notification`（默认 `new MessageBoxNotificationService()`）
- [ ] 将两处 `MessageBox.Show(..., Error)` 改为 `_notification.ShowError(...)`
- [ ] 生产路径 `CreateDefaultServices` 不变；错误路径走同一 `_notification` 实例

### Step 3: GREEN — RibbonController

- [ ] 构造函数注入或私有字段：`readonly INotificationService _notification = new MessageBoxNotificationService()`
- [ ] `OnShowLoggingSettingsSummary` / `OnAboutClick` 改用 `_notification.ShowInformation(...)`（Information 级别即可，行为与 MessageBox OK 一致）

- [ ] `dotnet test` 全绿

### Step 4: Commit

- [ ] `git commit -m "refactor(notify): route orchestrator and ribbon prompts via INotificationService"`

---

## Task 3：通知出口统一 — InsertPhotosForm（优先级 2b，可选但建议同 PR）

**Files:**
- Modify: `WordTools/Forms/InsertPhotosForm.cs`
- Modify: `WordTools/Services/InsertPhotosOrchestrator.cs`（构造 form 时传入 notification）
- Modify: `WordTools.Tests/Tests.cs`

### Step 1: RED

- [ ] 新增 `InsertPhotosForm_does_not_call_message_box_directly`（与 ProgressService 同类测试对齐）
- [ ] 运行 → FAIL

### Step 2: GREEN

- [ ] `InsertPhotosForm` 增加构造参数 `INotificationService notification = null`，默认 `new MessageBoxNotificationService()`
- [ ] 6 处 `MessageBox.Show` → `_notification.ShowWarning` / `ShowError`（按 icon 映射）
- [ ] `InsertPhotosOrchestrator.ShowFormAndExecuteIfConfirmed` 创建 form 时传入与 orchestrator 相同的 `_notification` 实例（保证 headless 测试可替换）

- [ ] `dotnet test` 全绿

### Step 3: Commit

- [ ] `git commit -m "refactor(notify): route InsertPhotosForm validation via INotificationService"`

---

## Task 4：Abstractions 状态标注（优先级 3）

**Files:**
- Modify: `WordTools/Services/Abstractions/IConfigService.cs`
- Modify: `WordTools/Services/Abstractions/IFileService.cs`
- Modify: `WordTools/Services/Abstractions/ITableService.cs`
- Modify: `WordTools/Services/Abstractions/ITableNumberingService.cs`
- Modify: `WordTools/Services/Abstractions/IBenchmarkLogService.cs`（若存在）
- Modify: `WordTools.Tests/Tests.cs`（可选：防回归测试）

### Step 1: 文档标注（无行为变更）

- [ ] 在每个**尚未落地**的接口 XML `<summary>` 追加 `<remarks>`：

```csharp
/// <remarks>
/// 规划抽象（Phase 2）。运行时仍使用 static <see cref="ConfigService"/>，尚无 Adapter 实现。
/// </remarks>
```

（`IFileService` → `FileService`，`ITableService` → `TableService`，以此类推）

- [ ] **已落地**的接口保持现状或补 remarks：
  - `INotificationService` — 已用于 ProgressService / NumberingRefreshService / Orchestrator
  - `IProgressReporter` / `IFailureDetailsPresenter` / `IWordApplicationContext` / `IDocumentContext`

### Step 2: RED（可选轻量测试）

- [ ] `AbstractionsExist` 测试旁新增 `Static_config_service_has_no_instance_adapter_yet`：
  - Glob 不存在 `ConfigServiceAdapter.cs` **或** assert 无 class implements IConfigService（反射扫描 `WordTools` 程序集，除测试外 implementors 数量为 0）

### Step 3: Commit

- [ ] `git commit -m "docs(abstractions): mark static-service interfaces as Phase 2 targets"`

---

## Task 5：验收与发布前检查

- [ ] `dotnet test` — WordTools.Tests 全绿
- [ ] Release MSBuild 成功（`WordTools.dll`）
- [ ] 手动冒烟（Word 64 位）：
  - Ribbon 批量插图 → 确认 → 插入成功
  - Ribbon 刷新编号 → 成功提示
  - 关于 / 日志设置 → 对话框正常
- [ ] automation smoke（若环境可用）：`Run-BatchInsertE2E-Smoke.ps1` 仍绿（Automation 路径未改行为）

**不 bump 版本**：纯 refactor，除非用户要求发 patch 包。

---

## 风险与回滚

| 风险 | 缓解 |
|------|------|
| ThisAddIn 测试对 Automation 块误判 | 测试用明确方法名过滤，Automation 区域允许 `InsertPhotosOrchestrator` |
| Form 构造签名变更影响 Designer | `InsertPhotosForm` 为纯代码 UI，无 Designer 依赖 |
| `ShowInformation` 替代 About 的 MessageBox | 图标从 Information 保持一致，用户无感 |

回滚：3 个 commit 可独立 revert（Task 1/2/3）。

---

## 预估工作量

| Task | 改动文件 | 预估 |
|------|----------|------|
| 1 Ribbon 统一 | 3 | ~30 min |
| 2a 通知 Orchestrator/Ribbon | 3 | ~25 min |
| 2b 通知 Form | 3 | ~35 min |
| 3 接口标注 | 5–6 | ~15 min |
| 5 验收 | — | ~20 min |
| **合计** | | **~2 h** |

---

## 后续（不在本计划，Phase 2 衔接）

1. `ConfigServiceAdapter : IConfigService` + `InsertPhotosForm` 注入（配合 god-module-split）
2. `NumberingRefreshService` 依赖 `ITableService`/`ITableNumberingService` 门面
3. 刷新编号 E2E / COM 自动化入口（若有回归需求）
