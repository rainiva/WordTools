# WordTools 拆分重构实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 TableService(2088行)、ProgressService(1742行) 等超大文件按单一职责拆分，提取接口抽象，补充单元测试，使每个模块权责清晰、入口唯一、接口一致。

**Architecture:** 分 6 个阶段递进：① 接口提取（基础层）→ ② TableService 拆分 → ③ ProgressService 拆分 → ④ ThisAddIn Ribbon 提取 → ⑤ FileService/Theme 清理 → ⑥ 测试覆盖补全。每阶段独立可交付、可构建、可测试。

**Tech Stack:** .NET Framework 4.8, xUnit 2.6.2, .NET 8.0 (测试), Microsoft.Office.Interop.Word PIA

## Global Constraints

- 目标框架 .NET Framework 4.8，不可升级
- 测试项目 .NET 8.0-windows + xUnit，通过 `<Compile Include="..">` 链接源文件
- 所有新增接口放在 `Services/Abstractions/`，适配器放在 `Services/Adapters/`
- 新增文件必须在 `WordTools.csproj` 和 `WordTools.Tests.csproj` 中注册
- 中文注释保留，不得破坏编码（UTF-8）
- 每次提交后必须 `msbuild WordTools.sln /p:Configuration=Release` 构建通过
- 每次提交后必须 `dotnet test` 测试通过

---

## 阶段总览

| 阶段 | 目标 | 涉及文件 | 预估任务数 |
|------|------|----------|-----------|
| 1 | 接口提取（基础层） | 5 个新接口 + 5 个适配器 | 5 |
| 2 | TableService 拆分 | TableService → TableService + TableNumberingService | 3 |
| 3 | ProgressService 拆分 | ProgressService → 6 个关注点模块 | 6 |
| 4 | ThisAddIn Ribbon 提取 | ThisAddIn → ThisAddIn + RibbonController | 2 |
| 5 | FileService/Theme 清理 | FileService UI 分离 + Theme 拆分 | 2 |
| 6 | 测试覆盖补全 | 为拆分后模块补充单元测试 | 4 |

---

## 阶段 1：接口提取（基础层）

> 为现有静态服务类提取接口，为后续拆分和测试打基础。本阶段不改变任何行为，只增加接口层。

### Task 1.1：提取 ITableService 接口

**Files:**
- Create: `WordTools/Services/Abstractions/ITableService.cs`
- Modify: `WordTools/WordTools.csproj`
- Modify: `WordTools.Tests/WordTools.Tests.csproj`

**Interfaces:**
- Produces: `ITableService` — 表格操作抽象接口

- [ ] **Step 1: 定义接口**

```csharp
// WordTools/Services/Abstractions/ITableService.cs
using System;
using Microsoft.Office.Interop.Word;

namespace WordTools.Services.Abstractions
{
    public interface ITableService
    {
        bool IsSelectionInTable(Selection selection);
        bool IsSelectionInFirstColumn(Selection selection);
        Table GetCurrentTable(Selection selection);
        bool IsCellSuitableForImage(Table tbl, int row, int col);
        int FindNextSuitableCell(Table tbl, int startRow, int startCol, out int resultRow, out int resultCol);
        bool HasFloatingShapeInCell(Cell cell);
        bool IsMergedCell(Table tbl, int row, int col);
        int GetMergedRowSpan(Table tbl, int row, int col);
        void EnsureRowExists(Table tbl, int rowIndex);
        void AdjustTableColumns(Table tbl, int targetCols);
        bool IsTableFixedColumnWidth(Table tbl);
        void SetTableFixedColumnWidth(Table tbl);
        void CreateTitleRow(Table tbl, string title);
        void InsertDescriptionRow(Table tbl, int imageRow, string description);
        void FillEmptyCellsWithNA(Table tbl);
    }
}
```

- [ ] **Step 2: 注册到 csproj**

在 `WordTools.csproj` 的 `<ItemGroup>` 中添加：
```xml
<Compile Include="Services\Abstractions\ITableService.cs" />
```

在 `WordTools.Tests.csproj` 的 `<ItemGroup>` 中添加：
```xml
<Compile Include="..\WordTools\Services\Abstractions\ITableService.cs" Link="Services\Abstractions\ITableService.cs" />
```

- [ ] **Step 3: 构建验证**

```bash
msbuild WordTools.sln /p:Configuration=Release
```
Expected: Build succeeded, 0 errors.

- [ ] **Step 4: 提交**

```bash
git add WordTools/Services/Abstractions/ITableService.cs WordTools/WordTools.csproj WordTools.Tests/WordTools.Tests.csproj
git commit -m "refactor: add ITableService interface"
```

---

### Task 1.2：提取 ITableNumberingService 接口

**Files:**
- Create: `WordTools/Services/Abstractions/ITableNumberingService.cs`
- Modify: `WordTools/WordTools.csproj`
- Modify: `WordTools.Tests/WordTools.Tests.csproj`

**Interfaces:**
- Produces: `ITableNumberingService` — 表格编号抽象接口

- [ ] **Step 1: 定义接口**

```csharp
// WordTools/Services/Abstractions/ITableNumberingService.cs
using System;
using Microsoft.Office.Interop.Word;

namespace WordTools.Services.Abstractions
{
    public interface ITableNumberingService
    {
        void RefreshTableNumbering(Table tbl, Document doc, int alignment = 2,
            Action<string> progressCallback = null);
        void ClearTableNumbering(Table tbl);
        void AddNumberingToDescriptionRows(Table tbl, int numberPosition, int alignment);
        bool HasNumbering(Table tbl);
    }
}
```

- [ ] **Step 2: 注册到 csproj**

在 `WordTools.csproj` 中添加：
```xml
<Compile Include="Services\Abstractions\ITableNumberingService.cs" />
```

在 `WordTools.Tests.csproj` 中添加：
```xml
<Compile Include="..\WordTools\Services\Abstractions\ITableNumberingService.cs" Link="Services\Abstractions\ITableNumberingService.cs" />
```

- [ ] **Step 3: 构建验证**

```bash
msbuild WordTools.sln /p:Configuration=Release
```

- [ ] **Step 4: 提交**

```bash
git add WordTools/Services/Abstractions/ITableNumberingService.cs WordTools/WordTools.csproj WordTools.Tests/WordTools.Tests.csproj
git commit -m "refactor: add ITableNumberingService interface"
```

---

### Task 1.3：提取 IConfigService 接口

**Files:**
- Create: `WordTools/Services/Abstractions/IConfigService.cs`
- Modify: `WordTools/WordTools.csproj`
- Modify: `WordTools.Tests/WordTools.Tests.csproj`

**Interfaces:**
- Produces: `IConfigService` — 配置管理抽象接口

- [ ] **Step 1: 定义接口**

```csharp
// WordTools/Services/Abstractions/IConfigService.cs
using Microsoft.Office.Interop.Word;

namespace WordTools.Services.Abstractions
{
    public interface IConfigService
    {
        string GetDocumentProperty(Document doc, string propertyName, string defaultValue = "");
        void SetDocumentProperty(Document doc, string propertyName, string value);
        string GetLastFolderPath();
        void SaveLastFolderPath(string path);
        float GetLastImageHeightCM();
        void SaveLastImageHeightCM(float height);
        bool GetDetailedLoggingEnabled();
        void SaveDetailedLoggingEnabled(bool enabled);
        bool GetBenchmarkLoggingEnabled();
        void SaveBenchmarkLoggingEnabled(bool enabled);
    }
}
```

- [ ] **Step 2: 注册到 csproj**

在 `WordTools.csproj` 中添加：
```xml
<Compile Include="Services\Abstractions\IConfigService.cs" />
```

在 `WordTools.Tests.csproj` 中添加：
```xml
<Compile Include="..\WordTools\Services\Abstractions\IConfigService.cs" Link="Services\Abstractions\IConfigService.cs" />
```

- [ ] **Step 3: 构建验证**

```bash
msbuild WordTools.sln /p:Configuration=Release
```

- [ ] **Step 4: 提交**

```bash
git add WordTools/Services/Abstractions/IConfigService.cs WordTools/WordTools.csproj WordTools.Tests/WordTools.Tests.csproj
git commit -m "refactor: add IConfigService interface"
```

---

### Task 1.4：提取 IFileService 接口

**Files:**
- Create: `WordTools/Services/Abstractions/IFileService.cs`
- Modify: `WordTools/WordTools.csproj`
- Modify: `WordTools.Tests/WordTools.Tests.csproj`

**Interfaces:**
- Produces: `IFileService` — 文件操作抽象接口

- [ ] **Step 1: 定义接口**

```csharp
// WordTools/Services/Abstractions/IFileService.cs
namespace WordTools.Services.Abstractions
{
    public interface IFileService
    {
        string[] GetImageFilesFromFolder(string folderPath, bool includeSubFolders);
        int CountTotalImageFiles(string folderPath, bool includeRoot, bool includeSubFolders);
        string[] GetSupportedExtensions();
        bool IsSupportedImageFile(string filePath);
    }
}
```

- [ ] **Step 2: 注册到 csproj**

在 `WordTools.csproj` 中添加：
```xml
<Compile Include="Services\Abstractions\IFileService.cs" />
```

在 `WordTools.Tests.csproj` 中添加：
```xml
<Compile Include="..\WordTools\Services\Abstractions\IFileService.cs" Link="Services\Abstractions\IFileService.cs" />
```

- [ ] **Step 3: 构建验证**

```bash
msbuild WordTools.sln /p:Configuration=Release
```

- [ ] **Step 4: 提交**

```bash
git add WordTools/Services/Abstractions/IFileService.cs WordTools/WordTools.csproj WordTools.Tests/WordTools.Tests.csproj
git commit -m "refactor: add IFileService interface"
```

---

### Task 1.5：提取 IBenchmarkLogService 接口

**Files:**
- Create: `WordTools/Services/Abstractions/IBenchmarkLogService.cs`
- Modify: `WordTools/WordTools.csproj`
- Modify: `WordTools.Tests/WordTools.Tests.csproj`

**Interfaces:**
- Produces: `IBenchmarkLogService` — 基准日志抽象接口

- [ ] **Step 1: 定义接口**

```csharp
// WordTools/Services/Abstractions/IBenchmarkLogService.cs
namespace WordTools.Services.Abstractions
{
    public interface IBenchmarkLogService
    {
        string GetDefaultLogPath(string documentPath);
        void AppendCsv(string filePath, BenchmarkLogEntry entry);
    }
}
```

- [ ] **Step 2: 注册到 csproj**

在 `WordTools.csproj` 中添加：
```xml
<Compile Include="Services\Abstractions\IBenchmarkLogService.cs" />
```

在 `WordTools.Tests.csproj` 中添加：
```xml
<Compile Include="..\WordTools\Services\Abstractions\IBenchmarkLogService.cs" Link="Services\Abstractions\IBenchmarkLogService.cs" />
```

- [ ] **Step 3: 构建验证**

```bash
msbuild WordTools.sln /p:Configuration=Release
```

- [ ] **Step 4: 提交**

```bash
git add WordTools/Services/Abstractions/IBenchmarkLogService.cs WordTools/WordTools.csproj WordTools.Tests/WordTools.Tests.csproj
git commit -m "refactor: add IBenchmarkLogService interface"
```

---

## 阶段 2：TableService 拆分

> 将 TableService.cs (2088行) 拆分为两个文件：TableService.cs（表格结构操作）和 TableNumberingService.cs（编号管理）。

### 方法归属映射

| 方法 | 归属 | 行号范围 |
|------|------|---------|
| `IsSelectionInTable` | TableService | :26 |
| `IsSelectionInFirstColumn` | TableService | :35 |
| `GetCurrentTable` | TableService | :44 |
| `IsCellSuitableForImage` | TableService | :53 |
| `GetCellAvailability` (4个重载) | TableService | :58-403 |
| `FindNextSuitableCell` | TableService | :175 |
| `HasFloatingShapeInCell` | TableService | :268 |
| `BuildFloatingShapeIndex` | TableService | :296 |
| `IsMergedCell` | TableService | :331 |
| `GetMergedRowSpan` | TableService | :376 |
| `GetImageRowAvailability` (2个重载) | TableService | :427-432 |
| `FindNextSuitableImageRow` (2个重载) | TableService | :452-457 |
| `AddMergedRow` | TableService | :491 |
| `ShouldTreatCellCountMismatchAsMerged` | TableService | :504 |
| `GetImageRowSearchEndRow` | TableService | :509 |
| `EnumerateFallbackRows` | TableService | :514 |
| `GetCellAvailabilityCore` | TableService | :535 |
| `EnsureRowExists` (2个重载) | TableService | :568-595 |
| `AdjustTableColumns` | TableService | :624 |
| `IsTableFixedColumnWidth` | TableService | :658 |
| `SetTableFixedColumnWidth` | TableService | :675 |
| `CreateTitleRow` | TableService | :695 |
| `InsertDescriptionRow` | TableService | :753 |
| `InsertFileNameDescriptionRow` | TableService | :768 |
| `FillEmptyCellsWithNA` | TableService | :809 |
| **RefreshTableNumbering** | **TableNumberingService** | :840 |
| **ClearTableNumbering** | **TableNumberingService** | :1448 |
| **AddNumberingToDescriptionRows** | **TableNumberingService** | :1689 |
| **CalculateNextSequenceNumber** | **TableNumberingService** | :1807 |
| **ExtractNumberFromCell** | **TableNumberingService** | :1844 |
| **HasNumbering** | **TableNumberingService** | :1880 |
| **UpdateCellNumber** | **TableNumberingService** | :1903 |
| **SetCellNumber** | **TableNumberingService** | :1924 |
| **ExtractSeqNumberFromCell** | **TableNumberingService** | :1965 |
| **InsertNumberText** | **TableNumberingService** | :1980 |
| **InsertSeqField** | **TableNumberingService** | :2044 |
| **ExtractNumberFromCellText** | **TableNumberingService** | :2055 |
| **CleanCellText** | **TableNumberingService** | :2078 |

### Task 2.1：创建 TableNumberingService.cs

**Files:**
- Create: `WordTools/Services/TableNumberingService.cs`
- Modify: `WordTools/Services/TableService.cs`（删除编号相关方法）
- Modify: `WordTools/WordTools.csproj`
- Modify: `WordTools/ThisAddIn.cs`（更新调用点）

**Interfaces:**
- Consumes: `ITableNumberingService` (Task 1.2)
- Produces: `TableNumberingService` 静态类

- [ ] **Step 1: 创建 TableNumberingService.cs**

从 `TableService.cs` 的行 834-2088（`#region 自动编号` 到文件末尾之前）提取所有编号相关方法到新文件：

```csharp
// WordTools/Services/TableNumberingService.cs
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace WordTools.Services
{
    public static class TableNumberingService
    {
        private static void SafeIgnore(Exception ex, string context)
        {
            Debug.WriteLine($"{context}: {ex.Message}");
        }

        // 从 TableService.cs 移入以下方法（保持原始实现不变）：
        // - RefreshTableNumbering (行840-1447)
        // - ClearTableNumbering (行1448-1688)
        // - AddNumberingToDescriptionRows (行1689-1806)
        // - CalculateNextSequenceNumber (行1807-1843)
        // - ExtractNumberFromCell (行1844-1879)
        // - HasNumbering (行1880-1902)
        // - UpdateCellNumber (行1903-1923)
        // - SetCellNumber (行1924-1964)
        // - ExtractSeqNumberFromCell (行1965-1979)
        // - InsertNumberText (行1980-2043)
        // - InsertSeqField (行2044-2054)
        // - ExtractNumberFromCellText (行2055-2077)
        // - CleanCellText (行2078-2088)
    }
}
```

**操作方式：** 复制 TableService.cs 行 834-2088 的全部内容到新文件，替换类名为 `TableNumberingService`，删除 `#region 自动编号` 和 `#endregion` 标记。

- [ ] **Step 2: 从 TableService.cs 删除编号方法**

删除 TableService.cs 中行 832（`#endregion` 之后）到行 2088（文件末尾 `}` 之前）的所有内容。保留 `SafeIgnore` 方法（两个类各自保留一份）。

- [ ] **Step 3: 更新 ThisAddIn.cs 调用点**

将 `ThisAddIn.cs` 行 239 的调用：
```csharp
Services.TableService.RefreshTableNumbering(tbl, doc, 2, (status) =>
```
改为：
```csharp
Services.TableNumberingService.RefreshTableNumbering(tbl, doc, 2, (status) =>
```

- [ ] **Step 4: 注册到 csproj**

在 `WordTools.csproj` 的 `<ItemGroup>` 中添加：
```xml
<Compile Include="Services\TableNumberingService.cs" />
```

- [ ] **Step 5: 构建验证**

```bash
msbuild WordTools.sln /p:Configuration=Release
```
Expected: Build succeeded, 0 errors.

- [ ] **Step 6: 测试验证**

```bash
dotnet test
```
Expected: All tests pass.

- [ ] **Step 7: 提交**

```bash
git add WordTools/Services/TableNumberingService.cs WordTools/Services/TableService.cs WordTools/ThisAddIn.cs WordTools/WordTools.csproj
git commit -m "refactor: split TableService into TableService + TableNumberingService"
```

---

### Task 2.2：为 TableNumberingService 添加特征测试

**Files:**
- Create: `WordTools.Tests/TableNumberingServiceTests.cs`
- Modify: `WordTools.Tests/WordTools.Tests.csproj`

**Interfaces:**
- Consumes: `TableNumberingService` (Task 2.1)

- [ ] **Step 1: 编写纯逻辑测试**

```csharp
// WordTools.Tests/TableNumberingServiceTests.cs
using System;
using System.Text.RegularExpressions;
using Xunit;

namespace WordTools.Tests
{
    public class TableNumberingServiceTests
    {
        [Theory]
        [InlineData("图 1-1", "1-1")]
        [InlineData("图1", "1")]
        [InlineData("1-2-3", "1-2-3")]
        [InlineData("abc", null)]
        [InlineData("", null)]
        public void ExtractNumberFromCellText_ReturnsCorrectNumber(string input, string expected)
        {
            // 测试编号提取的纯逻辑部分
            // 注意：ExtractNumberFromCellText 是 private，需要通过反射或改为 internal
            // 替代方案：测试 CleanCellText 等可公开访问的辅助方法
        }

        [Theory]
        [InlineData("  图 1-1  ", "图 1-1")]
        [InlineData("图\t1", "图 1")]
        [InlineData("", "")]
        public void CleanCellText_RemovesExtraWhitespace(string input, string expected)
        {
            // 测试文本清理的纯逻辑部分
        }
    }
}
```

**注意：** 由于 `CleanCellText` 和 `ExtractNumberFromCellText` 是 private 方法，需要将其改为 `internal` 并在测试项目中使用 `InternalsVisibleTo`，或者通过公开的 `HasNumbering` 等方法间接测试。具体实现时根据可访问性调整测试策略。

- [ ] **Step 2: 注册到测试 csproj**

在 `WordTools.Tests.csproj` 中添加：
```xml
<Compile Include="..\WordTools\Services\TableNumberingService.cs" Link="Services\TableNumberingService.cs" />
```

- [ ] **Step 3: 运行测试**

```bash
dotnet test
```
Expected: All tests pass.

- [ ] **Step 4: 提交**

```bash
git add WordTools.Tests/TableNumberingServiceTests.cs WordTools.Tests/WordTools.Tests.csproj
git commit -m "test: add TableNumberingService characterization tests"
```

---

### Task 2.3：验证 TableService 拆分完整性

- [ ] **Step 1: 确认 TableService.cs 行数**

```powershell
$content = [System.IO.File]::ReadAllText('WordTools\Services\TableService.cs')
$lines = ($content -split "`r`n|`n").Count
Write-Host "TableService.cs: $lines lines"
```
Expected: 行数 ≤ 850（原 2088 行减去编号部分约 1250 行）

- [ ] **Step 2: 确认 TableNumberingService.cs 行数**

```powershell
$content = [System.IO.File]::ReadAllText('WordTools\Services\TableNumberingService.cs')
$lines = ($content -split "`r`n|`n").Count
Write-Host "TableNumberingService.cs: $lines lines"
```
Expected: 行数 ≈ 1250

- [ ] **Step 3: 确认无遗漏方法**

检查 TableService.cs 中不再包含任何 `Numbering`、`Seq`、`Number` 相关方法名。

- [ ] **Step 4: 全量构建 + 测试**

```bash
msbuild WordTools.sln /p:Configuration=Release
dotnet test
```

- [ ] **Step 5: 提交（如有修正）**

```bash
git add -A
git commit -m "refactor: verify TableService split completeness"
```

---

## 阶段 3：ProgressService 拆分

> 将 ProgressService.cs (1742行) 按 6 个关注点拆分为独立模块。

### 关注点映射

| 关注点 | 新类名 | 方法 | 预估行数 |
|--------|--------|------|---------|
| 高性能模式管理 | `HighPerformanceModeController` | `EnterHighPerformanceMode`, `ExitHighPerformanceMode`, `GetOptimizedRefreshInterval`, `GetStatusBarUpdateInterval` | ~100 |
| ESC 键监控 | `EscapeKeyMonitor` | `CheckEscapeKey`, `ShouldCancel`, P/Invoke `GetAsyncKeyState` | ~50 |
| 窗口激活控制 | `WindowActivationService` | `EnsureWindowTopMost`, `EnsureWordWindowActive`, P/Invoke `SetWindowPos`/`IsWindow` | ~60 |
| 插入错误分类 | `InsertionErrorClassifier` | `ClassifyInsertionError`, `IsMergedCellError` | ~120 |
| 插入摘要展示 | `InsertionResultPresenter` | `ShowInsertionSummary`, `BuildTimeDetail`, `ShowFailureSummary`, `ShowMergedCellWarning`, `ShowOverwriteWarning`, `TryWriteBenchmarkLog` | ~300 |
| 批量插入编排 | `ProgressService`（保留） | `InsertPhotosWithProgress`, `InsertSelectedPhotosWithProgress`, `ProcessFileBatch`, `CloseProgressForm` | ~800 |

### Task 3.1：提取 HighPerformanceModeController

**Files:**
- Create: `WordTools/Services/HighPerformanceModeController.cs`
- Modify: `WordTools/Services/ProgressService.cs`
- Modify: `WordTools/WordTools.csproj`
- Modify: `WordTools.Tests/WordTools.Tests.csproj`

**Interfaces:**
- Consumes: `IWordApplicationContext`
- Produces: `HighPerformanceModeController`

- [ ] **Step 1: 创建 HighPerformanceModeController.cs**

```csharp
// WordTools/Services/HighPerformanceModeController.cs
using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using WordTools.Services.Abstractions;

namespace WordTools.Services
{
    public sealed class HighPerformanceModeController
    {
        private readonly IWordApplicationContext _appContext;
        private bool _originalScreenUpdating;
        private bool _originalDisplayAlerts;
        private bool _highPerformanceModeEntered;

        public HighPerformanceModeController(IWordApplicationContext appContext)
        {
            _appContext = appContext ?? throw new ArgumentNullException(nameof(appContext));
        }

        public void Enter()
        {
            // 从 ProgressService.EnterHighPerformanceMode() 移入（行113-142）
        }

        public void Exit()
        {
            // 从 ProgressService.ExitHighPerformanceMode() 移入（行148-167）
        }

        public int GetOptimizedRefreshInterval(int totalFiles)
        {
            // 从 ProgressService.GetOptimizedRefreshInterval() 移入（行172-177）
        }

        public int GetStatusBarUpdateInterval(int totalFiles)
        {
            // 从 ProgressService.GetStatusBarUpdateInterval() 移入（行183-189）
        }
    }
}
```

- [ ] **Step 2: 从 ProgressService.cs 删除已提取方法**

删除 ProgressService.cs 中以下方法和字段：
- `EnterHighPerformanceMode()` (行113-142)
- `ExitHighPerformanceMode()` (行148-167)
- `GetOptimizedRefreshInterval()` (行172-177)
- `GetStatusBarUpdateInterval()` (行183-189)
- 字段 `_originalScreenUpdating`, `_originalDisplayAlerts`, `_highPerformanceModeEntered` (行53-55)

在 ProgressService 构造函数中初始化 `_perfController = new HighPerformanceModeController(appContext)`，并将所有调用点改为 `_perfController.Enter()` / `_perfController.Exit()` 等。

- [ ] **Step 3: 注册到 csproj**

在 `WordTools.csproj` 中添加：
```xml
<Compile Include="Services\HighPerformanceModeController.cs" />
```

在 `WordTools.Tests.csproj` 中添加：
```xml
<Compile Include="..\WordTools\Services\HighPerformanceModeController.cs" Link="Services\HighPerformanceModeController.cs" />
```

- [ ] **Step 4: 构建验证**

```bash
msbuild WordTools.sln /p:Configuration=Release
dotnet test
```

- [ ] **Step 5: 提交**

```bash
git add WordTools/Services/HighPerformanceModeController.cs WordTools/Services/ProgressService.cs WordTools/WordTools.csproj WordTools.Tests/WordTools.Tests.csproj
git commit -m "refactor: extract HighPerformanceModeController from ProgressService"
```

---

### Task 3.2：提取 EscapeKeyMonitor

**Files:**
- Create: `WordTools/Services/EscapeKeyMonitor.cs`
- Modify: `WordTools/Services/ProgressService.cs`
- Modify: `WordTools/WordTools.csproj`

**Interfaces:**
- Consumes: `IWordApplicationContext`
- Produces: `EscapeKeyMonitor`

- [ ] **Step 1: 创建 EscapeKeyMonitor.cs**

```csharp
// WordTools/Services/EscapeKeyMonitor.cs
using System.Runtime.InteropServices;
using WordTools.Services.Abstractions;

namespace WordTools.Services
{
    public sealed class EscapeKeyMonitor
    {
        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);
        private const int VK_ESCAPE = 0x1B;

        private readonly IWordApplicationContext _appContext;
        private bool _isCancelled;

        public EscapeKeyMonitor(IWordApplicationContext appContext)
        {
            _appContext = appContext;
        }

        public bool IsCancelled => _isCancelled;

        public bool ShouldCancel(IProgressReporter progressReporter)
        {
            // 从 ProgressService.ShouldCancel() 移入（行84-104）
            if (_isCancelled) return true;
            if (progressReporter?.IsCancelled == true)
            {
                _isCancelled = true;
                return true;
            }
            if ((GetAsyncKeyState(VK_ESCAPE) & 0x8000) != 0)
            {
                _isCancelled = true;
                _appContext.Application.StatusBar = "检测到 ESC 键，正在取消操作...";
                _appContext.DoEvents();
                return true;
            }
            return false;
        }

        public void Reset()
        {
            _isCancelled = false;
        }
    }
}
```

- [ ] **Step 2: 从 ProgressService.cs 删除已提取方法**

删除：
- `CheckEscapeKey()` (行76-79)
- `ShouldCancel()` (行84-104)
- P/Invoke `GetAsyncKeyState` 声明和 `VK_ESCAPE` 常量 (行21-23)
- 字段 `_isCancelled` (行41)

在 ProgressService 中用 `_escapeMonitor.ShouldCancel(_progressReporter)` 替换所有 `ShouldCancel()` 调用。

- [ ] **Step 3: 注册到 csproj**

在 `WordTools.csproj` 中添加：
```xml
<Compile Include="Services\EscapeKeyMonitor.cs" />
```

- [ ] **Step 4: 构建验证**

```bash
msbuild WordTools.sln /p:Configuration=Release
dotnet test
```

- [ ] **Step 5: 提交**

```bash
git add WordTools/Services/EscapeKeyMonitor.cs WordTools/Services/ProgressService.cs WordTools/WordTools.csproj
git commit -m "refactor: extract EscapeKeyMonitor from ProgressService"
```

---

### Task 3.3：提取 WindowActivationService

**Files:**
- Create: `WordTools/Services/WindowActivationService.cs`
- Modify: `WordTools/Services/ProgressService.cs`
- Modify: `WordTools/WordTools.csproj`

**Interfaces:**
- Produces: `WindowActivationService`

- [ ] **Step 1: 创建 WindowActivationService.cs**

```csharp
// WordTools/Services/WindowActivationService.cs
using System;
using System.Runtime.InteropServices;

namespace WordTools.Services
{
    public static class WindowActivationService
    {
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_SHOWWINDOW = 0x0040;

        public static void EnsureWindowTopMost(IntPtr handle)
        {
            // 从 ProgressService.EnsureWindowTopMost() 移入（行1248-1262）
        }

        public static void EnsureWordWindowActive(IntPtr wordHandle)
        {
            // 从 ProgressService.EnsureWordWindowActive() 移入（行1263-1281）
        }
    }
}
```

- [ ] **Step 2: 从 ProgressService.cs 删除已提取方法**

删除：
- `EnsureWindowTopMost()` (行1248-1262)
- `EnsureWordWindowActive()` (行1263-1281)
- P/Invoke `SetWindowPos`, `IsWindow` 声明 (行25-29)
- 常量 `HWND_TOPMOST`, `HWND_NOTOPMOST`, `SWP_NOSIZE`, `SWP_NOMOVE`, `SWP_SHOWWINDOW` (行31-35)

将 ProgressService 中的调用改为 `WindowActivationService.EnsureWindowTopMost(handle)`。

- [ ] **Step 3: 注册到 csproj**

在 `WordTools.csproj` 中添加：
```xml
<Compile Include="Services\WindowActivationService.cs" />
```

- [ ] **Step 4: 构建验证**

```bash
msbuild WordTools.sln /p:Configuration=Release
dotnet test
```

- [ ] **Step 5: 提交**

```bash
git add WordTools/Services/WindowActivationService.cs WordTools/Services/ProgressService.cs WordTools/WordTools.csproj
git commit -m "refactor: extract WindowActivationService from ProgressService"
```

---

### Task 3.4：提取 InsertionErrorClassifier

**Files:**
- Create: `WordTools/Services/InsertionErrorClassifier.cs`
- Modify: `WordTools/Services/ProgressService.cs`
- Modify: `WordTools/WordTools.csproj`
- Modify: `WordTools.Tests/WordTools.Tests.csproj`

**Interfaces:**
- Produces: `InsertionErrorClassifier`

- [ ] **Step 1: 创建 InsertionErrorClassifier.cs**

```csharp
// WordTools/Services/InsertionErrorClassifier.cs
using System;
using System.IO;

namespace WordTools.Services
{
    public static class InsertionErrorClassifier
    {
        public static string Classify(Exception ex)
        {
            // 从 ProgressService.ClassifyInsertionError() 移入（行1282-1380）
            if (ex == null) return "未知错误";
            string msg = ex.Message ?? "";

            if (msg.IndexOf("rejected", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("retry", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("busy", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("忙", StringComparison.OrdinalIgnoreCase) >= 0 ||
                ex.HResult == unchecked((int)0x80010001) ||
                ex.HResult == unchecked((int)0x8001010A))
                return "Word 正忙，请关闭其他对话框后重试";

            if (ex is FileNotFoundException ||
                msg.IndexOf("找不到", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("not found", StringComparison.OrdinalIgnoreCase) >= 0)
                return "文件不存在或已被移动";

            if (ex is IOException ||
                msg.IndexOf("进程无法访问", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("being used", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("占用", StringComparison.OrdinalIgnoreCase) >= 0)
                return "文件被其他程序占用";

            if (ex is UnauthorizedAccessException ||
                msg.IndexOf("拒绝访问", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("access denied", StringComparison.OrdinalIgnoreCase) >= 0)
                return "没有文件访问权限";

            if (IsMergedCellError(ex))
                return "合并单元格";

            return msg;
        }

        public static bool IsMergedCellError(Exception ex)
        {
            // 从 ProgressService.IsMergedCellError() 移入（行1381-1391）
            if (ex == null) return false;
            string msg = ex.Message ?? "";
            return msg.IndexOf("合并", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   msg.IndexOf("merge", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   msg.IndexOf("单元格索引异常", StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
```

- [ ] **Step 2: 从 ProgressService.cs 删除已提取方法**

删除 `ClassifyInsertionError()` (行1282-1380) 和 `IsMergedCellError()` (行1381-1391)。
将调用改为 `InsertionErrorClassifier.Classify(ex)` 和 `InsertionErrorClassifier.IsMergedCellError(ex)`。

- [ ] **Step 3: 注册到 csproj**

在 `WordTools.csproj` 中添加：
```xml
<Compile Include="Services\InsertionErrorClassifier.cs" />
```

在 `WordTools.Tests.csproj` 中添加：
```xml
<Compile Include="..\WordTools\Services\InsertionErrorClassifier.cs" Link="Services\InsertionErrorClassifier.cs" />
```

- [ ] **Step 4: 编写特征测试**

```csharp
// WordTools.Tests/InsertionErrorClassifierTests.cs
using System;
using System.IO;
using WordTools.Services;
using Xunit;

namespace WordTools.Tests
{
    public class InsertionErrorClassifierTests
    {
        [Fact]
        public void Classify_NullException_ReturnsUnknown()
        {
            Assert.Equal("未知错误", InsertionErrorClassifier.Classify(null));
        }

        [Fact]
        public void Classify_FileNotFound_ReturnsFileMissing()
        {
            var ex = new FileNotFoundException("找不到文件");
            Assert.Equal("文件不存在或已被移动", InsertionErrorClassifier.Classify(ex));
        }

        [Fact]
        public void Classify_IOException_ReturnsFileLocked()
        {
            var ex = new IOException("进程无法访问该文件");
            Assert.Equal("文件被其他程序占用", InsertionErrorClassifier.Classify(ex));
        }

        [Fact]
        public void Classify_UnauthorizedAccess_ReturnsAccessDenied()
        {
            var ex = new UnauthorizedAccessException("拒绝访问");
            Assert.Equal("没有文件访问权限", InsertionErrorClassifier.Classify(ex));
        }

        [Fact]
        public void IsMergedCellError_MergeKeyword_ReturnsTrue()
        {
            var ex = new Exception("合并单元格操作失败");
            Assert.True(InsertionErrorClassifier.IsMergedCellError(ex));
        }

        [Fact]
        public void IsMergedCellError_NormalError_ReturnsFalse()
        {
            var ex = new Exception("普通错误");
            Assert.False(InsertionErrorClassifier.IsMergedCellError(ex));
        }
    }
}
```

- [ ] **Step 5: 构建 + 测试**

```bash
msbuild WordTools.sln /p:Configuration=Release
dotnet test
```

- [ ] **Step 6: 提交**

```bash
git add WordTools/Services/InsertionErrorClassifier.cs WordTools/Services/ProgressService.cs WordTools/WordTools.csproj WordTools.Tests/WordTools.Tests.csproj WordTools.Tests/InsertionErrorClassifierTests.cs
git commit -m "refactor: extract InsertionErrorClassifier with tests"
```

---

### Task 3.5：提取 InsertionResultPresenter

**Files:**
- Create: `WordTools/Services/InsertionResultPresenter.cs`
- Modify: `WordTools/Services/ProgressService.cs`
- Modify: `WordTools/WordTools.csproj`

**Interfaces:**
- Consumes: `INotificationService`, `IFailureDetailsPresenter`, `IBenchmarkLogService`
- Produces: `InsertionResultPresenter`

- [ ] **Step 1: 创建 InsertionResultPresenter.cs**

```csharp
// WordTools/Services/InsertionResultPresenter.cs
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using WordTools.Services.Abstractions;

namespace WordTools.Services
{
    public sealed class InsertionResultPresenter
    {
        private readonly INotificationService _notificationService;
        private readonly IFailureDetailsPresenter _failureDetailsPresenter;

        public InsertionResultPresenter(
            INotificationService notificationService,
            IFailureDetailsPresenter failureDetailsPresenter)
        {
            _notificationService = notificationService;
            _failureDetailsPresenter = failureDetailsPresenter;
        }

        public void ShowInsertionSummary(int successCount, int failCount,
            string timeInfo, string timeDetail,
            List<(string fileName, string errorReason)> failedFiles,
            List<int> mergedCellRows = null,
            List<string> overwriteWarnings = null)
        {
            // 从 ProgressService.ShowInsertionSummary() 移入（行1430-1480）
        }

        public string BuildTimeDetail(Stopwatch stopwatch, long t0, long t1,
            long t2, long t3, long t4, long t5, bool skippedClear)
        {
            // 从 ProgressService.BuildTimeDetail() 移入（行1481-1516）
        }

        public void ShowFailureSummary(
            List<(string fileName, string errorReason)> failedFiles)
        {
            // 从 ProgressService.ShowFailureSummary() 移入（行1517-1598）
        }

        public void ShowMergedCellWarning(List<int> mergedCellRows)
        {
            // 从 ProgressService.ShowMergedCellWarning() 移入（行1599-1639）
        }

        public void ShowOverwriteWarning(List<string> overwriteWarnings)
        {
            // 从 ProgressService.ShowOverwriteWarning() 移入（行1640-1683）
        }

        public void TryWriteBenchmarkLog(BenchmarkLogEntry entry, bool detailedLogging,
            bool benchmarkLogging, string documentPath)
        {
            // 从 ProgressService.TryWriteBenchmarkLog() 移入（行1684-1719）
        }
    }
}
```

- [ ] **Step 2: 从 ProgressService.cs 删除已提取方法**

删除行 1425-1731 的所有方法。在 ProgressService 构造函数中初始化 `_resultPresenter`，替换所有调用。

- [ ] **Step 3: 注册到 csproj**

在 `WordTools.csproj` 中添加：
```xml
<Compile Include="Services\InsertionResultPresenter.cs" />
```

- [ ] **Step 4: 构建验证**

```bash
msbuild WordTools.sln /p:Configuration=Release
dotnet test
```

- [ ] **Step 5: 提交**

```bash
git add WordTools/Services/InsertionResultPresenter.cs WordTools/Services/ProgressService.cs WordTools/WordTools.csproj
git commit -m "refactor: extract InsertionResultPresenter from ProgressService"
```

---

### Task 3.6：验证 ProgressService 拆分完整性

- [ ] **Step 1: 确认 ProgressService.cs 行数**

```powershell
$content = [System.IO.File]::ReadAllText('WordTools\Services\ProgressService.cs')
$lines = ($content -split "`r`n|`n").Count
Write-Host "ProgressService.cs: $lines lines"
```
Expected: 行数 ≤ 900（原 1742 行减去提取部分约 850 行）

- [ ] **Step 2: 确认 ProgressService 只保留编排职责**

ProgressService 应只包含：
- 构造函数（组装子组件）
- `InsertPhotosWithProgress()` — 文件夹批量插入编排
- `InsertSelectedPhotosWithProgress()` — 选中文件批量插入编排
- `ProcessFileBatch()` — 文件批次处理
- `CloseProgressForm()` — 关闭进度窗
- `SafeIgnore()` — 辅助方法
- `CleanupMemory()` — 内存清理（可保留或后续提取）
- `UpdateStatusBar()` — 状态栏更新（可保留或后续提取）

- [ ] **Step 3: 确认无 P/Invoke 残留**

ProgressService.cs 中不应再包含 `[DllImport]` 声明。

- [ ] **Step 4: 全量构建 + 测试**

```bash
msbuild WordTools.sln /p:Configuration=Release
dotnet test
```

- [ ] **Step 5: 提交（如有修正）**

```bash
git add -A
git commit -m "refactor: verify ProgressService split completeness"
```

---

## 阶段 4：ThisAddIn Ribbon 提取

> 将 ThisAddIn.cs (430行) 中的 Ribbon 回调方法提取到独立的 RibbonController 类。

### Task 4.1：创建 RibbonController.cs

**Files:**
- Create: `WordTools/RibbonController.cs`
- Modify: `WordTools/ThisAddIn.cs`
- Modify: `WordTools/WordTools.csproj`

**Interfaces:**
- Consumes: `ConfigService`, `LoggingOptionsStateController`, `AppVersionInfo`
- Produces: `RibbonController`

- [ ] **Step 1: 创建 RibbonController.cs**

```csharp
// WordTools/RibbonController.cs
using System;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using WordTools.Services;

namespace WordTools
{
    public class RibbonController
    {
        private Office.IRibbonUI _ribbonUI;

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbonUI = ribbonUI;
        }

        public void InvalidateRibbon()
        {
            _ribbonUI?.Invalidate();
        }

        public bool GetDetailedLoggingPressed(Office.IRibbonControl control)
        {
            return ConfigService.GetDetailedLoggingEnabled();
        }

        public bool GetBenchmarkLoggingPressed(Office.IRibbonControl control)
        {
            return LoggingOptionsStateController.Normalize(
                ConfigService.GetDetailedLoggingEnabled(),
                ConfigService.GetBenchmarkLoggingEnabled()).BenchmarkLoggingEnabled;
        }

        public bool GetBenchmarkLoggingEnabled(Office.IRibbonControl control)
        {
            return ConfigService.GetDetailedLoggingEnabled()
                && ConfigService.GetBenchmarkLoggingEnabled();
        }

        public void OnToggleDetailedLogging(Office.IRibbonControl control, bool pressed)
        {
            var state = LoggingOptionsStateController.Normalize(
                pressed, ConfigService.GetBenchmarkLoggingEnabled());
            ConfigService.SaveDetailedLoggingEnabled(state.DetailedLoggingEnabled);
            ConfigService.SaveBenchmarkLoggingEnabled(state.BenchmarkLoggingEnabled);
            InvalidateRibbon();
        }

        public void OnToggleBenchmarkLogging(Office.IRibbonControl control, bool pressed)
        {
            var state = LoggingOptionsStateController.Normalize(
                ConfigService.GetDetailedLoggingEnabled(), pressed);
            ConfigService.SaveDetailedLoggingEnabled(state.DetailedLoggingEnabled);
            ConfigService.SaveBenchmarkLoggingEnabled(state.BenchmarkLoggingEnabled);
            InvalidateRibbon();
        }

        public void OnShowLoggingSettingsSummary(Office.IRibbonControl control)
        {
            var state = LoggingOptionsStateController.Normalize(
                ConfigService.GetDetailedLoggingEnabled(),
                ConfigService.GetBenchmarkLoggingEnabled());

            string message =
                "当前日志设置：\n\n" +
                "详细日志：" + (state.DetailedLoggingEnabled ? "已开启" : "已关闭") + "\n" +
                "性能基准 CSV：" + (state.BenchmarkLoggingEnabled ? "已开启" : "已关闭") + "\n\n" +
                "提示：点击右侧下拉箭头可以直接调整这两项设置。";

            MessageBox.Show(message, "日志设置",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnAboutClick(Office.IRibbonControl control)
        {
            MessageBox.Show(AppVersionInfo.AboutMessage, "关于",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
```

- [ ] **Step 2: 修改 ThisAddIn.cs**

将以下方法从 ThisAddIn.cs 中删除，改为委托给 `_ribbonController`：
- `Ribbon_Load()` → `_ribbonController.OnRibbonLoad(ribbonUI)`
- `InvalidateRibbon()` → `_ribbonController.InvalidateRibbon()`
- `GetDetailedLoggingPressed()` → `_ribbonController.GetDetailedLoggingPressed(control)`
- `GetBenchmarkLoggingPressed()` → `_ribbonController.GetBenchmarkLoggingPressed(control)`
- `GetBenchmarkLoggingEnabled()` → `_ribbonController.GetBenchmarkLoggingEnabled(control)`
- `OnToggleDetailedLogging()` → `_ribbonController.OnToggleDetailedLogging(control, pressed)`
- `OnToggleBenchmarkLogging()` → `_ribbonController.OnToggleBenchmarkLogging(control, pressed)`
- `OnShowLoggingSettingsSummary()` → `_ribbonController.OnShowLoggingSettingsSummary(control)`
- `OnAboutClick()` → `_ribbonController.OnAboutClick(control)`

在 ThisAddIn 中添加字段：
```csharp
private readonly RibbonController _ribbonController = new RibbonController();
```

- [ ] **Step 3: 注册到 csproj**

在 `WordTools.csproj` 中添加：
```xml
<Compile Include="RibbonController.cs" />
```

- [ ] **Step 4: 构建验证**

```bash
msbuild WordTools.sln /p:Configuration=Release
dotnet test
```

- [ ] **Step 5: 提交**

```bash
git add WordTools/RibbonController.cs WordTools/ThisAddIn.cs WordTools/WordTools.csproj
git commit -m "refactor: extract RibbonController from ThisAddIn"
```

---

### Task 4.2：验证 ThisAddIn 瘦身效果

- [ ] **Step 1: 确认 ThisAddIn.cs 行数**

```powershell
$content = [System.IO.File]::ReadAllText('WordTools\ThisAddIn.cs')
$lines = ($content -split "`r`n|`n").Count
Write-Host "ThisAddIn.cs: $lines lines"
```
Expected: 行数 ≤ 250（原 430 行减去 Ribbon 回调约 180 行）

- [ ] **Step 2: 确认 ThisAddIn 只保留入口职责**

ThisAddIn 应只包含：
- COM 入口 (`OnConnection`, `OnDisconnection` 等)
- `GetCustomUI()` — Ribbon XML 加载
- `OnInsertPhotosClick()` — 委托给 `ShowInsertPhotosForm()`
- `OnRefreshNumberingClick()` — 委托给 `TableNumberingService`
- `OnExcelDataFillerClick()` — 创建 ExcelDataFillerForm
- `ShowInsertPhotosForm()` — 窗体创建
- `ExecuteInsertPhotosRequest()` — 服务组装
- `ExecuteInsertPhotosRequestDeferred()` — 延迟调度

- [ ] **Step 3: 全量构建 + 测试**

```bash
msbuild WordTools.sln /p:Configuration=Release
dotnet test
```

- [ ] **Step 4: 提交（如有修正）**

```bash
git add -A
git commit -m "refactor: verify ThisAddIn slimming"
```

---

## 阶段 5：FileService/Theme 清理

> 将 FileService 中的 UI 代码分离，拆分 Theme 为 Theme + UiToolkit。

### Task 5.1：FileService UI 分离

**Files:**
- Modify: `WordTools/Services/FileService.cs`
- Modify: `WordTools/Forms/InsertPhotosForm.cs`

**Interfaces:**
- 将 `SelectFolder()` 和 `SelectImageFiles()` 从 FileService 移至 InsertPhotosForm

- [ ] **Step 1: 识别 FileService 中的 UI 方法**

FileService 中包含以下 UI 方法：
- `SelectFolder()` — 使用 `FolderBrowserDialog`
- `SelectImageFiles()` — 使用 `OpenFileDialog`

这些方法应移至 `InsertPhotosForm.cs` 作为私有方法。

- [ ] **Step 2: 移动 UI 方法**

将 `SelectFolder()` 和 `SelectImageFiles()` 从 `FileService.cs` 移至 `InsertPhotosForm.cs`，改为实例方法。

更新 `InsertPhotosForm.cs` 中的调用点：
```csharp
// 原调用：FileService.SelectFolder(...)
// 新调用：SelectFolder(...)  (本类方法)
```

- [ ] **Step 3: 从 FileService 删除 UI 方法**

删除 `FileService.cs` 中的 `SelectFolder()` 和 `SelectImageFiles()` 方法。

- [ ] **Step 4: 构建验证**

```bash
msbuild WordTools.sln /p:Configuration=Release
dotnet test
```

- [ ] **Step 5: 提交**

```bash
git add WordTools/Services/FileService.cs WordTools/Forms/InsertPhotosForm.cs
git commit -m "refactor: move UI dialogs from FileService to InsertPhotosForm"
```

---

### Task 5.2：Theme 拆分为 Theme + UiToolkit

**Files:**
- Create: `WordTools/UiToolkit.cs`
- Modify: `WordTools/Theme.cs`
- Modify: `WordTools/WordTools.csproj`

**Interfaces:**
- Produces: `UiToolkit` — UI 控件工厂

- [ ] **Step 1: 创建 UiToolkit.cs**

```csharp
// WordTools/UiToolkit.cs
using System.Drawing;
using System.Windows.Forms;

namespace WordTools
{
    public static class UiToolkit
    {
        public static Button CreateButton(string text, ButtonStyle style = ButtonStyle.Default)
        {
            // 从 Theme.CreateButton() 移入（行276-312）
        }

        public static void ApplyButtonStyle(Button btn, ButtonStyle style)
        {
            // 从 Theme.ApplyButtonStyle() 移入（行319-340）
        }

        public static Label CreateDivider(int width)
        {
            // 从 Theme.CreateDivider() 移入（行347-355）
        }
    }
}
```

- [ ] **Step 2: 从 Theme.cs 删除已提取方法**

删除 `CreateButton()`、`ApplyButtonStyle()`、`CreateDivider()` 方法。
Theme.cs 只保留：`Colors` 类、`Fonts` 类、`DpiScale` 方法、`S()` 缩放方法。

- [ ] **Step 3: 更新所有调用点**

全局替换 `Theme.CreateButton` → `UiToolkit.CreateButton`，`Theme.CreateDivider` → `UiToolkit.CreateDivider`。

涉及文件：
- `WordTools/Forms/InsertPhotosForm.cs`
- `WordTools/Forms/ExcelDataFillerForm.cs`
- `WordTools/Forms/ProgressForm.cs`
- `WordTools/Forms/FailureDetailsForm.cs`

- [ ] **Step 4: 注册到 csproj**

在 `WordTools.csproj` 中添加：
```xml
<Compile Include="UiToolkit.cs" />
```

- [ ] **Step 5: 构建验证**

```bash
msbuild WordTools.sln /p:Configuration=Release
dotnet test
```

- [ ] **Step 6: 提交**

```bash
git add WordTools/UiToolkit.cs WordTools/Theme.cs WordTools/Forms/*.cs WordTools/WordTools.csproj
git commit -m "refactor: split Theme into Theme + UiToolkit"
```

---

## 阶段 6：测试覆盖补全

> 为拆分后的模块补充单元测试，重点覆盖纯逻辑部分。

### Task 6.1：HighPerformanceModeController 测试

**Files:**
- Create: `WordTools.Tests/HighPerformanceModeControllerTests.cs`
- Modify: `WordTools.Tests/WordTools.Tests.csproj`

- [ ] **Step 1: 编写测试**

```csharp
// WordTools.Tests/HighPerformanceModeControllerTests.cs
using WordTools.Services;
using WordTools.Services.Abstractions;
using Xunit;

namespace WordTools.Tests
{
    public class HighPerformanceModeControllerTests
    {
        [Theory]
        [InlineData(10, 10)]
        [InlineData(29, 10)]
        [InlineData(30, 15)]
        [InlineData(99, 15)]
        [InlineData(100, 20)]
        [InlineData(500, 20)]
        public void GetOptimizedRefreshInterval_ReturnsExpected(int totalFiles, int expected)
        {
            var controller = new HighPerformanceModeController(new FakeAppContext());
            Assert.Equal(expected, controller.GetOptimizedRefreshInterval(totalFiles));
        }

        [Theory]
        [InlineData(5, 1)]
        [InlineData(10, 1)]
        [InlineData(11, 5)]
        [InlineData(50, 5)]
        [InlineData(51, 15)]
        [InlineData(200, 15)]
        [InlineData(201, 25)]
        public void GetStatusBarUpdateInterval_ReturnsExpected(int totalFiles, int expected)
        {
            var controller = new HighPerformanceModeController(new FakeAppContext());
            Assert.Equal(expected, controller.GetStatusBarUpdateInterval(totalFiles));
        }
    }

    internal class FakeAppContext : IWordApplicationContext
    {
        public Microsoft.Office.Interop.Word.Application Application => null;
        public bool ScreenUpdating { get; set; } = true;
        public void SetStatusBar(string text) { }
        public void DoEvents() { }
    }
}
```

- [ ] **Step 2: 注册到测试 csproj**

在 `WordTools.Tests.csproj` 中添加：
```xml
<Compile Include="..\WordTools\Services\HighPerformanceModeController.cs" Link="Services\HighPerformanceModeController.cs" />
```

- [ ] **Step 3: 运行测试**

```bash
dotnet test
```
Expected: All tests pass.

- [ ] **Step 4: 提交**

```bash
git add WordTools.Tests/HighPerformanceModeControllerTests.cs WordTools.Tests/WordTools.Tests.csproj
git commit -m "test: add HighPerformanceModeController tests"
```

---

### Task 6.2：EscapeKeyMonitor 测试

**Files:**
- Create: `WordTools.Tests/EscapeKeyMonitorTests.cs`
- Modify: `WordTools.Tests/WordTools.Tests.csproj`

- [ ] **Step 1: 编写测试**

```csharp
// WordTools.Tests/EscapeKeyMonitorTests.cs
using WordTools.Services;
using Xunit;

namespace WordTools.Tests
{
    public class EscapeKeyMonitorTests
    {
        [Fact]
        public void IsCancelled_InitiallyFalse()
        {
            var monitor = new EscapeKeyMonitor(new FakeAppContext());
            Assert.False(monitor.IsCancelled);
        }

        [Fact]
        public void Reset_ClearsCancelledState()
        {
            var monitor = new EscapeKeyMonitor(new FakeAppContext());
            monitor.Reset();
            Assert.False(monitor.IsCancelled);
        }
    }
}
```

- [ ] **Step 2: 注册到测试 csproj**

在 `WordTools.Tests.csproj` 中添加：
```xml
<Compile Include="..\WordTools\Services\EscapeKeyMonitor.cs" Link="Services\EscapeKeyMonitor.cs" />
```

- [ ] **Step 3: 运行测试**

```bash
dotnet test
```

- [ ] **Step 4: 提交**

```bash
git add WordTools.Tests/EscapeKeyMonitorTests.cs WordTools.Tests/WordTools.Tests.csproj
git commit -m "test: add EscapeKeyMonitor tests"
```

---

### Task 6.3：WindowActivationService 测试

**Files:**
- Create: `WordTools.Tests/WindowActivationServiceTests.cs`
- Modify: `WordTools.Tests/WordTools.Tests.csproj`

- [ ] **Step 1: 编写测试**

```csharp
// WordTools.Tests/WindowActivationServiceTests.cs
using System;
using WordTools.Services;
using Xunit;

namespace WordTools.Tests
{
    public class WindowActivationServiceTests
    {
        [Fact]
        public void EnsureWindowTopMost_ZeroHandle_DoesNotThrow()
        {
            WindowActivationService.EnsureWindowTopMost(IntPtr.Zero);
        }

        [Fact]
        public void EnsureWordWindowActive_ZeroHandle_DoesNotThrow()
        {
            WindowActivationService.EnsureWordWindowActive(IntPtr.Zero);
        }
    }
}
```

- [ ] **Step 2: 注册到测试 csproj 并运行**

```bash
dotnet test
```

- [ ] **Step 3: 提交**

```bash
git add WordTools.Tests/WindowActivationServiceTests.cs WordTools.Tests/WordTools.Tests.csproj
git commit -m "test: add WindowActivationService tests"
```

---

### Task 6.4：全量验证

- [ ] **Step 1: 全量构建**

```bash
msbuild WordTools.sln /p:Configuration=Release
```
Expected: Build succeeded, 0 errors, 0 warnings (或仅已有的已知警告).

- [ ] **Step 2: 全量测试**

```bash
dotnet test --verbosity normal
```
Expected: All tests pass. 记录测试数量。

- [ ] **Step 3: 确认文件行数**

```powershell
Get-ChildItem -Recurse -File -Filter '*.cs' |
  Where-Object { $_.FullName -notmatch '\\obj\\' -and $_.FullName -notmatch '\\bin\\' } |
  ForEach-Object {
    $content = [System.IO.File]::ReadAllText($_.FullName)
    $lines = ($content -split "`r`n|`n").Count
    [PSCustomObject]@{ Lines = $lines; File = $_.Name }
  } | Sort-Object Lines -Descending | Format-Table -AutoSize
```

Expected: 无文件超过 900 行。

- [ ] **Step 4: 确认接口覆盖**

```powershell
Get-ChildItem -Path 'WordTools\Services\Abstractions' -Filter 'I*.cs' | Select-Object Name
```
Expected: 至少 10 个接口文件。

- [ ] **Step 5: 最终提交**

```bash
git add -A
git commit -m "refactor: complete split refactoring - all phases verified"
```

---

## 验收标准

### 硬性指标

| 指标 | 重构前 | 重构后目标 | 验证方式 |
|------|--------|-----------|---------|
| 最大单文件行数 | 2088 (TableService) | ≤ 900 | `Get-Content` 行数统计 |
| ProgressService 行数 | 1742 | ≤ 900 | `Get-Content` 行数统计 |
| ThisAddIn 行数 | 430 | ≤ 250 | `Get-Content` 行数统计 |
| 接口数量 | 5 | ≥ 10 | `Abstractions/` 目录文件数 |
| 适配器数量 | 5 | ≥ 5 | `Adapters/` 目录文件数 |
| 单元测试数量 | ~15 | ≥ 30 | `dotnet test` 输出 |
| 构建结果 | 0 errors | 0 errors | `msbuild` 输出 |
| 测试结果 | All pass | All pass | `dotnet test` 输出 |

### 架构指标

| 指标 | 重构前 | 重构后目标 |
|------|--------|-----------|
| ProgressService 关注点数 | 6 | 1（编排） |
| TableService 职责数 | 2（结构+编号） | 1（结构） |
| ThisAddIn 职责数 | 3（COM+Ribbon+编排） | 2（COM+编排） |
| P/Invoke 声明位置 | ProgressService 内部 | 独立模块 |
| UI 代码在服务层 | FileService 含对话框 | 服务层无 UI |

### 行为不变保证

| 场景 | 验证方式 |
|------|---------|
| 批量插图（文件夹模式） | 手动在 Word 中测试，功能不变 |
| 批量插图（选中文件模式） | 手动在 Word 中测试，功能不变 |
| 刷新表格编号 | 手动在 Word 中测试，功能不变 |
| Excel 数据填充 | 手动在 Word 中测试，功能不变 |
| Ribbon 日志设置 | 手动在 Word 中测试，功能不变 |
| 关于对话框 | 手动在 Word 中测试，功能不变 |

### 文件结构目标

```
WordTools/
├── ThisAddIn.cs                    (≤250行, COM入口+编排)
├── RibbonController.cs             (≤120行, Ribbon回调)
├── Theme.cs                        (≤200行, 颜色/字体/DPI)
├── UiToolkit.cs                    (≤80行, 控件工厂)
├── Forms/                          (不变)
├── Services/
│   ├── Abstractions/               (≥10个接口)
│   ├── Adapters/                   (≥5个适配器)
│   ├── TableService.cs             (≤850行, 表格结构操作)
│   ├── TableNumberingService.cs    (≤1250行, 编号管理)
│   ├── ProgressService.cs          (≤900行, 批量插入编排)
│   ├── HighPerformanceModeController.cs (≤100行)
│   ├── EscapeKeyMonitor.cs         (≤50行)
│   ├── WindowActivationService.cs  (≤60行)
│   ├── InsertionErrorClassifier.cs (≤120行)
│   ├── InsertionResultPresenter.cs (≤300行)
│   ├── ImageService.cs             (不变, 后续优化)
│   ├── FileService.cs              (≤350行, 无UI)
│   ├── ConfigService.cs            (不变, 后续优化)
│   └── ...                         (其他不变)
```

### 每阶段独立验收

| 阶段 | 验收条件 |
|------|---------|
| 1 接口提取 | 构建通过 + 测试通过 + 5个新接口文件存在 |
| 2 TableService 拆分 | 构建通过 + 测试通过 + TableService ≤850行 + TableNumberingService 独立 |
| 3 ProgressService 拆分 | 构建通过 + 测试通过 + ProgressService ≤900行 + 5个新模块 |
| 4 Ribbon 提取 | 构建通过 + 测试通过 + ThisAddIn ≤250行 + RibbonController 独立 |
| 5 FileService/Theme 清理 | 构建通过 + 测试通过 + FileService 无 UI + UiToolkit 独立 |
| 6 测试补全 | 构建通过 + 测试通过 + 测试数 ≥30 + 无文件超900行 |

---

## 风险与缓解

| 风险 | 缓解措施 |
|------|---------|
| 拆分后调用链断裂 | 每阶段构建+测试验证，不跨阶段提交 |
| 中文编码损坏 | 使用 Edit/Write 工具精确补丁，禁止 PowerShell 整文件覆盖 |
| COM 加载项行为回归 | 每阶段手动在 Word 中验证核心功能 |
| 测试项目链接文件遗漏 | 每新增文件同步更新两个 csproj |
| ProgressService 内部状态耦合 | 提取时保持字段引用，通过构造函数注入子组件 |
| TableNumberingService 依赖 TableService 内部方法 | 提取时保留 SafeIgnore 副本，或将共享方法提升为 internal |
