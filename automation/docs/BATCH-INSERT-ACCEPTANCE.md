# Batch Insert E2E Acceptance

Phase B covers orchestrated batch insert via `InsertPhotosOrchestrator.Execute` with headless adapters and real Word COM document assertions.

| AC ID | Scenario | Automated | Notes |
|-------|----------|-----------|-------|
| AC-B01 | Cursor not in table | Yes | Expect warning「请先选中一个表格！」; InlineShapes unchanged |
| AC-B02 | Cursor in column 2 | Yes | Expect warning「请将光标置于表格左侧单元格！」 |
| AC-B03 | Selected 4 files + numbering | Yes | InlineShapes=4; numbered descriptions |
| AC-B04 | Folder root + subfolder | Yes | InlineShapes=5; subfolder title row |
| AC-B05 | Single image | Yes | InlineShapes=1; last row col2=`N/A` |
| AC-B06 | Cancel mid-batch | Yes | HeadlessProgressReporter cancel after 2 updates |

## Run (dev machine)

Prerequisites: Release build, fixtures generated, **Word closed**.

```powershell
msbuild WordTools.sln /p:Configuration=Release
.\automation\scripts\generate-fixtures.ps1   # table-template.docx = 12×2 标准 Word 表格（插入→表格默认样式）

cd automation
pip install -r requirements.txt
python run_matrix_test.py --config configs/word64_batch_insert_e2e.json
python -m pytest tests/test_batch_insert_e2e.py -m integration
# Covers AC-B01 through AC-B06
python -m pytest tests/test_fixtures.py tests/test_batch_insert_expectations.py
```

Report layer: `automation/reports/word64_batch_insert_e2e/matrix-report.json` → `batch_insert`.

Phase A (full WinForms UI) uses real images from a local folder (default `C:\Users\coxte\Desktop\test2`, override with `WORDTOOLS_UI_IMAGE_ROOT` or `-ImageRoot`).

| AC ID | Scenario | Automated | Notes |
|-------|----------|-----------|-------|
| AC-UI-B03 | 选择文件 → 4 张真实图片 + ProgressForm | Yes (`ui_integration`) | Sorted first 4 under ImageRoot |
| AC-UI-B04 | 插入文件夹 → 根目录+子目录全部图片 | Yes (`ui_integration`) | All images under ImageRoot; long run |
| AC-UI-B05 | 选择文件 → 单张真实图片 | Yes (`ui_integration`) | First sorted image under ImageRoot |

```powershell
$env:WORDTOOLS_UI_IMAGE_ROOT = "C:\Users\coxte\Desktop\test2"
cd automation
python -m pytest tests/test_batch_insert_ui_e2e.py -m ui_integration -v
powershell -File ps/Matrix.BatchInsertUI.ps1 -RepoRoot .. -CaseId AC-UI-B04 -ImageRoot "C:\Users\coxte\Desktop\test2"
```

---

## Phase 2 Refactor Gate — PERF 性能基准

插图拆分重构不得降低插入性能。完整操作见 [`docs/superpowers/plans/PERF-BASELINE-RUNBOOK.md`](../../docs/superpowers/plans/PERF-BASELINE-RUNBOOK.md)。

| PERF ID | 场景 | run_mode | 文件数 | 说明 |
|---------|------|----------|--------|------|
| PERF-01 | 选中 4 文件 + 编号 | `SelectedFiles` | 4 | 等价 AC-B03 |
| PERF-02 | 文件夹根+子目录 | `Folder` | 5 | 等价 AC-B04 |
| PERF-03 | 大批量（Phase 4） | `SelectedFiles` | 50 | 可选 |

### 锁定基线（首次，各跑 3 次）

```powershell
# 1. Release 构建 + 注册加载项 + 开启 Ribbon「详细日志」「性能基准 CSV」
# 2. 按 RUNBOOK 手动跑 PERF-01 / PERF-02
# 3. 生成 median 基线：
.\scripts\capture-benchmark-baseline.ps1 `
  -SourceCsv "$env:USERPROFILE\Documents\WordTools\wordtools-benchmark.csv" `
  -OutputCsv "docs\superpowers\plans\perf-baseline-2026-06-21.csv" `
  -RunsPerScenario 3 -Force
```

### 每个涉及热路径的 Task / Phase 后必跑

```powershell
.\scripts\compare-benchmark.ps1 `
  -BaselineCsv "docs\superpowers\plans\perf-baseline-2026-06-21.csv" `
  -CurrentCsv "$env:USERPROFILE\Documents\WordTools\wordtools-benchmark.csv" `
  -ScenarioId All
# exit 0 = 通过；exit 1 = 性能回归
```

### 脚本自测

```powershell
.\scripts\compare-benchmark.ps1 -SelfTest
```
