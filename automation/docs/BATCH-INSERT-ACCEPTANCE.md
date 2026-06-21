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
| AC-B07 | Folder root only | Yes | IncludeSubFolder=false; 3 shapes; no sub-a title |
| AC-B08 | Folder sub only | Yes | IncludeRoot=false; 2 shapes; sub-a title |
| AC-B09 | No description | Yes | Selected 4; no filename/folder desc; no numbering |
| AC-B10 | Number after + center | Yes | Selected 4; `desc-N` suffix; center alignment |
| AC-B11 | Folder name description | Yes | Folder mode; UseFolderNameAsDescription |
| AC-B12 | Manual description + numbering | Yes | Empty description rows with `1.` style numbering |

### Parameter coverage matrix (Phase B)

| Parameter | B03 | B04 | B05 | B07 | B08 | B09 | B10 | B11 | B12 |
|-----------|-----|-----|-----|-----|-----|-----|-----|-----|-----|
| Mode SelectedFiles | ✓ | | ✓ | | | ✓ | ✓ | | ✓ |
| Mode Folder | | ✓ | | ✓ | ✓ | | | ✓ | |
| IncludeRoot | | ✓ | | ✓ | | | | ✓ | |
| IncludeSubFolder | | ✓ | | | ✓ | | | ✓ | |
| Filename desc | ✓ | ✓ | ✓ | ✓ | ✓ | | ✓ | | |
| Folder name desc | | | | | | | | ✓ | |
| Manual desc | | | | | | | | | ✓ |
| No description | | | | | | ✓ | | | |
| Auto numbering | ✓ | ✓ | | ✓ | ✓ | | ✓ | ✓ | ✓ |
| Number after desc | | | | | | | ✓ | | |
| Number center | | | | | | | ✓ | | |
| Single image N/A | | | ✓ | | | | | | |

## Run (dev machine)

Prerequisites: Release build, fixtures generated, **Word closed**.

```powershell
msbuild WordTools.sln /p:Configuration=Release
.\automation\scripts\generate-fixtures.ps1   # table-template.docx = 12×2 标准 Word 表格（插入→表格默认样式）

cd automation
pip install -r requirements.txt
python run_matrix_test.py --config configs/word64_batch_insert_e2e.json
python -m pytest tests/test_batch_insert_e2e.py -m integration
# Covers AC-B01 through AC-B12
python -m pytest tests/test_fixtures.py tests/test_batch_insert_expectations.py
```

Report layer: `automation/reports/word64_batch_insert_e2e/matrix-report.json` → `batch_insert`.

## E2E 分级规则

规则定义见 `automation/lib/e2e_tiers.py`（pytest markers: `smoke` / `standard` / `full`）。

| 级别 | 何时跑 | UI 用例 | 最多插入（test2） | Headless |
|------|--------|---------|-------------------|----------|
| **smoke** | 改插图代码后 | B05, B07（COM 直连） | ~5 张 | B01–B03,B05,B07,B09 |

**性能：** `--e2e-tier smoke` 时 UI 走 **COM 直连**（无 FlaUI），同层用例单次 Word 会话；目标 **2–3 分钟**。standard 额外跑 1 条 FlaUI（B05）。调试单用例加 `--e2e-per-case`。
| **standard** | 合并前 / 每日 | smoke + B03,B10,B12,B14,B15 | ≤4 张/用例 | B01–B12 全参数 |
| **full** | 发版 / nightly | standard + B04,B08,B11 | ~29 张/用例 | 同 standard |

```powershell
cd automation
# 冒烟（推荐日常）
.\ps\Run-BatchInsertE2E-Smoke.ps1 -ImageRoot "C:\Users\coxte\Desktop\test2"
# 或等价 pytest：
py -m pytest tests/test_batch_insert_e2e.py --e2e-tier smoke -m "integration and smoke" -v
$env:WORDTOOLS_UI_IMAGE_ROOT = "C:\Users\coxte\Desktop\test2"
py -m pytest tests/test_batch_insert_ui_e2e.py --e2e-tier smoke -m "ui_integration and smoke" -v

# 标准（不含全文件夹）
py -m pytest tests/test_batch_insert_ui_e2e.py --e2e-tier standard -m ui_integration -v

# 全量
py -m pytest tests/test_batch_insert_ui_e2e.py --e2e-tier full -m ui_integration -v
```

Phase A (full WinForms UI) uses real images from a local folder (default `C:\Users\coxte\Desktop\test2`, override with `WORDTOOLS_UI_IMAGE_ROOT` or `-ImageRoot`).

### UI 参数覆盖矩阵（真实图片）

| 界面参数 | 用例 | 说明 |
|----------|------|------|
| 选择文件 + 文件名 + 编号 + 居中 | AC-UI-B03 | 4 张 |
| 插入文件夹 + 根+子 + 文件名 + 编号 | AC-UI-B04 | 全部（test2 约 25 张，慢） |
| 选择文件 + 文件名 + 无编号 | AC-UI-B05 | 1 张 |
| 范围：仅根目录 | AC-UI-B07 | 根目录 4 张（不含子目录） |
| 范围：仅子目录 | AC-UI-B08 | 全部子目录图（慢） |
| 描述：无 | AC-UI-B09 | 4 张，无编号 |
| 编号在后 + 居中 | AC-UI-B10 | 4 张 |
| 描述：文件夹名 | AC-UI-B11 | 全部（慢） |
| 描述：手动 + 编号 | AC-UI-B12 | 4 张 |
| 编号在前 + 靠左 | AC-UI-B14 | 4 张 |
| 高度：3cm 固定 | AC-UI-B15 | 4 张（验证插入成功） |

| AC ID | Scenario | Automated | Notes |
|-------|----------|-----------|-------|
| AC-UI-B03 | 选择文件 → 4 张 + 编号 | Yes | 真实 ImageRoot |
| AC-UI-B04 | 插入文件夹 → 全部 | Yes | 大批量，>20 张需确认框 |
| AC-UI-B05 | 单张 + 无编号 | Yes | |
| AC-UI-B07 | 仅根目录 | Yes | 根目录图片数（test2 当前 4 张） |
| AC-UI-B08 | 仅子目录 | Yes | 预期子文件夹标题行 |
| AC-UI-B09 | 无说明 | Yes | 4 张 |
| AC-UI-B10 | 编号在后 + 居中 | Yes | |
| AC-UI-B11 | 文件夹名说明 | Yes | |
| AC-UI-B12 | 手动说明 + 编号 | Yes | |
| AC-UI-B14 | 编号在前 + 靠左 | Yes | |
| AC-UI-B15 | 固定高度 3cm | Yes | |

```powershell
$env:WORDTOOLS_UI_IMAGE_ROOT = "C:\Users\coxte\Desktop\test2"
cd automation
# 日常冒烟（~9 张图，见 E2E 分级规则）：
.\ps\Run-BatchInsertE2E-Smoke.ps1
# 标准参数矩阵（不含 B04/B08/B11 大批量）：
py -m pytest tests/test_batch_insert_ui_e2e.py -m "ui_integration and standard" -v
# 全量（含 ~29 张/用例）：
py -m pytest tests/test_batch_insert_ui_e2e.py -m "ui_integration and full" -v
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
