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
.\automation\scripts\generate-fixtures.ps1

cd automation
pip install -r requirements.txt
python run_matrix_test.py --config configs/word64_batch_insert_e2e.json
python -m pytest tests/test_batch_insert_e2e.py -m integration
# Covers AC-B01 through AC-B06
python -m pytest tests/test_fixtures.py tests/test_batch_insert_expectations.py
```

Report layer: `automation/reports/word64_batch_insert_e2e/matrix-report.json` → `batch_insert`.

Phase A (full WinForms UI) is deferred.
