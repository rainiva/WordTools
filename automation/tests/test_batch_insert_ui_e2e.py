from pathlib import Path

import pytest

from lib.expectations import evaluate_batch_insert_ui_case
from lib.powershell import invoke_matrix_script


@pytest.mark.ui_integration
def test_batch_insert_ui_ac_ui_b03_selected_four():
    repo_root = Path(__file__).resolve().parents[2]
    script_path = repo_root / "automation" / "ps" / "Matrix.BatchInsertUI.ps1"
    result = invoke_matrix_script(
        script_path=script_path,
        parameters={
            "RepoRoot": str(repo_root),
            "CaseId": "AC-UI-B03",
        },
    )

    if result["exit_code"] != 0 and "Word is already running" in str(result.get("payload", {})):
        pytest.skip("Word is running; close Word before UI integration test")

    if result["exit_code"] != 0 and "add-in not loaded" in str(result.get("payload", {})).lower():
        pytest.skip("WordTools add-in not registered for UI E2E")

    assert result["exit_code"] == 0, result["stderr"]
    evaluation = evaluate_batch_insert_ui_case("AC-UI-B03", result["payload"])
    assert evaluation["pass"] is True
