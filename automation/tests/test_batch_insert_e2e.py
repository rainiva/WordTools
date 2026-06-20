from pathlib import Path

import pytest

from lib.expectations import evaluate_batch_insert_case
from lib.powershell import invoke_matrix_script


def test_evaluate_ac_b04_folder_passes():
    result = evaluate_batch_insert_case(
        "AC-B04",
        {
            "pass": True,
            "inline_shape_count": 5,
            "success_count": 5,
            "has_subfolder_title": True,
        },
    )
    assert result["pass"] is True


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _ensure_fixtures(repo_root: Path) -> None:
    fixtures_script = repo_root / "automation" / "scripts" / "generate-fixtures.ps1"
    template = repo_root / "automation" / "assets" / "table-template.docx"
    if not fixtures_script.exists():
        pytest.skip("fixture generator missing")
    if not template.exists():
        invoke_matrix_script(script_path=fixtures_script, parameters={})


def _run_batch_insert(case_id: str) -> dict:
    repo_root = _repo_root()
    _ensure_fixtures(repo_root)

    script_path = repo_root / "automation" / "ps" / "Matrix.BatchInsert.ps1"
    result = invoke_matrix_script(
        script_path=script_path,
        parameters={
            "RepoRoot": str(repo_root),
            "CaseId": case_id,
            "Visible": "false",
        },
    )

    if result["exit_code"] != 0 and "Word is already running" in str(result.get("payload", {})):
        pytest.skip("Word is running; close Word before integration test")

    return result


@pytest.mark.integration
@pytest.mark.parametrize("case_id", ["AC-B01", "AC-B02", "AC-B03", "AC-B04", "AC-B05", "AC-B06"])
def test_batch_insert_acceptance_case(case_id: str):
    result = _run_batch_insert(case_id)

    assert result["exit_code"] == 0, result["stderr"]
    payload = result["payload"]
    evaluation = evaluate_batch_insert_case(case_id, payload)
    assert evaluation["pass"] is True, evaluation.get("details", [])

    if case_id == "AC-B03":
        assert payload["inline_shape_count"] == 4
    elif case_id == "AC-B01":
        assert payload["inline_shape_count"] == 0
    elif case_id == "AC-B06":
        assert payload.get("cancelled") is True
