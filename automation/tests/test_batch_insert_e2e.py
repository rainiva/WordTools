from pathlib import Path

import pytest

from lib.e2e_tiers import batch_case_tier_marks
from lib.expectations import evaluate_batch_insert_case
from lib.powershell import invoke_matrix_script
from lib.ui_batch_runner import batch_cases_for_tier, evaluate_headless_batch_run

BATCH_CASE_IDS = [
    "AC-B01", "AC-B02", "AC-B03", "AC-B04", "AC-B05", "AC-B06",
    "AC-B07", "AC-B08", "AC-B09", "AC-B10", "AC-B11", "AC-B12",
]


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _ensure_fixtures(repo_root: Path) -> None:
    fixtures_script = repo_root / "automation" / "scripts" / "generate-fixtures.ps1"
    template = repo_root / "automation" / "assets" / "table-template.docx"
    if not fixtures_script.exists():
        pytest.skip("fixture generator missing")
    if not template.exists():
        invoke_matrix_script(script_path=fixtures_script, parameters={})


def _run_batch_insert_cases(case_ids: list[str]) -> dict:
    repo_root = _repo_root()
    _ensure_fixtures(repo_root)

    script_path = repo_root / "automation" / "ps" / "Matrix.BatchInsert.ps1"
    parameters = {
        "RepoRoot": str(repo_root),
        "Visible": "false",
    }
    if len(case_ids) == 1:
        parameters["CaseId"] = case_ids[0]
    else:
        parameters["CaseIds"] = ";".join(case_ids)

    result = invoke_matrix_script(script_path=script_path, parameters=parameters)

    if result["exit_code"] != 0 and "Word is already running" in str(result.get("payload", {})):
        pytest.skip("Word is running; close Word before integration test")

    return result


def _run_batch_insert(case_id: str) -> dict:
    return _run_batch_insert_cases([case_id])


@pytest.mark.integration
@pytest.mark.smoke
@pytest.mark.standard
@pytest.mark.full
def test_batch_insert_headless_word_session(request):
    tier = request.config.getoption("--e2e-tier")
    if not tier:
        pytest.skip("word session batch requires --e2e-tier")

    case_ids = batch_cases_for_tier(tier)
    result = _run_batch_insert_cases(case_ids)
    assert result["exit_code"] == 0, result["stderr"]

    evaluation = evaluate_headless_batch_run(case_ids, result["payload"])
    assert evaluation["pass"] is True, evaluation


@pytest.mark.integration
@pytest.mark.parametrize(
    "case_id",
    [pytest.param(case_id, marks=batch_case_tier_marks(case_id), id=case_id) for case_id in BATCH_CASE_IDS],
)
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
