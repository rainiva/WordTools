from pathlib import Path

import pytest

from lib.e2e_tiers import UI_FLUI_STANDARD, UI_FULL_ONLY, ui_case_tier_marks, ui_flui_case_tier_marks
from lib.expectations import evaluate_batch_insert_ui_case
from lib.powershell import invoke_matrix_script
from lib.ui_batch_runner import evaluate_ui_batch_run, ui_cases_for_tier_direct
from lib.ui_image_fixtures import (
    default_real_image_root,
    expected_ui_case_count,
    subfolder_names,
)

UI_CASE_IDS = [
    "AC-UI-B05",
    "AC-UI-B07",
    "AC-UI-B09",
    "AC-UI-B03",
    "AC-UI-B10",
    "AC-UI-B12",
    "AC-UI-B14",
    "AC-UI-B15",
    "AC-UI-B04",
    "AC-UI-B08",
    "AC-UI-B11",
]


def _real_image_root() -> Path:
    root = default_real_image_root()
    if not root.is_dir():
        pytest.skip(f"real image root missing: {root}")
    return root


def _run_ui_cases(case_ids: list[str], image_root: Path, *, direct: bool = True) -> dict:
    repo_root = Path(__file__).resolve().parents[2]
    script_path = repo_root / "automation" / "ps" / "Matrix.BatchInsertUI.ps1"
    parameters = {
        "RepoRoot": str(repo_root),
        "ImageRoot": str(image_root),
        "Direct": "true" if direct else "false",
    }
    if len(case_ids) == 1:
        parameters["CaseId"] = case_ids[0]
    else:
        parameters["CaseIds"] = ";".join(case_ids)
    return invoke_matrix_script(script_path=script_path, parameters=parameters)


def _run_ui_case(case_id: str, image_root: Path, *, direct: bool = False) -> dict:
    return _run_ui_cases([case_id], image_root, direct=direct)


def _assert_ui_case(case_id: str, image_root: Path, *, direct: bool = False) -> None:
    expected_count = expected_ui_case_count(case_id, image_root)
    result = _run_ui_case(case_id, image_root, direct=direct)

    if result["exit_code"] != 0 and "Word is already running" in str(result.get("payload", {})):
        pytest.skip("Word is running; close Word before UI integration test")

    if result["exit_code"] != 0 and "add-in not loaded" in str(result.get("payload", {})).lower():
        pytest.skip("WordTools add-in not registered for UI E2E")

    assert result["exit_code"] == 0, result
    payload = dict(result["payload"])
    payload["expected_image_count"] = expected_count
    if case_id == "AC-UI-B07":
        payload["expect_zero_images"] = expected_count == 0
    if case_id in UI_FULL_ONLY:
        payload["subfolder_title_hints"] = subfolder_names(image_root)
    evaluation = evaluate_batch_insert_ui_case(case_id, payload)
    assert evaluation["pass"] is True, evaluation


def _skip_if_real_root_unsuitable(case_id: str, root: Path) -> None:
    if case_id in {"AC-UI-B03", "AC-UI-B09", "AC-UI-B10", "AC-UI-B12", "AC-UI-B14", "AC-UI-B15"}:
        if expected_ui_case_count(case_id, root) < 4:
            pytest.skip("need at least 4 real images")
    if case_id == "AC-UI-B08" and expected_ui_case_count(case_id, root) <= 0:
        pytest.skip("need subfolder images for AC-UI-B08")


@pytest.mark.ui_integration
@pytest.mark.smoke
@pytest.mark.standard
@pytest.mark.full
def test_batch_insert_ui_word_session(request):
    tier = request.config.getoption("--e2e-tier")
    if not tier:
        pytest.skip("word session batch requires --e2e-tier")

    root = _real_image_root()
    case_ids = ui_cases_for_tier_direct(tier)
    if not case_ids:
        pytest.skip("no direct UI cases for tier")

    for case_id in case_ids:
        _skip_if_real_root_unsuitable(case_id, root)

    result = _run_ui_cases(case_ids, root, direct=True)
    payload = result.get("payload", {})

    if result["exit_code"] != 0 and "Word is already running" in str(payload):
        pytest.skip("Word is running; close Word before UI integration test")

    if result["exit_code"] != 0 and "add-in not loaded" in str(payload).lower():
        pytest.skip("WordTools add-in not registered for UI E2E")

    assert result["exit_code"] == 0, result

    evaluation = evaluate_ui_batch_run(
        case_ids,
        payload,
        image_root_count_fn=lambda case_id: expected_ui_case_count(case_id, root),
        full_only_case_ids=UI_FULL_ONLY,
        subfolder_title_hints=subfolder_names(root),
    )
    assert evaluation["pass"] is True, evaluation


@pytest.mark.ui_integration
@pytest.mark.parametrize(
    "case_id",
    [
        pytest.param(
            case_id,
            marks=ui_flui_case_tier_marks(case_id),
            id=case_id,
        )
        for case_id in sorted(UI_FLUI_STANDARD)
    ],
)
def test_batch_insert_ui_flaui_path(case_id: str):
    root = _real_image_root()
    _assert_ui_case(case_id, root, direct=False)


@pytest.mark.ui_integration
@pytest.mark.parametrize(
    "case_id",
    [pytest.param(case_id, marks=ui_case_tier_marks(case_id), id=case_id) for case_id in UI_CASE_IDS],
)
def test_batch_insert_ui_real_image_parameter_matrix(case_id: str):
    root = _real_image_root()
    _skip_if_real_root_unsuitable(case_id, root)
    _assert_ui_case(case_id, root, direct=False)
