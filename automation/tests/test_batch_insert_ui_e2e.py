from pathlib import Path

import pytest

from lib.expectations import evaluate_batch_insert_ui_case
from lib.powershell import invoke_matrix_script
from lib.ui_image_fixtures import (
    count_all_images,
    default_real_image_root,
    selected_four_paths,
    single_image_path,
)


def _real_image_root() -> Path:
    root = default_real_image_root()
    if not root.is_dir():
        pytest.skip(f"real image root missing: {root}")
    return root


def _run_ui_case(case_id: str, image_root: Path) -> dict:
    repo_root = Path(__file__).resolve().parents[2]
    script_path = repo_root / "automation" / "ps" / "Matrix.BatchInsertUI.ps1"
    return invoke_matrix_script(
        script_path=script_path,
        parameters={
            "RepoRoot": str(repo_root),
            "CaseId": case_id,
            "ImageRoot": str(image_root),
        },
    )


def _assert_ui_case(case_id: str, image_root: Path, expected_count: int) -> None:
    result = _run_ui_case(case_id, image_root)

    if result["exit_code"] != 0 and "Word is already running" in str(result.get("payload", {})):
        pytest.skip("Word is running; close Word before UI integration test")

    if result["exit_code"] != 0 and "add-in not loaded" in str(result.get("payload", {})).lower():
        pytest.skip("WordTools add-in not registered for UI E2E")

    assert result["exit_code"] == 0, result
    payload = dict(result["payload"])
    payload["expected_image_count"] = expected_count
    evaluation = evaluate_batch_insert_ui_case(case_id, payload)
    assert evaluation["pass"] is True, evaluation


@pytest.mark.ui_integration
def test_batch_insert_ui_ac_ui_b03_real_selected_four():
    root = _real_image_root()
    if count_all_images(root) < 4:
        pytest.skip("need at least 4 real images")
    _assert_ui_case("AC-UI-B03", root, 4)


@pytest.mark.ui_integration
def test_batch_insert_ui_ac_ui_b04_real_folder_all():
    root = _real_image_root()
    expected = count_all_images(root)
    if expected < 2:
        pytest.skip("need folder with multiple real images")
    _assert_ui_case("AC-UI-B04", root, expected)


@pytest.mark.ui_integration
def test_batch_insert_ui_ac_ui_b05_real_single():
    root = _real_image_root()
    single_image_path(root)
    _assert_ui_case("AC-UI-B05", root, 1)
