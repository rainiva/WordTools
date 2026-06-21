"""Tests for single Word-session batch E2E helpers."""

import pytest

from lib.e2e_tiers import BATCH_SMOKE, UI_SMOKE
from lib.ui_batch_runner import (
    batch_cases_for_tier,
    evaluate_ui_batch_run,
    parse_batch_payload,
    ui_cases_for_tier,
)


def test_ui_cases_for_tier_smoke():
    assert ui_cases_for_tier("smoke") == sorted(UI_SMOKE)


def test_batch_cases_for_tier_smoke():
    assert batch_cases_for_tier("smoke") == sorted(BATCH_SMOKE)


def test_parse_batch_payload_single():
    payload = {"case_id": "AC-UI-B05", "pass": True}
    assert parse_batch_payload(payload) == [payload]


def test_parse_batch_payload_multi():
    payload = {
        "batch": True,
        "pass": True,
        "cases": [
            {"case_id": "AC-UI-B05", "pass": True, "inline_shape_count": 1},
            {"case_id": "AC-UI-B07", "pass": True, "inline_shape_count": 4},
        ],
    }
    cases = parse_batch_payload(payload)
    assert len(cases) == 2
    assert cases[0]["case_id"] == "AC-UI-B05"


def test_evaluate_ui_batch_run_all_pass():
    payload = {
        "batch": True,
        "pass": True,
        "cases": [
            {
                "case_id": "AC-UI-B05",
                "pass": True,
                "ui_flow_started": True,
                "form_clicked": True,
                "progress_seen": True,
                "inline_shape_count": 1,
                "expected_image_count": 1,
                "table_row_count": 8,
            },
            {
                "case_id": "AC-UI-B09",
                "pass": True,
                "ui_flow_started": True,
                "form_clicked": True,
                "progress_seen": True,
                "inline_shape_count": 4,
                "expected_image_count": 4,
                "table_row_count": 8,
                "has_numbered_description": False,
            },
        ],
    }
    result = evaluate_ui_batch_run(["AC-UI-B05", "AC-UI-B09"], payload)
    assert result["pass"] is True
    assert len(result["cases"]) == 2


def test_evaluate_ui_batch_run_mismatch_count():
    payload = {"batch": True, "pass": True, "cases": [{"case_id": "AC-UI-B05", "pass": True}]}
    result = evaluate_ui_batch_run(["AC-UI-B05", "AC-UI-B09"], payload)
    assert result["pass"] is False


def test_ui_cases_for_tier_unknown():
    with pytest.raises(ValueError, match="unknown UI E2E tier"):
        ui_cases_for_tier("invalid")
