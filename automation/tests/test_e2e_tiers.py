"""Tests for E2E tier definitions."""

from lib.e2e_tiers import (
    BATCH_FULL,
    BATCH_SMOKE,
    BATCH_STANDARD,
    UI_FULL,
    UI_FULL_ONLY,
    UI_SMOKE,
    UI_STANDARD,
)


def test_ui_smoke_is_minimal_and_not_bulk():
    assert "AC-UI-B05" in UI_SMOKE
    assert "AC-UI-B04" not in UI_SMOKE
    assert UI_FULL_ONLY.isdisjoint(UI_SMOKE)


def test_ui_standard_excludes_bulk_folder_cases():
    assert UI_FULL_ONLY.isdisjoint(UI_STANDARD)


def test_ui_full_covers_all_cases():
    assert UI_FULL == UI_STANDARD | UI_FULL_ONLY


def test_batch_smoke_subset_of_standard():
    assert BATCH_SMOKE <= BATCH_STANDARD
    assert BATCH_STANDARD == BATCH_FULL


def test_smoke_ui_case_count_cap():
    # smoke direct: 1 selected + root-only folder = 5 images max on typical test2
    assert len(UI_SMOKE) == 2
    assert "AC-UI-B09" not in UI_SMOKE


def test_batch_smoke_covers_root_only_and_no_description():
    assert "AC-B07" in BATCH_SMOKE
    assert "AC-B09" in BATCH_SMOKE
