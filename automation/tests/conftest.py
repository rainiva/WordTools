"""Pytest hooks for E2E tier selection and single Word-session batch runs."""

from __future__ import annotations

import pytest

UI_PARAM_TEST = "test_batch_insert_ui_real_image_parameter_matrix"
UI_BATCH_TEST = "test_batch_insert_ui_word_session"
UI_FLUI_TEST = "test_batch_insert_ui_flaui_path"
HEADLESS_PARAM_TEST = "test_batch_insert_acceptance_case"
HEADLESS_BATCH_TEST = "test_batch_insert_headless_word_session"


def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption(
        "--e2e-tier",
        action="store",
        default=None,
        choices=("smoke", "standard", "full"),
        help="Run only cases tagged for this E2E tier (smoke|standard|full).",
    )
    parser.addoption(
        "--e2e-per-case",
        action="store_true",
        default=False,
        help="When --e2e-tier is set, run one Word session per case instead of batching.",
    )


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    tier = config.getoption("--e2e-tier")
    per_case = config.getoption("--e2e-per-case")
    selected: list[pytest.Item] = []
    deselected: list[pytest.Item] = []

    for item in items:
        name = item.name
        if tier:
            if not per_case and (
                name.startswith(UI_PARAM_TEST) or name.startswith(HEADLESS_PARAM_TEST)
            ):
                deselected.append(item)
                continue
            if per_case and (name == UI_BATCH_TEST or name == HEADLESS_BATCH_TEST):
                deselected.append(item)
                continue
            if tier not in item.keywords:
                deselected.append(item)
                continue
        else:
            if name in {UI_BATCH_TEST, HEADLESS_BATCH_TEST}:
                deselected.append(item)
                continue

        selected.append(item)

    if deselected:
        config.hook.pytest_deselected(items=deselected)
    items[:] = selected
