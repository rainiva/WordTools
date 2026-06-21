"""Helpers for single Word-session UI E2E batch runs."""

from __future__ import annotations

from typing import Any

from lib.e2e_tiers import (
    BATCH_FULL,
    BATCH_SMOKE,
    BATCH_STANDARD,
    UI_FULL,
    UI_SMOKE,
    UI_STANDARD,
    ui_cases_for_tier_direct,
)
from lib.expectations import evaluate_batch_insert_case, evaluate_batch_insert_ui_case


def ui_cases_for_tier(tier: str) -> list[str]:
    tier = tier.lower()
    mapping = {
        "smoke": UI_SMOKE,
        "standard": UI_STANDARD,
        "full": UI_FULL,
    }
    if tier not in mapping:
        raise ValueError(f"unknown UI E2E tier: {tier}")
    return sorted(mapping[tier])


def batch_cases_for_tier(tier: str) -> list[str]:
    tier = tier.lower()
    mapping = {
        "smoke": BATCH_SMOKE,
        "standard": BATCH_STANDARD,
        "full": BATCH_FULL,
    }
    if tier not in mapping:
        raise ValueError(f"unknown batch E2E tier: {tier}")
    return sorted(mapping[tier])


def parse_batch_payload(payload: dict[str, Any]) -> list[dict[str, Any]]:
    if not payload.get("batch"):
        return [payload]
    cases = payload.get("cases")
    if not isinstance(cases, list):
        raise ValueError("batch payload missing cases array")
    return [dict(case) for case in cases]


def evaluate_ui_batch_run(
    case_ids: list[str],
    payload: dict[str, Any],
    *,
    image_root_count_fn=None,
    full_only_case_ids: frozenset[str] | None = None,
    subfolder_title_hints: list[str] | None = None,
) -> dict[str, Any]:
    """Evaluate a batch or single-case UI payload against expected case_ids order."""
    case_payloads = parse_batch_payload(payload)
    if len(case_payloads) != len(case_ids):
        return {
            "pass": False,
            "error": f"expected {len(case_ids)} case results, got {len(case_payloads)}",
            "cases": [],
        }

    evaluations: list[dict[str, Any]] = []
    all_pass = payload.get("pass", True) is not False

    for case_id, case_payload in zip(case_ids, case_payloads, strict=True):
        enriched = dict(case_payload)
        if image_root_count_fn is not None:
            enriched["expected_image_count"] = image_root_count_fn(case_id)
            if case_id == "AC-UI-B07":
                enriched["expect_zero_images"] = enriched["expected_image_count"] == 0
        if (
            full_only_case_ids
            and case_id in full_only_case_ids
            and subfolder_title_hints is not None
        ):
            enriched["subfolder_title_hints"] = subfolder_title_hints
        evaluation = evaluate_batch_insert_ui_case(case_id, enriched)
        evaluations.append({"case_id": case_id, **evaluation})
        if not evaluation.get("pass"):
            all_pass = False

    return {"pass": all_pass, "cases": evaluations}


def evaluate_headless_batch_run(case_ids: list[str], payload: dict[str, Any]) -> dict[str, Any]:
    case_payloads = parse_batch_payload(payload)
    if len(case_payloads) != len(case_ids):
        return {
            "pass": False,
            "error": f"expected {len(case_ids)} case results, got {len(case_payloads)}",
            "cases": [],
        }

    evaluations: list[dict[str, Any]] = []
    all_pass = payload.get("pass", True) is not False

    for case_id, case_payload in zip(case_ids, case_payloads, strict=True):
        evaluation = evaluate_batch_insert_case(case_id, case_payload)
        evaluations.append({"case_id": case_id, **evaluation})
        if not evaluation.get("pass"):
            all_pass = False

    return {"pass": all_pass, "cases": evaluations}


__all__ = [
    "batch_cases_for_tier",
    "evaluate_headless_batch_run",
    "evaluate_ui_batch_run",
    "parse_batch_payload",
    "ui_cases_for_tier",
    "ui_cases_for_tier_direct",
]
