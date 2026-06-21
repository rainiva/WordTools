"""E2E 分级规则（单一事实来源）。

分级原则：
- 控制每次插入的图片总数与耗时
- smoke：改代码后必跑，分钟级
- standard：合并前 / 每日，覆盖参数矩阵但不跑大批量文件夹
- full：发版前 / nightly，含全文件夹（~29 张）用例
"""

from __future__ import annotations

import pytest

# --- UI（真实图片，ImageRoot 默认 test2）---

# ≤5 张：选文件 + 文件夹仅根目录；COM 直连（无 FlaUI）
UI_SMOKE: frozenset[str] = frozenset({"AC-UI-B05", "AC-UI-B07"})

# standard 保留 1 条完整 FlaUI 路径（窗体 + 进度 + 完成）
UI_FLUI_STANDARD: frozenset[str] = frozenset({"AC-UI-B05"})

# 在 smoke 基础上 + 各参数变体（均 ≤4 张/次，或根目录少量）
UI_STANDARD: frozenset[str] = UI_SMOKE | frozenset(
    {
        "AC-UI-B03",  # 文件名 + 编号
        "AC-UI-B09",  # 无说明（headless AC-B09 亦覆盖 smoke）
        "AC-UI-B10",  # 编号在后 + 居中
        "AC-UI-B12",  # 手动说明 + 编号
        "AC-UI-B14",  # 编号在前 + 靠左
        "AC-UI-B15",  # 固定高度 3cm
    }
)

# 全文件夹大批量（根+子 或 仅子 全部图片，单次 25~29 张）
UI_FULL_ONLY: frozenset[str] = frozenset({"AC-UI-B04", "AC-UI-B08", "AC-UI-B11"})

UI_FULL: frozenset[str] = UI_STANDARD | UI_FULL_ONLY

# --- Headless（automation/assets 小 fixture，单次 ≤5 张）---

BATCH_SMOKE: frozenset[str] = frozenset(
    {"AC-B01", "AC-B02", "AC-B03", "AC-B05", "AC-B07", "AC-B09"}
)

BATCH_STANDARD: frozenset[str] = frozenset(
    {
        "AC-B01",
        "AC-B02",
        "AC-B03",
        "AC-B04",
        "AC-B05",
        "AC-B06",
        "AC-B07",
        "AC-B08",
        "AC-B09",
        "AC-B10",
        "AC-B11",
        "AC-B12",
    }
)

BATCH_FULL: frozenset[str] = BATCH_STANDARD


def _tier_marks(case_id: str, smoke: frozenset[str], standard: frozenset[str], full: frozenset[str]) -> list:
    marks: list = []
    if case_id in smoke:
        marks.append(pytest.mark.smoke)
    if case_id in standard:
        marks.append(pytest.mark.standard)
    if case_id in full:
        marks.append(pytest.mark.full)
    return marks


def ui_flui_case_tier_marks(case_id: str) -> list:
    marks: list = []
    if case_id in UI_FLUI_STANDARD:
        marks.append(pytest.mark.standard)
        marks.append(pytest.mark.full)
    return marks


def ui_cases_for_tier_direct(tier: str) -> list[str]:
    """UI cases executed via COM direct path (no FlaUI) for the given tier."""
    tier = tier.lower()
    if tier == "smoke":
        return sorted(UI_SMOKE)
    if tier == "standard":
        return sorted(UI_STANDARD - UI_FLUI_STANDARD)
    if tier == "full":
        return sorted(UI_FULL - UI_FLUI_STANDARD)
    raise ValueError(f"unknown UI E2E tier: {tier}")


def ui_case_tier_marks(case_id: str) -> list:
    return _tier_marks(case_id, UI_SMOKE, UI_STANDARD, UI_FULL)


def batch_case_tier_marks(case_id: str) -> list:
    return _tier_marks(case_id, BATCH_SMOKE, BATCH_STANDARD, BATCH_FULL)


def max_images_for_ui_case(case_id: str, root_count_all: int, root_count_root: int, root_count_sub: int) -> int | None:
    """返回该用例最多插入张数（用于文档/估算）；None 表示与 ImageRoot 规模相关。"""
    case_id = case_id.upper()
    caps = {
        "AC-UI-B03": 4,
        "AC-UI-B05": 1,
        "AC-UI-B07": root_count_root,
        "AC-UI-B09": 4,
        "AC-UI-B10": 4,
        "AC-UI-B12": 4,
        "AC-UI-B14": 4,
        "AC-UI-B15": 4,
        "AC-UI-B04": root_count_all,
        "AC-UI-B08": root_count_sub,
        "AC-UI-B11": root_count_all,
    }
    return caps.get(case_id)
