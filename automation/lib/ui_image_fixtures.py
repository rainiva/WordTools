"""Real-image fixtures for Phase A UI E2E (local dev; path overridable via env)."""

from __future__ import annotations

import os
from pathlib import Path

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}


def default_real_image_root() -> Path:
    raw = os.environ.get("WORDTOOLS_UI_IMAGE_ROOT", r"C:\Users\coxte\Desktop\test2")
    return Path(raw)


def iter_image_files(root: Path, *, include_root: bool = True, include_subfolders: bool = True):
    if not root.is_dir():
        return
    for path in sorted(root.rglob("*")):
        if not path.is_file() or path.suffix.lower() not in IMAGE_EXTENSIONS:
            continue
        if path.parent == root:
            if include_root:
                yield path
        elif include_subfolders:
            yield path


def count_all_images(root: Path) -> int:
    return sum(1 for _ in iter_image_files(root))


def count_root_images(root: Path) -> int:
    return sum(1 for _ in iter_image_files(root, include_root=True, include_subfolders=False))


def count_subfolder_images(root: Path) -> int:
    return sum(1 for _ in iter_image_files(root, include_root=False, include_subfolders=True))


def count_for_scope(root: Path, *, include_root: bool, include_subfolders: bool) -> int:
    return sum(
        1
        for _ in iter_image_files(
            root,
            include_root=include_root,
            include_subfolders=include_subfolders,
        )
    )


def subfolder_names(root: Path) -> list[str]:
    if not root.is_dir():
        return []
    names: set[str] = set()
    for path in iter_image_files(root, include_root=False, include_subfolders=True):
        if path.parent != root:
            names.add(path.parent.name)
    return sorted(names)


def selected_four_paths(root: Path) -> list[Path]:
    files = list(iter_image_files(root))
    if len(files) < 4:
        raise ValueError(f"need at least 4 images under {root}, found {len(files)}")
    return files[:4]


def single_image_path(root: Path) -> Path:
    files = list(iter_image_files(root))
    if not files:
        raise ValueError(f"need at least 1 image under {root}")
    return files[0]


def expected_ui_case_count(case_id: str, root: Path) -> int:
    case_id = case_id.upper()
    if case_id == "AC-UI-B03":
        return 4
    if case_id == "AC-UI-B04":
        return count_for_scope(root, include_root=True, include_subfolders=True)
    if case_id == "AC-UI-B05":
        return 1
    if case_id == "AC-UI-B07":
        return count_for_scope(root, include_root=True, include_subfolders=False)
    if case_id == "AC-UI-B08":
        return count_for_scope(root, include_root=False, include_subfolders=True)
    if case_id in {"AC-UI-B09", "AC-UI-B10", "AC-UI-B12", "AC-UI-B14", "AC-UI-B15"}:
        return 4
    if case_id == "AC-UI-B11":
        return count_for_scope(root, include_root=True, include_subfolders=True)
    raise KeyError(case_id)
