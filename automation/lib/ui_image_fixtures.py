"""Real-image fixtures for Phase A UI E2E (local dev; path overridable via env)."""

from __future__ import annotations

import os
from pathlib import Path

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}


def default_real_image_root() -> Path:
    raw = os.environ.get("WORDTOOLS_UI_IMAGE_ROOT", r"C:\Users\coxte\Desktop\test2")
    return Path(raw)


def iter_image_files(root: Path):
    if not root.is_dir():
        return
    for path in sorted(root.rglob("*")):
        if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS:
            yield path


def count_all_images(root: Path) -> int:
    return sum(1 for _ in iter_image_files(root))


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
