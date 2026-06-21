from pathlib import Path

import pytest

from lib.ui_image_fixtures import (
    count_all_images,
    default_real_image_root,
    selected_four_paths,
    single_image_path,
)


def test_real_image_root_has_enough_images_for_ui_cases():
    root = default_real_image_root()
    if not root.is_dir():
        pytest.skip(f"real image root missing: {root}")

    assert count_all_images(root) >= 4
    assert len(selected_four_paths(root)) == 4
    assert single_image_path(root).is_file()
