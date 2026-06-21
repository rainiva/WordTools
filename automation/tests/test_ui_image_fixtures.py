from pathlib import Path

import pytest

from lib.ui_image_fixtures import (
    count_all_images,
    count_root_images,
    count_subfolder_images,
    default_real_image_root,
    expected_ui_case_count,
    selected_four_paths,
    single_image_path,
    subfolder_names,
)


def test_real_image_root_has_enough_images_for_ui_cases():
    root = default_real_image_root()
    if not root.is_dir():
        pytest.skip(f"real image root missing: {root}")

    assert count_all_images(root) >= 4
    assert len(selected_four_paths(root)) == 4
    assert single_image_path(root).is_file()


def test_real_image_root_scope_counts():
    root = default_real_image_root()
    if not root.is_dir():
        pytest.skip(f"real image root missing: {root}")

    total = count_all_images(root)
    root_only = count_root_images(root)
    sub_only = count_subfolder_images(root)
    assert total == root_only + sub_only
    assert sub_only > 0
    assert len(subfolder_names(root)) >= 1


def test_expected_ui_case_counts_for_test2_layout():
    root = default_real_image_root()
    if not root.is_dir():
        pytest.skip(f"real image root missing: {root}")

    assert expected_ui_case_count("AC-UI-B03", root) == 4
    assert expected_ui_case_count("AC-UI-B04", root) == count_all_images(root)
    assert expected_ui_case_count("AC-UI-B07", root) == count_root_images(root)
    assert count_root_images(root) >= 1
    assert expected_ui_case_count("AC-UI-B08", root) == count_subfolder_images(root)
