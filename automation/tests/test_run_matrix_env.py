from pathlib import Path

import pytest

from run_matrix_test import find_repo_root
from lib.matrix_catalog import resolve_config_path


def test_resolve_config_path_points_to_word64_wps32():
    repo_root = find_repo_root()
    config_path = resolve_config_path(repo_root / "automation", "VM-03")

    assert config_path.name == "word64_wps32.json"
    assert config_path.exists()
