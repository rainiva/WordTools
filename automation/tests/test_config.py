import json
from pathlib import Path

import pytest

from lib.config import load_matrix_config, resolve_paths


def test_load_matrix_config_reads_env_name_and_hosts(tmp_path: Path):
    config_path = tmp_path / "word32_wps64.json"
    config_path.write_text(
        json.dumps(
            {
                "env_name": "word32_wps64",
                "hosts": {
                    "word": {"enabled": True, "expected_bitness": "32"},
                    "wps": {"enabled": True, "expected_bitness": "64"},
                },
                "output": {"report_dir": "reports/word32_wps64"},
            }
        ),
        encoding="utf-8",
    )

    config = load_matrix_config(config_path)

    assert config["env_name"] == "word32_wps64"
    assert config["hosts"]["word"]["expected_bitness"] == "32"
    assert config["hosts"]["wps"]["expected_bitness"] == "64"


def test_resolve_paths_makes_report_dir_absolute(tmp_path: Path):
    repo_root = tmp_path / "repo"
    repo_root.mkdir()
    config = {
        "env_name": "local",
        "output": {"report_dir": "reports/local", "screenshot_dir": "screenshots/local"},
    }

    resolved = resolve_paths(config, repo_root)

    assert resolved["report_dir"] == repo_root / "automation" / "reports" / "local"
    assert resolved["screenshot_dir"] == repo_root / "automation" / "screenshots" / "local"
