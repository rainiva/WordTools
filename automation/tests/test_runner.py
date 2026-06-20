import json
from pathlib import Path
from unittest.mock import patch

import pytest

from lib.runner import MatrixTestRunner


@pytest.fixture
def sample_config(tmp_path: Path) -> Path:
    config_path = tmp_path / "word32_wps64.json"
    config_path.write_text(
        json.dumps(
            {
                "env_name": "word32_wps64",
                "plugin": {
                    "configuration": "Release",
                    "prog_id": "WordTools.ThisAddIn",
                },
                "hosts": {
                    "word": {"enabled": True, "expected_bitness": "32"},
                    "wps": {"enabled": True, "expected_bitness": "64"},
                },
                "phases": {
                    "probe": True,
                    "register": False,
                    "verify_registration": False,
                    "smoke": False,
                    "unregister": False,
                    "verify_cleanup": False,
                },
                "output": {"report_dir": "reports/word32_wps64"},
            }
        ),
        encoding="utf-8",
    )
    return config_path


def test_runner_probe_phase_marks_report_pass_when_expectations_match(sample_config: Path):
    repo_root = Path(__file__).resolve().parents[2]
    fake_probe = {
        "layer": "host_probe",
        "word": {"installed": True, "bitness": "32"},
        "wps": {"installed": True, "bitness": "64"},
    }

    with patch("lib.runner.invoke_matrix_script", return_value={"exit_code": 0, "payload": fake_probe}):
        runner = MatrixTestRunner(repo_root=repo_root, config_path=sample_config)
        report = runner.run()

    probe_layer = next(layer for layer in report["layers"] if layer["name"] == "host_probe")
    assert probe_layer["payload"]["pass"] is True
    assert report["pass"] is True


def test_runner_probe_phase_fails_when_expected_host_missing(sample_config: Path):
    repo_root = Path(__file__).resolve().parents[2]
    fake_probe = {
        "layer": "host_probe",
        "word": {"installed": True, "bitness": "32"},
        "wps": {"installed": False, "bitness": None},
    }

    with patch("lib.runner.invoke_matrix_script", return_value={"exit_code": 0, "payload": fake_probe}):
        runner = MatrixTestRunner(repo_root=repo_root, config_path=sample_config)
        report = runner.run()

    assert report["pass"] is False
