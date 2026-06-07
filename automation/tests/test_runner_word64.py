import json
from pathlib import Path
from unittest.mock import patch

import pytest

from lib.runner import MatrixTestRunner


@pytest.fixture
def sample_word64_config(tmp_path: Path) -> Path:
    config_path = tmp_path / "word64_wps32.json"
    config_path.write_text(
        json.dumps(
            {
                "env_name": "word64_wps32",
                "hosts": {
                    "word": {"enabled": True, "expected_bitness": "64"},
                    "wps": {"enabled": True, "expected_bitness": "32"},
                },
                "phases": {
                    "probe": True,
                    "register": False,
                    "verify_registration": False,
                    "smoke": False,
                    "unregister": False,
                    "verify_cleanup": False,
                },
                "output": {"report_dir": "reports/word64_wps32"},
            }
        ),
        encoding="utf-8",
    )
    return config_path


def test_runner_accepts_word64_probe_result(sample_word64_config: Path):
    repo_root = Path(__file__).resolve().parents[2]
    fake_probe = {
        "layer": "host_probe",
        "word": {"installed": True, "bitness": "64", "path": r"C:\...\WINWORD.EXE"},
        "wps": {"installed": True, "bitness": "32", "path": r"C:\...\wps.exe"},
    }

    with patch("lib.runner.invoke_matrix_script", return_value={"exit_code": 0, "payload": fake_probe}):
        runner = MatrixTestRunner(repo_root=repo_root, config_path=sample_word64_config)
        report = runner.run()

    probe_layer = next(layer for layer in report["layers"] if layer["name"] == "host_probe")
    assert probe_layer["payload"]["word_bitness_ok"] is True
    assert report["pass"] is True

