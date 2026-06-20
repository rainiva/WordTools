import json
from pathlib import Path
from unittest.mock import patch

import pytest

from lib.runner import MatrixTestRunner


@pytest.fixture
def full_config(tmp_path: Path) -> Path:
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
                    "register": True,
                    "verify_registration": True,
                    "smoke": False,
                    "unregister": True,
                    "verify_cleanup": True,
                },
                "registration_cases": [
                    {
                        "case_id": "REG-022",
                        "case_name": "同时注册 Word + WPS",
                        "requested_host": "Both",
                        "execution_intent": "PreviewOnly",
                    }
                ],
                "output": {"report_dir": "reports/word32_wps64"},
            }
        ),
        encoding="utf-8",
    )
    return config_path


def test_runner_runs_registration_preview_case(full_config: Path):
    repo_root = Path(__file__).resolve().parents[2]

    def fake_invoke(script_path, parameters=None):
        name = script_path.name
        if name == "Matrix.HostProbe.ps1":
            return {
                "exit_code": 0,
                "payload": {
                    "layer": "host_probe",
                    "word": {"installed": True, "bitness": "32"},
                    "wps": {"installed": True, "bitness": "64"},
                },
            }
        if name == "Matrix.Register.ps1":
            return {
                "exit_code": 0,
                "payload": {
                    "layer": "register",
                    "case_id": parameters.get("CaseId", ""),
                    "word_register_success": True,
                    "wps_register_success": True,
                    "dll_bitness_match": True,
                    "pass": True,
                },
            }
        if name == "Matrix.VerifyRegistration.ps1":
            return {
                "exit_code": 0,
                "payload": {
                    "layer": "verify_registration",
                    "word_registry_ok": True,
                    "wps_registry_ok": True,
                    "pass": True,
                },
            }
        if name == "Matrix.Unregister.ps1":
            return {
                "exit_code": 0,
                "payload": {"layer": "unregister", "pass": True},
            }
        if name == "Matrix.VerifyCleanup.ps1":
            return {
                "exit_code": 0,
                "payload": {
                    "layer": "verify_cleanup",
                    "word_clean": True,
                    "wps_clean": True,
                    "pass": True,
                },
            }
        raise AssertionError(f"unexpected script {name}")

    with patch("lib.runner.invoke_matrix_script", side_effect=fake_invoke):
        runner = MatrixTestRunner(repo_root=repo_root, config_path=full_config)
        report = runner.run()

    layer_names = [layer["name"] for layer in report["layers"]]
    assert "host_probe" in layer_names
    assert "registration" in layer_names
    assert "verify_registration" in layer_names
    assert "unregister" in layer_names
    assert "verify_cleanup" in layer_names
