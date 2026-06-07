import pytest

from lib.runner import MatrixTestRunner


@pytest.mark.integration
def test_live_config_enables_all_phases(tmp_path):
    import json
    from pathlib import Path
    from unittest.mock import patch

    repo_root = Path(__file__).resolve().parents[2]
    config_path = repo_root / "automation" / "configs" / "word64_wps32_live_register.json"
    if not config_path.exists():
        pytest.skip("live config missing")

    config = json.loads(config_path.read_text(encoding="utf-8"))
    phases = config["phases"]
    assert phases["register"] is True
    assert config["registration_cases"][0]["execution_intent"] == "Live"
    assert phases["verify_registration"] is True
    assert phases["smoke"] is True
    assert config["smoke"]["mode"] == "com_load"


def test_runner_honors_skip_phases(sample_config, tmp_path):
    from pathlib import Path
    from unittest.mock import patch

    repo_root = Path(__file__).resolve().parents[2]

    with patch("lib.runner.invoke_matrix_script", return_value={"exit_code": 0, "payload": {"pass": True, "layer": "host_probe", "word": {"installed": True, "bitness": "64"}, "wps": {"installed": True, "bitness": "32"}}}) as mocked:
        runner = MatrixTestRunner(repo_root=repo_root, config_path=sample_config, skip_phases={"unregister", "verify_cleanup"})
        report = runner.run()

    layer_names = [layer["name"] for layer in report["layers"]]
    assert "unregister" not in layer_names
    assert mocked.call_count >= 1


@pytest.fixture
def sample_config(tmp_path):
    import json

    config_path = tmp_path / "live.json"
    config_path.write_text(
        json.dumps(
            {
                "env_name": "live",
                "hosts": {"word": {"enabled": True, "expected_bitness": "64"}, "wps": {"enabled": False, "expected_bitness": "32"}},
                "phases": {"probe": True, "register": False, "verify_registration": False, "smoke": False, "unregister": True, "verify_cleanup": True},
                "output": {"report_dir": "reports/live"},
            }
        ),
        encoding="utf-8",
    )
    return config_path
