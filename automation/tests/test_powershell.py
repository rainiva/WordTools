import json
import subprocess
from pathlib import Path

import pytest

from lib.powershell import invoke_matrix_script, parse_json_stdout


def test_parse_json_stdout_reads_multiline_json_object():
    stdout = '{\n    "layer": "host_probe",\n    "pass": true\n}\n'
    payload = parse_json_stdout(stdout)
    assert payload["layer"] == "host_probe"
    assert payload["pass"] is True


def test_parse_json_stdout_reads_last_json_object():
    stdout = "noise\n{\"layer\": \"host_probe\", \"pass\": true}\n"
    payload = parse_json_stdout(stdout)
    assert payload["layer"] == "host_probe"
    assert payload["pass"] is True


def test_invoke_matrix_script_runs_host_probe(tmp_path: Path):
    repo_root = Path(__file__).resolve().parents[2]
    ps_root = repo_root / "automation" / "ps"
    script_path = ps_root / "Matrix.HostProbe.ps1"
    if not script_path.exists():
        pytest.skip("Matrix.HostProbe.ps1 not implemented yet")

    result = invoke_matrix_script(
        script_path=script_path,
        parameters={"RepoRoot": str(repo_root), "OutputPath": str(tmp_path / "probe.json")},
    )

    assert result["exit_code"] == 0
    assert result["payload"]["layer"] == "host_probe"
    assert "word" in result["payload"]
    assert "wps" in result["payload"]
