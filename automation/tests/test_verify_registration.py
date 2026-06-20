from pathlib import Path

import pytest

from lib.powershell import invoke_matrix_script


@pytest.mark.integration
def test_verify_registration_detects_hklm_word_addin_when_present():
    repo_root = Path(__file__).resolve().parents[2]
    script_path = repo_root / "automation" / "ps" / "Matrix.VerifyRegistration.ps1"

    result = invoke_matrix_script(
        script_path=script_path,
        parameters={
            "RepoRoot": str(repo_root),
            "Configuration": "Release",
            "ExpectedWordBitness": "64",
        },
    )

    assert result["exit_code"] == 0
    payload = result["payload"]
    assert "word_registry_ok" in payload
    assert "word_registry_path" in payload
    if payload["word_registry_ok"]:
        assert "HKLM" in str(payload["word_registry_path"]) or "HKCU" in str(payload["word_registry_path"])
