from pathlib import Path

import pytest

from lib.powershell import invoke_matrix_script


@pytest.mark.integration
def test_register_script_uses_installer_core_when_available():
    repo_root = Path(__file__).resolve().parents[2]
    core_path = repo_root / "Installer.Core.ps1"
    if not core_path.exists():
        pytest.skip("Installer.Core.ps1 not present")

    script_path = repo_root / "automation" / "ps" / "Matrix.Register.ps1"
    result = invoke_matrix_script(
        script_path=script_path,
        parameters={
            "RepoRoot": str(repo_root),
            "RequestedHost": "Word",
            "ExecutionIntent": "PreviewOnly",
            "CaseId": "REG-020",
            "CaseName": "preview-word",
        },
    )

    assert result["exit_code"] == 0
    assert result["payload"]["source"] == "Installer.Core.ps1"
    assert "word_register_success" in result["payload"]


@pytest.mark.integration
def test_smoke_script_com_load_mode_reports_addin_fields():
    repo_root = Path(__file__).resolve().parents[2]
    script_path = repo_root / "automation" / "ps" / "Matrix.SmokeInsertImage.ps1"
    result = invoke_matrix_script(
        script_path=script_path,
        parameters={
            "RepoRoot": str(repo_root),
            "Mode": "com_load",
            "HostTarget": "Word",
        },
    )

    assert result["exit_code"] == 0
    assert result["payload"]["mode"] == "com_load"
    assert "word" in result["payload"]
    assert "addin_loaded" in result["payload"]["word"]
