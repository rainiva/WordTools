from lib.preflight import validate_live_preflight


def test_validate_live_preflight_warns_without_admin(tmp_path):
    config = {
        "phases": {"register": True},
        "registration_cases": [{"execution_intent": "Live"}],
        "plugin": {"configuration": "Release"},
    }
    repo_root = tmp_path / "repo"
    (repo_root / "WordTools" / "bin" / "Release").mkdir(parents=True)
    (repo_root / "WordTools" / "bin" / "Release" / "WordTools.dll").write_text("x", encoding="utf-8")

    issues = validate_live_preflight(config, repo_root)

    assert any("Administrator" in issue for issue in issues)
