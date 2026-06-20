import ctypes
from pathlib import Path
from typing import Any


def is_administrator() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except OSError:
        return False


def validate_live_preflight(
    config: dict[str, Any],
    repo_root: Path,
    *,
    is_admin: bool | None = None,
) -> list[str]:
    issues: list[str] = []

    phases = config.get("phases", {})
    registration_cases = config.get("registration_cases", [])
    unregister_intent = config.get("unregister", {}).get("execution_intent", "PreviewOnly")

    needs_live = any(case.get("execution_intent") == "Live" for case in registration_cases)
    needs_live = needs_live or unregister_intent == "Live"

    if not needs_live:
        return issues

    if is_admin is None:
        is_admin = is_administrator()

    if not is_admin:
        issues.append("Live phases require an elevated (Administrator) shell.")

    plugin = config.get("plugin", {})
    configuration = plugin.get("configuration", "Release")
    dll_path = repo_root / "WordTools" / "bin" / configuration / "WordTools.dll"
    if not dll_path.exists():
        issues.append(f"Plugin DLL not found for Live run: {dll_path}")

    if phases.get("smoke") and config.get("smoke", {}).get("mode") == "com_load":
        issues.append("com_load smoke will launch Word; close all Word instances before running.")

    return issues
