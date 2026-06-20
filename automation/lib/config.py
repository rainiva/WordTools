import json
from pathlib import Path
from typing import Any

VALID_BITNESS = {"32", "64"}


def load_matrix_config(config_path: Path) -> dict[str, Any]:
    with config_path.open(encoding="utf-8") as handle:
        config = json.load(handle)

    issues = validate_matrix_config(config)
    if issues:
        joined = "; ".join(issues)
        raise ValueError(f"Invalid matrix config {config_path}: {joined}")

    return config


def validate_matrix_config(config: dict[str, Any]) -> list[str]:
    issues: list[str] = []
    hosts = config.get("hosts", {})

    for host_name in ("word", "wps"):
        host = hosts.get(host_name, {})
        if not host.get("enabled", False):
            continue

        bitness = str(host.get("expected_bitness", ""))
        if bitness not in VALID_BITNESS:
            issues.append(f"{host_name}.expected_bitness must be one of {sorted(VALID_BITNESS)}")

    return issues


def resolve_paths(config: dict[str, Any], repo_root: Path) -> dict[str, Path]:
    output = config.get("output", {})
    report_dir = output.get("report_dir", "reports/default")
    screenshot_dir = output.get("screenshot_dir", "screenshots/default")

    automation_root = repo_root / "automation"
    return {
        "report_dir": _resolve_under_automation(automation_root, report_dir),
        "screenshot_dir": _resolve_under_automation(automation_root, screenshot_dir),
    }


def _resolve_under_automation(automation_root: Path, configured_path: str) -> Path:
    candidate = Path(configured_path)
    if candidate.is_absolute():
        return candidate
    return automation_root / candidate
