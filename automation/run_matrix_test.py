#!/usr/bin/env python3
"""Matrix test entry point for Word/WPS multi-host automation."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from lib.cli import print_json_report
from lib.config import load_matrix_config
from lib.matrix_catalog import list_environments, resolve_config_path
from lib.preflight import validate_live_preflight
from lib.runner import MatrixTestRunner


def find_repo_root() -> Path:
    current = Path(__file__).resolve().parent.parent
    if (current / "WordTools.sln").exists():
        return current
    raise SystemExit("Unable to locate WordTools repository root.")


def main() -> int:
    parser = argparse.ArgumentParser(description="Run Word/WPS matrix automation tests.")
    parser.add_argument(
        "--config",
        help="Path to matrix config JSON, e.g. configs/word32_wps64.json",
    )
    parser.add_argument(
        "--env",
        help="Matrix environment id from configs/matrix_index.json, e.g. VM-03",
    )
    parser.add_argument(
        "--list-envs",
        action="store_true",
        help="List all matrix environments and exit.",
    )
    parser.add_argument(
        "--repo-root",
        default=None,
        help="Repository root override.",
    )
    parser.add_argument(
        "--skip-phases",
        default="",
        help="Comma-separated phases to skip, e.g. unregister,verify_cleanup",
    )
    parser.add_argument(
        "--ignore-preflight",
        action="store_true",
        help="Run even when Live preflight checks fail.",
    )
    args = parser.parse_args()

    repo_root = Path(args.repo_root) if args.repo_root else find_repo_root()
    automation_root = repo_root / "automation"

    if args.list_envs:
        environments = list_environments(automation_root / "configs" / "matrix_index.json")
        print_json_report({"environments": environments})
        return 0

    if args.env:
        config_path = resolve_config_path(automation_root, args.env)
    elif args.config:
        config_path = Path(args.config)
        if not config_path.is_absolute():
            config_path = automation_root / config_path
    else:
        raise SystemExit("Provide --config or --env, or use --list-envs.")

    if not config_path.exists():
        raise SystemExit(f"Config not found: {config_path}")

    config = load_matrix_config(config_path)
    preflight_issues = validate_live_preflight(config, repo_root)
    if preflight_issues and not args.ignore_preflight:
        for issue in preflight_issues:
            print(f"PREFLIGHT: {issue}", file=sys.stderr)
        raise SystemExit("Live preflight failed. Re-run from an elevated shell or use --ignore-preflight.")

    skip_phases = {part.strip() for part in args.skip_phases.split(",") if part.strip()}
    report = MatrixTestRunner(repo_root=repo_root, config_path=config_path, skip_phases=skip_phases).run()
    print_json_report(report)
    return 0 if report.get("pass") else 1


if __name__ == "__main__":
    sys.exit(main())
