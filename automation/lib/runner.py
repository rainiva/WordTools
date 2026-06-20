from pathlib import Path
from typing import Any

from lib.config import load_matrix_config, resolve_paths
from lib.expectations import evaluate_cleanup_case, evaluate_host_probe, evaluate_registration_case
from lib.powershell import invoke_matrix_script
from lib.report import MatrixReport, write_report_json


class MatrixTestRunner:
    def __init__(self, repo_root: Path, config_path: Path, skip_phases: set[str] | None = None):
        self.repo_root = repo_root
        self.config_path = config_path
        self.config = load_matrix_config(config_path)
        self.paths = resolve_paths(self.config, repo_root)
        self.phases = self.config.get("phases", {})
        self.skip_phases = skip_phases or set()

    def _phase_enabled(self, name: str) -> bool:
        return bool(self.phases.get(name, False)) and name not in self.skip_phases

    def run(self) -> dict[str, Any]:
        report = MatrixReport(env_name=self.config["env_name"])

        if self.phases.get("probe", True):
            report.add_layer("host_probe", self._run_probe_phase())

        if self._phase_enabled("register"):
            report.add_layer("registration", self._run_registration_phase())

        if self._phase_enabled("verify_registration"):
            report.add_layer("verify_registration", self._run_verify_registration_phase())

        if self._phase_enabled("smoke"):
            report.add_layer("smoke", self._run_smoke_phase())

        if self._phase_enabled("batch_insert"):
            report.add_layer("batch_insert", self._run_batch_insert_phase())

        if self._phase_enabled("unregister"):
            report.add_layer("unregister", self._run_unregister_phase())

        if self._phase_enabled("verify_cleanup"):
            report.add_layer("verify_cleanup", self._run_verify_cleanup_phase())

        report_path = self.paths["report_dir"] / "matrix-report.json"
        write_report_json(report, report_path)
        return report.to_dict()

    def _invoke(self, script_name: str, parameters: dict[str, str] | None = None) -> dict[str, Any]:
        script_path = self.repo_root / "automation" / "ps" / script_name
        return invoke_matrix_script(script_path=script_path, parameters=parameters)

    def _run_probe_phase(self) -> dict[str, Any]:
        output_path = self.paths["report_dir"] / "host-probe.json"
        output_path.parent.mkdir(parents=True, exist_ok=True)

        result = self._invoke(
            "Matrix.HostProbe.ps1",
            {"RepoRoot": str(self.repo_root), "OutputPath": str(output_path)},
        )
        if result["exit_code"] != 0:
            return {"pass": False, "error": "host_probe_script_failed", "stderr": result["stderr"]}

        probe = result["payload"]
        evaluation = evaluate_host_probe(probe, self.config.get("hosts", {}))
        return {**probe, **evaluation}

    def _run_registration_phase(self) -> dict[str, Any]:
        cases = self.config.get("registration_cases", [])
        if not cases:
            return {"pass": False, "error": "no_registration_cases"}

        case_results = []
        plugin = self.config.get("plugin", {})
        for case in cases:
            result = self._invoke(
                "Matrix.Register.ps1",
                {
                    "RepoRoot": str(self.repo_root),
                    "Configuration": plugin.get("configuration", "Release"),
                    "RequestedHost": case.get("requested_host", "Both"),
                    "ExecutionIntent": case.get("execution_intent", "PreviewOnly"),
                    "CaseId": case.get("case_id", ""),
                    "CaseName": case.get("case_name", ""),
                },
            )
            payload = result["payload"] if result["exit_code"] == 0 else {"pass": False, "stderr": result["stderr"]}
            case_results.append(payload)

        return {
            "pass": all(item.get("pass", False) for item in case_results),
            "cases": case_results,
        }

    def _run_verify_registration_phase(self) -> dict[str, Any]:
        plugin = self.config.get("plugin", {})
        hosts = self.config.get("hosts", {})
        result = self._invoke(
            "Matrix.VerifyRegistration.ps1",
            {
                "RepoRoot": str(self.repo_root),
                "Configuration": plugin.get("configuration", "Release"),
                "ProgId": plugin.get("prog_id", "WordTools.ThisAddIn"),
                "ExpectedWordBitness": hosts.get("word", {}).get("expected_bitness", "64"),
            },
        )
        if result["exit_code"] != 0:
            return {"pass": False, "error": "verify_registration_failed", "stderr": result["stderr"]}
        return result["payload"]

    def _run_smoke_phase(self) -> dict[str, Any]:
        from lib.expectations import evaluate_smoke_case

        smoke = self.config.get("smoke", {})
        assets = self.config.get("test_assets", {})
        result = self._invoke(
            "Matrix.SmokeInsertImage.ps1",
            {
                "RepoRoot": str(self.repo_root),
                "Mode": smoke.get("mode", "discovery"),
                "HostTarget": smoke.get("host_target", "Both"),
                "ImagePath": assets.get("image", ""),
                "ScreenshotDir": str(self.paths["screenshot_dir"]),
                "ProgId": self.config.get("plugin", {}).get("prog_id", "WordTools.ThisAddIn"),
            },
        )
        if result["exit_code"] != 0:
            return {"pass": False, "error": "smoke_failed", "stderr": result["stderr"]}

        payload = result["payload"]
        evaluation = evaluate_smoke_case(
            word=payload.get("word") or {},
            wps=payload.get("wps") or {},
            hosts=self.config.get("hosts", {}),
            mode=smoke.get("mode", "discovery"),
            host_target=smoke.get("host_target", "Both"),
        )
        return {**payload, **evaluation}

    def _run_batch_insert_phase(self) -> dict[str, Any]:
        from lib.expectations import evaluate_batch_insert_case

        batch_insert = self.config.get("batch_insert", {})
        cases = batch_insert.get("cases", [])
        if not cases:
            return {"pass": False, "error": "no_batch_insert_cases"}

        visible = str(batch_insert.get("visible", "false"))
        plugin = self.config.get("plugin", {})
        configuration = plugin.get("configuration", "Release")
        case_results = []

        for case in cases:
            case_id = case.get("case_id", "")
            result = self._invoke(
                "Matrix.BatchInsert.ps1",
                {
                    "RepoRoot": str(self.repo_root),
                    "CaseId": case_id,
                    "Visible": visible,
                    "Configuration": configuration,
                },
            )
            if result["exit_code"] != 0:
                payload = result["payload"] if result["payload"] else {"pass": False, "stderr": result["stderr"]}
                evaluation = evaluate_batch_insert_case(case_id, {**payload, "pass": False})
                case_results.append({**payload, **evaluation, "case_name": case.get("case_name", "")})
                continue

            payload = result["payload"]
            evaluation = evaluate_batch_insert_case(case_id, payload)
            case_results.append({**payload, **evaluation, "case_name": case.get("case_name", "")})

        return {
            "pass": all(item.get("pass", False) for item in case_results),
            "cases": case_results,
        }

    def _run_unregister_phase(self) -> dict[str, Any]:
        plugin = self.config.get("plugin", {})
        result = self._invoke(
            "Matrix.Unregister.ps1",
            {
                "RepoRoot": str(self.repo_root),
                "Configuration": plugin.get("configuration", "Release"),
                "RequestedHost": self.config.get("unregister", {}).get("requested_host", "Word"),
                "ExecutionIntent": self.config.get("unregister", {}).get("execution_intent", "PreviewOnly"),
            },
        )
        if result["exit_code"] != 0:
            return {"pass": False, "error": "unregister_failed", "stderr": result["stderr"]}
        return result["payload"]

    def _run_verify_cleanup_phase(self) -> dict[str, Any]:
        plugin = self.config.get("plugin", {})
        hosts = self.config.get("hosts", {})
        result = self._invoke(
            "Matrix.VerifyCleanup.ps1",
            {
                "RepoRoot": str(self.repo_root),
                "ProgId": plugin.get("prog_id", "WordTools.ThisAddIn"),
                "ExpectedWordBitness": hosts.get("word", {}).get("expected_bitness", "64"),
            },
        )
        if result["exit_code"] != 0:
            return {"pass": False, "error": "verify_cleanup_failed", "stderr": result["stderr"]}

        payload = result["payload"]
        evaluation = evaluate_cleanup_case(
            case_id="REG-032",
            case_name="全部卸载",
            word_clean=bool(payload.get("word_clean")),
            wps_clean=bool(payload.get("wps_clean")),
        )
        return {**payload, **evaluation}
