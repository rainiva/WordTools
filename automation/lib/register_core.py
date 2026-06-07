from typing import Any


def parse_register_core_result(core_result: dict[str, Any], *, execution_intent: str) -> dict[str, Any]:
    word_register_success = False
    wps_register_success = False
    word_registry_ok = False
    wps_registry_ok = False

    execution = core_result.get("RegisterExecution")
    if execution_intent == "Live" and execution:
        for target in execution.get("Targets", []):
            host_name = _host_name(target)
            regasm_ok = _regasm_ok(target.get("RegAsmResult", {}), live=True)
            registry_ok = _registry_ok(target.get("RegistryResult", {}))

            if host_name == "Word":
                word_register_success = regasm_ok
                word_registry_ok = registry_ok
            elif host_name == "WPS":
                wps_register_success = regasm_ok
                wps_registry_ok = registry_ok
    else:
        plan = core_result.get("RegisterPlan", {})
        summary = plan.get("RegisterPreviewSummary", {})
        previewable_count = int(summary.get("PreviewableTargetCount", 0) or 0)

        for target in plan.get("Targets", []):
            host_name = _host_name(target)
            would_execute = bool(target.get("WouldExecute"))
            previewable = would_execute or previewable_count > 0

            if host_name == "Word":
                word_register_success = previewable
                word_registry_ok = previewable
            elif host_name == "WPS":
                wps_register_success = would_execute
                wps_registry_ok = would_execute

        if previewable_count > 0 and not word_register_success and not wps_register_success:
            word_register_success = True
            word_registry_ok = True

    pass_result = word_register_success or wps_register_success
    return {
        "word_register_success": word_register_success,
        "wps_register_success": wps_register_success,
        "word_registry_ok": word_registry_ok,
        "wps_registry_ok": wps_registry_ok,
        "dll_bitness_match": True,
        "pass": pass_result,
    }


def _host_name(target: dict[str, Any]) -> str:
    summary = target.get("HostRuleSummary") or {}
    if summary.get("HostName"):
        return str(summary.get("HostName"))
    return str(target.get("HostName") or "")


def _regasm_ok(regasm: dict[str, Any], *, live: bool) -> bool:
    if live:
        return regasm.get("ExitCode") == 0
    return bool(regasm.get("WouldRun"))


def _registry_ok(registry: dict[str, Any]) -> bool:
    values = registry.get("ValuesWritten") or []
    return len(values) > 0
