from typing import Any


def evaluate_host_probe(probe: dict[str, Any], expected_hosts: dict[str, Any]) -> dict[str, Any]:
    word = probe.get("word", {})
    wps = probe.get("wps", {})

    word_expected = expected_hosts.get("word", {})
    wps_expected = expected_hosts.get("wps", {})

    word_detected = bool(word.get("installed"))
    wps_detected = bool(wps.get("installed"))

    word_bitness_ok = _bitness_ok(word, word_expected, word_detected)
    wps_bitness_ok = _bitness_ok(wps, wps_expected, wps_detected)

    word_required = bool(word_expected.get("enabled", False))
    wps_required = bool(wps_expected.get("enabled", False))

    pass_result = True
    if word_required and not word_detected:
        pass_result = False
    if wps_required and not wps_detected:
        pass_result = False
    if word_required and word_detected and not word_bitness_ok:
        pass_result = False
    if wps_required and wps_detected and not wps_bitness_ok:
        pass_result = False

    return {
        "pass": pass_result,
        "word_detected": word_detected,
        "wps_detected": wps_detected,
        "word_bitness_ok": word_bitness_ok,
        "wps_bitness_ok": wps_bitness_ok,
    }


def _bitness_ok(host: dict[str, Any], expected: dict[str, Any], detected: bool) -> bool:
    if not expected.get("enabled", False):
        return True
    if not detected:
        return False
    expected_bitness = str(expected.get("expected_bitness", ""))
    actual_bitness = str(host.get("bitness", ""))
    return expected_bitness == actual_bitness


def evaluate_registration_case(
    *,
    case_id: str,
    case_name: str,
    word_detected: bool,
    wps_detected: bool,
    word_register_success: bool,
    wps_register_success: bool,
    word_registry_ok: bool,
    wps_registry_ok: bool,
    dll_bitness_match: bool,
) -> dict[str, Any]:
    pass_result = (
        word_register_success
        and wps_register_success
        and word_registry_ok
        and wps_registry_ok
        and dll_bitness_match
    )

    return {
        "case_id": case_id,
        "case_name": case_name,
        "word_detected": word_detected,
        "wps_detected": wps_detected,
        "word_register_success": word_register_success,
        "wps_register_success": wps_register_success,
        "word_registry_ok": word_registry_ok,
        "wps_registry_ok": wps_registry_ok,
        "dll_bitness_match": dll_bitness_match,
        "pass": pass_result,
    }


def evaluate_cleanup_case(
    *,
    case_id: str,
    case_name: str,
    word_clean: bool,
    wps_clean: bool,
) -> dict[str, Any]:
    pass_result = word_clean and wps_clean
    return {
        "case_id": case_id,
        "case_name": case_name,
        "word_clean": word_clean,
        "wps_clean": wps_clean,
        "pass": pass_result,
    }


def evaluate_smoke_case(
    *,
    word: dict[str, Any],
    wps: dict[str, Any],
    hosts: dict[str, Any],
    mode: str = "discovery",
    host_target: str = "Both",
) -> dict[str, Any]:
    word_required = bool(hosts.get("word", {}).get("enabled", False)) and host_target in {"Word", "Both"}
    wps_required = bool(hosts.get("wps", {}).get("enabled", False)) and host_target in {"WPS", "Both"}

    word_smoke_ok = _host_smoke_ok(word, word_required, mode)
    wps_smoke_ok = _host_smoke_ok(wps, wps_required, mode)

    pass_result = True
    if word_required and not word_smoke_ok:
        pass_result = False
    if wps_required and not wps_smoke_ok:
        pass_result = False

    return {
        "pass": pass_result,
        "word_smoke_ok": word_smoke_ok,
        "wps_smoke_ok": wps_smoke_ok,
    }


def _host_smoke_ok(host_result: dict[str, Any], required: bool, mode: str) -> bool:
    if not required:
        return True

    if mode == "discovery":
        return bool(host_result.get("launched"))

    if mode == "com_load":
        return bool(host_result.get("addin_loaded"))

    return bool(host_result.get("addin_loaded")) and bool(host_result.get("image_persisted"))


def evaluate_batch_insert_case(case_id: str, result: dict[str, Any]) -> dict[str, Any]:
    """Evaluate a single batch-insert E2E case payload from BatchInsertE2E / Matrix.BatchInsert.ps1."""
    base = {
        "case_id": case_id,
        "pass": False,
        "details": [],
    }

    if not result.get("pass", False):
        base["details"].append("host reported pass=false")
        return base

    warnings = [str(item) for item in result.get("warnings", [])]
    inline_shape_count = int(result.get("inline_shape_count", -1))
    success_count = int(result.get("success_count", -1))
    fail_count = int(result.get("fail_count", 0))

    if case_id == "AC-B01":
        ok = inline_shape_count == 0 and any("请先选中一个表格" in w for w in warnings)
        base["pass"] = ok
        if not ok:
            base["details"].append("expected no shapes and table-selection warning")
        return base

    if case_id == "AC-B02":
        ok = inline_shape_count == 0 and any("请将光标置于表格左侧单元格" in w for w in warnings)
        base["pass"] = ok
        if not ok:
            base["details"].append("expected no shapes and first-column warning")
        return base

    if case_id == "AC-B03":
        ok = (
            inline_shape_count == 4
            and success_count == 4
            and fail_count == 0
            and bool(result.get("has_numbered_description"))
        )
        base["pass"] = ok
        if not ok:
            base["details"].append("expected 4 shapes, success=4, numbered descriptions")
        return base

    if case_id == "AC-B04":
        ok = (
            inline_shape_count == 5
            and success_count == 5
            and bool(result.get("has_subfolder_title"))
        )
        base["pass"] = ok
        if not ok:
            base["details"].append("expected 5 shapes, subfolder title row")
        return base

    if case_id == "AC-B05":
        ok = (
            inline_shape_count == 1
            and success_count == 1
            and str(result.get("last_image_row_col2_text", "")).strip() == "N/A"
        )
        base["pass"] = ok
        if not ok:
            base["details"].append("expected 1 shape and col2 N/A")
        return base

    if case_id == "AC-B06":
        ok = bool(result.get("cancelled")) and any("操作已取消" in w for w in warnings)
        base["pass"] = ok
        if not ok:
            base["details"].append("expected cancelled with cancel notification")
        return base

    base["details"].append(f"unknown case_id: {case_id}")
    return base


def evaluate_batch_insert_ui_case(case_id: str, result: dict[str, Any]) -> dict[str, Any]:
    base = {
        "case_id": case_id,
        "pass": False,
        "details": [],
    }

    if not result.get("pass", False):
        base["details"].append("host reported pass=false")
        return base

    expected_count = int(result.get("expected_image_count", 0))
    inline_shape_count = int(result.get("inline_shape_count", 0))
    ui_flow_ok = (
        bool(result.get("ui_flow_started"))
        and bool(result.get("form_clicked"))
        and bool(result.get("progress_seen"))
    )

    if case_id in {"AC-UI-B03", "AC-UI-B04", "AC-UI-B05"}:
        if expected_count <= 0:
            base["details"].append("expected_image_count missing or invalid")
            return base

        ok = ui_flow_ok and inline_shape_count == expected_count
        if case_id == "AC-UI-B03":
            ok = ok and bool(result.get("has_numbered_description"))
            min_rows = int(result.get("min_table_row_count", 8))
            ok = ok and int(result.get("table_row_count", 0)) >= min_rows
        base["pass"] = ok
        if not ok:
            base["details"].append(
                f"expected UI flow + inline_shape_count={expected_count}, got {inline_shape_count}"
            )
        return base

    base["details"].append(f"unknown ui case_id: {case_id}")
    return base
