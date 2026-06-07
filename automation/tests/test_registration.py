from lib.expectations import evaluate_registration_case


def test_evaluate_registration_case_fails_when_registry_not_ok():
    outcome = evaluate_registration_case(
        case_id="REG-030",
        case_name="只卸载 Word 插件",
        word_detected=True,
        wps_detected=True,
        word_register_success=True,
        wps_register_success=True,
        word_registry_ok=False,
        wps_registry_ok=True,
        dll_bitness_match=True,
    )

    assert outcome["pass"] is False


def test_evaluate_cleanup_case_passes_when_all_hosts_clean():
    from lib.expectations import evaluate_cleanup_case

    outcome = evaluate_cleanup_case(
        case_id="REG-032",
        case_name="全部卸载",
        word_clean=True,
        wps_clean=True,
    )

    assert outcome["pass"] is True
    assert outcome["word_clean"] is True
