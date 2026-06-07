import pytest

from lib.expectations import evaluate_host_probe, evaluate_registration_case


def test_evaluate_host_probe_passes_when_detected_hosts_match_config():
    probe = {
        "word": {"installed": True, "bitness": "32", "path": r"C:\Program Files (x86)\...\WINWORD.EXE"},
        "wps": {"installed": True, "bitness": "64", "path": r"C:\Program Files\Kingsoft\...\wps.exe"},
    }
    expected = {
        "word": {"enabled": True, "expected_bitness": "32"},
        "wps": {"enabled": True, "expected_bitness": "64"},
    }

    result = evaluate_host_probe(probe, expected)

    assert result["pass"] is True
    assert result["word_detected"] is True
    assert result["wps_detected"] is True
    assert result["word_bitness_ok"] is True
    assert result["wps_bitness_ok"] is True


def test_evaluate_host_probe_fails_when_expected_host_missing():
    probe = {
        "word": {"installed": True, "bitness": "32", "path": r"C:\...\WINWORD.EXE"},
        "wps": {"installed": False, "bitness": None, "path": None},
    }
    expected = {
        "word": {"enabled": True, "expected_bitness": "32"},
        "wps": {"enabled": True, "expected_bitness": "64"},
    }

    result = evaluate_host_probe(probe, expected)

    assert result["pass"] is False
    assert result["wps_detected"] is False


def test_evaluate_host_probe_fails_when_bitness_mismatch():
    probe = {
        "word": {"installed": True, "bitness": "64", "path": r"C:\...\WINWORD.EXE"},
        "wps": {"installed": True, "bitness": "64", "path": r"C:\...\wps.exe"},
    }
    expected = {
        "word": {"enabled": True, "expected_bitness": "32"},
        "wps": {"enabled": True, "expected_bitness": "64"},
    }

    result = evaluate_host_probe(probe, expected)

    assert result["pass"] is False
    assert result["word_bitness_ok"] is False


def test_evaluate_registration_case_builds_structured_result():
    outcome = evaluate_registration_case(
        case_id="REG-022",
        case_name="同时注册 Word32 + WPS64",
        word_detected=True,
        wps_detected=True,
        word_register_success=True,
        wps_register_success=True,
        word_registry_ok=True,
        wps_registry_ok=True,
        dll_bitness_match=True,
    )

    assert outcome["case_id"] == "REG-022"
    assert outcome["pass"] is True
    assert outcome["word_register_success"] is True
    assert outcome["wps_register_success"] is True
