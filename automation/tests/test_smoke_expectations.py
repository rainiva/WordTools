import pytest

from lib.expectations import evaluate_smoke_case


def test_evaluate_smoke_case_passes_in_discovery_mode_when_host_launchable():
    result = evaluate_smoke_case(
        word={"launched": True, "addin_loaded": False},
        wps={"launched": True, "addin_loaded": False},
        hosts={"word": {"enabled": True}, "wps": {"enabled": True}},
        mode="discovery",
    )

    assert result["pass"] is True


def test_evaluate_smoke_case_passes_when_word_and_wps_insert_persist():
    result = evaluate_smoke_case(
        word={"addin_loaded": True, "image_persisted": True},
        wps={"addin_loaded": True, "image_persisted": True},
        hosts={"word": {"enabled": True}, "wps": {"enabled": True}},
        mode="insert_image",
    )

    assert result["pass"] is True
    assert result["word_smoke_ok"] is True
    assert result["wps_smoke_ok"] is True


def test_evaluate_smoke_case_word_only_ignores_wps_when_host_target_is_word():
    result = evaluate_smoke_case(
        word={"addin_loaded": True, "image_persisted": True},
        wps={"addin_loaded": False, "image_persisted": False},
        hosts={"word": {"enabled": True}, "wps": {"enabled": True}},
        mode="com_load",
        host_target="Word",
    )

    assert result["pass"] is True
    assert result["wps_smoke_ok"] is True


def test_evaluate_smoke_case_fails_when_enabled_host_missing_addin():
    result = evaluate_smoke_case(
        word={"addin_loaded": True, "image_persisted": True},
        wps={"addin_loaded": False, "image_persisted": False},
        hosts={"word": {"enabled": True}, "wps": {"enabled": True}},
        mode="insert_image",
    )

    assert result["pass"] is False
    assert result["wps_smoke_ok"] is False
