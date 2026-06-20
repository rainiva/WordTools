import pytest

from lib.config import validate_matrix_config


def test_validate_matrix_config_accepts_word64_bitness():
    config = {
        "env_name": "word64_wps32",
        "hosts": {
            "word": {"enabled": True, "expected_bitness": "64"},
            "wps": {"enabled": True, "expected_bitness": "32"},
        },
    }

    issues = validate_matrix_config(config)

    assert issues == []


def test_validate_matrix_config_rejects_invalid_word_bitness():
    config = {
        "env_name": "bad",
        "hosts": {
            "word": {"enabled": True, "expected_bitness": "128"},
        },
    }

    issues = validate_matrix_config(config)

    assert any("word.expected_bitness" in issue for issue in issues)
