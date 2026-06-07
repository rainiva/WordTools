import pytest

from lib.register_core import parse_register_core_result


def test_parse_register_core_result_marks_word_success_from_live_execution():
    core_result = {
        "RegisterExecution": {
            "AppliedTargetCount": 1,
            "Targets": [
                {
                    "HostRuleSummary": {"HostName": "Word", "HostBitness": "x64"},
                    "RegAsmResult": {"ExitCode": 0, "WouldRun": True},
                    "RegistryResult": {"ValuesWritten": ["LoadBehavior"]},
                }
            ],
        }
    }

    parsed = parse_register_core_result(core_result, execution_intent="Live")

    assert parsed["word_register_success"] is True
    assert parsed["word_registry_ok"] is True
    assert parsed["pass"] is True


def test_parse_register_core_result_marks_wps_skipped_in_preview():
    core_result = {
        "RegisterPlan": {
            "RegisterPreviewSummary": {
                "PreviewableTargetCount": 1,
            }
        },
        "RegisterExecution": None,
    }

    parsed = parse_register_core_result(core_result, execution_intent="PreviewOnly")

    assert parsed["word_register_success"] is True
    assert parsed["wps_register_success"] is False
    assert parsed["pass"] is True
