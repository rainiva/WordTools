from lib.cli import print_json_report


def test_print_json_report_writes_utf8_payload(capsys):
    print_json_report({"env_name": "word32_wps64", "pass": True})
    captured = capsys.readouterr()
    assert "word32_wps64" in captured.out
