from pathlib import Path

from lib.report import MatrixReport, write_report_json


def test_matrix_report_aggregates_layer_results():
    report = MatrixReport(env_name="word32_wps64")
    report.add_layer("host_probe", {"pass": True, "word_detected": True})
    report.add_layer("registration", {"pass": False, "case_id": "REG-022"})

    payload = report.to_dict()

    assert payload["env_name"] == "word32_wps64"
    assert payload["pass"] is False
    assert len(payload["layers"]) == 2
    assert payload["layers"][0]["name"] == "host_probe"


def test_write_report_json_creates_file(tmp_path: Path):
    report = MatrixReport(env_name="local")
    report.add_layer("host_probe", {"pass": True})

    output_path = tmp_path / "report.json"
    write_report_json(report, output_path)

    assert output_path.exists()
    assert '"env_name": "local"' in output_path.read_text(encoding="utf-8")
