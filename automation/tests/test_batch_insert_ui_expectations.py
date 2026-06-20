from lib.expectations import evaluate_batch_insert_ui_case


def test_evaluate_ac_ui_b03_passes():
    result = evaluate_batch_insert_ui_case(
        "AC-UI-B03",
        {
            "pass": True,
            "ui_flow_started": True,
            "form_clicked": True,
            "progress_seen": True,
            "inline_shape_count": 4,
        },
    )
    assert result["pass"] is True
