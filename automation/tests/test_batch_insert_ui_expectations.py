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
            "expected_image_count": 4,
            "has_numbered_description": True,
        },
    )
    assert result["pass"] is True


def test_evaluate_ac_ui_b04_real_folder_passes():
    result = evaluate_batch_insert_ui_case(
        "AC-UI-B04",
        {
            "pass": True,
            "ui_flow_started": True,
            "form_clicked": True,
            "progress_seen": True,
            "inline_shape_count": 23,
            "expected_image_count": 23,
        },
    )
    assert result["pass"] is True


def test_evaluate_ac_ui_b05_real_single_passes():
    result = evaluate_batch_insert_ui_case(
        "AC-UI-B05",
        {
            "pass": True,
            "ui_flow_started": True,
            "form_clicked": True,
            "progress_seen": True,
            "inline_shape_count": 1,
            "expected_image_count": 1,
        },
    )
    assert result["pass"] is True
