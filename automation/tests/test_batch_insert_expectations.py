from lib.expectations import evaluate_batch_insert_case


def test_evaluate_ac_b03_selected_four_passes():
    result = evaluate_batch_insert_case(
        "AC-B03",
        {
            "pass": True,
            "inline_shape_count": 4,
            "success_count": 4,
            "fail_count": 0,
            "has_numbered_description": True,
        },
    )
    assert result["pass"] is True
    assert result["case_id"] == "AC-B03"


def test_evaluate_ac_b03_fails_when_shape_count_wrong():
    result = evaluate_batch_insert_case(
        "AC-B03",
        {
            "pass": True,
            "inline_shape_count": 3,
            "success_count": 4,
            "fail_count": 0,
            "has_numbered_description": True,
        },
    )
    assert result["pass"] is False


def test_evaluate_ac_b05_single_image_na_passes():
    result = evaluate_batch_insert_case(
        "AC-B05",
        {
            "pass": True,
            "inline_shape_count": 1,
            "success_count": 1,
            "fail_count": 0,
            "last_image_row_col2_text": "N/A",
        },
    )
    assert result["pass"] is True


def test_evaluate_ac_b01_warning_only_passes():
    result = evaluate_batch_insert_case(
        "AC-B01",
        {
            "pass": True,
            "inline_shape_count": 0,
            "warnings": ["请先选中一个表格！"],
        },
    )
    assert result["pass"] is True


def test_evaluate_ac_b06_cancel_passes():
    result = evaluate_batch_insert_case(
        "AC-B06",
        {
            "pass": True,
            "cancelled": True,
            "success_count": 2,
            "warnings": ["操作已取消"],
        },
    )
    assert result["pass"] is True
