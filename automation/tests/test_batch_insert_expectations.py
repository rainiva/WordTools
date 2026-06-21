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


def test_evaluate_ac_b07_folder_root_only_passes():
    result = evaluate_batch_insert_case(
        "AC-B07",
        {
            "pass": True,
            "inline_shape_count": 3,
            "success_count": 3,
            "has_subfolder_title": False,
        },
    )
    assert result["pass"] is True


def test_evaluate_ac_b08_folder_sub_only_passes():
    result = evaluate_batch_insert_case(
        "AC-B08",
        {
            "pass": True,
            "inline_shape_count": 2,
            "success_count": 2,
            "has_subfolder_title": True,
        },
    )
    assert result["pass"] is True


def test_evaluate_ac_b09_no_description_passes():
    result = evaluate_batch_insert_case(
        "AC-B09",
        {
            "pass": True,
            "inline_shape_count": 4,
            "success_count": 4,
            "has_numbered_description": False,
        },
    )
    assert result["pass"] is True


def test_evaluate_ac_b10_number_after_center_passes():
    result = evaluate_batch_insert_case(
        "AC-B10",
        {
            "pass": True,
            "inline_shape_count": 4,
            "success_count": 4,
            "has_number_after_description": True,
            "has_center_aligned_numbered_description": True,
        },
    )
    assert result["pass"] is True


def test_evaluate_ac_b11_folder_name_description_passes():
    result = evaluate_batch_insert_case(
        "AC-B11",
        {
            "pass": True,
            "inline_shape_count": 5,
            "success_count": 5,
            "has_folder_name_description": True,
        },
    )
    assert result["pass"] is True


def test_evaluate_ac_b12_manual_description_numbered_passes():
    result = evaluate_batch_insert_case(
        "AC-B12",
        {
            "pass": True,
            "inline_shape_count": 4,
            "success_count": 4,
            "has_numbered_description": True,
            "has_manual_description_rows": True,
        },
    )
    assert result["pass"] is True
