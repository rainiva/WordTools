from lib.expectations import evaluate_batch_insert_ui_case


def _ui_payload(**overrides):
    base = {
        "pass": True,
        "ui_flow_started": True,
        "form_clicked": True,
        "progress_seen": True,
    }
    base.update(overrides)
    return base


def test_evaluate_ac_ui_b03_passes():
    result = evaluate_batch_insert_ui_case(
        "AC-UI-B03",
        _ui_payload(
            inline_shape_count=4,
            expected_image_count=4,
            has_numbered_description=True,
            table_row_count=12,
            min_table_row_count=8,
        ),
    )
    assert result["pass"] is True


def test_evaluate_ac_ui_b05_no_numbering_passes():
    result = evaluate_batch_insert_ui_case(
        "AC-UI-B05",
        _ui_payload(
            inline_shape_count=1,
            expected_image_count=1,
            has_numbered_description=False,
        ),
    )
    assert result["pass"] is True


def test_evaluate_ac_ui_b07_root_only_passes():
    result = evaluate_batch_insert_ui_case(
        "AC-UI-B07",
        _ui_payload(
            inline_shape_count=4,
            expected_image_count=4,
            has_subfolder_title=False,
        ),
    )
    assert result["pass"] is True


def test_evaluate_ac_ui_b08_subfolder_title_passes():
    result = evaluate_batch_insert_ui_case(
        "AC-UI-B08",
        _ui_payload(
            inline_shape_count=25,
            expected_image_count=25,
            has_subfolder_title=True,
        ),
    )
    assert result["pass"] is True


def test_evaluate_ac_ui_b09_no_description_passes():
    result = evaluate_batch_insert_ui_case(
        "AC-UI-B09",
        _ui_payload(
            inline_shape_count=4,
            expected_image_count=4,
            has_numbered_description=False,
        ),
    )
    assert result["pass"] is True


def test_evaluate_ac_ui_b10_number_after_center_passes():
    result = evaluate_batch_insert_ui_case(
        "AC-UI-B10",
        _ui_payload(
            inline_shape_count=4,
            expected_image_count=4,
            has_number_after_description=True,
            has_center_aligned_numbered_description=True,
        ),
    )
    assert result["pass"] is True


def test_evaluate_ac_ui_b11_folder_name_passes():
    result = evaluate_batch_insert_ui_case(
        "AC-UI-B11",
        _ui_payload(
            inline_shape_count=25,
            expected_image_count=25,
            has_folder_name_description=True,
        ),
    )
    assert result["pass"] is True


def test_evaluate_ac_ui_b12_manual_description_passes():
    result = evaluate_batch_insert_ui_case(
        "AC-UI-B12",
        _ui_payload(
            inline_shape_count=4,
            expected_image_count=4,
            has_numbered_description=True,
            has_manual_description_rows=True,
        ),
    )
    assert result["pass"] is True


def test_evaluate_ac_ui_b14_number_before_left_passes():
    result = evaluate_batch_insert_ui_case(
        "AC-UI-B14",
        _ui_payload(
            inline_shape_count=4,
            expected_image_count=4,
            has_numbered_description=True,
            has_left_aligned_numbered_description=True,
        ),
    )
    assert result["pass"] is True
