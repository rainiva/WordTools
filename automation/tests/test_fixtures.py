from pathlib import Path


def _assets_root() -> Path:
    return Path(__file__).resolve().parents[1] / "assets"


def test_table_template_docx_exists():
    path = _assets_root() / "table-template.docx"
    assert path.is_file(), f"missing fixture: {path}"
    assert path.stat().st_size > 1000


def test_test_small_jpg_exists():
    path = _assets_root() / "test-small.jpg"
    assert path.is_file(), f"missing fixture: {path}"


def test_selected_four_images_exist():
    folder = _assets_root() / "images" / "selected-4"
    names = ["01.jpg", "02.jpg", "03.jpg", "04.jpg"]
    for name in names:
        path = folder / name
        assert path.is_file(), f"missing fixture: {path}"


def test_folder_root_images_exist():
    root = _assets_root() / "images" / "folder-root"
    for name in ["01.jpg", "02.jpg", "03.jpg"]:
        assert (root / name).is_file()
    sub = root / "sub-a"
    for name in ["01.jpg", "02.jpg"]:
        assert (sub / name).is_file()


def test_single_image_fixture_exists():
    path = _assets_root() / "images" / "single" / "01.jpg"
    assert path.is_file(), f"missing fixture: {path}"
