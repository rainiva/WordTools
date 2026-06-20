from pathlib import Path

import pytest

from lib.matrix_catalog import (
    get_environment,
    list_environments,
    load_matrix_index,
    word64_environment_ids,
)


@pytest.fixture
def automation_root() -> Path:
    return Path(__file__).resolve().parents[1]


def test_matrix_index_includes_all_four_bitness_combinations(automation_root: Path):
    index = load_matrix_index(automation_root / "configs" / "matrix_index.json")
    env_ids = {item["id"] for item in index["environments"]}

    assert env_ids == {"VM-01", "VM-02", "VM-03", "VM-04"}


def test_word64_environments_declare_expected_word_bitness_64(automation_root: Path):
    index = load_matrix_index(automation_root / "configs" / "matrix_index.json")
    word64_ids = word64_environment_ids(index)

    assert "VM-03" in word64_ids
    assert "VM-04" in word64_ids

    for env_id in word64_ids:
        env = get_environment(index, env_id)
        assert env["word_bitness"] == "64"


def test_word64_environment_configs_exist_on_disk(automation_root: Path):
    index = load_matrix_index(automation_root / "configs" / "matrix_index.json")

    for env_id in word64_environment_ids(index):
        env = get_environment(index, env_id)
        config_path = automation_root / env["config"]
        assert config_path.exists(), f"missing config for {env_id}: {config_path}"


def test_list_environments_marks_word64_as_supported_or_planned(automation_root: Path):
    environments = list_environments(automation_root / "configs" / "matrix_index.json")
    vm03 = next(item for item in environments if item["id"] == "VM-03")
    vm04 = next(item for item in environments if item["id"] == "VM-04")

    assert vm03["word_support"] == "supported"
    assert vm04["word_support"] == "supported"
