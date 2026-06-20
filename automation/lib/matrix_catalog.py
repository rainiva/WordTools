import json
from pathlib import Path
from typing import Any


def load_matrix_index(index_path: Path) -> dict[str, Any]:
    with index_path.open(encoding="utf-8") as handle:
        return json.load(handle)


def list_environments(index_path: Path) -> list[dict[str, Any]]:
    return load_matrix_index(index_path)["environments"]


def get_environment(index: dict[str, Any], env_id: str) -> dict[str, Any]:
    for item in index["environments"]:
        if item["id"] == env_id:
            return item
    raise KeyError(f"Unknown environment id: {env_id}")


def word64_environment_ids(index: dict[str, Any]) -> list[str]:
    return [item["id"] for item in index["environments"] if item.get("word_bitness") == "64"]


def resolve_config_path(automation_root: Path, env_id: str) -> Path:
    index_path = automation_root / "configs" / "matrix_index.json"
    index = load_matrix_index(index_path)
    env = get_environment(index, env_id)
    return automation_root / env["config"]
