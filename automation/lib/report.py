import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


class MatrixReport:
    def __init__(self, env_name: str):
        self.env_name = env_name
        self.layers: list[dict[str, Any]] = []

    def add_layer(self, name: str, payload: dict[str, Any]) -> None:
        self.layers.append({"name": name, "payload": payload})

    def to_dict(self) -> dict[str, Any]:
        return {
            "env_name": self.env_name,
            "generated_at_utc": datetime.now(timezone.utc).isoformat(),
            "pass": all(layer["payload"].get("pass", False) for layer in self.layers),
            "layers": self.layers,
        }


def write_report_json(report: MatrixReport, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as handle:
        json.dump(report.to_dict(), handle, ensure_ascii=False, indent=2)
