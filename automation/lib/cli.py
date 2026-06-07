import json
import sys


def print_json_report(report: dict) -> None:
    text = json.dumps(report, ensure_ascii=False, indent=2)
    encoding = sys.stdout.encoding or "utf-8"
    sys.stdout.buffer.write(text.encode(encoding, errors="replace"))
    sys.stdout.buffer.write(b"\n")
