from __future__ import annotations

from pathlib import Path
from typing import Any


def load_yaml(path: str | Path) -> dict[str, Any]:
    try:
        import yaml
    except ImportError as exc:
        raise ImportError("YAML 설정을 읽으려면 PyYAML이 필요합니다. pip install PyYAML") from exc

    src = Path(path)
    if not src.exists():
        raise FileNotFoundError(f"설정 파일을 찾을 수 없습니다: {src}")
    with src.open("r", encoding="utf-8") as fp:
        return yaml.safe_load(fp) or {}
