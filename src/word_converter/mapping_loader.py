"""Load optional JSON mapping overrides for conversion."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


class MappingConfigError(ValueError):
    """Raised when mapping config is invalid."""


def load_mapping_overrides(config_path: str | Path) -> dict[str, dict[str, str]]:
    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {path}")

    data: Any = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise MappingConfigError("Config must be a JSON object.")

    supported = {"table_header_mapping", "cell_code_mapping", "fixed_text_mapping"}
    unknown = set(data.keys()) - supported
    if unknown:
        raise MappingConfigError(f"Unsupported config keys: {', '.join(sorted(unknown))}")

    normalized: dict[str, dict[str, str]] = {}
    for key in supported:
        value = data.get(key, {})
        if not isinstance(value, dict):
            raise MappingConfigError(f"{key} must be an object of string-to-string.")

        casted: dict[str, str] = {}
        for mapping_key, mapping_value in value.items():
            if not isinstance(mapping_key, str) or not isinstance(mapping_value, str):
                raise MappingConfigError(f"{key} keys and values must be strings.")
            casted[mapping_key] = mapping_value

        normalized[key] = casted

    return normalized
