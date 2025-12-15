from __future__ import annotations

import os
from pathlib import Path
from typing import Any


def _normalize_key(key: Any) -> str:
    return str(key).strip().replace("-", "_")


def _ensure_list(value: Any) -> list[Any]:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    if isinstance(value, tuple):
        return list(value)
    return [value]


def load_yaml_config(path: str | None) -> dict[str, Any]:
    """
    Load YAML config file and return normalized defaults dict for CLI.

    - Keys are normalized to snake_case (hyphen -> underscore).
    - Relative paths are resolved relative to the config file directory.
    - Strings expand env vars and ~.
    """
    if not path:
        return {}

    config_path = Path(path).expanduser().resolve()
    raw_text = config_path.read_text(encoding="utf-8")

    try:
        import yaml
    except Exception as e:  # pragma: no cover
        raise RuntimeError("YAML config requires 'PyYAML'. Install dependencies first.") from e

    raw = yaml.safe_load(raw_text) or {}
    if not isinstance(raw, dict):
        raise ValueError("Config root must be a mapping (YAML dict)")

    if "clearml_dataset_excel" in raw and isinstance(raw["clearml_dataset_excel"], dict):
        raw = raw["clearml_dataset_excel"]

    cfg: dict[str, Any] = {_normalize_key(k): v for k, v in raw.items()}

    # Expand strings
    for k, v in list(cfg.items()):
        if isinstance(v, str):
            cfg[k] = os.path.expanduser(os.path.expandvars(v))

    # Backward compatible aliases
    if "excel" in cfg and "manifest" not in cfg:
        cfg["manifest"] = cfg["excel"]

    config_dir = config_path.parent

    # Resolve local paths
    for key in ("manifest", "base_dir"):
        v = cfg.get(key)
        if isinstance(v, str) and v:
            p = Path(v).expanduser()
            if not p.is_absolute():
                p = (config_dir / p).resolve()
            else:
                p = p.resolve()
            cfg[key] = p.as_posix()

    # Normalize lists
    for key in ("tags", "parent", "include", "exclude"):
        if key in cfg:
            cfg[key] = [str(x) for x in _ensure_list(cfg.get(key))]

    # Fix file:// output_uri if relative
    output_uri = cfg.get("output_uri")
    if isinstance(output_uri, str) and output_uri.startswith("file://"):
        uri_path = output_uri[len("file://") :]
        if uri_path and not uri_path.startswith("/"):
            abs_path = (config_dir / Path(uri_path)).expanduser().resolve()
            cfg["output_uri"] = "file://" + abs_path.as_posix()

    return cfg

