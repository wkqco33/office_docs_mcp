from __future__ import annotations

import json
import os
import tomllib
from collections.abc import Mapping
from pathlib import Path
from typing import Any

from wconfig import user_config_dir

DEFAULT_CONFIG: dict[str, Any] = {
    "app": {
        "name": "office_docs_mcp",
    },
    "logging": {
        "level": "INFO",
    },
}


def get_platform_config_path(
    platform: str | None = None,
    environ: dict[str, str] | None = None,
    home: Path | None = None,
) -> Path:
    """Return the platform-specific default configuration file path.

    - Linux/Other: $XDG_CONFIG_HOME/office_docs_mcp/config.toml (or ~/.config/office_docs_mcp/config.toml)
    - macOS: ~/Library/Application Support/office_docs_mcp/config.toml
    - Windows: %APPDATA%/office_docs_mcp/config.toml
    """
    return (
        user_config_dir("office_docs_mcp", platform=platform, environ=environ, home=home)
        / "config.toml"
    )


def dump_toml_value(val: Any) -> str | None:
    """Format a single value into TOML representation."""
    if val is None:
        return None
    if isinstance(val, bool):
        return "true" if val else "false"
    if isinstance(val, (int, float)):
        return str(val)
    if isinstance(val, str):
        escaped = val.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")
        return f'"{escaped}"'
    if isinstance(val, list):
        items = [dump_toml_value(v) for v in val]
        valid_items = [i for i in items if i is not None]
        return "[" + ", ".join(valid_items) + "]"
    return f'"{val}"'


def dump_toml(data: Mapping[str, Any], prefix: str = "") -> str:
    """Serialize a dictionary mapping into a clean TOML formatted string."""
    lines: list[str] = []

    # 1. Output top-level scalar values first
    scalars = {k: v for k, v in data.items() if not isinstance(v, (dict, Mapping))}
    for k, v in scalars.items():
        v_str = dump_toml_value(v)
        if v_str is not None:
            lines.append(f"{k} = {v_str}")

    # 2. Output table sections
    tables = {k: v for k, v in data.items() if isinstance(v, (dict, Mapping))}
    for k, table in tables.items():
        sec_name = f"{prefix}.{k}" if prefix else k
        if lines and lines[-1] != "":
            lines.append("")
        lines.append(f"[{sec_name}]")

        sub_scalars = {sk: sv for sk, sv in table.items() if not isinstance(sv, (dict, Mapping))}
        for sk, sv in sub_scalars.items():
            sv_str = dump_toml_value(sv)
            if sv_str is not None:
                lines.append(f"{sk} = {sv_str}")

        sub_tables = {sk: sv for sk, sv in table.items() if isinstance(sv, (dict, Mapping))}
        for sub_k, sub_v in sub_tables.items():
            sub_res = dump_toml({sub_k: sub_v}, prefix=sec_name)
            if sub_res.strip():
                if lines and lines[-1] != "":
                    lines.append("")
                lines.append(sub_res.strip())

    return "\n".join(lines).strip() + "\n"


def parse_config_value(raw: str) -> Any:
    """Parse a string argument into a Python typed value (bool, int, float, json, or str)."""
    raw_s = raw.strip()
    lower = raw_s.lower()
    if lower == "true":
        return True
    if lower == "false":
        return False
    if lower in ("none", "null"):
        return None
    try:
        return int(raw_s)
    except ValueError:
        pass
    try:
        return float(raw_s)
    except ValueError:
        pass
    if (raw_s.startswith("[") and raw_s.endswith("]")) or (
        raw_s.startswith("{") and raw_s.endswith("}")
    ):
        try:
            return json.loads(raw_s)
        except json.JSONDecodeError:
            pass
    if (raw_s.startswith('"') and raw_s.endswith('"')) or (
        raw_s.startswith("'") and raw_s.endswith("'")
    ):
        return raw_s[1:-1]
    return raw_s


def write_toml_file(path: Path, data: Mapping[str, Any]) -> None:
    """Atomically write data to a TOML file."""
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    content = dump_toml(data)
    temp_file = path.with_suffix(f".tmp.{os.getpid()}")
    try:
        temp_file.write_text(content, encoding="utf-8")
        temp_file.replace(path)
    except Exception:
        if temp_file.exists():
            temp_file.unlink(missing_ok=True)
        raise


def resolve_config_path(
    flag_path: str | None = None,
    root_config_flag: str | None = None,
    local: bool = False,
) -> Path:
    """Resolve the effective configuration file path.

    Priority:
      1. Explicit flag path (--path / -p)
      2. Root persistent config flag (--config / -c)
      3. Local flag (--local / -l -> ./config.toml)
      4. Platform-standard default configuration path (e.g. ~/.config/office_docs_mcp/config.toml)
    """
    if flag_path:
        return Path(flag_path).resolve()
    if root_config_flag:
        return Path(root_config_flag).resolve()
    if local:
        return Path("config.toml").resolve()
    return get_platform_config_path()


def init_config(
    path: str | Path | None = None,
    force: bool = False,
    template: dict[str, Any] | None = None,
) -> Path:
    """Initialize a new config.toml file with default template values."""
    target = Path(path).resolve() if path is not None else get_platform_config_path()
    if target.exists() and not force:
        raise FileExistsError(
            f"Configuration file already exists at '{target}'. Use --force to overwrite."
        )
    write_toml_file(target, template or DEFAULT_CONFIG)
    return target


def load_raw_config(path: str | Path | None = None) -> dict[str, Any]:
    """Load configuration dictionary from a TOML file."""
    target = Path(path).resolve() if path is not None else get_platform_config_path()
    if not target.exists():
        return {}
    content = target.read_text(encoding="utf-8")
    return tomllib.loads(content)


def get_nested_key(data: Mapping[str, Any], key: str) -> Any:
    """Retrieve a nested value using a dotted key path."""
    parts = [p.strip() for p in key.split(".") if p.strip()]
    if not parts:
        raise KeyError("Empty configuration key")
    curr: Any = data
    for p in parts:
        if not isinstance(curr, Mapping) or p not in curr:
            raise KeyError(f"Configuration key '{key}' not found.")
        curr = curr[p]
    return curr


def set_nested_key(data: dict[str, Any], key: str, value: Any) -> None:
    """Set a nested value in a dictionary using a dotted key path."""
    parts = [p.strip() for p in key.split(".") if p.strip()]
    if not parts:
        raise ValueError("Empty configuration key")
    curr = data
    for p in parts[:-1]:
        if p not in curr or not isinstance(curr[p], dict):
            curr[p] = {}
        curr = curr[p]
    curr[parts[-1]] = value


def update_config_file(
    path: str | Path | None = None,
    key: str = "",
    value_raw: str = "",
) -> tuple[Any, Path]:
    """Update a key-value pair in a TOML configuration file."""
    target = Path(path).resolve() if path is not None else get_platform_config_path()
    if target.exists():
        try:
            data = load_raw_config(target)
        except Exception:
            data = dict(DEFAULT_CONFIG)
    else:
        data = dict(DEFAULT_CONFIG)

    parsed_value = parse_config_value(value_raw)
    set_nested_key(data, key, parsed_value)
    write_toml_file(target, data)
    return parsed_value, target
