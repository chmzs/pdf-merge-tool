"""Configuration file handling.

Location: %APPDATA%/PDFToolkit/config.toml (override the directory with PDFTOOLKIT_CONFIG_DIR).
Keys are stored once and shared between the GUI and the CLI. Reading uses tomllib; writing
uses a tiny serializer, so the tool has no dependency beyond the standard library.
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import tomllib

APP_NAME = "PDFToolkit"

DEFAULTS = {
    "general": {
        "out_dir": "",
        "name_template": "{main}",
        "open_output_folder": True,
        "keep_merged": False,
        "max_parallel": 2,
        "theme": "light",
    },
    "compress": {
        "enabled": True,
        "level": "low",
        "backend": "ilovepdf",
    },
    "ilovepdf": {
        "public_key": "",
        "secret_key": "",
        "proxy": "",
        "timeout_seconds": 300,
        "sdk_path": "",
    },
    "docx": {
        "convert": True,
        "word_single_session": True,
    },
}

VALID_LEVELS = ("low", "recommended", "extreme")


def config_path() -> Path:
    base = os.environ.get("PDFTOOLKIT_CONFIG_DIR")
    if not base:
        base = os.path.join(os.environ.get("APPDATA") or str(Path.home()), APP_NAME)
    return Path(base) / "config.toml"


def legacy_key_files() -> list[Path]:
    return [Path(os.path.expanduser("~")) / ".ilovepdf.json"]


def _merge_defaults(raw: dict) -> dict:
    merged = {}
    for section, values in DEFAULTS.items():
        src = raw.get(section) if isinstance(raw.get(section), dict) else {}
        merged[section] = dict(values)
        for key, default in values.items():
            if key in src and isinstance(src[key], type(default)):
                merged[section][key] = src[key]
    return merged


def _format_value(value) -> str:
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return repr(value)
    if isinstance(value, str):
        return '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'
    raise TypeError(f"unsupported config value: {type(value)!r}")


def dumps(config: dict) -> str:
    lines = [
        "# pdf-merge-tool configuration (generated, safe to edit by hand)",
        "# Windows paths keep their backslashes; the writer escapes them for you",
        "",
    ]
    for section, values in config.items():
        lines.append(f"[{section}]")
        for key, value in values.items():
            lines.append(f"{key} = {_format_value(value)}")
        lines.append("")
    return "\n".join(lines)


def migrate_keys(config: dict) -> bool:
    """Import keys from the older ~/.ilovepdf.json once."""
    if config["ilovepdf"]["public_key"] and config["ilovepdf"]["secret_key"]:
        return False
    for legacy in legacy_key_files():
        try:
            data = json.loads(legacy.read_text(encoding="utf-8"))
        except Exception:
            continue
        public_key = str(data.get("public_key", "")).strip()
        secret_key = str(data.get("secret_key", "")).strip()
        if public_key and secret_key:
            config["ilovepdf"]["public_key"] = public_key
            config["ilovepdf"]["secret_key"] = secret_key
            return True
    return False


def load(create: bool = True) -> dict:
    path = config_path()
    raw = {}
    if path.exists():
        try:
            with open(path, "rb") as fh:
                raw = tomllib.load(fh)
        except Exception as exc:
            print(f"[config] unreadable, falling back to defaults: {exc}", file=sys.stderr)
    config = _merge_defaults(raw)
    changed = migrate_keys(config)
    if create and (changed or not path.exists()):
        save(config)
    return config


def save(config: dict) -> Path:
    path = config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(dumps(config), encoding="utf-8")
    return path


def validate(config: dict) -> list[str]:
    problems = []
    if config["compress"]["level"] not in VALID_LEVELS:
        problems.append("compress.level must be one of " + "/".join(VALID_LEVELS))
    if config["compress"]["backend"] != "ilovepdf":
        problems.append("compress.backend only supports ilovepdf")
    if not 1 <= int(config["general"]["max_parallel"]) <= 8:
        problems.append("general.max_parallel must be within 1-8")
    return problems


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    config = load()
    if "--path" in argv:
        print(config_path())
        return 0
    if "--show" in argv:
        print(dumps(config))
        return 0
    if "--validate" in argv:
        problems = validate(config)
        if not config["ilovepdf"]["public_key"] or not config["ilovepdf"]["secret_key"]:
            problems.append("ilovepdf public_key/secret_key missing")
        print(f"config: {config_path()}")
        print("FAIL: " + "; ".join(problems) if problems else "OK")
        return 1 if problems else 0
    print(__doc__)
    return 0


if __name__ == "__main__":
    sys.exit(main())
