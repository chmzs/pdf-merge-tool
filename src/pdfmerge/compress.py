"""PDF compression through the iLovePDF cloud service (official SDK)."""

from __future__ import annotations

import os
import sys
from pathlib import Path

from pdfmerge.config import ILOVEPDF_KEYS_URL


class CompressConfigError(RuntimeError):
    """Keys or settings are missing."""


class CompressSDKMissing(RuntimeError):
    """The ilovepdf package cannot be imported."""


def _load_sdk(config: dict):
    """Import the SDK, honouring an explicit local path when one is configured."""
    sdk_path = config["ilovepdf"]["sdk_path"].strip()
    if sdk_path:
        sys.path.insert(0, sdk_path)
    from ilovepdf.compress_task import CompressTask

    return CompressTask


def check_available(config: dict) -> list[str]:
    """Everything that makes a compression call fail, checked before any work is spent."""
    problems: list[str] = []
    ilovepdf = config["ilovepdf"]
    if not ilovepdf["public_key"].strip() or not ilovepdf["secret_key"].strip():
        problems.append(
            f"iLovePDF keys are missing - get a free key at {ILOVEPDF_KEYS_URL} "
            "and paste it into the settings page"
        )
    try:
        _load_sdk(config)
    except ImportError:
        problems.append(
            "iLovePDF SDK is not importable - run 'pip install ilovepdf', "
            "or point ilovepdf.sdk_path at a local copy of the package"
        )
    return problems


def compress(inp: Path, out: Path, level: str, config: dict, log=print) -> tuple[int, int]:
    """Upload one PDF, download the compressed result. Returns (in_bytes, out_bytes)."""
    ilovepdf = config["ilovepdf"]
    public_key = ilovepdf["public_key"].strip()
    secret_key = ilovepdf["secret_key"].strip()
    if not public_key or not secret_key:
        raise CompressConfigError(
            "iLovePDF keys are not configured. Get a free key at "
            f"{ILOVEPDF_KEYS_URL} and paste it into the GUI settings page "
            "(or edit the config file: python -m pdfmerge.config --path)."
        )

    os.environ["DEFAULT_TIMEOUT_SECONDS"] = os.environ.get(
        "DEFAULT_TIMEOUT_SECONDS", str(int(ilovepdf["timeout_seconds"]))
    )
    proxy = ilovepdf["proxy"].strip()
    if proxy and os.environ.get("PDFMERGE_IGNORE_PROXY") != "1":
        os.environ.setdefault("HTTP_PROXY", proxy)
        os.environ.setdefault("HTTPS_PROXY", proxy)

    try:
        CompressTask = _load_sdk(config)
    except ImportError as exc:  # pragma: no cover - depends on local install
        raise CompressSDKMissing(f"iLovePDF SDK is not importable: {exc}") from exc

    in_size = inp.stat().st_size
    task = CompressTask(public_key=public_key, secret_key=secret_key)
    task.compression_level = level
    task.add_file(str(inp))
    log(f"[1] task started, file added: {inp.name} ({in_size} bytes)")
    task.execute()
    log(f"[2] process status: {task.result}")
    out.parent.mkdir(parents=True, exist_ok=True)
    task.set_output_filename(out.name)
    task.download(str(out.parent))
    out_size = out.stat().st_size
    log(f"[3] compressed saved: {out} ({out_size} bytes)")
    log(f"    ratio: {100 * out_size / in_size:.0f}% of original")
    return in_size, out_size
