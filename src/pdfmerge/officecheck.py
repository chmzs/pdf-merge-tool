"""Detect an installed Word-compatible COM converter.

Microsoft Word and WPS Office expose near-identical COM object models, so one PowerShell
routine drives either. The registry probe decides which ProgIDs to try before a task is
spent; the actual conversion falls through to the next candidate if the first one is broken.
"""

from __future__ import annotations

import sys

# Order matters: a genuine Microsoft Word install wins over WPS's compatibility shim.
CANDIDATES: tuple[tuple[str, str], ...] = (
    ("Word.Application", "Microsoft Word"),
    ("KWPS.Application", "WPS Office"),
)


def _progid_registered(progid: str) -> bool:
    try:
        import winreg
    except ImportError:  # pragma: no cover - non-Windows
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + "\\CLSID"):
            return True
    except OSError:
        return False


def detect_converters() -> list[tuple[str, str]]:
    """Installed converters as [(ProgID, display name)], most preferred first."""
    if sys.platform != "win32":
        return []
    return [(progid, label) for progid, label in CANDIDATES if _progid_registered(progid)]


def converter_labels() -> str:
    """Human readable summary for logs and error messages."""
    found = detect_converters()
    return ", ".join(label for _progid, label in found) if found else "none"
