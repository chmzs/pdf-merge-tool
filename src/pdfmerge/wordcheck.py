"""Microsoft Word availability check (Windows registry, no pywin32 dependency)."""

from __future__ import annotations

import sys


def word_available() -> bool:
    """True when a Word COM ProgID is registered.

    Cheap (sub-millisecond) and dependency-free: we only need to know whether the
    conversion would work before spending a task on it, not whether Word is happy.
    """
    if sys.platform != "win32":
        return False
    try:
        import winreg
    except ImportError:  # pragma: no cover - non-Windows
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application\\CLSID"):
            return True
    except OSError:
        return False
