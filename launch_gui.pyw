"""Windows launcher: double-click (or shortcut) starts the GUI without a console window."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from pdfmerge.gui import main

if __name__ == "__main__":
    sys.exit(main())
