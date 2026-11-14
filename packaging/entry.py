"""PyInstaller entry point: behaves like "python -m pdfmerge"."""
import sys

from pdfmerge.__main__ import main

if __name__ == "__main__":
    sys.exit(main())
