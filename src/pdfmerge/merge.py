"""PDF merging (pypdf)."""

from __future__ import annotations

from pathlib import Path

from pypdf import PdfReader, PdfWriter


def merge_pdfs(paths: list[Path], out_path: Path) -> int:
    """Append every page of every input into one PDF. Returns the page count."""
    writer = PdfWriter()
    pages = 0
    for path in paths:
        reader = PdfReader(str(path))
        for page in reader.pages:
            writer.add_page(page)
        pages += len(reader.pages)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "wb") as fh:
        writer.write(fh)
    return pages
