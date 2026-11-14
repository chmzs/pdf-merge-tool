"""Smoke tests: no network, no Word required."""

from __future__ import annotations

from pathlib import Path

import pytest
from pypdf import PdfReader, PdfWriter

from pdfmerge import config as pdfconfig
from pdfmerge.engine import run
from pdfmerge.merge import merge_pdfs


def make_pdf(path: Path, pages: int = 1) -> Path:
    writer = PdfWriter()
    for _ in range(pages):
        writer.add_blank_page(width=200, height=200)
    with open(path, "wb") as handle:
        writer.write(handle)
    return path


def test_merge_counts_pages(tmp_path: Path):
    first = make_pdf(tmp_path / "a.pdf", pages=2)
    second = make_pdf(tmp_path / "b.pdf", pages=3)
    out = tmp_path / "merged.pdf"
    assert merge_pdfs([first, second], out) == 5
    assert len(PdfReader(str(out)).pages) == 5


def test_config_roundtrip(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("PDFTOOLKIT_CONFIG_DIR", str(tmp_path))
    cfg = pdfconfig.load()
    cfg["general"]["max_parallel"] = 3
    path = pdfconfig.save(cfg)
    assert Path(path).exists()
    assert pdfconfig.load()["general"]["max_parallel"] == 3
    assert pdfconfig.validate(pdfconfig.load()) == []


def test_engine_merge_only(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("PDFTOOLKIT_CONFIG_DIR", str(tmp_path))
    cfg = pdfconfig.load()
    first = make_pdf(tmp_path / "main.pdf", pages=1)
    second = make_pdf(tmp_path / "si.pdf", pages=2)
    out_dir = tmp_path / "out"
    stages: list[str] = []
    result = run(
        files=[first, second],
        out_dir=out_dir,
        do_compress=False,
        config=cfg,
        log=lambda _message: None,
        stage=lambda name, detail: stages.append(name),
    )
    assert result.output is not None and result.output.exists()
    assert len(PdfReader(str(result.output)).pages) == 3
    assert stages == ["merge"]
