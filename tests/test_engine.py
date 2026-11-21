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


def test_engine_rejects_docx_without_office(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("PDFTOOLKIT_CONFIG_DIR", str(tmp_path))
    from pdfmerge import engine
    from pdfmerge.docx2pdf import DocxConvertError

    monkeypatch.setattr(engine, "detect_converters", lambda: [])
    docx = tmp_path / "supp.docx"
    docx.write_bytes(b"PK\x03\x04placeholder")
    with pytest.raises(DocxConvertError):
        engine.run(files=[docx], do_compress=False, config=pdfconfig.load(), log=lambda _m: None)


def test_check_available_flags_missing_keys(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("PDFTOOLKIT_CONFIG_DIR", str(tmp_path))
    from pdfmerge.compress import check_available

    cfg = pdfconfig.load()
    cfg["ilovepdf"]["public_key"] = ""
    cfg["ilovepdf"]["secret_key"] = ""
    cfg["ilovepdf"]["sdk_path"] = str(tmp_path / "missing")
    problems = check_available(cfg)
    assert any("keys are missing" in problem for problem in problems)


def test_check_available_flags_unimportable_sdk(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("PDFTOOLKIT_CONFIG_DIR", str(tmp_path))
    from pdfmerge import compress as compress_module

    cfg = pdfconfig.load()
    cfg["ilovepdf"]["public_key"] = "public"
    cfg["ilovepdf"]["secret_key"] = "secret"

    def broken(_config):
        raise ImportError("no such module")

    monkeypatch.setattr(compress_module, "_load_sdk", broken)
    problems = compress_module.check_available(cfg)
    assert problems and "pip install ilovepdf" in problems[0]


def test_engine_refuses_to_start_without_compress_prereqs(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("PDFTOOLKIT_CONFIG_DIR", str(tmp_path))
    from pdfmerge import engine
    from pdfmerge.compress import CompressConfigError

    cfg = pdfconfig.load()
    cfg["ilovepdf"]["public_key"] = ""
    cfg["ilovepdf"]["secret_key"] = ""
    monkeypatch.setattr(engine, "check_available", lambda _cfg: ["no keys"])
    pdf = make_pdf(tmp_path / "main.pdf", pages=1)
    with pytest.raises(CompressConfigError):
        engine.run(files=[pdf], do_compress=True, config=cfg, log=lambda _m: None)


def test_convert_batch_times_out_cleanly(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    import subprocess

    from pdfmerge import docx2pdf

    docx = tmp_path / "supp.docx"
    docx.write_bytes(b"PK\x03\x04placeholder")

    def fake_run(*_args, **_kwargs):
        raise subprocess.TimeoutExpired(cmd="powershell", timeout=1)

    monkeypatch.setattr(docx2pdf.subprocess, "run", fake_run)
    with pytest.raises(docx2pdf.DocxConvertError) as excinfo:
        docx2pdf.convert_batch([docx], timeout=1, log=lambda _m: None)
    assert "timed out" in str(excinfo.value)


def test_detect_converters_returns_list():
    from pdfmerge.officecheck import detect_converters

    found = detect_converters()
    assert isinstance(found, list)
    for progid, label in found:
        assert progid and label


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
