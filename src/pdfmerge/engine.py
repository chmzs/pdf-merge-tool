"""Pipeline: convert docx -> merge -> compress.

Stage notifications go through the stage callback so the CLI can print them and the GUI
can move its progress bar without parsing log text.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from pdfmerge import config as pdfconfig
from pdfmerge.compress import compress
from pdfmerge.docx2pdf import DocxConvertError, convert_batch
from pdfmerge.docx2pdf import supported as docx_supported
from pdfmerge.merge import merge_pdfs
from pdfmerge.officecheck import detect_converters

KNOWN_SUFFIXES = (".pdf", ".docx", ".doc")


@dataclass
class Result:
    output: Path | None = None
    warnings: list[str] = field(default_factory=list)


def run(
    files: list[Path],
    out_dir: Path | None = None,
    do_compress: bool = True,
    level: str | None = None,
    keep_merged: bool | None = None,
    config: dict | None = None,
    log=print,
    stage=None,
) -> Result:
    config = config or pdfconfig.load()
    level = level or config["compress"]["level"]
    if keep_merged is None:
        keep_merged = bool(config["general"]["keep_merged"])
    result = Result()

    def notify(name: str, detail: str = "") -> None:
        if stage:
            stage(name, detail)

    docx_inputs = [p for p in files if p.exists() and docx_supported(p)]
    converters = detect_converters()
    if docx_inputs and not converters:
        raise DocxConvertError(
            "Neither Microsoft Word nor WPS Office was detected, so DOCX cannot be converted. "
            "Install one of them, save the document as PDF first, or pass PDFs only."
        )
    for path in files:
        if not path.exists():
            log(f"  [skip] missing: {path}")
            result.warnings.append(f"missing: {path}")
        elif path.suffix.lower() not in KNOWN_SUFFIXES:
            log(f"  [skip] unsupported type: {path.name}")
            result.warnings.append(f"unsupported: {path.name}")

    converted: dict[str, Path] = {}
    if docx_inputs:
        notify("convert", f"0/{len(docx_inputs)}")
        force_saveas = bool(config["docx"]["use_saveas"])
        convert_timeout = float(config["docx"]["timeout_seconds"])
        failures: list[Exception] = []
        for progid, label in converters:
            log(f"  converting {len(docx_inputs)} docx via {label}")
            try:
                converted, failed = convert_batch(
                    docx_inputs,
                    converter=progid,
                    force_saveas=force_saveas,
                    timeout=convert_timeout,
                    log=log,
                )
            except DocxConvertError as exc:
                log(f"  {label} did not work: {exc}")
                failures.append(exc)
                converted, failed = {}, []
                continue
            for source, _message in failed:
                result.warnings.append(f"docx conversion failed: {Path(source).name}")
            break
        else:
            raise failures[-1]
        notify("convert", f"{len(converted)}/{len(docx_inputs)}")

    pdfs: list[Path] = []
    for path in files:
        if not path.exists():
            continue
        suffix = path.suffix.lower()
        if suffix == ".pdf":
            pdfs.append(path.resolve())
        elif docx_supported(path) and path.stem in converted:
            pdfs.append(converted[path.stem])

    if not pdfs:
        raise ValueError("no usable PDF input")

    target_dir = Path(out_dir).resolve() if out_dir else pdfs[0].parent
    target_dir.mkdir(parents=True, exist_ok=True)
    base = Path(files[0]).stem

    if len(pdfs) == 1:
        if not do_compress:
            log(f"  single file, compression disabled: {pdfs[0]}")
            result.output = pdfs[0]
            return result
        out = target_dir / f"{base}_compression.pdf"
        notify("compress", pdfs[0].name)
        compress(pdfs[0], out, level, config, log=log)
        result.output = out
        return result

    merged = target_dir / f"{base}_merged.pdf"
    notify("merge", f"{len(pdfs)} files")
    log(f"  merging {len(pdfs)} files -> {merged.name}")
    pages = merge_pdfs(pdfs, merged)
    log(f"  merged {pages} pages")
    if not do_compress:
        result.output = merged
        return result

    out = target_dir / f"{base}_merged_compression.pdf"
    notify("compress", merged.name)
    compress(merged, out, level, config, log=log)
    result.output = out
    if not keep_merged:
        try:
            merged.unlink()
        except OSError:
            pass
    return result
