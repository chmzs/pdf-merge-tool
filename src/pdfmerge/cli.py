"""Command line interface: pdf-merge <files...> [options]."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pdfmerge import __version__
from pdfmerge import config as pdfconfig
from pdfmerge.compress import CompressConfigError, CompressSDKMissing
from pdfmerge.docx2pdf import DocxConvertError
from pdfmerge.engine import run


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pdf-merge",
        description="Merge PDF/DOCX files into one PDF and shrink it with iLovePDF.",
    )
    parser.add_argument("files", nargs="+", type=Path, help="input files, first one names the output")
    parser.add_argument("--out", type=Path, help="output directory (default: first input's folder)")
    parser.add_argument("--level", choices=pdfconfig.VALID_LEVELS, help="compression level")
    parser.add_argument("--no-compress", action="store_true", help="merge only, skip iLovePDF")
    parser.add_argument("--keep-merged", action="store_true", help="keep the intermediate _merged.pdf")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    config = pdfconfig.load()

    def stage(name: str, detail: str = "") -> None:
        print(f"STAGE:{name}:{detail}", flush=True)

    try:
        result = run(
            files=args.files,
            out_dir=args.out,
            do_compress=not args.no_compress,
            level=args.level,
            keep_merged=True if args.keep_merged else None,
            config=config,
            log=lambda message: print(message, flush=True),
            stage=stage,
        )
    except (CompressConfigError, CompressSDKMissing, DocxConvertError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # pragma: no cover - network and SDK errors
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1

    for warning in result.warnings:
        print(f"WARN: {warning}", file=sys.stderr)
    if result.output:
        print(f"DONE: {result.output}", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
