"""DOCX -> PDF conversion through the Microsoft Word COM interface (Windows only).

The whole batch shares one Word instance. Starting Word costs a few seconds, so
converting N files in one session is the difference between roughly N * 5 s and ~5 s.
"""

from __future__ import annotations

import json
import os
import subprocess
import tempfile
from pathlib import Path

WORD_SUFFIXES = (".docx", ".doc")

POWERSHELL_SCRIPT = """
param(
    [Parameter(Mandatory=$true)][string]$JobFile
)
$ErrorActionPreference = 'Stop'
$jobs = Get-Content -Raw -Encoding UTF8 $JobFile | ConvertFrom-Json
$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    foreach ($job in $jobs) {
        $doc = $null
        try {
            $doc = $word.Documents.Open($job.In, $false, $true)
            foreach ($sec in $doc.Sections) {
                if ($sec.PageSetup.LineNumbering.Active) { $sec.PageSetup.LineNumbering.Active = 0 }
            }
            $doc.ExportAsFixedFormat($job.Out, 17, $false, 0)
            Write-Output ("OK" + [char]9 + $job.Out)
        }
        catch {
            $message = $_.Exception.Message.Replace([char]10, " ").Replace([char]9, " ")
            Write-Output ("FAIL" + [char]9 + $job.In + [char]9 + $message)
        }
        finally {
            if ($doc) {
                $doc.Close($false)
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
            }
        }
    }
}
catch {
    Write-Error "FAILED: $($_.Exception.Message)"
    exit 1
}
finally {
    if ($word) { $word.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
}
"""

NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


class DocxConvertError(RuntimeError):
    """Word is missing or the conversion process failed."""


def supported(path: Path) -> bool:
    return path.suffix.lower() in WORD_SUFFIXES


def convert_batch(paths: list[Path], log=print) -> tuple[dict[str, Path], list[tuple[str, str]]]:
    """Convert every docx in one Word session.

    Returns ({original stem: converted pdf path}, [(failed source, message), ...]).
    """
    if not paths:
        return {}, []
    jobs = [
        {"In": str(path.resolve()), "Out": str(path.with_name(path.stem + "_conv.pdf"))} for path in paths
    ]
    handle, job_file = tempfile.mkstemp(prefix="pdfmerge_docx_", suffix=".json")
    os.close(handle)
    script_file = str(Path(job_file).with_suffix(".ps1"))
    Path(script_file).write_text(POWERSHELL_SCRIPT, encoding="utf-8")
    Path(job_file).write_text(json.dumps(jobs, ensure_ascii=False), encoding="utf-8")
    try:
        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                script_file,
                "-JobFile",
                job_file,
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=NO_WINDOW,
        )
        converted: dict[str, Path] = {}
        failed: list[tuple[str, str]] = []
        for line in (result.stdout or "").splitlines():
            parts = line.split("\t")
            if parts[0] == "OK" and len(parts) >= 2:
                out_path = Path(parts[1])
                converted[out_path.name.replace("_conv.pdf", "")] = out_path
            elif parts[0] == "FAIL" and len(parts) >= 2:
                failed.append((parts[1], parts[2] if len(parts) > 2 else "unknown error"))
        if result.returncode != 0 and not converted and not failed:
            raise DocxConvertError((result.stderr or "Word conversion failed").strip()[:400])
        for source, message in failed:
            log(f"  [docx failed] {os.path.basename(source)}: {message}")
        for out_path in converted.values():
            log(f"  [docx->pdf] {out_path.name}")
        return converted, failed
    finally:
        for temp_file in (job_file, script_file):
            try:
                os.unlink(temp_file)
            except OSError:
                pass
