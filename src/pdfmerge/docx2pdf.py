"""DOCX -> PDF conversion through a Word-compatible COM interface (Windows only).

Microsoft Word and WPS Office both expose Documents.Open / ExportAsFixedFormat (or SaveAs),
so the same PowerShell routine drives either one. The whole batch shares a single
application instance: starting Word/WPS costs seconds, so converting N files in one session
is the difference between roughly N * 5 s and ~5 s.
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
    [Parameter(Mandatory=$true)][string]$JobFile,
    [string]$Converter = 'Word.Application',
    [switch]$ForceSaveAs
)
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
$ErrorActionPreference = 'Stop'
$jobs = Get-Content -Raw -Encoding UTF8 $JobFile | ConvertFrom-Json
$app = $null
try {
    $app = New-Object -ComObject $Converter
    try { $app.Visible = $false } catch { }
    try { $app.DisplayAlerts = 0 } catch { }
    $appName = ''
    try { $appName = $app.Name } catch { }
    Write-Output ("INFO" + [char]9 + $Converter + [char]9 + $appName)
    foreach ($job in $jobs) {
        $doc = $null
        try {
            $doc = $app.Documents.Open($job.In, $false, $true)
            try {
                foreach ($sec in $doc.Sections) {
                    if ($sec.PageSetup.LineNumbering.Active) { $sec.PageSetup.LineNumbering.Active = 0 }
                }
            } catch { }
            $saved = $false
            if (-not $ForceSaveAs) {
                try {
                    $doc.ExportAsFixedFormat($job.Out, 17, $false, 0)
                    $saved = $true
                } catch { }
            }
            if (-not $saved) {
                try {
                    $doc.SaveAs2($job.Out, 17)
                    $saved = $true
                } catch {
                    $doc.SaveAs([ref]$job.Out, [ref]17)
                    $saved = $true
                }
            }
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
    $reason = $_.Exception.Message.Replace([char]10, " ").Replace([char]13, " ")
    [Console]::Error.WriteLine("FAILED: " + $reason)
    exit 1
}
finally {
    if ($app) {
        $app.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    }
}
"""

NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


class DocxConvertError(RuntimeError):
    """No usable Word/WPS installation, or the conversion process failed."""


def supported(path: Path) -> bool:
    return path.suffix.lower() in WORD_SUFFIXES


def convert_batch(
    paths: list[Path],
    converter: str = "Word.Application",
    force_saveas: bool = False,
    timeout: float = 900,
    log=print,
) -> tuple[dict[str, Path], list[tuple[str, str]]]:
    """Convert every docx in one application session.

    Returns ({original stem: converted pdf path}, [(failed source, message), ...]).
    Raises DocxConvertError when the converter cannot even be started.
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
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        script_file,
        "-JobFile",
        job_file,
        "-Converter",
        converter,
    ]
    if force_saveas:
        command.append("-ForceSaveAs")
    try:
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=NO_WINDOW,
                timeout=timeout,
            )
        except subprocess.TimeoutExpired as exc:
            raise DocxConvertError(
                f"conversion timed out after {timeout:g}s - Word/WPS may be blocked by a "
                "dialog (sign-in or advertisement). Close the popup, or set docx.timeout_seconds."
            ) from exc
        converted: dict[str, Path] = {}
        failed: list[tuple[str, str]] = []
        for line in (result.stdout or "").splitlines():
            parts = line.split("\t")
            if parts[0] == "INFO" and len(parts) >= 3:
                shown = parts[2].strip() or parts[1]
                log(f"  [converter] {parts[1]} -> {shown}")
            elif parts[0] == "OK" and len(parts) >= 2:
                out_path = Path(parts[1])
                converted[out_path.name.replace("_conv.pdf", "")] = out_path
            elif parts[0] == "FAIL" and len(parts) >= 2:
                failed.append((parts[1], parts[2] if len(parts) > 2 else "unknown error"))
        if result.returncode != 0 and not converted and not failed:
            raise DocxConvertError((result.stderr or "conversion process failed").strip()[:400])
        if failed and not converted:
            reasons = "; ".join(message for _source, message in failed[:2])
            raise DocxConvertError(f"all conversions failed ({reasons})"[:400])
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
