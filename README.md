# pdf-merge-tool

Merge a main PDF with its supplementary material (or any pile of PDFs/DOCX files) into a
single PDF, then shrink it with [iLovePDF](https://www.ilovepdf.com). Windows GUI + CLI,
one small dependency set, no local build step required to use it.

主要用途：把文献正文与补充材料（SI）合成一个 PDF 并压缩到便于阅读/上传的体积。

## Features

- **Merge then compress** — DOCX is converted to PDF through Microsoft Word **or WPS Office** (auto-detected, one application session per batch); PDFs are concatenated with pypdf; the result is compressed by the iLovePDF API (`low` by default).
- **Batch DOCX conversion in one Word session** — N documents cost one Word start instead of N.
- **GUI** — file table with a "main" file (double-click to set; it is merged first and names the output), reorder, live streamed log, progress bar, cancel (kills the child process tree), failed rows stay for retry, settings page.
- **Task queue** — configurable concurrency (default 2; iLovePDF rate-limits aggressive use).
- **Config file** — `%APPDATA%\PDFToolkit\config.toml`, keys can also come from an existing `~/.ilovepdf.json` (migrated once).
- **CLI** — same engine, scriptable: `pdf-merge main.pdf si.docx --out out`.

## Install

### Windows binary

Download `pdf-merge-tool-win64.zip` from [Releases](https://github.com/chmzs/pdf-merge-tool/releases),
unzip anywhere, run `pdf-merge-tool.exe`. No Python needed.

### From PyPI dependencies / source

```powershell
pip install git+https://github.com/chmzs/pdf-merge-tool
pdf-merge-gui            # GUI
pdf-merge a.pdf b.docx   # CLI
```

Requires Python 3.10+. On Windows only for the GUI and DOCX conversion (DOCX needs Microsoft Word).

## iLovePDF keys

No key ships with this tool — a shared key would be extractable from any public binary, so every
user brings their own (free tier is enough for personal use).

1. Register at <https://developer.ilovepdf.com/user/projects> and create a project.
2. Open the GUI → 设置 → click **获取免费密钥** (or paste the two keys manually).
3. `config.toml` lives at `%APPDATA%\PDFToolkit\config.toml`; the CLI reads the same file.

The tool uploads the merged PDF to iLovePDF and downloads the result — nothing else leaves the
machine.

```toml
[ilovepdf]
public_key = "project_public_..."
secret_key  = "secret_key_..."
proxy       = ""      # e.g. http://127.0.0.1:7890 when the API is unreachable directly
```

## CLI

```powershell
pdf-merge main.pdf si.docx --out D:\papers\hu2023          # merge + compress (level low)
pdf-merge main.pdf si.docx --no-compress                    # merge only
pdf-merge scanned.pdf --level extreme                       # single file: compress only
pdf-merge --help
```

Output naming (base name comes from the first input):

| case | output |
|---|---|
| single input + compress | `<name>_compression.pdf` |
| multiple inputs, no compress | `<name>_merged.pdf` |
| multiple inputs + compress | `<name>_merged_compression.pdf` |

Progress is printed as `STAGE:convert|merge|compress:<detail>` so wrappers can parse it.

## Development

```powershell
pip install -e ".[dev]"
ruff check src tests
pytest -q
python -m pdfmerge          # GUI from source
```

Builds run in GitHub Actions only: `.github/workflows/ci.yml` lints and tests,
`.github/workflows/build.yml` freezes the exe on `windows-latest` and attaches it to a
release when a `v*` tag is pushed. Nothing needs to be built locally.

## Limitations

- DOCX conversion needs Windows plus Microsoft Word or WPS Office (`Word.Application`, then `KWPS.Application`). If both are missing the task fails before any upload, so no quota is wasted. LibreOffice is not supported.
- Compression is a cloud call: quality/size is whatever iLovePDF returns for the chosen level.
- pypdf merging is single-threaded; merging is fast, the network upload dominates.

## License

MIT
