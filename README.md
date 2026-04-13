# Bulk PPTX to PDF Converter

> Bulk PowerPoint to PDF converter using Microsoft PowerPoint's native export engine. Designed for large files (450MB+) with a GUI and CLI, optimized for RAG ingestion pipelines.

[![Platform](https://img.shields.io/badge/platform-Windows-blue)]()
[![Python](https://img.shields.io/badge/python-3.7%2B-blue)]()
[![License](https://img.shields.io/badge/license-MIT-green)]()

## Features

- **PowerPoint COM engine** — uses PowerPoint's own export, not LibreOffice. Produces the same output as File > Save As > PDF
- **Massive compression** — 93% average reduction across tested files (767 MB → 51 MB total)
- **Reliable batch processing** — per-file PowerPoint restart, automatic retry (x2), hard timeout per file
- **GUI + CLI** — graphical interface for manual use, command-line for automation and scripting
- **Queue processing** — converts one file at a time, logs every result, continues on failure
- **RAG-ready output** — lightweight PDFs with text fully preserved, images at screen resolution

## Requirements

- Windows 10/11
- Microsoft PowerPoint installed (any recent version)
- Python 3.7+
- `pywin32` package

```bash
pip install pywin32
```

## Installation

```bash
git clone https://github.com/squesadacx/Bulk-pptx-converter-to-PDF.git
cd Bulk-pptx-converter-to-PDF
pip install pywin32
```

## Usage

### GUI

```bash
python converter_gui.py
```

- **Convert Single File** — pick one PPTX
- **Convert Multiple Files** — Ctrl+Click to select several
- **Convert Entire Folder** — processes all PPTX files recursively
- Output directory defaults to same folder as input; use Browse to override
- Status log shows per-file result and size in real time

### CLI

```bash
# Convert a single file
python convert_pptx_to_pdf.py presentation.pptx

# Convert an entire folder
python convert_pptx_to_pdf.py "C:\path\to\folder"

# Specify output directory
python convert_pptx_to_pdf.py "C:\path\to\folder" -o "C:\output"

# Force PowerPoint engine explicitly
python convert_pptx_to_pdf.py presentation.pptx --engine powerpoint
```

## Proven Results

Tested on 11 real-world presentations (Microsoft Fabric / FabCon / Ignite session decks):

| File | PPTX | PDF | Reduction |
|---|---|---|---|
| Beyond Monitoring AI Driven Spark... | 241 MB | 5.5 MB | 98% |
| FABCON-SQLCON SQL 2025 Developers... | 193 MB | 6.1 MB | 97% |
| FabCon Atlanta 2026 - OneLake Spark... | 98 MB | 4.3 MB | 96% |
| FCSC26 - Instant insights pipeline... | 117 MB | 8.4 MB | 93% |
| FCSC26 - Building Next-Gen Apps... | 57 MB | 7.7 MB | 86% |
| Best Practices Library Management... | 17.5 MB | 3.2 MB | 82% |
| Adapting to Fabric Spark | 11 MB | 3.2 MB | 71% |
| SQL Server 2025 for DBAs | 9 MB | 3.1 MB | 66% |
| FabricIQ_FoundryIQ_300Level | 8.7 MB | 3.3 MB | 62% |
| Unity in Action - Building Enterprise AI | 8.3 MB | 4.0 MB | 52% |
| MLV_FabCon_Final | 6.9 MB | 3.0 MB | 57% |
| **TOTAL (11 files)** | **768 MB** | **52 MB** | **93%** |

> Files heavy with embedded images compress 95-98%. Text-heavy files compress 50-70%. All above the 60% target threshold.

## How It Works

1. **PowerPoint COM** — opens each PPTX via `win32com` and calls `SaveAs(..., ppSaveAsPDF)`, the same path PowerPoint uses internally for PDF export
2. **Per-file isolation** — a fresh PowerPoint COM instance is created and destroyed for every single file. No shared state between files
3. **Retry logic** — on failure, kills any lingering `POWERPNT.EXE` processes and retries up to 2 times
4. **Hard timeout** — each file has a 10-minute deadline enforced via a background thread. Hung conversions are killed and skipped
5. **Thread-safe GUI** — conversion runs in a background thread; the UI stays responsive and streams log output in real time

## Project Structure

```
pptx-to-pdf-converter/
├── converter_gui.py         # GUI application
├── convert_pptx_to_pdf.py   # CLI tool and batch engine
├── powerpoint_converter.py  # PowerPoint COM core (reliability layer)
├── Start Converter.bat      # Windows double-click launcher
└── README.md
```

## Troubleshooting

**"PowerPoint COM automation is NOT available"**
- Install pywin32: `pip install pywin32`
- Make sure Microsoft PowerPoint is installed and licensed

**"Application.Visible: Invalid request"**
- PowerPoint is already open. Close it before running the converter, or ignore — the tool handles this

**Files fail consistently**
- Run the CLI directly to see the full error: `python convert_pptx_to_pdf.py yourfile.pptx`
- Check the file opens normally in PowerPoint
- Ensure there is enough disk space for the output

**GUI shows "standard" quality but I want smallest files**
- Select "screen - Screen/Web (smallest, like PowerPoint)" in the PDF Quality dropdown
- The screen preset is the default and matches PowerPoint's built-in PDF export quality

## Notes

- Windows only — requires Microsoft PowerPoint (COM automation is not available on macOS/Linux)
- The quality dropdown in the GUI is informational; all presets currently use PowerPoint's default screen-quality export via `SaveAs`, which produces the smallest files
- LibreOffice is no longer used as primary engine — it produced files 10-20x larger than PowerPoint COM for the same input

## License

MIT — see [LICENSE](LICENSE)
