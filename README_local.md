# PPTX to PDF Bulk Converter

High-quality bulk converter for PowerPoint files (PPTX/PPT) to PDF format. Designed to handle large files (450MB+) while preserving formatting and layout.

## Features

- **Handles Large Files**: Tested with 450MB+ presentations
- **Batch Processing**: Convert multiple files or entire directories
- **Format Preservation**: Maintains layouts, fonts, images, and formatting
- **Offline Operation**: No internet connection required
- **Cross-Platform**: Works on Windows, Linux, and macOS
- **Free & Open Source**: No licensing costs or API limits

## Prerequisites

### 1. Install LibreOffice

LibreOffice is required for the conversion engine.

**Windows:**
1. Download from: https://www.libreoffice.org/download/download/
2. Run the installer
3. Default installation path: `C:\Program Files\LibreOffice\program\soffice.exe`

**Linux:**
```bash
sudo apt-get update
sudo apt-get install libreoffice
```

**macOS:**
```bash
brew install --cask libreoffice
```

### 2. Python 3.7+

Python is already installed on your system (Python 3.14.0).

## Installation

No additional Python packages required! The script uses only standard library modules.

## Usage

### Option 1: Graphical User Interface (GUI) - Recommended

The easiest way to use the converter is through the graphical interface:

```bash
python converter_gui.py
```

The GUI provides three simple options:
1. **Convert Single File** - Select one PPTX file to convert
2. **Convert Multiple Files** - Select multiple PPTX files to convert in batch
3. **Convert Entire Folder** - Convert all PPTX files in a selected folder

Features:
- Simple point-and-click interface
- Real-time conversion progress
- Status log with success/failure indicators
- Optional custom output directory
- LibreOffice status check on startup

### Option 2: Command-Line Interface

For automation and scripting, use the command-line tool:

#### Convert a Single File

```bash
python convert_pptx_to_pdf.py presentation.pptx
```

#### Convert All PPTX Files in a Directory

```bash
python convert_pptx_to_pdf.py "C:\Users\squesada\Documents\Presentations"
```

#### Convert Multiple Files

```bash
python convert_pptx_to_pdf.py file1.pptx file2.pptx file3.pptx
```

#### Specify Output Directory

```bash
python convert_pptx_to_pdf.py presentation.pptx -o "C:\Users\squesada\Output"
```

#### Convert with Custom LibreOffice Path

```bash
python convert_pptx_to_pdf.py presentation.pptx --libreoffice "C:\Program Files\LibreOffice\program\soffice.exe"
```

#### Quiet Mode (Less Verbose)

```bash
python convert_pptx_to_pdf.py presentation.pptx -q
```

## Command-Line Options

```
positional arguments:
  inputs                PPTX file(s) or directory containing PPTX files

optional arguments:
  -h, --help            Show help message and exit
  -o OUTPUT, --output OUTPUT
                        Output directory for PDF files (default: same as input)
  --libreoffice PATH    Custom path to LibreOffice executable
  -q, --quiet           Suppress verbose output
```

## Examples

### Example 1: Convert Single Large File
```bash
cd C:\Users\squesada\pptx-to-pdf-converter
python convert_pptx_to_pdf.py "C:\Users\squesada\Documents\large_presentation.pptx"
```

### Example 2: Batch Convert Entire Folder
```bash
python convert_pptx_to_pdf.py "C:\Users\squesada\Documents\Presentations" -o "C:\Users\squesada\Documents\PDFs"
```

### Example 3: Convert Multiple Specific Files
```bash
python convert_pptx_to_pdf.py report1.pptx report2.pptx slides.pptx -o ./output
```

## Features in Detail

### Large File Support
- Handles files up to 450MB+ without issues
- 10-minute timeout per file for very large conversions
- Memory-efficient headless LibreOffice processing

### Format Preservation
- Maintains original slide layouts
- Preserves fonts and typography
- Keeps images at original quality
- Retains shapes, charts, and diagrams
- Animations are rendered as static images

### Batch Processing
- Recursively finds all PPTX files in directories
- Progress tracking for multiple files
- Summary report at completion
- Continues processing even if individual files fail

## Troubleshooting

### "LibreOffice not found" Error

If you see this error, do one of the following:

1. **Install LibreOffice** using the links above
2. **Specify the path manually**:
   ```bash
   python convert_pptx_to_pdf.py presentation.pptx --libreoffice "C:\Program Files\LibreOffice\program\soffice.exe"
   ```

### Conversion Takes Too Long

For very large files (450MB+):
- The script has a 10-minute timeout per file
- This is normal for complex presentations
- Watch the console for progress messages

### File Not Converting Properly

1. Open the file in LibreOffice Impress to verify it opens correctly
2. Check if the file is corrupted
3. Ensure you have enough disk space for the PDF output
4. Try converting with LibreOffice GUI first to diagnose issues

### Permission Errors

Run your terminal/command prompt as Administrator (Windows) or use `sudo` (Linux/Mac) if you encounter permission issues.

## Performance

Typical conversion times (approximate):
- Small files (< 10MB): 5-15 seconds
- Medium files (10-100MB): 30-90 seconds
- Large files (100-450MB): 2-10 minutes

Performance depends on:
- File size
- Number of slides
- Image quality and quantity
- System resources (CPU, RAM)

## Technical Details

- **Engine**: LibreOffice (headless mode)
- **Format Support**: .pptx, .ppt, .PPTX, .PPT
- **Output Format**: PDF
- **Dependencies**: Python 3.7+, LibreOffice 7.0+
- **Platform**: Windows, Linux, macOS

## License

This script is provided as-is for bulk PowerPoint to PDF conversion. LibreOffice is licensed under MPL 2.0.

## Support

For issues or questions:
1. Ensure LibreOffice is properly installed
2. Check that your PPTX files open correctly in LibreOffice Impress
3. Verify Python 3.7+ is installed: `python --version`

## Quick Start Guide

1. **Install LibreOffice** (if not already installed)
   - Download: https://www.libreoffice.org/download/download/

2. **Launch the GUI**
   ```bash
   cd C:\Users\squesada\pptx-to-pdf-converter
   python converter_gui.py
   ```

3. **Start Converting**
   - Click one of the three conversion options
   - Select your PPTX file(s) or folder
   - Watch the progress and enjoy!

**Alternative:** Use command-line for automation:
```bash
python convert_pptx_to_pdf.py file.pptx
```

Enjoy fast, high-quality PPTX to PDF conversions!
