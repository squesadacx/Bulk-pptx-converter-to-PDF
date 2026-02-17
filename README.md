# Bulk PPTX Converter to PDF

> High-quality bulk PowerPoint to PDF converter with a user-friendly GUI. Designed to handle large files (450MB+) while preserving formatting perfectly.

[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS-blue)]()
[![Python](https://img.shields.io/badge/python-3.7%2B-blue)]()
[![License](https://img.shields.io/badge/license-MIT-green)]()

## ✨ Features

- **🎯 GUI Application** - Simple, intuitive interface with three conversion modes
- **📦 Handles Large Files** - Tested with 450MB+ presentations without issues
- **🎨 Perfect Format Preservation** - Maintains layouts, fonts, images, and formatting
- **⚡ Batch Processing** - Convert single files, multiple files, or entire folders
- **🔒 Offline Operation** - No internet connection required, completely private
- **💰 100% Free** - No licensing costs, API limits, or subscriptions
- **🌍 Cross-Platform** - Works on Windows, Linux, and macOS

## 📸 Screenshots

![GUI Interface](https://via.placeholder.com/700x600.png?text=PPTX+to+PDF+Converter+GUI)

*Simple and clean interface with three conversion options*

## 🚀 Quick Start

### Prerequisites

1. **Python 3.7 or higher**
   ```bash
   python --version
   ```

2. **LibreOffice** (conversion engine)
   - Download: https://www.libreoffice.org/download/download/

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/squesadacx/Bulk-pptx-converter-to-PDF.git
   cd Bulk-pptx-converter-to-PDF
   ```

2. **Install LibreOffice** (if not already installed)

   **Windows:**
   - Download from https://www.libreoffice.org/download/download/
   - Run the installer with default settings

   **Linux (Ubuntu/Debian):**
   ```bash
   sudo apt-get update
   sudo apt-get install libreoffice
   ```

   **macOS:**
   ```bash
   brew install --cask libreoffice
   ```

3. **Launch the converter**
   ```bash
   python converter_gui.py
   ```

   Or double-click `Start Converter.bat` (Windows)

That's it! No Python packages to install - uses only standard library.

## 📖 Usage Guide

### GUI Mode (Recommended)

Launch the graphical interface:
```bash
python converter_gui.py
```

**Three Simple Options:**

1. **📄 Convert Single File**
   - Click the button
   - Select one PPTX file
   - Done!

2. **📑 Convert Multiple Files**
   - Click the button
   - Select multiple files (Ctrl+Click or Shift+Click)
   - All selected files will be converted

3. **📁 Convert Entire Folder**
   - Click the button
   - Select a folder
   - All PPTX files (including subfolders) will be converted

**Optional Settings:**
- **Output Directory**: By default, PDFs are saved next to the original files
- Click "Browse..." to choose a custom output location
- Click "Reset" to return to default behavior

### Command-Line Mode

For automation, scripting, or advanced users:

**Convert a single file:**
```bash
python convert_pptx_to_pdf.py presentation.pptx
```

**Convert all PPTX files in a folder:**
```bash
python convert_pptx_to_pdf.py /path/to/presentations/
```

**Convert multiple specific files:**
```bash
python convert_pptx_to_pdf.py file1.pptx file2.pptx file3.pptx
```

**Specify output directory:**
```bash
python convert_pptx_to_pdf.py presentation.pptx -o /path/to/output/
```

**Custom LibreOffice path:**
```bash
python convert_pptx_to_pdf.py presentation.pptx --libreoffice "C:\Program Files\LibreOffice\program\soffice.exe"
```

**Quiet mode (less verbose):**
```bash
python convert_pptx_to_pdf.py presentation.pptx -q
```

**View all options:**
```bash
python convert_pptx_to_pdf.py --help
```

## 🎯 Use Cases

- **Business Presentations** - Convert sales decks, quarterly reports, training materials
- **Academic Work** - Convert lecture slides, thesis presentations, research posters
- **Archival** - Create PDF archives of PowerPoint presentations
- **Sharing** - Convert presentations to universally viewable PDFs
- **Batch Processing** - Convert entire directories of presentations at once

## ⚙️ How It Works

1. **LibreOffice Engine** - Uses LibreOffice's powerful conversion engine in headless mode
2. **Format Fidelity** - Preserves all formatting, fonts, layouts, images, and charts
3. **Batch Processing** - Processes files sequentially with progress tracking
4. **Thread-Safe GUI** - Background conversion keeps UI responsive
5. **Error Handling** - Continues processing even if individual files fail

## 📊 Performance

Typical conversion times:

| File Size | Conversion Time |
|-----------|----------------|
| Small (< 10MB) | 5-15 seconds |
| Medium (10-100MB) | 30-90 seconds |
| Large (100-450MB) | 2-10 minutes |

Performance depends on:
- File size and number of slides
- Image quality and quantity
- System resources (CPU, RAM)
- Disk I/O speed

## 🛠️ Technical Details

- **Language**: Python 3.7+
- **GUI Framework**: Tkinter (built into Python)
- **Conversion Engine**: LibreOffice 7.0+ (headless mode)
- **Input Formats**: .pptx, .ppt (case-insensitive)
- **Output Format**: PDF
- **Dependencies**: None (uses Python standard library only)
- **Platform**: Windows, Linux, macOS

## 📁 Project Structure

```
Bulk-pptx-converter-to-PDF/
├── converter_gui.py           # GUI application (main entry point)
├── convert_pptx_to_pdf.py     # Command-line tool & conversion engine
├── Start Converter.bat        # Windows launcher (double-click to run)
├── README.md                  # This file
├── GUI_GUIDE.txt              # Quick GUI reference guide
├── INSTALL.txt                # Installation instructions
└── LICENSE                    # MIT License
```

## 🔧 Troubleshooting

### "LibreOffice not found" Error

**Solution 1:** Install LibreOffice from https://www.libreoffice.org/download/download/

**Solution 2:** Specify the path manually:
```bash
python convert_pptx_to_pdf.py presentation.pptx --libreoffice "/path/to/soffice"
```

Common LibreOffice paths:
- Windows: `C:\Program Files\LibreOffice\program\soffice.exe`
- Linux: `/usr/bin/soffice` or `/usr/bin/libreoffice`
- macOS: `/Applications/LibreOffice.app/Contents/MacOS/soffice`

### Conversion Takes Too Long

For very large files (450MB+):
- 10-minute timeout per file is normal
- Watch the progress bar and status log
- Ensure sufficient disk space for output PDFs

### File Not Converting Properly

1. Try opening the PPTX file in LibreOffice Impress first
2. Check if the file is corrupted
3. Ensure enough disk space for output
4. Try converting with LibreOffice GUI to diagnose issues

### GUI Doesn't Launch

1. Check Python version: `python --version` (must be 3.7+)
2. Ensure Tkinter is installed (usually included with Python)
3. On Linux, install: `sudo apt-get install python3-tk`

### Permission Errors

- Windows: Run terminal as Administrator
- Linux/macOS: Use `sudo` or check file permissions

## 🤝 Contributing

Contributions are welcome! Here are some ways you can help:

- 🐛 Report bugs by opening an issue
- 💡 Suggest new features or improvements
- 📝 Improve documentation
- 🔧 Submit pull requests

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **LibreOffice** - Powerful open-source conversion engine
- Built with Python's Tkinter for cross-platform GUI support
- Inspired by the need to handle large PowerPoint files efficiently

## ⚠️ Disclaimer

This tool uses LibreOffice for conversion. Ensure your PPTX files open correctly in LibreOffice Impress for best results. Some advanced PowerPoint features (macros, embedded videos) may not convert perfectly to PDF.

## 📞 Support

- **Issues**: https://github.com/squesadacx/Bulk-pptx-converter-to-PDF/issues
- **Documentation**: See [README.md](README.md) and [GUI_GUIDE.txt](GUI_GUIDE.txt)
- **LibreOffice Help**: https://www.libreoffice.org/get-help/

## 🌟 Star This Repository

If this tool helped you, please consider giving it a star ⭐ on GitHub!

---

**Made with ❤️ for efficient document conversion**
