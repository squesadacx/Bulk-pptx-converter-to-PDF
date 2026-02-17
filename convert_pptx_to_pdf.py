#!/usr/bin/env python3
"""
Bulk PPTX to PDF Converter using LibreOffice
Handles large files (450MB+) with high-quality format preservation
"""

import subprocess
import sys
import os
from pathlib import Path
import argparse
from typing import List

class PPTXtoPDFConverter:
    def __init__(self, libreoffice_path=None):
        """
        Initialize converter with LibreOffice path

        Args:
            libreoffice_path: Custom path to LibreOffice executable (optional)
        """
        self.libreoffice_path = self._find_libreoffice(libreoffice_path)

    def _find_libreoffice(self, custom_path=None):
        """Find LibreOffice executable on the system"""
        if custom_path and os.path.exists(custom_path):
            return custom_path

        # Common LibreOffice paths on Windows
        common_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            r"C:\Program Files\LibreOffice 7\program\soffice.exe",
            r"C:\Program Files\LibreOffice 24\program\soffice.exe",
        ]

        for path in common_paths:
            if os.path.exists(path):
                return path

        # Try command line (for Linux/Mac or if in PATH)
        for cmd in ['soffice', 'libreoffice']:
            try:
                result = subprocess.run([cmd, '--version'],
                                       capture_output=True,
                                       text=True,
                                       timeout=5)
                if result.returncode == 0:
                    return cmd
            except (subprocess.TimeoutExpired, FileNotFoundError):
                continue

        return None

    def convert_file(self, input_file: str, output_dir: str = None, verbose: bool = True):
        """
        Convert a single PPTX file to PDF

        Args:
            input_file: Path to PPTX file
            output_dir: Output directory (defaults to same as input)
            verbose: Print conversion status

        Returns:
            True if successful, False otherwise
        """
        if not self.libreoffice_path:
            print("ERROR: LibreOffice not found. Please install LibreOffice first.")
            return False

        input_path = Path(input_file)

        if not input_path.exists():
            print(f"ERROR: File not found: {input_file}")
            return False

        if input_path.suffix.lower() not in ['.pptx', '.ppt']:
            print(f"WARNING: {input_file} is not a PowerPoint file, skipping...")
            return False

        # Set output directory
        if output_dir:
            out_dir = Path(output_dir)
            out_dir.mkdir(parents=True, exist_ok=True)
        else:
            out_dir = input_path.parent

        if verbose:
            file_size = input_path.stat().st_size / (1024 * 1024)  # MB
            print(f"Converting: {input_path.name} ({file_size:.2f} MB)")

        try:
            # LibreOffice command for conversion
            cmd = [
                str(self.libreoffice_path),
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', str(out_dir),
                str(input_path)
            ]

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=600  # 10 minute timeout for large files
            )

            if result.returncode == 0:
                pdf_path = out_dir / f"{input_path.stem}.pdf"
                if pdf_path.exists():
                    pdf_size = pdf_path.stat().st_size / (1024 * 1024)
                    if verbose:
                        print(f"✓ Success: {pdf_path.name} ({pdf_size:.2f} MB)")
                    return True
                else:
                    print(f"✗ Failed: PDF not created for {input_path.name}")
                    return False
            else:
                print(f"✗ Failed: {input_path.name}")
                if verbose and result.stderr:
                    print(f"  Error: {result.stderr}")
                return False

        except subprocess.TimeoutExpired:
            print(f"✗ Timeout: {input_path.name} (took longer than 10 minutes)")
            return False
        except Exception as e:
            print(f"✗ Error converting {input_path.name}: {str(e)}")
            return False

    def convert_batch(self, input_paths: List[str], output_dir: str = None, verbose: bool = True):
        """
        Convert multiple PPTX files to PDF

        Args:
            input_paths: List of file paths or directories
            output_dir: Output directory (defaults to same as input)
            verbose: Print conversion status

        Returns:
            Dictionary with success/failure counts
        """
        files_to_convert = []

        # Collect all PPTX files
        for path in input_paths:
            p = Path(path)
            if p.is_file() and p.suffix.lower() in ['.pptx', '.ppt']:
                files_to_convert.append(p)
            elif p.is_dir():
                # Recursively find all PPTX files in directory
                files_to_convert.extend(p.rglob('*.pptx'))
                files_to_convert.extend(p.rglob('*.ppt'))
                files_to_convert.extend(p.rglob('*.PPTX'))
                files_to_convert.extend(p.rglob('*.PPT'))

        if not files_to_convert:
            print("No PPTX files found to convert.")
            return {'total': 0, 'success': 0, 'failed': 0}

        print(f"\nFound {len(files_to_convert)} file(s) to convert\n")
        print("=" * 60)

        success_count = 0
        failed_count = 0

        for i, file_path in enumerate(files_to_convert, 1):
            print(f"\n[{i}/{len(files_to_convert)}]")
            if self.convert_file(str(file_path), output_dir, verbose):
                success_count += 1
            else:
                failed_count += 1

        print("\n" + "=" * 60)
        print(f"\nConversion Summary:")
        print(f"  Total files: {len(files_to_convert)}")
        print(f"  Successful: {success_count}")
        print(f"  Failed: {failed_count}")

        return {
            'total': len(files_to_convert),
            'success': success_count,
            'failed': failed_count
        }


def main():
    parser = argparse.ArgumentParser(
        description='Bulk convert PPTX files to PDF using LibreOffice',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Convert single file
  python convert_pptx_to_pdf.py presentation.pptx

  # Convert all PPTX in a directory
  python convert_pptx_to_pdf.py /path/to/presentations/

  # Convert multiple files with custom output directory
  python convert_pptx_to_pdf.py file1.pptx file2.pptx -o ./output/

  # Convert with custom LibreOffice path
  python convert_pptx_to_pdf.py presentation.pptx --libreoffice "C:/Program Files/LibreOffice/program/soffice.exe"
        """
    )

    parser.add_argument('inputs', nargs='+', help='PPTX file(s) or directory containing PPTX files')
    parser.add_argument('-o', '--output', help='Output directory for PDF files (default: same as input)')
    parser.add_argument('--libreoffice', help='Custom path to LibreOffice executable')
    parser.add_argument('-q', '--quiet', action='store_true', help='Suppress verbose output')

    args = parser.parse_args()

    # Initialize converter
    converter = PPTXtoPDFConverter(args.libreoffice)

    if not converter.libreoffice_path:
        print("\n" + "=" * 60)
        print("ERROR: LibreOffice not found!")
        print("=" * 60)
        print("\nPlease install LibreOffice:")
        print("  Windows: https://www.libreoffice.org/download/download/")
        print("  Linux: sudo apt-get install libreoffice")
        print("  Mac: brew install --cask libreoffice")
        print("\nOr specify the path with --libreoffice flag")
        sys.exit(1)

    print(f"Using LibreOffice: {converter.libreoffice_path}")

    # Convert files
    results = converter.convert_batch(
        args.inputs,
        args.output,
        verbose=not args.quiet
    )

    # Exit with error code if any failed
    sys.exit(0 if results['failed'] == 0 else 1)


if __name__ == '__main__':
    main()
