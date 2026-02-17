#!/usr/bin/env python3
"""
PowerPoint COM Automation for PDF Conversion
Uses Microsoft PowerPoint's native export engine for best quality and compression
Windows-only module
"""

import os
import sys
from pathlib import Path


def is_powerpoint_available():
    """Check if PowerPoint is available via COM"""
    if sys.platform != 'win32':
        return False

    try:
        import win32com.client
        return True
    except ImportError:
        return False


def check_powerpoint_installation():
    """Check if PowerPoint is actually installed and working"""
    if not is_powerpoint_available():
        return False

    try:
        import win32com.client
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Quit()
        return True
    except Exception:
        return False


class PowerPointConverter:
    """
    PowerPoint COM automation for PDF conversion
    Provides exact PowerPoint quality and compression
    """

    # PowerPoint PDF export quality constants
    QUALITY_SETTINGS = {
        'screen': {
            'name': 'Screen/Web (PowerPoint Standard)',
            'ppFixedFormatIntent': 1,  # ppFixedFormatIntentScreen
            'description': 'Optimized for screen viewing, smallest files'
        },
        'standard': {
            'name': 'Standard Quality',
            'ppFixedFormatIntent': 1,  # ppFixedFormatIntentScreen
            'description': 'Balanced quality for most uses'
        },
        'high': {
            'name': 'Print Quality',
            'ppFixedFormatIntent': 2,  # ppFixedFormatIntentPrint
            'description': 'High quality for printing'
        },
        'maximum': {
            'name': 'Maximum Quality',
            'ppFixedFormatIntent': 2,  # ppFixedFormatIntentPrint
            'description': 'Highest quality, larger files'
        }
    }

    def __init__(self, quality='screen'):
        """
        Initialize PowerPoint converter

        Args:
            quality: Quality preset (screen, standard, high, maximum)
        """
        if not is_powerpoint_available():
            raise ImportError("PowerPoint COM automation requires pywin32 package")

        self.quality = quality if quality in self.QUALITY_SETTINGS else 'screen'
        self.powerpoint = None

    def _start_powerpoint(self):
        """Start PowerPoint application"""
        if self.powerpoint is None:
            import win32com.client
            self.powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            # Don't show PowerPoint window
            self.powerpoint.Visible = 0

    def _quit_powerpoint(self):
        """Quit PowerPoint application"""
        if self.powerpoint:
            try:
                self.powerpoint.Quit()
            except:
                pass
            self.powerpoint = None

    def convert_file(self, input_file: str, output_dir: str = None, verbose: bool = True):
        """
        Convert a single PPTX file to PDF using PowerPoint

        Args:
            input_file: Path to PPTX file
            output_dir: Output directory (defaults to same as input)
            verbose: Print conversion status

        Returns:
            True if successful, False otherwise
        """
        input_path = Path(input_file)

        if not input_path.exists():
            if verbose:
                print(f"ERROR: File not found: {input_file}")
            return False

        if input_path.suffix.lower() not in ['.pptx', '.ppt']:
            if verbose:
                print(f"WARNING: {input_file} is not a PowerPoint file, skipping...")
            return False

        # Set output directory
        if output_dir:
            out_dir = Path(output_dir)
            out_dir.mkdir(parents=True, exist_ok=True)
        else:
            out_dir = input_path.parent

        output_path = out_dir / f"{input_path.stem}.pdf"

        if verbose:
            file_size = input_path.stat().st_size / (1024 * 1024)  # MB
            print(f"Converting: {input_path.name} ({file_size:.2f} MB)")
            preset = self.QUALITY_SETTINGS[self.quality]
            print(f"Quality: {preset['name']} (PowerPoint COM)")

        presentation = None
        try:
            # Start PowerPoint
            self._start_powerpoint()

            # Open presentation
            # Use absolute path for COM
            abs_input = str(input_path.absolute())
            abs_output = str(output_path.absolute())

            if verbose:
                print(f"Opening presentation in PowerPoint...")

            presentation = self.powerpoint.Presentations.Open(
                abs_input,
                ReadOnly=True,
                Untitled=True,
                WithWindow=False
            )

            # Get quality settings
            preset = self.QUALITY_SETTINGS[self.quality]
            quality_setting = preset['ppFixedFormatIntent']

            if verbose:
                print(f"Exporting to PDF with quality intent: {quality_setting}...")

            # Export to PDF using ExportAsFixedFormat for quality control
            presentation.ExportAsFixedFormat(
                abs_output,
                2,  # ppFixedFormatTypePDF
                quality_setting,  # Intent (1=Screen, 2=Print)
                False,  # FrameSlides
                0,  # HandoutOrder
                0,  # OutputType (0=Slides)
                False,  # PrintHiddenSlides
                None,  # PrintRange
                0,  # RangeType (0=All)
                "",  # SlideShowName
                True,  # IncludeDocProperties
                True,  # KeepIRMSettings
                True,  # DocStructureTags
                True,  # BitmapMissingFonts
                True   # UseISO19005_1 (PDF/A)
            )

            # Check if PDF was created
            if output_path.exists():
                pdf_size = output_path.stat().st_size / (1024 * 1024)
                if verbose:
                    print(f"OK Success: {output_path.name} ({pdf_size:.2f} MB)")
                return True
            else:
                if verbose:
                    print(f"X Failed: PDF not created for {input_path.name}")
                return False

        except Exception as e:
            import traceback
            if verbose:
                print(f"X Error converting {input_path.name}:")
                print(f"  Error type: {type(e).__name__}")
                print(f"  Error message: {str(e)}")
                print(f"  Full traceback:")
                traceback.print_exc()
            return False

        finally:
            # Always close presentation
            if presentation:
                try:
                    presentation.Close()
                except:
                    pass

    def __del__(self):
        """Cleanup: Quit PowerPoint when object is destroyed"""
        self._quit_powerpoint()


def install_pywin32():
    """Helper function to install pywin32 package"""
    import subprocess

    print("PowerPoint COM automation requires pywin32 package.")
    print("Installing pywin32...")

    try:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pywin32'])
        print("✓ pywin32 installed successfully!")
        print("Please restart the application.")
        return True
    except subprocess.CalledProcessError:
        print("✗ Failed to install pywin32")
        print("Please install manually: pip install pywin32")
        return False


if __name__ == '__main__':
    # Test if PowerPoint is available
    if check_powerpoint_installation():
        print("OK: PowerPoint COM automation is available")
    else:
        print("ERROR: PowerPoint COM automation is NOT available")
        if not is_powerpoint_available():
            print("  Reason: pywin32 package not installed")
            print("  Install: pip install pywin32")
        else:
            print("  Reason: PowerPoint may not be installed or not accessible")
