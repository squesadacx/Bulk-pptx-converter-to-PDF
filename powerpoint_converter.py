#!/usr/bin/env python3
"""
PowerPoint COM Automation for PDF Conversion
Uses Microsoft PowerPoint's native export engine for best quality and compression
Windows-only module
"""

import os
import sys
import time
import threading
from pathlib import Path


def is_powerpoint_available():
    """Check if pywin32 is installed"""
    if sys.platform != 'win32':
        return False
    try:
        import win32com.client
        return True
    except ImportError:
        return False


def check_powerpoint_installation():
    """Check if PowerPoint is actually installed and responsive"""
    if not is_powerpoint_available():
        return False
    try:
        import win32com.client
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Quit()
        return True
    except Exception:
        return False


def _kill_powerpoint_processes():
    """Force-kill any lingering PowerPoint processes"""
    try:
        import subprocess
        subprocess.run(
            ['taskkill', '/F', '/IM', 'POWERPNT.EXE'],
            capture_output=True
        )
        time.sleep(2)
    except Exception:
        pass


class PowerPointConverter:
    """
    PowerPoint COM automation for PDF conversion.
    Starts a fresh PowerPoint instance per file for maximum reliability.
    """

    QUALITY_SETTINGS = {
        'screen': {
            'name': 'Screen/Web (smallest files)',
            'ppFixedFormatIntent': 1,  # ppFixedFormatIntentScreen
        },
        'standard': {
            'name': 'Standard Quality',
            'ppFixedFormatIntent': 1,
        },
        'high': {
            'name': 'Print Quality',
            'ppFixedFormatIntent': 2,  # ppFixedFormatIntentPrint
        },
        'maximum': {
            'name': 'Maximum Quality',
            'ppFixedFormatIntent': 2,
        },
    }

    # How long to wait for a single file conversion before giving up (seconds)
    CONVERSION_TIMEOUT = 600  # 10 minutes

    # How many times to retry a failed file before marking it as failed
    MAX_RETRIES = 2

    def __init__(self, quality='screen'):
        if not is_powerpoint_available():
            raise ImportError(
                "PowerPoint COM automation requires pywin32. "
                "Install with: pip install pywin32"
            )
        self.quality = quality if quality in self.QUALITY_SETTINGS else 'screen'

    def _fresh_powerpoint(self):
        """
        Create and return a brand-new PowerPoint COM instance.
        Note: We do NOT set Visible=0 — PowerPoint may already be running and
        setting Visible on an existing instance raises an error.
        The presentation itself is opened with WithWindow=False so no slide
        window appears even if the app is visible.
        """
        import win32com.client
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        return powerpoint

    def _quit_safely(self, powerpoint):
        """Quit a PowerPoint COM instance, ignoring all errors."""
        try:
            powerpoint.Quit()
        except Exception:
            pass

    def _do_convert(self, abs_input, abs_output, quality_intent, result_holder):
        """
        Worker that runs inside a thread so we can enforce a hard timeout.
        Writes True/exception into result_holder[0].
        """
        import pythoncom
        pythoncom.CoInitialize()  # Required: COM must be initialized on every thread that uses it
        powerpoint = None
        presentation = None
        try:
            powerpoint = self._fresh_powerpoint()
            presentation = powerpoint.Presentations.Open(
                abs_input,
                ReadOnly=True,
                Untitled=True,
                WithWindow=False,
            )
            # ppSaveAsPDF = 32
            # SaveAs uses PowerPoint's default PDF export (screen quality, smallest files)
            # This is equivalent to File > Save As > PDF in PowerPoint
            presentation.SaveAs(abs_output, 32)
            result_holder[0] = True
        except Exception as exc:
            result_holder[0] = exc
        finally:
            if presentation:
                try:
                    presentation.Close()
                except Exception:
                    pass
            if powerpoint:
                self._quit_safely(powerpoint)
            pythoncom.CoUninitialize()

    def convert_file(self, input_file: str, output_dir: str = None, verbose: bool = True):
        """
        Convert a single PPTX file to PDF.

        Starts a fresh PowerPoint instance for every file and enforces a hard
        timeout. Retries up to MAX_RETRIES times on failure, killing leftover
        PowerPoint processes between attempts.

        Returns True on success, False on failure.
        """
        input_path = Path(input_file)

        if not input_path.exists():
            if verbose:
                print(f"ERROR: File not found: {input_file}")
            return False

        if input_path.suffix.lower() not in ('.pptx', '.ppt'):
            if verbose:
                print(f"SKIP: Not a PowerPoint file: {input_file}")
            return False

        # Resolve output path
        if output_dir:
            out_dir = Path(output_dir)
            out_dir.mkdir(parents=True, exist_ok=True)
        else:
            out_dir = input_path.parent

        output_path = out_dir / f"{input_path.stem}.pdf"
        abs_input  = str(input_path.resolve())
        abs_output = str(output_path.resolve())

        quality_intent = self.QUALITY_SETTINGS[self.quality]['ppFixedFormatIntent']

        if verbose:
            mb = input_path.stat().st_size / (1024 * 1024)
            print(f"Converting: {input_path.name} ({mb:.1f} MB) "
                  f"[{self.QUALITY_SETTINGS[self.quality]['name']}]")

        for attempt in range(1, self.MAX_RETRIES + 2):  # attempts: 1, 2, 3
            if attempt > 1:
                if verbose:
                    print(f"  Retry {attempt - 1}/{self.MAX_RETRIES} for {input_path.name}...")
                # Clean up any stale PowerPoint processes before retrying
                _kill_powerpoint_processes()

            result_holder = [None]
            thread = threading.Thread(
                target=self._do_convert,
                args=(abs_input, abs_output, quality_intent, result_holder),
                daemon=True,
            )
            thread.start()
            thread.join(timeout=self.CONVERSION_TIMEOUT)

            if thread.is_alive():
                # Hard timeout — conversion hung
                if verbose:
                    print(f"  TIMEOUT: {input_path.name} exceeded "
                          f"{self.CONVERSION_TIMEOUT}s — killing PowerPoint")
                _kill_powerpoint_processes()
                # Don't retry a timeout — the file is probably corrupt/too large
                return False

            result = result_holder[0]

            if result is True and output_path.exists():
                pdf_mb = output_path.stat().st_size / (1024 * 1024)
                if verbose:
                    print(f"  OK: {output_path.name} ({pdf_mb:.1f} MB)")
                return True

            # Conversion returned an exception or PDF wasn't created
            error_msg = str(result) if isinstance(result, Exception) else "PDF not created"
            if verbose:
                print(f"  FAILED (attempt {attempt}): {error_msg}")

            # Remove partial output file if it exists
            if output_path.exists():
                try:
                    output_path.unlink()
                except Exception:
                    pass

        if verbose:
            print(f"  GIVING UP: {input_path.name} failed after "
                  f"{self.MAX_RETRIES + 1} attempts")
        return False


def install_pywin32():
    """Helper to install pywin32"""
    import subprocess
    print("Installing pywin32...")
    try:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pywin32'])
        print("pywin32 installed. Please restart the application.")
        return True
    except subprocess.CalledProcessError:
        print("Failed to install pywin32. Run manually: pip install pywin32")
        return False


if __name__ == '__main__':
    if check_powerpoint_installation():
        print("OK: PowerPoint COM automation is available")
    else:
        print("ERROR: PowerPoint COM automation is NOT available")
        if not is_powerpoint_available():
            print("  Reason: pywin32 not installed — run: pip install pywin32")
        else:
            print("  Reason: PowerPoint may not be installed or accessible")
