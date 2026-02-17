@echo off
REM PPTX to PDF Converter - GUI Launcher
REM Double-click this file to start the converter

cd /d "%~dp0"
python converter_gui.py

if errorlevel 1 (
    echo.
    echo Error: Python may not be installed or not in PATH
    echo.
    pause
)
