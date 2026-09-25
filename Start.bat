@echo off
title ESSLCE Result Extractor
cd /d "%~dp0"
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" launcher.py
) else if exist ".venv\Scripts\python.exe" (
    start "" ".venv\Scripts\python.exe" launcher.py
) else (
    echo Python not found inside .venv
    echo Please install Python and run: pip install -r requirements.txt
    pause
)
