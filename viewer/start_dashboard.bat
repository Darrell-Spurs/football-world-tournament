@echo off
REM Double-click to launch the live squad dashboard (read/write).
cd /d "%~dp0"

REM Ensure the Excel library is available for saving swaps / finalizing.
python -m pip install --quiet openpyxl 2>nul || py -m pip install --quiet openpyxl 2>nul

python serve.py
if errorlevel 1 py serve.py
pause
