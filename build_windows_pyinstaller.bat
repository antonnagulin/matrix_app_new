@echo off
setlocal enabledelayedexpansion

REM Build GUI executable for Windows (tkinter) using PyInstaller.
REM You must have Python 3.12 x64 installed and `pyinstaller` available.

cd /d "%~dp0"
set "PYTHONPATH=."

REM Using --onedir first (more reliable than --onefile).
REM Output: dist\matrix-check\matrix-check.exe
pyinstaller --clean --noconfirm --windowed --onedir --name matrix-check ^
  --paths . --paths src ^
  src\main.py

echo.
echo Build finished.
echo Executable: dist\matrix-check\matrix-check.exe
pause
