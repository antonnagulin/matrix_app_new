@echo off
cd /d "%~dp0"
python -m src.main
if errorlevel 1 pause
