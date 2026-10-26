@echo off
set "PYTHONDONTWRITEBYTECODE=1"
cd /d "%~dp0scripts"
python -B pipeline_explorer.py %*
if errorlevel 1 pause
