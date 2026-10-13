@echo off
cd /d "%~dp0"
python pipeline_explorer.py
if errorlevel 1 pause
