@echo off
setlocal
set "PYTHONDONTWRITEBYTECODE=1"
set "PYTHONUTF8=1"
cd /d "%~dp0scripts"
python -B preparar_proyecto.py %*
if errorlevel 1 (
  pause
  exit /b 1
)
python -B pipeline_explorer.py %*
if errorlevel 1 pause
endlocal
