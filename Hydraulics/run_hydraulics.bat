@echo off
REM Windows: double-click to open the hydraulics launcher (text menu).
REM For the GUI instead, run:  run_hydraulics.bat gui
cd /d "%~dp0"

where python >nul 2>nul
if %errorlevel%==0 (
  set "PY=python"
) else (
  where py >nul 2>nul && set "PY=py"
)

if "%PY%"=="" (
  echo Python 3 was not found. Install it from https://www.python.org/downloads/
  pause
  exit /b 1
)

if /I "%~1"=="gui" (
  %PY% run_hydraulics.py --gui
) else (
  %PY% run_hydraulics.py
)

echo.
pause
