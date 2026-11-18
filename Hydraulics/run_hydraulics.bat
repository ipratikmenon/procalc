@echo off
REM ============================================================================
REM  Procalc Hydraulics launcher (Windows) - double-click for the text menu.
REM
REM    run_hydraulics.bat                       menu: build the HMB from a PRO/II dump,
REM                                             pick the input workbook and circuit, run the hydraulics
REM    run_hydraulics.bat gui                   small window instead of the text menu
REM    run_hydraulics.bat build DUMP.xlsx [HMB_OUT.xlsx]
REM    run_hydraulics.bat run   INPUT.xlsx HMB.xlsx [CIRCUIT] [--case "Case 1"] [--isenthalpic] [--light]
REM    run_hydraulics.bat all   DUMP.xlsx INPUT.xlsx [CIRCUIT]     build the HMB, then run
REM
REM  Keep this file, run_hydraulics.py, hmbbuilder.py and hydraulics.py in one folder.
REM ============================================================================
cd /d "%~dp0"

set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY (
  where py >nul 2>nul && set "PY=py"
)
if not defined PY (
  echo Python 3 was not found. Install it from https://www.python.org/downloads/
  pause
  exit /b 1
)

%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo The 'openpyxl' package is missing - installing it now...
  %PY% -m pip install openpyxl
)

if /I "%~1"=="gui" (
  %PY% run_hydraulics.py --gui
) else (
  %PY% run_hydraulics.py %*
)

echo.
pause
