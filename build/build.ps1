# =============================================================================
#  build.ps1  --  One-shot Windows build for Procalc Hydraulics
#
#  What it does:
#    1. Creates (or reuses) a local virtual environment  .\.venv
#    2. Installs the app requirements + PyInstaller into it
#    3. Runs PyInstaller against procalc.spec  ->  dist\Procalc\
#    4. Runs Inno Setup (iscc) against installer.iss  ->  Output\ProcalcSetup-*.exe
#    5. Prints the path to the finished installer
#
#  Prerequisites (see README.md):
#    * Python 3.11+ (x64) on PATH  (the "py" launcher is preferred if present)
#    * Inno Setup 6 installed (provides iscc.exe)
#    * Run from THIS directory:   repo_root\build
#
#  Usage (from a PowerShell prompt in repo_root\build):
#      .\build.ps1
#
#  If script execution is blocked, launch once with:
#      powershell -ExecutionPolicy Bypass -File .\build.ps1
# =============================================================================

# Stop on the first uncaught error; make cmdlet errors terminating.
$ErrorActionPreference = "Stop"

# Resolve directories relative to this script so it works from any CWD.
$BuildDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$RepoRoot = Split-Path -Parent $BuildDir
$VenvDir  = Join-Path $BuildDir ".venv"
$Reqs     = Join-Path $RepoRoot "procalc_app\requirements.txt"
$Spec     = Join-Path $BuildDir "procalc.spec"
$Iss      = Join-Path $BuildDir "installer.iss"

Write-Host "=== Procalc Windows build ===" -ForegroundColor Cyan
Write-Host "Repo root : $RepoRoot"
Write-Host "Build dir : $BuildDir"

# --- pick a base Python interpreter -----------------------------------------
# Prefer the "py" launcher (py -3), fall back to "python" on PATH.
$BasePython = $null
if (Get-Command py -ErrorAction SilentlyContinue) {
    $BasePython = "py -3"
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $BasePython = "python"
} else {
    throw "No Python found on PATH. Install Python 3.11+ (x64) and re-run."
}
Write-Host "Base Python: $BasePython"

# --- 1) create the venv if it does not exist --------------------------------
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"
if (-not (Test-Path $VenvPython)) {
    Write-Host "`n[1/4] Creating virtual environment at $VenvDir ..." -ForegroundColor Yellow
    # Invoke the base interpreter to build the venv (handles "py -3" as two tokens).
    Invoke-Expression "$BasePython -m venv `"$VenvDir`""
} else {
    Write-Host "`n[1/4] Reusing existing virtual environment at $VenvDir" -ForegroundColor Yellow
}
if (-not (Test-Path $VenvPython)) {
    throw "venv creation failed: $VenvPython not found."
}

# --- 2) install dependencies ------------------------------------------------
Write-Host "`n[2/4] Installing dependencies (requirements + PyInstaller) ..." -ForegroundColor Yellow
& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -r $Reqs
& $VenvPython -m pip install pyinstaller

# --- 3) run PyInstaller (one-dir) -------------------------------------------
Write-Host "`n[3/4] Running PyInstaller ..." -ForegroundColor Yellow
# Run from the build dir so dist\ / build\ land here and relative paths in the
# .spec resolve as intended.
Push-Location $BuildDir
try {
    & $VenvPython -m PyInstaller $Spec --clean --noconfirm
} finally {
    Pop-Location
}

$DistExe = Join-Path $BuildDir "dist\Procalc\Procalc.exe"
if (-not (Test-Path $DistExe)) {
    throw "PyInstaller did not produce $DistExe"
}
Write-Host "PyInstaller output: $DistExe" -ForegroundColor Green

# --- 4) run Inno Setup ------------------------------------------------------
Write-Host "`n[4/4] Building the installer with Inno Setup ..." -ForegroundColor Yellow

# Locate iscc.exe: PATH first, then the standard install locations.
$Iscc = $null
if (Get-Command iscc.exe -ErrorAction SilentlyContinue) {
    $Iscc = (Get-Command iscc.exe).Source
} else {
    foreach ($cand in @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles}\Inno Setup 6\ISCC.exe"
    )) {
        if ($cand -and (Test-Path $cand)) { $Iscc = $cand; break }
    }
}
if (-not $Iscc) {
    throw "iscc.exe (Inno Setup 6) not found. Install Inno Setup 6 or add it to PATH."
}
Write-Host "Using Inno Setup: $Iscc"

Push-Location $BuildDir
try {
    & $Iscc $Iss
} finally {
    Pop-Location
}

# --- report the produced installer ------------------------------------------
$OutDir = Join-Path $BuildDir "Output"
$Installer = Get-ChildItem -Path $OutDir -Filter "ProcalcSetup-*.exe" -ErrorAction SilentlyContinue |
             Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($Installer) {
    Write-Host "`n=== BUILD COMPLETE ===" -ForegroundColor Cyan
    Write-Host "Installer: $($Installer.FullName)" -ForegroundColor Green
} else {
    Write-Warning "Build finished but no installer was found in $OutDir"
}
