# -*- mode: python ; coding: utf-8 -*-
# =============================================================================
# PyInstaller spec  --  Procalc Hydraulics  (ONE-DIR build, not one-file)
#
# Build from THIS directory (repo_root/build) with:
#
#     pyinstaller procalc.spec --clean --noconfirm
#
# Output: dist/Procalc/  (a folder containing Procalc.exe + _internal/ payload)
# which installer.iss then packages into a Windows installer.
#
# Layout facts this spec encodes (see repo, not editable here):
#   * Entry point .......... procalc_app/__main__.py  (python -m procalc_app)
#   * App package .......... procalc_app/  (top-level imports: app, engine_api,
#                            resources, model.*, run.*, results.*, pms.*)
#   * Engine (OUTSIDE pkg).. Hydraulics/*.py  and  common/*.py, put on sys.path
#                            at runtime by engine_api.py.  Collected here as
#                            top-level hidden imports via pathex.
#   * Data that MUST ship .. Hydraulics/gems_extracted.json  (~2.6 MB PMS seed),
#                            Hydraulics/assets/ten_logo.png,
#                            procalc_app/resources/ (ten_logo.png + theme.py).
#
# IMPORTANT -- gems_extracted.json placement
#   pms_classes.py computes its seed path as  _HERE/gems_extracted.json  where
#   _HERE = dirname(pms_classes.__file__).  Under the freeze pms_classes is a
#   top-level module, so its __file__ resolves to sys._MEIPASS and it therefore
#   looks for the seed at  <bundle>/gems_extracted.json  (the ROOT).  We ship
#   the json at BOTH the bundle root (so the seed is actually found at runtime)
#   AND under a Hydraulics/ sub-folder (mirrors the source layout and is picked
#   up by the runtime hook's defensive path additions).  Costs one duplicate
#   ~2.6 MB file -- deliberate, so first-run seeding of %LOCALAPPDATA%\Procalc\
#   pms.json cannot silently fall back to the 2 curated-only classes.
# =============================================================================

import os
import sys

block_cipher = None

# ---------------------------------------------------------------------------
# Paths.  SPECPATH is the directory containing this .spec (repo_root/build).
# Everything else is expressed relative to the repo root, one level up.
# ---------------------------------------------------------------------------
BUILD_DIR = SPECPATH
ROOT      = os.path.abspath(os.path.join(BUILD_DIR, os.pardir))
APP       = os.path.join(ROOT, "procalc_app")
HYDRA     = os.path.join(ROOT, "Hydraulics")
COMMON    = os.path.join(ROOT, "common")

# Optional application icon.  Supply build/procalc.ico (Windows) or
# build/procalc.icns (macOS) to brand the exe/app/installer; if it's absent
# PyInstaller falls back to its default icon (the build still succeeds).
# See README.md.
_ICON = os.path.join(BUILD_DIR, "procalc.icns" if sys.platform == "darwin" else "procalc.ico")
ICON  = _ICON if os.path.exists(_ICON) else None

# ---------------------------------------------------------------------------
# Data files bundled into the one-dir payload.  Tuple = (source, dest_subdir),
# dest_subdir relative to the bundle root (sys._MEIPASS at runtime).
# ---------------------------------------------------------------------------
datas = [
    # App resources: ten_logo.png (+ theme.py / __init__.py).  engine_api.py
    # resolves the logo as  <pkg>/resources/ten_logo.png ; as a top-level
    # frozen module that becomes  <bundle>/resources/ten_logo.png.
    (os.path.join(APP, "resources"), "resources"),

    # PMS catalogue seed.  Shipped at the bundle ROOT (where pms_classes.py
    # actually reads it) AND under Hydraulics/ (source-mirroring + hook path).
    (os.path.join(HYDRA, "gems_extracted.json"), "."),
    (os.path.join(HYDRA, "gems_extracted.json"), "Hydraulics"),

    # Engine-side logo (used by the engine's Excel/PDF output).
    (os.path.join(HYDRA, "assets", "ten_logo.png"), os.path.join("Hydraulics", "assets")),
]

# ---------------------------------------------------------------------------
# Hidden imports.  PyInstaller's static analysis misses:
#   * the engine modules (imported by string / via sys.path at runtime),
#   * a few app submodules imported lazily inside functions,
#   * Qt/openpyxl/PIL sub-packages loaded only on demand.
# pathex below makes the bare engine module names resolvable.
# ---------------------------------------------------------------------------
hiddenimports = [
    # --- hydraulics engine (Hydraulics/*.py, imported top-level) ---
    "hydraulics_XOM",
    "hmb_proii_reader",
    "td_parser",
    "pms_classes",
    "comp_constants",
    # Live simulation connectors: importable on every platform (the actual
    # win32com dependency inside each is guarded by its own try/except
    # ImportError, so these are a no-op — "feature unavailable" — anywhere
    # pywin32 isn't present, e.g. macOS/Linux).
    "hysys_com_reader",
    "proii_com_reader",
    # --- shared helpers (common/*.py, imported top-level) ---
    "units",
    "style",
    # --- app submodules imported lazily / dynamically ---
    "app",
    "engine_api",
    "resources",
    "resources.theme",
    "model.grid_model",
    "model.delegates",
    "run.controller",
    "results.results_view",
    "results.workbook_model",
    "results.charts",
    "results.export",
    "pms.pms_manager",
    # --- Qt modules used only inside on-demand widgets/exports ---
    "PySide6.QtCharts",        # results/charts.py (pressure & flow-pattern plots)
    "PySide6.QtPrintSupport",  # print/PDF plumbing (QPdfWriter path)
    # --- openpyxl / PIL bits loaded lazily by the engine's spreadsheet I/O ---
    "openpyxl.cell._writer",   # optimized write path, imported by name
    "PIL",                     # Pillow: openpyxl.drawing.image embeds the logo
]

# Windows only: pywin32's COM plumbing, used by hysys_com_reader.py /
# proii_com_reader.py for live HYSYS/PRO-II automation. Not needed (and not
# installed — see requirements.txt's sys_platform marker) on macOS/Linux.
if sys.platform == "win32":
    hiddenimports += ["win32com", "win32com.client", "pythoncom", "pywintypes"]

# ---------------------------------------------------------------------------
# Excludes.  Trim heavy / unused libraries and Qt sub-modules the app never
# touches to keep the bundle small.  Verified against the source: neither the
# app nor the engine import pandas/numpy/matplotlib/tkinter, and no QML/Quick/
# 3D/WebEngine Qt module is used (charts are QtCharts + openpyxl, not QML).
# ---------------------------------------------------------------------------
excludes = [
    "pandas",
    "numpy",
    "matplotlib",
    "tkinter",
    "PySide6.QtQml",
    "PySide6.QtQuick",
    "PySide6.Qt3DCore",
    "PySide6.Qt3DRender",
    "PySide6.Qt3DInput",
    "PySide6.Qt3DLogic",
    "PySide6.Qt3DAnimation",
    "PySide6.Qt3DExtras",
    "PySide6.QtWebEngineCore",
    "PySide6.QtWebEngineWidgets",
    "PySide6.QtWebEngineQuick",
]

a = Analysis(
    [os.path.join(APP, "__main__.py")],
    # pathex: repo root + the three source dirs so that top-level names
    # (app, engine_api, hydraulics_XOM, units, style, ...) all resolve.
    pathex=[ROOT, APP, HYDRA, COMMON],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    # Runtime hook: injects the bundled engine dirs onto sys.path before the
    # app runs (defensive -- see rthook_paths.py).
    runtime_hooks=[os.path.join(BUILD_DIR, "rthook_paths.py")],
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,          # one-dir: binaries live in COLLECT, not EXE
    name="Procalc",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                      # UPX off: avoids AV false positives on Windows
    console=False,                  # GUI app -- no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=ICON,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="Procalc",                 # -> dist/Procalc/
)

# macOS only: wrap the one-dir payload into a real double-clickable .app
# bundle (-> dist/Procalc.app).  No-op on every other platform -- the
# Windows/Linux build stops at the COLLECT() above, unchanged.
if sys.platform == "darwin":
    app_bundle = BUNDLE(
        coll,
        name="Procalc.app",
        icon=ICON,
        bundle_identifier="com.ten.procalc",
        info_plist={
            "CFBundleShortVersionString": "0.1.0",
            "NSHighResolutionCapable": True,
        },
    )
