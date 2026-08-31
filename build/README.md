# Procalc Hydraulics — packaging (Windows + macOS)

This directory builds **Procalc Hydraulics** into a native desktop package.
It wraps the PySide6 app in `procalc_app/` together with the hydraulics engine
in `Hydraulics/` and the shared helpers in `common/`. `procalc.spec` is a
single, cross-platform PyInstaller spec: on Windows it produces a one-dir
bundle that `installer.iss` packages into an Inno Setup installer; on macOS
it additionally wraps that bundle into a `Procalc.app`, packaged into a
`.dmg` via `hdiutil`. Both are built by CI (`.github/workflows/build-windows.yml`
/ `build-macos.yml`) since PyInstaller never cross-compiles — a Windows
build can only come from a Windows host, a macOS build only from a macOS
host (see each section below).

Everything here is self-contained in `build/` — nothing outside it is modified.

---

## Contents

| File | Purpose |
|------|---------|
| `procalc.spec` | PyInstaller spec, **shared by both platforms** — one-dir build (`dist/Procalc/`) on Windows; also wraps a `dist/Procalc.app` bundle on macOS. |
| `rthook_paths.py` | PyInstaller runtime hook — puts the bundled engine dirs on `sys.path`. |
| `installer.iss` | Inno Setup 6 script (Windows only) — packages `dist/Procalc/` into an installer. |
| `build.ps1` | PowerShell driver (Windows only) — venv → deps → PyInstaller → Inno Setup. |
| `procalc.ico` | *(optional, you supply)* Windows app/installer icon. |
| `procalc.icns` | *(optional — CI generates it)* macOS app icon, built from `procalc_app/resources/ten_logo.png`. |

---

## Prerequisites (build machine — Windows)

1. **Python 3.11+ (x64)** — install from python.org and tick *Add Python to
   PATH*. A 64-bit interpreter yields a 64-bit exe (what `installer.iss`
   expects).
2. **Inno Setup 6** — https://jrsoftware.org/isdl.php . Provides `iscc.exe`
   (found automatically under `Program Files (x86)\Inno Setup 6\`).
3. **(Optional) `procalc.ico`** — drop a Windows `.ico` next to these files to
   brand the exe (and, if you uncomment `SetupIconFile` in `installer.iss`, the
   installer). If absent, PyInstaller uses its default icon and the build still
   succeeds.

> The build must run **on Windows**. PyInstaller does not cross-compile; a
> Windows exe/installer can only be produced from a Windows host (or VM).

---

## Build (one command)

From a PowerShell prompt in this `build/` directory:

```powershell
.\build.ps1
```

If script execution is blocked by policy:

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

`build.ps1` will:

1. create/reuse `.\.venv`,
2. `pip install -r ..\procalc_app\requirements.txt` + `pyinstaller`,
3. run `pyinstaller procalc.spec --clean --noconfirm` → `dist\Procalc\`,
4. run `iscc installer.iss` → `Output\ProcalcSetup-<version>.exe`,
5. print the installer path.

### Manual build (equivalent steps)

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r ..\procalc_app\requirements.txt pyinstaller
pyinstaller procalc.spec --clean --noconfirm
& "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe" installer.iss
```

The installer version is set by `#define AppVersion "0.1.0"` at the top of
`installer.iss` — bump it there for each release.

---

## macOS build (Apple Silicon, unsigned)

Built by `.github/workflows/build-macos.yml` on a `macos-latest` (Apple
Silicon) GitHub Actions runner — there's no local build step to run since
this repo's dev sandbox isn't macOS and PyInstaller can't cross-compile.
The workflow: installs deps, generates `procalc.icns` from
`procalc_app/resources/ten_logo.png` (via macOS's built-in `sips`/`iconutil`,
no extra tooling), runs `pyinstaller procalc.spec --clean --noconfirm` →
`dist/Procalc.app`, smoke-tests it headless, then packages it with
`hdiutil create -format UDZO` → `ProcalcSetup-0.1.0.dmg` — the macOS
equivalent of Inno Setup, built into the OS, no third-party installer needed.

**The app is unsigned** (no Apple Developer account) — on first launch macOS
Gatekeeper will refuse to open it with a plain double-click. To run it:

1. Drag `Procalc.app` out of the mounted `.dmg` into `/Applications` (or run
   it in place).
2. **Right-click** (or Control-click) the app → **Open** → **Open** again in
   the confirmation dialog. This is a one-time step per machine; every launch
   after that works normally with a plain double-click.
3. If step 2 doesn't offer an Open option, go to **System Settings → Privacy
   & Security**, scroll to the blocked-app notice, and click **Open Anyway**.

This is the standard, expected trade-off of shipping unsigned — there is no
cost/waiting-free way around it without an Apple Developer Program
membership ($99/yr) to sign and notarize the build.

---

## Running the app from source (no packaging)

From the repository root (`procalc/`):

```powershell
python -m pip install -r procalc_app\requirements.txt
python -m procalc_app
```

`procalc_app/__main__.py` inserts the package dir on `sys.path`, and
`engine_api.py` inserts `../common` and `../Hydraulics`, so the engine imports
resolve without any packaging.

---

## Where runtime data lives

* **PMS catalogue (`pms.json`)** — on first launch the app copies the bundled
  seed `gems_extracted.json` to a per-user location:

  ```
  Windows:  %LOCALAPPDATA%\Procalc\pms.json   (e.g. C:\Users\<you>\AppData\Local\Procalc\pms.json)
  macOS:    ~/Library/Application Support/Procalc/pms.json
  Linux:    ~/.local/share/Procalc/pms.json
  ```

  This per-user file is what the app edits/uploads; the bundled seed is never
  modified. Delete it to reset the catalogue to the shipped seed. The
  `PROCALC_PMS_JSON` environment variable overrides the location on any
  platform. The **installer/DMG does not create `pms.json`** — the app seeds
  it per-user on first run.

* **Run outputs** — result workbooks are written to a temp dir (via Python's
  `tempfile`: `%TEMP%\procalc_out_*` on Windows, `/var/folders/.../procalc_out_*`
  on macOS) unless the app is told otherwise.

---

## How the bundle is laid out (for troubleshooting)

PyInstaller one-dir puts everything under `dist\Procalc\` (`Procalc.exe` plus an
`_internal\` payload). Key bundled items and why:

* `resources\ten_logo.png` — app logo; `engine_api.LOGO_PATH` reads it here.
* `gems_extracted.json` at the **bundle root** — `pms_classes.py` reads its seed
  from next to itself (`_MEIPASS`), so the root copy is the one used at runtime.
  A second copy under `Hydraulics\` mirrors the source layout and is covered by
  the runtime hook's defensive `sys.path` additions.
* `Hydraulics\assets\ten_logo.png` — logo used by the engine's Excel/PDF output.
* Engine modules (`hydraulics_XOM`, `pms_classes`, `units`, `style`, …) are
  collected as hidden imports into the frozen archive; `rthook_paths.py` adds
  `Hydraulics\`/`common\` to `sys.path` as a fallback for any that ship as data.

If a `ModuleNotFoundError` appears for an engine module at runtime, add it to
`hiddenimports` in `procalc.spec`. If the PMS list shows only 2 classes
(`G1S-1`, `G1S-2`), the seed json was not found — confirm the root-level
`gems_extracted.json` is present in the bundle.

---

## Smoke tests on a clean Windows VM

Install `ProcalcSetup-<version>.exe`, then verify on a machine **without** Python
or the source tree:

1. **Launch** — Start-menu (and, if selected, desktop) shortcut opens the app;
   the T.EN-themed window appears with the toolbar and no console window.
2. **Load HMB** — *Load HMB* opens a stream/PRO-II export; the case dropdown and
   stream list populate (`Hydraulics/HMB.xlsx` or a PRO-II export is a good sample).
3. **Run** — with a starter/loaded circuit, *Run* completes without error and
   produces result workbook(s).
4. **Results tabs** — the results view shows the output sheets rendered in the
   Excel-like table.
5. **PMS upload** — open the PMS manager; the catalogue lists the full ~250
   classes (proves the seed shipped and seeded `%LOCALAPPDATA%\Procalc\pms.json`);
   upload a `gems_extracted.json` and confirm it re-indexes.
6. **Excel / PDF export** — export a result to `.xlsx` and to PDF; open both and
   confirm the T.EN logo, formatting, and charts render.
7. **Charts** — the pressure-profile and flow-pattern charts draw (exercises the
   bundled `PySide6.QtCharts`).
8. **Uninstall** — Add/Remove Programs removes the app cleanly (the per-user
   `pms.json` is intentionally left behind).

---

## Notes / assumptions

* One-dir (not one-file): faster startup, easier antivirus whitelisting, and the
  installer ships a normal folder. UPX is disabled to avoid AV false positives.
* Excludes trim `pandas`/`numpy`/`matplotlib`/`tkinter` and unused Qt modules
  (QML/Quick/3D/WebEngine) — verified none are imported by the app or engine.
* Icon is optional; supply `build/procalc.ico` to brand the exe.
* Signing: to code-sign, sign `dist\Procalc\Procalc.exe` before running `iscc`,
  and optionally sign the finished installer (configure `SignTool` in Inno Setup).
