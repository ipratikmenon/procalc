# HyTrays

## Apex concept (severe-service, step-change-capacity tray)

`HyTrays_Apex_concept.md` is a design-concept memo for a flagship tray
targeting extreme fouling resistance, a 2-5 psi structural (uplift) deck
rating and ~2x ULTRA-FRAC-class capacity via co-current "vortex-grid" boost
zones. It is a paper exercise (no parameters in `tray_library.py` yet) --
see its Section 7 for the path to integrate it into the comparison engine
once the open R&D questions (Sec. 6) have a basis. Supporting reference
images are in `References/`.

## Datasheets

Drop the HyTrays tray-family PDF/Excel datasheets (and the Ansys files,
once they're small enough to add) into `HyTrays/Datasheets/`. This folder
is currently empty -- no PDF/Excel files for the new HyTrays trays were
found in the repository to move yet. Once they're added, the
`RIPPLE` entry in `tray_library.py` should be updated from the manufacturer
geometry/performance data (currently placeholder values, see the comments
in that file).

## Distillation column tray comparison tool

A small numerical tray-hydraulics engine comparing five tray families --
**Sieve, HyTrays Ripple, Valve, Dualflow and a generic High-Performance
tray** -- on the same column section using Fair's flooding correlation,
dry/wet tray pressure drop, downcomer backup, turndown and an
O'Connell-based efficiency estimate.

| File | Purpose |
|---|---|
| `tray_library.py` | Tray-type geometry/performance parameters (`TrayType` dataclass + the 5 definitions). |
| `column_hydraulics.py` | The hydraulics engine: Fair capacity correlation, pressure drop, downcomer backup, turndown, O'Connell efficiency (`design_tray`). |
| `run_tray_comparison.py` | Run script: defines the design-basis case, runs all 5 tray types, prints a table, and writes the CSV/plots/report below. |

### Run it

```bash
cd HyTrays
python3 run_tray_comparison.py
```

### Output (`HyTrays/output/`)

- `tray_comparison.csv` -- full numeric results
- `01_capacity_and_sizing.png` -- flooding velocity & required column diameter
- `02_pressure_drop.png` -- per-tray and total column pressure drop
- `03_efficiency_and_trays.png` -- tray efficiency, actual tray count & column height
- `04_operating_window.png` -- stable operating range (weep -> flood) / turndown
- `05_downcomer_backup.png` -- downcomer backup vs. the 50%-of-spacing limit
- `tray_comparison_report.md` -- design basis, results table and written discussion

### Design basis

The default case in `run_tray_comparison.py` sizes against the T801
rectifying-section top-tray vapor/liquid traffic taken from
`Hydraulics/HMB.xlsx` (T801-OH, T801-REFLUX). Liquid density, viscosity,
surface tension and relative volatility are **engineering estimates**
(the workbook doesn't carry tray-level physical properties) -- replace
those four values with simulator output for a fully rigorous case. Edit
`CASE` in `run_tray_comparison.py` and re-run to evaluate other sections /
services.

## Multi-service tray selection study

A second, broader comparison that sizes all five tray families against
**five representative services** (general rectification, fouling/heavy-ends,
vacuum, high-pressure/high-liquid-load, foaming) -- the classic
tray-selection discriminators in Kister, *Distillation Operation* Ch. 1.
Intended as the numerical backbone for a tray-selection paper/report.

| File | Purpose |
|---|---|
| `service_cases.py` | The 5 `ColumnCase` definitions (one per service), incl. foaming system factor and fouling flag. |
| `run_service_comparison.py` | Run script: sizes/rates all 5 tray types for all 5 services, prints per-service tables, and writes the CSV/plots/report below. |

### Run it

```bash
cd HyTrays
python3 run_service_comparison.py
```

### Output (`HyTrays/output/`)

- `service_comparison.csv` -- full numeric results, one row per (service, tray)
- `11_diameter_by_service.png` -- column diameter, all trays x all services
- `12_total_dp_by_service.png` -- total column pressure drop, all trays x all services
- `13_downcomer_backup_by_service.png` -- downcomer backup vs. the 50% limit
- `14_turndown_by_service.png` -- operating-range turndown ratio
- `15_fouling_sensitivity.png` -- dry-tray dP increase vs. open-area retention (flow-independent)
- `tray_selection_study.md` -- full report: tray/service tables, cross-service
  plots, per-service discussion and a tray-selection guidance matrix

Only the General Rectification case is anchored to the T801 stream data
above; the other four services use representative property sets so the
five tray types can be compared on a common, reproducible basis -- edit
`service_cases.py` to match a specific column once real data is available.
