# HyTrays

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
