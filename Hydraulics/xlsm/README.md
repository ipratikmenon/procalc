# Procalc Hydraulics — macro workbook (v1)

`Procalc_Hydraulics.xlsm` runs the hydraulics **inside Excel**: edit the **Input** sheet, press **RUN HYDRAULICS**
(Control sheet), and **Pressure_Profile** and **CV_Summary** are rebuilt in the same workbook (~1-2 s).
No Python needed. Total size ≈ 55 KB (≈ 100 KB with all 440 streams of the sample HMB).

| Sheet | Purpose |
|---|---|
| Control | RUN button, flow-case multiplier % (C4), last-run status (C8) |
| Input | Same 37 columns as `Pipeline_Input` (paste your sheet; header row = the row with `Circuit` in column A; unit suffixes ignored). FPS units |
| Streams | Stream property table. Row 1 = `Stream` + HMB keys (same as `hmbbuilder.py wide`). Paste yours, or rebuild with `build_xlsm.py --hmb HMB.xlsx` |
| Component_Constants | Reserved for the flash model (phase 2) |
| Data_K / Data_OD / Data_Pipe / Data_PMS | Fitting K, ASME B36.10 OD/walls, PMS class → schedule (edit class names / schedules here) |
| Pressure_Profile, CV_Summary | Outputs (generated) |

## What v1 does
* Darcy-Weisbach (Swamee-Jain), K-method fittings, elevation, ID from PMS class + bore (same tables as `hydraulics.py`)
* Single phase and homogeneous two-phase (McAdams µ); gas ρ = P·MW/(Z·R·T) at the local pressure (isothermal)
* Boundaries / devices: Source (Set P), Destination (Set P), Fix Pressure, Fix K, orifice / inline FT (Instr dP)
* Control valve: Fixed dP > Set P (outlet) > flow control solved to the Destination > Fixed K; Cv (ISA 75.01 liquid / gas)
* Pump-discharge back-calculation: Source without Set P + Destination with Set P → source pressure is solved
* Flow case: multiplier on Control!C4 scales all stream flows

Validated against `hydraulics.py`: single-phase gas and two-phase kernels agree to < 1e-4 psi per station; the L-810 pump
case gives 246.632 psia (same as Python).

## Not in v1 (phase 2 — still in `hydraulics.py`)
Rachford-Rice flash from composition, Peng-Robinson / JT temperature, tees & branches, FIV / AIV / flow-regime maps,
Component_Detail / Composition_Splits, thermowell / inline instrument datasheets, light output theme.

## If Excel does not show the RUN macro
Open the workbook, Alt+F11 → File → Import File → `Procalc.bas`, save as `.xlsm`. The button is already linked to `RunHydraulics`.

## Rebuild
```
python build_xlsm.py [--hmb ../HMB.xlsx] [--out Procalc_Hydraulics.xlsm]
```
`ovba.py` writes the VBA project (source-only, no p-code) so the file stays small; Excel compiles it on open.
