# Procalc Hydraulics — macro workbook (v1)

`Procalc_Hydraulics.xlsm` runs the hydraulics **inside Excel**: edit the **Input** sheet, press **RUN HYDRAULICS**
(Control sheet), and **Pressure_Profile** and **CV_Summary** are rebuilt in the same workbook (~1-2 s).
No Python needed. Total size ≈ 55 KB (≈ 100 KB with all 440 streams of the sample HMB).

| Sheet | Purpose |
|---|---|
| Control | RUN button, flow-case multiplier % (C4), last-run status (C8) |
| Input | Same 37 columns as `Pipeline_Input` (paste your sheet; header row = the row with `Circuit` in column A; unit suffixes ignored). FPS units |
| `<stream name>` sheets (e.g. `DX10-OUT`) | Paste the HMB per-stream sheet as is (composition included). A circuit whose *Stream Lookup* matches a sheet name is flashed |
| Component_Constants | Paste the HMB `COMP_CONSTANTS` sheet as is (B Name, E MW, I Tc, J Pc, M ω) — used by the flash (Wilson T-correction) and for vapour µ / Z |
| Streams | Fallback property table (no composition → no flash). Row 1 = `Stream` + HMB keys |
| Data_K / Data_OD / Data_Pipe / Data_PMS | Fitting K, ASME B36.10 OD/walls, PMS class → schedule (edit class names / schedules here) |
| Instrument_Data / Data_TWClass | Optional per-tag data (thermowell geometry, orifice d, CV rated Cv / FL / xT / curve) and per-piping-class thermowell defaults (material otherwise from the class MOC) |
| Pressure_Profile, CV_Summary, FIV_EI_T2.2, AIV, Two_Phase_Regime, Inline_Instruments, Composition_Splits, Component_Detail | Outputs (generated; the last two when a stream is flashed): per-station z / K / y / x and vapour / liquid lbmol/h per component |

## What it does (v1 + phase 2)
* Darcy-Weisbach (Swamee-Jain), K-method fittings, elevation, ID from PMS class + bore (same tables as `hydraulics.py`)
* Single phase and homogeneous two-phase (McAdams µ); gas ρ = P·MW/(Z·R·T) at the local pressure (isothermal)
* Boundaries / devices: Source (Set P), Destination (Set P), Fix Pressure, Fix K, orifice / inline FT (Instr dP)
* Control valve: Fixed dP > Set P (outlet) > flow control solved to the Destination > Fixed K; Cv (ISA 75.01 liquid / gas)
* Pump-discharge back-calculation: Source without Set P + Destination with Set P → source pressure is solved
* **Tees and branches:** `Tee Split Branch Flow N` / merge rows with *Flow Fraction from Main*, Mass Vapor / Liquid Carry Over, `Run Type = Branch` blocks (seeded from the Main split via Upstream Line No / Upstream Seq, or from their own Stream Lookup), and `Tee Join Branch Flow N` re-flashes the blended mixture. Branch rows are listed as `C1 [Branch n]` in every output sheet
* **Peng-Robinson package** (Data_PR / Data_PRAlias / Data_PRKij, 98 pure components, H2-tuned like the Python tool): when every component of a flashed stream is in the databank, flash-generated vapour gets its Z, viscosity and Cp/Cv from PR instead of the generic correlations. Control!C9 = ON also marches the temperature isenthalpically (Joule-Thomson) along the line; default OFF because `hydraulics.py` only does this for typed-composition streams, not HMB streams
* Flow case: multiplier on Control!C4 scales all stream flows
* **Flash-coupled marching** (Control!C7 = ON): Rachford-Rice at the local pressure with K(P) = K_ref·P_ref/P, reference K from the HMB phase split (y/x, else z/x with a solved λ), Wilson T-correction when a Temp override is typed; vapour µ / Z from Lucas / Peng-Robinson using Component_Constants (generic fallback if constants are incomplete). Output adds vapour fraction, β, vapour / liquid flow per station

Validated against `hydraulics.py`: single-phase gas, two-phase and flash-coupled (stream 198, flashing along the line, incl. choking) agree to < 1e-4 psi at every station; the L-810 pump
case gives 246.632 psia (same as Python).

## Also in the workbook now
FIV (EI T2.2), AIV (source attenuation), Two_Phase_Regime (six correlations + slug), Inline_Instruments (PT / TT / FT / FO / RO + thermowell wake check), CV_Summary with FL / xT / rated Cv / travel check and liquid choking (bubble-point Pv). All checked against `hydraulics.py`.

## Not in the workbook (still in `hydraulics.py`)
typed-composition input (Stream_Composition sheet), temperature mixing at a join,
light output theme. Tees: temperature mixing at a join is not applied; a branch seeded from a split uses the circuit feed; a join of two *different* HMB streams is ported but only same-stream split / join was checked against Python.

Note: a sub-cooled single-liquid stream has equal proxy K values, so (as in Python) it flashes all-or-nothing around its HMB pressure; keep the line above the HMB pressure or supply a stream that carries its own vapour split.

## If Excel does not show the RUN macro
Open the workbook, Alt+F11 → File → Import File → `Procalc.bas`, save as `.xlsm`. The button is already linked to `RunHydraulics`.

## Rebuild
```
python build_xlsm.py [--hmb ../HMB.xlsx] [--out Procalc_Hydraulics.xlsm]
```
`ovba.py` writes the VBA project (source-only, no p-code) so the file stays small; Excel compiles it on open.
