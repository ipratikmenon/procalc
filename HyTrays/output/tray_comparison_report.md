# HyTrays tray comparison -- distillation column design

**Case:** T801 rectifying section -- top tray (HMB.xlsx basis)

## Design basis

| Quantity | Value | Source |
|---|---|---|
| Vapor traffic, V | 52,449 lb/hr | T801-OH, `Hydraulics/HMB.xlsx` (measured) |
| Vapor density, rho_V | 1.018 lb/ft3 | T801-OH actual vol. rate, `Hydraulics/HMB.xlsx` (measured) |
| Liquid (reflux) traffic, L | 14,798 lb/hr | T801-REFLUX, `Hydraulics/HMB.xlsx` (measured) |
| Liquid density, rho_L | 30.0 lb/ft3 | **estimated** for ~344 F tray liquid |
| Surface tension, sigma | 9.0 dyn/cm | **estimated** |
| Liquid viscosity, mu_L | 0.12 cP | **estimated** |
| Relative volatility, alpha | 1.40 | **estimated** |
| Tray spacing | 24 in | assumed |
| Design flood fraction | 80% | assumed |
| Theoretical stages (this section) | 20 | illustrative |

Flow parameter FLV = (L/V)*sqrt(rho_V/rho_L) = **0.0520** -- a low-FLV, vapor-dominated system, typical of an atmospheric/light-pressure hydrocarbon rectifying section.

> Replace the *estimated* rows with simulator output once available; everything else in this report is recomputed automatically by `run_tray_comparison.py`.

## Results

| Tray | Diameter (ft) | u_nf (ft/s) | dP/tray (psi) | dP/tray (in liq) | DC backup (% lim) | Operating window (% of design) | Turndown | Eff (%) | N_actual | Height (ft) | Total dP (psi) |
|---|---|---|---|---|---|---|---|---|---|---|---|
| Sieve | 4.39 | 1.52 | 0.0557 | 3.21 | 34 | 62-125 | 2.00:1 | 76.2 | 27 | 64.0 | 1.503 |
| HyTrays Ripple | 4.13 | 1.67 | 0.0484 | 2.79 | 30 | 44-125 | 2.86:1 | 82.3 | 25 | 60.0 | 1.211 |
| Valve | 4.39 | 1.52 | 0.0386 | 2.22 | 26 | 38-125 | 3.33:1 | 76.2 | 27 | 64.0 | 1.042 |
| Dualflow | 3.61 | 1.75 | 0.0170 | 0.98 | n/a | 81-125 | 1.54:1 | 64.8 | 31 | 72.0 | 0.526 |
| High-Performance | 3.76 | 1.90 | 0.0351 | 2.02 | 23 | 50-125 | 2.50:1 | 73.9 | 28 | 66.0 | 0.983 |

## Discussion

**Capacity / column diameter.** At the same vapor and liquid traffic and the same 80% design flood point, the required column diameter ranges from **3.61 ft** (Dualflow) to **4.39 ft** (Sieve). High-Performance and Dualflow trays need the least active area per unit of vapor handled (capacity factors of 1.25x and 1.15x the sieve baseline, plus little/no downcomer area), so they shrink the shell the most. The new HyTrays Ripple deck comes in at 4.13 ft -- about 6% smaller than a plain sieve tray -- from its 1.10x capacity factor and slightly larger active-area fraction.

**Pressure drop.** Per-tray pressure drop tracks hole velocity (u_hole) and the dry-tray discharge coefficient. Valve trays show the lowest *clean* dry-tray drop here because of their high discharge coefficient (C0=0.85); in practice their floor (closed-valve weight, modeled as 0.40 in. liquid) dominates at low loads instead. Dualflow is lowest overall (17.0 mpsi/tray) because it has no weir and a large open area. HyTrays Ripple (48.4 mpsi/tray) sits in the middle of the pack -- a modest premium over Dualflow/High-Performance for the efficiency gain described below.

**Downcomer backup.** FLV is low in this case (vapor-dominated), so every tray with downcomers operates well inside its 50%-of-spacing backup limit (Sieve 34%, HyTrays Ripple 30%, Valve 26%, High-Performance 23%). Downcomer area is not the limiting consideration for this section; flooding/entrainment (Fair capacity) governs sizing instead.

**Turndown / operating window.** Bubble-cap-style decks and the corrugated HyTrays Ripple deck retain liquid at low vapor rates and resist weeping, giving the widest stable operating windows (3.33:1 for Valve, 2.86:1 for HyTrays Ripple). Dualflow has the narrowest window (1.54:1) because its large, unweired perforations weep heavily as soon as vapor rate drops -- it needs a fairly steady load to stay efficient.

**Efficiency & actual trays.** The O'Connell baseline efficiency for this system is **76.2%**. HyTrays Ripple's extra interfacial area from the corrugated deck (1.08x factor) raises this to **82.3%**, the highest of the five -- for 20 theoretical stages it needs only **25 actual trays** (column height 60 ft), versus **27 trays** (64 ft) for a plain sieve tray and **31 trays** (72 ft) for Dualflow, whose lower contacting efficiency (0.85x) offsets its low per-tray pressure drop once the whole column is added up: total column dP is 0.53 psi for Dualflow vs 1.21 psi for HyTrays Ripple and 1.50 psi for plain sieve.

## Summary / selection guidance

- **Sieve** -- baseline. Cheapest per-tray hardware, but the largest shell, the most trays for a given separation and the worst turndown of the conventional designs.
- **HyTrays Ripple** -- best overall balance in this case: smaller shell than sieve/valve, the *highest* efficiency (fewest actual trays / shortest column), a wide operating window, and total column dP close to the High-Performance tray -- without the cost or mechanical complexity of a multi-downcomer/grid design. *(Parameters are placeholders -- refine once the HyTrays Ripple datasheets are added to `HyTrays/Datasheets/`.)*
- **Valve** -- same shell diameter as sieve at design, but the widest operating window of all five -- the right choice when the column must run efficiently over a wide turndown range.
- **Dualflow** -- smallest per-tray dP and a compact-ish shell, at the cost of the lowest efficiency, the most actual trays/tallest column, and a narrow operating window. Best suited to fouling or high-liquid-load services that value simplicity over turndown.
- **High-Performance** -- the smallest shell of all (lowest capex on the column shell) with efficiency close to sieve; a strong choice when plot space / shell diameter is the controlling constraint.

---
*Generated by `HyTrays/run_tray_comparison.py`. Re-run after editing `CASE` in that file (operating conditions) or the tray parameters in `HyTrays/tray_library.py` (e.g. once HyTrays Ripple datasheets are available).*
