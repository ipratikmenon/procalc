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
| HT-01A Hinge | 4.23 | 1.59 | 0.0426 | 2.45 | 28 | 10-125 | 12.50:1 | 80.0 | 26 | 62.0 | 1.107 |
| HT-01B Spring | 4.23 | 1.59 | 0.0453 | 2.61 | 29 | 11-125 | 11.11:1 | 80.0 | 26 | 62.0 | 1.178 |
| HT-01C LipSeal | 4.12 | 1.64 | 0.0380 | 2.19 | 25 | 48-125 | 2.63:1 | 77.7 | 26 | 62.0 | 0.989 |
| HT-02 CVS | 3.48 | 2.50 | 0.0545 | 3.14 | 35 | 50-125 | 2.50:1 | 76.2 | 27 | 64.0 | 1.473 |
| HT-03 GRADEX | 3.91 | 1.82 | 0.0478 | 2.75 | 30 | 38-125 | 3.33:1 | 85.3 | 24 | 58.0 | 1.146 |
| HT-05 PULSAR | 4.23 | 1.59 | 0.0500 | 2.88 | 31 | 31-125 | 4.00:1 | 83.8 | 24 | 58.0 | 1.201 |

## Discussion

**Capacity / column diameter.** At the same vapor and liquid traffic and the same 80% design flood point, the required column diameter ranges from **3.48 ft** (HT-02 CVS) to **4.39 ft** (Sieve). The smallest shell belongs to **HT-02 CVS**, whose 1.65x capacity factor (on Fair's C_SBF) lets it handle the same vapor in the least active area. Among the HT-series decks, **HT-02 CVS** gives the smallest shell (3.48 ft -- about 21% smaller than the plain sieve baseline) from its 1.65x capacity factor.

**Pressure drop.** Per-tray pressure drop tracks hole velocity (u_hole) and the dry-tray discharge coefficient (C0). The lowest total column dP here is **HT-01C LipSeal** (0.99 psi), a product of both its per-tray drop and the number of trays its efficiency requires. The plain sieve baseline runs 55.7 mpsi/tray (1.50 psi total); decks with higher discharge coefficients (valve-/push-valve-style C0) cut the dry-tray component, while those that add stages cut the *total* by needing fewer trays.

**Downcomer backup.** FLV is low in this case (vapor-dominated), so every tray with downcomers operates well inside its 50%-of-spacing backup limit (Sieve 34%, HT-01A Hinge 28%, HT-01B Spring 29%, HT-01C LipSeal 25%, HT-02 CVS 35%, HT-03 GRADEX 30%, HT-05 PULSAR 31%). Downcomer area is not the limiting consideration for this section; flooding/entrainment (Fair capacity) governs sizing instead.

**Turndown / operating window.** The adaptive HT-01 decks (living hinge / leaf spring) follow the vapor load by opening and closing their flaps, so they resist weeping far below the design rate and post the widest stable operating windows. The widest here is **HT-01A Hinge** (12.50:1); the widest HT-series deck is **HT-01A Hinge** (12.50:1), versus 2.00:1 for the plain sieve baseline. Capacity-oriented decks (centrifugal swirl) need a minimum vapor rate to work and so have narrower windows.

**Efficiency & actual trays.** The O'Connell baseline efficiency for this system is **76.2%**. The highest effective efficiency here is **HT-03 GRADEX** at **85.3%** (1.12x factor); the best HT-series deck on this metric is **HT-03 GRADEX** (85.3%) -- for 20 theoretical stages it needs only **24 actual trays** (column height 58 ft), versus **27 trays** (64 ft) for a plain sieve tray. Higher contacting efficiency shortens the column and, by needing fewer trays, usually lowers the total column pressure drop as well.

## Summary / selection guidance

- **Sieve** (baseline, not a HyTrays product) -- cheapest per-tray hardware, but the largest shell, the most trays for a given separation and the worst turndown. Every HT-series deck below is rated relative to it.
- **HT-01A Hinge / HT-01B Spring** -- adaptive decks whose flaps track the vapor load, giving the widest turndown in the family (~12-15:1 / ~11:1); pick these where the column must run efficiently across a very wide load range. The spring variant trades a little turndown for lower stress and better robustness/fouling resistance.
- **HT-01C LipSeal** -- depending check-valve lip deck; a >=30% wider stable window than a plain sieve from its weep-sealing lips, a straightforward upgrade from sieve where mild turndown is the issue.
- **HT-02 CVS** -- centrifugal swirl deck; the highest capacity in the family (smallest shell) and capacity decoupled from tray spacing -- the choice when shell diameter / plot space dominates, at the cost of narrower turndown.
- **HT-03 GRADEX** -- radially-graded push-valve deck; the biggest efficiency lever on large-diameter trays (Peclet/plug-flow gain), shortening the column where many stages are needed.
- **HT-05 PULSAR** -- fluidic-oscillator self-sweeping deck; modest capacity/turndown but the best fouling resistance in the family and a useful efficiency bump -- aimed at fouling-prone service.

> The DCX (HT-04 downcomer module), VortiValve (HT-06 inlet device) and AEGIS (HT-07 structural overlay) family members are not standalone contacting decks and so are not sized here -- see `tray_library.NON_DECK_MODULES`.

---
*Generated by `HyTrays/run_tray_comparison.py`. Re-run after editing `CASE` in that file (operating conditions) or the tray parameters in `HyTrays/tray_library.py`. HT-series parameters are derived from the datasheets in `HyTrays/Datasheets/` and are engineering estimates for relative screening.*
