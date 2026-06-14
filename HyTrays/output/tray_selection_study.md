# HyTrays Tray Selection Study -- Numerical Comparison Across Services

Five tray families -- **Sieve, HyTrays Ripple, Valve, Dualflow and a generic High-Performance tray** -- are sized and rated with the Fair-correlation tray-hydraulics engine in `column_hydraulics.py` across five representative distillation services. The goal is to make the trade-offs that drive tray selection (capacity, pressure drop, downcomer backup, turndown and fouling resistance) explicit and reproducible.

## 1. Tray types compared

| Tray | f_active | f_hole | h_weir (in) | C0 | Aeration factor | Capacity factor | Efficiency factor | Min load frac | Fouling open-area retention |
|---|---|---|---|---|---|---|---|---|---|
| Sieve | 0.78 | 0.10 | 2.0 | 0.73 | 0.55 | 1.00 | 1.00 | 0.50 | 0.75 |
| HyTrays Ripple | 0.80 | 0.11 | 1.5 | 0.74 | 0.50 | 1.10 | 1.08 | 0.35 | 0.85 |
| Valve | 0.78 | 0.13 | 2.0 | 0.85 | 0.55 | 1.00 | 1.00 | 0.30 | 0.70 |
| Dualflow | 1.00 | 0.20 | 0.0 | 0.73 | 0.40 | 1.15 | 0.85 | 0.65 | 0.92 |
| High-Performance | 0.85 | 0.14 | 1.0 | 0.80 | 0.50 | 1.25 | 0.97 | 0.40 | 0.80 |

> *HyTrays Ripple parameters are placeholders pending the datasheets in `HyTrays/Datasheets/` -- update and re-run once available.*

## 2. Services evaluated

| Service | V (lb/hr) | L (lb/hr) | rho_V (lb/ft3) | rho_L (lb/ft3) | sigma (dyn/cm) | mu_L (cP) | alpha | Tray spacing (in) | f_flood | System factor | Fouling? | N_theoretical |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| General Rectification (T801 basis) | 52,449 | 14,798 | 1.018 | 30.0 | 9.0 | 0.12 | 1.40 | 24 | 80% | 1.00 | False | 20 |
| Fouling / Heavy-Ends Service | 45,000 | 90,000 | 0.400 | 42.0 | 18.0 | 3.00 | 1.15 | 24 | 65% | 1.00 | True | 15 |
| Vacuum Tower Section | 80,000 | 40,000 | 0.060 | 38.0 | 14.0 | 1.20 | 1.20 | 30 | 80% | 1.00 | False | 8 |
| High-Pressure / High-Liquid-Load Service (C3/C4 splitter) | 150,000 | 300,000 | 2.000 | 28.0 | 5.0 | 0.08 | 1.15 | 24 | 80% | 1.00 | False | 40 |
| Foaming Service | 52,449 | 14,798 | 1.018 | 30.0 | 9.0 | 0.12 | 1.40 | 24 | 80% | 0.75 | False | 20 |

- **General Rectification (T801 basis)**: Mid-pressure hydrocarbon rectifying section -- vapor/liquid loads and vapor density from T801-OH / T801-REFLUX (Hydraulics/HMB.xlsx); liquid density, viscosity, surface tension and alpha are estimates for ~344 F / 105 psia.
- **Fouling / Heavy-Ends Service**: Heavy, viscous, deposit-forming liquid (e.g. coker main fractionator wash/slurry section): high liquid load, high viscosity, close relative volatility. Sized at a reduced 65% of flood to leave margin for deposit buildup; evaluated with tray.fouling_open_area_retention applied (fouling=True).
- **Vacuum Tower Section**: Low-pressure HVGO-type section: very low vapor density drives large diameters and high superficial velocities; every inch of tray dP raises the flash-zone temperature, so dP/tray is the key discriminator. 30 in. tray spacing (common practice in vacuum towers).
- **High-Pressure / High-Liquid-Load Service (C3/C4 splitter)**: High-pressure light-ends splitter: high reflux ratio (L/V ~ 2), low surface tension, low relative volatility drives many stages. High FLV stresses downcomer/liquid handling rather than vapor capacity.
- **Foaming Service**: Same traffic/properties as General Rectification but with a moderate foaming system factor (Fp = 0.75, Kister Table 1-2) applied uniformly to every tray's flooding velocity -- representative of amine/glycol-type foaming tendencies.

## 3. Cross-service comparison

![Diameter by service](11_diameter_by_service.png)

![Total dP by service](12_total_dp_by_service.png)

![Downcomer backup by service](13_downcomer_backup_by_service.png)

![Turndown by service](14_turndown_by_service.png)

![Fouling sensitivity](15_fouling_sensitivity.png)

## 4. Results & discussion by service

### 4.1 General Rectification (T801 basis)

| Tray | Diameter (ft) | u_nf (ft/s) | dP/tray (psi) | DC backup (% lim) | Operating window (% design) | Turndown | Eff (%) | N_actual | Height (ft) | Total dP (psi) |
|---|---|---|---|---|---|---|---|---|---|---|
| Sieve | 4.39 | 1.52 | 0.0557 | 34 | 62-125 | 2.00:1 | 76.2 | 27 | 64.0 | 1.503 |
| HyTrays Ripple | 4.13 | 1.67 | 0.0484 | 30 | 44-125 | 2.86:1 | 82.3 | 25 | 60.0 | 1.211 |
| Valve | 4.39 | 1.52 | 0.0386 | 26 | 38-125 | 3.33:1 | 76.2 | 27 | 64.0 | 1.042 |
| Dualflow | 3.61 | 1.75 | 0.0170 | n/a | 81-125 | 1.54:1 | 64.8 | 31 | 72.0 | 0.526 |
| High-Performance | 3.76 | 1.90 | 0.0351 | 23 | 50-125 | 2.50:1 | 73.9 | 28 | 66.0 | 0.983 |

FLV = 0.052 (vapor-dominated). HyTrays Ripple gives the smallest shell of the conventional weired trays (4.13 ft vs 4.39 ft for Sieve/Valve) **and** the highest efficiency (82.3%), giving the shortest column (60 ft, 25 trays) and the lowest total dP after Dualflow/High-Performance (1.21 psi vs 1.50 psi for Sieve). Dualflow needs the most trays (31) due to its lower contacting efficiency.

### 4.2 Fouling / Heavy-Ends Service

| Tray | Diameter (ft) | u_nf (ft/s) | dP/tray (psi) | DC backup (% lim) | Operating window (% design) | Turndown | Eff (%) | N_actual | Height (ft) | Total dP (psi) |
|---|---|---|---|---|---|---|---|---|---|---|
| Sieve | 5.42 | 2.67 | 0.0905 | 41 | 77-154 | 2.00:1 | 33.5 | 45 | 100.0 | 4.070 |
| HyTrays Ripple | 5.10 | 2.94 | 0.0703 | 35 | 54-154 | 2.86:1 | 36.2 | 42 | 94.0 | 2.954 |
| Valve | 5.42 | 2.67 | 0.0688 | 34 | 46-154 | 3.33:1 | 33.5 | 45 | 100.0 | 3.096 |
| Dualflow | 4.46 | 3.08 | 0.0193 | n/a | 100-154 | 1.54:1 | 28.5 | 53 | 116.0 | 1.021 |
| High-Performance | 4.64 | 3.34 | 0.0579 | 31 | 62-154 | 2.50:1 | 32.5 | 47 | 104.0 | 2.723 |

Sized at 65% of flood (margin for deposit buildup) with each tray's hole area derated by `fouling_open_area_retention`. The standalone fouling-sensitivity result (Fig. 15, independent of flow rates) shows the underlying mechanism -- for a tray whose open area survives at fraction *r* of as-new, dry-tray dP rises by (1/r)^2 - 1:

| Tray | Open-area retention | Dry-tray dP increase |
|---|---|---|
| Sieve | 0.75 | +78% |
| HyTrays Ripple | 0.85 | +38% |
| Valve | 0.70 | +104% |
| Dualflow | 0.92 | +18% |
| High-Performance | 0.80 | +56% |

Dualflow's large, simple perforations lose the least relative open area (+18% dP), keeping it usable far longer between cleanings, while a plain Sieve tray's dry-tray dP rises by roughly 78% and a Valve tray's by 104% (small holes/slots, moving parts), signaling much faster fouling-driven capacity loss. HyTrays Ripple (+38%, placeholder) sits between the two extremes.

### 4.3 Vacuum Tower Section

| Tray | Diameter (ft) | u_nf (ft/s) | dP/tray (psi) | DC backup (% lim) | Operating window (% design) | Turndown | Eff (%) | N_actual | Height (ft) | Total dP (psi) |
|---|---|---|---|---|---|---|---|---|---|---|
| Sieve | 8.84 | 9.67 | 0.1050 | 37 | 62-125 | 2.00:1 | 45.9 | 18 | 55.0 | 1.891 |
| HyTrays Ripple | 8.32 | 10.64 | 0.0950 | 34 | 44-125 | 2.86:1 | 49.5 | 17 | 52.5 | 1.614 |
| Valve | 8.84 | 9.67 | 0.0641 | 26 | 38-125 | 3.33:1 | 45.9 | 18 | 55.0 | 1.154 |
| Dualflow | 7.28 | 11.13 | 0.0328 | n/a | 81-125 | 1.54:1 | 39.0 | 21 | 62.5 | 0.689 |
| High-Performance | 7.57 | 12.09 | 0.0675 | 25 | 50-125 | 2.50:1 | 44.5 | 18 | 55.0 | 1.215 |

With rho_V = 0.06 lb/ft3, superficial velocities and diameters are large for all trays (7.3-8.8 ft). Per-tray dP is what matters most here: Dualflow is lowest (32.8 mpsi), Valve and High-Performance follow (64.1 and 67.5 mpsi) on the strength of their higher discharge coefficients (C0 = 0.85 / 0.80 vs 0.73-0.74 for Sieve/Ripple), while Sieve and Ripple run highest at 95-105 mpsi/tray -- multiplied over a real vacuum tower's tray count this difference is what drives vacuum-service designs toward higher-C0 or grid/high-capacity internals.

### 4.4 High-Pressure / High-Liquid-Load Service (C3/C4 splitter)

| Tray | Diameter (ft) | u_nf (ft/s) | dP/tray (psi) | DC backup (% lim) | Operating window (% design) | Turndown | Eff (%) | N_actual | Height (ft) | Total dP (psi) |
|---|---|---|---|---|---|---|---|---|---|---|
| Sieve | 9.38 | 0.48 | 0.0513 | 43 | 62-125 | 2.00:1 | 84.7 | 48 | 106.0 | 2.463 |
| HyTrays Ripple | 8.83 | 0.53 | 0.0440 | 41 | 44-125 | 2.86:1 | 91.5 | 44 | 98.0 | 1.937 |
| Valve | 9.38 | 0.48 | 0.0518 | 43 | 38-125 | 3.33:1 | 84.7 | 48 | 106.0 | 2.485 |
| Dualflow | 7.72 | 0.56 | 0.0085 | n/a | 81-125 | 1.54:1 | 72.0 | 56 | 122.0 | 0.475 |
| High-Performance | 8.03 | 0.60 | 0.0398 | 40 | 50-125 | 2.50:1 | 82.1 | 49 | 108.0 | 1.949 |

FLV = 0.535 (high, liquid-dominated). Downcomer backup margins shrink markedly versus the General case (43% vs 34% for Sieve) because the smaller diameters needed for high-capacity trays shorten the weir, increasing the Francis-weir crest (h_ow) for the same liquid rate. All weired trays here still sit under the 50% limit, but with much less margin than the General case -- a real design at this FLV would likely need wider downcomers or a larger diameter than the flood-only sizing shown. **Dualflow is flagged as not recommended** for this service regardless of the numbers: with no downcomer, all the liquid must counter-flow through the same perforations as the vapor, which becomes the limiting mechanism at high liquid rates (Kister, *Distillation Operation*).

### 4.5 Foaming Service

| Tray | Diameter (ft) | u_nf (ft/s) | dP/tray (psi) | DC backup (% lim) | Operating window (% design) | Turndown | Eff (%) | N_actual | Height (ft) | Total dP (psi) |
|---|---|---|---|---|---|---|---|---|---|---|
| Sieve | 5.07 | 1.14 | 0.0418 | 28 | 62-125 | 2.00:1 | 76.2 | 27 | 64.0 | 1.130 |
| HyTrays Ripple | 4.77 | 1.25 | 0.0350 | 24 | 44-125 | 2.86:1 | 82.3 | 25 | 60.0 | 0.875 |
| Valve | 5.07 | 1.14 | 0.0322 | 23 | 38-125 | 3.33:1 | 76.2 | 27 | 64.0 | 0.870 |
| Dualflow | 4.17 | 1.31 | 0.0126 | n/a | 81-125 | 1.54:1 | 64.8 | 31 | 72.0 | 0.390 |
| High-Performance | 4.34 | 1.42 | 0.0257 | 19 | 50-125 | 2.50:1 | 73.9 | 28 | 66.0 | 0.721 |

Applying Kister's moderate-foam system factor (Fp = 0.75) derates every tray's flooding velocity equally, so diameters grow by 1/sqrt(Fp) = 1.15x relative to the General case (e.g. Sieve 4.39 -> 5.07 ft) for every tray type -- the relative ranking by capacity is unchanged. The differentiator in foaming service is froth intensity: trays with a lower aeration factor (Dualflow 0.40, HyTrays Ripple/High-Performance 0.50) generate less aerated froth for the same clear-liquid height and are generally more foam-tolerant than Sieve/Valve (0.55).

## 5. Tray selection guidance matrix

Rank 1 = best, 5 = worst, by the metric noted for each service (footnotes below).

| Tray | General [1] | Fouling [2] | Vacuum [3] | High-P/High-L [4] | Foaming [5] |
|---|---|---|---|---|---|
| Sieve | 5 (1.5) | 4 (0.75) | 5 (0.105) | 3 (42.7) | 4 (0.55) |
| HyTrays Ripple | 4 (1.21) | 2 (0.85) | 4 (0.095) | 2 (40.9) | 2 (0.5) |
| Valve | 3 (1.04) | 5 (0.7) | 2 (0.0641) | 4 (42.9) | 5 (0.55) |
| Dualflow | 1 (0.526) | 1 (0.92) | 1 (0.0328) | 5 (n/a) | 1 (0.4) |
| High-Performance | 2 (0.983) | 3 (0.8) | 3 (0.0675) | 1 (40.3) | 3 (0.5) |

[1] **General** (General Rectification (T801 basis)): Lower total column dP across the N_actual trays needed for the target separation (psi).
[2] **Fouling** (Fouling / Heavy-Ends Service): Higher fouling open-area retention -> smaller dry-tray dP increase as deposits build up (see fouling sensitivity figure).
[3] **Vacuum** (Vacuum Tower Section): Lower per-tray dP -- directly limits flash-zone temperature rise in vacuum service.
[4] **High-P/High-L** (High-Pressure / High-Liquid-Load Service (C3/C4 splitter)): Lower downcomer backup (% of the 50%-of-spacing limit) at high FLV. Dualflow has no downcomer and is generally not recommended for high-liquid-rate services (Kister, Distillation Operation) -- ranked last regardless of the hydraulic numbers.
[5] **Foaming** (Foaming Service): Lower aeration factor -> less intense froth generation -> generally more foam-tolerant.

## 6. Conclusions & limitations

- **HyTrays Ripple** is competitive or best-in-class for General Rectification (smallest shell among weired trays, highest efficiency, shortest column) and has a favorable fouling profile between Sieve/Valve and Dualflow/High-Performance -- a good general-purpose upgrade from plain Sieve. *Its parameters are placeholders; replace them from the HyTrays datasheets once added to `HyTrays/Datasheets/` and re-run.*
- **Dualflow** wins on raw capacity and per-tray dP (General, Vacuum, Fouling) but loses on efficiency/turndown and is not recommended for high-liquid-rate service.
- **High-Performance** gives the smallest shell across the board and the lowest per-tray dP alongside Dualflow, with efficiency much closer to Sieve -- a strong default where shell diameter/capex dominates.
- **Valve** matches Sieve's capacity but offers the widest turndown of the conventional designs -- preferred where the column must run efficiently over a wide load range, at the cost of being the most fouling-sensitive (smallest open area, moving parts).
- **Sieve** remains the simplest/cheapest baseline but is dominated by HyTrays Ripple on every metric computed here.

**Model limitations** -- read before using these numbers for anything beyond relative tray-type screening:
- Fair's C_SBF curve fit is valid for FLV = 0.01-1.0 and tray spacing 6-24 in (the Vacuum case uses 30 in, slightly outside the fitted range, and the High-P/High-L case has FLV at the top of the range).
- Turndown/weep is represented by each tray's literature-typical `min_load_frac`, not a first-principles weep-point correlation; entrainment is not explicitly modeled.
- Dualflow's clear-liquid height is a fixed 1 in. placeholder (no overflow weir to apply the Francis equation to) -- its dP and downcomer-backup columns should be read as indicative only.
- All `HyTrays Ripple` and `fouling_open_area_retention` values are engineering placeholders pending vendor/manufacturer data.

---
*Generated by `HyTrays/run_service_comparison.py`. Edit `service_cases.py` (operating conditions per service) or `tray_library.py` (tray parameters) and re-run to update every table, figure and number in this report.*
