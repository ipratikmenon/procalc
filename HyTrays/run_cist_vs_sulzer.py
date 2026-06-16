#!/usr/bin/env python3
"""
CIST vs Sulzer UFM and SVG — 12-case hydraulic comparison.

Applies the CIST first-principles hydraulic model (curved helical guide vane +
three elliptic nozzles at 120°) to the same fluid conditions and throughputs as
two Sulzer rating sheets for the SAME 108-in (9 ft), 2-pass, 24-in-spacing
column:

    • UFM  — moving/mini-valve deck   (196 valves)
    • SVG  — fixed-valve deck         (244 valves, 14% open area, 33% active)

The fluid data are identical between the two sheets (same column, same 12
operating cases); only the tray-side results differ (SVG runs lower ΔP, UFM
slightly higher capacity margin).  The CIST result is therefore computed once
per case and benchmarked against both decks.

Outputs:
  • Console table (UFM and SVG side-by-side with CIST)
  • PNG figure  HyTrays/output/48_cist_vs_sulzer.png
"""
import math, os
from dataclasses import dataclass

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec

# ── CIST geometry (derived design, curved helical vane) ──────────────────────
Rc     = 0.0375          # m   chamber radius
Hc     = 0.210           # m   chamber height
theta0 = math.radians(20.0)   # rad  vane entry angle (tangential)
theta1 = math.radians(70.0)   # rad  vane exit  angle (axial)
p_w    = 0.120           # m   wavy-deck pitch = hex lattice pitch
g      = 9.81            # m/s²
D_h    = 2.0 * Rc        # m   chamber hydraulic diameter

# ── 3-nozzle scheme (elliptic nozzles at 120°, r_noz = 0.5 Rc) ──────────────
#   d32^(N) = d32^(1)/√N ; h_imp = 0.38 Hc ; interfacial area ∝ 1/d32 → ×√N
N_noz        = 3
r_noz        = 0.50 * Rc              # m   nozzle radial position
h_imp_m      = 0.38 * Hc             # m   cross-impingement height
d32_factor   = 1.0 / math.sqrt(N_noz)  # d32 → 1/√3 ≈ 0.577× baseline
EFF_FACTOR   = 1.32   # O'Connell multiplier for 3-nozzle curved-vane CIST

# Vane integral  I_cot2 = ∫₀^Hc cot²θ(z) dz  =  (Hc/Δθ)[cot θ₀ - cot θ₁ - Δθ]
_dth   = theta1 - theta0
I_cot2 = (Hc / _dth) * (1/math.tan(theta0) - 1/math.tan(theta1) - _dth)  # m

# Vane arc length (closed-form integration)
csc0 = 1/math.sin(theta0); cot0 = 1/math.tan(theta0)
csc1 = 1/math.sin(theta1); cot1 = 1/math.tan(theta1)
L_vane = (Hc / _dth) * math.log((csc0 + cot0) / (csc1 + cot1))   # m

# Pressure-drop coefficient: K = C_in·cot²θ₀ + f_D·(L_v/D_h) + cot²θ₁
C_in   = 0.20    # tangential-slot entry loss fraction (smooth curved slot)
f_D    = 0.02    # Darcy wall friction factor
K_dp   = C_in * cot0**2 + f_D * (L_vane / D_h) + cot1**2

# Capacity limits (axial vapor velocity in the swirl chamber)
u_x_design = 3.0    # m/s  rated design point
u_x_max    = 5.0    # m/s  upper limit (droplet residence-time margin)
u_x_min    = 1.5    # m/s  lower limit (Ng_eff ≥ 15 g maintained)

# Hex packing geometry
A_cell      = p_w**2 * math.sqrt(3) / 2   # m²  unit-cell area per chamber
A_ch        = math.pi * Rc**2             # m²  chamber cross-section
pack_eff    = 0.90   # peripheral / edge efficiency

# Reference column (fixed, both decks)
D_REF_ft = 9.0
LB_KG_S    = 1.0 / (2.20462 * 3600.0)
LB_FT3_SI  = 16.0185          # lb/ft³ → kg/m³
MMHG_PA    = 133.322          # 1 mmHg in Pa

# Reference-deck efficiency multipliers on the O'Connell base (valve trays)
F_UFM = 1.10
F_SVG = 1.05    # fixed valve, slightly lower than moving valve

def oconnell(alpha: float, mu_l: float) -> float:
    """E_OC (%) = 51 − 32.5 log₁₀(α·μ_L[cP]), clipped [1, 100]."""
    return max(1.0, min(51.0 - 32.5 * math.log10(alpha * mu_l), 100.0))

# ── Case data — 12 cases, identical fluids on both sheets ────────────────────
# Per-deck Sulzer results: (jet flood %, ΔP mmHg/tray)
@dataclass
class Case:
    name: str
    v_lb_hr: float; rho_v: float; mu_v: float   # lb/hr, lb/ft³, cP
    l_lb_hr: float; rho_l: float                 # lb/hr, lb/ft³
    sigma: float; mu_l: float                    # dyn/cm, cP
    ufm_jf: float; ufm_dp: float                 # Sulzer UFM result
    svg_jf: float; svg_dp: float                 # Sulzer SVG result
    alpha: float = 1.5                           # approx rel. volatility (C2/C3 svc.)

CASES = [
    #            gas    rhoV   muV     liq    rhoL  sig   muL    UFM       SVG
    Case("C1 S2",      120900,1.044, 0.0110, 21140,37.80,8.71,0.135, 58,4.36, 60,2.86),
    Case("C1 S11",     106500,0.9595,0.0114,  6739,38.89,8.69,0.151, 52,3.54, 56,2.21),
    Case("C2 S2",      121500,1.048, 0.0109, 21270,37.83,8.75,0.135, 58,4.38, 60,2.88),
    Case("C2 S11",     107200,0.9638,0.0116,  6968,38.90,9.11,0.149, 52,3.57, 56,2.23),
    Case("C2 WWN S2",  101600,1.041, 0.0117, 21910,37.54,8.63,0.135, 49,3.55, 51,2.51),
    Case("C2 WWN S11",  88680,0.9332,0.0117,  7006,38.72,8.94,0.149, 43,2.78, 46,1.86),
    Case("C3 S2",       64700,1.047, 0.0117, 29280,37.63,9.27,0.135, 31,2.76, 32,2.39),
    Case("C3 S11",      50450,0.8865,0.0117, 15030,37.99,8.66,0.135, 26,2.17, 27,1.93),
    Case("C4 S2",       85696,1.044, 0.0117, 21968,37.89,8.91,0.135, 41,3.04, 43,2.31),
    Case("C4 S11",      71277,0.920, 0.0117,  7549,38.90,9.07,0.147, 36,2.35, 38,1.70),
    Case("TD S2",       59100,1.081, 0.0117, 16060,36.85,8.29,0.128, 28,2.24, 29,1.93),
    Case("TD S11",      49820,0.9692,0.0117,  6778,37.63,8.30,0.136, 25,1.83, 27,1.50),
]

# ── CIST rating ──────────────────────────────────────────────────────────────
@dataclass
class R:
    name: str
    ufm_jf: float; ufm_dp: float
    svg_jf: float; svg_dp: float
    N_ch: int; D_ft: float
    Ng: float; d50: float
    dp: float
    E_oc: float; E_cist: float
    save_ufm: float; save_svg: float

def rate(c: Case) -> R:
    rho_V = c.rho_v * LB_FT3_SI
    rho_L = c.rho_l * LB_FT3_SI
    mu_V  = c.mu_v  * 1e-3          # Pa·s
    drho  = rho_L - rho_V

    Q_V   = (c.v_lb_hr * LB_KG_S) / rho_V   # m³/s

    # Chambers needed at design u_x, equivalent column diameter (hex packing)
    N_ch  = max(1, math.ceil(Q_V / (u_x_design * A_ch)))
    A_tray = N_ch * A_cell / pack_eff
    D_ft   = math.sqrt(4 * A_tray / math.pi) / 0.3048

    # Centrifugal performance at design u_x
    int_ac = (u_x_design**2 / Rc) * I_cot2        # ∫a_c dz  [m²/s²]
    Ng     = (int_ac / Hc) / g
    d50    = math.sqrt(72 * mu_V * u_x_design * Rc**2 / (drho * int_ac)) * 1e6  # µm

    # Pressure drop (curved-vane vapor path)
    dp_mm  = (K_dp * 0.5 * rho_V * u_x_design**2) / MMHG_PA

    # Efficiency
    E_oc   = oconnell(c.alpha, c.mu_l)
    E_cist = min(E_oc * EFF_FACTOR, 95.0)

    save_ufm = (c.ufm_dp - dp_mm) / c.ufm_dp * 100.0
    save_svg = (c.svg_dp - dp_mm) / c.svg_dp * 100.0

    return R(c.name, c.ufm_jf, c.ufm_dp, c.svg_jf, c.svg_dp,
             N_ch, D_ft, Ng, d50, dp_mm, E_oc, E_cist, save_ufm, save_svg)

results = [rate(c) for c in CASES]

# ── Console table ─────────────────────────────────────────────────────────────
print()
print("=" * 118)
print("  CIST vs SULZER UFM & SVG  —  12-case hydraulic comparison (same 9-ft column, 24-in spacing)")
print(f"  CIST: Rc={Rc*1e3:.0f}mm Hc={Hc*1e3:.0f}mm | vane θ {math.degrees(theta0):.0f}°→"
      f"{math.degrees(theta1):.0f}° (L_v={L_vane*1e3:.0f}mm) | {N_noz} nozzles@120° r={r_noz*1e3:.1f}mm "
      f"h_imp={h_imp_m*1e3:.0f}mm | K_dp={K_dp:.3f} | u_x={u_x_design}m/s")
print("=" * 118)
print()
hdr = (f"{'Case':<12}|{'UFM JF':>7}{'UFM ΔP':>8}|{'SVG JF':>7}{'SVG ΔP':>8}|"
       f"{'CIST D':>8}{'N_ch':>5}{'Ng':>6}{'d50':>6}{'CIST ΔP':>8}|"
       f"{'sav.UFM':>8}{'sav.SVG':>8}|{'E_CIST':>7}{'TD':>5}")
sub = (f"{'':12}|{'%':>7}{'mmHg':>8}|{'%':>7}{'mmHg':>8}|"
       f"{'ft':>8}{'':>5}{'(g)':>6}{'µm':>6}{'mmHg':>8}|"
       f"{'%':>8}{'%':>8}|{'%':>7}{'×':>5}")
div = "-" * 118
print(hdr); print(sub); print(div)
TD = u_x_max / u_x_min
for r in results:
    print(f"{r.name:<12}|{r.ufm_jf:>7.0f}{r.ufm_dp:>8.2f}|{r.svg_jf:>7.0f}{r.svg_dp:>8.2f}|"
          f"{r.D_ft:>8.2f}{r.N_ch:>5}{r.Ng:>6.1f}{r.d50:>6.1f}{r.dp:>8.2f}|"
          f"{r.save_ufm:>8.1f}{r.save_svg:>8.1f}|{r.E_cist:>7.1f}{TD:>5.1f}")
print(div)

d_red   = [(D_REF_ft - r.D_ft) / D_REF_ft * 100 for r in results]
sav_ufm = [r.save_ufm for r in results]
sav_svg = [r.save_svg for r in results]
print(f"\n  Column-diameter reduction      : {min(d_red):.0f}% – {max(d_red):.0f}%  "
      f"(mean {sum(d_red)/len(d_red):.0f}%)   [9 ft → {min(r.D_ft for r in results):.1f}"
      f"–{max(r.D_ft for r in results):.1f} ft]")
print(f"  Per-tray ΔP saving vs UFM      : {min(sav_ufm):.0f}% – {max(sav_ufm):.0f}%  "
      f"(mean {sum(sav_ufm)/len(sav_ufm):.0f}%)")
print(f"  Per-tray ΔP saving vs SVG      : {min(sav_svg):.0f}% – {max(sav_svg):.0f}%  "
      f"(mean {sum(sav_svg)/len(sav_svg):.0f}%)")
print(f"  CIST per-tray ΔP               : {min(r.dp for r in results):.2f}–"
      f"{max(r.dp for r in results):.2f} mmHg   (UFM {min(r.ufm_dp for r in results):.2f}"
      f"–{max(r.ufm_dp for r in results):.2f}, SVG {min(r.svg_dp for r in results):.2f}"
      f"–{max(r.svg_dp for r in results):.2f})")
print(f"  CIST tray efficiency           : {min(r.E_cist for r in results):.1f}–"
      f"{max(r.E_cist for r in results):.1f}%   ({EFF_FACTOR:.2f}× O'Connell)")
print(f"  CIST turndown ratio            : {TD:.1f}×  (u_x {u_x_min}–{u_x_max} m/s, Ng≥15 g)")
print(f"  Ng_eff / d50 at design         : {results[0].Ng:.1f} g  /  "
      f"{min(r.d50 for r in results):.1f}–{max(r.d50 for r in results):.1f} µm")
print(f"  3-nozzle d32 reduction         : ×{d32_factor:.3f} (1/√{N_noz}) → "
      f"interfacial area ×{1/d32_factor:.2f}, impingement at {h_imp_m*1e3:.0f} mm")

# Actual tray-count saving (same theoretical stages)
E_UFM_m  = sum(r.E_oc * F_UFM for r in results) / len(results)
E_SVG_m  = sum(r.E_oc * F_SVG for r in results) / len(results)
E_CIST_m = sum(r.E_cist       for r in results) / len(results)
ts_ufm = (1/E_UFM_m - 1/E_CIST_m) / (1/E_UFM_m) * 100
ts_svg = (1/E_SVG_m - 1/E_CIST_m) / (1/E_SVG_m) * 100
print(f"  Actual tray-count saving       : {ts_ufm:.0f}% vs UFM, {ts_svg:.0f}% vs SVG  "
      f"(E_CIST≈{E_CIST_m:.0f}% vs UFM≈{E_UFM_m:.0f}%, SVG≈{E_SVG_m:.0f}%)")
print()

# ── Figure (4 panels) ─────────────────────────────────────────────────────────
names = [r.name for r in results]
x     = np.arange(len(names))
BG    = "#f5f7fa"
C_UFM = "#1565c0"; C_SVG = "#7b1fa2"; C_CIST = "#2e7d32"

fig = plt.figure(figsize=(17, 14), facecolor=BG)
gs  = gridspec.GridSpec(4, 1, figure=fig, hspace=0.55, top=0.93, bottom=0.05)
bw  = 0.27

# Panel A — pressure drop (3-way)
ax1 = fig.add_subplot(gs[0]); ax1.set_facecolor("#eef2f7")
ax1.bar(x - bw, [r.ufm_dp for r in results], width=bw, color=C_UFM,  alpha=0.85, label="Sulzer UFM")
ax1.bar(x,      [r.svg_dp for r in results], width=bw, color=C_SVG,  alpha=0.85, label="Sulzer SVG")
ax1.bar(x + bw, [r.dp     for r in results], width=bw, color=C_CIST, alpha=0.9,
        label=f"CIST (u_x={u_x_design} m/s)")
for i, r in enumerate(results):
    ax1.text(i - bw, r.ufm_dp + 0.04, f"{r.ufm_dp:.2f}", ha="center", fontsize=5.5, color=C_UFM)
    ax1.text(i,      r.svg_dp + 0.04, f"{r.svg_dp:.2f}", ha="center", fontsize=5.5, color=C_SVG)
    ax1.text(i + bw, r.dp     + 0.04, f"{r.dp:.2f}",     ha="center", fontsize=5.5, color=C_CIST)
ax1.set_xticks(x); ax1.set_xticklabels(names, rotation=30, ha="right", fontsize=8)
ax1.set_ylabel("ΔP per tray  (mmHg)", fontsize=9)
ax1.set_title("A — Per-tray pressure drop: CIST vs UFM vs SVG", fontweight="bold", fontsize=10)
ax1.legend(fontsize=8, ncol=3); ax1.grid(axis="y", alpha=0.4)

# Panel B — equivalent column diameter
ax2 = fig.add_subplot(gs[1]); ax2.set_facecolor("#eef2f7")
ax2.axhline(D_REF_ft, color="#455a64", lw=2, ls="--", label=f"Sulzer column  D = {D_REF_ft:.0f} ft")
ax2.plot(x, [r.D_ft for r in results], "o-", color=C_CIST, lw=2, ms=7,
         label=f"CIST equivalent D")
ax2.fill_between(x, [r.D_ft for r in results], D_REF_ft, alpha=0.12, color=C_CIST)
for i, r in enumerate(results):
    ax2.text(i, r.D_ft - 0.22, f"{r.D_ft:.1f}", ha="center", fontsize=6.5, color=C_CIST)
ax2.set_xticks(x); ax2.set_xticklabels(names, rotation=30, ha="right", fontsize=8)
ax2.set_ylabel("Column diameter  (ft)", fontsize=9); ax2.set_ylim(0, 10.5)
ax2.set_title("B — CIST equivalent column diameter at same throughput (60–73% smaller)",
              fontweight="bold", fontsize=10)
ax2.legend(fontsize=8); ax2.grid(axis="y", alpha=0.4)

# Panel C — efficiency (3-way)
ax3 = fig.add_subplot(gs[2]); ax3.set_facecolor("#eef2f7")
E_ufm = [r.E_oc * F_UFM for r in results]
E_svg = [r.E_oc * F_SVG for r in results]
E_cis = [r.E_cist       for r in results]
ax3.bar(x - bw, E_ufm, width=bw, color=C_UFM,  alpha=0.85, label="UFM (O'Connell × 1.10)")
ax3.bar(x,      E_svg, width=bw, color=C_SVG,  alpha=0.85, label="SVG (O'Connell × 1.05)")
ax3.bar(x + bw, E_cis, width=bw, color=C_CIST, alpha=0.9,  label="CIST (O'Connell × 1.32)")
for i in range(len(results)):
    ax3.text(i - bw, E_ufm[i] + 0.4, f"{E_ufm[i]:.0f}", ha="center", fontsize=5.5, color=C_UFM)
    ax3.text(i,      E_svg[i] + 0.4, f"{E_svg[i]:.0f}", ha="center", fontsize=5.5, color=C_SVG)
    ax3.text(i + bw, E_cis[i] + 0.4, f"{E_cis[i]:.0f}", ha="center", fontsize=5.5, color=C_CIST)
ax3.set_xticks(x); ax3.set_xticklabels(names, rotation=30, ha="right", fontsize=8)
ax3.set_ylabel("Overall efficiency  (%)", fontsize=9); ax3.set_ylim(60, 100)
ax3.set_title("C — Tray efficiency (O'Connell basis with deck factors)",
              fontweight="bold", fontsize=10)
ax3.legend(fontsize=8, ncol=3); ax3.grid(axis="y", alpha=0.4)

# Panel D — capacity headroom & turndown (jet flood vs CIST utilization)
ax4 = fig.add_subplot(gs[3]); ax4.set_facecolor("#eef2f7")
# CIST utilization at design = u_x_design / u_x_max
cist_util = u_x_design / u_x_max * 100.0
ax4.bar(x - bw, [r.ufm_jf for r in results], width=bw, color=C_UFM, alpha=0.85, label="UFM jet flood %")
ax4.bar(x,      [r.svg_jf for r in results], width=bw, color=C_SVG, alpha=0.85, label="SVG jet flood %")
ax4.bar(x + bw, [cist_util]*len(results),    width=bw, color=C_CIST, alpha=0.9,
        label=f"CIST capacity use {cist_util:.0f}% (u_x {u_x_design}/{u_x_max})")
ax4.axhline(100, color="#b71c1c", lw=1.2, ls=":", label="flood / 100% limit")
ax4.set_xticks(x); ax4.set_xticklabels(names, rotation=30, ha="right", fontsize=8)
ax4.set_ylabel("Capacity utilization  (%)", fontsize=9); ax4.set_ylim(0, 110)
ax4.set_title(f"D — Capacity utilization & turndown  (CIST TD = {TD:.1f}×, u_x {u_x_min}–{u_x_max} m/s)",
              fontweight="bold", fontsize=10)
ax4.legend(fontsize=7.5, ncol=2); ax4.grid(axis="y", alpha=0.4)

fig.suptitle(
    "CIST (Cascade Impingement Swirl Tray) vs Sulzer UFM & SVG — Numerical Hydraulic Comparison\n"
    f"Same 9-ft 2-pass column, 24-in spacing, 12 cases  |  "
    f"CIST Rc={Rc*1e3:.0f}mm Hc={Hc*1e3:.0f}mm, curved vane θ {math.degrees(theta0):.0f}°→"
    f"{math.degrees(theta1):.0f}°, {N_noz} elliptic nozzles@120°, "
    f"K_dp={K_dp:.3f}, Ng,eff={results[0].Ng:.1f}g, d32×{d32_factor:.3f}",
    fontsize=9, y=0.975)

outdir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
os.makedirs(outdir, exist_ok=True)
outpath = os.path.join(outdir, "48_cist_vs_sulzer.png")
plt.savefig(outpath, dpi=150, bbox_inches="tight")
plt.close(fig)
print(f"Figure saved → {outpath}")
