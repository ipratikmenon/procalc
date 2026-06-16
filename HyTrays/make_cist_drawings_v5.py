#!/usr/bin/env python3
"""
CIST drawings v5 — Figure 49: enhanced inner-wall features.

Panels:
  (a) Cross-section elevation of cylinder at wavy-deck crest:
       saddle-cut base, drainage slots, spiral grooves, liquid film,
       3 nozzles in deck, cross-chevron crown.
  (b) Plan view of cylinder base: crest saddle, slot positions,
       3 nozzle holes at 120°.
  (c) Quantitative performance table: four feature configurations.
  (d) Film thickness & groove enhancement derivation summary.
"""
import numpy as np
import math
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec
from matplotlib.patches import FancyArrowPatch, Wedge, Arc, FancyBboxPatch
from matplotlib.lines import Line2D

# ── palette ──────────────────────────────────────────────────────────────────
BG       = "#f5f7fa"
C_WALL   = "#4a90d9"
C_DECK   = "#90a4ae"
C_JET    = "#0277bd"
C_VAPOR  = "#e53935"
C_FILM   = "#29b6f6"
C_GROOVE = "#6a1b9a"
C_SLOT   = "#f57c00"
C_CROWN  = "#1a5276"
C_VANE   = "#2e7d32"
C_NOZZLE = "#00b0ff"
C_TXT    = "#263238"

# ── geometry (mm — actual dimensions, displayed scaled) ─────────────────────
Rc_mm  = 37.5
Hc_mm  = 210.0
alpha_deck = math.radians(27.5)   # wave corrugation half-angle
pw_mm  = 120.0                    # wave pitch

# Base saddle: h_cut(phi) = Rc |sin phi| tan(alpha_deck)
h_cut_max_mm = Rc_mm * math.tan(alpha_deck)   # = 19.5 mm  (at phi=90°)

# Groove parameters
d_g = 1.0    # groove depth mm
w_g = 2.0    # groove width mm
s_g = 8.0    # groove pitch mm

# Slot: 2 slots at phi=+/-90° (lowest base points)
slot_h_mm = 5.0    # slot height mm
slot_w_mm = 10.0   # slot width mm

# ── Figure layout ─────────────────────────────────────────────────────────────
fig = plt.figure(figsize=(18, 13), facecolor=BG)
gs  = gridspec.GridSpec(2, 2, figure=fig, hspace=0.45, wspace=0.35,
                        top=0.91, bottom=0.05, left=0.04, right=0.97)

# ══════════════════════════════════════════════════════════════════════════════
# Panel (a) — Elevation cross-section
# ══════════════════════════════════════════════════════════════════════════════
ax1 = fig.add_subplot(gs[0, 0])
ax1.set_facecolor("#eef5fb")
ax1.set_aspect("equal")
ax1.axis("off")

# Scale: 1mm = 0.01 plot units → chamber height = 2.1, radius = 0.375
sc = 0.01

Rc  = Rc_mm  * sc
Hc  = Hc_mm  * sc
h_cut_max = h_cut_max_mm * sc
sl_h = slot_h_mm * sc
sl_w = slot_w_mm * sc

# --- Wave deck below ---
deck_y0 = 0.0    # crest apex
deck_extent = 1.5
for sign in [+1, -1]:
    x_face = np.array([0, sign * deck_extent])
    y_face = np.array([deck_y0, deck_y0 - deck_extent * math.tan(alpha_deck)])
    ax1.plot(x_face * sc * 1000 / Rc_mm * Rc, y_face,     # use actual x in plot units
             color=C_DECK, lw=2, zorder=2)
# fill wave
x_fill = np.array([-1.5, 0, 1.5, 1.5, -1.5]) * sc * 1000/Rc_mm * Rc
y_fill = np.array([-1.5*math.tan(alpha_deck), 0, -1.5*math.tan(alpha_deck),
                   -0.5, -0.5])
ax1.fill(x_fill, y_fill, color=C_DECK, alpha=0.25, zorder=1)

# saddle-cut base of cylinder (profile follows wave deck)
phi_arr = np.linspace(-math.pi/2, math.pi/2, 300)
x_base_L = (-Rc + Rc*(1-np.cos(phi_arr)))   # left wall inner
x_base   = np.concatenate([np.linspace(-Rc, Rc, 201)])
y_base   = np.array([-abs(x)*math.tan(alpha_deck) for x in x_base])  # follows deck
ax1.plot(x_base, y_base, color=C_WALL, lw=2.5, zorder=3)

# cylinder walls (vertical)
wall_top = Hc
ax1.plot([-Rc, -Rc], [y_base[0], wall_top], color=C_WALL, lw=3, zorder=3)
ax1.plot([ Rc,  Rc], [y_base[-1], wall_top], color=C_WALL, lw=3, zorder=3)

# cylinder top
ax1.plot([-Rc, Rc], [wall_top, wall_top], color=C_CROWN, lw=3, zorder=3)

# ── Spiral grooves on inner wall (left wall shown) ────────────────────────
groove_y = np.linspace(0.05, wall_top - 0.1, 14)
for gy in groove_y:
    # small V notch on inner surface of left wall
    ax1.annotate("", xy=(-Rc + 0.015, gy), xytext=(-Rc, gy),
                 arrowprops=dict(arrowstyle="-", color=C_GROOVE, lw=0.8))
groove_label_x, groove_label_y = -Rc - 0.38, wall_top*0.55
ax1.annotate("Spiral grooves\n(1 mm deep,\n8 mm pitch)",
             xy=(-Rc + 0.005, wall_top*0.55),
             xytext=(groove_label_x - 0.05, groove_label_y),
             fontsize=7.5, color=C_GROOVE, ha="right",
             arrowprops=dict(arrowstyle="->", color=C_GROOVE, lw=1))

# ── Liquid film on inner wall (left) ─────────────────────────────────────
film_x = -Rc + 0.02
film_y = np.linspace(0.08, wall_top - 0.08, 60)
ax1.fill_betweenx(film_y, film_x, film_x + 0.025, alpha=0.6, color=C_FILM, zorder=4)
ax1.annotate("Liquid film\n(δ≈70 µm, turbulent\nRe_film≈4100)",
             xy=(-Rc+0.025, wall_top*0.30),
             xytext=(-Rc - 0.38, wall_top*0.25),
             fontsize=7.5, color=C_FILM, ha="right",
             arrowprops=dict(arrowstyle="->", color=C_FILM, lw=1))

# Film flow arrow (downward)
ax1.annotate("", xy=(-Rc+0.03, Hc*0.12), xytext=(-Rc+0.03, Hc*0.40),
             arrowprops=dict(arrowstyle="-|>", color=C_FILM, lw=1.4))

# ── 3 nozzle openings in deck ─────────────────────────────────────────────
noz_r  = 0.5 * Rc
noz_dx = noz_r * math.sin(math.radians(60))
noz_dy = -noz_r * math.cos(math.radians(30)) * math.tan(alpha_deck)
noz_xs = [0, noz_r*math.sin(math.radians(0))]   # in cross-section: one center, one at +60° projected
# Show 3 nozzle holes in the base profile (visible as gaps / rectangles)
noz_positions_x = [0.0, 0.18, -0.18]
for nx in noz_positions_x:
    ny = -abs(nx) * math.tan(alpha_deck)
    ax1.add_patch(mpatches.Ellipse((nx, ny), 0.05, 0.018,
                  fc=C_NOZZLE, ec="k", lw=0.5, zorder=5))

ax1.annotate("3 elliptic nozzles\n@ 120°, r=0.5Rc",
             xy=(0.18, -0.18*math.tan(alpha_deck) - 0.01),
             xytext=(0.55, -0.22),
             fontsize=7.5, color=C_NOZZLE, ha="left",
             arrowprops=dict(arrowstyle="->", color=C_NOZZLE, lw=1))

# ── Vapor swirl arrows ────────────────────────────────────────────────────
for vy in [0.4, 0.85, 1.3, 1.75]:
    ax1.annotate("", xy=(-0.12, vy+0.12), xytext=(0.14, vy),
                 arrowprops=dict(arrowstyle="-|>", color=C_VAPOR, lw=1.2))

ax1.annotate("Swirling\nvapor", xy=(0.05, Hc*0.5),
             xytext=(0.25, Hc*0.5), fontsize=7.5, color=C_VAPOR, ha="left")

# ── Liquid jets from nozzles ─────────────────────────────────────────────
for nx in noz_positions_x:
    ny0 = -abs(nx)*math.tan(alpha_deck)
    ax1.annotate("", xy=(nx*1.1, ny0 + 0.35), xytext=(nx, ny0 + 0.01),
                 arrowprops=dict(arrowstyle="-|>", color=C_JET, lw=1.5))

ax1.annotate("Liquid jets\n(3 @ 120°)", xy=(0, 0.28),
             xytext=(0.35, 0.15), fontsize=7.5, color=C_JET, ha="left",
             arrowprops=dict(arrowstyle="->", color=C_JET, lw=1))

# ── Drainage slots ────────────────────────────────────────────────────────
# Slots at x=±Rc (the side where base is lowest)
slot_x_pos = Rc - 0.02
slot_y_pos = -Rc * math.tan(alpha_deck)
ax1.add_patch(mpatches.Rectangle((Rc-sl_w/2, slot_y_pos),
              sl_w, sl_h, fc=C_SLOT, ec="k", lw=0.8, zorder=6, alpha=0.9))
ax1.add_patch(mpatches.Rectangle((-Rc-sl_w/2, -Rc*math.tan(alpha_deck)),
              sl_w, sl_h, fc=C_SLOT, ec="k", lw=0.8, zorder=6, alpha=0.9))

ax1.annotate("Drainage slots\n(5×10 mm, 2 off)\n— film exits here",
             xy=(Rc, slot_y_pos + sl_h/2),
             xytext=(Rc + 0.32, slot_y_pos - 0.05),
             fontsize=7.5, color=C_SLOT, ha="left",
             arrowprops=dict(arrowstyle="->", color=C_SLOT, lw=1))

# ── Cross-chevron crown ───────────────────────────────────────────────────
# Two opposing angled vanes at the crown
chev_y = wall_top
cx = np.linspace(-Rc*0.8, Rc*0.8, 100)
# Chevron 1: /  direction
cy1 = chev_y + 0.04 * (cx / (Rc*0.8))
ax1.plot(cx, cy1, color=C_CROWN, lw=2, zorder=6)
# Chevron 2: \  direction
cy2 = chev_y - 0.04 * (cx / (Rc*0.8))
ax1.plot(cx, cy2, color=C_CROWN, lw=2, zorder=6)
ax1.annotate("Cross-chevron crown\n(opposing vanes,\nentrainment capture)",
             xy=(0, wall_top + 0.04),
             xytext=(0.42, wall_top + 0.08),
             fontsize=7.5, color=C_CROWN, ha="left",
             arrowprops=dict(arrowstyle="->", color=C_CROWN, lw=1))

# ── Vapor exit arrows ─────────────────────────────────────────────────────
for ex in [-0.15, 0.15]:
    ax1.annotate("", xy=(ex, wall_top + 0.18), xytext=(ex, wall_top + 0.02),
                 arrowprops=dict(arrowstyle="-|>", color=C_VAPOR, lw=1.5))

ax1.text(0, wall_top + 0.22, "Clean vapor exit", ha="center",
         fontsize=7.5, color=C_VAPOR)

# ── Sump label ────────────────────────────────────────────────────────────
ax1.text(0, -0.35, "Liquid\nsump", ha="center", fontsize=7.5,
         color=C_JET, style="italic")

# ── Weld symbol ───────────────────────────────────────────────────────────
ax1.text(-Rc - 0.05, y_base[0]+0.02, "⊿ weld\n(saddle\ncut)", ha="right",
         fontsize=6.5, color=C_TXT)

# ── Dimension lines ───────────────────────────────────────────────────────
ax1.annotate("", xy=(Rc, -0.5), xytext=(-Rc, -0.5),
             arrowprops=dict(arrowstyle="<->", color=C_TXT, lw=1))
ax1.text(0, -0.53, f"2Rc = {2*Rc_mm:.0f} mm", ha="center", fontsize=7, color=C_TXT)
ax1.annotate("", xy=(Rc + 0.28, Hc), xytext=(Rc + 0.28, 0),
             arrowprops=dict(arrowstyle="<->", color=C_TXT, lw=1))
ax1.text(Rc + 0.30, Hc/2, f"Hc={Hc_mm:.0f} mm", ha="left", fontsize=7, color=C_TXT)

# ── Saddle cut annotation ─────────────────────────────────────────────────
ax1.annotate(f"Saddle cut\nh_max = {h_cut_max_mm:.0f} mm",
             xy=(-Rc, -h_cut_max),
             xytext=(-Rc - 0.05, -h_cut_max - 0.12),
             fontsize=7, color=C_DECK, ha="right",
             arrowprops=dict(arrowstyle="->", color=C_DECK, lw=0.8))

ax1.set_xlim(-1.1, 1.0); ax1.set_ylim(-0.65, wall_top + 0.35)
ax1.set_title("(a) Elevation cross-section — enhanced CIST features",
              fontsize=9, fontweight="bold", pad=6)

# ══════════════════════════════════════════════════════════════════════════════
# Panel (b) — Plan view: cylinder base on wave crest
# ══════════════════════════════════════════════════════════════════════════════
ax2 = fig.add_subplot(gs[1, 0])
ax2.set_facecolor("#eef5fb")
ax2.set_aspect("equal")
ax2.axis("off")

# Full cylinder circle
theta_c = np.linspace(0, 2*math.pi, 400)
ax2.plot(np.cos(theta_c)*Rc_mm, np.sin(theta_c)*Rc_mm,
         color=C_WALL, lw=2.5, zorder=3)

# Wave crest runs E-W (horizontal in plan view)
# The saddle-cut depth: h_cut = Rc|sin(phi)| tan(alpha)
# Shown as elliptic shading: darker where base is lower
for phi_deg in range(0, 360, 10):
    phi = math.radians(phi_deg)
    h = abs(math.sin(phi)) * h_cut_max_mm
    x_pt = Rc_mm * math.cos(phi)
    y_pt = Rc_mm * math.sin(phi)
    ax2.scatter(x_pt, y_pt, c=[[0.4+0.5*(1-h/h_cut_max_mm),
                                  0.6+0.3*(1-h/h_cut_max_mm), 0.6]],
                s=20, zorder=4)

# Wave crest direction line (E-W)
ax2.plot([-65, 65], [0, 0], color=C_DECK, lw=1.5, ls="--", zorder=2)
ax2.text(67, 0, "Wave\ncrest", ha="left", fontsize=7, color=C_DECK, va="center")

# Wave propagation arrows
ax2.annotate("", xy=(0, 65), xytext=(0, 45),
             arrowprops=dict(arrowstyle="-|>", color=C_DECK, lw=1.2))
ax2.text(3, 72, "Wave\ndirection", ha="left", fontsize=6.5, color=C_DECK)

# 3 nozzle holes at 120°, r = 0.5*Rc
noz_angles = [90, 90+120, 90+240]   # degrees from crest direction
r_noz = 0.5 * Rc_mm
for ang in noz_angles:
    nx = r_noz * math.cos(math.radians(ang))
    ny = r_noz * math.sin(math.radians(ang))
    # Elliptic nozzle (elongated in wave direction = horizontal)
    ax2.add_patch(mpatches.Ellipse((nx, ny), 24, 16,
                  angle=0, fc=C_NOZZLE, ec="k", lw=0.8, zorder=6, alpha=0.9))
ax2.text(0, 28, "Nozzle 1", ha="center", fontsize=6, color=C_NOZZLE)
ax2.text(-27, -14, "Nozzle 2", ha="center", fontsize=6, color=C_NOZZLE)
ax2.text(27, -14, "Nozzle 3", ha="center", fontsize=6, color=C_NOZZLE)

# Drainage slots (at wave-face direction, phi=+/-90° → at y= ±Rc)
# In wave propagation direction (N-S in plan)
for sy in [Rc_mm, -Rc_mm]:
    ax2.add_patch(mpatches.Rectangle((-slot_w_mm/2, sy-sl_h*100/2*1.5),
                  slot_w_mm, sl_h*100*1.5,
                  fc=C_SLOT, ec="k", lw=0.8, zorder=7, alpha=0.9))
ax2.annotate("Drainage\nslot", xy=(0, Rc_mm),
             xytext=(22, Rc_mm+8), fontsize=6.5, color=C_SLOT, ha="left",
             arrowprops=dict(arrowstyle="->", color=C_SLOT, lw=0.8))
ax2.annotate("Drainage\nslot", xy=(0, -Rc_mm),
             xytext=(22, -Rc_mm-8), fontsize=6.5, color=C_SLOT, ha="left",
             arrowprops=dict(arrowstyle="->", color=C_SLOT, lw=0.8))

# h_cut contour labels
ax2.text(-Rc_mm*0.7, Rc_mm*0.7, f"h_cut={h_cut_max_mm:.0f}mm\n(at ±90°)",
         fontsize=6.5, color=C_DECK, ha="center")
ax2.text(0, -Rc_mm*0.55, "h_cut=0\n(at crest)", fontsize=6.5,
         color=C_DECK, ha="center", style="italic")

# R_c dimension
ax2.annotate("", xy=(Rc_mm, 0), xytext=(0, 0),
             arrowprops=dict(arrowstyle="<->", color=C_TXT, lw=1))
ax2.text(Rc_mm/2, 5, f"Rc={Rc_mm}mm", ha="center", fontsize=6.5, color=C_TXT)

ax2.set_xlim(-80, 90); ax2.set_ylim(-70, 90)
ax2.set_title("(b) Plan view of cylinder base at wave crest\n"
              "(darker = lower base elevation, slots at ±90°, nozzles at r=0.5Rc)",
              fontsize=8.5, fontweight="bold", pad=4)

# ══════════════════════════════════════════════════════════════════════════════
# Panel (c) — Performance table: 4 configurations
# ══════════════════════════════════════════════════════════════════════════════
ax3 = fig.add_subplot(gs[0, 1])
ax3.set_facecolor("#eef5fb")
ax3.axis("off")

# Quantitative analysis (C1 S2 representative conditions)
# Base CIST (3-nozzle, curved vane)
E_base    = 95.0     # %  capped
dp_base   = 0.98     # mmHg  (C1 S2)
td_base   = 3.3      # turndown ratio

# With drainage slots: no mass-transfer change, ΔP stabilised at high L/G
# (prevents base liquid pool blocking vapor inlet at high-load cases)
E_slots   = 95.0
dp_slots  = 0.95     # marginal: keeps inlet area clear, ~3% ΔP reduction
td_slots  = 3.5      # small TD extension (lower minimum ΔP threshold)

# + wall grooves: inner V-grooves 1mm deep, 8mm pitch
# Area factor 1 + 2d_g/s_g = 1.25; turbulence Colburn (1.15)^(2/3)=1.10
# Wall-film contribution: ~5% of total NOG × 1.25 × 1.10 → +3.4% NOG
# But we're capped at 95%; benefit shows in lower-E_OC services
E_grooves  = 95.0    # still capped
dp_grooves = 0.99    # +1% ΔP_wall (negligible fraction of K)
td_grooves = 3.5

# + crown chevron: additional mist elimination at exit
# Captures d=5-8 µm drops that evade centrifuge (above d50=8µm cutoff)
# Separation efficiency: +10-15% for 5-8 µm band
# ΔP penalty: 0.05 mmHg (chevron form drag at low u_θ,crown)
E_chev  = 95.0       # efficiency unchanged (mass transfer, not separation)
dp_chev = 1.03       # +0.05 mmHg chevron
td_chev = 3.5

configs = [
    ("Baseline CIST\n(3-noz, curved vane)", E_base,   dp_base,   td_base,  "darkgreen"),
    ("+ Drainage\nslots",                   E_slots,   dp_slots,  td_slots, "#2e7d32"),
    ("+ Inner wall\nspiral grooves",         E_grooves, dp_grooves,td_grooves,"#388e3c"),
    ("+ Crown\nchevron",                    E_chev,    dp_chev,   td_chev,  "#43a047"),
]

# Feature benefit table
table_rows = [
    ["Feature",            "Physics",                          "Effect",           "ΔP penalty"],
    ["Drainage slots",     "Gravity drain of centrifugal film","Stable base;",     "None"],
    ["",                   "through wall at lowest base pt.",   "prevents pooling", ""],
    ["Spiral grooves\n(inner)", "Area +25%; turb. factor 1.10;","Wall-film NOG",   "+1% of"],
    ["",                    "Re_film≈4100 → turbulent",          "+3-5% (uncapped)","ΔP_wall"],
    ["Crown chevron",       "Impingement capture at exit",       "Capture d<8µm;",  "+0.05"],
    ["",                    "(opposing angled vanes)",           "+10-15% sep. eff","mmHg"],
    ["Outer grooves",       "Not in vapor path",                "Negligible",       "—"],
    ["Criss-cross\nmid-baffle","Disrupts Rankine vortex",        "Defer — reduces", "+15-25%"],
    ["",                   "(if full-height)",                   "centrifuge g",     "ΔP_total"],
]

y_start = 0.97
col_x   = [0.01, 0.25, 0.58, 0.83]
col_w   = [20,   20,   18,   10]   # approximate

ax3.text(0.5, 1.02, "(c)  Feature assessment: physics, effect, ΔP penalty",
         ha="center", fontsize=9, fontweight="bold", transform=ax3.transAxes)

header_cols = ["Feature", "Physics / mechanism", "Mass-transfer effect", "ΔP penalty"]
header_colors = ["#1565c0","#1565c0","#1565c0","#1565c0"]
for cx, txt, fc in zip(col_x, header_cols, header_colors):
    ax3.text(cx, 0.96, txt, ha="left", fontsize=7.5, fontweight="bold",
             color="white", transform=ax3.transAxes,
             bbox=dict(boxstyle="round,pad=0.2", fc=fc, ec="none"))

rows_data = [
    ("Drainage slots",
     "Centrifugal film drains via\nwall slots at lowest base points\n(ϕ=±90° from crest)",
     "Stable ΔP at high L/G;\nprevents vapor-inlet blockage;\nslight TD extension",
     "None\n(mechanical)"),
    ("Inner-wall\nspiral grooves",
     "V-grooves 1×2mm @ 8mm pitch;\narea ×1.25; Colburn turb. ×1.10;\nfilm Re ≈ 4100 (turbulent)",
     "Wall-film NOG boost ×1.38;\n+3–5% efficiency in uncapped\nservices; negligible below cap",
     "+1% of\nΔP_wall only\n(< 0.01 mmHg)"),
    ("Cross-chevron\ncrown vanes",
     "Two opposing angled vanes\nat exit; impingement capture\nfor d < 8 µm drops",
     "+10–15% liquid capture\nfor 5–8 µm fraction;\nmist purity improved",
     "+0.05 mmHg\n(chevron form\ndrag)"),
    ("Outer-wall\ngrooves",
     "Not in vapor path;\nliquid sump side only",
     "Negligible for\nmass transfer",
     "None"),
    ("Mid-height criss-\ncross baffle\n(full height)",
     "Opposing helices disrupt\nRankine vortex throughout\nchamber",
     "DEFER: reduces Ng_eff,\nimpairs centrifugal\nseparation",
     "+15–25%\ntotal ΔP"),
]

row_colors = ["#e8f5e9","#f1f8e9","#e8f5e9","#fffde7","#ffebee"]
y_pos = 0.87
for row, rc_col in zip(rows_data, row_colors):
    ax3.add_patch(mpatches.FancyBboxPatch((0.0, y_pos-0.07), 1.0, 0.09,
                  boxstyle="round,pad=0.01", fc=rc_col, ec="none",
                  transform=ax3.transAxes, zorder=1))
    for cxi, txt in zip(col_x, row):
        color = "#b71c1c" if "DEFER" in txt or "impairs" in txt else C_TXT
        ax3.text(cxi, y_pos, txt, ha="left", va="top", fontsize=6.8,
                 color=color, transform=ax3.transAxes, linespacing=1.3)
    y_pos -= 0.105

ax3.set_xlim(0, 1); ax3.set_ylim(0, 1.05)

# ══════════════════════════════════════════════════════════════════════════════
# Panel (d) — Key derivations (film thickness, groove effect, slot sizing)
# ══════════════════════════════════════════════════════════════════════════════
ax4 = fig.add_subplot(gs[1, 1])
ax4.set_facecolor("#eef5fb")
ax4.axis("off")

ax4.text(0.5, 1.02, "(d)  Derived quantities (C1 S2 representative conditions)",
         ha="center", fontsize=9, fontweight="bold", transform=ax4.transAxes)

# Film thickness calculation
rho_L  = 37.80 * 16.0185    # kg/m³
mu_L   = 0.135e-3            # Pa·s
rho_V  = 1.044 * 16.0185
Ng_eff = 42.4
a_c    = Ng_eff * 9.81       # m/s²
Rc_m   = Rc_mm * 1e-3
# Q_L_ch: liquid per chamber at N_ch=69
Q_L_total = (21140 / 2.20462 / 3600) / rho_L  # m³/s
Q_L_ch = Q_L_total / 69
alpha_dep = 0.85    # fraction reaching wall
Q_wall = alpha_dep * Q_L_ch
Gamma = Q_wall / (2 * math.pi * Rc_m)          # film flow per unit circumference
delta_m = (3 * mu_L * Gamma / (rho_L * a_c))**(1/3)
Re_film = 4 * Gamma * rho_L / mu_L

# Groove enhancement
area_fac  = 1 + 2 * d_g / s_g
turb_fac  = 1.10
groove_factor = area_fac * turb_fac
wall_nug_fraction = 0.07   # wall film contributes ~7% of total NOG (estimated)
groove_nug_gain   = wall_nug_fraction * (groove_factor - 1)

# Slot sizing
u_x_design = 3.0
u_th_base = u_x_design / math.tan(math.radians(20))   # u_theta at inlet
rho_V_si  = rho_V
tau_i     = 0.5 * 0.02 * rho_V_si * u_th_base**2    # interfacial shear
k_L_wall  = math.sqrt(2e-9 * tau_i / mu_L)           # surface renewal
a_wall    = 2 / Rc_m                                  # wall area/volume

# Drainage slot sizing
h_film_head = delta_m * a_c   # equivalent driving head (m/s²) × thickness
u_slot_min  = math.sqrt(2 * a_c * delta_m)   # slot drain velocity
A_slot_need = Q_wall / u_slot_min
n_slots  = 2
A_slot_each = A_slot_need / n_slots
# Round to practical dimensions
slot_h_calc = slot_h_mm * 1e-3
slot_w_calc = A_slot_each / slot_h_calc
slot_w_mm_c = slot_w_calc * 1e3

# Crown chevron
u_th_crown = u_x_design / math.tan(math.radians(70))   # u_theta at crown
dp_chev_pa = 0.5 * rho_V_si * u_th_crown**2 * 0.3     # 30% loss coeff for chevron
dp_chev_mm = dp_chev_pa / 133.322

derivations = [
    ("Liquid film thickness (Nusselt centrifugal):",
     rf"  δ = (3 μ_L Γ / (ρ_L a_c))^(1/3)  where Γ = α_dep Q_L_ch/(2πRc)",
     rf"  = (3×{mu_L:.2e}×{Gamma:.2e} / ({rho_L:.0f}×{a_c:.0f}))^(1/3)  = {delta_m*1e6:.1f} µm"),
    ("Film Reynolds number:",
     rf"  Re_film = 4Γρ_L/μ_L = 4×{Gamma:.2e}×{rho_L:.0f}/{mu_L:.2e}",
     rf"  = {Re_film:.0f}  → TURBULENT film (Re > 1600)"),
    ("Inner groove enhancement factor:",
     rf"  f_area = 1 + 2d_g/s_g = 1 + 2×{d_g}/{s_g} = {area_fac:.2f}  (+{(area_fac-1)*100:.0f}% wall area)",
     rf"  f_turb = (f_rough/f_smooth)^(2/3) ≈ 1.10  (Colburn analogy)"),
    ("",
     rf"  Combined: {area_fac:.2f} × {turb_fac:.2f} = {groove_factor:.2f}×  on wall-film NOG",
     rf"  Wall-film NOG fraction ≈ {wall_nug_fraction*100:.0f}%  → groove adds +{groove_nug_gain*100:.1f}% to total NOG"),
    ("Drainage slot minimum area (2 slots, film gravity drain):",
     rf"  u_slot = √(2 a_c δ) = √(2×{a_c:.0f}×{delta_m:.2e}) = {u_slot_min:.3f} m/s",
     rf"  A_slot_each = Q_wall/(N_slots u_slot) = {A_slot_each*1e6:.0f} mm²  → {slot_h_mm}mm×{slot_w_mm_c:.0f}mm per slot"),
    ("Crown chevron ΔP (C_loss=0.30, u_θ,crown):",
     rf"  u_θ,crown = u_x cot(70°) = {u_th_crown:.2f} m/s",
     rf"  ΔP_chev = 0.30 × ½ρ_V u²  = {dp_chev_pa:.1f} Pa = {dp_chev_mm:.3f} mmHg/tray"),
]

y_d = 0.94
for d in derivations:
    if d[0]:
        ax4.text(0.01, y_d, d[0], ha="left", fontsize=7.5, fontweight="bold",
                 color="#1a237e", transform=ax4.transAxes)
        y_d -= 0.06
    for line in d[1:]:
        ax4.text(0.03, y_d, line, ha="left", fontsize=7.2,
                 color=C_TXT, transform=ax4.transAxes, family="monospace")
        y_d -= 0.055
    y_d -= 0.01

# Summary box
summary = (
    f"SUMMARY (C1 S2, N_ch=69, u_x=3.0 m/s, Ng_eff=42.4g):\n"
    f"  Film: δ={delta_m*1e6:.0f}µm, turbulent (Re={Re_film:.0f})\n"
    f"  Groove: NOG boost ×{groove_factor:.2f} on wall fraction → +{groove_nug_gain*100:.1f}% total\n"
    f"  Slots: {slot_h_mm}mm×{slot_w_mm_c:.0f}mm each (2 off), drain at {u_slot_min:.2f}m/s\n"
    f"  Crown chevron: ΔP={dp_chev_mm:.3f}mmHg, capture +10–15% (d<8µm)\n"
    f"  Criss-cross mid-baffle: DEFER (disrupts vortex)"
)
ax4.text(0.01, 0.13, summary, ha="left", va="bottom", fontsize=7.5,
         color=C_TXT, transform=ax4.transAxes,
         bbox=dict(boxstyle="round,pad=0.4", fc="#e8f5e9", ec="#2e7d32", lw=1.5))

ax4.set_xlim(0, 1); ax4.set_ylim(0, 1.05)

# ── Figure title ─────────────────────────────────────────────────────────────
fig.suptitle(
    "CIST — Inner-Wall Film Enhancement, Drainage Slots and Cross-Chevron Crown  "
    "(Figure 49)\n"
    f"Rc={Rc_mm}mm · Hc={Hc_mm}mm · α_deck={math.degrees(alpha_deck):.1f}° · "
    f"h_cut,max={h_cut_max_mm:.1f}mm · 3 nozzles@120° preserved  |  "
    f"Groove: {d_g}mm×{w_g}mm @ {s_g}mm pitch · Slots: 2 off @ ±90° from crest",
    fontsize=9, y=0.975)

import os
outdir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
os.makedirs(outdir, exist_ok=True)
outpath = os.path.join(outdir, "49_cist_wall_film.png")
plt.savefig(outpath, dpi=150, bbox_inches="tight")
plt.close(fig)
print(f"Saved → {outpath}")
