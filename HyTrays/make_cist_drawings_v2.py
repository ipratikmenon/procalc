#!/usr/bin/env python3
"""Generate enhanced CIST drawings v2.

Outputs
-------
output/43_cist_multiview_3d.png  — six-view 3-D assembly (plan, 3 elevations, 2 isometrics)
output/44_cist_wavy_deck.png     — wavy-deck variant: cross-section, 3D, nozzle detail, table
"""
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec
from mpl_toolkits.mplot3d import Axes3D            # noqa: F401 – registers projection
from mpl_toolkits.mplot3d.art3d import Poly3DCollection

# ── shared palette ───────────────────────────────────────────────────────────
C_CHAMBER = "#4a90d9"    # steel-blue   – chamber wall
C_CROWN   = "#1a5276"    # dark blue    – crown vane disc
C_DECK    = "#b0bec5"    # grey         – tray deck
C_LIQUID  = "#81d4fa"    # pale blue    – liquid
C_JET     = "#0277bd"    # mid-blue     – liquid jet
C_VAPOR   = "#e53935"    # red          – vapour
C_SLOT    = "#f57c00"    # orange       – tangential inlet slot
C_WAVE    = "#607d8b"    # blue-grey    – wavy deck material
BG        = "#f5f7fa"

# ── representative geometry (dimensionless; Rc = 1 ≡ 75 mm) ─────────────────
Rc = 1.00   # unit chamber radius
Hc = 2.80   # chamber height / Rc

# ════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ════════════════════════════════════════════════════════════════════════════

def cyl_mesh(r, z0, z1, n=80):
    """Cylindrical wall as surface arrays (X, Y, Z)."""
    th = np.linspace(0, 2*np.pi, n)
    Z  = np.array([z0, z1])
    T, Zg = np.meshgrid(th, Z)
    return r*np.cos(T), r*np.sin(T), Zg


def disk_mesh(r_in, r_out, z, n=80):
    """Annular disk as surface arrays."""
    th = np.linspace(0, 2*np.pi, n)
    R  = np.array([r_in, r_out])
    T, Rg = np.meshgrid(th, R)
    return Rg*np.cos(T), Rg*np.sin(T), np.full_like(T, z)


def arc_slot(r, th0, th1, z0, z1, n=24):
    """Tangential inlet slot represented as a short arc-segment on the cylinder wall."""
    th  = np.linspace(th0, th1, n)
    Z   = np.array([z0, z1])
    T, Zg = np.meshgrid(th, Z)
    return (r+0.04)*np.cos(T), (r+0.04)*np.sin(T), Zg


# ════════════════════════════════════════════════════════════════════════════
#  DRAW ONE CIST CHAMBER in a 3-D axes
# ════════════════════════════════════════════════════════════════════════════

def draw_chamber_3d(ax, a=0.28):
    """Populate ax with CIST chamber surfaces; a = translucency."""
    # outer wall
    ax.plot_surface(*cyl_mesh(Rc, 0, Hc), color=C_CHAMBER, alpha=a, linewidth=0, zorder=2)
    # inner wall (thin shell)
    ax.plot_surface(*cyl_mesh(Rc*0.92, 0.05, Hc-0.05), color="white", alpha=a*0.5,
                    linewidth=0, zorder=1)
    # crown disc
    ax.plot_surface(*disk_mesh(0, Rc, Hc), color=C_CROWN, alpha=a+0.20, linewidth=0, zorder=3)
    # top ring above crown
    ax.plot_surface(*cyl_mesh(Rc, Hc, Hc+0.12), color=C_CROWN, alpha=a+0.20, linewidth=0)
    # tray deck
    ax.plot_surface(*disk_mesh(0, 1.55*Rc, -0.18), color=C_DECK, alpha=a+0.12, linewidth=0)
    # sump cylinder below deck
    ax.plot_surface(*cyl_mesh(0.35*Rc, -0.65, 0), color=C_DECK, alpha=a+0.05, linewidth=0)
    # tangential inlet slot (orange arc patch at bottom of wall)
    ax.plot_surface(*arc_slot(Rc, -np.pi/6, np.pi/6, 0.10, 0.55),
                    color=C_SLOT, alpha=0.85, linewidth=0, zorder=4)
    # liquid jet (upward arrow)
    ax.quiver(0, 0, -0.55, 0, 0, 0.60, color=C_JET,
              arrow_length_ratio=0.28, linewidth=2.2, zorder=5)
    # vapour exit (dashed arrow above crown)
    ax.quiver(0, 0, Hc+0.12, 0, 0, 0.65, color=C_VAPOR,
              arrow_length_ratio=0.28, linewidth=2.0, linestyle="--", zorder=5)
    # collection ring drain line
    ax.plot([Rc*0.85, Rc*0.85], [0, 0], [Hc-0.20, -0.18],
            color=C_LIQUID, linewidth=2.2, linestyle="-.")

    lim = 1.75
    ax.set_xlim(-lim, lim)
    ax.set_ylim(-lim, lim)
    ax.set_zlim(-0.80, Hc + 0.90)
    ax.set_axis_off()
    try:
        ax.set_box_aspect([1, 1, Hc/Rc * 0.90])
    except AttributeError:
        pass


# ════════════════════════════════════════════════════════════════════════════
#  FIGURE 43 — SIX-VIEW ASSEMBLY DRAWING
# ════════════════════════════════════════════════════════════════════════════

VIEWS = [
    # (elev, azim, title_line1, title_line2)
    (90,  -90, "PLAN VIEW",          "(looking down — top)"),
    (0,   -90, "FRONT ELEVATION",    "(North face)"),
    (0,     0, "SIDE ELEVATION",     "(East face)"),
    (0,   180, "REAR ELEVATION",     "(South face)"),
    (30,   45, "ISOMETRIC — NE",     ""),
    (30,  225, "ISOMETRIC — SW",     ""),
]

fig43 = plt.figure(figsize=(16, 11))
fig43.patch.set_facecolor(BG)

gs43 = gridspec.GridSpec(2, 3, figure=fig43, hspace=0.08, wspace=0.04,
                          left=0.02, right=0.98, top=0.90, bottom=0.08)

for idx, (elev, azim, t1, t2) in enumerate(VIEWS):
    ax = fig43.add_subplot(gs43[idx // 3, idx % 3], projection="3d")
    draw_chamber_3d(ax, a=0.30)
    ax.view_init(elev=elev, azim=azim)
    ax.set_title(f"{t1}\n{t2}", fontsize=8.5, fontweight="bold",
                 color="#1a252f", pad=3)

fig43.suptitle(
    "Cascade Impingement Swirl Tray (CIST) — Six-View Assembly Drawing\n"
    "Plan  ·  Front / Rear / Side Elevations  ·  NE & SW Isometrics",
    fontsize=13, fontweight="bold", color="#1a252f", y=0.975
)

legend_items = [
    mpatches.Patch(facecolor=C_CHAMBER, alpha=0.55, label="Swirl chamber wall (stainless steel)"),
    mpatches.Patch(facecolor=C_CROWN,   alpha=0.80, label="Crown vane separator disc"),
    mpatches.Patch(facecolor=C_DECK,    alpha=0.65, label="Tray deck / sump shell"),
    mpatches.Patch(facecolor=C_SLOT,    alpha=0.90, label="Tangential vapour inlet slot"),
    mpatches.Patch(facecolor=C_JET,     alpha=0.90, label="Upward liquid jet"),
    mpatches.Patch(facecolor=C_VAPOR,   alpha=0.90, label="Treated vapour exit"),
    mpatches.Patch(facecolor=C_LIQUID,  alpha=0.90, label="Collection-ring drain conduit"),
]
fig43.legend(handles=legend_items, loc="lower center", ncol=4, fontsize=8,
             frameon=True, fancybox=True, framealpha=0.9,
             bbox_to_anchor=(0.50, 0.00))

fig43.savefig("output/43_cist_multiview_3d.png", dpi=150, bbox_inches="tight",
              facecolor=BG)
plt.close(fig43)
print("Saved  output/43_cist_multiview_3d.png")


# ════════════════════════════════════════════════════════════════════════════
#  FIGURE 44 — WAVY DECK VARIANT
# ════════════════════════════════════════════════════════════════════════════

ALPHA_DEG = 27.5
alpha_rad = np.radians(ALPHA_DEG)
p_w       = 120.0          # wave pitch crest-to-crest, mm
h_w       = (p_w/2) * np.tan(alpha_rad)   # ≈ 31.2 mm

fig44 = plt.figure(figsize=(17, 10))
fig44.patch.set_facecolor(BG)
gs44 = gridspec.GridSpec(2, 3, figure=fig44, hspace=0.42, wspace=0.32,
                          left=0.06, right=0.97, top=0.90, bottom=0.06)

# ── Panel A: 2-D wave cross-section (left column, full height) ──────────────
ax_A = fig44.add_subplot(gs44[:, 0])

N_WAVES = 5
x_wave, y_wave = [], []
for i in range(N_WAVES):
    x0 = i * p_w
    x_wave += [x0, x0 + p_w/2, x0 + p_w]
    y_wave += [0.0, h_w, 0.0]
xw = np.array(x_wave)
yw = np.array(y_wave)

ax_A.fill_between(xw, yw, y2=-20, color=C_WAVE, alpha=0.28, label="Deck material")
ax_A.plot(xw, yw, color="#37474f", linewidth=2.8, label="Wave profile")

# Nozzles at troughs
for i in range(N_WAVES):
    xn = i * p_w
    ax_A.plot(xn, 0, "o", color=C_JET, markersize=10, zorder=6)
    ax_A.annotate("", xy=(xn, h_w*1.55), xytext=(xn, 2),
                  arrowprops=dict(arrowstyle="-|>", color=C_JET, lw=2.0))

# Chambers on crests
r_box = 18
for i in range(N_WAVES):
    xc = (i + 0.5) * p_w
    rect = mpatches.FancyBboxPatch(
        (xc - r_box, h_w), 2*r_box, 58,
        boxstyle="round,pad=1.5",
        facecolor=C_CHAMBER, edgecolor=C_CHAMBER, alpha=0.38, linewidth=1.5
    )
    ax_A.add_patch(rect)
    if i == 1:
        ax_A.text(xc, h_w + 63, "Swirl\nchamber", ha="center",
                  fontsize=7.5, color="#1a5f8a", fontweight="bold")

# Wave angle annotation (on second face, left to right)
x_s, y_s = p_w, 0.0
x_e, y_e = p_w + p_w/2, h_w
arc_r = 25
th_arc = np.linspace(0, alpha_rad, 50)
ax_A.plot(x_s + arc_r*np.cos(th_arc), y_s + arc_r*np.sin(th_arc),
          color="#c62828", linewidth=1.8)
ax_A.text(x_s + arc_r*1.25*np.cos(alpha_rad/2),
          y_s + arc_r*1.25*np.sin(alpha_rad/2),
          f"α = {ALPHA_DEG:.1f}°", fontsize=9.5, color="#c62828", fontweight="bold")

# Pitch dimension arrow
y_dim = -12
ax_A.annotate("", xy=(p_w, y_dim), xytext=(0, y_dim),
              arrowprops=dict(arrowstyle="<->", color="#555", lw=1.5))
ax_A.text(p_w/2, y_dim - 6, f"$p_w$ = {p_w:.0f} mm", ha="center", fontsize=8.5)

# Wave-height dimension
ax_A.annotate("", xy=(p_w/2 + 4, h_w), xytext=(p_w/2 + 4, 0),
              arrowprops=dict(arrowstyle="<->", color="#555", lw=1.5))
ax_A.text(p_w/2 + 12, h_w/2, f"$h_w$ = {h_w:.1f} mm",
          fontsize=8.5, va="center", color="#555")

ax_A.text(0, -7, "Trough\n(nozzle)", ha="center", fontsize=7.5, color=C_JET)
ax_A.text(p_w/2, h_w+1, "Crest\n(chamber\nseat)", ha="center", va="bottom",
          fontsize=7.5, color="#37474f")

ax_A.set_xlim(-20, N_WAVES*p_w + 15)
ax_A.set_ylim(-25, h_w + 88)
ax_A.set_aspect("equal")
ax_A.set_xlabel("Horizontal distance (mm)", fontsize=9)
ax_A.set_ylabel("Height (mm)", fontsize=9)
ax_A.set_title(
    "(a)  Wave profile — cross-section\n"
    r"$\alpha = 27.5°$,  $p_w = 120$ mm,  $h_w = (p_w/2)\tan\alpha$",
    fontsize=9, fontweight="bold"
)
ax_A.grid(True, alpha=0.22, linestyle="--")
ax_A.legend(fontsize=8, loc="upper right")


# ── Panel B: 3-D wavy deck perspective (top-right 2 columns) ────────────────
ax_B = fig44.add_subplot(gs44[0, 1:], projection="3d")

Ny   = 5          # crests in y
Lx   = 220.0      # flow-direction length, mm
pw   = p_w        # 120 mm
n_mx, n_my = 100, 160
xd = np.linspace(0, Lx, n_mx)
yd = np.linspace(0, Ny*pw, n_my)
Xd, Yd = np.meshgrid(xd, yd)
# triangular wave (0 at trough, h_w at crest)
Zd = h_w * (1.0 - np.abs(2.0*(Yd % pw)/pw - 1.0))

ax_B.plot_surface(Xd, Yd, Zd, cmap="Blues", alpha=0.38, linewidth=0,
                  vmin=-10, vmax=h_w*2.2)

# Chambers at crests
r_ch, h_ch = 17.0, 62.0
n_th = 48
th_c = np.linspace(0, 2*np.pi, n_th)
xc_list = [Lx*0.32, Lx*0.68]
yc_list = [(i + 0.5)*pw for i in range(Ny)]
for xcc in xc_list:
    for ycc in yc_list:
        zc = h_w
        # cylinder wall
        Zwall = np.array([zc, zc + h_ch])
        Tc, Zcyl = np.meshgrid(th_c, Zwall)
        ax_B.plot_surface(xcc + r_ch*np.cos(Tc), ycc + r_ch*np.sin(Tc), Zcyl,
                          color=C_CHAMBER, alpha=0.42, linewidth=0)
        # crown disc
        Rd = np.linspace(0, r_ch, 12)
        Tdisc, Rdg = np.meshgrid(th_c, Rd)
        ax_B.plot_surface(xcc + Rdg*np.cos(Tdisc), ycc + Rdg*np.sin(Tdisc),
                          np.full_like(Rdg, zc + h_ch),
                          color=C_CROWN, alpha=0.60, linewidth=0)

# Liquid jets from troughs
yt_arr = [i*pw for i in range(Ny+1)]
for yt in yt_arr[1:-1]:
    for xj in xc_list:
        ax_B.quiver(xj, yt, 0.0, 0, 0, h_w*0.88,
                    color=C_JET, arrow_length_ratio=0.30, linewidth=1.8)

# Vapor exits
for xcc in xc_list:
    for ycc in yc_list:
        ax_B.quiver(xcc, ycc, h_w + h_ch, 0, 0, h_ch*0.40,
                    color=C_VAPOR, arrow_length_ratio=0.35, linewidth=1.4,
                    linestyle="--")

ax_B.set_xlabel("Flow direction (mm)", fontsize=8, labelpad=1)
ax_B.set_ylabel("Width (mm)", fontsize=8, labelpad=1)
ax_B.set_zlabel("Height (mm)", fontsize=8, labelpad=1)
ax_B.set_title("(b)  3-D wavy deck — hex chambers at crests, jets at troughs",
               fontsize=9, fontweight="bold")
ax_B.view_init(elev=26, azim=28)
try:
    ax_B.set_box_aspect([Lx, Ny*pw, (h_ch + h_w)*1.5])
except AttributeError:
    pass


# ── Panel C: Single-wave nozzle detail ──────────────────────────────────────
ax_C = fig44.add_subplot(gs44[1, 1])

xw1 = np.array([0.0, p_w/2, p_w])
yw1 = np.array([0.0, h_w,   0.0])
ax_C.fill_between(xw1, yw1, -18, color=C_WAVE, alpha=0.28)
ax_C.plot(xw1, yw1, color="#37474f", lw=2.5)

# Liquid pool in trough
ax_C.fill_between([-14, 14], [-18, -18], [0, 0], color=C_LIQUID, alpha=0.50)
ax_C.text(0, -9, "Liquid sump", ha="center", fontsize=7.5, color="#01579b",
          fontweight="bold")

# Nozzle
ax_C.plot(0, 0, "o", color=C_JET, markersize=11, zorder=6)
ax_C.annotate("", xy=(0, h_w*1.5), xytext=(0, 2.0),
              arrowprops=dict(arrowstyle="-|>", color=C_JET, lw=2.2))
ax_C.text(4, h_w*0.7, "Liquid\njet", fontsize=7.5, color=C_JET, fontweight="bold")

# Liquid film on left wave face
for offset in [1.5, 3.5, 5.5]:
    ax_C.plot([0 + offset*np.cos(alpha_rad + np.pi/2),
               p_w/2 + offset*np.cos(alpha_rad + np.pi/2)],
              [0 + offset*np.sin(alpha_rad + np.pi/2),
               h_w + offset*np.sin(alpha_rad + np.pi/2)],
              color=C_LIQUID, lw=1.2, alpha=0.75)
ax_C.text(-12, h_w*0.45, "Liquid\nfilm", fontsize=7.5, color="#0277bd",
          ha="right", rotation=0)

# Vapour path arrow (along right wave face towards chamber)
ax_C.annotate("", xy=(p_w/2 - 18, h_w + 5),
              xytext=(p_w - 4, 5),
              arrowprops=dict(arrowstyle="-|>", color=C_VAPOR, lw=1.8,
                              connectionstyle="arc3,rad=0.20"))
ax_C.text(p_w*0.82, h_w*0.50, "Vapour\npath", fontsize=7.5, color=C_VAPOR,
          ha="center", rotation=-50)

# Chamber rectangle
ch_rect = mpatches.FancyBboxPatch(
    (p_w/2 - 18, h_w), 36, 58,
    boxstyle="round,pad=2",
    facecolor=C_CHAMBER, edgecolor=C_CHAMBER, alpha=0.38, linewidth=1.5
)
ax_C.add_patch(ch_rect)
ax_C.text(p_w/2, h_w + 63, "Swirl chamber", ha="center", fontsize=7.5, color="#1a5f8a")

# Angle arc annotation
th_a = np.linspace(0, alpha_rad, 30)
ax_C.plot(p_w + 22*np.cos(th_a), 22*np.sin(th_a), color="#c62828", lw=1.8)
ax_C.text(p_w + 26, 10, f"α = {ALPHA_DEG:.1f}°", fontsize=8.5, color="#c62828",
          fontweight="bold")

ax_C.set_xlim(-22, p_w + 45)
ax_C.set_ylim(-22, h_w + 82)
ax_C.set_aspect("equal")
ax_C.set_xlabel("Lateral position (mm)", fontsize=8)
ax_C.set_ylabel("Height (mm)", fontsize=8)
ax_C.set_title("(c)  Nozzle-at-trough detail\n(jet · liquid film · vapour path)",
               fontsize=9, fontweight="bold")
ax_C.grid(True, alpha=0.20, linestyle="--")


# ── Panel D: Flat vs wavy deck comparison table ──────────────────────────────
ax_D = fig44.add_subplot(gs44[1, 2])
ax_D.axis("off")

rows = [
    ["Deck surface area",       "1.000",          r"+12.8 % (1/cosα)"],
    ["Mass-transfer area",      "Baseline",        "+12–18 %"],
    ["Bending rigidity",        "1 ×",             "~500 × (same t)"],
    ["Deck plate thickness",    r"$t_0$",          r"0.55–0.65 $t_0$"],
    ["Deck weight",             "Baseline",        "−35 to −45 %"],
    ["Pressure drop",           "Baseline",        "+5–8 %"],
    ["Turndown (inherent)",     "~2.8 : 1",       "~3.5 : 1"],
    ["Self-draining at idle",   "Moderate",       "Excellent (α ≫ 5°)"],
    ["Fouling tendency",        "Moderate",       "Lower"],
    ["Manufacture",             "Flat press/stamp","Roll-form"],
]
col_labels = ["Property", "Flat deck", r"Wavy deck (α=27.5°)"]

tbl = ax_D.table(
    cellText  = rows,
    colLabels = col_labels,
    cellLoc   = "center",
    loc       = "center",
)
tbl.auto_set_font_size(False)
tbl.set_fontsize(7.5)
tbl.scale(1.02, 1.52)

for j in range(3):
    cell = tbl[(0, j)]
    cell.set_facecolor("#1565c0")
    cell.set_text_props(color="white", fontweight="bold")

for i in range(1, len(rows) + 1):
    bg = "#e3f2fd" if i % 2 == 0 else "#ffffff"
    for j in range(3):
        tbl[(i, j)].set_facecolor(bg)
        if j == 2:
            tbl[(i, j)].set_text_props(color="#1b5e20", fontweight="bold")

tbl.auto_set_column_width([0, 1, 2])
ax_D.set_title("(d)  Flat vs. wavy deck — key metrics",
               fontsize=9, fontweight="bold", pad=6)


fig44.suptitle(
    r"CIST Wavy-Deck Variant  |  Wave face angle $\alpha$ = 25–30°"
    "\nTriangular corrugation: liquid at troughs · chambers at crests",
    fontsize=12, fontweight="bold", color="#1a252f"
)

fig44.savefig("output/44_cist_wavy_deck.png", dpi=150, bbox_inches="tight",
              facecolor=BG)
plt.close(fig44)
print("Saved  output/44_cist_wavy_deck.png")
print("Done.")
