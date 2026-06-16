#!/usr/bin/env python3
"""CIST drawings v3 — fixes and new detail figures.

Regenerates:
  output/43_cist_multiview_3d.png  — 6-view with elliptical nozzle now visible
  output/45_cist_design_details.png — 4-panel: nozzle cuts, multi-jet
                                       cross-section, wavy-deck crest/trough
                                       clarification, internal features
"""
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.patheffects as pe
import matplotlib.gridspec as gridspec
from matplotlib.patches import Ellipse, Circle, Rectangle, FancyArrowPatch, FancyBboxPatch
from mpl_toolkits.mplot3d import Axes3D            # noqa: F401
from mpl_toolkits.mplot3d.art3d import Poly3DCollection

# ── palette ──────────────────────────────────────────────────────────────────
C_CHAMBER = "#4a90d9"
C_CROWN   = "#1a5276"
C_DECK    = "#b0bec5"
C_LIQUID  = "#81d4fa"
C_JET     = "#0277bd"
C_VAPOR   = "#e53935"
C_SLOT    = "#f57c00"
C_WAVE    = "#607d8b"
C_ELLIPSE = "#00b0ff"   # bright cyan — elliptical nozzle
BG        = "#f5f7fa"

# ── chamber geometry (unit: Rc = 1) ──────────────────────────────────────────
Rc = 1.00
Hc = 2.80
a_ell, b_ell = 0.28, 0.14   # nozzle semi-axes

# ════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ════════════════════════════════════════════════════════════════════════════

def cyl_mesh(r, z0, z1, n=80):
    th = np.linspace(0, 2*np.pi, n)
    T, Zg = np.meshgrid(th, np.array([z0, z1]))
    return r*np.cos(T), r*np.sin(T), Zg


def disk_mesh(r_in, r_out, z, n=80):
    th = np.linspace(0, 2*np.pi, n)
    T, Rg = np.meshgrid(th, np.array([r_in, r_out]))
    return Rg*np.cos(T), Rg*np.sin(T), np.full_like(T, z)


def arc_slot(r, th0, th1, z0, z1, n=24):
    th = np.linspace(th0, th1, n)
    T, Zg = np.meshgrid(th, np.array([z0, z1]))
    return (r+0.04)*np.cos(T), (r+0.04)*np.sin(T), Zg


# ════════════════════════════════════════════════════════════════════════════
#  DRAW ONE CIST CHAMBER — now with visible elliptical nozzle at base
# ════════════════════════════════════════════════════════════════════════════

def draw_chamber_3d(ax, a=0.28):
    # outer wall
    ax.plot_surface(*cyl_mesh(Rc, 0, Hc), color=C_CHAMBER, alpha=a, linewidth=0)
    # inner wall (thin shell)
    ax.plot_surface(*cyl_mesh(Rc*0.92, 0.05, Hc-0.05), color="white", alpha=a*0.4, linewidth=0)
    # crown disc
    ax.plot_surface(*disk_mesh(0, Rc, Hc), color=C_CROWN, alpha=a+0.20, linewidth=0)
    # crown top ring
    ax.plot_surface(*cyl_mesh(Rc, Hc, Hc+0.10), color=C_CROWN, alpha=a+0.18, linewidth=0)
    # tray deck (floor)
    ax.plot_surface(*disk_mesh(0, 1.50*Rc, -0.16), color=C_DECK, alpha=a+0.10, linewidth=0)
    # sump cylinder below deck
    ax.plot_surface(*cyl_mesh(0.40*Rc, -0.60, 0), color=C_DECK, alpha=a+0.05, linewidth=0)
    # tangential inlet slot (orange arc on cylinder wall)
    ax.plot_surface(*arc_slot(Rc, -np.pi/7, np.pi/7, 0.08, 0.52),
                    color=C_SLOT, alpha=0.85, linewidth=0)

    # ── ELLIPTICAL NOZZLE at chamber floor ──────────────────────────────────
    n_e = 120
    th_e = np.linspace(0, 2*np.pi, n_e)
    X_ell = a_ell * np.cos(th_e)
    Y_ell = b_ell * np.sin(th_e)
    Z_ell = np.zeros(n_e)
    ax.plot(X_ell, Y_ell, Z_ell, color=C_ELLIPSE, linewidth=3.5, zorder=10)
    # fill of ellipse (visible in plan view)
    ax.plot_surface(
        np.outer(np.linspace(0, 1, 8), a_ell*np.cos(th_e[:80])),
        np.outer(np.linspace(0, 1, 8), b_ell*np.sin(th_e[:80])),
        np.zeros((8, 80)),
        color=C_ELLIPSE, alpha=0.55, linewidth=0
    )

    # liquid jet arrow (upward from sump)
    ax.quiver(0, 0, -0.52, 0, 0, 0.58, color=C_JET,
              arrow_length_ratio=0.28, linewidth=2.4, zorder=8)
    # vapour exit arrow above crown
    ax.quiver(0, 0, Hc+0.10, 0, 0, 0.60, color=C_VAPOR,
              arrow_length_ratio=0.28, linewidth=2.0, linestyle="--", zorder=8)
    # collection ring drain
    ax.plot([Rc*0.80]*2, [0]*2, [Hc-0.18, -0.16],
            color=C_LIQUID, linewidth=2.2, linestyle="-.")

    lim = 1.80
    ax.set_xlim(-lim, lim); ax.set_ylim(-lim, lim)
    ax.set_zlim(-0.75, Hc + 0.85)
    ax.set_axis_off()
    try:
        ax.set_box_aspect([1, 1, Hc*0.90])
    except AttributeError:
        pass


# ════════════════════════════════════════════════════════════════════════════
#  FIGURE 43 — 6-VIEW ASSEMBLY (regenerated with ellipse)
# ════════════════════════════════════════════════════════════════════════════

VIEWS = [
    (90,  -90, "PLAN VIEW",         "(looking down — top)"),
    (0,   -90, "FRONT ELEVATION",   "(North face)"),
    (0,     0, "SIDE ELEVATION",    "(East face)"),
    (0,   180, "REAR ELEVATION",    "(South face)"),
    (30,   45, "ISO — NE",          ""),
    (30,  225, "ISO — SW",          ""),
]

fig43 = plt.figure(figsize=(16, 11))
fig43.patch.set_facecolor(BG)
gs43 = gridspec.GridSpec(2, 3, figure=fig43, hspace=0.08, wspace=0.04,
                          left=0.02, right=0.98, top=0.90, bottom=0.09)

for idx, (elev, azim, t1, t2) in enumerate(VIEWS):
    ax = fig43.add_subplot(gs43[idx // 3, idx % 3], projection="3d")
    draw_chamber_3d(ax, a=0.30)
    ax.view_init(elev=elev, azim=azim)
    ax.set_title(f"{t1}\n{t2}", fontsize=8.5, fontweight="bold", color="#1a252f", pad=3)

fig43.suptitle(
    "CIST — Six-View Assembly Drawing (updated: elliptical nozzle at chamber base visible)\n"
    "Plan  ·  Front / Rear / Side Elevations  ·  NE & SW Isometrics",
    fontsize=12, fontweight="bold", color="#1a252f", y=0.975
)
legend_items = [
    mpatches.Patch(facecolor=C_CHAMBER, alpha=0.55, label="Swirl chamber wall"),
    mpatches.Patch(facecolor=C_CROWN,   alpha=0.80, label="Crown vane separator"),
    mpatches.Patch(facecolor=C_DECK,    alpha=0.65, label="Tray deck / sump"),
    mpatches.Patch(facecolor=C_SLOT,    alpha=0.90, label="Tangential vapour inlet slot"),
    mpatches.Patch(facecolor=C_ELLIPSE, alpha=0.90, label="Elliptical liquid-jet nozzle"),
    mpatches.Patch(facecolor=C_JET,     alpha=0.90, label="Liquid jet (upward)"),
    mpatches.Patch(facecolor=C_VAPOR,   alpha=0.90, label="Treated vapour exit"),
    mpatches.Patch(facecolor=C_LIQUID,  alpha=0.90, label="Collection-ring drain"),
]
fig43.legend(handles=legend_items, loc="lower center", ncol=4, fontsize=7.5,
             frameon=True, framealpha=0.9, bbox_to_anchor=(0.50, 0.00))

fig43.savefig("output/43_cist_multiview_3d.png", dpi=150, bbox_inches="tight", facecolor=BG)
plt.close(fig43)
print("Saved  output/43_cist_multiview_3d.png")


# ════════════════════════════════════════════════════════════════════════════
#  FIGURE 45 — DESIGN REFINEMENT DETAILS  (4-panel)
# ════════════════════════════════════════════════════════════════════════════

fig45 = plt.figure(figsize=(16, 12))
fig45.patch.set_facecolor(BG)
gs45 = gridspec.GridSpec(2, 2, figure=fig45, hspace=0.40, wspace=0.30,
                          left=0.06, right=0.97, top=0.91, bottom=0.05)

# ── Panel A: Nozzle-cut options (top view of cylinder base) ─────────────────
ax_A = fig45.add_subplot(gs45[0, 0])

# Three circles representing cylinder cross-sections (top view)
r_cy = 1.0
centers_x = [-3.0, 0.0, 3.0]
labels_opt = [
    "(i) Single centered\nellipse (current)",
    "(ii) Three ellipses\nat 120°, r = 0.5Rc",
    "(iii) Four circular\nholes at 90°, r = 0.5Rc",
]

for ci, cx in enumerate(centers_x):
    # Cylinder outline
    theta_c = np.linspace(0, 2*np.pi, 200)
    ax_A.plot(cx + r_cy*np.cos(theta_c), r_cy*np.sin(theta_c),
              color="#37474f", linewidth=2.0)
    ax_A.fill(cx + r_cy*np.cos(theta_c), r_cy*np.sin(theta_c),
              color="white", alpha=0.7, zorder=1)

    if ci == 0:
        # Single centered ellipse
        a0, b0 = 0.30, 0.15
        ell = Ellipse((cx, 0), 2*a0, 2*b0, angle=30,
                      facecolor=C_ELLIPSE, edgecolor="#0277bd",
                      linewidth=2.0, alpha=0.75, zorder=3)
        ax_A.add_patch(ell)
        ax_A.annotate("a = 30 mm\nb = 15 mm", (cx, 0), fontsize=7,
                      ha="center", va="center", color="#01579b", fontweight="bold")

    elif ci == 1:
        # Three ellipses at 120° at r = 0.5Rc
        r_ring = 0.5
        a_i, b_i = 0.30/np.sqrt(3), 0.15/np.sqrt(3)  # same total area
        for k, phi in enumerate([90, 210, 330]):
            phir = np.radians(phi)
            ex = cx + r_ring * np.cos(phir)
            ey = r_ring * np.sin(phir)
            ell = Ellipse((ex, ey), 2*a_i, 2*b_i, angle=phi+30,
                          facecolor=C_ELLIPSE, edgecolor="#0277bd",
                          linewidth=1.8, alpha=0.80, zorder=3)
            ax_A.add_patch(ell)
        # Ring dotted circle
        ax_A.plot(cx + r_ring*np.cos(theta_c), r_ring*np.sin(theta_c),
                  color="#aaa", linewidth=0.8, linestyle=":", zorder=2)
        ax_A.text(cx, 0, r"$r = 0.5R_c$", ha="center", va="center",
                  fontsize=7, color="#555")
        # Radial dimension
        ax_A.annotate("", xy=(cx + r_ring*np.cos(np.radians(90)),
                               r_ring*np.sin(np.radians(90))),
                      xytext=(cx, 0),
                      arrowprops=dict(arrowstyle="<->", color="#777", lw=1.2))

    else:
        # Four circular holes at 90°
        r_ring = 0.52
        r_h = 0.22  # hole radius (same total area: 4*pi*r_h^2 = pi*0.30*0.15 → r_h ≈ 0.212)
        r_h = np.sqrt(0.30*0.15)  # actual matching
        for k, phi in enumerate([45, 135, 225, 315]):
            phir = np.radians(phi)
            hx = cx + r_ring * np.cos(phir)
            hy = r_ring * np.sin(phir)
            circ = Circle((hx, hy), r_h,
                           facecolor=C_ELLIPSE, edgecolor="#0277bd",
                           linewidth=1.5, alpha=0.80, zorder=3)
            ax_A.add_patch(circ)
        ax_A.plot(cx + r_ring*np.cos(theta_c), r_ring*np.sin(theta_c),
                  color="#aaa", linewidth=0.8, linestyle=":", zorder=2)

for ci, (cx, lbl) in enumerate(zip(centers_x, labels_opt)):
    ax_A.text(cx, -1.30, lbl, ha="center", va="top", fontsize=7.5,
              color="#1a252f",
              fontweight="bold" if ci == 1 else "normal")

ax_A.set_xlim(-4.5, 4.5)
ax_A.set_ylim(-1.8, 1.8)
ax_A.set_aspect("equal")
ax_A.set_axis_off()
ax_A.set_title("(a)  Nozzle-cut options — top view of cylinder base\n"
               "(all options drawn to same total flow area)",
               fontsize=9, fontweight="bold")
# Recommended label
ax_A.text(0, 1.45, "★ Recommended", ha="center", fontsize=8.5,
          color="#2e7d32", fontweight="bold")


# ── Panel B: Multi-jet cross-section (side view, y = 0 plane) ───────────────
ax_B = fig45.add_subplot(gs45[0, 1])

R_cy = 50.0   # mm
H_cy = 140.0  # mm

# Cylinder walls
ax_B.plot([-R_cy, -R_cy], [0, H_cy], color="#37474f", linewidth=2.5)
ax_B.plot([ R_cy,  R_cy], [0, H_cy], color="#37474f", linewidth=2.5)
ax_B.plot([-R_cy, R_cy], [H_cy, H_cy], color=C_CROWN, linewidth=3.0)  # crown
ax_B.fill([-R_cy, R_cy, R_cy, -R_cy], [H_cy, H_cy, H_cy+8, H_cy+8],
          color=C_CROWN, alpha=0.7)

# Chamber fill (light blue background)
ax_B.fill_betweenx([0, H_cy], -R_cy, R_cy, color="#e3f2fd", alpha=0.30)

# Three nozzle positions (3 ellipses at 120° — we see 3 in cross-section at y=0)
# Two side nozzles visible in this cross-section plane
r_ring = 0.5 * R_cy   # 25 mm offset
nozzle_xs = [-r_ring, r_ring]     # the two nozzles in the cross-section plane
# Third nozzle is at phi=270 (straight down, but at y=0) -- show dotted
for nx, label in zip(nozzle_xs, ["Nozzle A", "Nozzle B"]):
    # Jet trajectory (slightly angled inward, converging)
    angle_in = np.radians(8)  # 8° inward tilt
    jx_end = nx - np.sign(nx) * H_cy * np.tan(angle_in)
    ax_B.annotate("", xy=(jx_end * 0.82, H_cy * 0.82),
                  xytext=(nx, 2),
                  arrowprops=dict(arrowstyle="-|>", color=C_JET, lw=2.2))
    # Nozzle mark
    ax_B.plot(nx, 0, "o", color=C_ELLIPSE, markersize=12, zorder=6,
              markeredgecolor="#0277bd", markeredgewidth=1.5)
    ax_B.text(nx, -8, label, ha="center", fontsize=7.5, color=C_JET,
              fontweight="bold")

# Third nozzle (at 90° — comes from behind, projected as centered)
ax_B.plot(0, 0, "s", color=C_ELLIPSE, markersize=10, zorder=6,
          markeredgecolor="#0277bd", markeredgewidth=1.5, alpha=0.5)
ax_B.text(0, -8, "Nozzle C\n(into page)", ha="center", fontsize=7.0,
          color="#0277bd", style="italic")

# Impingement zone (where jets converge)
h_imp = H_cy * 0.38
imp_r = 14.0
imp_ell = Ellipse((0, h_imp), 2*imp_r, 2*imp_r*0.6,
                   facecolor="#ffd54f", edgecolor="#f9a825",
                   linewidth=2.0, alpha=0.80, zorder=5)
ax_B.add_patch(imp_ell)
ax_B.text(imp_r + 4, h_imp, "Impingement\nzone", fontsize=7.5,
          color="#f57f17", va="center", fontweight="bold")
ax_B.annotate("", xy=(imp_r, h_imp), xytext=(imp_r + 22, h_imp),
              arrowprops=dict(arrowstyle="->", color="#f57f17", lw=1.3))

# Swirl arrows at upper chamber
for y_sw, r_sw in [(H_cy*0.75, 35), (H_cy*0.65, 30)]:
    arc_th = np.linspace(np.pi*0.1, np.pi*1.7, 60)
    ax_B.plot(r_sw*np.cos(arc_th), y_sw + r_sw*0.25*np.sin(arc_th),
              color=C_VAPOR, linewidth=1.4, alpha=0.60)

# Tangential slot entry (right side)
ax_B.fill_betweenx([20, 46], [R_cy, R_cy], [R_cy+10, R_cy+10],
                   color=C_SLOT, alpha=0.80)
ax_B.text(R_cy + 13, 33, "Tang.\nslot", ha="left", fontsize=7.5,
          color=C_SLOT, fontweight="bold")

# Vapour swirl annotation
ax_B.text(-R_cy + 4, H_cy*0.72, "Swirl\ncontact\nzone", fontsize=7.5,
          color="#1565c0", ha="left", va="center")

# Crown exit arrow
ax_B.annotate("", xy=(0, H_cy + 28), xytext=(0, H_cy + 8),
              arrowprops=dict(arrowstyle="-|>", color=C_VAPOR, lw=2.0))
ax_B.text(0, H_cy + 32, "Clean vapour\nexit", ha="center", fontsize=7.5,
          color=C_VAPOR, fontweight="bold")

ax_B.set_xlim(-R_cy - 30, R_cy + 55)
ax_B.set_ylim(-20, H_cy + 50)
ax_B.set_aspect("equal")
ax_B.set_xlabel("Lateral position (mm)", fontsize=8)
ax_B.set_ylabel("Height in chamber (mm)", fontsize=8)
ax_B.set_title("(b)  Three-nozzle cross-section (side view)\n"
               "Jets from nozzles A/B/C converge at impingement zone ≈ 0.38Hc",
               fontsize=9, fontweight="bold")
ax_B.grid(True, alpha=0.18, linestyle="--")


# ── Panel C: Wavy deck — chamber at CREST, clear elevation ──────────────────
ax_C = fig45.add_subplot(gs45[1, 0])

ALPHA = np.radians(27.5)
pw    = 120.0
hw    = pw/2 * np.tan(ALPHA)

# One full wave (trough – crest – trough)
xw = np.array([0.0, pw/2, pw])
yw = np.array([0.0, hw,   0.0])
ax_C.fill_between(xw, yw, y2=-18, color=C_WAVE, alpha=0.30)
ax_C.plot(xw, yw, color="#37474f", linewidth=2.8)

# Sump below
ax_C.fill_between([0, pw], [-18, -18], [-38, -38], color=C_LIQUID, alpha=0.35)
ax_C.text(pw/2, -28, "Sump (liquid)", ha="center", fontsize=8, color="#01579b")

# Liquid fills trough to nozzle level
ax_C.fill_between([-8, 8], [-18, -18], [0, 0], color=C_LIQUID, alpha=0.50)
ax_C.fill_between([pw-8, pw+8], [-18, -18], [0, 0], color=C_LIQUID, alpha=0.50)

# Nozzle at trough (circles)
for xn in [0.0, pw]:
    ax_C.plot(xn, 0, "o", color=C_ELLIPSE, markersize=10, zorder=6,
              markeredgecolor="#0277bd", markeredgewidth=1.5)
    ax_C.text(xn, -10, "Elliptical\nnozzle", ha="center", fontsize=6.5,
              color=C_JET)

# Liquid jet rising from trough to crest level (free-flight zone)
ax_C.annotate("", xy=(0, hw*1.55), xytext=(0, 2),
              arrowprops=dict(arrowstyle="-|>", color=C_JET, lw=2.2,
                              linestyle="dashed"))
ax_C.annotate("", xy=(pw, hw*1.55), xytext=(pw, 2),
              arrowprops=dict(arrowstyle="-|>", color=C_JET, lw=2.2,
                              linestyle="dashed"))

# Free-flight / axis-switching zone annotation
ax_C.annotate("", xy=(0, hw), xytext=(0, 1),
              arrowprops=dict(arrowstyle="<->", color="#555", lw=1.4))
ax_C.text(-12, hw/2, f"$h_w$ = {hw:.1f} mm\n(free-flight\naxis-switching\nzone)",
          ha="right", fontsize=7, color="#555", va="center")

# Chamber at CREST
ch_w = 32.0
ch_h = 62.0
ch_rect = FancyBboxPatch((pw/2 - ch_w/2, hw), ch_w, ch_h,
                          boxstyle="round,pad=1.5",
                          facecolor=C_CHAMBER, edgecolor=C_CHAMBER,
                          alpha=0.45, linewidth=2.0)
ax_C.add_patch(ch_rect)
# Crown disc
ax_C.fill_between([pw/2 - ch_w/2, pw/2 + ch_w/2],
                  [hw + ch_h - 3, hw + ch_h - 3],
                  [hw + ch_h + 3, hw + ch_h + 3],
                  color=C_CROWN, alpha=0.80)

# Elliptical nozzle opening at chamber floor (visible as ellipse at crest level)
ell_ch = Ellipse((pw/2, hw), 9, 4.5, angle=0,
                  facecolor=C_ELLIPSE, edgecolor="#0277bd",
                  linewidth=1.8, alpha=0.80, zorder=7)
ax_C.add_patch(ell_ch)
ax_C.text(pw/2 + ch_w/2 + 4, hw, "Elliptical\nopening\nat crest", fontsize=6.5,
          color=C_JET, va="center")

# Chamber label
ax_C.text(pw/2, hw + ch_h/2, "Swirl\nchamber\nat CREST", ha="center", va="center",
          fontsize=8, color="white", fontweight="bold")

# Vapour path arrows (along wave face toward chamber)
x_vap = np.linspace(pw + 2, pw/2 + ch_w/2, 20)
y_vap = np.linspace(3, hw + ch_h*0.25, 20)
ax_C.annotate("", xy=(pw/2 + ch_w/2 + 2, hw + 18), xytext=(pw - 5, 6),
              arrowprops=dict(arrowstyle="-|>", color=C_VAPOR, lw=1.8,
                              connectionstyle="arc3,rad=0.25"))
ax_C.text(pw*0.80, hw*0.35, "Vapour path\n(along wave face)", fontsize=7,
          color=C_VAPOR, ha="center", rotation=-52)

# Clean vapour exit arrow above crown
ax_C.annotate("", xy=(pw/2, hw + ch_h + 22), xytext=(pw/2, hw + ch_h + 5),
              arrowprops=dict(arrowstyle="-|>", color=C_VAPOR, lw=1.8))

# Wave angle arc
th_arc = np.linspace(0, ALPHA, 40)
r_arc = 22
ax_C.plot(xw[2] + r_arc*np.cos(np.pi - th_arc),
          r_arc*np.sin(th_arc),
          color="#c62828", linewidth=1.8)
ax_C.text(xw[2] - r_arc*1.45, 8, f"α={27.5:.1f}°", fontsize=8,
          color="#c62828", fontweight="bold")

ax_C.set_xlim(-38, pw + 38)
ax_C.set_ylim(-42, hw + ch_h + 34)
ax_C.set_aspect("equal")
ax_C.set_xlabel("Lateral position (mm)", fontsize=8)
ax_C.set_ylabel("Height (mm)", fontsize=8)
ax_C.set_title("(c)  Wavy deck — chambers at CRESTS (correct arrangement)\n"
               "Free-flight zone = axis-switching pre-contact  ·  Nozzles at troughs",
               fontsize=9, fontweight="bold")
ax_C.grid(True, alpha=0.18, linestyle="--")


# ── Panel D: Internal chamber features ──────────────────────────────────────
ax_D = fig45.add_subplot(gs45[1, 1])

R_d = 40.0
H_d = 120.0
configs = [
    ("(i) Smooth\nwall", -110.0,  [], "Baseline\nN_OG baseline"),
    ("(ii) Helical\nguide vane", 0.0, ["helix"], "+15–25% N_OG\nreduced dead core"),
    ("(iii) Ring\nbaffle at ½H", 110.0, ["ring"],  "+10–20% k_G,k_L\nhigh-shear zone"),
]

for label, cx, features, annotation in configs:
    # Cylinder walls
    ax_D.plot([cx-R_d, cx-R_d], [0, H_d], color="#37474f", linewidth=2.0)
    ax_D.plot([cx+R_d, cx+R_d], [0, H_d], color="#37474f", linewidth=2.0)
    # Crown
    ax_D.fill_between([cx-R_d, cx+R_d], [H_d, H_d], [H_d+7, H_d+7],
                      color=C_CROWN, alpha=0.75)
    # Vapour swirl lines
    for hy in [0.65*H_d, 0.78*H_d, 0.88*H_d]:
        th_sw = np.linspace(0, 2*np.pi*0.85, 60)
        r_sw  = (R_d - 4) * (1 - 0.05*hy/H_d)
        ax_D.plot(cx + r_sw*np.cos(th_sw)*np.linspace(0.3, 1, 60),
                  hy + (R_d*0.2)*np.sin(th_sw),
                  color="#90caf9", lw=0.9, alpha=0.60)

    # SMOOTH (i): dead core shading
    if not features:
        ax_D.fill_betweenx([5, H_d-5], cx-8, cx+8, color="#ffccbc", alpha=0.60)
        ax_D.text(cx, H_d*0.45, "Dead\ncore", ha="center", fontsize=6.5,
                  color="#bf360c", fontweight="bold")

    # HELIX (ii): helical coil drawn as sinusoidal line on wall
    if "helix" in features:
        z_h = np.linspace(5, H_d-5, 200)
        x_h = cx + (R_d-6) * np.cos(z_h * 2*np.pi / 35.0 + np.pi/2)
        # Plot helix as projection
        ax_D.plot(x_h, z_h, color="#43a047", linewidth=2.0, alpha=0.85)
        ax_D.text(cx, H_d*0.50, "Guide\nvane", ha="center", fontsize=6.5,
                  color="#2e7d32", fontweight="bold")

    # RING BAFFLE (iii): horizontal obstruction at H/2
    if "ring" in features:
        r_ring_b = R_d * 0.40   # inner radius of annular baffle
        ax_D.fill_between([cx - R_d, cx - r_ring_b], [H_d*0.50, H_d*0.50],
                          [H_d*0.50 + 5, H_d*0.50 + 5], color="#8d6e63", alpha=0.90)
        ax_D.fill_between([cx + r_ring_b, cx + R_d], [H_d*0.50, H_d*0.50],
                          [H_d*0.50 + 5, H_d*0.50 + 5], color="#8d6e63", alpha=0.90)
        ax_D.text(cx, H_d*0.50 + 9, f"Ring\nbaffle", ha="center", fontsize=6.5,
                  color="#4e342e", fontweight="bold")
        # High-shear zone annotation
        ax_D.fill_between([cx - R_d, cx + R_d],
                          [H_d*0.50-5, H_d*0.50-5],
                          [H_d*0.50+10, H_d*0.50+10],
                          color="#fff176", alpha=0.40)

    # Nozzle at base
    ax_D.plot(cx, 0, "o", color=C_ELLIPSE, markersize=8, zorder=6,
              markeredgecolor="#0277bd", markeredgewidth=1.5)

    # Label + performance annotation
    ax_D.text(cx, -10, label, ha="center", fontsize=7.5, fontweight="bold",
              color="#1a252f")
    ax_D.text(cx, H_d + 12, annotation, ha="center", fontsize=7, color="#1b5e20",
              fontweight="bold")

    # Tangential slot (right side)
    ax_D.fill_between([cx + R_d, cx + R_d + 8], [10, 10], [32, 32],
                      color=C_SLOT, alpha=0.80)

ax_D.set_xlim(-165, 165)
ax_D.set_ylim(-22, H_d + 38)
ax_D.set_aspect("equal")
ax_D.set_axis_off()
ax_D.set_title("(d)  Internal chamber features — effect on turbulence & mass transfer\n"
               "(i) smooth wall  ·  (ii) helical guide vane  ·  (iii) ring orifice baffle",
               fontsize=9, fontweight="bold")


fig45.suptitle(
    "CIST Design Refinements\n"
    "(a) Nozzle-cut options  ·  (b) Multi-jet impingement  "
    "·  (c) Wavy-deck chamber position  ·  (d) Internal features",
    fontsize=12, fontweight="bold", color="#1a252f"
)

fig45.savefig("output/45_cist_design_details.png", dpi=150, bbox_inches="tight", facecolor=BG)
plt.close(fig45)
print("Saved  output/45_cist_design_details.png")
print("Done.")
