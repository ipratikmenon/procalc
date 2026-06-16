"""
make_cist_drawings.py
Generates three schematic figures for the Cascade Impingement Swirl Tray (CIST).
Outputs saved to ./output/
"""

import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyArrowPatch, Arc, Ellipse, FancyBboxPatch, Rectangle
import matplotlib.patheffects as pe
import os

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), 'output')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ─────────────────────────────────────────────────────────────────────────────
# FIGURE 1 — CIST Chamber Cross-Section Elevation
# ─────────────────────────────────────────────────────────────────────────────

def make_figure1():
    fig, ax = plt.subplots(figsize=(7, 10))
    ax.set_xlim(-110, 200)
    ax.set_ylim(-50, 230)
    ax.set_aspect('equal')
    ax.axis('off')

    fig.patch.set_facecolor('white')

    # ── dimensions (all in mm, approximate) ──
    cx = 0          # centre x
    r = 35          # chamber radius (70 mm diam)
    floor_y = 0     # floor of chamber
    top_y = 105     # top of chamber

    # ── SUMP box ──
    sump_h = 30
    sump_w = 90
    sump = FancyBboxPatch((cx - sump_w/2, -sump_h - 5), sump_w, sump_h,
                          boxstyle="square,pad=0", linewidth=1.5,
                          edgecolor='#333333', facecolor='#cccccc')
    ax.add_patch(sump)
    ax.text(cx, -sump_h/2 - 5 + sump_h/2, 'SUMP — liquid pool',
            ha='center', va='center', fontsize=8, color='#333333', fontweight='bold')

    # ── Chamber walls ──
    wall_lw = 2.5
    ax.plot([-r, -r], [floor_y, top_y], color='#222222', lw=wall_lw)  # left wall
    ax.plot([ r,  r], [floor_y, top_y], color='#222222', lw=wall_lw)  # right wall

    # ── Chamber floor plate (horizontal line with hatch) ──
    floor_patch = Rectangle((-r, floor_y - 4), 2*r, 4,
                             facecolor='#888888', edgecolor='#222222', lw=1.2)
    ax.add_patch(floor_patch)

    # ── Elliptical nozzle in floor ──
    nozzle = Ellipse((cx, floor_y - 2), width=18, height=4,
                     facecolor='#eeeeee', edgecolor='#111111', lw=1.5)
    ax.add_patch(nozzle)
    ax.annotate('Elliptical nozzle (2:1)',
                xy=(cx + 9, floor_y - 2), xytext=(65, floor_y - 2),
                fontsize=7.5, color='#111111',
                arrowprops=dict(arrowstyle='->', color='#444444', lw=1.0),
                va='center')

    # ── Liquid jet arrow (dashed, blue) ──
    ax.annotate('', xy=(cx, floor_y + 30), xytext=(cx, -5),
                arrowprops=dict(arrowstyle='->', color='#1a6faf', lw=1.8,
                                linestyle='dashed',
                                connectionstyle='arc3,rad=0'))
    # dashed line below floor
    ax.plot([cx, cx], [-5, -sump_h + 2], color='#1a6faf', lw=1.8, ls='--')
    ax.text(cx + 5, floor_y + 16, 'Liquid jet', fontsize=7.5, color='#1a6faf',
            va='center', ha='left', style='italic')

    # ── Swirl chamber wall labels ──
    ax.text(-r - 5, top_y/2 + 10, 'Swirl\nchamber\nwall',
            fontsize=7.5, ha='right', va='center', color='#444444')
    ax.annotate('', xy=(-r - 1, top_y/2 + 10), xytext=(-r - 5, top_y/2 + 10),
                arrowprops=dict(arrowstyle='->', color='#444444', lw=0.8))

    # ── Tangential vapor inlet arrow (from right) ──
    vapor_y = 15
    ax.annotate('', xy=(r, vapor_y), xytext=(100, vapor_y),
                arrowprops=dict(arrowstyle='->', color='#c0392b', lw=2.0))
    ax.text(105, vapor_y, 'Vapor\n(tangential inlet)', fontsize=7.5,
            color='#c0392b', va='center', ha='left', fontweight='bold')

    # ── Swirl arrows inside chamber ──
    # Draw a few curved arc arrows to simulate swirl
    for angle_start, angle_extent in [(30, 200), (220, 200)]:
        arc = Arc((cx, top_y/2), width=40, height=50,
                  angle=0, theta1=angle_start, theta2=angle_start + angle_extent,
                  color='#777777', lw=1.2, linestyle='-')
        ax.add_patch(arc)
    # arrowhead at end of arcs to suggest rotation
    ax.annotate('', xy=(-8, top_y/2 + 25), xytext=(-9, top_y/2 + 22),
                arrowprops=dict(arrowstyle='->', color='#777777', lw=1.2))
    ax.annotate('', xy=(8, top_y/2 - 25), xytext=(9, top_y/2 - 22),
                arrowprops=dict(arrowstyle='->', color='#777777', lw=1.2))

    # ── Droplet cloud (scatter dots) ──
    rng = np.random.default_rng(42)
    n_drops = 40
    theta_d = rng.uniform(0, 2*np.pi, n_drops)
    rad_d = rng.uniform(0, 28, n_drops)
    xd = cx + rad_d * np.cos(theta_d)
    yd = top_y/2 + rad_d * np.sin(theta_d) * 0.7
    # only plot drops inside cylinder
    mask = (xd > -r + 3) & (xd < r - 3) & (yd > floor_y + 5) & (yd < top_y - 10)
    ax.scatter(xd[mask], yd[mask], s=6, color='#1a6faf', alpha=0.6, zorder=3)
    ax.text(-r + 5, top_y/2 + 30, 'Droplet\ncontact zone', fontsize=7.5,
            color='#1a6faf', ha='left', va='bottom', style='italic')

    # ── Crown vanes (hatched bar at top) ──
    crown_h = 8
    crown = Rectangle((-r, top_y), 2*r, crown_h,
                      facecolor='#aaaaaa', edgecolor='#222222', lw=1.5,
                      hatch='/////')
    ax.add_patch(crown)
    ax.text(cx, top_y + crown_h/2, 'Crown vanes',
            ha='center', va='center', fontsize=7.5, color='#111111', fontweight='bold')

    crown_top = top_y + crown_h

    # ── Clean vapor arrow upward ──
    ax.annotate('', xy=(cx - 10, crown_top + 22), xytext=(cx - 10, crown_top + 2),
                arrowprops=dict(arrowstyle='->', color='#c0392b', lw=1.8))
    ax.text(cx - 10, crown_top + 25, 'Clean vapor →', fontsize=7.5,
            color='#c0392b', ha='center', va='bottom')

    # ── Separated liquid arrow outward ──
    ax.annotate('', xy=(r + 30, crown_top + 4), xytext=(r + 2, crown_top + 4),
                arrowprops=dict(arrowstyle='->', color='#1a6faf', lw=1.8))
    ax.text(r + 32, crown_top + 4, 'Separated\nliquid →', fontsize=7.5,
            color='#1a6faf', ha='left', va='center')

    # ── Collection ring (dashed box right side) ──
    ring_x = r + 30
    ring_y = 55
    ring_w = 55
    ring_h = 50
    ring_box = FancyBboxPatch((ring_x, ring_y), ring_w, ring_h,
                               boxstyle="square,pad=2", linewidth=1.2,
                               edgecolor='#1a6faf', facecolor='#ddeeff',
                               linestyle='--')
    ax.add_patch(ring_box)
    ax.text(ring_x + ring_w/2, ring_y + ring_h/2 + 5, 'Collection\nring',
            ha='center', va='center', fontsize=7.5, color='#1a6faf', fontweight='bold')

    # line from separated liquid arrow to collection ring
    ax.plot([ring_x + ring_w/2, ring_x + ring_w/2], [crown_top + 4, ring_y + ring_h],
            color='#1a6faf', lw=1.0, ls='--')

    # ── Drain downward from collection ring ──
    ax.annotate('', xy=(ring_x + ring_w/2, ring_y - 18), xytext=(ring_x + ring_w/2, ring_y),
                arrowprops=dict(arrowstyle='->', color='#1a6faf', lw=1.5))
    ax.text(ring_x + ring_w/2, ring_y - 22, 'Drain to\nlower sump',
            ha='center', va='top', fontsize=7, color='#1a6faf')

    # ── Vapor space label at very top ──
    ax.text(cx, crown_top + 45, 'Vapor space', fontsize=8,
            ha='center', va='bottom', color='#c0392b',
            bbox=dict(boxstyle='round,pad=0.3', fc='#fff0f0', ec='#c0392b', lw=0.8))

    # ── Dimension tick: 70 mm diameter ──
    tick_y = -42
    ax.annotate('', xy=(r, tick_y), xytext=(-r, tick_y),
                arrowprops=dict(arrowstyle='<->', color='#555555', lw=1.0))
    ax.text(cx, tick_y - 3, '70 mm', ha='center', va='top', fontsize=7, color='#555555')

    # ── Dimension tick: 105 mm height ──
    tick_x = -80
    ax.annotate('', xy=(tick_x, top_y), xytext=(tick_x, floor_y),
                arrowprops=dict(arrowstyle='<->', color='#555555', lw=1.0))
    ax.text(tick_x - 3, top_y/2, '105 mm', ha='right', va='center', fontsize=7,
            color='#555555', rotation=90)

    # ── Title & note ──
    ax.set_title('Fig. 1 — CIST Chamber: Cross-Section Elevation',
                 fontsize=11, fontweight='bold', pad=12)
    ax.text(0.5, -0.02, 'Schematic — not to scale',
            transform=ax.transAxes, ha='center', va='top',
            fontsize=8, style='italic', color='#666666')

    plt.tight_layout()
    out = os.path.join(OUTPUT_DIR, '40_cist_chamber_elevation.png')
    fig.savefig(out, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return out


# ─────────────────────────────────────────────────────────────────────────────
# FIGURE 2 — CIST Tray Plan View
# ─────────────────────────────────────────────────────────────────────────────

def make_figure2():
    fig, ax = plt.subplots(figsize=(8, 8))
    ax.set_aspect('equal')

    # Column: 36 inch = 914.4 mm diameter
    col_r = 914.4 / 2   # mm

    ax.set_xlim(-col_r * 1.18, col_r * 1.18)
    ax.set_ylim(-col_r * 1.18, col_r * 1.18)
    ax.axis('off')
    fig.patch.set_facecolor('white')

    # ── Column shell ──
    col_circle = plt.Circle((0, 0), col_r, fill=False,
                             edgecolor='#222222', lw=2.5)
    ax.add_patch(col_circle)

    # ── Downcomers ──
    dc_w = col_r * 0.38   # width of downcomer chord section
    dc_half_h = col_r * 0.90

    # Left downcomer (inlet)
    dc_left_x = -col_r
    dc_left = Rectangle((dc_left_x, -dc_half_h), dc_w, 2 * dc_half_h,
                         facecolor='#f5deb3', edgecolor='#8B6914', lw=1.5)
    ax.add_patch(dc_left)
    # clip to circle
    dc_left.set_clip_path(col_circle)
    ax.text(dc_left_x + dc_w/2, 0, 'Liquid\ninlet DC',
            ha='center', va='center', fontsize=8, fontweight='bold', color='#5a3e00')

    # Right downcomer (outlet)
    dc_right_x = col_r - dc_w
    dc_right = Rectangle((dc_right_x, -dc_half_h), dc_w, 2 * dc_half_h,
                          facecolor='#f5deb3', edgecolor='#8B6914', lw=1.5)
    ax.add_patch(dc_right)
    dc_right.set_clip_path(col_circle)
    ax.text(dc_right_x + dc_w/2, 0, 'Liquid\noutlet DC',
            ha='center', va='center', fontsize=8, fontweight='bold', color='#5a3e00')

    # ── Liquid collection ring (thin grey annular ring inside column, outside chamber array) ──
    ring_inner = col_r * 0.72
    ring_outer = col_r * 0.80
    ring = mpatches.Annulus((0, 0), ring_outer, ring_outer - ring_inner,
                             facecolor='#cccccc', edgecolor='#888888', lw=1.0, alpha=0.7)
    ax.add_patch(ring)
    ax.annotate('Liquid collection ring',
                xy=(ring_inner * np.cos(np.radians(50)), ring_inner * np.sin(np.radians(50))),
                xytext=(col_r * 0.45, col_r * 0.88),
                fontsize=7.5, color='#555555',
                arrowprops=dict(arrowstyle='->', color='#888888', lw=0.9))

    # ── Hex-packed CIST chamber circles ──
    # 70 mm diam, ~155 mm pitch → scale to mm
    chamber_r = 35        # mm radius
    pitch = 155           # mm hex pitch
    active_r = ring_inner * 0.97  # stay inside ring

    # Generate hex grid
    chambers = []
    rows = 8
    for row in range(-rows, rows + 1):
        y = row * pitch * np.sqrt(3) / 2
        offset = pitch / 2 if (row % 2 != 0) else 0
        for col in range(-rows, rows + 1):
            x = col * pitch + offset
            dist = np.sqrt(x**2 + y**2)
            # must be inside active ring and not inside downcomers
            in_left_dc  = (x < dc_left_x + dc_w + chamber_r + 5) and (abs(y) < dc_half_h + chamber_r)
            in_right_dc = (x > dc_right_x - chamber_r - 5) and (abs(y) < dc_half_h + chamber_r)
            if dist + chamber_r < active_r and not in_left_dc and not in_right_dc:
                chambers.append((x, y, dist))

    # Sort by distance so inner ones get plotted on top (cosmetic)
    chambers.sort(key=lambda c: c[2])

    inner_limit = col_r * 0.55
    labeled_primary = False
    labeled_outer = False

    for (x, y, d) in chambers:
        if d < inner_limit:
            fc = '#aad4f5'
            ec = '#1a6faf'
        else:
            fc = '#d0e8ff'
            ec = '#4a90c4'
        c = plt.Circle((x, y), chamber_r, facecolor=fc, edgecolor=ec, lw=1.2, zorder=3)
        ax.add_patch(c)

    # Labels for chamber zones
    ax.annotate('Primary contact\nchambers (hex-packed)',
                xy=(0, -inner_limit * 0.5),
                xytext=(col_r * 0.05, -col_r * 0.78),
                fontsize=8, color='#1a6faf', ha='center',
                arrowprops=dict(arrowstyle='->', color='#1a6faf', lw=1.0))

    # ── Column diameter annotation ──
    ann_y = -col_r * 1.08
    ax.annotate('', xy=(col_r, ann_y), xytext=(-col_r, ann_y),
                arrowprops=dict(arrowstyle='<->', color='#333333', lw=1.2))
    ax.text(0, ann_y - 18, '36 in (914 mm) column I.D.',
            ha='center', va='top', fontsize=8, color='#333333')

    # ── Tick marks on column ──
    ax.plot([-col_r, -col_r], [ann_y - 6, ann_y + 6], color='#333333', lw=1.0)
    ax.plot([ col_r,  col_r], [ann_y - 6, ann_y + 6], color='#333333', lw=1.0)

    # ── Chamber count note ──
    ax.text(0, col_r * 1.10, f'{len(chambers)} CIST chambers shown',
            ha='center', va='bottom', fontsize=8, color='#555555', style='italic')

    # ── Title ──
    ax.set_title('Fig. 2 — CIST Tray Plan View',
                 fontsize=12, fontweight='bold', pad=14)
    ax.text(0.5, -0.03, 'Schematic — not to scale',
            transform=ax.transAxes, ha='center', va='top',
            fontsize=8, style='italic', color='#666666')

    plt.tight_layout()
    out = os.path.join(OUTPUT_DIR, '41_cist_plan_view.png')
    fig.savefig(out, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return out


# ─────────────────────────────────────────────────────────────────────────────
# FIGURE 3 — Elliptical Nozzle Geometry and Axis-Switching
# ─────────────────────────────────────────────────────────────────────────────

def make_figure3():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 6))
    fig.patch.set_facecolor('white')

    # ── LEFT PANEL: Ellipse geometry ──
    ax1.set_aspect('equal')
    ax1.set_xlim(-3.5, 3.5)
    ax1.set_ylim(-2.5, 3.0)
    ax1.axis('off')
    ax1.set_title('Ellipse Geometry', fontsize=10, fontweight='bold', pad=8)

    k = 2.0
    # Let b = 1 → a = 2
    a = 2.0
    b = 1.0
    r_eq = np.sqrt(a * b)  # equivalent circle radius = sqrt(2) ≈ 1.414

    # Equivalent circle (dashed grey)
    eq_circle = plt.Circle((0, 0), r_eq, fill=False,
                            edgecolor='#888888', lw=1.5, linestyle='--', zorder=2)
    ax1.add_patch(eq_circle)
    ax1.text(r_eq * np.cos(np.radians(135)),
             r_eq * np.sin(np.radians(135)) + 0.15,
             'Equal-area\ncircle (r = √(a·b))',
             fontsize=7.5, color='#888888', ha='center')

    # Ellipse fill
    ell = Ellipse((0, 0), width=2*a, height=2*b,
                  facecolor='#aad4f5', edgecolor='#1a6faf', lw=2.0, zorder=3)
    ax1.add_patch(ell)

    # Axis lines
    ax1.annotate('', xy=(a + 0.1, 0), xytext=(-a - 0.1, 0),
                arrowprops=dict(arrowstyle='<->', color='#1a6faf', lw=1.5))
    ax1.annotate('', xy=(0, b + 0.1), xytext=(0, -b - 0.1),
                arrowprops=dict(arrowstyle='<->', color='#c0392b', lw=1.5))

    # Labels
    ax1.text(a + 0.15, 0.15, 'a = major axis', fontsize=8, color='#1a6faf', va='bottom')
    ax1.text(-0.15, b + 0.15, 'b = minor axis', fontsize=8, color='#c0392b',
             ha='right', va='bottom')

    # Aspect ratio box
    ax1.text(0, -b - 0.55, f'Aspect ratio  k = a/b = {k:.1f}',
             ha='center', va='top', fontsize=9, fontweight='bold',
             bbox=dict(boxstyle='round,pad=0.4', fc='#fffce0', ec='#ccaa00', lw=1.2))

    # r_eq annotation
    theta_ann = np.radians(45)
    ax1.annotate('',
                xy=(r_eq * np.cos(theta_ann), r_eq * np.sin(theta_ann)),
                xytext=(0, 0),
                arrowprops=dict(arrowstyle='->', color='#555555', lw=1.0))
    ax1.text(r_eq * np.cos(theta_ann) + 0.1, r_eq * np.sin(theta_ann) + 0.1,
             'r = √(a·b)', fontsize=7.5, color='#555555')

    # b vs r relationship
    ax1.text(0, 2.7,
             f'For k=2:   a = r√2,   b = r/√2\n(b < r < a)',
             ha='center', va='top', fontsize=8, color='#333333',
             bbox=dict(boxstyle='round,pad=0.35', fc='#f0f0f0', ec='#bbbbbb', lw=1.0))

    ax1.text(0, -2.3, 'Equal flow area', ha='center', va='bottom',
             fontsize=8, style='italic', color='#555555',
             bbox=dict(boxstyle='rarrow,pad=0.3', fc='#eaf4ff', ec='#1a6faf', lw=0.8))

    # ── RIGHT PANEL: Axis-switching schematic ──
    ax2.set_xlim(-1.5, 5.5)
    ax2.set_ylim(-2.2, 3.0)
    ax2.set_aspect('equal')
    ax2.axis('off')
    ax2.set_title('Axis-Switching Phenomenon', fontsize=10, fontweight='bold', pad=8)

    # Three positions along jet travel (horizontal layout)
    positions = [0, 2.0, 4.0]
    labels_z  = ['z = 0\n(nozzle exit)', 'z = L_AS / 2\n(axes switching)', 'z = L_AS\n(axes switched)']
    # Ellipses: at z=0 major horiz, z=mid circular, z=end major vertical
    ellipse_params = [
        (1.0, 0.5),   # a_x, a_y at z=0: wide horizontal
        (0.72, 0.72), # circular midpoint
        (0.5, 1.0),   # wide vertical at z=L_AS
    ]
    colors_fill = ['#aad4f5', '#c8e6c9', '#f5c6a0']
    colors_edge = ['#1a6faf', '#2e7d32', '#bf360c']

    for i, (xc, (aw, ah), zlab, fc, ec) in enumerate(
            zip(positions, ellipse_params, labels_z, colors_fill, colors_edge)):
        e = Ellipse((xc, 0), width=2*aw, height=2*ah,
                    facecolor=fc, edgecolor=ec, lw=2.0, zorder=3)
        ax2.add_patch(e)
        ax2.text(xc, -ah - 0.35, zlab, ha='center', va='top', fontsize=7.5,
                 color=ec, fontweight='bold')

    # Dashed centerline arrow
    ax2.annotate('', xy=(4.0 + 0.55, 0), xytext=(-0.55, 0),
                arrowprops=dict(arrowstyle='->', color='#555555', lw=1.2,
                                linestyle='dashed'))
    ax2.text(4.6, 0.08, 'Jet travel\n(→ downward in chamber)',
             fontsize=7, color='#555555', va='bottom', ha='left')

    # Brace for L_AS
    brace_y = 1.35
    brace_x0 = positions[0]
    brace_x1 = positions[2]
    # Draw a simple bracket
    ax2.annotate('', xy=(brace_x1, brace_y), xytext=(brace_x0, brace_y),
                arrowprops=dict(arrowstyle='<->', color='#333333', lw=1.2))
    ax2.plot([brace_x0, brace_x0], [1.10, brace_y], color='#333333', lw=0.8)
    ax2.plot([brace_x1, brace_x1], [1.10, brace_y], color='#333333', lw=0.8)
    ax2.text((brace_x0 + brace_x1) / 2, brace_y + 0.12,
             'Axis-switching length  $L_{AS}$',
             ha='center', va='bottom', fontsize=8.5, color='#333333', fontweight='bold')

    # Axis direction labels for first ellipse
    ax2.annotate('', xy=(positions[0] + 1.1, 0), xytext=(positions[0], 0),
                arrowprops=dict(arrowstyle='->', color='#1a6faf', lw=1.0))
    ax2.text(positions[0] + 1.15, -0.15, 'major', fontsize=7, color='#1a6faf')

    # Axis direction labels for last ellipse
    ax2.annotate('', xy=(positions[2], 1.1), xytext=(positions[2], 0),
                arrowprops=dict(arrowstyle='->', color='#bf360c', lw=1.0))
    ax2.text(positions[2] + 0.08, 1.15, 'major', fontsize=7, color='#bf360c')

    # Note box
    ax2.text(2.0, -1.9,
             'The elliptical jet spontaneously rotates its cross-section\n'
             'due to surface-tension and inertial axis-switching instability.',
             ha='center', va='top', fontsize=7.5, color='#333333',
             bbox=dict(boxstyle='round,pad=0.4', fc='#f8f8f8', ec='#aaaaaa', lw=1.0))

    # ── Overall title ──
    fig.suptitle('Fig. 3 — Elliptical Nozzle: Geometry and Axis-Switching Phenomenon',
                 fontsize=12, fontweight='bold', y=1.01)
    ax1.text(0.5, -0.04, 'Schematic — not to scale',
             transform=ax1.transAxes, ha='center', va='top',
             fontsize=7.5, style='italic', color='#666666')

    plt.tight_layout()
    out = os.path.join(OUTPUT_DIR, '42_cist_ellipse_nozzle.png')
    fig.savefig(out, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return out


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    files = []
    files.append(make_figure1())
    files.append(make_figure2())
    files.append(make_figure3())
    print('Done:', files)
