#!/usr/bin/env python3
"""HyTrays — schematic technical drawings for the Apex technical paper.

These are *conceptual* engineering schematics (not CAD/scaled drawings):
proportions follow Table 2 of ``output/apex_paper_tables.md`` (the General
Rectification design point, D = 32 in) where practical, but cartridge counts,
slot counts etc. are illustrative. They exist to make the Integrated
Cartridge-Grid architecture (``HyTrays_Apex_concept.md`` Section 3) and the
conventional-tray comparison (Section 5 of the technical paper) legible at a
glance.

Outputs (``HyTrays/output/``)
------------------------------
    30_apex_cross_section.png        -- Apex deck cross-section (cutaway)
    31_apex_plan_view.png             -- Apex deck plan view (top-down)
    32_tray_family_cross_sections.png -- 7-panel comparison vs. conventional trays

Usage
-----
    python make_apex_drawings.py
"""
from __future__ import annotations

import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Circle, Ellipse, FancyArrow, Polygon, Rectangle, Wedge

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")

C_DECK = "#8C8C8C"
C_LIQUID = "#9FD3E8"
C_FROTH = "#CFEAF5"
C_CARTRIDGE = "#4C72B0"
C_CAP = "#2F4B7C"
C_LIP = "#DD8452"
C_AEGIS = "#55A868"
C_VAPOR = "#C44E52"
C_PULSAR = "#8172B2"


# ══════════════════════════════════════════════════════════════════════════
# Figure 30 -- Apex cross-section (cutaway elevation)
# ══════════════════════════════════════════════════════════════════════════
def draw_cross_section(ax) -> None:
    D = 32.0  # in, General Rectification design point (Table 2)
    deck_y = 0.0
    deck_t = 0.6

    # zone boundaries (Table 2: primary 70%, secondary 15%, downcomers 15%
    # split 7.5% each side)
    x0, x1 = 0.0, D
    dc_w = 0.075 * D          # 2.4 in each side
    sec_w = 0.15 * D           # 4.8 in
    pri_w = 0.70 * D           # 22.4 in
    x_pri0 = dc_w
    x_pri1 = x_pri0 + pri_w
    x_sec1 = x_pri1 + sec_w
    assert abs(x_sec1 + dc_w - D) < 1e-6

    # ── deck plate (with gaps left open visually for cartridges/lips) ──
    ax.add_patch(Rectangle((x0, deck_y - deck_t), D, deck_t,
                            facecolor=C_DECK, edgecolor="k", zorder=2))

    # ── downcomers (liquid chutes, both sides) ──────────────────────────
    for xa in (x0, D - dc_w):
        ax.add_patch(Rectangle((xa, deck_y - deck_t - 7.5), dc_w, 8.0,
                                facecolor=C_LIQUID, edgecolor="k",
                                alpha=0.6, zorder=1))
    ax.text(dc_w / 2, -5.5, "downcomer\n(in)", ha="center", va="center", fontsize=7)
    ax.text(D - dc_w / 2, -5.5, "downcomer\n(out, over\nweir)", ha="center",
            va="center", fontsize=7)

    # ── liquid / froth layer across the active area ─────────────────────
    h_clear = 1.4  # in (Table 2, ~to scale)
    ax.add_patch(Rectangle((x_pri0, deck_y), pri_w + sec_w, h_clear,
                            facecolor=C_FROTH, edgecolor="none", zorder=0,
                            hatch="....", alpha=0.7))

    # ── weir before exit downcomer ───────────────────────────────────────
    weir_h = 0.5
    ax.add_patch(Rectangle((x_sec1 - 0.3, deck_y), 0.3, weir_h,
                            facecolor="#444444", zorder=3))
    ax.text(x_sec1 + 0.5, weir_h + 1.0, "weir\n0.5 in", ha="center",
            fontsize=6.5)

    # ── primary zone: swirl-tube / VORTEXA cartridges ───────────────────
    n_cart = 5
    cart_w = pri_w / n_cart * 0.6
    for i in range(n_cart):
        xc = x_pri0 + pri_w * (i + 0.5) / n_cart
        body_h = 3.2
        ax.add_patch(Rectangle((xc - cart_w / 2, deck_y), cart_w, body_h,
                                facecolor=C_CARTRIDGE, edgecolor="k", zorder=3))
        # helical-sleeve indexing marks (HT-08)
        for k in range(3):
            yk = body_h * (k + 1) / 4
            ax.plot([xc - cart_w / 2, xc + cart_w / 2], [yk, yk + 0.35],
                    color="white", lw=1.0, zorder=4)
        # vane cap (mechanical disengagement) -- domed ellipse with a gap
        cap_y = body_h + 0.9
        ax.add_patch(Ellipse((xc, cap_y), cart_w * 1.7, 1.0,
                              facecolor=C_CAP, edgecolor="k", zorder=4))
        # vapor path: up the tube, swirling out sideways under the cap
        ax.add_patch(FancyArrow(xc, deck_y - 2.2, 0, body_h + 1.4,
                                 width=0.05, head_width=0.5, head_length=0.6,
                                 length_includes_head=True,
                                 color=C_VAPOR, zorder=5))
        for sgn in (-1, 1):
            ax.add_patch(FancyArrow(xc, cap_y, sgn * cart_w * 1.1, 0.0,
                                     width=0.04, head_width=0.35, head_length=0.4,
                                     length_includes_head=True,
                                     color=C_VAPOR, zorder=5, alpha=0.85))
        # disengaged liquid falling back
        ax.add_patch(FancyArrow(xc + cart_w * 0.8, cap_y - 0.3,
                                 cart_w * 0.3, -(cap_y - h_clear - 0.3),
                                 width=0.03, head_width=0.22, head_length=0.3,
                                 length_includes_head=True,
                                 color=C_LIQUID, zorder=5))
    ax.text((x_pri0 + x_pri1) / 2, 5.6,
            "Primary zone -- HT-02/HT-08 swirl-tube\ncartridges, 70% of deck area",
            ha="center", fontsize=8, fontweight="bold")

    # ── secondary zone: lip-seal trickle valves ─────────────────────────
    n_lip = 2
    for i in range(n_lip):
        xc = x_pri1 + sec_w * (i + 0.5) / n_lip
        lw_ = sec_w / n_lip * 0.6
        # hinged lip flap, slightly open
        verts = [(xc - lw_ / 2, deck_y), (xc + lw_ / 2, deck_y),
                 (xc + lw_ / 2 - 0.3, deck_y + 0.9), (xc - lw_ / 2 + 0.1, deck_y + 0.6)]
        ax.add_patch(Polygon(verts, closed=True, facecolor=C_LIP,
                              edgecolor="k", zorder=4))
        ax.add_patch(FancyArrow(xc, deck_y - 2.2, 0, 2.0,
                                 width=0.04, head_width=0.35, head_length=0.4,
                                 length_includes_head=True,
                                 color=C_VAPOR, zorder=5, alpha=0.85))
    ax.text(x_pri1 + sec_w / 2, 3.0,
            "Secondary\nzone --\nHT-01C\nlip-seal\ntrickle,\n15%",
            ha="center", fontsize=7.5, fontweight="bold")

    # ── PULSAR self-sweeping relief slots (between cartridges, in deck) ──
    for i in range(n_cart - 1):
        xc = x_pri0 + pri_w * (i + 1) / n_cart
        ax.add_patch(Rectangle((xc - 0.35, deck_y - deck_t), 0.7, deck_t,
                                facecolor=C_PULSAR, edgecolor="k", zorder=3))
        ax.annotate("", xy=(xc, deck_y - deck_t - 1.6), xytext=(xc, deck_y - deck_t - 0.1),
                     arrowprops=dict(arrowstyle="-|>", color=C_PULSAR,
                                      connectionstyle="arc3,rad=0.6", lw=1.2))
    ax.text(x_pri0 + 1.0, deck_y - deck_t - 2.6,
            "PULSAR oscillator\nrelief slots\n(self-sweeping, 5-30 Hz)",
            fontsize=7, color=C_PULSAR, ha="left")

    # ── AEGIS structural grid (below deck) ──────────────────────────────
    for xc in (D * 0.2, D * 0.5, D * 0.8):
        ax.add_patch(Rectangle((xc - 0.6, deck_y - deck_t - 4.0), 1.2, 4.0,
                                facecolor=C_AEGIS, edgecolor="k", alpha=0.6, zorder=1))
    ax.text(D * 0.5, deck_y - deck_t - 4.6,
            "AEGIS bolted cartridge-grid (structural beams, 2-5 psi rating)",
            ha="center", fontsize=8, color="#2E5C3D", fontweight="bold")
    for xa, ha in ((0.0, "left"), (D, "right")):
        ax.annotate("hook\nclamp", xy=(xa, deck_y - deck_t - 0.3),
                     xytext=(xa + (1.5 if ha == "left" else -1.5), deck_y - deck_t - 2.5),
                     ha="center", fontsize=6.5, color="#2E5C3D",
                     arrowprops=dict(arrowstyle="-", color="#2E5C3D", lw=0.8))

    # ── vapor inlet arrows from tray below ──────────────────────────────
    ax.text(D / 2, -9.3, "vapor up from tray below", ha="center", fontsize=8,
            color=C_VAPOR, style="italic")

    # ── tray-spacing dimension ───────────────────────────────────────────
    ax.annotate("", xy=(D + 1.2, 0), xytext=(D + 1.2, 24),
                 arrowprops=dict(arrowstyle="<->", color="k", lw=1))
    ax.plot([D, D + 1.2], [0, 0], color="k", lw=0.6, ls=":")
    ax.plot([D, D + 1.2], [24, 24], color="k", lw=0.6, ls=":")
    ax.text(D + 1.8, 12, "24 in tray\nspacing", rotation=90, va="center",
            fontsize=8)
    ax.plot([0, D], [24, 24], color="k", lw=0.8, ls="--")
    ax.text(D / 2, 24.4, "underside of tray above", ha="center", fontsize=7,
            style="italic")

    ax.set_xlim(-2, D + 4)
    ax.set_ylim(-11, 27)
    ax.set_aspect("equal")
    ax.axis("off")
    ax.set_title("HyTrays Apex -- deck cross-section (cutaway)\n"
                  "General Rectification design point, D = 32 in (2.66 ft)",
                  fontsize=11, fontweight="bold")


# ══════════════════════════════════════════════════════════════════════════
# Figure 31 -- Apex plan view (top-down)
# ══════════════════════════════════════════════════════════════════════════
def draw_plan_view(ax) -> None:
    R = 16.0  # in, D = 32 in

    theta = np.linspace(0, 2 * np.pi, 200)
    ax.plot(R * np.cos(theta), R * np.sin(theta), color="k", lw=1.5)

    # downcomer chords (left = inlet, right = outlet over weir) -- ~7.5% each
    for sign, label in ((-1, "downcomer\n(liquid in)"), (1, "downcomer\n(liquid out)")):
        chord_x = sign * R * 0.80
        y_half = np.sqrt(max(R ** 2 - chord_x ** 2, 0))
        verts = [(chord_x, -y_half), (chord_x, y_half),
                 (sign * R, 0)]
        # approximate the circular segment with a polygon arc
        arc_theta = np.linspace(np.arctan2(y_half, chord_x), np.arctan2(-y_half, chord_x),
                                 30) if sign > 0 else \
            np.linspace(np.pi - np.arctan2(y_half, -chord_x),
                        np.pi + np.arctan2(y_half, -chord_x), 30)
        arc_pts = [(R * np.cos(t), R * np.sin(t)) for t in arc_theta]
        poly = [(chord_x, -y_half)] + arc_pts + [(chord_x, y_half)]
        ax.add_patch(Polygon(poly, closed=True, facecolor=C_LIQUID,
                              edgecolor="k", alpha=0.6, zorder=1))
        ax.text(sign * R * 0.92, 0, label, ha="center", va="center",
                fontsize=6.5, rotation=90 if sign < 0 else -90)

    # active deck region (between the two chords)
    x_lo, x_hi = -R * 0.80, R * 0.80

    # secondary lip-seal trickle strip along the outlet edge (15% of width)
    sec_x0 = x_hi - (x_hi - x_lo) * 0.15
    y_band = np.sqrt(np.maximum(R ** 2 - np.linspace(sec_x0, x_hi, 2) ** 2, 0))
    ax.axvspan(sec_x0, x_hi, color=C_LIP, alpha=0.18)
    for yy in np.linspace(-12, 12, 7):
        if sec_x0 ** 2 + yy ** 2 < R ** 2:
            ax.add_patch(Rectangle((sec_x0 + 0.3, yy - 0.6), 2.5, 1.2,
                                    facecolor=C_LIP, edgecolor="k", lw=0.4, zorder=2))
    ax.text((sec_x0 + x_hi) / 2, R * 1.05, "Secondary zone (15%)\nHT-01C lip-seal",
            ha="center", va="bottom", fontsize=7, fontweight="bold")

    # primary zone: hex-packed swirl-tube cartridges with GRADEX radial grading
    pri_x0, pri_x1 = x_lo, sec_x0
    pitch = 4.0
    cart_r = 1.6
    rows = np.arange(pri_x0 + pitch / 2, pri_x1, pitch)
    n_cart = 0
    for i, xc in enumerate(rows):
        y_off = (pitch / 2) if i % 2 else 0.0
        for yc in np.arange(-R + pitch / 2 + y_off, R, pitch):
            r_frac = np.hypot(xc, yc) / R
            if r_frac > 0.97:
                continue
            n_cart += 1
            # GRADEX radial grading: open-area fraction 8% (center) -> 13% (rim),
            # rendered as cartridge shading intensity
            open_frac = 0.08 + 0.05 * r_frac
            ax.add_patch(Circle((xc, yc), cart_r, facecolor=C_CARTRIDGE,
                                 edgecolor="k", lw=0.5,
                                 alpha=0.35 + 1.8 * open_frac, zorder=2))
            ax.add_patch(Circle((xc, yc), cart_r * 0.45, facecolor=C_CAP,
                                 edgecolor="none", zorder=3))

    ax.text((pri_x0 + sec_x0) / 2, R * 1.05,
            "Primary zone (70%) -- HT-02/HT-08 swirl-tube cartridges",
            ha="center", va="bottom", fontsize=7, fontweight="bold")
    ax.text((pri_x0 + sec_x0) / 2, -R - 1.9,
            f"{n_cart} cartridges shown, hex-packed @ {pitch:.0f} in pitch  --  "
            "shading = GRADEX radial grading, 8% (center) -> 13% (rim) open area",
            ha="center", va="top", fontsize=6.5, style="italic", color="0.35")

    # multi-chordal liquid sweep arrows (left -> right, fanning by chord)
    for yy in np.linspace(-11, 11, 5):
        x_start = x_lo + 0.5
        x_end = sec_x0 - 0.5
        ax.annotate("", xy=(x_end, yy * 0.95), xytext=(x_start, yy),
                     arrowprops=dict(arrowstyle="-|>", color="#1f6f1f", lw=1.0,
                                      alpha=0.7))
    ax.text(0, -R * 0.92, "multi-chordal liquid sweep (toward outlet weir)",
            ha="center", fontsize=7, color="#1f6f1f", style="italic")

    # PULSAR slots scattered between cartridges near the inlet edge
    for xc, yc in [(pri_x0 + 2.0, 6), (pri_x0 + 2.0, -6), (pri_x0 + 6.5, 0)]:
        ax.add_patch(Rectangle((xc - 0.9, yc - 0.25), 1.8, 0.5,
                                facecolor=C_PULSAR, edgecolor="k", lw=0.4, zorder=4))
    ax.text(pri_x0 + 4, -R * 0.55, "PULSAR\nrelief slots", ha="center",
            fontsize=6.5, color=C_PULSAR)

    ax.set_xlim(-R - 2, R + 2)
    ax.set_ylim(-R - 3.5, R + 3)
    ax.set_aspect("equal")
    ax.axis("off")
    ax.set_title("HyTrays Apex -- plan view (top-down)\n"
                  "D = 32 in (2.66 ft); flow left -> right",
                  fontsize=11, fontweight="bold")


# ══════════════════════════════════════════════════════════════════════════
# Figure 32 -- seven-panel tray-family cross-section comparison
# ══════════════════════════════════════════════════════════════════════════
def _panel_frame(ax, title: str) -> None:
    ax.set_xlim(0, 10)
    ax.set_ylim(-3.2, 4.5)
    ax.set_aspect("equal")
    ax.axis("off")
    ax.set_title(title, fontsize=9, fontweight="bold")
    ax.add_patch(Rectangle((0, -3.2), 10, 7.7, fill=False, edgecolor="#cccccc", lw=0.8))


def _liquid_layer(ax, x0, x1, h, weir_x=None, weir_h=0.0):
    ax.add_patch(Rectangle((x0, 0), x1 - x0, h, facecolor=C_FROTH,
                            edgecolor="none", hatch="....", alpha=0.7, zorder=0))
    if weir_x is not None:
        ax.add_patch(Rectangle((weir_x - 0.15, 0), 0.3, weir_h,
                                facecolor="#444444", zorder=3))


def _downcomer(ax, x0, x1, label=""):
    ax.add_patch(Rectangle((x0, -3.0), x1 - x0, 3.0, facecolor=C_LIQUID,
                            edgecolor="k", alpha=0.6, zorder=1))
    if label:
        ax.text((x0 + x1) / 2, -1.5, label, rotation=90, ha="center",
                va="center", fontsize=6)


def _vapor_arrow(ax, x, y0, y1, **kw):
    ax.add_patch(FancyArrow(x, y0, 0, y1 - y0, width=0.04, head_width=0.3,
                             head_length=0.3, length_includes_head=True,
                             color=C_VAPOR, zorder=5, **kw))


def panel_sieve(ax):
    _panel_frame(ax, "1. Sieve (baseline)\nf_hole=0.10, no moving parts")
    _downcomer(ax, 8.2, 9.6, "downcomer")
    ax.add_patch(Rectangle((0.4, -0.3), 7.8, 0.3, facecolor=C_DECK, edgecolor="k", zorder=2))
    for xc in np.linspace(1.2, 7.4, 7):
        ax.add_patch(Circle((xc, 0), 0.18, facecolor="white", edgecolor="k", zorder=3))
        _vapor_arrow(ax, xc, -2.2, 1.6)
    _liquid_layer(ax, 0.4, 8.2, 1.6, weir_x=8.0, weir_h=2.0)
    ax.text(4.2, 3.3, "small round holes\nuniform pattern", ha="center", fontsize=7)


def panel_dualflow(ax):
    _panel_frame(ax, "2. Dual-Flow\nf_hole=0.20, no downcomer")
    ax.add_patch(Rectangle((0.4, -0.3), 9.2, 0.3, facecolor=C_DECK, edgecolor="k", zorder=2))
    for xc in np.linspace(1.0, 9.0, 7):
        ax.add_patch(Circle((xc, 0), 0.32, facecolor="white", edgecolor="k", zorder=3))
        _vapor_arrow(ax, xc - 0.12, -2.2, 1.4)
        ax.add_patch(FancyArrow(xc + 0.12, 1.4, 0, -2.2 - 1.4, width=0.03,
                                 head_width=0.22, head_length=0.3,
                                 length_includes_head=True, color=C_LIQUID, zorder=5))
    _liquid_layer(ax, 0.4, 9.6, 1.0)
    ax.text(5, 3.3, "same large holes pass\nvapor up & liquid down", ha="center", fontsize=7)


def panel_moving_valve(ax):
    _panel_frame(ax, "3. Moving Valve\nf_hole=0.13, weight-loaded caps")
    _downcomer(ax, 8.2, 9.6, "downcomer")
    ax.add_patch(Rectangle((0.4, -0.3), 7.8, 0.3, facecolor=C_DECK, edgecolor="k", zorder=2))
    for xc in np.linspace(1.2, 7.4, 5):
        ax.add_patch(Circle((xc, 0), 0.18, facecolor="white", edgecolor="k", zorder=3))
        ax.add_patch(Ellipse((xc, 0.55), 0.7, 0.35, facecolor=C_CAP, edgecolor="k", zorder=4))
        _vapor_arrow(ax, xc - 0.3, -2.0, 0.35)
        _vapor_arrow(ax, xc + 0.3, -2.0, 0.35)
    _liquid_layer(ax, 0.4, 8.2, 1.6, weir_x=8.0, weir_h=2.0)
    ax.text(4.2, 3.3, "valves lift with vapor rate,\nclose at low load", ha="center", fontsize=7)


def panel_fixed_valve(ax):
    _panel_frame(ax, "4. Fixed Valve\nf_hole=0.12, fixed venturi louvers")
    _downcomer(ax, 8.2, 9.6, "downcomer")
    ax.add_patch(Rectangle((0.4, -0.3), 7.8, 0.3, facecolor=C_DECK, edgecolor="k", zorder=2))
    for xc in np.linspace(1.2, 7.4, 5):
        verts = [(xc - 0.4, 0), (xc + 0.4, 0), (xc + 0.15, 0.5), (xc - 0.15, 0.5)]
        ax.add_patch(Polygon(verts, closed=True, facecolor=C_CARTRIDGE, edgecolor="k", zorder=3))
        _vapor_arrow(ax, xc - 0.35, -2.0, 0.35)
        _vapor_arrow(ax, xc + 0.35, -2.0, 0.35)
    _liquid_layer(ax, 0.4, 8.2, 1.6, weir_x=8.0, weir_h=2.0)
    ax.text(4.2, 3.3, "fixed shaped orifices,\nno moving parts", ha="center", fontsize=7)


def panel_high_perf(ax):
    _panel_frame(ax, "5. High-Performance MD\nf_hole=0.14, multi-downcomer")
    _downcomer(ax, 0.4, 1.4, "DC")
    _downcomer(ax, 4.6, 5.4, "DC")
    _downcomer(ax, 8.6, 9.6, "DC")
    ax.add_patch(Rectangle((1.4, -0.3), 3.2, 0.3, facecolor=C_DECK, edgecolor="k", zorder=2))
    ax.add_patch(Rectangle((5.4, -0.3), 3.2, 0.3, facecolor=C_DECK, edgecolor="k", zorder=2))
    for xc in [2.0, 2.8, 3.6, 4.0, 6.0, 6.8, 7.6, 8.0]:
        ax.add_patch(Circle((xc, 0), 0.16, facecolor="white", edgecolor="k", zorder=3))
        _vapor_arrow(ax, xc, -2.0, 1.0)
    _liquid_layer(ax, 1.4, 4.6, 0.9, weir_x=4.5, weir_h=1.0)
    _liquid_layer(ax, 5.4, 8.6, 0.9, weir_x=8.5, weir_h=1.0)
    ax.text(5, 3.3, "short liquid path per pass,\nlower weir, more capacity", ha="center", fontsize=7)


def panel_ripple(ax):
    _panel_frame(ax, "6. Ripple (Corrugated)\nf_hole=0.11, corrugated deck")
    _downcomer(ax, 8.2, 9.6, "downcomer")
    xs = np.linspace(0.4, 8.2, 100)
    ys = -0.15 + 0.18 * np.sin(xs * 2.4)
    ax.plot(xs, ys, color="k", lw=1.5, zorder=2)
    ax.fill_between(xs, ys - 0.25, ys, color=C_DECK, zorder=1)
    for xc in np.linspace(1.2, 7.4, 7):
        yc = -0.15 + 0.18 * np.sin(xc * 2.4)
        ax.add_patch(Circle((xc, yc), 0.16, facecolor="white", edgecolor="k", zorder=3))
        _vapor_arrow(ax, xc, -2.2, yc + 1.5)
    _liquid_layer(ax, 0.4, 8.2, 1.5, weir_x=8.0, weir_h=2.0)
    ax.text(4.2, 3.3, "corrugated deck: extra\ninterfacial area, low-rate\nliquid retention in troughs",
            ha="center", fontsize=7)


def panel_apex(ax):
    _panel_frame(ax, "7. HyTrays Apex\nf_hole=0.25, cartridge-grid")
    _downcomer(ax, 0.0, 0.7, "in")
    _downcomer(ax, 9.3, 10.0, "out")
    ax.add_patch(Rectangle((0.7, -0.3), 8.6, 0.3, facecolor=C_DECK, edgecolor="k", zorder=2))
    # primary cartridges (70%)
    pri_x = np.linspace(1.2, 6.6, 4)
    for xc in pri_x:
        ax.add_patch(Rectangle((xc - 0.35, 0), 0.7, 1.4, facecolor=C_CARTRIDGE,
                                edgecolor="k", zorder=3))
        ax.add_patch(Ellipse((xc, 1.9), 1.1, 0.55, facecolor=C_CAP, edgecolor="k", zorder=4))
        _vapor_arrow(ax, xc, -2.2, 1.6)
        for sgn in (-1, 1):
            ax.add_patch(FancyArrow(xc, 1.9, sgn * 0.6, 0, width=0.03,
                                     head_width=0.2, head_length=0.2,
                                     length_includes_head=True, color=C_VAPOR,
                                     zorder=5, alpha=0.8))
    # secondary lip-seal (15%)
    for xc in (7.6, 8.4):
        verts = [(xc - 0.35, 0), (xc + 0.35, 0), (xc + 0.2, 0.6), (xc - 0.25, 0.45)]
        ax.add_patch(Polygon(verts, closed=True, facecolor=C_LIP, edgecolor="k", zorder=4))
        _vapor_arrow(ax, xc, -2.2, 0.3)
    # weir + low h_weir
    ax.add_patch(Rectangle((9.05, 0), 0.15, 0.5, facecolor="#444444", zorder=3))
    _liquid_layer(ax, 0.7, 9.3, 0.9)
    # AEGIS grid below
    for xc in (2.0, 5.0, 8.0):
        ax.add_patch(Rectangle((xc - 0.3, -3.0), 0.6, 2.7, facecolor=C_AEGIS,
                                edgecolor="k", alpha=0.6, zorder=1))
    ax.text(5, 3.3, "70% swirl-tube cartridges +\n15% lip-seal trickle + AEGIS grid;\nh_weir=0.5 in, C0=0.85",
            ha="center", fontsize=7)


def make_family_comparison(out_path: str) -> None:
    fig, axes = plt.subplots(2, 4, figsize=(15, 7.5))
    panels = [panel_sieve, panel_dualflow, panel_moving_valve, panel_fixed_valve,
              panel_high_perf, panel_ripple, panel_apex]
    for ax, fn in zip(axes.flat, panels):
        fn(ax)
    axes.flat[-1].axis("off")
    axes.flat[-1].text(0.5, 0.5,
                        "Vapor path  ->  red arrows\n"
                        "Liquid / froth  ->  light blue (hatched)\n"
                        "Moving parts  ->  navy\n"
                        "Fixed/no-moving-part elements -> blue-grey\n"
                        "Apex-only: cartridge caps, lip-seal,\n"
                        "AEGIS grid (green), PULSAR slots (purple)\n\n"
                        "Schematic cross-sections, not to a common\n"
                        "scale -- proportions follow Table 1 of\n"
                        "apex_paper_tables.md (f_hole, h_weir, etc.)",
                        ha="center", va="center", fontsize=9,
                        transform=axes.flat[-1].transAxes,
                        bbox=dict(boxstyle="round", facecolor="#f5f5f5"))
    fig.suptitle("Tray-family cross-section comparison: Apex vs. conventional decks",
                  fontsize=13, fontweight="bold")
    fig.tight_layout(rect=(0, 0, 1, 0.96))
    fig.savefig(out_path, dpi=140)
    plt.close(fig)


def main() -> None:
    os.makedirs(OUT_DIR, exist_ok=True)

    fig, ax = plt.subplots(figsize=(11, 7))
    draw_cross_section(ax)
    fig.tight_layout()
    fig.savefig(os.path.join(OUT_DIR, "30_apex_cross_section.png"), dpi=140)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(8, 8))
    draw_plan_view(ax)
    fig.tight_layout()
    fig.savefig(os.path.join(OUT_DIR, "31_apex_plan_view.png"), dpi=140)
    plt.close(fig)

    make_family_comparison(os.path.join(OUT_DIR, "32_tray_family_cross_sections.png"))

    print(f"Wrote drawings to {OUT_DIR}")


if __name__ == "__main__":
    main()
