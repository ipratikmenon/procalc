#!/usr/bin/env python3
"""CIST drawings v4 — curved helical guide vane (variable-pitch helix).

Generates output/46_cist_curved_vane.png:
  (a) 3-D chamber cross-section with curved helix vane
  (b) Helix angle θ(z) profile: 20° → 70°
  (c) Centripetal-acceleration profile Ng(z) in units of g
  (d) Cut-diameter d₅₀(z) along chamber height
"""
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec
from mpl_toolkits.mplot3d import Axes3D            # noqa: F401
from mpl_toolkits.mplot3d.art3d import Line3DCollection

# ── palette ──────────────────────────────────────────────────────────────────
C_CHAMBER = "#4a90d9"
C_CROWN   = "#1a5276"
C_DECK    = "#b0bec5"
C_JET     = "#0277bd"
C_VAPOR   = "#e53935"
C_SLOT    = "#f57c00"
C_ELLIPSE = "#00b0ff"
C_VANE    = "#2e7d32"      # dark green — helical vane
C_VANE_LT = "#a5d6a7"     # light green — vane shading
BG        = "#f5f7fa"

# ── representative geometry (mm) ─────────────────────────────────────────────
Rc_mm = 37.5          # chamber radius, mm
Hc_mm = 210.0         # chamber height, mm
Rv_mm = 0.92 * Rc_mm  # vane radius (92% of wall), mm

# Helix angles
THETA0_DEG = 20.0   # bottom: mostly tangential
THETA1_DEG = 70.0   # top:    mostly axial
theta0 = np.radians(THETA0_DEG)
theta1 = np.radians(THETA1_DEG)
beta   = (theta1 - theta0) / Hc_mm   # rad/mm

# Physical properties (representative light hydrocarbon)
u_x   = 3.0           # m/s axial vapor velocity
mu_V  = 1.5e-5        # Pa·s vapor viscosity
rho_diff = 450.0      # ρ_L - ρ_V, kg/m³
D_L   = 2.0e-9        # m²/s liquid diffusivity

# SI versions
Rc = Rc_mm * 1e-3     # m
Hc = Hc_mm * 1e-3     # m
Rv = Rv_mm * 1e-3     # m
g  = 9.81             # m/s²

# ── Helix centrelinearray (in mm for plotting) ───────────────────────────────
N_pts = 1200
z_v = np.linspace(0.0, Hc_mm, N_pts)
theta_z = theta0 + beta * z_v   # rad, varies along z

# Azimuthal angle: phi(z) = (1/(Rv_mm*beta)) * ln(sin(theta(z))/sin(theta0))
phi_z   = (1.0 / (Rv_mm * beta)) * np.log(np.sin(theta_z) / np.sin(theta0))
x_v     = Rv_mm * np.cos(phi_z)
y_v     = Rv_mm * np.sin(phi_z)

# ── Performance arrays (SI) ──────────────────────────────────────────────────
z_si = z_v * 1e-3                            # m
theta_si = theta_z.copy()                    # same array

# Tangential velocity imposed by vane
u_theta_z = u_x / np.tan(theta_si)          # m/s   (= u_x * cot(θ))

# Centripetal acceleration at wall
a_c_z = u_theta_z**2 / Rc                   # m/s²
Ng_z  = a_c_z / g                           # dimensionless (g's)

# Cut diameter along height: d50(z) from Stokes criterion at each z
# d50(z) = sqrt(72 mu_V u_x Rc² / (rho_diff * a_c(z) * Hc))
# This is the local cut diameter — a droplet centered at height z, radial
# position r=Rc/2, must cross to the wall within the remaining height (Hc-z).
# Simplified as d50(z) ≈ sqrt(18 mu_V u_x Rc / (rho_diff * a_c(z) * (Hc-z+0.01)))
# using remaining height as the capture length:
dz_remain = Hc - z_si + 0.001              # avoid div-by-zero at top
d50_z = np.sqrt(18 * mu_V * u_x * Rc / (rho_diff * a_c_z * dz_remain)) * 1e6  # µm

# Effective (integrated) Ng over the full chamber:
#   ∫₀^Hc a_c(z) dz = (u_x/Rc)² × (Hc/(θ₁-θ₀)) × [cot(θ₀)-cot(θ₁)-(θ₁-θ₀)]
I_cot2 = (Hc / (theta1-theta0)) * (1/np.tan(theta0) - 1/np.tan(theta1) - (theta1-theta0))
a_c_eff = (u_x**2 / Rc) * I_cot2 / Hc    # effective mean a_c (m/s²)
Ng_eff  = a_c_eff / g                      # effective g's

# Overall cut diameter (integrated):
#   d50_overall = sqrt(72 mu_V u_x Rc² / (rho_diff * ∫₀^Hc a_c dz))
int_ac = (u_x**2 / Rc) * I_cot2           # ∫a_c dz  [m²/s²]
d50_eff = np.sqrt(72 * mu_V * u_x * Rc**2 / (rho_diff * int_ac)) * 1e6  # µm

# Arc length of curved helix (closed form)
# L = (Hc/(θ₁-θ₀)) × ln[(csc θ₀ + cot θ₀)/(csc θ₁ + cot θ₁)]
csc0, cot0 = 1/np.sin(theta0), 1/np.tan(theta0)
csc1, cot1 = 1/np.sin(theta1), 1/np.tan(theta1)
L_vane_m = (Hc / (theta1-theta0)) * np.log((csc0+cot0)/(csc1+cot1))
L_vane_mm = L_vane_m * 1e3

# Total azimuthal sweep
phi_total = phi_z[-1]  # radians
turns = phi_total / (2*np.pi)

print(f"Curved helix vane geometry:")
print(f"  θ₀ = {THETA0_DEG}°, θ₁ = {THETA1_DEG}°")
print(f"  Total turns   = {turns:.3f}")
print(f"  Arc length    = {L_vane_mm:.1f} mm")
print(f"  φ_total       = {np.degrees(phi_total):.1f}°")
print(f"  Ng at base    = {Ng_z[0]:.1f}g")
print(f"  Ng at crown   = {Ng_z[-1]:.2f}g")
print(f"  Ng,eff        = {Ng_eff:.1f}g")
print(f"  d50,overall   = {d50_eff:.2f} µm")

# ════════════════════════════════════════════════════════════════════════════
#  FIGURE 46 — 2×2 layout
# ════════════════════════════════════════════════════════════════════════════

fig46 = plt.figure(figsize=(15, 12))
fig46.patch.set_facecolor(BG)
gs = gridspec.GridSpec(2, 2, figure=fig46, hspace=0.38, wspace=0.30,
                        left=0.07, right=0.96, top=0.91, bottom=0.06)

# ── Panel A: 3-D chamber with curved helical vane ───────────────────────────
ax_A = fig46.add_subplot(gs[0, 0], projection="3d")

# Cylinder wall (transparent)
th_cy = np.linspace(0, 2*np.pi, 80)
T_cy, Z_cy = np.meshgrid(th_cy, np.array([0.0, Hc_mm]))
ax_A.plot_surface(Rc_mm*np.cos(T_cy), Rc_mm*np.sin(T_cy), Z_cy,
                  color=C_CHAMBER, alpha=0.12, linewidth=0)

# Cylinder wire edges (top and bottom rings)
ax_A.plot(Rc_mm*np.cos(th_cy), Rc_mm*np.sin(th_cy), np.zeros_like(th_cy),
          color=C_CHAMBER, linewidth=1.2, alpha=0.60)
ax_A.plot(Rc_mm*np.cos(th_cy), Rc_mm*np.sin(th_cy), Hc_mm*np.ones_like(th_cy),
          color=C_CHAMBER, linewidth=1.2, alpha=0.60)

# Crown disc
Rd_arr = np.linspace(0, Rc_mm, 10)
T_dk, R_dk = np.meshgrid(th_cy, Rd_arr)
ax_A.plot_surface(R_dk*np.cos(T_dk), R_dk*np.sin(T_dk),
                  np.full_like(R_dk, Hc_mm),
                  color=C_CROWN, alpha=0.55, linewidth=0)

# Tray deck
Rd2 = np.linspace(0, Rc_mm*1.4, 6)
T_dk2, R_dk2 = np.meshgrid(th_cy, Rd2)
ax_A.plot_surface(R_dk2*np.cos(T_dk2), R_dk2*np.sin(T_dk2),
                  np.full_like(R_dk2, -12.0),
                  color=C_DECK, alpha=0.35, linewidth=0)

# Tangential slot (orange arc near base)
sl_th = np.linspace(-np.pi/6, np.pi/6, 20)
sl_z  = np.array([8, 30])
T_sl, Z_sl = np.meshgrid(sl_th, sl_z)
ax_A.plot_surface((Rc_mm+2)*np.cos(T_sl), (Rc_mm+2)*np.sin(T_sl), Z_sl,
                  color=C_SLOT, alpha=0.90, linewidth=0)

# ── CURVED HELICAL VANE — the key new feature ──────────────────────────────
# Plot the vane as a thick colored tube along the helix centrelinearray
# Colour-code by Ng(z): low = light green, high = dark green
Ng_norm = (Ng_z - Ng_z.min()) / (Ng_z.max() - Ng_z.min())

# Draw vane centerline as coloured segments
from matplotlib.colors import LinearSegmentedColormap
cmap_vane = LinearSegmentedColormap.from_list("vane", ["#a5d6a7", "#1b5e20"], N=256)

N_seg = 100
idx_seg = np.linspace(0, N_pts-2, N_seg, dtype=int)
for i in idx_seg:
    col = cmap_vane(Ng_norm[i])
    ax_A.plot(x_v[i:i+2], y_v[i:i+2], z_v[i:i+2],
              color=col, linewidth=4.0, solid_capstyle="round", zorder=5)

# Thick boundary lines for the vane (show vane extent, e=5mm height)
e_vane = 4.0   # mm (vane fin height)
r_inner = (Rv_mm - e_vane)
x_vi = r_inner * np.cos(phi_z)
y_vi = r_inner * np.sin(phi_z)
ax_A.plot(x_vi[::8], y_vi[::8], z_v[::8],
          color=C_VANE, linewidth=1.5, alpha=0.65, linestyle="--")

# Liquid jet (cyan arrow from sump)
ax_A.quiver(0, 0, -35, 0, 0, 42, color=C_JET,
            arrow_length_ratio=0.28, linewidth=2.5)

# Elliptical nozzle at base
th_e = np.linspace(0, 2*np.pi, 100)
ax_A.plot(10.5*np.cos(th_e), 5.2*np.sin(th_e), np.zeros(100),
          color=C_ELLIPSE, linewidth=3.0, zorder=8)

# Vapor exit
ax_A.quiver(0, 0, Hc_mm, 0, 0, 42, color=C_VAPOR,
            arrow_length_ratio=0.28, linewidth=2.0, linestyle="--")

ax_A.set_xlabel("x (mm)", fontsize=7, labelpad=1)
ax_A.set_ylabel("y (mm)", fontsize=7, labelpad=1)
ax_A.set_zlabel("z (mm)", fontsize=7, labelpad=1)
lim = Rc_mm * 1.6
ax_A.set_xlim(-lim, lim); ax_A.set_ylim(-lim, lim)
ax_A.set_zlim(-40, Hc_mm + 50)
ax_A.view_init(elev=22, azim=38)
try:
    ax_A.set_box_aspect([1, 1, Hc_mm/Rc_mm*0.80])
except AttributeError:
    pass
ax_A.set_title("(a)  Curved helical guide vane in swirl chamber\n"
               "Colour: green-dark = high $N_g$, light = low $N_g$",
               fontsize=9, fontweight="bold")

# Vane legend arrow
sm = plt.cm.ScalarMappable(cmap=cmap_vane,
                            norm=plt.Normalize(vmin=Ng_z.min(), vmax=Ng_z.max()))
sm.set_array([])
cbar = fig46.colorbar(sm, ax=ax_A, shrink=0.40, pad=0.04, aspect=18)
cbar.set_label("$N_g(z)$ [g]", fontsize=7.5)
cbar.ax.tick_params(labelsize=7)


# ── Panel B: Helix angle θ(z) profile ───────────────────────────────────────
ax_B = fig46.add_subplot(gs[0, 1])

z_plot = z_v   # mm
theta_deg = np.degrees(theta_z)
ax_B.plot(z_plot, theta_deg, color=C_VANE, linewidth=2.8)
ax_B.fill_between(z_plot, THETA0_DEG, theta_deg, color=C_VANE_LT, alpha=0.50)

# Annotations
ax_B.axhline(THETA0_DEG, color="#555", linewidth=1.0, linestyle="--")
ax_B.axhline(THETA1_DEG, color="#555", linewidth=1.0, linestyle="--")
ax_B.text(5, THETA0_DEG + 1.5, f"θ₀ = {THETA0_DEG:.0f}° (tangential-heavy)",
          fontsize=8.5, color="#555")
ax_B.text(5, THETA1_DEG + 1.5, f"θ₁ = {THETA1_DEG:.0f}° (axial-heavy)",
          fontsize=8.5, color="#555")

# Zone annotations
ax_B.fill_betweenx([18, 38], 0, Hc_mm*0.35,
                   color="#e8f5e9", alpha=0.50)
ax_B.text(Hc_mm*0.17, 22, "High-swirl\nzone", ha="center",
          fontsize=8, color="#2e7d32", fontweight="bold")
ax_B.fill_betweenx([55, 73], Hc_mm*0.65, Hc_mm,
                   color="#fce4ec", alpha=0.50)
ax_B.text(Hc_mm*0.83, 58, "Axial-flow\nzone", ha="center",
          fontsize=8, color="#b71c1c", fontweight="bold")

ax_B.set_xlabel("Height z (mm)", fontsize=9)
ax_B.set_ylabel("Helix angle θ(z) (degrees)", fontsize=9)
ax_B.set_xlim(0, Hc_mm)
ax_B.set_ylim(10, 80)
ax_B.grid(True, alpha=0.22, linestyle="--")
ax_B.set_title(
    r"(b)  Helix angle profile  $\theta(z) = \theta_0 + (\theta_1-\theta_0)\,z/H_c$",
    fontsize=9, fontweight="bold")


# ── Panel C: Centripetal acceleration Ng(z) ─────────────────────────────────
ax_C = fig46.add_subplot(gs[1, 0])

ax_C.semilogy(z_plot, Ng_z, color="#1565c0", linewidth=2.8)
ax_C.fill_between(z_plot, 1, Ng_z, color="#e3f2fd", alpha=0.55)

# Reference lines
ax_C.axhline(Ng_eff, color="#e65100", linewidth=1.6, linestyle="-.",
             label=f"$N_{{g,\\mathrm{{eff}}}}$ = {Ng_eff:.1f} g (integrated mean)")
ax_C.axhline(30.0, color="#888", linewidth=1.2, linestyle="--",
             label="Baseline design target: 30–40 g")
ax_C.axhline(40.0, color="#888", linewidth=1.2, linestyle="--")

# Zone shading
ax_C.axvspan(0, Hc_mm*0.35, color="#e8f5e9", alpha=0.35, label="High-swirl zone")
ax_C.axvspan(Hc_mm*0.65, Hc_mm, color="#fce4ec", alpha=0.35, label="Axial-flow zone")

ax_C.text(8, Ng_z[0]*0.55, f"$N_g$ = {Ng_z[0]:.0f} g at base",
          fontsize=8.5, color="#0d47a1", fontweight="bold")
ax_C.text(Hc_mm*0.60, Ng_z[-1]*1.8,
          f"$N_g$ = {Ng_z[-1]:.1f} g at crown", fontsize=8.5, color="#0d47a1")

ax_C.set_xlabel("Height z (mm)", fontsize=9)
ax_C.set_ylabel("Centripetal acceleration  $N_g(z) = u_x^2 \\cot^2\\!\\theta(z)\\,/\\,(R_c g)$  [g]",
                fontsize=8.5)
ax_C.set_xlim(0, Hc_mm)
ax_C.set_ylim(1, Ng_z[0]*2.5)
ax_C.legend(fontsize=7.5, loc="upper right")
ax_C.grid(True, alpha=0.22, linestyle="--")
ax_C.set_title(f"(c)  Centripetal acceleration profile  ({Ng_z[-1]:.1f} g → {Ng_z[0]:.0f} g,  base → crown)",
               fontsize=9, fontweight="bold")


# ── Panel D: Cut diameter d₅₀(z) ─────────────────────────────────────────────
ax_D = fig46.add_subplot(gs[1, 1])

# Clip d50 to reasonable range (very bottom is unrealistically small)
d50_plot = np.clip(d50_z, 0.5, 200)
ax_D.semilogy(z_plot[:-5], d50_plot[:-5], color="#6a1b9a", linewidth=2.8)
ax_D.fill_between(z_plot[:-5], 0.1, d50_plot[:-5], color="#f3e5f5", alpha=0.55)

# Reference lines
ax_D.axhline(d50_eff, color="#e65100", linewidth=1.8, linestyle="-.",
             label=f"Overall $d_{{50}}$ = {d50_eff:.1f} µm (integrated)")
ax_D.axhline(150, color="#555", linewidth=1.2, linestyle="--",
             label="$d_{\\min}$ ≈ 150 µm (axis-switching)")

ax_D.text(8, d50_plot[0]*0.45, f"$d_{{50}}$ = {d50_plot[0]:.1f} µm at base",
          fontsize=8.5, color="#4a148c", fontweight="bold")

# Zones
ax_D.axvspan(0, Hc_mm*0.35, color="#e8f5e9", alpha=0.35)
ax_D.axvspan(Hc_mm*0.65, Hc_mm, color="#fce4ec", alpha=0.35)

ax_D.set_xlabel("Height z (mm)", fontsize=9)
ax_D.set_ylabel("Local cut diameter  $d_{50}(z)$  (µm)", fontsize=9)
ax_D.set_xlim(0, Hc_mm * 0.95)
ax_D.set_ylim(0.5, 250)
ax_D.legend(fontsize=7.5, loc="upper left")
ax_D.grid(True, alpha=0.22, linestyle="--")
ax_D.set_title(f"(d)  Local cut-diameter profile  (overall $d_{{50}}$ = {d50_eff:.1f} µm)",
               fontsize=9, fontweight="bold")


fig46.suptitle(
    "CIST Curved Helical Guide Vane  |  θ(z) linear: 20° → 70° (tangential → axial)\n"
    f"Total turns = {turns:.2f}  ·  Arc length = {L_vane_mm:.0f} mm  "
    f"·  $N_{{g,\\mathrm{{eff}}}}$ = {Ng_eff:.1f} g  ·  Overall $d_{{50}}$ = {d50_eff:.1f} µm",
    fontsize=11, fontweight="bold", color="#1a252f"
)

fig46.savefig("output/46_cist_curved_vane.png", dpi=150, bbox_inches="tight", facecolor=BG)
plt.close(fig46)
print("Saved  output/46_cist_curved_vane.png")
print("Done.")
