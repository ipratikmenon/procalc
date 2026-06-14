#!/usr/bin/env python3
"""HyTrays — distillation-column tray hydraulics engine.

Implements the standard Fair (Souders-Brown) capacity/flooding correlation
and the classic tray-rating hand calculations (dry-tray pressure drop,
Francis-weir clear-liquid height, downcomer backup, turndown/weep margin
and an O'Connell-based efficiency estimate) so that the five tray families
in ``tray_library.py`` can be sized and compared on a common basis.

All internal units are FPS (the Procalc convention -- see
``common/units.py``):

    mass flow      lb/hr
    density        lb/ft3
    surface tension dyn/cm (= mN/m)
    viscosity       cP
    length          ft (diameters, areas) / in (weir, pressure-drop heads)
    velocity        ft/s
    pressure drop   psi (per tray) and inches of liquid (heads)

References
----------
* Fair, J.R., "How to Predict Sieve Tray Entrainment and Flooding,"
  Petro/Chem Engineer, 1961 -- capacity (C_SBF) correlation.
* The closed-form fit to Fair's C_SBF chart used below,
  ``C_SBF = 0.0105 + 8.127e-4 * HS_mm**0.755 * exp(-1.463 * FLV**0.842)``
  (HS = tray spacing, FLV = flow parameter), is the curve fit widely used
  in process simulators and textbooks (e.g. Kister, *Distillation Design*).
* Dry-tray orifice equation, Francis weir formula and downcomer apron loss:
  Kister, *Distillation Design* (1992), Ch. 6.
* O'Connell, H.E., "Plate Efficiency of Fractionating Columns and
  Absorbers," Trans. AIChE 42 (1946) -- overall column efficiency
  correlation E_o (%) = 51 - 32.5 log10(alpha * mu_L).
"""
from __future__ import annotations

import math
from dataclasses import dataclass

from tray_library import TrayType

# ── physical constants ──────────────────────────────────────────────────
GPM_PER_FT3HR_AT_RHO = 0.124675   # gpm = (lb/hr / rho[lb/ft3]) * this factor
SIGMA_REF = 20.0                  # dyn/cm reference for Fair surface-tension correction


@dataclass(frozen=True)
class ColumnCase:
    """Operating point + physical properties for one column section.

    Vapor and liquid loads are the internal traffic *through the tray
    being rated* (not feed/product draws).
    """

    name: str
    v_mass_lb_hr: float     # vapor mass flow, lb/hr
    l_mass_lb_hr: float     # liquid (reflux) mass flow, lb/hr
    rho_v: float            # vapor density, lb/ft3
    rho_l: float            # liquid density, lb/ft3
    sigma: float            # liquid surface tension, dyn/cm
    mu_l: float             # liquid viscosity, cP
    alpha: float            # light/heavy-key relative volatility (O'Connell)
    tray_spacing_in: float = 24.0
    f_flood: float = 0.80   # design fraction of flooding velocity
    n_theoretical: int = 20  # theoretical stages for the actual-tray comparison

    # ── service modifiers ───────────────────────────────────────────────
    system_factor: float = 1.00  # foaming/system derate on Fair's u_nf (Kister Fp);
                                  # 1.0 = non-foaming, ~0.75 = moderate foam, ~0.5-0.6 = severe foam
    fouling: bool = False         # apply tray.fouling_open_area_retention to the hole area
    description: str = ""         # one-line service description, used in reports


@dataclass(frozen=True)
class TrayDesignResult:
    tray: TrayType
    case: ColumnCase

    flv: float              # flow parameter (L/V)*sqrt(rho_v/rho_l)
    csbf_ft_s: float        # Fair capacity parameter, ft/s
    u_nf_ft_s: float        # flooding (superficial, active-area) velocity, ft/s
    u_design_ft_s: float    # design vapor velocity, ft/s

    a_active_ft2: float
    a_col_ft2: float
    diameter_ft: float

    f_hole_eff: float        # effective hole-area fraction used (fouling-derated if case.fouling)
    u_hole_ft_s: float
    dp_dry_in: float        # dry-tray pressure drop, in. liquid
    h_ow_in: float          # weir crest (Francis), in. liquid
    h_clear_in: float       # clear liquid height = h_weir + h_ow, in.
    dp_tray_in: float        # total wet tray pressure drop, in. liquid
    dp_tray_psi: float

    h_dc_in: float | None       # downcomer clear-liquid backup, in.
    h_dc_limit_in: float | None  # 50% backup limit, in.
    dc_backup_pct: float | None  # h_dc / h_dc_limit * 100

    pct_flood_design: float   # = f_flood * 100, by construction
    max_load_pct: float       # max throughput before flooding, % of design
    min_load_pct: float       # min stable throughput (weep), % of design
    turndown_ratio: float     # max_load_pct / min_load_pct

    e_oconnell_pct: float
    e_tray_pct: float
    n_actual_trays: int
    column_height_ft: float
    total_dp_psi: float


def flow_parameter(case: ColumnCase) -> float:
    """FLV = (L/V) * sqrt(rho_v / rho_l) -- dimensionless flow parameter."""
    return (case.l_mass_lb_hr / case.v_mass_lb_hr) * math.sqrt(case.rho_v / case.rho_l)


def csbf_fair(flv: float, tray_spacing_in: float) -> float:
    """Fair capacity parameter C_SBF (ft/s) from the FLV / tray-spacing fit.

    ``C_SBF[m/s] = 0.0105 + 8.127e-4 * HS[mm]**0.755 * exp(-1.463 * FLV**0.842)``
    valid for HS = 150-600 mm (6-24 in) and FLV = 0.01-1.0; converted to ft/s.
    """
    hs_mm = tray_spacing_in * 25.4
    csbf_m_s = 0.0105 + 8.127e-4 * hs_mm ** 0.755 * math.exp(-1.463 * flv ** 0.842)
    return csbf_m_s * 3.28084


def flooding_velocity(csbf_ft_s: float, sigma: float, rho_v: float, rho_l: float) -> float:
    """Souders-Brown flooding velocity (ft/s), surface-tension corrected."""
    return csbf_ft_s * (sigma / SIGMA_REF) ** 0.2 * math.sqrt((rho_l - rho_v) / rho_v)


def oconnell_efficiency_pct(alpha: float, mu_l: float) -> float:
    """O'Connell (1946) overall column efficiency, % (clipped to 100%)."""
    e = 51.0 - 32.5 * math.log10(alpha * mu_l)
    return max(1.0, min(e, 100.0))


def design_tray(tray: TrayType, case: ColumnCase) -> TrayDesignResult:
    """Size a column for ``tray`` at ``case.f_flood`` and rate it.

    Returns a :class:`TrayDesignResult` with the sizing (diameter), per-tray
    pressure drop / downcomer backup, turndown range, efficiency and the
    resulting actual tray count / column height / total pressure drop for
    ``case.n_theoretical`` theoretical stages.
    """
    flv = flow_parameter(case)
    csbf = csbf_fair(flv, case.tray_spacing_in)
    u_nf = (flooding_velocity(csbf, case.sigma, case.rho_v, case.rho_l)
            * tray.capacity_factor * case.system_factor)
    u_design = case.f_flood * u_nf

    v_dot_cfs = case.v_mass_lb_hr / 3600.0 / case.rho_v
    a_active = v_dot_cfs / u_design
    a_col = a_active / tray.f_active
    diameter_ft = math.sqrt(4.0 * a_col / math.pi)

    # ── dry-tray pressure drop (orifice equation) ──────────────────────
    f_hole_eff = tray.f_hole * (tray.fouling_open_area_retention if case.fouling else 1.0)
    a_hole = f_hole_eff * a_active
    u_hole = v_dot_cfs / a_hole
    dp_dry = 0.186 * (u_hole / tray.c0) ** 2 * (case.rho_v / case.rho_l)
    dp_dry = max(dp_dry, tray.dp_dry_floor_in)

    # ── clear liquid height (Francis weir) ─────────────────────────────
    q_l_gpm = case.l_mass_lb_hr / case.rho_l * GPM_PER_FT3HR_AT_RHO
    if tray.has_downcomer:
        weir_len_in = 0.73 * diameter_ft * 12.0
        h_ow = 0.48 * (q_l_gpm / weir_len_in) ** (2.0 / 3.0)
        h_clear = tray.h_weir_in + h_ow
    else:
        # Dualflow: no overflow weir -- liquid level set by the perforations
        # themselves. Use a representative operating liquid height.
        h_ow = 0.0
        h_clear = 1.0

    # ── total wet tray pressure drop ───────────────────────────────────
    dp_tray_in = dp_dry + tray.aeration_factor * h_clear
    dp_tray_psi = dp_tray_in / 12.0 * case.rho_l / 144.0

    # ── downcomer backup ────────────────────────────────────────────────
    if tray.has_downcomer:
        a_dc = (a_col - a_active) / 2.0
        h_da = 0.03 * (q_l_gpm / (100.0 * a_dc)) ** 2
        h_dc = dp_dry + h_clear + h_da
        h_dc_limit = 0.5 * (case.tray_spacing_in + tray.h_weir_in)
        dc_backup_pct = h_dc / h_dc_limit * 100.0
    else:
        h_dc = None
        h_dc_limit = None
        dc_backup_pct = None

    # ── operating range (turndown) ──────────────────────────────────────
    pct_flood_design = case.f_flood * 100.0
    max_load_pct = 100.0 / case.f_flood          # 100% flood point, % of design throughput
    min_load_pct = tray.min_load_frac * max_load_pct
    turndown_ratio = max_load_pct / min_load_pct

    # ── efficiency & actual tray count ──────────────────────────────────
    e_oc = oconnell_efficiency_pct(case.alpha, case.mu_l)
    e_tray = min(e_oc * tray.efficiency_factor, 100.0)
    n_actual = math.ceil(case.n_theoretical / (e_tray / 100.0))
    column_height_ft = n_actual * case.tray_spacing_in / 12.0 + 10.0  # +10 ft sump/disengagement
    total_dp_psi = n_actual * dp_tray_psi

    return TrayDesignResult(
        tray=tray, case=case,
        flv=flv, csbf_ft_s=csbf, u_nf_ft_s=u_nf, u_design_ft_s=u_design,
        a_active_ft2=a_active, a_col_ft2=a_col, diameter_ft=diameter_ft,
        f_hole_eff=f_hole_eff, u_hole_ft_s=u_hole, dp_dry_in=dp_dry, h_ow_in=h_ow, h_clear_in=h_clear,
        dp_tray_in=dp_tray_in, dp_tray_psi=dp_tray_psi,
        h_dc_in=h_dc, h_dc_limit_in=h_dc_limit, dc_backup_pct=dc_backup_pct,
        pct_flood_design=pct_flood_design, max_load_pct=max_load_pct,
        min_load_pct=min_load_pct, turndown_ratio=turndown_ratio,
        e_oconnell_pct=e_oc, e_tray_pct=e_tray, n_actual_trays=n_actual,
        column_height_ft=column_height_ft, total_dp_psi=total_dp_psi,
    )
