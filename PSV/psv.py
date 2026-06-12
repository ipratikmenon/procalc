#!/usr/bin/env python3
"""Procalc PSV — sizing & contingency engine (single file to run).

Implements the project PSV methodology (PSV_Calculation_Methodology.md):
contingency classification (design / remote / fire) with allowable-pressure
rules, scenario relief loads (blocked outlet, API 521 fire incl. vapor-only
screening & two-phase swell, CV failure per ISA/IEC 60534, gas blowby &
check-valve reverse flow, thermal expansion, tube rupture / 6 mm leak,
utility failure, tower LOC (HMVRAF flash), liquid overfill, boil-up two-flash
λ method, abnormal heat input with steam-CV back-calculation, runaway),
rigorous flash at relieving conditions (composition from the HMB via the
hydraulics flash engine), static-elevation-corrected set pressure
("PSV Inputs plus"), API 520 Part I sizing (vapor / liquid / two-phase Leung
omega), API 526 orifice selection, inlet 3 % & back-pressure compliance, and
a disposal-routing advisor.

Workbook units are declared on the UNITS sheet (FPS or SI + per-quantity
overrides) — shared Procalc convention (../common/units.py).  The engine
computes internally in SI.

Usage
-----
    python psv.py --template           # blank FPS input workbook
    python psv.py --template --si      # blank SI input workbook
    python psv.py psv_input.xlsx       # run (writes psv_output_<tag>.xlsx)

Companions: ../Hydraulics/hydraulics.py (HMB stream reading + flash VLE),
            ../common/units.py (units layer).
"""
from __future__ import annotations

import math
import os
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_HERE, "..", "common"))
sys.path.insert(0, os.path.join(_HERE, "..", "Hydraulics"))
import units as UN                                     # noqa: E402
from style import (                                    # noqa: E402
    NAVY, WHITE, LGRAY, AMBER, TEAL, DGRAY, ORANGE, CYAN, C, _title,
    _U, _hdr_formula, _kv_block, _result_rows,
)

try:
    import hydraulics as HYD          # HMB stream loading + flash VLE
except Exception as _exc:             # pragma: no cover
    HYD = None
    print(f"  NOTE: hydraulics engine not importable ({_exc}) — "
          "HMB stream linkage & rigorous flash disabled; manual properties only.")

# ═══════════════════════════════════════════════════════════════════════════
#  Constants
# ═══════════════════════════════════════════════════════════════════════════
R_UNIV   = 8314.462618        # J/(kmol·K)
G_STD    = 9.80665            # m/s²
ATM_PA   = 101_325.0
BTU_LB_TO_JKG = 2326.0
PSI_TO_PA = 6894.757293

# API 526 orifice designations (letter → area cm²)
API_ORIFICE = {
    "D": 0.710, "E": 1.267, "F": 1.980, "G": 3.245, "H": 5.065,
    "J": 8.303, "K": 12.65, "L": 19.61, "M": 29.03, "N": 41.16,
    "P": 71.00, "Q": 109.7, "R": 167.7, "T": 258.1,
}

# Fittings catalogue (K-table) — reuse the Hydraulics tool's list so
# INLET/OUTLET piping fittings match the hydraulics engine exactly.
if HYD is not None:
    FITTING_NAMES = HYD.FITTING_NAMES
    fitting_k = HYD.fitting_k
else:                                  # pragma: no cover — HYD unavailable
    FITTING_NAMES = ["Straight Pipeline", "Elbow 90 Short", "Gate Full Open"]

    def fitting_k(name: str | None) -> float:
        return 0.0

N_FIT_ROWS = 12                        # fitting entry rows per piping sheet

# Overpressure allowance by contingency class (% of set pressure, gauge)
OP_PCT = {"DESIGN": 10.0, "DESIGN_MULTI": 16.0, "FIRE": 21.0}

SCEN_CODES = [
    ("BO",       "Blocked Outlet"),
    ("FIRE",     "External Fire (API 521)"),
    ("CV",       "Control Valve Failure"),
    ("GASBB",    "Gas Blowby / Check-Valve Reverse Flow"),
    ("THERM",    "Thermal (Hydraulic) Expansion"),
    ("TUBE",     "HX Tube Rupture / Leak"),
    ("UTIL",     "Utility Failure / Loss of Cooling"),
    ("LOC",      "Tower Loss of Condensing (HMVRAF)"),
    ("OVERFILL", "Liquid Overfill (Tower §5.6)"),
    ("BOILUP",   "Liquid Boil-Up (Power Failure §5.7)"),
    ("ABHEAT",   "Abnormal Heat Input (§5.4)"),
    ("RUNAWAY",  "Runaway Reaction"),
    ("OTHER",    "Other / Manual"),
]

# ═══════════════════════════════════════════════════════════════════════════
#  Units I/O adapter:  user units ⇄ internal SI  (via the shared FPS layer)
# ═══════════════════════════════════════════════════════════════════════════
_FPS2SI = {
    "P":    lambda v: v * PSI_TO_PA,                       # psia → Pa(a)
    "dP":   lambda v: v * PSI_TO_PA,                       # psi  → Pa
    "T":    lambda v: (v - 32.0) / 1.8 + 273.15,           # °F   → K
    "dT":   lambda v: v / 1.8,
    "mflow": lambda v: v * 0.45359237 / 3600.0,            # lb/hr → kg/s
    "molflow": lambda v: v * 0.45359237 / 3600.0,          # lbmol/hr → kmol/s
    "rho":  lambda v: v * 16.018463,                       # lb/ft³ → kg/m³
    "L":    lambda v: v * 0.3048,                          # ft → m
    "Lin":  lambda v: v * 0.0254,                          # in → m
    "A":    lambda v: v * 0.09290304,                      # ft² → m²
    "Ain":  lambda v: v * 6.4516e-4,                       # in² → m²
    "V":    lambda v: v * 0.028316846592,                  # ft³ → m³
    "Q":    lambda v: v * 0.29307107,                      # BTU/hr → W
    "h":    lambda v: v * BTU_LB_TO_JKG,                   # BTU/lb → J/kg
    "v":    lambda v: v * 0.3048,
    "visc": lambda v: v,                                   # cP stays cP
    "MW":   lambda v: v,
    "-":    lambda v: v,
    "pct":  lambda v: v,
}
_SI2FPS = {
    "P":    lambda v: v / PSI_TO_PA,
    "dP":   lambda v: v / PSI_TO_PA,
    "T":    lambda v: (v - 273.15) * 1.8 + 32.0,
    "dT":   lambda v: v * 1.8,
    "mflow": lambda v: v * 3600.0 / 0.45359237,
    "molflow": lambda v: v * 3600.0 / 0.45359237,
    "rho":  lambda v: v / 16.018463,
    "L":    lambda v: v / 0.3048,
    "Lin":  lambda v: v / 0.0254,
    "A":    lambda v: v / 0.09290304,
    "Ain":  lambda v: v / 6.4516e-4,
    "V":    lambda v: v / 0.028316846592,
    "Q":    lambda v: v / 0.29307107,
    "h":    lambda v: v / BTU_LB_TO_JKG,
    "v":    lambda v: v / 0.3048,
    "visc": lambda v: v,
    "MW":   lambda v: v,
    "-":    lambda v: v,
    "pct":  lambda v: v,
}


class UIO:
    """User-facing units (per UNITS sheet) ⇄ internal SI."""

    def __init__(self, usys: UN.UnitSystem):
        self.u = usys

    @property
    def system(self) -> str:
        return self.u.system

    def si(self, qty: str, val):
        """user value → SI (None-safe)."""
        if val is None:
            return None
        v = self.u.to_internal(qty, val)        # user → FPS internal
        try:
            return _FPS2SI[qty](float(v))
        except (KeyError, TypeError, ValueError):
            return v

    def user(self, qty: str, si_val, nd: int | None = 4):
        """SI value → user display units."""
        if si_val is None:
            return None
        try:
            fps = _SI2FPS[qty](float(si_val))
        except (KeyError, TypeError, ValueError):
            return si_val
        out = self.u.from_internal(qty, fps)
        if nd is not None and isinstance(out, float):
            out = round(out, nd)
        return out

    def label(self, qty: str) -> str:
        return self.u.label(qty)

    def hdr(self, base: str, qty: str) -> str:
        return self.u.hdr(base, qty)


# ═══════════════════════════════════════════════════════════════════════════
#  Small helpers
# ═══════════════════════════════════════════════════════════════════════════
def num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return None if (isinstance(v, float) and math.isnan(v)) else float(v)
    s = str(v).strip()
    if not s or s.upper() in {"N/A", "NA", "NONE", "—", "-"}:
        return None
    m = re.search(r"[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?", s)
    return float(m.group(0)) if m else None


def txt(v):
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def yes(v) -> bool:
    return str(v or "").strip().upper().startswith("Y")


def bisect(f, lo, hi, it=200, tol=1e-10):
    """Robust bisection; returns None if no sign change."""
    flo, fhi = f(lo), f(hi)
    if flo == 0:
        return lo
    if fhi == 0:
        return hi
    if flo * fhi > 0:
        return None
    for _ in range(it):
        mid = 0.5 * (lo + hi)
        fm = f(mid)
        if abs(fm) < tol or (hi - lo) < tol:
            return mid
        if flo * fm <= 0:
            hi, fhi = mid, fm
        else:
            lo, flo = mid, fm
    return 0.5 * (lo + hi)


# ═══════════════════════════════════════════════════════════════════════════
#  API 520 Part I sizing (internal SI)
# ═══════════════════════════════════════════════════════════════════════════
def gas_mass_flux(p1_pa: float, p2_pa: float, t_k: float, mw: float,
                  z: float, k: float) -> tuple[float, bool]:
    """Isentropic nozzle mass flux G [kg/m²·s] and choked flag.

    Critical:    G = P1·√(k·M/(Z·R̄·T)) · (2/(k+1))^((k+1)/(2(k−1)))
    Subcritical: G = √( 2·P1·ρ1·(k/(k−1)) · (r^(2/k) − r^((k+1)/k)) )
    """
    k = max(1.001, k or 1.4)
    z = z or 1.0
    mw = mw or 28.96
    if p1_pa <= 0 or t_k <= 0:
        return 0.0, False
    r_crit = (2.0 / (k + 1.0)) ** (k / (k - 1.0))
    r = max(0.0, min(1.0, (p2_pa or 0.0) / p1_pa))
    if r <= r_crit:                       # choked
        g = p1_pa * math.sqrt(k * mw / (z * R_UNIV * t_k)) * \
            (2.0 / (k + 1.0)) ** ((k + 1.0) / (2.0 * (k - 1.0)))
        return g, True
    rho1 = p1_pa * mw / (z * R_UNIV * t_k)
    term = (k / (k - 1.0)) * (r ** (2.0 / k) - r ** ((k + 1.0) / k))
    g = math.sqrt(max(0.0, 2.0 * p1_pa * rho1 * term))
    return g, False


def size_vapor(w_kgs: float, t_k: float, mw: float, z: float, k: float,
               p1_pa: float, p2_pa: float, kd: float, kb: float,
               kc: float) -> float:
    """Required area m² — vapor/gas, critical or subcritical."""
    if w_kgs <= 0 or p1_pa <= 0:
        return 0.0
    g, _ = gas_mass_flux(p1_pa, p2_pa, t_k, mw, z, k)
    if g <= 0:
        return 0.0
    return w_kgs / (max(1e-9, kd * kb * kc) * g)


def size_liquid(w_kgs: float, rho: float, p1_pa: float, p2_pa: float,
                kd: float, kw: float, kc: float, kv: float) -> float:
    """Required area m² — liquid (API 520 Eq. 12 theoretical-orifice form)."""
    dp = p1_pa - p2_pa
    if w_kgs <= 0 or dp <= 0 or rho <= 0:
        return 0.0
    q = w_kgs / rho                                  # m³/s
    return q / (max(1e-9, kd * kw * kc * kv) * math.sqrt(2.0 * dp / rho))


def kv_viscosity(area_m2: float, rho: float, mu_cp: float,
                 w_kgs: float) -> float:
    """API 520 Kv viscosity correction (Fig.-36 fit on Reynolds number)."""
    if area_m2 <= 0 or mu_cp <= 0 or rho <= 0 or w_kgs <= 0:
        return 1.0
    d = 2.0 * math.sqrt(area_m2 / math.pi)
    v = w_kgs / (rho * area_m2)
    re = rho * v * d / (mu_cp * 1e-3)
    if re >= 1e5:
        return 1.0
    if re <= 10:
        return 0.3
    kv = 1.0 / (0.9935 + 2.878 / re ** 0.5 + 342.75 / re ** 1.5)
    return max(0.3, min(1.0, kv))


def kw_bellows(bp_pct_of_set: float) -> float:
    """Liquid back-pressure factor Kw for balanced bellows (API Fig.-31 proxy)."""
    if bp_pct_of_set <= 15.0:
        return 1.0
    return max(0.5, 1.0 - 0.011 * (bp_pct_of_set - 15.0))


def kb_factor(psv_type: str, bp_pct_of_set: float, op_pct: float,
              kb_user: float | None = None) -> float:
    """Vapor back-pressure correction Kb per PSV type."""
    pt = (psv_type or "conventional").lower()
    if kb_user:
        return kb_user
    if "pilot" in pt:
        return 1.0
    if "bellow" in pt or "balanced" in pt:
        if bp_pct_of_set <= 30.0:
            return 1.0
        return max(0.55, 1.0 - 0.012 * (bp_pct_of_set - 30.0))
    # conventional: Kb covers superimposed BP above 10% at 10% OP (API Fig.-30)
    if bp_pct_of_set <= op_pct:
        return 1.0
    return max(0.5, 1.0 - 0.014 * (bp_pct_of_set - op_pct))


def leung_omega(x0: float, rho_l: float, rho_v: float, cpl_jkgk: float,
                t0_k: float, p0_pa: float, latent_jkg: float) -> float:
    """API 520 Annex C / Leung ω parameter (saturated two-phase inlet)."""
    vl = 1.0 / max(1e-6, rho_l)
    vg = 1.0 / max(1e-6, rho_v)
    v0 = x0 * vg + (1.0 - x0) * vl
    if v0 <= 0:
        return 1.0
    omega = x0 * vg / v0
    if latent_jkg and latent_jkg > 0 and cpl_jkgk and t0_k > 0:
        vfg = vg - vl
        omega += (cpl_jkgk * t0_k * p0_pa / v0) * (vfg / latent_jkg) ** 2
    return max(1e-3, omega)


def size_twophase(w_kgs: float, p1_pa: float, p2_pa: float, x0: float,
                  rho_l: float, rho_v: float, cpl_jkgk: float, t0_k: float,
                  latent_jkg: float, kd: float, kc: float) -> tuple[float, float]:
    """Leung omega-method two-phase sizing.  Returns (area m², omega)."""
    if w_kgs <= 0 or p1_pa <= 0:
        return 0.0, 0.0
    vl = 1.0 / max(1e-6, rho_l)
    vg = 1.0 / max(1e-6, rho_v)
    v0 = x0 * vg + (1.0 - x0) * vl
    omega = leung_omega(x0, rho_l, rho_v, cpl_jkgk, t0_k, p1_pa, latent_jkg)

    def eta_resid(eta: float) -> float:
        return (eta ** 2 + (omega ** 2 - 2.0 * omega) * (1.0 - eta) ** 2
                + 2.0 * omega ** 2 * math.log(eta)
                + 2.0 * omega ** 2 * (1.0 - eta))

    eta_c = bisect(eta_resid, 1e-4, 0.9999) or 0.55
    eta_a = max(1e-6, min(0.9999, (p2_pa or ATM_PA) / p1_pa))
    if eta_a <= eta_c:                                  # choked
        g = eta_c * math.sqrt(p1_pa / (v0 * omega))
    else:
        numer = math.sqrt(max(0.0, -2.0 * (omega * math.log(eta_a)
                                           + (omega - 1.0) * (1.0 - eta_a))))
        denom = omega * (1.0 / eta_a - 1.0) + 1.0
        g = (numer / max(1e-9, denom)) * math.sqrt(p1_pa / v0)
    if g <= 0:
        return 0.0, omega
    return w_kgs / (max(1e-9, kd * kc) * g), omega


def select_orifice(req_m2: float) -> tuple[str, float]:
    req_cm2 = req_m2 * 1e4
    for letter, a in API_ORIFICE.items():
        if a >= req_cm2:
            return letter, a * 1e-4
    n = math.ceil(req_cm2 / API_ORIFICE["T"])
    return f"{n}×T", n * API_ORIFICE["T"] * 1e-4


# ═══════════════════════════════════════════════════════════════════════════
#  Fire — API 521 §5.15 + methodology §3.10
# ═══════════════════════════════════════════════════════════════════════════
def fire_env_factor(drainage: bool, firefight: bool, insulated: bool,
                    f_override: float | None) -> float:
    if f_override:
        f = f_override
    elif drainage and firefight:
        f = 0.15
    elif firefight:
        f = 0.30
    elif drainage:
        f = 0.50
    else:
        f = 1.0
    if insulated:
        f = max(0.075, min(f, 0.30))   # methodology §3.10: floor F = 0.075
    return f


def fire_q_w(wetted_m2: float, f: float) -> float:
    """Q [W] = F × 43,200 × A^0.82 (project convention, API 521 SI)."""
    if wetted_m2 <= 0:
        return 0.0
    return f * 43200.0 * wetted_m2 ** 0.82


def wetted_area_m2(orient: str, od_m: float, tt_m: float, head: str,
                   nll_m: float, grade_elev_m: float, fire_lim_m: float,
                   extra_m2: float) -> float:
    """Wetted surface to NLL, capped at the API 521 fire height limit."""
    r = od_m / 2.0
    if r <= 0:
        return max(0.0, extra_m2)
    head_one = {"2:1": 1.09 * math.pi * r * r,
                "HEMI": 2.0 * math.pi * r * r,
                "FLAT": math.pi * r * r}.get(_head_key(head),
                                             1.09 * math.pi * r * r)
    o = (orient or "Vertical").upper()
    if o.startswith("SPHERE"):
        area = 4.0 * math.pi * r * r                     # full sphere (conservative)
    elif o.startswith("HORIZ"):
        # partial wetted circumference at fill height nll (from bottom of shell)
        h = max(0.0, min(od_m, nll_m))
        theta = 2.0 * math.acos(max(-1.0, min(1.0, 1.0 - h / r)))
        arc = theta * r
        frac = h / od_m
        area = arc * tt_m + 2.0 * head_one * frac
    else:                                                # vertical
        # cap wetted height at fire limit above grade
        wet_h = max(0.0, min(nll_m, fire_lim_m - grade_elev_m))
        area = math.pi * od_m * wet_h + head_one          # bottom head + shell
    return area + max(0.0, extra_m2)


def _head_key(head: str) -> str:
    h = (head or "").upper()
    if "HEMI" in h:
        return "HEMI"
    if "FLAT" in h:
        return "FLAT"
    return "2:1"


def fire_vapor_only_ok(dia_m: float, fill_pct: float, foaming: bool,
                       dp_set_pa_g: float, orient: str) -> tuple[bool, float, str]:
    """Methodology §3.10 vapor-only screening.

    Returns (vapor_only_ok, area_multiplier, note).  area_multiplier = 2.0 for
    the vertical 1.5–3 m approximation band, else 1.0.
    """
    if dia_m < 3.0 and fill_pct < 85.0:
        return True, 1.0, "dia < 3.0 m and level < 85% expanded volume"
    if dia_m >= 3.0 and fill_pct < 95.0:
        return True, 1.0, "dia ≥ 3.0 m and level < 95% expanded volume"
    if not foaming and dp_set_pa_g >= 345_000.0:
        if (orient or "").upper().startswith("VERT") and 1.5 <= dia_m <= 3.0:
            return False, 2.0, ("vertical 1.5–3 m, non-foaming, DP ≥ 50 psig "
                                "→ two-phase area ≈ 2 × vapor-only area")
        if dia_m > 3.0 or dia_m > 1.5:
            return True, 1.0, ("non-foaming, DP ≥ 345 kPa(g), vessel category "
                               "exemption (§3.10) — verify drainage/service")
    return False, 1.0, "vapor-only criteria not met → rigorous two-phase"


def fire_twophase_w2(w_vap_kgs: float, q_w: float, cpl_jkgk: float,
                     dt_k: float, alpha: float, vf: float, vg: float,
                     latent_jkg: float) -> tuple[float, float, str]:
    """Two-phase swell estimate (§3.10 rigorous step, homogeneous-vessel form).

    x_rel = vessel-average quality at void fraction α (homogeneous):
            x = (α/vg) / (α/vg + (1−α)/vf)
    Energy: W₂ = Q / (Cp·ΔT + x_rel·λ);  take W₂ only if > vapor-only W.
    Returns (w2_kgs, x_rel, note).  This is a documented approximation —
    review against the corporate energy/swell balance for final design.
    """
    if q_w <= 0 or latent_jkg <= 0 or vg <= 0 or vf <= 0:
        return 0.0, 1.0, "insufficient data for two-phase swell — vapor-only used"
    a = max(1e-4, min(0.999, alpha))
    x_rel = (a / vg) / (a / vg + (1.0 - a) / vf)
    denom = max(1e-3, (cpl_jkgk or 0.0) * max(0.0, dt_k) + x_rel * latent_jkg)
    w2 = q_w / denom
    note = (f"homogeneous swell: x_rel={x_rel:.4f}, α={a:.3f}")
    if w2 <= w_vap_kgs:
        return w_vap_kgs, 1.0, note + " — W₂ ≤ vapor-only W → vapor-only governs"
    if w_vap_kgs > 0 and w2 > 3.0 * w_vap_kgs:
        note += "  ⚠ W₂ > 3·W — methodology requires safety/risk specialist review"
    return w2, x_rel, note


# ═══════════════════════════════════════════════════════════════════════════
#  CV failure (ISA / IEC 60534) & orifice flows
# ═══════════════════════════════════════════════════════════════════════════
def cv_liquid_kgs(cv: float, p1_pa: float, p2_pa: float, rho: float,
                  fl: float = 0.9, pv_pa: float = 0.0,
                  pc_pa: float = 22_064_000.0) -> float:
    """ISA liquid flow through a valve at Cv.  W [kg/s]."""
    if cv <= 0 or p1_pa <= p2_pa or rho <= 0:
        return 0.0
    sg = rho / 999.0
    ff = 0.96 - 0.28 * math.sqrt(max(0.0, pv_pa) / pc_pa)
    dp_max = (fl ** 2) * (p1_pa - ff * max(0.0, pv_pa))
    dp = min(p1_pa - p2_pa, max(1.0, dp_max))
    q_m3h = 0.865 * cv * math.sqrt((dp / 1e5) / max(1e-6, sg))
    return q_m3h * rho / 3600.0


def cv_gas_kgs(cv: float, p1_pa: float, p2_pa: float, t_k: float, mw: float,
               z: float, k: float, xt: float = 0.7) -> float:
    """ISA gas/vapor flow through a valve at Cv (choked-capped).  W [kg/s]."""
    if cv <= 0 or p1_pa <= 0:
        return 0.0
    fk = max(0.1, (k or 1.4) / 1.4)
    x = max(0.0, min(1.0, (p1_pa - max(0.0, p2_pa)) / p1_pa))
    x = min(x, fk * xt)
    y = max(2.0 / 3.0, 1.0 - x / (3.0 * fk * xt))
    rho1 = p1_pa * (mw or 28.96) / ((z or 1.0) * R_UNIV * max(1.0, t_k))
    w_kgh = 27.3 * cv * y * math.sqrt(x * (p1_pa / 1000.0) * rho1)
    return w_kgh / 3600.0


def orifice_gas_kgs(area_m2: float, cd: float, p1_pa: float, p2_pa: float,
                    t_k: float, mw: float, z: float, k: float) -> float:
    g, _ = gas_mass_flux(p1_pa, p2_pa, t_k, mw, z, k)
    return max(0.0, cd) * max(0.0, area_m2) * g


def orifice_liq_kgs(area_m2: float, cd: float, p1_pa: float, p2_pa: float,
                    rho: float) -> float:
    dp = p1_pa - p2_pa
    if dp <= 0 or rho <= 0:
        return 0.0
    return cd * area_m2 * math.sqrt(2.0 * rho * dp)


# ═══════════════════════════════════════════════════════════════════════════
#  Stream resolution — HMB linkage (per failure case) + manual override
# ═══════════════════════════════════════════════════════════════════════════
@dataclass
class StreamData:
    """Relieving-fluid properties for one scenario (internal SI)."""
    source: str = "manual"          # 'HMB:<file>:<stream>' or 'manual'
    t_k: float | None = None
    p_pa: float | None = None
    phase: str | None = None        # Vapor | Liquid | Two-Phase
    mw: float | None = None
    z: float | None = None
    k: float | None = None
    quality: float | None = None    # mass vapor fraction
    rho_v: float | None = None      # kg/m³
    rho_l: float | None = None
    mu_v: float | None = None       # cP
    mu_l: float | None = None
    latent: float | None = None     # J/kg
    cp_l: float | None = None       # J/(kg·K)
    sg: float | None = None
    feed: object | None = None      # FlashFeed (composition) when available
    sp: object | None = None        # hydraulics StreamProps
    notes: list[str] = field(default_factory=list)


def _hmb_path(fname: str, base_dir: str) -> str | None:
    if not fname:
        return None
    cands = [fname,
             os.path.join(base_dir, fname),
             os.path.join(base_dir, "..", "Hydraulics", fname)]
    for c in cands:
        if os.path.exists(c):
            return c
    return None


def resolve_stream(row: dict, uio: UIO, base_dir: str) -> StreamData:
    """Build StreamData from a STREAM_INPUTS row: HMB link first, then manual
    orange-cell overrides (manual wins where filled)."""
    sd = StreamData()
    fname, case, lookup = row.get("hmb_file"), row.get("hmb_case"), row.get("stream")
    if fname and lookup and HYD is not None:
        path = _hmb_path(fname, base_dir)
        if path is None:
            sd.notes.append(f"HMB file '{fname}' not found — manual props used")
        else:
            sp = HYD.load_stream_props(path, lookup, case or "Case 1")
            if sp is None:
                sd.notes.append(f"stream '{lookup}' not in {os.path.basename(path)}"
                                f" (case {case}) — manual props used")
            else:
                sd.source = f"HMB:{os.path.basename(path)}:{lookup}"
                sd.sp = sp
                sd.t_k = _FPS2SI["T"](sp.temp_f) if sp.temp_f is not None else None
                sd.p_pa = _FPS2SI["P"](sp.pres_psia) if sp.pres_psia else None
                sd.phase = sp.phase
                sd.mw = sp.mol_weight or sp.vap_mw
                sd.z = sp.vap_z
                sd.k = sp.vap_cp_cv
                tot = (sp.total_mass or 0.0)
                sd.quality = (sp.vap_mass or 0.0) / tot if tot > 0 else None
                sd.rho_v = _FPS2SI["rho"](sp.vap_density) if sp.vap_density else None
                sd.rho_l = _FPS2SI["rho"](sp.liq_density) if sp.liq_density else None
                sd.mu_v, sd.mu_l = sp.vap_visc, sp.liq_visc
                if sp.vap_sp_enthalpy is not None and sp.liq_sp_enthalpy is not None:
                    sd.latent = (sp.vap_sp_enthalpy - sp.liq_sp_enthalpy) * BTU_LB_TO_JKG
                if sp.liq_cp is not None:
                    sd.cp_l = sp.liq_cp * BTU_LB_TO_JKG * 1.8   # BTU/lb·°F → J/kg·K
                if sd.rho_l:
                    sd.sg = sd.rho_l / 999.0
                try:
                    is_case = HYD.is_proii_export(path)
                    sd.feed = HYD.read_feed(path, lookup, case or "Case 1",
                                            sp, is_casesheet=is_case)
                except Exception as exc:
                    sd.notes.append(f"composition not loaded: {exc}")
    elif fname and lookup and HYD is None:
        sd.notes.append("hydraulics engine unavailable — manual props used")

    # manual orange-cell overrides (user units → SI)
    ov = [("t_k", "T", "t"), ("p_pa", "P", "p"), ("mw", "MW", "mw"),
          ("z", "-", "z"), ("k", "-", "k"), ("quality", "-", "x"),
          ("rho_v", "rho", "rho_v"), ("rho_l", "rho", "rho_l"),
          ("mu_v", "visc", "mu_v"), ("mu_l", "visc", "mu_l"),
          ("latent", "h", "latent"), ("cp_l", "h", "cp_l"), ("sg", "-", "sg")]
    for attr, qty, key in ov:
        v = num(row.get(key))
        if v is None:
            continue
        if attr == "cp_l":                       # h-unit per dT
            setattr(sd, attr, uio.si("h", v) * (1.8 if qty else 1.0)
                    if uio.u.sel.get("h") == "BTU/lb" else uio.si("h", v))
        else:
            setattr(sd, attr, uio.si(qty, v))
    ph = txt(row.get("phase"))
    if ph:
        sd.phase = ph
    if sd.sg is None and sd.rho_l:
        sd.sg = sd.rho_l / 999.0
    return sd


# ── flash helpers (composition-based, via hydraulics flash VLE) ─────────────
def _flash(feed, p_pa: float, t_k: float | None = None):
    p_psia = p_pa / PSI_TO_PA
    t_f = (t_k - 273.15) * 1.8 + 32.0 if t_k is not None else None
    return HYD.flash(feed, p_psia, t_f)


def flash_adiabatic(feed, p_pa: float):
    return HYD.flash_isenthalpic(feed, p_pa / PSI_TO_PA)


def bubble_dew_t_k(feed, p_pa: float, target: str = "bubble") -> float | None:
    """Solve T where molar vapor fraction crosses 0 (bubble) or 1 (dew)."""
    if feed is None or not feed.has_constants():
        return None
    goal = 1e-4 if target == "bubble" else 1.0 - 1e-4
    t_ref = feed.temp_f if feed.temp_f is not None else 100.0

    def f(t_f):
        return _flash_beta(feed, p_pa, t_f) - goal

    lo, hi = t_ref - 400.0, t_ref + 400.0
    t = bisect(f, lo, hi, it=80, tol=1e-6)
    return (t - 32.0) / 1.8 + 273.15 if t is not None else None


def _flash_beta(feed, p_pa: float, t_f: float) -> float:
    try:
        return HYD.flash(feed, p_pa / PSI_TO_PA, t_f).beta
    except Exception:
        return 0.5


def latent_from_feed(feed, p_pa: float) -> float | None:
    """λ [J/kg] between bubble and dew at P from the HMB enthalpy model."""
    if feed is None or feed.vap_h_ref is None or feed.liq_h_ref is None:
        return None
    tb = bubble_dew_t_k(feed, p_pa, "bubble")
    td = bubble_dew_t_k(feed, p_pa, "dew")
    if tb is None or td is None:
        lam_btu = feed.vap_h_ref - feed.liq_h_ref
        return lam_btu * BTU_LB_TO_JKG if lam_btu and lam_btu > 0 else None
    tb_f = (tb - 273.15) * 1.8 + 32.0
    td_f = (td - 273.15) * 1.8 + 32.0
    t_ref = feed.temp_f or tb_f
    h_liq = feed.liq_h_ref + (feed.liq_cp or 0.5) * (tb_f - t_ref)
    h_vap = feed.vap_h_ref + (feed.vap_cp or 0.4) * (td_f - t_ref)
    lam = (h_vap - h_liq) * BTU_LB_TO_JKG
    return lam if lam > 0 else None


# ═══════════════════════════════════════════════════════════════════════════
#  Template builder
# ═══════════════════════════════════════════════════════════════════════════
# Palette, C()/_title()/_kv_block()/_result_rows() and the live unit-label
# helpers (_U, _unit_*, _hdr_formula, UNIT_ROWS) are shared with the CV tool
# via ../common/style.py — imported at the top of this file.


def create_input_template(out_path: str = "psv_input.xlsx",
                          unit_system: str = "FPS") -> str:
    usys = UN.UnitSystem(unit_system)
    uio = UIO(usys)
    L = lambda *parts: _U(*parts)
    wb = Workbook()

    # ── GENERAL ────────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "GENERAL"
    _title(ws, 4, f"PSV GENERAL DATA  —  identification · operating/design "
                  f"conditions · set pressure · static elevation correction   "
                  f"(units: {usys.system}, see UNITS sheet)")
    rows = [
        ("Identification", "", "", "", "sec"),
        ("Unit", "", "", "Process unit", "in"),
        ("P&ID No", "", "", "", "in"),
        ("PSV Tag", "PSV-001", "", "", "in"),
        ("Location", "", "", "", "in"),
        ("Service / fluid description", "", "", "", "in"),
        ("PSV Operating Conditions", "", "", "", "sec"),
        ("Operating pressure", None, L("P"), "Normal", "in"),
        ("Max operating pressure", None, L("P"), "Highest credible steady state", "in"),
        ("Min operating pressure", None, L("P"), "", "in"),
        ("Operating temperature", None, L("T"), "", "in"),
        ("Max operating temperature", None, L("T"), "", "in"),
        ("Min operating temperature", None, L("T"), "", "in"),
        ("PSV Design Conditions", "", "", "", "sec"),
        ("Design pressure", None, L("P"), "Usually = MAWP", "in"),
        ("MAWP", None, L("P"), "Max Allowable Working Pressure", "in"),
        ("MAP", None, L("P"), "Max Allowable Pressure (new & cold), optional", "in"),
        ("Design temperature", None, L("T"), "", "in"),
        ("Set pressure", None, L("P"), "Must be ≤ MAWP", "in"),
        ("Overpressure allowance", 10, "%",
         "10% single non-fire | 16% multiple PSVs | 21% fire — engine applies "
         "by scenario class automatically; this cell = design-class default", "in"),
        ("Number of PSVs (identical)", 1, "", "", "in"),
        ("Remote Contingency Rule (§2.2)", "", "", "", "sec"),
        ("Hydrotest factor HTF", 1.5, "",
         "Front-end default 1.5 — verify vs vessel code", "in"),
        ("Actual hydrotest pressure (optional)", None, L("P"),
         "If known.  Corrected for temperature stress-ratio & corrosion below", "in"),
        ("Test-pressure correction factor (optional)", None, "",
         "e.g. 0.833 × 0.91 — temperature & corrosion correction multiplier", "in"),
        ("Remote allowable pressure  (engine)", None, L("P"),
         "min(HTF × DP, corrected test P); fallback 1.3 × MAWP (ASME VIII-1)", "calc"),
        ("Static Elevation Correction (PSV Inputs plus)", "", "", "", "sec"),
        ("Max EL of protected equipment from grade", None, L("L"), "", "in"),
        ("Min EL of protected equipment from grade", None, L("L"), "", "in"),
        ("PSV elevation from grade", None, L("L"), "", "in"),
        ("Liquid density for static head", None, L("rho"),
         "Highest credible coincident density (§5.6)", "in"),
        ("Static head  (engine)", None, L("dP"),
         "ρ·g·(EL_max − EL_PSV); sign follows elevation difference", "calc"),
        ("Corrected set pressure  (engine)", None, L("P"),
         "Set P − static head seen by the PSV.  PSV above liquid top → head "
         "adds (flagged if top equipment exceeds MAWP at lift)", "calc"),
    ]
    _kv_block(ws, 3, rows)

    # ── EQUIPMENT ──────────────────────────────────────────────────────────
    ws = wb.create_sheet("EQUIPMENT")
    _title(ws, 10, "PROTECTED EQUIPMENT  —  up to 10 items (PSV Inputs plus) "
                   "+ fire geometry of the governing vessel")
    hdrs = ["#", "Equipment Tag", "Type",
            _hdr_formula("Op P", "P"), _hdr_formula("Op T", "T"),
            _hdr_formula("Des P", "P"), _hdr_formula("Des T", "T"),
            _hdr_formula("MAWP", "P"), _hdr_formula("EL from grade", "L"), "Notes"]
    for ci, h in enumerate(hdrs, 2):
        C(ws, 3, ci, h, bg=NAVY, fg=DGRAY, sz=9, bold=True, ha="center", wrap=True)
    ws.row_dimensions[3].height = 26
    for i in range(10):
        C(ws, 4 + i, 2, i + 1, bg=LGRAY, sz=9, ha="center")
        for ci in range(3, 12):
            C(ws, 4 + i, ci, None, bg=AMBER, fg=CYAN, sz=9)
    widths = [4, 16, 14, 10, 10, 10, 10, 10, 14, 28]
    for ci, w in enumerate(widths, 2):
        ws.column_dimensions[get_column_letter(ci)].width = w
    r = 16
    rows = [
        ("Fire Geometry — governing vessel (auto wetted area)", "", "", "", "sec"),
        ("Vessel orientation", "Vertical", "",
         "Vertical | Horizontal | Sphere", "in"),
        ("Vessel OD", None, L("L"), "", "in"),
        ("Tangent-to-tangent length", None, L("L"), "", "in"),
        ("Head type", "2:1 Ellipsoidal", "",
         "2:1 Ellipsoidal | Hemispherical | Flat", "in"),
        ("NLL from bottom TL", None, L("L"), "Liquid level for wetted area", "in"),
        ("Bottom TL elevation above grade", 0, L("L"), "", "in"),
        ("API 521 fire height limit", None, L("L"),
         "7.5 m / 25 ft above grade — enter in your units", "in"),
        ("Additional wetted area (other vessels/piping)", 0, L("A"),
         "Include piping when its volume > 15% of system wetted volume (§3.10)", "in"),
        ("Manual wetted area override", None, L("A"),
         "Leave blank to use the calculated area", "in"),
        ("Fire Credits & Screening (§3.10)", "", "", "", "sec"),
        ("Drainage adequate?", "N", "Y/N", "1% slope away, free drainage", "in"),
        ("Prompt fire-fighting?", "N", "Y/N", "water + trained crew ≤ 10 min", "in"),
        ("Insulated (fire-rated)?", "N", "Y/N",
         "Min environment factor F = 0.075 regardless of thickness", "in"),
        ("Manual F-factor override", None, "",
         "Leave blank → engine from credits", "in"),
        ("Fill fraction of expanded liquid volume", None, "%",
         "For vapor-only screening (85% / 95% rules)", "in"),
        ("Foaming system?", "N", "Y/N", "", "in"),
        ("Heat Exchanger (tube rupture / leak)", "", "", "", "sec"),
        ("High-pressure side", "Tube", "", "Shell | Tube", "in"),
        ("Shell-side MAWP", None, L("P"), "", "in"),
        ("Tube-side MAWP", None, L("P"), "", "in"),
        ("Tube OD", None, "in", "Always inches", "in"),
        ("Tube wall thickness", None, "in", "", "in"),
        ("Number of tubes", None, "", "", "in"),
        ("Vessel Volume", "", "", "", "sec"),
        ("Vessel internal volume", None, L("V"),
         "Thermal expansion / runaway / overfill time", "in"),
    ]
    _kv_block(ws, r, rows)

    # ── PRESSURE (back pressure & compliance) ─────────────────────────────
    ws = wb.create_sheet("PRESSURE")
    _title(ws, 4, "BACK PRESSURE & COMPLIANCE")
    rows = [
        ("Back Pressure", "", "", "", "sec"),
        ("Superimposed back pressure — constant", 0, L("P"),
         "Header pressure always present (gauge basis = same unit as set P)", "in"),
        ("Superimposed back pressure — variable max", 0, L("P"), "", "in"),
        ("Built-up back pressure  (engine)", None, L("dP"),
         "From OUTLET_PIPING at relief flow — iterative", "calc"),
        ("Total back pressure  (engine)", None, L("P"), "", "calc"),
        ("Allowable Limits", "", "", "", "sec"),
        ("Max BP — conventional", 10, "% of set", "API 520 (at 10% OP)", "in"),
        ("Max BP — balanced bellows", 35, "% of set", "Manufacturer curve", "in"),
        ("Max BP — pilot operated", 100, "% of set", "", "in"),
        ("Kb manufacturer override", None, "",
         "Leave blank → engine Kb from PSV type + BP", "in"),
        ("Compliance (engine)", "", "", "", "sec"),
        ("Set P ≤ MAWP", None, "", "", "calc"),
        ("Accumulation ≤ class allowable", None, "", "", "calc"),
        ("Total BP ≤ limit for PSV type", None, "", "", "calc"),
        ("Inlet dP ≤ 3% of set pressure", None, "", "", "calc"),
    ]
    _kv_block(ws, 3, rows)

    # ── PSV_CONFIG ────────────────────────────────────────────────────────
    ws = wb.create_sheet("PSV_CONFIG")
    _title(ws, 4, "PSV CONFIGURATION — hardware & discharge coefficients")
    rows = [
        ("Hardware", "", "", "", "sec"),
        ("PSV type", "Conventional", "",
         "Conventional | Balanced Bellows | Pilot-Operated", "in"),
        ("Inlet connection NPS", None, "in", "Always inches", "in"),
        ("Outlet connection NPS", None, "in", "", "in"),
        ("Rupture disc upstream?", "N", "Y/N", "Kc = 0.9 if Y (uncertified)", "in"),
        ("Discharge Coefficients (API 520)", "", "", "", "sec"),
        ("Kd — vapor/gas", 0.975, "", "", "in"),
        ("Kd — liquid", 0.65, "", "", "in"),
        ("Kd — two-phase", 0.85, "", "Leung / Annex C recommendation", "in"),
        ("Kc — combination factor", 1.0, "", "0.9 with rupture disc", "in"),
    ]
    _kv_block(ws, 3, rows)
    return _finish_template(wb, uio, out_path)


# ── STREAM_INPUTS column map (canonical key → header base, qty) ─────────────
STREAM_COLS = [
    ("code",    "Scenario",            None),
    ("name",    "Description",         None),
    ("active",  "Active (Y/N)",        None),
    ("klass",   "Class",               None),
    ("hmb_file", "HMB File",           None),
    ("hmb_case", "HMB Case",           None),
    ("stream",  "Stream Lookup",       None),
    ("t",       "Temp",                "T"),
    ("p",       "Pressure",            "P"),
    ("phase",   "Phase",               None),
    ("mw",      "MW",                  "MW"),
    ("z",       "Z",                   None),
    ("k",       "Cp/Cv k",             None),
    ("x",       "Quality x",           None),
    ("rho_v",   "Vap Density",         "rho"),
    ("rho_l",   "Liq Density",         "rho"),
    ("mu_v",    "Vap Visc (cP)",       None),
    ("mu_l",    "Liq Visc (cP)",       None),
    ("latent",  "Latent Heat λ",       "h"),
    ("cp_l",    "Liq Cp",              "h"),
    ("sg",      "Liq SG",              None),
    ("notes",   "Notes",               None),
]


def _finish_template(wb: Workbook, uio: UIO, out_path: str) -> str:
    L = lambda *parts: _U(*parts)

    # ── STREAM_INPUTS ─────────────────────────────────────────────────────
    ws = wb.create_sheet("STREAM_INPUTS")
    ncol = len(STREAM_COLS)
    _title(ws, ncol, "STREAM INPUTS — one row per failure case.  Link an HMB "
                     "workbook + stream per scenario, or fill the orange "
                     "manual-property cells (manual overrides HMB).")
    C(ws, 2, 2, "Class: Design | Remote | Fire — sets the allowable "
                "accumulation per the methodology (§2).  Orange = manual "
                "properties (same convention as the hydraulics input).",
      sz=9, fg="808080")
    for ci, (key, base, qty) in enumerate(STREAM_COLS, 2):
        h = _hdr_formula(base, qty) if qty else base
        C(ws, 4, ci, h, bg=NAVY, fg=DGRAY, sz=9, bold=True, ha="center", wrap=True)
    ws.row_dimensions[4].height = 28
    manual_keys = {"t", "p", "phase", "mw", "z", "k", "x", "rho_v", "rho_l",
                   "mu_v", "mu_l", "latent", "cp_l", "sg"}
    for ri, (code, name) in enumerate(SCEN_CODES, 5):
        for ci, (key, _, _q) in enumerate(STREAM_COLS, 2):
            if key == "code":
                C(ws, ri, ci, code, bg=LGRAY, sz=9, bold=True)
            elif key == "name":
                C(ws, ri, ci, name, bg=LGRAY, sz=9)
            elif key == "active":
                C(ws, ri, ci, "N", bg=AMBER, fg=CYAN, sz=9, ha="center")
            elif key == "klass":
                default = ("Fire" if code == "FIRE" else "Design")
                C(ws, ri, ci, default, bg=AMBER, fg=CYAN, sz=9, ha="center")
            elif key in manual_keys:
                C(ws, ri, ci, None, bg=AMBER, fg=CYAN, sz=9, ha="right")
            else:
                C(ws, ri, ci, None, bg=TEAL, sz=9)
    dv = DataValidation(type="list", formula1='"Design,Remote,Fire"',
                        allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("E5:E17")
    dv2 = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    ws.add_data_validation(dv2)
    dv2.add("D5:D17")
    for ci, w in enumerate([10, 30, 9, 9, 16, 10, 13] + [9] * 14 + [24], 2):
        ws.column_dimensions[get_column_letter(ci)].width = w

    # ── SCENARIOS (per-scenario user inputs; props come from STREAM_INPUTS)
    ws = wb.create_sheet("SCENARIOS")
    _title(ws, 4, "SCENARIO INPUTS — fill only the ACTIVE scenarios "
                  "(activation on STREAM_INPUTS).  Engine results land in "
                  "the output workbook.")
    rows = [
        ("1 · BLOCKED OUTLET", "", "", "", "sec"),
        ("BO: Relieving rate", None, L("mflow"),
         "From simulation at blocked-outlet conditions (peak flow)", "in"),
        ("BO: Basis / simulation reference", "", "", "", "in"),
        ("2 · EXTERNAL FIRE — geometry on EQUIPMENT sheet", "", "", "", "sec"),
        ("FIRE: Manual Q override", None, L("Q"),
         "Blank → engine from wetted area & F", "in"),
        ("FIRE: Manual relief rate override", None, L("mflow"),
         "Blank → Q/λ (λ from stream/flash)", "in"),
        ("FIRE: Manual two-phase W2 override", None, L("mflow"),
         "Blank → engine homogeneous-swell estimate when screening fails", "in"),
        ("3 · CONTROL VALVE FAILURE — valve data on CV_DATA sheet", "", "", "", "sec"),
        ("CV: Evaluate as", "Design", "",
         "Design (bypass at 50% of normal operating Cv) | Remote (bypass 100% "
         "rated Cv) — §3.3", "in"),
        ("CV: Outflow credit", 0, L("mflow"),
         "Pre-contingency outflow at relief ΔP (documented basis only)", "in"),
        ("4 · GAS BLOWBY / CHECK-VALVE REVERSE — source on GAS_DATA sheet",
         "", "", "", "sec"),
        ("5 · THERMAL EXPANSION", "", "", "", "sec"),
        ("THERM: Heat input rate", None, L("Q"),
         "Exchanger hot-side / tracing / solar duty into the blocked liquid", "in"),
        ("THERM: Cubic expansion coefficient β", None, _U("1/", "dT"),
         "At operating T (water 60°F ≈ 0.000113/°F)", "in"),
        ("6 · TUBE RUPTURE / LEAK — HX data on EQUIPMENT sheet", "", "", "", "sec"),
        ("TUBE: Number of tube failures", 1, "",
         "1 full tube (both ends) typical; 10 when ΔP ≥ 1000 psi & not excluded "
         "(§3.11)", "in"),
        ("TUBE: Evaluate 6 mm leak instead of rupture", "N", "Y/N",
         "Tube-leak design basis: single 6 mm (0.25 in) hole", "in"),
        ("TUBE: HP-side max pressure override", None, L("P"),
         "Blank → HP-side MAWP from EQUIPMENT", "in"),
        ("7 · UTILITY FAILURE / LOSS OF COOLING", "", "", "", "sec"),
        ("UTIL: Lost cooling/condensing duty", None, L("Q"), "", "in"),
        ("UTIL: Vaporising fraction", 1.0, "", "≤ 1.0", "in"),
        ("8 · TOWER LOSS OF CONDENSING (SI method, §5.3)", "", "", "", "sec"),
        ("LOC: HMVRAF rate", None, L("mflow"),
         "Highest Molar Vapor Rate Above the Feed (mass units).  Engine "
         "flashes it adiabatically to relief pressure when composition is "
         "linked; else uses it directly", "in"),
        ("LOC: Outflow credit", 0, L("mflow"),
         "Per §5.3 matrix (condenser type × PRV location)", "in"),
        ("9 · LIQUID OVERFILL (§5.6 simple method)", "", "", "", "sec"),
        ("OVERFILL: Feed rate", None, L("mflow"), "", "in"),
        ("OVERFILL: Continuing liquid outflow credit", 0, L("mflow"),
         "Reflux + liquid distillate at pre-contingency volumetric rates", "in"),
        ("OVERFILL: Max reboiler vapor rate", 0, L("mflow"),
         "For relief quality assembly", "in"),
        ("OVERFILL: Max feed vapor rate", 0, L("mflow"), "", "in"),
        ("OVERFILL: Friction dP equipment→PSV at overfill flow", 0, L("dP"),
         "For the associated-equipment DPeq check (§3.4)", "in"),
        ("10 · LIQUID BOIL-UP (§5.7 two-flash λ method)", "", "", "", "sec"),
        ("BOILUP: Reboiler duty", None, L("Q"),
         "Pre-contingency duty (or pinched — see PINCH below)", "in"),
        ("BOILUP: Manual λ override", None, L("h"),
         "Blank → engine two-flash λ from linked composition", "in"),
        ("REBOILER PINCH (§5.1) — applied to BOILUP / OVERFILL / ABHEAT",
         "", "", "", "sec"),
        ("PINCH: Apply pinch credit?", "N", "Y/N", "", "in"),
        ("PINCH: U × A", None, _U("Q", "/", "dT"),
         "Overall coefficient × area (clean U if medium flow rises)", "in"),
        ("PINCH: Heating medium temperature", None, L("T"), "", "in"),
        ("PINCH: Reboiler-feed bubble point at accumulation", None, L("T"),
         "Blank → engine from linked composition", "in"),
        ("11 · ABNORMAL HEAT INPUT (§5.4)", "", "", "", "sec"),
        ("ABHEAT: Reboiler type", "Steam", "", "Steam | Sensible | Fired", "in"),
        ("ABHEAT: Pre-contingency duty Q0", None, L("Q"), "", "in"),
        ("ABHEAT: Steam latent heat", None, L("h"),
         "At condensing pressure (steam reboilers)", "in"),
        ("ABHEAT: CV normal operating Cv", None, "", "", "in"),
        ("ABHEAT: CV rated full-open Cv", None, "", "", "in"),
        ("ABHEAT: Design duty (fired only)", None, L("Q"),
         "Engine uses 1.25 × design duty for fired reboilers", "in"),
        ("ABHEAT: Base overhead vapor rate", None, L("mflow"),
         "Pre-contingency overhead, flashed adiabatically to relief P", "in"),
        ("ABHEAT: Condenser type", "Total", "", "Total | Partial", "in"),
        ("ABHEAT: Pre-contingency drum vapor rate (partial)", 0, L("mflow"), "", "in"),
        ("12 · RUNAWAY REACTION", "", "", "", "sec"),
        ("RUNAWAY: Heat generation at runaway", None, L("Q"),
         "From calorimetry (ARC/DSC)", "in"),
        ("RUNAWAY: System type", "Tempered", "", "Tempered | Gassy", "in"),
        ("RUNAWAY: DIERS vapor rate (if available)", None, L("mflow"), "", "in"),
        ("13 · OTHER / MANUAL", "", "", "", "sec"),
        ("OTHER: Relieving rate", None, L("mflow"), "", "in"),
        ("OTHER: Description", "", "", "", "in"),
    ]
    _kv_block(ws, 3, rows)

    # ── CV_DATA ───────────────────────────────────────────────────────────
    ws = wb.create_sheet("CV_DATA")
    _title(ws, 4, "CONTROL VALVE FAILURE DATA — inlet CV that can "
                  "overpressure the protected system (§3.3)")
    rows = []
    for i in (1, 2):
        rows += [
            (f"CV{i}", "", "", "", "sec"),
            (f"CV{i}: Tag", "", "", "", "in"),
            (f"CV{i}: Phase", "Gas", "", "Gas | Liquid", "in"),
            (f"CV{i}: Rated Cv (full open)", None, "", "", "in"),
            (f"CV{i}: Normal operating Cv", None, "",
             "For the 50%-bypass design case", "in"),
            (f"CV{i}: Bypass rated Cv", None, "",
             "For the remote case (bypass 100% open)", "in"),
            (f"CV{i}: FL", 0.9, "", "Liquid pressure recovery", "in"),
            (f"CV{i}: xT", 0.7, "", "Choked pressure-drop ratio (gas)", "in"),
            (f"CV{i}: Max upstream pressure", None, L("P"), "", "in"),
            (f"CV{i}: Upstream temperature", None, L("T"), "", "in"),
            (f"CV{i}: Stream code", "CV", "",
             "STREAM_INPUTS row giving fluid props (default CV)", "in"),
        ]
    _kv_block(ws, 3, rows)

    # ── GAS_DATA ──────────────────────────────────────────────────────────
    ws = wb.create_sheet("GAS_DATA")
    _title(ws, 4, "GAS BLOWBY / CHECK-VALVE REVERSE FLOW — source & "
                  "restriction (§3.7)")
    rows = [
        ("GB1", "", "", "", "sec"),
        ("GB1: Source description", "", "", "", "in"),
        ("GB1: Restriction type", "Cv", "",
         "Cv | Orifice | CheckValve", "in"),
        ("GB1: Cv (full open)", None, "", "If type = Cv", "in"),
        ("GB1: xT", 0.7, "", "", "in"),
        ("GB1: Orifice area", None, L("Ain"), "If type = Orifice", "in"),
        ("GB1: Discharge coefficient Cd", 0.61, "",
         "0.61 sharp edge | 0.82 rounded", "in"),
        ("GB1: Check valve min internal diameter", None, "in",
         "If type = CheckValve — always inches", "in"),
        ("GB1: Number of check valves in series", 1, "", "1 or 2", "in"),
        ("GB1: Partial-failure orifice — HP valve", 6, "% of ID",
         "§3.7: 6% of min internal diameter (design)", "in"),
        ("GB1: Partial-failure orifice — LP valve", 2, "% of ID",
         "2% for the second (LP) valve in series", "in"),
        ("GB1: Max source pressure", None, L("P"), "", "in"),
        ("GB1: Source temperature", None, L("T"), "", "in"),
        ("GB1: Stream code", "GASBB", "", "STREAM_INPUTS props row", "in"),
    ]
    _kv_block(ws, 3, rows)

    # ── INLET / OUTLET piping ─────────────────────────────────────────────
    # Shared dropdown source for the fittings tables (hidden sheet, mirrors
    # the Hydraulics tool's "_FittingList" convention).
    ws_fit = wb.create_sheet("_FittingList")
    ws_fit.sheet_state = "hidden"
    for i, fn in enumerate(FITTING_NAMES, 1):
        ws_fit.cell(row=i, column=1, value=fn)
    n_fit = len(FITTING_NAMES)

    for nm, ttl in (("INLET_PIPING", "INLET PIPING — 3% rule check"),
                    ("OUTLET_PIPING", "OUTLET PIPING — built-up back pressure")):
        ws = wb.create_sheet(nm)
        _title(ws, 4, ttl + "  (define once — per-case velocity/Re/dP results "
                            "appear on each CASE sheet)")
        rows = [
            ("Geometry", "", "", "", "sec"),
            ("Line inside diameter", None, "in", "Always inches", "in"),
            ("Straight length", None, L("L"), "", "in"),
            ("Elevation change", 0, L("L"), "+ up", "in"),
            ("Fittings", "", "", "", "sec"),
        ]
        r = _kv_block(ws, 3, rows)
        C(ws, r, 2, "Fitting type (Hydraulics catalogue)", bg=NAVY, fg=DGRAY,
          sz=9, bold=True, ha="center", wrap=True)
        C(ws, r, 3, "Quantity", bg=NAVY, fg=DGRAY, sz=9, bold=True, ha="center")
        C(ws, r, 4, "", bg=NAVY)
        C(ws, r, 5, "K-each and ΣK are looked up and totalled per case by "
                     "the engine from the Hydraulics fittings catalogue.",
          sz=9, fg="808080", wrap=True)
        dv_fit = DataValidation(type="list",
                                 formula1=f"_FittingList!$A$1:$A${n_fit}",
                                 allow_blank=True, showErrorMessage=True,
                                 error="Choose a fitting from the list",
                                 errorTitle="Invalid Fitting")
        ws.add_data_validation(dv_fit)
        fit_first = r + 1
        for i in range(N_FIT_ROWS):
            rr = fit_first + i
            C(ws, rr, 2, None, bg=AMBER, fg=CYAN, sz=9)
            C(ws, rr, 3, None, bg=AMBER, fg=CYAN, sz=9, ha="right")
        dv_fit.add(f"B{fit_first}:B{fit_first + N_FIT_ROWS - 1}")
        rows2 = [
            ("Override", "", "", "", "sec"),
            ("Manual dP override at relief flow", None, L("dP"),
             "From the hydraulics engine (recommended for complex circuits).  "
             "Blank → engine line model from the geometry & fittings above",
             "in"),
        ]
        _kv_block(ws, fit_first + N_FIT_ROWS, rows2)

    # ── DISPOSAL ──────────────────────────────────────────────────────────
    ws = wb.create_sheet("DISPOSAL")
    _title(ws, 4, "DISPOSAL ROUTING ADVISOR (§7) — answer Y/N; engine "
                  "recommends Flare / Atmosphere")
    rows = [
        ("Stream character", "", "", "", "sec"),
        ("Liquid or partially liquid at PSV inlet?", "N", "Y/N", "", "in"),
        ("Flammable vapor MW > 80 or condenses at ambient?", "N", "Y/N", "", "in"),
        ("Toxic service (AEGL limits at platforms/fence)?", "N", "Y/N", "", "in"),
        ("H2S-containing?", "N", "Y/N",
         "PRDs route to the HC flare — never the acid-gas flare (§7.3)", "in"),
        ("Corrosive condensables (e.g. phenol)?", "N", "Y/N", "", "in"),
        ("Release / dispersion", "", "", "", "sec"),
        ("> 50% LFL at grade/platforms/fence?", "N", "Y/N", "", "in"),
        ("Radiant flux > 9.5 kW/m² if ignited?", "N", "Y/N", "", "in"),
        ("Operating above auto-ignition at PSV inlet?", "N", "Y/N", "", "in"),
        ("≥ 50 mol% H2 or T > min(AIT, 315°C)?", "N", "Y/N", "", "in"),
        ("Overfill basis", "", "", "", "sec"),
        ("Overfill is a DESIGN contingency on this vessel?", "N", "Y/N",
         "Vapor-space PSVs with design-contingency overfill must go to flare", "in"),
        ("Benign water-type liquid < 65°C?", "N", "Y/N", "", "in"),
        ("Verdict (engine)", "", "", "", "sec"),
        ("Recommended destination", None, "", "", "calc"),
    ]
    _kv_block(ws, 3, rows)

    UN.add_units_sheet(wb, uio.system, position=len(wb.worksheets))
    wb.save(out_path)
    return out_path


# ═══════════════════════════════════════════════════════════════════════════
#  Reader
# ═══════════════════════════════════════════════════════════════════════════
def _read_kv(ws) -> dict:
    out = {}
    for r in range(1, ws.max_row + 1):
        lbl = ws.cell(r, 2).value
        if lbl is None:
            continue
        out[str(lbl).strip()] = ws.cell(r, 3).value
    return out


def _read_piping(ws) -> dict:
    """Read an INLET/OUTLET_PIPING sheet: geometry/override key-values plus
    the fittings table (rows ``fit_first .. fit_first + N_FIT_ROWS - 1``,
    columns B = fitting name, C = quantity).  Adds ``"Total fittings K"``
    (ΣK over the table, via the Hydraulics catalogue) and ``"_fittings"``
    (non-zero ``(name, qty, k)`` rows, for per-case result display)."""
    kv = _read_kv(ws)
    fit_first = 3 + 5 + 1          # geometry block (5 rows) + table header
    fittings, ktot = [], 0.0
    for i in range(N_FIT_ROWS):
        rr = fit_first + i
        name = txt(ws.cell(rr, 2).value)
        qty = num(ws.cell(rr, 3).value) or 0.0
        if not name or qty == 0:
            continue
        k = fitting_k(name)
        fittings.append((name, qty, k))
        ktot += k * qty
    kv["Total fittings K"] = ktot
    kv["_fittings"] = fittings
    return kv


@dataclass
class Model:
    uio: UIO = None
    gen: dict = field(default_factory=dict)        # GENERAL (SI where numeric)
    equip_rows: list = field(default_factory=list)
    eq: dict = field(default_factory=dict)         # EQUIPMENT kv (SI)
    press: dict = field(default_factory=dict)
    cfg: dict = field(default_factory=dict)
    streams: dict = field(default_factory=dict)    # code → raw row dict
    scen: dict = field(default_factory=dict)       # SCENARIOS kv (raw)
    cv: dict = field(default_factory=dict)
    gas: dict = field(default_factory=dict)
    inlet: dict = field(default_factory=dict)
    outlet: dict = field(default_factory=dict)
    disposal: dict = field(default_factory=dict)
    base_dir: str = "."


def read_model(path: str) -> Model:
    wb = load_workbook(path, data_only=True)
    uio = UIO(UN.UnitSystem.from_workbook(wb))
    m = Model(uio=uio, base_dir=os.path.dirname(os.path.abspath(path)) or ".")

    g = _read_kv(wb["GENERAL"])
    sg = lambda lbl: txt(g.get(lbl))
    sn = lambda lbl, qty=None: (uio.si(qty, num(g.get(lbl)))
                                 if qty else num(g.get(lbl)))
    m.gen = {
        "unit": sg("Unit"), "pid": sg("P&ID No"),
        "tag": sg("PSV Tag") or "PSV", "location": sg("Location"),
        "service": sg("Service / fluid description"),
        "op_p": sn("Operating pressure", "P"),
        "max_op_p": sn("Max operating pressure", "P"),
        "min_op_p": sn("Min operating pressure", "P"),
        "op_t": sn("Operating temperature", "T"),
        "max_op_t": sn("Max operating temperature", "T"),
        "min_op_t": sn("Min operating temperature", "T"),
        "design_p": sn("Design pressure", "P"),
        "mawp": sn("MAWP", "P"),
        "map": sn("MAP", "P"),
        "design_t": sn("Design temperature", "T"),
        "set_p": sn("Set pressure", "P"),
        "op_pct": sn("Overpressure allowance") or 10.0,
        "n_psv": int(sn("Number of PSVs (identical)") or 1),
        "htf": sn("Hydrotest factor HTF") or 1.5,
        "test_p": sn("Actual hydrotest pressure (optional)", "P"),
        "test_corr": sn("Test-pressure correction factor (optional)"),
        "el_max": sn("Max EL of protected equipment from grade", "L"),
        "el_min": sn("Min EL of protected equipment from grade", "L"),
        "el_psv": sn("PSV elevation from grade", "L"),
        "rho_head": sn("Liquid density for static head", "rho"),
    }

    ws = wb["EQUIPMENT"]
    for r in range(4, 14):
        tag = txt(ws.cell(r, 3).value)
        if not tag:
            continue
        m.equip_rows.append({
            "tag": tag, "type": txt(ws.cell(r, 4).value),
            "op_p": uio.si("P", num(ws.cell(r, 5).value)),
            "op_t": uio.si("T", num(ws.cell(r, 6).value)),
            "des_p": uio.si("P", num(ws.cell(r, 7).value)),
            "des_t": uio.si("T", num(ws.cell(r, 8).value)),
            "mawp": uio.si("P", num(ws.cell(r, 9).value)),
            "el": uio.si("L", num(ws.cell(r, 10).value)),
            "notes": txt(ws.cell(r, 11).value),
        })
    e = _read_kv(ws)
    en = lambda lbl, qty=None: (uio.si(qty, num(e.get(lbl)))
                                 if qty else num(e.get(lbl)))
    m.eq = {
        "orient": txt(e.get("Vessel orientation")) or "Vertical",
        "od": en("Vessel OD", "L"),
        "tt": en("Tangent-to-tangent length", "L"),
        "head": txt(e.get("Head type")) or "2:1",
        "nll": en("NLL from bottom TL", "L"),
        "grade_el": en("Bottom TL elevation above grade", "L") or 0.0,
        "fire_lim": en("API 521 fire height limit", "L") or 7.62,
        "extra_a": en("Additional wetted area (other vessels/piping)", "A") or 0.0,
        "a_override": en("Manual wetted area override", "A"),
        "drain": yes(e.get("Drainage adequate?")),
        "ff": yes(e.get("Prompt fire-fighting?")),
        "insul": yes(e.get("Insulated (fire-rated)?")),
        "f_override": en("Manual F-factor override"),
        "fill_pct": en("Fill fraction of expanded liquid volume"),
        "foaming": yes(e.get("Foaming system?")),
        "hp_side": txt(e.get("High-pressure side")) or "Tube",
        "shell_mawp": en("Shell-side MAWP", "P"),
        "tube_mawp": en("Tube-side MAWP", "P"),
        "tube_od_in": en("Tube OD"),
        "tube_wall_in": en("Tube wall thickness"),
        "n_tubes": en("Number of tubes"),
        "volume": en("Vessel internal volume", "V"),
    }

    p = _read_kv(wb["PRESSURE"])
    pn = lambda lbl, qty=None: (uio.si(qty, num(p.get(lbl))) if qty
                                 else num(p.get(lbl)))
    m.press = {
        "super_const": pn("Superimposed back pressure — constant", "dP") or 0.0,
        "super_var": pn("Superimposed back pressure — variable max", "dP") or 0.0,
        "bp_conv": pn("Max BP — conventional") or 10.0,
        "bp_bell": pn("Max BP — balanced bellows") or 35.0,
        "bp_pilot": pn("Max BP — pilot operated") or 100.0,
        "kb_user": pn("Kb manufacturer override"),
    }

    c = _read_kv(wb["PSV_CONFIG"])
    m.cfg = {
        "type": txt(c.get("PSV type")) or "Conventional",
        "rd": yes(c.get("Rupture disc upstream?")),
        "kd_v": num(c.get("Kd — vapor/gas")) or 0.975,
        "kd_l": num(c.get("Kd — liquid")) or 0.65,
        "kd_2": num(c.get("Kd — two-phase")) or 0.85,
        "kc": num(c.get("Kc — combination factor")) or 1.0,
    }
    if m.cfg["rd"] and m.cfg["kc"] >= 1.0:
        m.cfg["kc"] = 0.9

    ws = wb["STREAM_INPUTS"]
    for r in range(5, 5 + len(SCEN_CODES)):
        code = txt(ws.cell(r, 2).value)
        if not code:
            continue
        row = {}
        for ci, (key, _b, _q) in enumerate(STREAM_COLS, 2):
            row[key] = ws.cell(r, ci).value
        m.streams[code.upper()] = row

    m.scen = _read_kv(wb["SCENARIOS"])
    m.cv = _read_kv(wb["CV_DATA"])
    m.gas = _read_kv(wb["GAS_DATA"])
    m.inlet = _read_piping(wb["INLET_PIPING"])
    m.outlet = _read_piping(wb["OUTLET_PIPING"])
    m.disposal = _read_kv(wb["DISPOSAL"])
    wb.close()
    return m


# ═══════════════════════════════════════════════════════════════════════════
#  Scenario engine
# ═══════════════════════════════════════════════════════════════════════════
@dataclass
class Scen:
    code: str
    name: str
    klass: str = "Design"
    active: bool = False
    w: float = 0.0                  # kg/s
    t: float = 300.0                # K
    p1: float = 0.0                 # Pa abs (relieving)
    p2: float = ATM_PA              # Pa abs (back pressure)
    allowable_g: float = 0.0        # Pa gauge allowable accumulation
    phase: str = "Vapor"
    quality: float = 1.0
    mw: float = 28.96
    z: float = 1.0
    k: float = 1.4
    rho_v: float | None = None
    rho_l: float | None = None
    mu: float = 0.3
    latent: float | None = None
    cp_l: float | None = None
    area: float = 0.0               # m² required
    area_mult: float = 1.0
    omega: float | None = None
    kv_visc: float = 1.0            # liquid viscosity correction (display)
    kw_bp: float = 1.0              # liquid back-pressure factor (display)
    notes: list = field(default_factory=list)
    # per-case back-pressure / inlet-line results (filled by size_case)
    kb: float = 1.0
    builtup_pa: float = 0.0         # built-up back pressure, this case's flow
    total_bp_g: float = 0.0         # super + built-up, Pa gauge
    bp_pct: float = 0.0             # total BP as % of set pressure
    bp_lim_pct: float = 10.0        # allowable BP per PSV type
    bp_ok: bool = True
    inlet_dp_pa: float = 0.0        # inlet-line dP, this case's flow
    in_pct: float = 0.0             # inlet dP as % of set pressure
    in_ok: bool = True
    accum_ok: bool = True
    orifice: str = ""
    orifice_area: float = 0.0       # m²
    outlet_info: dict = field(default_factory=dict)
    inlet_info: dict = field(default_factory=dict)


def _remote_allowable_g(gen: dict) -> tuple[float, str]:
    dp_g = (gen.get("design_p") or gen.get("mawp") or 0.0) - ATM_PA
    mawp_g = (gen.get("mawp") or gen.get("design_p") or 0.0) - ATM_PA
    cands = []
    note = []
    if dp_g > 0:
        cands.append(gen["htf"] * dp_g)
        note.append(f"HTF×DP = {gen['htf']:.2f}×DP")
    if gen.get("test_p"):
        tp_g = gen["test_p"] - ATM_PA
        corr = gen.get("test_corr") or 1.0
        cands.append(tp_g * corr)
        note.append("corrected test P")
    if not cands and mawp_g > 0:
        return 1.3 * mawp_g, "1.3 × MAWP fallback (ASME VIII-1 2007)"
    return (min(cands) if cands else 0.0), " | ".join(note)


def _class_allowable_g(klass: str, gen: dict) -> tuple[float, str]:
    set_g = (gen.get("set_p") or 0.0) - ATM_PA
    kl = (klass or "Design").upper()
    if kl.startswith("FIRE"):
        return set_g * 1.21, "121% set (fire)"
    if kl.startswith("REMOTE"):
        ra, note = _remote_allowable_g(gen)
        return ra, f"remote rule: {note}"
    pct = 16.0 if gen.get("n_psv", 1) > 1 else (gen.get("op_pct") or 10.0)
    return set_g * (1.0 + pct / 100.0), f"{100 + pct:.0f}% set"


def _relief_pressure(klass: str, gen: dict) -> float:
    g, _ = _class_allowable_g(klass, gen)
    return g + ATM_PA


def _apply_stream(s: Scen, sd: StreamData):
    if sd.t_k is not None:
        s.t = sd.t_k
    if sd.phase:
        ph = sd.phase.upper()
        s.phase = ("Two-Phase" if "TWO" in ph or "MIX" in ph
                   else "Liquid" if "LIQ" in ph else "Vapor")
    if sd.quality is not None:
        s.quality = sd.quality
    if sd.mw:
        s.mw = sd.mw
    if sd.z:
        s.z = sd.z
    if sd.k:
        s.k = sd.k
    if sd.rho_v:
        s.rho_v = sd.rho_v
    if sd.rho_l:
        s.rho_l = sd.rho_l
    s.mu = (sd.mu_l if s.phase == "Liquid" else sd.mu_v) or s.mu
    if sd.latent:
        s.latent = sd.latent
    if sd.cp_l:
        s.cp_l = sd.cp_l
    s.notes.extend(sd.notes)


def _rho_v_ideal(s: Scen) -> float:
    if s.rho_v:
        return s.rho_v
    return s.p1 * s.mw / (s.z * R_UNIV * max(1.0, s.t))


def compute_scenarios(m: Model) -> list[Scen]:
    uio, gen, scen = m.uio, m.gen, m.scen
    sv = lambda lbl, qty=None: (uio.si(qty, num(scen.get(lbl)))
                                 if qty else num(scen.get(lbl)))
    out: list[Scen] = []

    for code, name in SCEN_CODES:
        row = m.streams.get(code, {})
        s = Scen(code=code, name=name)
        s.active = yes(row.get("active"))
        s.klass = txt(row.get("klass")) or ("Fire" if code == "FIRE" else "Design")
        if not s.active:
            out.append(s)
            continue
        s.allowable_g, allow_note = _class_allowable_g(s.klass, gen)
        s.p1 = _relief_pressure(s.klass, gen)
        s.notes.append(f"allowable: {allow_note}")
        sd = resolve_stream(row, uio, m.base_dir)
        _apply_stream(s, sd)
        if sd.source != "manual":
            s.notes.append(f"props: {sd.source}")

        if code == "BO":
            s.w = sv("BO: Relieving rate", "mflow") or 0.0

        elif code == "FIRE":
            _fire(s, m, sd, sv)

        elif code == "CV":
            _cv_failure(s, m, sd)

        elif code == "GASBB":
            _gas_blowby(s, m, sd)

        elif code == "THERM":
            q = sv("THERM: Heat input rate", "Q") or 0.0
            beta = sv("THERM: Cubic expansion coefficient β")
            if beta and uio.u.sel.get("dT", "degF") in ("degF",):
                beta *= 1.8                       # per °F → per K
            cp = s.cp_l or 4180.0
            s.w = q * (beta or 0.0) / cp
            s.phase = "Liquid"
            s.notes.append("thermal expansion W = q·β/Cp")

        elif code == "TUBE":
            _tube(s, m, sd, sv)

        elif code == "UTIL":
            q = sv("UTIL: Lost cooling/condensing duty", "Q") or 0.0
            frac = sv("UTIL: Vaporising fraction") or 1.0
            lam = s.latent or latent_from_feed(sd.feed, s.p1)
            if lam:
                s.w = q * frac / lam
                s.phase = "Vapor"
            else:
                s.notes.append("no latent heat available — rate not computed")

        elif code == "LOC":
            w_in = sv("LOC: HMVRAF rate", "mflow") or 0.0
            credit = sv("LOC: Outflow credit", "mflow") or 0.0
            if sd.feed is not None and HYD is not None:
                fr = flash_adiabatic(sd.feed, s.p1)
                s.t = ((fr.temp_f - 32.0) / 1.8 + 273.15
                       if fr.temp_f is not None else s.t)
                s.quality = fr.quality
                s.mw = fr.vap_mw or s.mw
                s.phase = ("Vapor" if fr.quality > 0.999 else "Two-Phase")
                s.notes.append(f"HMVRAF flashed adiabatically to relief P: "
                               f"x={fr.quality:.3f}")
                if s.p1 > 1.05 * sd.feed.p_ref_psia * PSI_TO_PA:
                    s.notes.append(
                        "⚠ relief P exceeds the HMB stream reference P — "
                        "verify the linked stream represents relieving "
                        "conditions")
            s.w = max(0.0, w_in - credit)

        elif code == "OVERFILL":
            _overfill(s, m, sd, sv)

        elif code == "BOILUP":
            _boilup(s, m, sd, sv)

        elif code == "ABHEAT":
            _abheat(s, m, sd, sv)

        elif code == "RUNAWAY":
            q = sv("RUNAWAY: Heat generation at runaway", "Q") or 0.0
            diers = sv("RUNAWAY: DIERS vapor rate (if available)", "mflow")
            typ = txt(scen.get("RUNAWAY: System type")) or "Tempered"
            lam = s.latent or latent_from_feed(sd.feed, s.p1)
            if diers:
                s.w = diers
            elif typ.lower().startswith("temper") and lam:
                s.w = q / lam
                s.notes.append("tempered: W = Q_rxn/λ")
            else:
                s.notes.append("gassy system — DIERS rate required")
            s.phase = "Two-Phase"
            s.quality = min(1.0, max(0.01, s.quality if s.quality < 1 else 0.1))
            s.notes.append("two-phase venting assumed (DIERS)")

        elif code == "OTHER":
            s.w = sv("OTHER: Relieving rate", "mflow") or 0.0

        out.append(s)
    return out


def _fire(s: Scen, m: Model, sd: StreamData, sv):
    eq, gen = m.eq, m.gen
    area = eq.get("a_override")
    if area is None:
        area = wetted_area_m2(eq["orient"], eq.get("od") or 0.0,
                              eq.get("tt") or 0.0, eq["head"],
                              eq.get("nll") or 0.0, eq["grade_el"],
                              eq["fire_lim"], eq["extra_a"])
        s.notes.append(f"wetted area (calc) = {area:.2f} m²")
    f = fire_env_factor(eq["drain"], eq["ff"], eq["insul"], eq.get("f_override"))
    q = sv("FIRE: Manual Q override", "Q") or fire_q_w(area, f)
    s.notes.append(f"F = {f:.3f}; Q = {q/1000.0:.1f} kW")
    lam = s.latent or latent_from_feed(sd.feed, s.p1)
    if sd.feed is not None:
        tb = bubble_dew_t_k(sd.feed, s.p1, "bubble")
        if tb:
            s.t = tb
            s.notes.append("relieving T = bubble point at accumulation")
    w_vap = sv("FIRE: Manual relief rate override", "mflow")
    if w_vap is None:
        w_vap = q / lam if lam else 0.0
        if not lam:
            s.notes.append("no λ — enter latent heat or relief rate manually")
    s.w = w_vap
    s.phase = "Vapor"
    # §3.10 vapor-only screening
    dia = eq.get("od") or 0.0
    fill = eq.get("fill_pct")
    set_g = (gen.get("set_p") or ATM_PA) - ATM_PA
    if fill is not None and dia > 0:
        ok, mult, note = fire_vapor_only_ok(dia, fill, eq["foaming"],
                                            set_g, eq["orient"])
        s.notes.append(f"vapor-only screen: {note}")
        s.area_mult = mult
        if not ok and mult == 1.0:
            w2_manual = sv("FIRE: Manual two-phase W2 override", "mflow")
            vf = 1.0 / (s.rho_l or 700.0)
            vg = 1.0 / max(1e-6, _rho_v_ideal(s))
            dt = 0.0
            if sd.feed is not None:
                t_set = bubble_dew_t_k(sd.feed, gen.get("set_p") or s.p1, "bubble")
                t_acc = bubble_dew_t_k(sd.feed, s.p1, "bubble")
                if t_set and t_acc:
                    dt = max(0.0, t_acc - t_set)
            alpha = max(0.0, 1.0 - (fill or 95.0) / 100.0)
            if w2_manual:
                s.w, x2 = w2_manual, max(0.01, min(1.0, s.quality))
                s.notes.append("manual two-phase W2 used")
            else:
                w2, x2, note2 = fire_twophase_w2(
                    w_vap, q, s.cp_l or 2400.0, dt, alpha, vf, vg,
                    lam or 2.0e5)
                s.w = w2
                s.notes.append(f"two-phase swell: {note2}")
            if s.w > w_vap:
                s.phase = "Two-Phase"
                s.quality = x2


def _cv_failure(s: Scen, m: Model, sd: StreamData):
    uio, cvd = m.uio, m.cv
    mode = (txt(m.scen.get("CV: Evaluate as")) or "Design").upper()
    credit = (uio.si("mflow", num(m.scen.get("CV: Outflow credit"))) or 0.0)
    best = 0.0
    for i in (1, 2):
        pre = f"CV{i}: "
        rated = num(cvd.get(pre + "Rated Cv (full open)"))
        if not rated:
            continue
        normal = num(cvd.get(pre + "Normal operating Cv")) or 0.0
        bypass = num(cvd.get(pre + "Bypass rated Cv")) or 0.0
        cv_eff = rated + (bypass if mode.startswith("REMOTE")
                          else 0.5 * normal)
        p1 = uio.si("P", num(cvd.get(pre + "Max upstream pressure"))) or 0.0
        t1 = uio.si("T", num(cvd.get(pre + "Upstream temperature"))) or s.t
        phase = (txt(cvd.get(pre + "Phase")) or "Gas").upper()
        fl = num(cvd.get(pre + "FL")) or 0.9
        xt = num(cvd.get(pre + "xT")) or 0.7
        if phase.startswith("L"):
            w = cv_liquid_kgs(cv_eff, p1, s.p1, s.rho_l or 800.0, fl)
        else:
            w = cv_gas_kgs(cv_eff, p1, s.p1, t1, s.mw, s.z, s.k, xt)
        s.notes.append(f"CV{i}: Cv_eff={cv_eff:.1f} ({mode.title()}) "
                       f"→ {w*3600:.0f} kg/h")
        best = max(best, w)
    s.w = max(0.0, best - credit)
    if mode.startswith("REMOTE") and not s.klass.upper().startswith("REMOTE"):
        s.notes.append("⚠ evaluated as Remote but class is Design — check class")


def _gas_blowby(s: Scen, m: Model, sd: StreamData):
    uio, gd = m.uio, m.gas
    typ = (txt(gd.get("GB1: Restriction type")) or "Cv").upper()
    p1 = uio.si("P", num(gd.get("GB1: Max source pressure"))) or 0.0
    t1 = uio.si("T", num(gd.get("GB1: Source temperature"))) or s.t
    cd = num(gd.get("GB1: Discharge coefficient Cd")) or 0.61
    if typ.startswith("CV"):
        cv = num(gd.get("GB1: Cv (full open)")) or 0.0
        xt = num(gd.get("GB1: xT")) or 0.7
        s.w = cv_gas_kgs(cv, p1, s.p1, t1, s.mw, s.z, s.k, xt)
        s.notes.append(f"blowby through full-open Cv={cv:.0f}")
    elif typ.startswith("OR"):
        a = uio.si("Ain", num(gd.get("GB1: Orifice area"))) or 0.0
        s.w = orifice_gas_kgs(a, cd, p1, s.p1, t1, s.mw, s.z, s.k)
        s.notes.append("blowby through fixed orifice")
    else:                                           # check-valve reverse flow
        id_in = num(gd.get("GB1: Check valve min internal diameter")) or 0.0
        nv = int(num(gd.get("GB1: Number of check valves in series")) or 1)
        pct_hp = (num(gd.get("GB1: Partial-failure orifice — HP valve")) or 6.0)
        pct_lp = (num(gd.get("GB1: Partial-failure orifice — LP valve")) or 2.0)
        d_m = id_in * 0.0254
        a_hp = math.pi / 4.0 * (d_m * pct_hp / 100.0) ** 2
        a_lp = math.pi / 4.0 * (d_m * pct_lp / 100.0) ** 2
        a_eff = a_hp if nv == 1 else min(a_hp, a_lp)
        s.w = orifice_gas_kgs(a_eff, cd, p1, s.p1, t1, s.mw, s.z, s.k)
        s.notes.append(f"§3.7 reverse flow: {nv} valve(s), governing orifice "
                       f"d = {pct_lp if nv > 1 else pct_hp:.0f}% of "
                       f"{id_in:.2f} in ID (series approximated by the "
                       f"smaller orifice)")
    if (s.phase or "").startswith("Liq"):
        s.phase = "Vapor"
        s.notes.append("gas blowby sized as vapor at PSV")


def _tube(s: Scen, m: Model, sd: StreamData, sv):
    eq, uio = m.eq, m.uio
    leak = yes(m.scen.get("TUBE: Evaluate 6 mm leak instead of rupture"))
    nfail = int(sv("TUBE: Number of tube failures") or 1)
    hp_p = (sv("TUBE: HP-side max pressure override", "P")
            or (eq.get("tube_mawp") if eq["hp_side"].upper().startswith("T")
                else eq.get("shell_mawp")) or 0.0)
    if leak:
        a = math.pi / 4.0 * 0.006 ** 2
        n_holes = 1
        s.notes.append("tube LEAK basis: single 6 mm hole (§3.11)")
    else:
        od = (eq.get("tube_od_in") or 0.75) * 0.0254
        wall = (eq.get("tube_wall_in") or 0.083) * 0.0254
        di = max(1e-4, od - 2.0 * wall)
        a = math.pi / 4.0 * di ** 2
        n_holes = 2 * nfail                      # both ends of each broken tube
        s.notes.append(f"tube RUPTURE: {nfail} tube(s) × both ends")
    cd = 0.62
    if (s.phase or "Vapor").startswith("Liq"):
        s.w = n_holes * orifice_liq_kgs(a, cd, hp_p, s.p1, s.rho_l or 800.0)
        s.notes.append("HP fluid liquid — check flashing across the breach")
    else:
        s.w = n_holes * orifice_gas_kgs(a, cd, hp_p, s.p1, s.t, s.mw, s.z, s.k)
    dp_psi = (hp_p - s.p1) / PSI_TO_PA
    if dp_psi >= 1000.0 and not leak:
        s.notes.append("⚠ ΔP ≥ 1000 psi: check 10-tube multiple-failure case "
                       "and LP piping surge (§3.11)")
    elif dp_psi >= 750.0 and not leak:
        s.notes.append("⚠ ΔP ≥ 750 psi: transient (surge) analysis required "
                       "for liquid-continuous LP outlets (§3.11)")


def _pinched_duty(m: Model, sd: StreamData, p1: float, q_base: float,
                  notes: list) -> float:
    scen, uio = m.scen, m.uio
    if not yes(scen.get("PINCH: Apply pinch credit?")):
        return q_base
    ua_raw = num(scen.get("PINCH: U × A"))
    t_med = uio.si("T", num(scen.get("PINCH: Heating medium temperature")))
    t_bub = uio.si("T", num(scen.get("PINCH: Reboiler-feed bubble point at "
                                     "accumulation")))
    if t_bub is None and sd.feed is not None:
        t_bub = bubble_dew_t_k(sd.feed, p1, "bubble")
    if ua_raw is None or t_med is None or t_bub is None:
        notes.append("pinch requested but UA / T_medium / T_bubble missing — "
                     "no credit taken")
        return q_base
    ua_si = (uio.si("Q", ua_raw) or 0.0) * (1.8 if uio.u.sel.get("dT") == "degF"
                                            else 1.0)
    if t_bub >= t_med:
        notes.append("COMPLETE reboiler pinch: bubble point ≥ medium T — "
                     "no boil-up for this contingency")
        return 0.0
    q_p = ua_si * (t_med - t_bub)
    if q_p < q_base:
        notes.append(f"partial pinch: duty {q_base/1000:.0f} → {q_p/1000:.0f} kW")
        return q_p
    return q_base


def _boilup(s: Scen, m: Model, sd: StreamData, sv):
    q = sv("BOILUP: Reboiler duty", "Q") or 0.0
    q = _pinched_duty(m, sd, s.p1, q, s.notes)
    if q <= 0:
        s.w = 0.0
        return
    lam = sv("BOILUP: Manual λ override", "h")
    if lam is None:
        lam = latent_from_feed(sd.feed, s.p1) or s.latent
        if lam and sd.feed is not None:
            s.notes.append("λ from two-flash (bubble→dew) of linked feed (§5.7)")
    if not lam:
        s.notes.append("no λ available — enter BOILUP λ override")
        return
    s.w = q / lam
    s.phase = "Vapor"
    if sd.feed is not None:
        td = bubble_dew_t_k(sd.feed, s.p1, "dew")
        if td:
            s.t = td
            s.notes.append("relieving T = dew point at accumulation")
        s.mw = sd.feed.total_mw or s.mw


def _overfill(s: Scen, m: Model, sd: StreamData, sv):
    feed = sv("OVERFILL: Feed rate", "mflow") or 0.0
    credit = sv("OVERFILL: Continuing liquid outflow credit", "mflow") or 0.0
    reb_v = sv("OVERFILL: Max reboiler vapor rate", "mflow") or 0.0
    feed_v = sv("OVERFILL: Max feed vapor rate", "mflow") or 0.0
    s.w = max(0.0, feed - credit)
    if s.w <= 0:
        s.notes.append("net overfill rate ≤ 0 — check credits")
        return
    x = min(1.0, (reb_v + feed_v) / s.w) if s.w > 0 else 0.0
    s.quality = x
    s.phase = "Liquid" if x <= 1e-4 else "Two-Phase"
    s.notes.append(f"relief quality assembled per §5.6: x = {x:.3f} "
                   "(max reboiler vapor + max feed vapor)")
    if sd.feed is not None:
        tb = bubble_dew_t_k(sd.feed, s.p1, "bubble")
        if tb:
            s.t = tb
    # DPeq check (§3.4) for associated equipment
    gen = m.gen
    dpfric = sv("OVERFILL: Friction dP equipment→PSV at overfill flow", "dP") or 0.0
    rho = s.rho_l or gen.get("rho_head") or 999.0
    el_max, el_psv = gen.get("el_max"), gen.get("el_psv")
    if el_max is not None and el_psv is not None and gen.get("set_p"):
        lh = rho * G_STD * (el_psv - (gen.get("el_min") or 0.0))
        dpeq = ((gen["set_p"] - ATM_PA) * gen["htf"]) - lh - dpfric
        s.notes.append(f"DPeq (min reqd design P of associated equipment) = "
                       f"{m.uio.user('dP', dpeq)} {m.uio.label('dP')} (g) — "
                       "verify against the EQUIPMENT table")


def _abheat(s: Scen, m: Model, sd: StreamData, sv):
    scen = m.scen
    typ = (txt(scen.get("ABHEAT: Reboiler type")) or "Steam").upper()
    q0 = sv("ABHEAT: Pre-contingency duty Q0", "Q") or 0.0
    base = sv("ABHEAT: Base overhead vapor rate", "mflow") or 0.0
    cond = (txt(scen.get("ABHEAT: Condenser type")) or "Total").upper()
    drum = sv("ABHEAT: Pre-contingency drum vapor rate (partial)", "mflow") or 0.0
    if typ.startswith("FIRED"):
        qd = sv("ABHEAT: Design duty (fired only)", "Q") or q0
        qc = 1.25 * qd
        s.notes.append("fired reboiler: Qc = 1.25 × design duty (§5.4); "
                       "furnace reboilers do not pinch (§5.1)")
    else:
        cvn = num(scen.get("ABHEAT: CV normal operating Cv")) or 0.0
        cvf = num(scen.get("ABHEAT: CV rated full-open Cv")) or 0.0
        ratio_cv = (cvf / cvn) if cvn > 0 else 1.0
        qc = q0 * ratio_cv
        s.notes.append(f"medium CV wide open: m_c/m₀ = Cv_full/Cv_norm = "
                       f"{ratio_cv:.2f} (same P1/P2 — flow ∝ Cv)")
        qc = _pinched_duty(m, sd, s.p1, qc, s.notes)
    ratio = qc / q0 if q0 > 0 else 1.0
    scaled = base * ratio
    if cond.startswith("PART"):
        s.w = max(0.0, scaled + drum - base)
        s.notes.append("partial condenser flooded: relief = scaled + drum "
                       "vapor − pre-contingency overhead (§5.4)")
    else:
        s.w = max(0.0, scaled - base)
        s.notes.append("total condenser: relief = scaled − pre-contingency "
                       "overhead (§5.4)")
    if sd.feed is not None and HYD is not None:
        fr = flash_adiabatic(sd.feed, s.p1)
        if fr.temp_f is not None:
            s.t = (fr.temp_f - 32.0) / 1.8 + 273.15
        s.mw = fr.vap_mw or s.mw
    s.phase = "Vapor"


# ═══════════════════════════════════════════════════════════════════════════
#  Sizing, piping & compliance
# ═══════════════════════════════════════════════════════════════════════════
def size_scenario(s: Scen, m: Model, kb: float) -> float:
    cfg = m.cfg
    # drop sizing notes from a previous iteration (BP loop re-sizes)
    s.notes = [n for n in s.notes
               if not n.startswith(("Leung ω", "Kv (visc", "area multiplier"))]
    if s.w <= 0 or s.p1 <= 0:
        s.area = 0.0
        return 0.0
    ph = (s.phase or "Vapor").upper()
    if ph.startswith("LIQ"):
        kw = (kw_bellows(100.0 * (s.p2 - ATM_PA) / max(1.0, s.p1 - ATM_PA))
              if "bellow" in cfg["type"].lower() else 1.0)
        a0 = size_liquid(s.w, s.rho_l or 800.0, s.p1, s.p2,
                         cfg["kd_l"], kw, cfg["kc"], 1.0)
        kv = kv_viscosity(a0, s.rho_l or 800.0, s.mu or 1.0, s.w)
        s.area = size_liquid(s.w, s.rho_l or 800.0, s.p1, s.p2,
                             cfg["kd_l"], kw, cfg["kc"], kv)
        s.kw_bp, s.kv_visc = kw, kv
        if kv < 0.999:
            s.notes.append(f"Kv (viscosity) = {kv:.3f}")
    elif ph.startswith("TWO"):
        rho_v = _rho_v_ideal(s)
        s.area, s.omega = size_twophase(
            s.w, s.p1, s.p2, max(1e-4, min(0.9999, s.quality)),
            s.rho_l or 700.0, rho_v, s.cp_l or 2400.0, s.t,
            s.latent or 2.0e5, cfg["kd_2"], cfg["kc"])
        s.notes.append(f"Leung ω = {s.omega:.3f}")
    else:
        s.area = size_vapor(s.w, s.t, s.mw, s.z, s.k, s.p1, s.p2,
                            cfg["kd_v"], kb, cfg["kc"])
    s.area *= s.area_mult
    if s.area_mult != 1.0:
        s.notes.append(f"area multiplier ×{s.area_mult:g} applied "
                       "(§3.10 two-phase approximation)")
    return s.area


def simple_line_dp(kv: dict, uio: UIO, w_kgs: float, rho: float,
                   mu_cp: float) -> tuple[float, dict]:
    """Darcy single-line dP [Pa] from an INLET/OUTLET piping sheet.

    Returns ``(dp_pa, info)``.  ``info["mode"]`` is ``"manual"`` (override
    cell used), ``"calc"`` (line model — info carries v/Re/f/geometry for
    display), or ``"none"`` (no line diameter / no flow — dP = 0).
    """
    manual = uio.si("dP", num(kv.get("Manual dP override at relief flow")))
    if manual is not None:
        return manual, {"mode": "manual"}
    d_in = num(kv.get("Line inside diameter"))
    if not d_in or w_kgs <= 0 or rho <= 0:
        return 0.0, {"mode": "none", "d_in": d_in}
    d = d_in * 0.0254
    a = math.pi / 4.0 * d * d
    v = w_kgs / (rho * a)
    length = uio.si("L", num(kv.get("Straight length"))) or 0.0
    ktot = num(kv.get("Total fittings K")) or 0.0
    dz = uio.si("L", num(kv.get("Elevation change"))) or 0.0
    re = rho * v * d / ((mu_cp or 0.3) * 1e-3)
    if re < 2100:
        f = 64.0 / max(1.0, re)
    else:
        f = 0.25 / (math.log10(1.524e-3 / (3.7 * d * 1000.0)
                               + 5.74 / re ** 0.9)) ** 2
    dp = max(0.0, (f * length / d + ktot) * 0.5 * rho * v * v + rho * G_STD * dz)
    return dp, {"mode": "calc", "d_in": d_in, "v": v, "re": re, "f": f,
                "length": length, "ktot": ktot, "dz": dz, "rho": rho,
                "fittings": kv.get("_fittings", [])}


def size_case(s: Scen, m: Model, super_g: float, remote_allow_g: float) -> None:
    """Full per-case sizing: iterate built-up back pressure from the OUTLET
    line at THIS case's relief load/properties, then evaluate the INLET
    line at the same flow.  Populates the case-result fields on ``s`` so
    every active scenario gets its own back-pressure + inlet/outlet line
    sizing (written to its own CASE sheet)."""
    cfg, gen, press = m.cfg, m.gen, m.press
    set_g = (gen.get("set_p") or ATM_PA) - ATM_PA
    op_pct = gen.get("op_pct") or 10.0
    kb_user = press.get("kb_user")

    kb = kb_factor(cfg["type"], 100.0 * super_g / max(1.0, set_g), op_pct, kb_user)
    s.p2 = ATM_PA + super_g
    size_scenario(s, m, kb)

    bp_g, outlet_info = super_g, {"mode": "none"}
    if s.w > 0 and s.area > 0:
        for _ in range(8):
            rho_out = (_rho_v_ideal(s) if not s.phase.startswith("Liq")
                       else (s.rho_l or 800.0))
            dp_out, outlet_info = simple_line_dp(m.outlet, m.uio, s.w, rho_out, s.mu)
            bp_g = super_g + dp_out
            kb_new = kb_factor(cfg["type"], 100.0 * bp_g / max(1.0, set_g),
                               op_pct, kb_user)
            s.p2 = ATM_PA + bp_g
            size_scenario(s, m, kb_new)
            if abs(kb_new - kb) < 1e-4:
                kb = kb_new
                break
            kb = kb_new
    s.kb, s.builtup_pa, s.total_bp_g = kb, bp_g - super_g, bp_g
    s.outlet_info = outlet_info

    if s.w > 0:
        rho_in = (_rho_v_ideal(s) if not s.phase.startswith("Liq")
                  else (s.rho_l or 800.0))
        s.inlet_dp_pa, s.inlet_info = simple_line_dp(m.inlet, m.uio, s.w, rho_in, s.mu)
    else:
        s.inlet_dp_pa, s.inlet_info = 0.0, {"mode": "none"}

    s.orifice, s.orifice_area = (select_orifice(s.area) if s.area > 0
                                  else ("—", 0.0))

    s.bp_lim_pct = (press["bp_pilot"] if "pilot" in cfg["type"].lower() else
                    press["bp_bell"] if "bellow" in cfg["type"].lower() else
                    press["bp_conv"])
    s.bp_pct = 100.0 * s.total_bp_g / set_g if set_g > 0 else 0.0
    s.in_pct = 100.0 * s.inlet_dp_pa / set_g if set_g > 0 else 0.0
    s.bp_ok = s.bp_pct <= s.bp_lim_pct
    s.in_ok = s.in_pct <= 3.0
    mawp_g = (gen.get("mawp") or 0.0) - ATM_PA
    s.accum_ok = (s.allowable_g <= max(remote_allow_g, 1.21 * mawp_g) + 1.0
                  if mawp_g > 0 else True)


def static_head_correction(gen: dict) -> tuple[float | None, float | None, list]:
    """Returns (head_pa, corrected_set_pa_abs, notes)."""
    notes = []
    rho = gen.get("rho_head")
    el_max, el_psv = gen.get("el_max"), gen.get("el_psv")
    if rho is None or el_max is None or el_psv is None or not gen.get("set_p"):
        return None, None, ["static correction skipped — elevation/density "
                            "inputs incomplete"]
    head = rho * G_STD * (el_max - el_psv)       # >0 when PSV below liquid top
    corrected = gen["set_p"] - head
    if head < 0:
        notes.append("PSV ABOVE the protected liquid top: the top equipment "
                     "sees MORE than set pressure at lift — §5.6 flags this "
                     "as unacceptable for overfill; verify set P / elevation")
    return head, corrected, notes


def disposal_verdict(d: dict) -> tuple[str, list]:
    Y = lambda lbl: yes(d.get(lbl))
    reasons = []
    if Y("Liquid or partially liquid at PSV inlet?") and not \
            Y("Benign water-type liquid < 65°C?"):
        reasons.append("liquid/partially-liquid inlet")
    if Y("Flammable vapor MW > 80 or condenses at ambient?"):
        reasons.append("condensable flammable vapor (MW > 80)")
    if Y("Toxic service (AEGL limits at platforms/fence)?"):
        reasons.append("toxic — AEGL limits")
    if Y("Corrosive condensables (e.g. phenol)?"):
        reasons.append("corrosive condensables (rain-out)")
    if Y("> 50% LFL at grade/platforms/fence?"):
        reasons.append("> 50% LFL")
    if Y("Radiant flux > 9.5 kW/m² if ignited?"):
        reasons.append("radiant flux > 9.5 kW/m²")
    if Y("Operating above auto-ignition at PSV inlet?"):
        reasons.append("above auto-ignition temperature")
    if Y("≥ 50 mol% H2 or T > min(AIT, 315°C)?"):
        reasons.append("H2-rich / hot service (tightened criteria)")
    if Y("Overfill is a DESIGN contingency on this vessel?") and not \
            Y("Benign water-type liquid < 65°C?"):
        reasons.append("vapor-space PSV with design-contingency overfill")
    if reasons:
        verdict = "FLARE (required)"
        if Y("H2S-containing?"):
            verdict += " — HC flare, never the acid-gas flare (§7.3)"
        return verdict, reasons
    return ("ATMOSPHERE permissible — verify §7.2: vapor-only, dispersion, "
            "radiation, exit velocity ≥ 30 m/s at 25% capacity, ≤ 75% sonic"),\
           ["no flare trigger answered Y"]


@dataclass
class RunResult:
    scens: list = field(default_factory=list)
    governing: Scen | None = None
    orifice: str = ""
    orifice_area: float = 0.0
    kb: float = 1.0
    builtup_pa: float = 0.0
    total_bp_g: float = 0.0
    inlet_dp_pa: float = 0.0
    head_pa: float | None = None
    corrected_set_pa: float | None = None
    remote_allow_g: float = 0.0
    remote_note: str = ""
    checks: list = field(default_factory=list)
    disposal: str = ""
    disposal_reasons: list = field(default_factory=list)
    head_notes: list = field(default_factory=list)


def run_model(m: Model) -> RunResult:
    r = RunResult()
    gen, press, cfg = m.gen, m.press, m.cfg
    set_g = (gen.get("set_p") or ATM_PA) - ATM_PA
    r.remote_allow_g, r.remote_note = _remote_allowable_g(gen)
    r.head_pa, r.corrected_set_pa, r.head_notes = static_head_correction(gen)

    scens = compute_scenarios(m)
    r.scens = scens

    # per-case sizing: every active scenario gets its own built-up
    # back-pressure iteration (own outlet-line dP) and inlet-line dP
    super_g = press["super_const"]
    active = [s for s in scens if s.active]
    for s in active:
        size_case(s, m, super_g, r.remote_allow_g)

    gov = max(active, key=lambda s: s.area, default=None)
    r.governing = gov
    if gov and gov.area > 0:
        r.kb, r.builtup_pa, r.total_bp_g = gov.kb, gov.builtup_pa, gov.total_bp_g
        r.inlet_dp_pa = gov.inlet_dp_pa
        r.orifice, r.orifice_area = gov.orifice, gov.orifice_area
    else:
        r.kb = kb_factor(cfg["type"], 100.0 * super_g / max(1.0, set_g),
                         gen.get("op_pct") or 10.0, press.get("kb_user"))

    # ── compliance checks ─────────────────────────────────────────────────
    mawp_g = (gen.get("mawp") or 0.0) - ATM_PA
    ck = r.checks
    if gen.get("set_p") and gen.get("mawp"):
        ck.append(("Set P ≤ MAWP", set_g <= mawp_g + 1.0,
                   f"set {m.uio.user('P', gen['set_p'])} vs MAWP "
                   f"{m.uio.user('P', gen['mawp'])} {m.uio.label('P')}"))
    for s in active:
        if mawp_g > 0:
            ck.append((f"{s.code}: allowable within class limit", s.accum_ok,
                       f"{s.klass}: {m.uio.user('dP', s.allowable_g)} "
                       f"{m.uio.label('dP')} (g)"))
    bp_lim_pct = gov.bp_lim_pct if gov else (
        press["bp_pilot"] if "pilot" in cfg["type"].lower() else
        press["bp_bell"] if "bellow" in cfg["type"].lower() else
        press["bp_conv"])
    if set_g > 0 and gov:
        ck.append((f"Total BP ≤ {bp_lim_pct:.0f}% of set ({cfg['type']})",
                   gov.bp_ok, f"total BP = {gov.bp_pct:.1f}% of set"))
        ck.append(("Inlet dP ≤ 3% of set", gov.in_ok,
                   f"inlet dP = {gov.in_pct:.2f}% of set"))
    if r.corrected_set_pa is not None:
        des_ps = [e["des_p"] for e in m.equip_rows if e.get("des_p")]
        if des_ps:
            ok = min(des_ps) >= (gen.get("set_p") or 0.0)
            ck.append(("Min equipment design P ≥ set P", ok,
                       f"min design P = {m.uio.user('P', min(des_ps))} "
                       f"{m.uio.label('P')}"))

    r.disposal, r.disposal_reasons = disposal_verdict(m.disposal)
    return r


# ═══════════════════════════════════════════════════════════════════════════
#  Output workbook
# ═══════════════════════════════════════════════════════════════════════════
def _sizing_rows(s: Scen, m: Model) -> list:
    """Phase-specific API 520 sizing-detail rows for a per-case sheet."""
    uio, cfg = m.uio, m.cfg
    U, L = uio.user, uio.label
    rows = []
    ph = (s.phase or "Vapor").upper()
    if ph.startswith("LIQ"):
        rows += [
            ("Differential pressure (P1 − P2)", U("dP", s.p1 - s.p2), L("dP"), ""),
            ("Liquid density", U("rho", s.rho_l), L("rho"), ""),
            ("Viscosity", round(s.mu, 4), L("visc"), ""),
            ("Kd (discharge coeff., liquid)", cfg["kd_l"], "", ""),
            ("Kw (back-pressure factor)", round(s.kw_bp, 4), "", ""),
            ("Kv (viscosity correction)", round(s.kv_visc, 4), "", ""),
            ("Kc (combination factor)", cfg["kc"], "", ""),
        ]
    elif ph.startswith("TWO"):
        rho_v = _rho_v_ideal(s)
        rows += [
            ("Quality x (mass vapor fraction)", round(s.quality, 4), "", ""),
            ("Vapor density (relieving)", U("rho", rho_v), L("rho"), ""),
            ("Liquid density", U("rho", s.rho_l), L("rho"), ""),
            ("Latent heat", U("h", s.latent), L("h"), "" if s.latent
             else "not provided — default used"),
            ("Leung ω (Annex C)",
             round(s.omega, 4) if s.omega is not None else None, "", ""),
            ("Kd (discharge coeff., two-phase)", cfg["kd_2"], "", ""),
            ("Kc (combination factor)", cfg["kc"], "", ""),
        ]
    else:
        g, choked = gas_mass_flux(s.p1, s.p2, s.t, s.mw, s.z, s.k)
        rows += [
            ("Molecular weight", round(s.mw, 2), L("MW"), ""),
            ("Compressibility Z", round(s.z, 3), "", ""),
            ("k = Cp/Cv", round(s.k, 3), "", ""),
            ("Mass flux G", round(g, 2), "kg/(m²·s)", ""),
            ("Flow regime", "Choked (critical)" if choked else "Subcritical",
             "", f"P2/P1 = {(s.p2 / s.p1):.3f}" if s.p1 > 0 else ""),
            ("Kd (discharge coeff., vapor)", cfg["kd_v"], "", ""),
            ("Kb (back-pressure correction)", round(s.kb, 4), "", ""),
            ("Kc (combination factor)", cfg["kc"], "", ""),
        ]
    if s.area_mult != 1.0:
        rows.append(("Area multiplier (two-phase swell, §3.10)",
                      f"×{s.area_mult:g}", "", ""))
    return rows


def _line_rows(prefix: str, info: dict) -> list:
    """Render velocity/Re/f/geometry rows for a calculated INLET/OUTLET line,
    or a note when an override was used / the line is unsized."""
    if info.get("mode") == "manual":
        return [(f"{prefix} line dP", "manual override", "",
                 "from the piping sheet's override cell")]
    if info.get("mode") == "calc":
        rows = [
            (f"{prefix} line inside diameter", info.get("d_in"), "in", ""),
            (f"{prefix} line velocity", round(info["v"], 3), "m/s", ""),
            (f"{prefix} line Reynolds number", f"{info['re']:.0f}", "", ""),
            (f"{prefix} line friction factor f", round(info["f"], 5), "", ""),
            (f"{prefix} line length / ΣK / Δz",
             f"{info['length']:.2f} m / {info['ktot']:.2f} / {info['dz']:.2f} m",
             "", ""),
        ]
        fittings = info.get("fittings") or []
        if fittings:
            rows.append((f"{prefix} line fittings breakdown", "", "", "sec"))
            for name, qty, k in fittings:
                rows.append((f"  {name} (qty {qty:g}, K = {k:g} each)",
                              round(k * qty, 3), "K", ""))
        return rows
    return [(f"{prefix} line dP", "not evaluated", "",
             f"no line size or no flow — {prefix.upper()}_PIPING incomplete")]


def _write_case_sheet(wb, m: Model, r: RunResult, s: Scen) -> None:
    """One results sheet per active scenario: relieving conditions, sizing,
    back-pressure (this case's own outlet-line dP/Kb iteration) and inlet-line
    sizing (this case's own flow through the shared INLET_PIPING line)."""
    uio = m.uio
    U, L = uio.user, uio.label
    ws = wb.create_sheet(f"CASE_{s.code}"[:31])
    title = f"{s.code} — {s.name}"
    if r.governing is s:
        title += "   ★ GOVERNING CASE"
    _title(ws, 4, f"{title}   |   units: {uio.system}")

    rows = [("Relieving Conditions", "", "", "sec"),
            ("Class", s.klass, "", ""),
            ("Allowable accumulation pressure (g)",
             U("dP", s.allowable_g), L("dP"), ""),
            ("Relieving pressure P1 (abs)", U("P", s.p1), L("P"), ""),
            ("Relieving temperature", U("T", s.t, 1), L("T"), ""),
            ("Phase", s.phase, "", "")]
    if (s.phase or "").upper().startswith("TWO"):
        rows.append(("Quality x", round(s.quality, 4), "",
                      "mass vapor fraction"))
    rows.append(("Molecular weight", round(s.mw, 2), L("MW"), ""))
    rows.append(("Compressibility Z", round(s.z, 3), "", ""))
    rows.append(("k = Cp/Cv", round(s.k, 3), "", ""))
    if s.rho_v:
        rows.append(("Vapor density", U("rho", s.rho_v), L("rho"), ""))
    if s.rho_l:
        rows.append(("Liquid density", U("rho", s.rho_l), L("rho"), ""))
    rows.append(("Viscosity", round(s.mu, 4), L("visc"), ""))
    if s.latent:
        rows.append(("Latent heat", U("h", s.latent), L("h"), ""))

    rows.append(("Relief Load", "", "", "sec"))
    rows.append(("Relieving mass flow W", U("mflow", s.w, 2), L("mflow"), ""))

    rows.append(("Sizing (API 520 Part I)", "", "", "sec"))
    rows += _sizing_rows(s, m)
    rows.append(("Required orifice area", U("Ain", s.area, 5), L("Ain"), ""))
    rows.append(("Selected orifice (API 526)", s.orifice, "",
                  f"actual {U('Ain', s.orifice_area, 4)} {L('Ain')}"))

    rows.append((f"Back Pressure — Outlet Line ({s.code} flow)", "", "", "sec"))
    rows.append(("Superimposed back pressure (constant)",
                  U("dP", m.press["super_const"]), L("dP"), ""))
    rows += _line_rows("Outlet", s.outlet_info)
    rows.append(("Built-up back pressure (this case)",
                  U("dP", s.builtup_pa), L("dP"), ""))
    rows.append(("Total back pressure (g)", U("dP", s.total_bp_g), L("dP"), ""))
    rows.append(("Kb (back-pressure correction)", round(s.kb, 4), "", ""))
    rows.append((f"Total BP vs {s.bp_lim_pct:.0f}% limit ({m.cfg['type']})",
                  "PASS" if s.bp_ok else "FAIL", "", f"{s.bp_pct:.1f}% of set"))

    rows.append((f"Inlet Line ({s.code} flow) — 3% Rule", "", "", "sec"))
    rows += _line_rows("Inlet", s.inlet_info)
    rows.append(("Inlet line pressure drop", U("dP", s.inlet_dp_pa), L("dP"), ""))
    rows.append(("Inlet dP vs 3% limit", "PASS" if s.in_ok else "FAIL", "",
                  f"{s.in_pct:.2f}% of set"))

    rows.append(("Accumulation Check", "", "", "sec"))
    rows.append(("Allowable within class limit",
                  "PASS" if s.accum_ok else "FAIL", "",
                  f"{s.klass}: {U('dP', s.allowable_g)} {L('dP')} (g)"))

    if s.notes:
        rows.append(("Notes / Basis", "", "", "sec"))
        for n in s.notes:
            rows.append((n, "", "", ""))

    _result_rows(ws, 3, rows, widths=(36, 24, 10, 56))


def write_output(m: Model, r: RunResult, out_path: str) -> str:
    uio = m.uio
    U = uio.user
    L = uio.label
    wb = Workbook()

    ws = wb.active
    ws.title = "SUMMARY"
    _title(ws, 4, f"PSV SIZING SUMMARY — {m.gen.get('tag')}   |   "
                  f"units: {uio.system}   |   {datetime.now():%d-%b-%Y %H:%M}")
    r0 = 3
    rows = [
        ("Identification", "", "", "sec"),
        ("PSV Tag", m.gen.get("tag"), "", ""),
        ("Unit / P&ID", f"{m.gen.get('unit') or ''} / {m.gen.get('pid') or ''}", "", ""),
        ("Service", m.gen.get("service"), "", ""),
        ("Pressures", "", "", "sec"),
        ("Set pressure", U("P", m.gen.get("set_p")), L("P"), ""),
        ("MAWP", U("P", m.gen.get("mawp")), L("P"), ""),
        ("Static head (max equip → PSV)", U("dP", r.head_pa), L("dP"),
         "ρ·g·(EL_max − EL_PSV)"),
        ("Corrected set pressure", U("P", r.corrected_set_pa), L("P"),
         "Set P − static head (PSV Inputs plus)"),
        ("Remote-contingency allowable (g)", U("dP", r.remote_allow_g), L("dP"),
         r.remote_note),
        ("Governing Result", "", "", "sec"),
        ("Governing scenario",
         f"{r.governing.code} — {r.governing.name}" if r.governing else "—",
         "", ""),
        ("Relieving rate", U("mflow", r.governing.w) if r.governing else None,
         L("mflow"), ""),
        ("Required orifice area",
         U("Ain", r.governing.area) if r.governing else None, L("Ain"), ""),
        ("Selected orifice (API 526)", r.orifice, "",
         f"actual {U('Ain', r.orifice_area)} {L('Ain')}"),
        ("Kb (back-pressure)", round(r.kb, 4), "", ""),
        ("Built-up back pressure", U("dP", r.builtup_pa), L("dP"), ""),
        ("Total back pressure (g)", U("dP", r.total_bp_g), L("dP"), ""),
        ("Inlet pressure drop", U("dP", r.inlet_dp_pa), L("dP"), ""),
        ("Disposal", "", "", "sec"),
        ("Recommended destination", r.disposal, "",
         "; ".join(r.disposal_reasons)),
    ]
    r0 = _result_rows(ws, r0, rows, widths=(34, 30, 10, 56))
    r0 += 1
    check_rows = [("COMPLIANCE CHECKS", "", "", "sec")]
    for name, ok, note in r.checks:
        check_rows.append((name, "PASS" if ok else "FAIL", "", note))
    r0 = _result_rows(ws, r0, check_rows, widths=(34, 30, 10, 56))
    for n in r.head_notes:
        C(ws, r0, 2, "⚠ " + n, sz=9, fg=ORANGE, wrap=True)
        r0 += 1

    # ── SCENARIO_RESULTS ─────────────────────────────────────────────────
    ws = wb.create_sheet("SCENARIO_RESULTS")
    hdrs = ["Code", "Scenario", "Active", "Class",
            uio.hdr("Allowable (g)", "dP"), uio.hdr("Relieving P", "P"),
            uio.hdr("Relieving T", "T"), "Phase", "Quality x",
            uio.hdr("Relief rate", "mflow"), uio.hdr("Required area", "Ain"),
            "Orifice", "Case Sheet", "Notes"]
    _title(ws, len(hdrs), "SCENARIO RESULTS — all contingencies")
    for ci, h in enumerate(hdrs, 2):
        C(ws, 3, ci, h, bg=NAVY, fg=DGRAY, sz=9, bold=True, ha="center",
          wrap=True)
    ws.row_dimensions[3].height = 26
    rr = 4
    gov_code = r.governing.code if r.governing else None
    for s in r.scens:
        bg = ("FFF2CC" if s.code == gov_code and s.active else
              WHITE if s.active else LGRAY)
        letter = select_orifice(s.area)[0] if s.area > 0 else "—"
        case_sheet = f"CASE_{s.code}"[:31] if s.active else None
        vals = [s.code, s.name, "Y" if s.active else "N", s.klass,
                U("dP", s.allowable_g) if s.active else None,
                U("P", s.p1) if s.active else None,
                U("T", s.t, 1) if s.active else None,
                s.phase if s.active else None,
                round(s.quality, 3) if s.active else None,
                U("mflow", s.w, 1) if s.active else None,
                U("Ain", s.area, 4) if s.active else None,
                letter if s.active else None,
                case_sheet,
                "; ".join(s.notes)]
        for ci, v in enumerate(vals, 2):
            cell = C(ws, rr, ci, "—" if v is None else v, bg=bg, sz=8,
                     ha="right" if isinstance(v, (int, float)) else "left",
                     wrap=(ci == len(hdrs) + 1))
            if case_sheet and v is case_sheet:
                cell.hyperlink = f"#'{case_sheet}'!A1"
                cell.font = Font(name="Calibri", size=8, underline="single",
                                  color="0563C1")
        rr += 1
    for ci, w in enumerate([8, 26, 7, 9, 12, 11, 10, 10, 8, 12, 12, 8, 12, 70], 2):
        ws.column_dimensions[get_column_letter(ci)].width = w

    # ── per-case sheets ──────────────────────────────────────────────────
    for s in r.scens:
        if s.active:
            _write_case_sheet(wb, m, r, s)

    # ── EQUIPMENT_CHECK ──────────────────────────────────────────────────
    if m.equip_rows:
        ws = wb.create_sheet("EQUIPMENT_CHECK")
        hdrs = ["Tag", "Type", uio.hdr("Design P", "P"), uio.hdr("MAWP", "P"),
                uio.hdr("EL", "L"), "Design P ≥ Set P?", "Notes"]
        _title(ws, len(hdrs), "PROTECTED EQUIPMENT CHECK")
        for ci, h in enumerate(hdrs, 2):
            C(ws, 3, ci, h, bg=NAVY, fg=DGRAY, sz=9, bold=True, ha="center")
        for i, e in enumerate(m.equip_rows):
            ok = (e.get("des_p") or 0) >= (m.gen.get("set_p") or 0)
            vals = [e["tag"], e.get("type"), U("P", e.get("des_p")),
                    U("P", e.get("mawp")), U("L", e.get("el"), 2),
                    "OK" if ok else "CHECK", e.get("notes")]
            for ci, v in enumerate(vals, 2):
                C(ws, 4 + i, ci, "—" if v is None else v,
                  bg=(WHITE if ok else "FFC7CE"), sz=9)
        for ci, w in enumerate([14, 12, 12, 12, 10, 14, 30], 2):
            ws.column_dimensions[get_column_letter(ci)].width = w

    # ── NOTES ────────────────────────────────────────────────────────────
    ws = wb.create_sheet("NOTES")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 120
    notes = [
        "METHODOLOGY & ASSUMPTIONS",
        "",
        "· Sizing per API 520 Part I: vapor (isentropic nozzle, critical/subcritical),",
        "  liquid (theoretical orifice with Kd/Kw/Kc/Kv), two-phase (Leung omega, Annex C).",
        "· Fire per API 521 §5.15 with project convention Q = F·43200·A^0.82 [W, m²];",
        "  vapor-only screening and the two-phase swell estimate per methodology §3.10.",
        "  The two-phase W₂ uses a homogeneous-vessel swell approximation — review",
        "  against the corporate energy/swell balance before final design.",
        "· Contingency classes per §2: Design 110%/116%, Fire 121%, Remote =",
        "  min(HTF×DP, corrected test pressure) with 1.3×MAWP fallback.",
        "· CV failure per §3.3: design = bypass at 50% of normal operating Cv;",
        "  remote = bypass 100% open at rated Cv.  ISA/IEC 60534 flow equations.",
        "· Check-valve reverse flow per §3.7 (6% / 2% of minimum internal diameter).",
        "· Tower methods per §5: LOC = HMVRAF flashed adiabatically; boil-up =",
        "  two-flash λ; overfill quality = max reboiler + max feed vapor; abnormal",
        "  heat = duty-ratio scaling with optional LMTD pinch credit (§5.1).",
        "· Static-elevation corrected set pressure per 'PSV Inputs plus'.",
        "· Relieving temperature uses bubble/dew/adiabatic-flash of the linked HMB",
        "  composition (hydraulics flash engine) where available.",
        "· INLET/OUTLET piping use a single-line Darcy model — for relief circuits",
        "  run the hydraulics engine and paste the dP into the override cells.",
        "· Every active scenario gets its own CASE_<code> sheet with a full",
        "  back-pressure iteration (Kb, built-up BP, outlet-line v/Re/f at that",
        "  case's relief rate/properties) and inlet-line dP through the shared",
        "  INLET/OUTLET piping geometry — see SCENARIO_RESULTS 'Case Sheet' links.",
        "",
        "VERIFY all results against API 520/521 and the corporate practice before",
        "issuing for design.  This workbook is a calculation aid, not a substitute",
        "for engineering review.",
    ]
    for i, t in enumerate(notes, 2):
        ws.cell(i, 1).value = t
        ws.cell(i, 1).font = Font(name="Calibri", size=9,
                                  bold=t.isupper() and len(t) > 3,
                                  color=DGRAY)
    wb.save(out_path)
    return out_path


# ═══════════════════════════════════════════════════════════════════════════
#  CLI
# ═══════════════════════════════════════════════════════════════════════════
def run(input_path: str, out_path: str | None = None) -> str:
    m = read_model(input_path)
    print(f"  Units: {m.uio.system}   |   PSV: {m.gen.get('tag')}")
    r = run_model(m)
    if r.governing:
        print(f"  Governing: {r.governing.code}  W="
              f"{m.uio.user('mflow', r.governing.w)} {m.uio.label('mflow')}  "
              f"area={m.uio.user('Ain', r.governing.area)} "
              f"{m.uio.label('Ain')}  orifice {r.orifice}")
    else:
        print("  No active scenarios produced a relief load — check inputs.")
    for name, ok, note in r.checks:
        print(f"  [{'PASS' if ok else 'FAIL'}] {name}  ({note})")
    op = out_path or f"psv_output_{re.sub(r'[^A-Za-z0-9_-]', '_', m.gen.get('tag') or 'PSV')}.xlsx"
    op = os.path.join(os.path.dirname(os.path.abspath(input_path)), op) \
        if not os.path.isabs(op) else op
    write_output(m, r, op)
    print(f"  Saved: {op}")
    return op


def main():
    argv = sys.argv[1:]
    if "--template" in argv or "-t" in argv:
        usys = "SI" if any(a.lower() in ("--si", "-si") for a in argv) else "FPS"
        args = [a for a in argv if not a.startswith("-")]
        out = args[0] if args else "psv_input.xlsx"
        path = create_input_template(out, unit_system=usys)
        print(f"  Template created: {path}  (units: {usys})")
        print("  Activate scenarios on STREAM_INPUTS, fill SCENARIOS/CV_DATA/"
              "GAS_DATA, then run:  python psv.py " + out)
        return
    args = [a for a in argv if not a.startswith("-")]
    inp = args[0] if args else "psv_input.xlsx"
    if not os.path.exists(inp):
        raise SystemExit(f"input not found: {inp}  (use --template to create)")
    run(inp)


if __name__ == "__main__":
    main()
