#!/usr/bin/env python3
"""Procalc CV — control valve sizing & adequacy check (IEC 60534, single file).

Implements a Control Valve sizing/adequacy tool per IEC 60534-2-1 (sizing
equations, liquid & gas/vapor), -2-2 / ISA-75.01.01 conventions for Cv,
IEC 60534-4 & FCI 70-2 (seat leakage classes), and screening-level checks
for cavitation (sigma index / FF), choked flow, two-phase / flashing
service, beta ratio, and noise (IEC 60534-8-3 aerodynamic / -8-4
hydrodynamic — referenced, not fully implemented as dB levels).

Five flow cases (Minimum / Normal / Maximum / Special Case 1 / Special
Case 2) — each may link an HMB workbook + stream (upstream conditions) or
use manual properties.  When a composition is linked, the downstream
(valve-outlet) state at P2 is *predicted* via an adiabatic (isenthalpic)
flash — giving the two-phase / flashing fraction and downstream properties
used for the noise screen.

A separate gas-blowby / check-valve restriction (GASBB) may be linked or
entered manually, mirroring the PSV tool's GAS_DATA convention.

Workbook units are declared on the UNITS sheet (FPS or SI + per-quantity
overrides) — shared Procalc convention (../common/units.py).  The engine
computes internally in SI; Cv itself is reported on the universal US
customary basis (gpm/psi/SG for liquid, lb/h-psia-lb/ft3 for gas) per
industry datasheet convention, independent of the workbook's unit system.

Usage
-----
    python cv.py --template           # blank FPS input workbook
    python cv.py --template --si      # blank SI input workbook
    python cv.py cv_input.xlsx         # run (writes cv_output_<tag>.xlsx)

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
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_HERE, "..", "common"))
sys.path.insert(0, os.path.join(_HERE, "..", "Hydraulics"))
import units as UN                                     # noqa: E402

try:
    import hydraulics as HYD          # HMB stream loading + flash VLE
except Exception as _exc:             # pragma: no cover
    HYD = None
    print(f"  NOTE: hydraulics engine not importable ({_exc}) — "
          "HMB stream linkage & rigorous two-phase prediction disabled; "
          "manual properties only.")

# ═══════════════════════════════════════════════════════════════════════════
#  Constants
# ═══════════════════════════════════════════════════════════════════════════
R_UNIV   = 8314.462618        # J/(kmol·K)
ATM_PA   = 101_325.0
BTU_LB_TO_JKG = 2326.0
PSI_TO_PA = 6894.757293

M3S_TO_GPM = 15850.323         # m3/s -> US gpm
KGS_TO_LBH = 7936.6414         # kg/s -> lb/h

CASE_CODES = [
    ("MIN", "Minimum Flow"),
    ("NOR", "Normal Flow"),
    ("MAX", "Maximum Flow"),
    ("SC1", "Special Case 1"),
    ("SC2", "Special Case 2"),
]
CASE_NAMES = dict(CASE_CODES)

TRIM_CHARS = ["Linear", "Equal Percentage", "Quick Opening"]
LEAK_CLASSES = ["II", "III", "IV", "V", "VI"]

# Class II/III/IV seat leakage as % of rated Cv (FCI 70-2 / IEC 60534-4)
CLASS_PCT_CV = {"II": 0.5, "III": 0.1, "IV": 0.01}

# Class VI — max bubbles/min of air at standard test dP, by port dia (in)
# (FCI 70-2 Table 2 — indicative; confirm against manufacturer certification)
CLASS_VI_TABLE = [
    (1.0, 0.15), (1.5, 0.30), (2.0, 0.45), (2.5, 0.60),
    (3.0, 0.90), (4.0, 1.70), (6.0, 4.00), (8.0, 6.75),
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


# ═══════════════════════════════════════════════════════════════════════════
#  IEC 60534-2-1 sizing engine (internal SI)
# ═══════════════════════════════════════════════════════════════════════════
def gas_mass_flux(p1_pa: float, p2_pa: float, t_k: float, mw: float,
                  z: float, k: float) -> tuple[float, bool]:
    """Isentropic nozzle mass flux G [kg/m²·s] and choked flag (used for the
    GASBB orifice / check-valve restriction estimate)."""
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


def ff_factor(pv_pa: float | None, pc_pa: float | None) -> float:
    """Liquid critical-pressure-ratio factor FF = 0.96 − 0.28·√(Pv/Pc)."""
    if not pv_pa or not pc_pa or pc_pa <= 0:
        return 0.96
    return max(0.0, min(1.0, 0.96 - 0.28 * math.sqrt(max(0.0, pv_pa) / pc_pa)))


def sigma_index(p1_pa: float, p2_pa: float,
                pv_pa: float | None) -> float | None:
    """Service cavitation index σ = (P1 − Pv) / (P1 − P2)  (ISA RP75.23)."""
    dp = p1_pa - p2_pa
    if dp <= 0:
        return None
    return (p1_pa - (pv_pa or 0.0)) / dp


def liquid_dp_choked(fl: float | None, p1_pa: float, pv_pa: float | None,
                      ff: float) -> float:
    """Choked ΔP for liquids: FL²·(P1 − FF·Pv)."""
    fl = fl or 0.9
    return (fl ** 2) * (p1_pa - ff * (pv_pa or 0.0))


def cv_required_liquid(w_kgs: float, rho: float, dp_eff_pa: float) -> float:
    """Required Cv (US customary basis) — liquid, IEC 60534-2-1 Eq.1.

    Cv = Q[gpm] · √(SG / ΔP[psi])
    """
    if w_kgs <= 0 or rho <= 0 or dp_eff_pa <= 0:
        return 0.0
    q_gpm = (w_kgs / rho) * M3S_TO_GPM
    sg = rho / 999.0
    dp_psi = dp_eff_pa / PSI_TO_PA
    return q_gpm * math.sqrt(sg / dp_psi)


def gas_xy(p1_pa: float, p2_pa: float, k: float | None,
           xt: float | None) -> tuple[float, float, float, float, bool]:
    """Gas pressure-drop ratio x, choked limit x_T = Fk·xT, effective x,
    expansion factor Y, and choked flag — IEC 60534-2-1 §5.5/5.6."""
    if p1_pa <= 0:
        return 0.0, 0.0, 0.0, 1.0, False
    fk = max(0.1, (k or 1.4) / 1.4)
    x = max(0.0, (p1_pa - p2_pa) / p1_pa)
    x_choked = max(1e-6, fk * (xt or 0.7))
    x_eff = min(x, x_choked)
    y = max(2.0 / 3.0, 1.0 - x_eff / (3.0 * x_choked))
    return x, x_choked, x_eff, y, (x >= x_choked)


def cv_required_gas(w_kgs: float, p1_pa: float, rho1: float, x_eff: float,
                     y: float) -> float:
    """Required Cv (US customary basis) — gas/vapor, mass-flow form.

    Cv = W[lb/h] / (63.3 · Y · √(x_eff · P1[psia] · ρ1[lb/ft³]))
    """
    if w_kgs <= 0 or p1_pa <= 0 or rho1 <= 0 or x_eff <= 0 or y <= 0:
        return 0.0
    w_lbh = w_kgs * KGS_TO_LBH
    p1_psia = p1_pa / PSI_TO_PA
    rho1_lbft3 = rho1 / 16.018463
    return w_lbh / (63.3 * y * math.sqrt(x_eff * p1_psia * rho1_lbft3))


def cv_required_two_phase(w_kgs: float, x_quality: float | None,
                           rho_l: float | None, rho_v: float | None,
                           dp_eff_pa: float) -> tuple[float, float | None]:
    """Required Cv — flashing / two-phase, homogeneous-mixture estimate.

    Simplified first-pass method: mixes liquid + vapor densities at the
    given mass quality into a single effective density and applies the
    liquid Cv equation.  For final sizing of severe flashing/cavitating
    service use a vendor two-phase method (e.g. ISA / manufacturer
    software).
    """
    if w_kgs <= 0 or dp_eff_pa <= 0:
        return 0.0, None
    rho_l = rho_l or 999.0
    rho_v = rho_v or 1.2
    x_quality = max(0.0, min(1.0, x_quality or 0.0))
    rho_mix = 1.0 / (x_quality / rho_v + (1.0 - x_quality) / rho_l)
    return cv_required_liquid(w_kgs, rho_mix, dp_eff_pa), rho_mix


def travel_pct(cv_req: float | None, cv100: float | None, char: str | None,
               rangeability: float | None = 50.0) -> float | None:
    """Estimated % travel from required Cv, rated Cv100 & trim
    characteristic (Linear / Equal Percentage / Quick Opening)."""
    if not cv100 or cv100 <= 0 or cv_req is None:
        return None
    ratio = cv_req / cv100
    c = (char or "Linear").strip().lower()
    if c.startswith("equal"):
        if ratio <= 0:
            return 0.0
        r = max(rangeability or 50.0, 2.0)
        travel = 100.0 * (1.0 + math.log(ratio) / math.log(r))
    elif c.startswith("quick"):
        travel = 100.0 * math.sqrt(max(0.0, ratio))
    else:                                    # Linear
        travel = 100.0 * ratio
    return max(0.0, min(100.0, travel))


def beta_ratio(bore_m: float | None, pipe_m: float | None) -> float | None:
    """β = valve bore/seat ID ÷ pipe ID."""
    if not bore_m or not pipe_m or pipe_m <= 0:
        return None
    return bore_m / pipe_m


def seat_leakage(klass: str | None, cv100: float | None, dp_pa: float,
                 seat_dia_mm: float | None) -> dict:
    """Seat leakage estimate by class (FCI 70-2 / IEC 60534-4)."""
    k = (klass or "IV").strip().upper()
    if k in CLASS_PCT_CV:
        pct = CLASS_PCT_CV[k]
        leak_cv = cv100 * pct / 100.0 if cv100 else None
        return {"class": k, "pct_cv": pct, "leak_cv": leak_cv,
                "desc": f"Class {k}: ≤ {pct}% of rated Cv "
                        "(FCI 70-2 / IEC 60534-4)"}
    if k == "V":
        dp_bar = max(0.0, dp_pa) / 1e5
        leak = 0.18 * (seat_dia_mm or 0.0) * math.sqrt(dp_bar)
        return {"class": "V", "leak_ml_min": leak,
                "desc": "Class V: indicative liquid leakage estimate "
                        "(order-of-magnitude) — confirm against FCI 70-2 / "
                        "manufacturer certified test data"}
    if k == "VI":
        seat_in = (seat_dia_mm or 0.0) / 25.4
        bubbles = CLASS_VI_TABLE[-1][1]
        for dia, b in CLASS_VI_TABLE:
            if seat_in <= dia:
                bubbles = b
                break
        return {"class": "VI", "bubbles_per_min": bubbles,
                "desc": "Class VI: max bubbles/min of air at standard test "
                        "dP, by port size (FCI 70-2 Table 2 — indicative)"}
    return {"class": k, "desc": "Unrecognized leakage class — "
                                 "see IEC 60534-4 / FCI 70-2"}


def sonic_velocity_gas(k: float | None, z: float | None, t_k: float | None,
                        mw: float | None) -> float | None:
    if not t_k or not mw:
        return None
    return math.sqrt((k or 1.4) * (z or 1.0) * R_UNIV * t_k / mw)


def noise_screen(regime: str, w_kgs: float | None, rho2: float | None,
                  t2_k: float | None, mw: float | None, k: float | None,
                  z: float | None, pipe_id_m: float | None) -> dict:
    """Outlet-velocity / Mach-number screen — a rule-of-thumb flag, not a
    substitute for IEC 60534-8-3 (aerodynamic) / -8-4 (hydrodynamic /
    cavitation) sound-pressure-level calculations."""
    if not pipe_id_m or pipe_id_m <= 0 or not rho2 or rho2 <= 0 or not w_kgs:
        return {}
    a = math.pi / 4.0 * pipe_id_m ** 2
    v2 = (w_kgs / rho2) / a
    out = {"v2": v2}
    if regime == "gas":
        c2 = sonic_velocity_gas(k, z, t2_k, mw)
        if c2:
            mach2 = v2 / c2
            out["c2"], out["mach2"] = c2, mach2
            if mach2 > 0.8:
                out["flag"] = ("HIGH — outlet Mach > 0.8: significant "
                                "aerodynamic noise/vibration risk; low-noise "
                                "trim and an IEC 60534-8-3 acoustic analysis "
                                "are strongly recommended")
            elif mach2 > 0.3:
                out["flag"] = ("MODERATE — outlet Mach > 0.3: screen per "
                                "IEC 60534-8-3; consider low-noise trim / "
                                "diffuser")
            else:
                out["flag"] = "LOW — outlet velocity within typical limits"
    else:
        if v2 > 15.0:                      # ~50 ft/s
            out["flag"] = ("HIGH — outlet velocity > 15 m/s (50 ft/s): "
                            "erosion / hydrodynamic-noise / cavitation-noise "
                            "risk — see IEC 60534-8-4")
        else:
            out["flag"] = "LOW — outlet velocity within typical limits"
    return out


def gasbb_cv_flow_kgs(cv: float, p1_pa: float, p2_pa: float, t1_k: float,
                       mw: float | None, z: float | None, k: float | None,
                       xt: float | None) -> float:
    """Gas blowby flow through a Cv-rated restriction (IEC 60534 gas eq.)."""
    if cv <= 0 or p1_pa <= 0:
        return 0.0
    _x, _xc, x_eff, y, _ch = gas_xy(p1_pa, p2_pa, k, xt)
    rho1 = p1_pa * (mw or 28.96) / ((z or 1.0) * R_UNIV * max(1.0, t1_k))
    w_lbh = 63.3 * cv * y * math.sqrt(max(0.0, x_eff) * (p1_pa / PSI_TO_PA)
                                       * (rho1 / 16.018463))
    return w_lbh / KGS_TO_LBH


def gasbb_orifice_flow_kgs(area_m2: float, cd: float, p1_pa: float,
                             p2_pa: float, t1_k: float, mw: float | None,
                             z: float | None, k: float | None) -> float:
    """Gas blowby flow through a fixed restriction (orifice / partially-open
    check valve) — isentropic nozzle mass flux × area."""
    g, _ = gas_mass_flux(p1_pa, p2_pa, t1_k, mw, z, k)
    return max(0.0, cd) * max(0.0, area_m2) * g


# ═══════════════════════════════════════════════════════════════════════════
#  Stream resolution — HMB linkage (per case) + manual overrides
# ═══════════════════════════════════════════════════════════════════════════
@dataclass
class StreamData:
    """Upstream fluid properties for one case (internal SI)."""
    source: str = "manual"          # 'HMB:<file>:<stream>' or 'manual'
    t_k: float | None = None
    p_pa: float | None = None
    phase: str | None = None        # Vapor | Liquid | Two-Phase
    w: float | None = None          # kg/s, from HMB total stream flow
    mw: float | None = None
    z: float | None = None
    k: float | None = None
    quality: float | None = None    # mass vapor fraction
    rho_v: float | None = None      # kg/m³
    rho_l: float | None = None
    mu_v: float | None = None       # cP
    mu_l: float | None = None
    pv: float | None = None         # Pa — vapor pressure (manual only)
    pc: float | None = None         # Pa — critical pressure (manual only)
    sg: float | None = None
    feed: object | None = None      # FlashFeed (composition) when available
    notes: list = field(default_factory=list)


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


def resolve_stream(fname, case, lookup, row: dict, uio: UIO,
                    base_dir: str) -> StreamData:
    """Build StreamData: HMB link first, then manual orange-cell overrides
    (manual wins where filled)."""
    sd = StreamData()
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
                sd.t_k = _FPS2SI["T"](sp.temp_f) if sp.temp_f is not None else None
                sd.p_pa = _FPS2SI["P"](sp.pres_psia) if sp.pres_psia else None
                sd.phase = sp.phase
                sd.mw = sp.mol_weight or sp.vap_mw
                sd.z = sp.vap_z
                sd.k = sp.vap_cp_cv
                tot = (sp.total_mass or 0.0)
                sd.quality = (sp.vap_mass or 0.0) / tot if tot > 0 else None
                if tot > 0:
                    sd.w = _FPS2SI["mflow"](tot)
                sd.rho_v = _FPS2SI["rho"](sp.vap_density) if sp.vap_density else None
                sd.rho_l = _FPS2SI["rho"](sp.liq_density) if sp.liq_density else None
                sd.mu_v, sd.mu_l = sp.vap_visc, sp.liq_visc
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
    ov = [("t_k", "T", "t1"), ("p_pa", "P", "p1"), ("mw", "MW", "mw"),
          ("z", "-", "z"), ("k", "-", "k"), ("quality", "-", "x1"),
          ("rho_v", "rho", "rho_v"), ("rho_l", "rho", "rho_l"),
          ("mu_v", "visc", "mu_v"), ("mu_l", "visc", "mu_l"),
          ("pv", "P", "pv"), ("pc", "P", "pc"), ("sg", "-", "sg")]
    for attr, qty, key in ov:
        v = num(row.get(key))
        if v is None:
            continue
        setattr(sd, attr, uio.si(qty, v))
    w = num(row.get("w"))
    if w is not None:
        sd.w = uio.si("mflow", w)
    ph = txt(row.get("phase"))
    if ph:
        sd.phase = ph
    if sd.sg is None and sd.rho_l:
        sd.sg = sd.rho_l / 999.0
    return sd


def flash_adiabatic(feed, p_pa: float):
    return HYD.flash_isenthalpic(feed, p_pa / PSI_TO_PA)


# ═══════════════════════════════════════════════════════════════════════════
#  Template builder — styling helpers (shared Procalc palette/conventions)
# ═══════════════════════════════════════════════════════════════════════════
NAVY, WHITE, LGRAY = "1F4973", "FFFFFF", "F2F2F2"
AMBER, TEAL, EGRAY = "FFF2CC", "DDEBF7", "E0E0E0"
GRNHDR, DGRAY, ORANGE = "375623", "404040", "C55A11"


def _fillc(h):
    return PatternFill("solid", fgColor=h)


def _bord():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)


def C(ws, r, c, v=None, bg=None, fg=DGRAY, sz=10, bold=False, wrap=False,
      ha="left"):
    x = ws.cell(row=r, column=c)
    if v is not None:
        x.value = v
    x.font = Font(name="Calibri", size=sz, bold=bold, color=fg)
    if bg:
        x.fill = _fillc(bg)
    x.alignment = Alignment(horizontal=ha, vertical="center", wrap_text=wrap)
    x.border = _bord()
    return x


def _title(ws, ncol, text):
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=1 + ncol)
    C(ws, 1, 2, text, bg=NAVY, fg=WHITE, sz=11, bold=True)
    ws.row_dimensions[1].height = 22


# ── live unit labels (track the UNITS sheet) ────────────────────────────────
# UNITS sheet layout (see common/units.py add_units_sheet): "Unit System" in
# B3, then one row per quantity code (column A) — in QUANTITY_NAMES order,
# starting row 6 — with the per-quantity override in column C.
UNIT_ROWS = {qty: 6 + i for i, qty in enumerate(UN.QUANTITY_NAMES.keys())}


class _U:
    """Marks a `_kv_block`/header unit as a *live* formula referencing the
    UNITS sheet, rather than literal text fixed at template-creation time.

    ``_U("P")``            → the current label for quantity "P"
    ``_U("Q", "/", "dT")``  → composite, e.g. "kW/°C" (literals pass through)
    """
    __slots__ = ("parts",)

    def __init__(self, *parts):
        self.parts = parts


def _pretty_unit(u: str) -> str:
    return {"degF": "°F", "degC": "°C"}.get(u, u)


def _unit_expr(qty: str) -> str:
    """Excel expression (no leading '=') for the live label of `qty`:
    UNITS!C<row> override if set, else the system default for UNITS!B3."""
    row = UNIT_ROWS.get(qty)
    if row is None:
        return f'"{_pretty_unit(UN.SYSTEM_DEFAULTS["FPS"].get(qty, ""))}"'
    ovr = f"UNITS!$C${row}"
    fps_def = _pretty_unit(UN.SYSTEM_DEFAULTS["FPS"][qty])
    si_def = _pretty_unit(UN.SYSTEM_DEFAULTS["SI"][qty])
    pretty_ovr = f'IF({ovr}="degF","°F",IF({ovr}="degC","°C",{ovr}))'
    default = f'IF(UNITS!$B$3="SI","{si_def}","{fps_def}")'
    return f'IF({ovr}<>"",{pretty_ovr},{default})'


def _unit_formula(qty: str) -> str:
    return "=" + _unit_expr(qty)


def _unit_formula_parts(parts) -> str:
    pieces = [_unit_expr(p) if p in UNIT_ROWS else f'"{p}"' for p in parts]
    return "=" + " & ".join(pieces)


def _unit_cell(unit):
    """Resolve a `_kv_block` unit-hint: `_U(...)` → live formula string,
    anything else passes through unchanged (literal text)."""
    if isinstance(unit, _U):
        if len(unit.parts) == 1:
            return _unit_formula(unit.parts[0])
        return _unit_formula_parts(unit.parts)
    return unit


def _hdr_formula(base: str, qty: str) -> str:
    """Live column-header formula: '<base> (<live unit>)'."""
    return f'="{base} (" & {_unit_expr(qty)} & ")"'


def _kv_block(ws, r, rows, widths=(34, 16, 12, 64)):
    """rows = list of (label, default, unit_hint, note, kind) — kind: in/calc/sec."""
    for col, w in zip("BCDE", widths):
        ws.column_dimensions[col].width = w
    for label, default, unit, note, kind in rows:
        if kind == "sec":
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
            C(ws, r, 2, label, bg=GRNHDR, fg=WHITE, sz=10, bold=True)
        else:
            C(ws, r, 2, label, bg=LGRAY, sz=9, bold=False)
            C(ws, r, 3, default,
              bg=(EGRAY if kind == "calc" else AMBER), sz=9, ha="right")
            C(ws, r, 4, _unit_cell(unit), sz=9, fg="808080")
            C(ws, r, 5, note, sz=9, fg="808080", wrap=True)
        r += 1
    return r


# ── CASES table columns (mirrors PSV's STREAM_INPUTS convention) ───────────
CASE_COLS = [
    ("code",     "Case",                       None),
    ("name",     "Description",                None),
    ("active",   "Active (Y/N)",                None),
    ("hmb_file", "HMB File",                    None),
    ("hmb_case", "HMB Case",                    None),
    ("stream",   "Stream Lookup (upstream)",    None),
    ("p1",       "P1 — upstream",               "P"),
    ("p2",       "P2 — downstream",             "P"),
    ("t1",       "T1 — upstream",               "T"),
    ("w",        "Mass Flow W",                 "mflow"),
    ("phase",    "Phase (upstream)",            None),
    ("x1",       "Quality x1 (upstream)",       None),
    ("mw",       "MW",                          "MW"),
    ("z",        "Z",                           None),
    ("k",        "Cp/Cv k",                     None),
    ("rho_l",    "Liq Density",                 "rho"),
    ("rho_v",    "Vap Density",                 "rho"),
    ("mu_l",     "Liq Visc (cP)",               None),
    ("mu_v",     "Vap Visc (cP)",               None),
    ("pv",       "Vapor Pressure Pv",           "P"),
    ("pc",       "Critical Pressure Pc",        "P"),
    ("sg",       "Liquid SG",                   None),
    ("notes",    "Notes",                       None),
]
CASE_COL_WIDTHS = [10, 28, 9, 14, 10, 16, 9, 9, 9, 9, 10, 8, 7, 6, 6,
                   9, 9, 8, 8, 9, 9, 7, 24]


def create_input_template(out_path: str = "cv_input.xlsx",
                           unit_system: str = "FPS") -> str:
    usys = UN.UnitSystem(unit_system)
    uio = UIO(usys)
    L = lambda *parts: _U(*parts)
    wb = Workbook()

    # ── GENERAL ────────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "GENERAL"
    _title(ws, 4, "CONTROL VALVE GENERAL DATA — identification · valve data · "
                  f"piping · service/material   (units: {usys.system}, "
                  "see UNITS sheet)")
    rows = [
        ("Identification", "", "", "", "sec"),
        ("Unit", "", "", "Process unit", "in"),
        ("P&ID No", "", "", "", "in"),
        ("Valve Tag", "FCV-001", "", "", "in"),
        ("Service / fluid description", "", "", "", "in"),
        ("Valve Data", "", "", "", "sec"),
        ("Valve type", "Globe", "",
         "Globe | Ball | Butterfly | Angle | ...", "in"),
        ("Trim characteristic", "Equal Percentage", "",
         "Linear | Equal Percentage | Quick Opening", "in"),
        ("Rated Cv @ 100% travel (Cv100)", None, "",
         "US customary Cv (gpm·√SG/√psi liquid basis)", "in"),
        ("Rangeability R (equal-% trim)", 50, "",
         "Typical 30–50; used for % travel back-calc", "in"),
        ("Valve size (NPS)", None, "in", "Always inches", "in"),
        ("Valve seat / bore inside diameter", None, "in",
         "Always inches — used for β ratio & Class V/VI leakage", "in"),
        ("FL — liquid pressure-recovery factor", 0.90, "",
         "IEC 60534-2-1; manufacturer data preferred", "in"),
        ("xT — pressure-differential ratio factor (choked, gas)", 0.70, "",
         "Manufacturer data preferred", "in"),
        ("Piping (for β ratio & noise screening)", "", "", "", "sec"),
        ("Inlet line inside diameter", None, "in", "Always inches", "in"),
        ("Outlet line inside diameter", None, "in", "Always inches", "in"),
        ("Service / Material Considerations", "", "", "", "sec"),
        ("Seat leakage class", "IV", "",
         "II | III | IV | V | VI (FCI 70-2 / IEC 60534-4)", "in"),
        ("NACE MR0175 / ISO 15156 compliance required?", "N", "Y/N", "", "in"),
        ("H2S service?", "N", "Y/N", "", "in"),
        ("H2S concentration", None, "ppm", "", "in"),
        ("Hydrogen (H2) service?", "N", "Y/N",
         "API RP 941 Nelson-curve screening for HT/HP H2 service", "in"),
        ("Materials / special-metallurgy notes", "", "", "", "in"),
    ]
    r_end = _kv_block(ws, 3, rows)

    # data validations (column C of GENERAL)
    def _row_of(label):
        for i, rr in enumerate(rows):
            if rr[0] == label:
                return 3 + i
        return None

    dv_trim = DataValidation(type="list",
                              formula1='"Linear,Equal Percentage,Quick Opening"',
                              allow_blank=True)
    ws.add_data_validation(dv_trim)
    dv_trim.add(f"C{_row_of('Trim characteristic')}")

    dv_class = DataValidation(type="list", formula1='"II,III,IV,V,VI"',
                               allow_blank=True)
    ws.add_data_validation(dv_class)
    dv_class.add(f"C{_row_of('Seat leakage class')}")

    dv_yn = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    ws.add_data_validation(dv_yn)
    for lbl in ("NACE MR0175 / ISO 15156 compliance required?",
                "H2S service?", "Hydrogen (H2) service?"):
        dv_yn.add(f"C{_row_of(lbl)}")

    _finish_template(wb, uio, out_path)
    return out_path


def _finish_template(wb: Workbook, uio: UIO, out_path: str) -> str:
    L = lambda *parts: _U(*parts)

    # ── CASES ────────────────────────────────────────────────────────────
    ws = wb.create_sheet("CASES")
    ncol = len(CASE_COLS)
    _title(ws, ncol, "CASES — Minimum / Normal / Maximum / Special Case 1-2.  "
                     "Link an HMB workbook + upstream stream per case, or "
                     "fill the orange manual-property cells (manual "
                     "overrides HMB).  P2 (downstream) is always a manual "
                     "design input.")
    C(ws, 2, 2, "Two-phase prediction: when an HMB composition is linked, the "
                "engine flashes the upstream stream adiabatically to P2 to "
                "predict downstream phase / flashing fraction / properties.",
      sz=9, fg="808080")
    for ci, (key, base, qty) in enumerate(CASE_COLS, 2):
        h = _hdr_formula(base, qty) if qty else base
        C(ws, 4, ci, h, bg=NAVY, fg=WHITE, sz=9, bold=True, ha="center", wrap=True)
    ws.row_dimensions[4].height = 28
    manual_keys = {"p1", "p2", "t1", "w", "phase", "x1", "mw", "z", "k",
                    "rho_l", "rho_v", "mu_l", "mu_v", "pv", "pc", "sg"}
    for ri, (code, name) in enumerate(CASE_CODES, 5):
        for ci, (key, _b, _q) in enumerate(CASE_COLS, 2):
            if key == "code":
                C(ws, ri, ci, code, bg=LGRAY, sz=9, bold=True)
            elif key == "name":
                C(ws, ri, ci, name, bg=LGRAY, sz=9)
            elif key == "active":
                C(ws, ri, ci, "N", bg=AMBER, sz=9, ha="center")
            elif key in manual_keys:
                C(ws, ri, ci, None, bg=AMBER, sz=9, ha="right")
            else:
                C(ws, ri, ci, None, bg=TEAL, sz=9)
    last_row = 4 + len(CASE_CODES)
    dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"D5:D{last_row}")
    dv2 = DataValidation(type="list", formula1='"Liquid,Vapor,Gas,Two-Phase"',
                          allow_blank=True)
    ws.add_data_validation(dv2)
    dv2.add(f"L5:L{last_row}")
    for ci, w in enumerate(CASE_COL_WIDTHS, 2):
        ws.column_dimensions[get_column_letter(ci)].width = w

    # ── GASBB ────────────────────────────────────────────────────────────
    ws = wb.create_sheet("GASBB")
    _title(ws, 4, "GAS BLOWBY / CHECK-VALVE REVERSE-FLOW RESTRICTION — "
                  "optional, link an HMB workbook + stream or use manual "
                  "properties")
    rows = [
        ("Gas Blowby / Reverse Flow", "", "", "", "sec"),
        ("Active?", "N", "Y/N", "", "in"),
        ("Source description", "", "", "", "in"),
        ("HMB File", "", "", "", "in"),
        ("HMB Case", "", "", "", "in"),
        ("Stream Lookup", "", "", "", "in"),
        ("Restriction type", "Cv", "", "Cv | Orifice | CheckValve", "in"),
        ("Cv (full open)", None, "", "if type = Cv", "in"),
        ("xT", 0.7, "", "if type = Cv (gas)", "in"),
        ("Orifice area", None, L("Ain"), "if type = Orifice", "in"),
        ("Discharge coefficient Cd", 0.61, "",
         "0.61 sharp edge | 0.82 rounded", "in"),
        ("Check valve min internal diameter", None, "in",
         "if type = CheckValve — always inches", "in"),
        ("Partial-failure orifice — % of ID", 6, "% of ID",
         "Typical 6% (API 521 convention)", "in"),
        ("Source pressure P1 (manual)", None, L("P"), "if HMB not linked", "in"),
        ("Source temperature T1 (manual)", None, L("T"), "if HMB not linked", "in"),
        ("Downstream pressure P2", None, L("P"),
         "destination pressure (flare header / downstream vessel)", "in"),
        ("MW (manual)", None, L("MW"), "if HMB not linked", "in"),
        ("Z (manual)", None, "", "if HMB not linked", "in"),
        ("k = Cp/Cv (manual)", None, "", "if HMB not linked", "in"),
    ]
    r_end = _kv_block(ws, 3, rows)

    def _row_of(label):
        for i, rr in enumerate(rows):
            if rr[0] == label:
                return 3 + i
        return None

    dv_a = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    ws.add_data_validation(dv_a)
    dv_a.add(f"C{_row_of('Active?')}")
    dv_r = DataValidation(type="list", formula1='"Cv,Orifice,CheckValve"',
                          allow_blank=True)
    ws.add_data_validation(dv_r)
    dv_r.add(f"C{_row_of('Restriction type')}")

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


@dataclass
class Case:
    code: str
    name: str
    active: bool = False
    source: str = "manual"
    p1: float | None = None
    p2: float | None = None
    t1: float | None = None
    w: float | None = None
    phase: str | None = None
    x1: float | None = None
    mw: float | None = None
    z: float | None = None
    k: float | None = None
    rho_l: float | None = None
    rho_v: float | None = None
    mu_l: float | None = None
    mu_v: float | None = None
    pv: float | None = None
    pc: float | None = None
    sg: float | None = None
    feed: object | None = None
    notes: list = field(default_factory=list)
    res: dict = field(default_factory=dict)


@dataclass
class Model:
    uio: UIO = None
    gen: dict = field(default_factory=dict)
    cases: list = field(default_factory=list)
    gasbb: dict = field(default_factory=dict)
    base_dir: str = "."


def read_model(path: str) -> Model:
    wb = load_workbook(path, data_only=True)
    uio = UIO(UN.UnitSystem.from_workbook(wb))
    m = Model(uio=uio, base_dir=os.path.dirname(os.path.abspath(path)) or ".")

    g = _read_kv(wb["GENERAL"])
    sg_ = lambda lbl: txt(g.get(lbl))
    sn = lambda lbl, qty=None: (uio.si(qty, num(g.get(lbl)))
                                 if qty else num(g.get(lbl)))
    bore_in = sn("Valve seat / bore inside diameter")
    inlet_in = sn("Inlet line inside diameter")
    outlet_in = sn("Outlet line inside diameter")
    m.gen = {
        "unit": sg_("Unit"), "pid": sg_("P&ID No"),
        "tag": sg_("Valve Tag") or "FCV",
        "service": sg_("Service / fluid description"),
        "valve_type": sg_("Valve type") or "Globe",
        "trim_char": sg_("Trim characteristic") or "Linear",
        "cv100": sn("Rated Cv @ 100% travel (Cv100)"),
        "rangeability": sn("Rangeability R (equal-% trim)") or 50.0,
        "valve_nps": sn("Valve size (NPS)"),
        "bore_in": bore_in,
        "bore_m": bore_in * 0.0254 if bore_in else None,
        "fl": sn("FL — liquid pressure-recovery factor") or 0.9,
        "xt": sn("xT — pressure-differential ratio factor (choked, gas)") or 0.7,
        "inlet_in": inlet_in,
        "inlet_m": inlet_in * 0.0254 if inlet_in else None,
        "outlet_in": outlet_in,
        "outlet_m": outlet_in * 0.0254 if outlet_in else None,
        "leak_class": (sg_("Seat leakage class") or "IV").upper(),
        "nace": yes(g.get("NACE MR0175 / ISO 15156 compliance required?")),
        "h2s": yes(g.get("H2S service?")),
        "h2s_ppm": sn("H2S concentration"),
        "h2": yes(g.get("Hydrogen (H2) service?")),
        "matl_notes": sg_("Materials / special-metallurgy notes"),
    }

    ws = wb["CASES"]
    for r in range(5, 5 + len(CASE_CODES)):
        code = txt(ws.cell(r, 2).value)
        if not code:
            continue
        row = {}
        for ci, (key, _b, _q) in enumerate(CASE_COLS, 2):
            row[key] = ws.cell(r, ci).value
        c = Case(code=code.upper(), name=txt(row.get("name")) or CASE_NAMES.get(code.upper(), code))
        c.active = yes(row.get("active"))
        if c.active:
            sd = resolve_stream(txt(row.get("hmb_file")), txt(row.get("hmb_case")),
                                 txt(row.get("stream")), row, uio, m.base_dir)
            c.source = sd.source
            c.notes.extend(sd.notes)
            c.p1, c.t1 = sd.p_pa, sd.t_k
            c.w = sd.w
            c.phase = sd.phase
            c.x1 = sd.quality
            c.mw, c.z, c.k = sd.mw, sd.z, sd.k
            c.rho_l, c.rho_v = sd.rho_l, sd.rho_v
            c.mu_l, c.mu_v = sd.mu_l, sd.mu_v
            c.pv, c.pc, c.sg = sd.pv, sd.pc, sd.sg
            c.feed = sd.feed
            c.p2 = uio.si("P", num(row.get("p2")))
        m.cases.append(c)

    gb = _read_kv(wb["GASBB"])
    active = yes(gb.get("Active?"))
    m.gasbb = {"active": active}
    if active:
        row = {"p1": gb.get("Source pressure P1 (manual)"),
               "t1": gb.get("Source temperature T1 (manual)"),
               "mw": gb.get("MW (manual)"), "z": gb.get("Z (manual)"),
               "k": gb.get("k = Cp/Cv (manual)")}
        sd = resolve_stream(txt(gb.get("HMB File")), txt(gb.get("HMB Case")),
                             txt(gb.get("Stream Lookup")), row, uio, m.base_dir)
        cv_id_in = num(gb.get("Check valve min internal diameter"))
        m.gasbb.update({
            "sd": sd,
            "desc": txt(gb.get("Source description")),
            "rtype": txt(gb.get("Restriction type")) or "Cv",
            "cv": num(gb.get("Cv (full open)")) or 0.0,
            "xt": num(gb.get("xT")) or 0.7,
            "area_m2": uio.si("Ain", num(gb.get("Orifice area"))),
            "cd": num(gb.get("Discharge coefficient Cd")) or 0.61,
            "cv_id_m": cv_id_in * 0.0254 if cv_id_in else None,
            "pct_id": num(gb.get("Partial-failure orifice — % of ID")) or 6.0,
            "p1": sd.p_pa,
            "t1": sd.t_k,
            "mw": sd.mw, "z": sd.z, "k": sd.k,
            "p2": uio.si("P", num(gb.get("Downstream pressure P2"))),
        })
    wb.close()
    return m


# ═══════════════════════════════════════════════════════════════════════════
#  Compute engine
# ═══════════════════════════════════════════════════════════════════════════
def compute_case(c: Case, m: Model) -> None:
    if not c.active:
        return
    res = c.res
    p1, p2 = c.p1 or 0.0, c.p2 or 0.0
    if p1 <= 0 or p2 <= 0 or p2 >= p1:
        c.notes.append("P1/P2 missing or P2 ≥ P1 — case not evaluated")
        return
    dp = p1 - p2
    res["dp"] = dp

    # ── downstream prediction via adiabatic flash (two-phase prediction) ──
    t2 = x2 = rho_v2 = rho_l2 = None
    if c.feed is not None and HYD is not None:
        try:
            fr = flash_adiabatic(c.feed, p2)
            t2 = _FPS2SI["T"](fr.temp_f) if fr.temp_f is not None else c.t1
            x2 = fr.quality
            if fr.vap_mw and t2:
                rho_v2 = p2 * fr.vap_mw / ((c.z or 1.0) * R_UNIV * t2)
            rho_l2 = c.rho_l            # ~unchanged across moderate ΔP
        except Exception as exc:
            c.notes.append(f"downstream flash failed: {exc}")
    res.update(t2=t2, x2=x2, rho_v2=rho_v2, rho_l2=rho_l2)

    # ── regime classification ──────────────────────────────────────────
    ph = (c.phase or "").strip().lower()
    if "two" in ph or "mixed" in ph:
        regime, x_q = "two-phase", (c.x1 if c.x1 is not None else 0.5)
    elif "liq" in ph:
        if x2 is not None and x2 > 0.01:
            regime, x_q = "two-phase", x2
            c.notes.append("flashing across the valve predicted — downstream "
                            f"quality x ≈ {x2:.3f} (adiabatic flash at P2)")
        else:
            regime, x_q = "liquid", 0.0
    elif "vap" in ph or "gas" in ph:
        regime, x_q = "gas", 1.0
    else:
        regime = "liquid" if (c.rho_l and not c.rho_v) else "gas"
        x_q = c.x1 if c.x1 is not None else (0.0 if regime == "liquid" else 1.0)
    res["regime"], res["x_q"] = regime, x_q

    # ── cavitation factors ────────────────────────────────────────────
    ff = ff_factor(c.pv, c.pc)
    res["ff"] = ff
    res["sigma"] = sigma_index(p1, p2, c.pv)

    if regime == "gas":
        x, x_ch, x_eff, y, choked = gas_xy(p1, p2, c.k, m.gen["xt"])
        rho1 = c.rho_v
        if not rho1 and c.t1:
            rho1 = p1 * (c.mw or 28.96) / ((c.z or 1.0) * R_UNIV * c.t1)
        cv_req = cv_required_gas(c.w or 0.0, p1, rho1 or 0.0, x_eff, y)
        res.update(x=x, x_choked=x_ch, y=y, choked=choked,
                   cv_req=cv_req, rho_eff=rho1, sg_eff=None)
    else:
        dp_choked = liquid_dp_choked(m.gen["fl"], p1, c.pv, ff)
        choked = dp_choked > 0 and dp >= dp_choked
        dp_eff = min(dp, dp_choked) if dp_choked > 0 else dp
        res.update(dp_choked=dp_choked, choked=choked, dp_eff=dp_eff)
        if regime == "liquid":
            cv_req = cv_required_liquid(c.w or 0.0, c.rho_l or 999.0, dp_eff)
            res.update(cv_req=cv_req, rho_eff=c.rho_l, sg_eff=c.sg)
        else:                              # two-phase
            cv_req, rho_mix = cv_required_two_phase(c.w or 0.0, x_q, c.rho_l,
                                                       c.rho_v, dp_eff)
            res.update(cv_req=cv_req, rho_eff=rho_mix,
                       sg_eff=(rho_mix / 999.0 if rho_mix else None))

    # ── % travel, beta ratio ──────────────────────────────────────────
    res["travel"] = travel_pct(res.get("cv_req"), m.gen["cv100"],
                                m.gen["trim_char"], m.gen["rangeability"])
    res["beta_in"] = beta_ratio(m.gen["bore_m"], m.gen["inlet_m"])
    res["beta_out"] = beta_ratio(m.gen["bore_m"], m.gen["outlet_m"])

    # ── noise screening (downstream / outlet conditions) ───────────────
    if regime == "gas":
        rho2 = rho_v2 or res.get("rho_eff")
    elif regime == "liquid":
        rho2 = rho_l2 or c.rho_l
    else:
        if rho_v2 and rho_l2 and x2 is not None and rho_v2 > 0 and rho_l2 > 0:
            rho2 = 1.0 / (x2 / rho_v2 + (1.0 - x2) / rho_l2)
        else:
            rho2 = res.get("rho_eff")
    res["noise"] = noise_screen(regime, c.w, rho2, t2 or c.t1, c.mw, c.k,
                                 c.z, m.gen["outlet_m"])

    # ── seat leakage ───────────────────────────────────────────────────
    res["leak"] = seat_leakage(m.gen["leak_class"], m.gen["cv100"], dp,
                                (m.gen["bore_in"] or 0.0) * 25.4)


def compute_gasbb(m: Model) -> dict:
    g = m.gasbb
    res = {}
    if not g.get("active"):
        return res
    p1, p2 = g.get("p1") or 0.0, g.get("p2") or 0.0
    if p1 <= 0 or p2 <= 0 or p2 >= p1:
        res["note"] = "P1/P2 missing or P2 ≥ P1 — not evaluated"
        return res
    t1 = g.get("t1") or 288.7
    mw, z, k = g.get("mw"), g.get("z"), g.get("k")
    rtype = (g.get("rtype") or "Cv").strip().lower()
    if rtype.startswith("cv"):
        w = gasbb_cv_flow_kgs(g.get("cv") or 0.0, p1, p2, t1, mw, z, k, g.get("xt"))
        res["basis"] = f"Cv = {g.get('cv')}, xT = {g.get('xt')}"
    elif rtype.startswith("orif"):
        area = g.get("area_m2") or 0.0
        w = gasbb_orifice_flow_kgs(area, g.get("cd") or 0.61, p1, p2, t1, mw, z, k)
        res["basis"] = f"orifice area = {area:.6g} m²,  Cd = {g.get('cd')}"
    else:                                  # CheckValve
        id_m = g.get("cv_id_m") or 0.0
        pct = (g.get("pct_id") or 6.0) / 100.0
        area = math.pi / 4.0 * (id_m * pct) ** 2
        w = gasbb_orifice_flow_kgs(area, g.get("cd") or 0.61, p1, p2, t1, mw, z, k)
        res["basis"] = (f"{pct * 100:.0f}% of {id_m * 1000:.1f} mm ID "
                        f"→ area {area * 1e6:.2f} mm²,  Cd = {g.get('cd')}")
    res.update(w=w, p1=p1, p2=p2, t1=t1, mw=mw, z=z, k=k, rtype=g.get("rtype"))
    return res


# ═══════════════════════════════════════════════════════════════════════════
#  Output workbook
# ═══════════════════════════════════════════════════════════════════════════
def _result_rows(ws, r0, rows, widths=(34, 30, 10, 56)):
    """Render a key/value result block.

    ``rows`` = list of ``(label, value, unit, note)``; ``note == "sec"``
    marks a section-header row.  ``value`` of literal ``"PASS"``/``"FAIL"``
    is colour-coded as a compliance result.
    """
    for col, w in zip("BCDE", widths):
        ws.column_dimensions[col].width = w
    for label, val, unit, note in rows:
        if note == "sec":
            ws.merge_cells(start_row=r0, start_column=2, end_row=r0, end_column=5)
            C(ws, r0, 2, label, bg=GRNHDR, fg=WHITE, sz=10, bold=True)
        else:
            pf = val in ("PASS", "FAIL")
            bg = "C6EFCE" if val == "PASS" else "FFC7CE" if val == "FAIL" else TEAL
            C(ws, r0, 2, label, bg=LGRAY, sz=9)
            C(ws, r0, 3, "—" if val is None else val, bg=bg, sz=9, bold=pf,
              ha="center" if pf else "right")
            C(ws, r0, 4, unit, sz=9, fg="808080")
            C(ws, r0, 5, note, sz=9, fg="808080", wrap=True)
        r0 += 1
    return r0


def _write_case_sheet(wb, m: Model, c: Case) -> None:
    uio = m.uio
    U, L = uio.user, uio.label
    res = c.res
    ws = wb.create_sheet(f"CASE_{c.code}"[:31])
    _title(ws, 4, f"{c.code} — {c.name}   |   units: {uio.system}")

    rows = [("Upstream / Flow Conditions", "", "", "sec"),
            ("Source", c.source, "", ""),
            ("Phase (upstream)", c.phase, "", ""),
            ("P1 — upstream pressure", U("P", c.p1), L("P"), ""),
            ("P2 — downstream pressure", U("P", c.p2), L("P"), "")]
    if "dp" in res:
        rows.append(("Differential pressure ΔP", U("dP", res["dp"]), L("dP"), ""))
    rows.append(("T1 — upstream temperature", U("T", c.t1, 1), L("T"), ""))
    rows.append(("Mass flow W", U("mflow", c.w, 2), L("mflow"), ""))
    rows.append(("MW", round(c.mw, 2) if c.mw else None, L("MW"), ""))
    rows.append(("Z", round(c.z, 3) if c.z else None, "", ""))
    rows.append(("k = Cp/Cv", round(c.k, 3) if c.k else None, "", ""))
    if c.rho_l:
        rows.append(("Liquid density", U("rho", c.rho_l), L("rho"), ""))
    if c.rho_v:
        rows.append(("Vapor density", U("rho", c.rho_v), L("rho"), ""))
    if c.mu_l:
        rows.append(("Liquid viscosity", round(c.mu_l, 4), "cP", ""))
    if c.mu_v:
        rows.append(("Vapor viscosity", round(c.mu_v, 4), "cP", ""))
    if c.sg:
        rows.append(("Liquid SG", round(c.sg, 4), "", ""))
    if c.pv:
        rows.append(("Vapor pressure Pv", U("P", c.pv), L("P"), ""))
    if c.pc:
        rows.append(("Critical pressure Pc", U("P", c.pc), L("P"), ""))

    if not res:
        rows.append(("Result", "not evaluated", "", "; ".join(c.notes) or
                      "incomplete inputs"))
        _result_rows(ws, 3, rows, widths=(36, 24, 10, 56))
        return

    rows.append(("Downstream Prediction (adiabatic flash to P2)", "", "", "sec"))
    if res.get("t2") is not None:
        rows.append(("T2 — downstream temperature (predicted)",
                      U("T", res["t2"], 1), L("T"), ""))
        rows.append(("x2 — downstream vapor mass fraction (predicted)",
                      round(res["x2"], 4), "",
                      f"flashing ≈ {res['x2'] * 100:.2f}% by mass"))
        if res.get("rho_v2"):
            rows.append(("Vapor density (downstream, predicted)",
                          U("rho", res["rho_v2"]), L("rho"), "ideal-gas estimate"))
    else:
        rows.append(("Downstream prediction", "not available", "",
                      "no linked HMB composition — manual properties used"))

    rows.append(("Flow Regime / Cavitation / Choked Flow", "", "", "sec"))
    rows.append(("Governing regime", res["regime"].replace("-", " ").title(), "", ""))
    rows.append(("FF — liquid critical pressure-ratio factor",
                  round(res["ff"], 4), "", ""))
    if res.get("sigma") is not None:
        rows.append(("Cavitation index σ = (P1−Pv)/(P1−P2)",
                      round(res["sigma"], 3), "",
                      "compare against manufacturer σi/σc curves (ISA RP75.23)"))
    if res["regime"] == "gas":
        rows.append(("x = ΔP / P1", round(res["x"], 4), "", ""))
        rows.append(("Choked limit xT_eff = Fk·xT", round(res["x_choked"], 4)
                      if "x_choked" in res else None, "", ""))
        rows.append(("Y — expansion factor", round(res["y"], 4), "", ""))
        rows.append(("Choked (critical) flow?", "Y" if res["choked"] else "N", "", ""))
    else:
        rows.append(("ΔP choked = FL²·(P1−FF·Pv)", U("dP", res["dp_choked"]), L("dP"), ""))
        rows.append(("Choked / cavitating flow?", "Y" if res["choked"] else "N", "", ""))
        rows.append(("Effective ΔP used for sizing", U("dP", res["dp_eff"]), L("dP"), ""))

    rows.append(("Valve Sizing", "", "", "sec"))
    rows.append(("Required Cv (this case)",
                  round(res["cv_req"], 3) if res.get("cv_req") else None, "",
                  "US customary Cv basis (gpm·√SG/√psi liquid; "
                  "lb/h, psia, lb/ft³ gas)"))
    rows.append(("Rated Cv100", m.gen["cv100"], "", ""))
    if res.get("travel") is not None:
        rows.append(("Estimated % travel", round(res["travel"], 1), "%",
                      f"trim: {m.gen['trim_char']}"))
        if res["travel"] < 10:
            rows.append(("Travel check", "LOW — poor controllability/resolution "
                          "risk near closed", "", ""))
        elif res["travel"] > 90:
            rows.append(("Travel check", "HIGH — limited rangeability margin; "
                          "verify Cv100 / valve size", "", ""))
    else:
        rows.append(("Estimated % travel", "—", "", "Cv100 not specified"))
    if m.gen["cv100"]:
        ok = m.gen["cv100"] >= (res.get("cv_req") or 0)
        rows.append(("Valve capacity check (Cv100 ≥ Cv required)",
                      "PASS" if ok else "FAIL", "", ""))

    rows.append(("Beta Ratio", "", "", "sec"))
    rows.append(("β (inlet) = bore ID / inlet line ID",
                  round(res["beta_in"], 3) if res.get("beta_in") else None, "", ""))
    rows.append(("β (outlet) = bore ID / outlet line ID",
                  round(res["beta_out"], 3) if res.get("beta_out") else None, "", ""))

    rows.append(("Seat Leakage (FCI 70-2 / IEC 60534-4)", "", "", "sec"))
    leak = res.get("leak") or {}
    rows.append((f"Leakage class {leak.get('class', '—')}", leak.get("desc", ""), "", ""))
    if "pct_cv" in leak:
        rows.append(("Allowable leakage", leak.get("leak_cv") and
                      round(leak["leak_cv"], 4), "Cv equiv.",
                      f"{leak['pct_cv']}% of Cv100"))
    if "leak_ml_min" in leak:
        rows.append(("Indicative leakage rate",
                      round(leak["leak_ml_min"], 3), "mL/min (water equiv.)", ""))
    if "bubbles_per_min" in leak:
        rows.append(("Allowable bubbles/min (air)",
                      leak["bubbles_per_min"], "bubbles/min", ""))

    rows.append(("Noise Screening (rule-of-thumb)", "", "", "sec"))
    ns = res.get("noise") or {}
    if ns:
        rows.append(("Outlet velocity v2", U("v", ns["v2"]), L("v"), ""))
        if "mach2" in ns:
            rows.append(("Outlet Mach number M2", round(ns["mach2"], 3), "", ""))
        rows.append(("Screening result", ns.get("flag", ""), "",
                      "see IEC 60534-8-3 (aerodynamic) / -8-4 (hydrodynamic / "
                      "cavitation) for rigorous dB(A) prediction"))
    else:
        rows.append(("Screening result", "not evaluated", "",
                      "outlet line ID or downstream density not available"))

    if c.notes:
        rows.append(("Notes", "", "", "sec"))
        for n in c.notes:
            rows.append((n, "", "", ""))

    _result_rows(ws, 3, rows, widths=(36, 24, 10, 56))


def _write_gasbb_sheet(wb, m: Model, gres: dict) -> None:
    uio = m.uio
    U, L = uio.user, uio.label
    ws = wb.create_sheet("GASBB_RESULT")
    _title(ws, 4, f"GAS BLOWBY / CHECK-VALVE RESTRICTION RESULT   |   "
                  f"units: {uio.system}")
    g = m.gasbb
    rows = [("Source", "", "", "sec"),
            ("Description", g.get("desc") or "", "", ""),
            ("Restriction type", gres.get("rtype") or g.get("rtype"), "", "")]
    if "note" in gres:
        rows.append(("Result", gres["note"], "", ""))
        _result_rows(ws, 3, rows, widths=(36, 24, 10, 56))
        return
    rows += [
        ("P1 — source pressure", U("P", gres["p1"]), L("P"), ""),
        ("P2 — downstream pressure", U("P", gres["p2"]), L("P"), ""),
        ("T1 — source temperature", U("T", gres["t1"], 1), L("T"), ""),
        ("MW", round(gres["mw"], 2) if gres.get("mw") else None, L("MW"), ""),
        ("Z", round(gres["z"], 3) if gres.get("z") else None, "", ""),
        ("k = Cp/Cv", round(gres["k"], 3) if gres.get("k") else None, "", ""),
        ("Result", "", "", "sec"),
        ("Basis", gres.get("basis"), "", ""),
        ("Estimated blowby flow", U("mflow", gres["w"], 3), L("mflow"), ""),
    ]
    _result_rows(ws, 3, rows, widths=(36, 24, 10, 56))


def write_output(m: Model, out_path: str) -> str:
    uio = m.uio
    U, L = uio.user, uio.label
    wb = Workbook()

    active = [c for c in m.cases if c.active]
    sized = [c for c in active if c.res.get("cv_req")]
    gres = compute_gasbb(m)

    # ── SUMMARY ────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "SUMMARY"
    _title(ws, 4, f"CONTROL VALVE SIZING SUMMARY — {m.gen.get('tag')}   |   "
                  f"units: {uio.system}   |   {datetime.now():%d-%b-%Y %H:%M}")
    rows = [
        ("Identification", "", "", "sec"),
        ("Valve Tag", m.gen.get("tag"), "", ""),
        ("Unit / P&ID", f"{m.gen.get('unit') or ''} / {m.gen.get('pid') or ''}", "", ""),
        ("Service", m.gen.get("service"), "", ""),
        ("Valve Data", "", "", "sec"),
        ("Valve type", m.gen.get("valve_type"), "", ""),
        ("Trim characteristic", m.gen.get("trim_char"), "", ""),
        ("Rated Cv100", m.gen.get("cv100"), "", ""),
        ("Valve size (NPS)", m.gen.get("valve_nps"), "in", ""),
        ("Seat / bore inside diameter", m.gen.get("bore_in"), "in", ""),
        ("FL", m.gen.get("fl"), "", ""),
        ("xT", m.gen.get("xt"), "", ""),
        ("Rangeability R", m.gen.get("rangeability"), "", ""),
        ("Piping", "", "", "sec"),
        ("Inlet line ID", m.gen.get("inlet_in"), "in", ""),
        ("Outlet line ID", m.gen.get("outlet_in"), "in", ""),
        ("Service / Material", "", "", "sec"),
        ("Seat leakage class", m.gen.get("leak_class"), "", ""),
        ("NACE MR0175 / ISO 15156 required?", "Y" if m.gen.get("nace") else "N", "", ""),
        ("H2S service?", "Y" if m.gen.get("h2s") else "N",
         m.gen.get("h2s_ppm"), "ppm H2S" if m.gen.get("h2s_ppm") else ""),
        ("Hydrogen (H2) service?", "Y" if m.gen.get("h2") else "N", "",
         "screen against API RP 941 Nelson curves if Y" if m.gen.get("h2") else ""),
        ("Materials notes", m.gen.get("matl_notes") or "", "", ""),
    ]
    r0 = _result_rows(ws, 3, rows, widths=(34, 30, 10, 56))
    r0 += 1

    rows2 = [("Governing Results", "", "", "sec")]
    if sized:
        worst_cv = max(sized, key=lambda c: c.res["cv_req"])
        rows2.append(("Case requiring max Cv (governs valve size)",
                       f"{worst_cv.code} — {worst_cv.name}", "",
                       f"Cv req = {worst_cv.res['cv_req']:.2f}"))
        best_cv = min(sized, key=lambda c: c.res["cv_req"])
        rows2.append(("Case requiring min Cv (controllability check)",
                       f"{best_cv.code} — {best_cv.name}", "",
                       f"Cv req = {best_cv.res['cv_req']:.2f}"))
    else:
        rows2.append(("Governing case", "—", "", "no case produced a sizing result"))
    r0 = _result_rows(ws, r0, rows2, widths=(34, 30, 10, 56))
    r0 += 1

    check_rows = [("COMPLIANCE CHECKS", "", "", "sec")]
    for c in active:
        if not c.res.get("cv_req"):
            continue
        cv100 = m.gen.get("cv100")
        ok_cap = (cv100 or 0) >= c.res["cv_req"]
        check_rows.append((f"{c.code}: Cv100 ≥ Cv required", "PASS" if ok_cap else "FAIL",
                            "", f"req {c.res['cv_req']:.2f} vs rated {cv100 or '—'}"))
        trav = c.res.get("travel")
        if trav is not None:
            ok_t = 10.0 <= trav <= 90.0
            check_rows.append((f"{c.code}: % travel within 10–90%",
                                "PASS" if ok_t else "FAIL", "", f"{trav:.1f}%"))
        if c.res.get("choked"):
            check_rows.append((f"{c.code}: choked / cavitating flow",
                                "FAIL", "", "review trim selection / σ index"))
    if len(check_rows) == 1:
        check_rows.append(("No active cases with a sizing result", "", "", ""))
    r0 = _result_rows(ws, r0, check_rows, widths=(34, 30, 10, 56))

    # ── CASES_RESULTS ──────────────────────────────────────────────────
    ws = wb.create_sheet("CASES_RESULTS")
    hdrs = ["Case", "Description", "Active", "Regime",
            uio.hdr("P1", "P"), uio.hdr("P2", "P"), uio.hdr("ΔP", "dP"),
            uio.hdr("W", "mflow"), "Cv req", "Cv100", "% Travel",
            "Choked?", "σ", "β (in)", "Noise", "Case Sheet", "Notes"]
    _title(ws, len(hdrs), "CASE RESULTS — all flow cases")
    for ci, h in enumerate(hdrs, 2):
        C(ws, 3, ci, h, bg=NAVY, fg=WHITE, sz=9, bold=True, ha="center", wrap=True)
    ws.row_dimensions[3].height = 26
    rr = 4
    for c in m.cases:
        res = c.res
        bg = WHITE if c.active else LGRAY
        case_sheet = f"CASE_{c.code}"[:31] if c.active else None
        noise_flag = ""
        if res.get("noise", {}).get("flag"):
            noise_flag = res["noise"]["flag"].split(" — ")[0]
        vals = [c.code, c.name, "Y" if c.active else "N",
                res.get("regime", "").title() if c.active else None,
                U("P", c.p1) if c.active else None,
                U("P", c.p2) if c.active else None,
                U("dP", res.get("dp")) if c.active else None,
                U("mflow", c.w, 2) if c.active else None,
                round(res["cv_req"], 2) if res.get("cv_req") else None,
                m.gen.get("cv100") if c.active else None,
                round(res["travel"], 1) if res.get("travel") is not None else None,
                ("Y" if res.get("choked") else "N") if c.active and "choked" in res else None,
                round(res["sigma"], 2) if res.get("sigma") is not None else None,
                round(res["beta_in"], 3) if res.get("beta_in") else None,
                noise_flag,
                case_sheet,
                "; ".join(c.notes)]
        for ci, v in enumerate(vals, 2):
            cell = C(ws, rr, ci, "—" if v is None else v, bg=bg, sz=8,
                     ha="right" if isinstance(v, (int, float)) else "left",
                     wrap=(ci == len(hdrs) + 1))
            if case_sheet and v is case_sheet:
                cell.hyperlink = f"#'{case_sheet}'!A1"
                cell.font = Font(name="Calibri", size=8, underline="single",
                                  color="0563C1")
        rr += 1
    for ci, w in enumerate([8, 18, 7, 11, 9, 9, 9, 11, 8, 8, 9, 8, 7, 8, 26, 12, 50], 2):
        ws.column_dimensions[get_column_letter(ci)].width = w

    # ── per-case sheets ──────────────────────────────────────────────────
    for c in m.cases:
        if c.active:
            _write_case_sheet(wb, m, c)

    # ── GASBB_RESULT ───────────────────────────────────────────────────
    if m.gasbb.get("active"):
        _write_gasbb_sheet(wb, m, gres)

    # ── NOTES ────────────────────────────────────────────────────────────
    ws = wb.create_sheet("NOTES")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 120
    notes = [
        "METHODOLOGY & ASSUMPTIONS",
        "",
        "· Sizing per IEC 60534-2-1 / ISA-75.01.01: liquid Cv = Q[gpm]·√(SG/ΔP[psi]);",
        "  gas/vapor Cv = W[lb/h] / (63.3·Y·√(x_eff·P1[psia]·ρ1[lb/ft³])) — Cv is",
        "  reported on the universal US-customary basis regardless of the",
        "  workbook's unit system, per standard valve-datasheet convention.",
        "· Regime per case is taken from the upstream phase; for liquid streams",
        "  the engine flashes the linked HMB composition adiabatically to P2 —",
        "  if that predicts vapor formation (flashing service), the two-phase",
        "  homogeneous-mixture method is used instead.",
        "· Two-phase Cv uses a simplified homogeneous-mixture density (mass",
        "  quality x weighting 1/ρ_v and 1/ρ_l) with the liquid Cv equation.",
        "  This is a first-pass estimate — for severe flashing/cavitating",
        "  service use a vendor two-phase sizing method.",
        "· Cavitation/choked flow: liquids & two-phase per ΔP_choked =",
        "  FL²·(P1 − FF·Pv) with FF = 0.96 − 0.28·√(Pv/Pc); gas/vapor per the",
        "  x_T = Fk·xT choked-flow limit and expansion factor Y (IEC 60534-2-1).",
        "  The cavitation index σ = (P1−Pv)/(P1−P2) (ISA RP75.23) is reported for",
        "  comparison against manufacturer σi/σc trim-selection curves.",
        "· % travel is back-calculated from Cv_required / Cv100 via the selected",
        "  trim characteristic (Linear / Equal Percentage with rangeability R /",
        "  Quick Opening).  Reynolds-number (FR) laminar-flow corrections per",
        "  IEC 60534-2-1 Annex are NOT applied (turbulent flow assumed).",
        "· β ratio = valve seat/bore ID ÷ inlet (or outlet) line ID.",
        "· Seat leakage: Class II/III/IV as %Cv100 (FCI 70-2 / IEC 60534-4);",
        "  Class V as an indicative order-of-magnitude liquid leakage estimate;",
        "  Class VI as max bubbles/min of air by port size (FCI 70-2 Table 2).",
        "  All leakage figures are indicative — confirm with manufacturer",
        "  certified test data for the selected trim.",
        "· Noise is screened via outlet velocity / Mach number only (rule-of-",
        "  thumb risk flags).  Rigorous sound-pressure-level prediction requires",
        "  IEC 60534-8-3 (aerodynamic) or IEC 60534-8-4 (hydrodynamic /",
        "  cavitation) — not implemented here.",
        "· NACE MR0175 / ISO 15156 (sour service materials), and API RP 941",
        "  Nelson curves (H2 service, high-temperature hydrogen attack) are",
        "  flagged as inputs for materials selection — no calculation performed.",
        "· Gas blowby / check-valve reverse flow (GASBB) mirrors the PSV tool's",
        "  GAS_DATA convention: Cv-rated restriction, fixed orifice, or a",
        "  partially-open check valve (% of minimum internal diameter).",
        "",
        "RECENT DEVELOPMENTS / FURTHER READING",
        "",
        "· Multi-stage / tortuous-path anti-cavitation trims and trim-selection",
        "  guidance — ISA RP75.23 (cavitation index method).",
        "· Smart/digital valve positioners with online diagnostics (valve",
        "  signature / step-response analysis) for predictive maintenance and",
        "  in-service seat-leakage / stiction detection.",
        "· CFD-assisted Cv, cavitation-inception and noise prediction is",
        "  increasingly used by manufacturers to supplement IEC 60534 hand",
        "  calculations for severe service.",
        "· Check the current editions of IEC 60534 (parts 2-1, 2-2, 4, 8-3,",
        "  8-4) and ISA-75 series — these are periodically revised.",
        "",
        "VERIFY all results against IEC 60534 / manufacturer sizing software and",
        "the corporate practice before issuing for design.  This workbook is a",
        "calculation aid, not a substitute for engineering review.",
    ]
    for i, t in enumerate(notes, 2):
        ws.cell(i, 1).value = t
        ws.cell(i, 1).font = Font(name="Calibri", size=9,
                                  bold=t.isupper() and len(t) > 3,
                                  color=NAVY if t.isupper() else DGRAY)
    wb.save(out_path)
    return out_path


# ═══════════════════════════════════════════════════════════════════════════
#  CLI
# ═══════════════════════════════════════════════════════════════════════════
def run_model(m: Model) -> None:
    for c in m.cases:
        compute_case(c, m)


def run(input_path: str, out_path: str | None = None) -> str:
    m = read_model(input_path)
    print(f"  Units: {m.uio.system}   |   Valve: {m.gen.get('tag')}")
    run_model(m)
    for c in m.cases:
        if not c.active:
            continue
        if c.res.get("cv_req"):
            print(f"  {c.code}: regime={c.res['regime']:<10} "
                  f"Cv_req={c.res['cv_req']:.2f}  "
                  f"travel={c.res.get('travel')}")
        else:
            print(f"  {c.code}: not evaluated ({'; '.join(c.notes) or 'incomplete inputs'})")
    op = out_path or f"cv_output_{re.sub(r'[^A-Za-z0-9_-]', '_', m.gen.get('tag') or 'CV')}.xlsx"
    op = os.path.join(os.path.dirname(os.path.abspath(input_path)), op) \
        if not os.path.isabs(op) else op
    write_output(m, op)
    print(f"  Saved: {op}")
    return op


def main():
    argv = sys.argv[1:]
    if "--template" in argv or "-t" in argv:
        usys = "SI" if any(a.lower() in ("--si", "-si") for a in argv) else "FPS"
        args = [a for a in argv if not a.startswith("-")]
        out = args[0] if args else "cv_input.xlsx"
        path = create_input_template(out, unit_system=usys)
        print(f"  Template created: {path}  (units: {usys})")
        print("  Activate cases on CASES, fill GENERAL/GASBB, then run:  "
              "python cv.py " + out)
        return
    args = [a for a in argv if not a.startswith("-")]
    inp = args[0] if args else "cv_input.xlsx"
    if not os.path.exists(inp):
        raise SystemExit(f"input not found: {inp}  (use --template to create)")
    run(inp)


if __name__ == "__main__":
    main()
