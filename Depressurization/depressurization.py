#!/usr/bin/env python3
"""Depressurization Dynamics — vessel blowdown time-march.

Builds directly on the existing Procalc engines:
  * hydraulics.py  — Pipeline_Input sheet/reader, pipe ID resolution
                      (Piping Spec + Bore -> internal diameter), units layer.
  * psv.py         — API 520/521 relief-device flow equations (orifice / Cv)
                      and API 521 fire heat-input equations.

Scope (v1 — documented on the output workbook's Notes sheet):
  * Single well-mixed VAPOUR gas space = Vessel Volume + downstream relief
    tailpipe volume (computed from the Pipeline_Input sheet geometry).
  * Adiabatic open-system ("isenthalpic blowdown") energy balance using
    ideal-gas Cp/Cv derived from Vap Cp/Cv (k) and Vap MW; Z held constant.
  * Relief flow sized fresh each timestep from current vessel P, T through
    a fixed orifice area or Cv (no resizing during the run).
  * Fire Case = Yes adds a constant API 521 fire heat input (wetted area
    held fixed at its initial value — no liquid-level tracking).
  * Composition split (if a Stream Lookup / HMB feed is supplied) is a
    real K-value flash at each timestep's P, T for REPORTING only — it does
    not feed back into the simplified single-phase mass/energy balance.
  * Builtup backpressure is solved each timestep with a self-contained
    Darcy-Weisbach + fittings march along the Pipeline_Input tailpipe row(s)
    (not hydraulics.py's full multi-stream noiso engine, which is built for
    HMB circuit marching, not a synthetic per-timestep blowdown state).

Final Pressure 1/2 & Final Time 1/2 are PASS/FAIL design-check targets read
off the computed P-vs-t curve; they do not alter the simulation itself.
"""
from __future__ import annotations

import math
import os
import sys
from dataclasses import dataclass, field
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_HERE, "..", "common"))
sys.path.insert(0, os.path.join(_HERE, "..", "Hydraulics"))
sys.path.insert(0, os.path.join(_HERE, "..", "PSV"))
import units as UN                                                # noqa: E402
from style import (                                               # noqa: E402
    NAVY, WHITE, LGRAY, DGRAY, C, _title, _kv_block, _result_rows, _U,
)
import hydraulics as HYD                                          # noqa: E402
import psv as PSVM                                                # noqa: E402

CASE_SHEET = "Depressurization_Case"
PROFILE_SHEET = "Depressurization_Profile"
CHECKS_SHEET = "Design_Checks"

ROUGHNESS_M = 0.0018 * HYD.FT_TO_M / 12.0   # 0.0018 in commercial steel -> m


# ════════════════════════════════════════════════════════════════════════
#  Unit-conversion helpers (engine works in SI internally, mirrors psv.py)
# ════════════════════════════════════════════════════════════════════════
def _to_pa(psia: float) -> float:
    return psia * HYD.PSIA_TO_PA


def _to_psia(pa: float) -> float:
    return pa / HYD.PSIA_TO_PA


def _to_k(degf: float) -> float:
    return (degf + HYD.F_TO_K_OFFSET) / 1.8


def _to_f(k: float) -> float:
    return k * 1.8 - HYD.F_TO_K_OFFSET


def _ft3_to_m3(v: float) -> float:
    return v * HYD.FT_TO_M ** 3


def _m3_to_ft3(v: float) -> float:
    return v / HYD.FT_TO_M ** 3


def _kgs_to_lbhr(v: float) -> float:
    return v / HYD.LBHR_TO_KGS


def _in_to_m(v: float) -> float:
    return v * HYD.FT_TO_M / 12.0


def _ft_to_m(v: float) -> float:
    return v * HYD.FT_TO_M


def _in2_to_m2(v: float) -> float:
    return v * (HYD.FT_TO_M / 12.0) ** 2


# ════════════════════════════════════════════════════════════════════════
#  Depressurization_Case input schema  (label, default, unit, note, kind)
# ════════════════════════════════════════════════════════════════════════
CASE_FIELDS: list[tuple] = [
    ("GENERAL", None, None, "", "sec"),
    ("Case ID", "Case-1", None, "", "in"),
    ("Description", "", None, "", "in"),
    ("Tailpipe Line No", "", None,
     "Line No in the Pipeline_Input sheet for the relief/tailpipe line — "
     "used for line volume + builtup backpressure; blank = no downstream line",
     "in"),

    ("VESSEL & INITIAL STATE", None, None, "", "sec"),
    ("Vessel Volume", 1000.0, _U("V"), "vessel internal gas volume", "in"),
    ("Initial Pressure", 250.0, _U("P"), "vessel pressure at t = 0", "in"),
    ("Initial Temperature", 100.0, _U("T"), "vessel temperature at t = 0", "in"),
    ("Superimposed BP", 14.7, _U("P"),
     "backpressure already present at the discharge tie-in, independent of "
     "this valve's own flow", "in"),
    ("Builtup BP", None, _U("P"),
     "CALCULATED on run — peak pressure rise in the tailpipe caused by this "
     "valve's own flow", "calc"),

    ("RELIEF GAS PROPERTIES (Stream Lookup optional — composition reporting only)",
     None, None, "", "sec"),
    ("Stream Lookup", "", None,
     "HMB stream key — used ONLY for composition-split reporting", "in"),
    ("HMB File", "", None, "", "in"),
    ("Case", "Case 1", None, "", "in"),
    ("Vap MW", 44.0, None, "drives the mass/energy balance", "in"),
    ("Vap Z", 1.0, None, "held constant through the run", "in"),
    ("Vap Cp/Cv (k)", 1.20, None, "held constant through the run", "in"),
    ("Vap Visc (cP)", 0.012, None, "tailpipe friction only", "in"),

    ("RELIEF DEVICE", None, None, "", "sec"),
    ("Relief Device Type", "Orifice", None, "Orifice | Cv", "in"),
    ("Relief Orifice Area", 1.287, _U("Ain"), "used when Relief Device Type = Orifice", "in"),
    ("Discharge Coeff Cd", 0.975, None, "", "in"),
    ("Relief Cv", 0.0, None, "used when Relief Device Type = Cv", "in"),

    ("TIME", None, None, "", "sec"),
    ("Initial Time", 0.0, None, "seconds — fixed at 0", "calc"),
    ("Timestep", 5.0, None, "seconds", "in"),
    ("Max Duration", 900.0, None, "seconds", "in"),
    ("Final Pressure 1", 100.0, _U("P"), "design-check target #1", "in"),
    ("Final Time 1", 300.0, None, "seconds — design-check target #1", "in"),
    ("Final Pressure 2", 50.0, _U("P"), "design-check target #2", "in"),
    ("Final Time 2", 900.0, None, "seconds — design-check target #2", "in"),

    ("FIRE CASE", None, None, "", "sec"),
    ("Fire Case", "No", None,
     "Yes/No — adds API 521 fire heat input to the energy balance", "in"),
    ("Vessel Orientation", "Vertical", None, "Vertical | Horizontal | Sphere", "in"),
    ("Vessel OD", 96.0, _U("Lin"), "", "in"),
    ("Tan-Tan Length", 30.0, _U("L"), "", "in"),
    ("Head Type", "2:1", None, "2:1 | HEMI | FLAT", "in"),
    ("Normal Liquid Level", 10.0, _U("L"), "from bottom of shell", "in"),
    ("Grade Elevation", 0.0, _U("L"), "", "in"),
    ("Fire Limit Height", 25.0, _U("L"), "API 521 fire height limit above grade", "in"),
    ("Insulated", "No", None, "", "in"),
    ("Drainage", "No", None, "adequate drainage & spacing per API 521", "in"),
    ("Firefighting", "No", None, "", "in"),
    ("Env Factor Override", "", None,
     "blank = computed from Drainage / Firefighting / Insulated", "in"),
    ("Extra Wetted Area", 0.0, _U("A"),
     "additional wetted area not captured by shell/heads", "in"),

    ("NOTES", None, None, "", "sec"),
    ("Notes", "", None, "", "in"),
]


# ════════════════════════════════════════════════════════════════════════
#  Input-template creator  (reuses hydraulics.py's Pipeline_Input + UNITS)
# ════════════════════════════════════════════════════════════════════════
def create_input_template(out_path: str = "depressurization_input.xlsx",
                          unit_system: str = "FPS") -> str:
    """Build the Pipeline_Input/_FittingList/Notes/UNITS sheets via
    hydraulics.py, then add the Depressurization_Case sheet in front."""
    HYD.create_input_template(out_path, unit_system)
    wb = load_workbook(out_path)

    ws = wb.create_sheet(CASE_SHEET, 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    _title(ws, 4, "DEPRESSURIZATION DYNAMICS — vessel blowdown case input")
    _kv_block(ws, 3, CASE_FIELDS)

    ws_n = wb["Notes"] if "Notes" in wb.sheetnames else wb.create_sheet("Notes")
    r = ws_n.max_row + 3 if ws_n.max_row else 1
    for t, bold in [
        ("", False),
        ("DEPRESSURIZATION DYNAMICS — QUICK REFERENCE", True),
        ("  Fill the Depressurization_Case sheet for the vessel blowdown case.", False),
        ("  Tailpipe Line No must match a Line No used in the Pipeline_Input sheet below — "
         "its pipe/fitting rows give the relief line's internal volume and tailpipe friction "
         "(builtup backpressure).  Leave blank if there is no downstream line to model.", False),
        ("  Stream Lookup/HMB File/Case are used ONLY to report a per-timestep composition "
         "split (real K-value flash) — they do not drive the mass/energy balance, which runs "
         "on Vap MW/Z/Cp-Cv as a single-phase ideal gas.", False),
        ("  Fire Case = Yes adds a constant API 521 fire heat input using the vessel geometry "
         "fields; wetted area is held fixed at its initial value through the run.", False),
        ("  Final Pressure/Time 1 & 2 are PASS/FAIL design-check targets only — they do not "
         "change the simulated flow profile or relief device sizing.", False),
        ("  SIMPLIFICATIONS: single well-mixed vapour phase (no liquid-level tracking), "
         "constant Z and Cp/Cv, constant wetted area for the fire case.", False),
    ]:
        ws_n.cell(r, 1).value = t
        ws_n.cell(r, 1).font = Font(name="Calibri", size=9, bold=bold,
                                     color=NAVY if bold else DGRAY)
        r += 1

    wb.active = 0
    wb.save(out_path)
    return out_path


# ════════════════════════════════════════════════════════════════════════
#  Input reader
# ════════════════════════════════════════════════════════════════════════
def read_case_input(path: str) -> tuple[dict, "UN.UnitSystem"]:
    wb = load_workbook(path, data_only=True)
    usys = UN.UnitSystem.from_workbook(wb)
    if CASE_SHEET not in wb.sheetnames:
        raise ValueError(f"{path}: no '{CASE_SHEET}' sheet found")
    ws = wb[CASE_SHEET]

    qty_by_label: dict[str, str | None] = {}
    for label, _default, unit, _note, kind in CASE_FIELDS:
        if kind == "sec":
            continue
        qty_by_label[label] = unit.parts[0] if isinstance(unit, _U) else None

    case: dict = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=2, max_col=3,
                            values_only=True):
        label, val = (row[0], row[1]) if row and len(row) > 1 else (None, None)
        if label is None or label not in qty_by_label:
            continue
        qty = qty_by_label[label]
        if qty and val is not None:
            try:
                val = usys.to_internal(qty, float(val))
            except (TypeError, ValueError):
                pass
        case[str(label).strip()] = val
    return case, usys


# ════════════════════════════════════════════════════════════════════════
#  Line geometry — internal volume (reuses hydraulics.py's ID resolver)
# ════════════════════════════════════════════════════════════════════════
def line_volume_ft3(pipeline_rows: list[dict], line_no: str | None) -> float:
    if not line_no:
        return 0.0
    vol = 0.0
    for row in pipeline_rows:
        if row.get("Line No") != line_no:
            continue
        bore, spec = row.get("Bore (in)"), row.get("Piping Spec")
        length_ft = row.get("Length (ft)") or 0.0
        if not bore or not spec or length_ft <= 0:
            continue
        idr = HYD.resolve_id(spec, bore)
        if not idr.id_in:
            continue
        area_ft2 = math.pi / 4.0 * (idr.id_in / 12.0) ** 2
        vol += area_ft2 * length_ft
    return vol


# ════════════════════════════════════════════════════════════════════════
#  Relief device flow  (API 520 — psv.py, SI internally)
# ════════════════════════════════════════════════════════════════════════
def relief_mdot_kgs(case: dict, p_vessel_pa: float, p_back_pa: float,
                    t_k: float) -> float:
    if p_vessel_pa <= p_back_pa:
        return 0.0
    mw = case.get("Vap MW") or 28.96
    z = case.get("Vap Z") or 1.0
    k = max(1.01, case.get("Vap Cp/Cv (k)") or 1.2)
    device = str(case.get("Relief Device Type") or "Orifice").strip().lower()
    if device.startswith("cv"):
        return PSVM.cv_gas_kgs(case.get("Relief Cv") or 0.0,
                               p_vessel_pa, p_back_pa, t_k, mw, z, k)
    area_m2 = _in2_to_m2(case.get("Relief Orifice Area") or 0.0)
    cd = case.get("Discharge Coeff Cd") or 0.975
    return PSVM.orifice_gas_kgs(area_m2, cd, p_vessel_pa, p_back_pa, t_k, mw, z, k)


# ════════════════════════════════════════════════════════════════════════
#  Tailpipe march — self-contained Darcy-Weisbach + fittings (builtup BP)
# ════════════════════════════════════════════════════════════════════════
def _line_outlet_pa(rows: list[dict], line_no: str, mdot_kgs: float, t_k: float,
                    mw: float, z: float, mu_pas: float, p_start_pa: float) -> float:
    p = p_start_pa
    for row in rows:
        if row.get("Line No") != line_no:
            continue
        length_m = _ft_to_m(row.get("Length (ft)") or 0.0)
        elev_m = _ft_to_m(row.get("Elev Change (ft)") or 0.0)
        fixed_k = row.get("Fixed K") or 0.0
        bore, spec = row.get("Bore (in)"), row.get("Piping Spec")
        if not bore or not spec:
            continue
        idr = HYD.resolve_id(spec, bore)
        if not idr.id_in:
            continue
        id_m = _in_to_m(idr.id_in)
        rho = max(1e-6, p * mw / (z * PSVM.R_UNIV * t_k))
        area_m2 = math.pi / 4.0 * id_m ** 2
        v = mdot_kgs / (rho * area_m2)
        re = rho * v * id_m / max(1e-9, mu_pas)
        if re < 1e-6:
            f = 0.0
        elif re < 2300.0:
            f = 64.0 / re
        else:
            f = 0.25 / (math.log10(ROUGHNESS_M / (3.7 * id_m) + 5.74 / re ** 0.9)) ** 2
        dp_fric = f * (length_m / id_m) * 0.5 * rho * v * v if length_m > 0 else 0.0
        dp_k = fixed_k * 0.5 * rho * v * v
        dp_elev = rho * HYD.G_SI * elev_m
        p = max(1.0, p - dp_fric - dp_k - dp_elev)
    return p


def solve_builtup_bp_pa(rows: list[dict], line_no: str | None, mdot_kgs: float,
                        t_k: float, mw: float, z: float, mu_pas: float,
                        superimposed_bp_pa: float, p_vessel_pa: float) -> float:
    """Builtup BP (Pa) = the tailpipe inlet pressure (just past the relief
    device) minus Superimposed BP, solved so the tailpipe's OWN pressure drop
    (at the current flow) lands the outlet exactly on Superimposed BP."""
    if not line_no or mdot_kgs <= 0 or p_vessel_pa <= superimposed_bp_pa:
        return 0.0

    def resid(p_start_pa: float) -> float:
        return _line_outlet_pa(rows, line_no, mdot_kgs, t_k, mw, z, mu_pas,
                               p_start_pa) - superimposed_bp_pa

    sol = PSVM.bisect(resid, superimposed_bp_pa, p_vessel_pa)
    if sol is None:
        return 0.0
    return max(0.0, sol - superimposed_bp_pa)


# ════════════════════════════════════════════════════════════════════════
#  Fire heat input  (API 521 — psv.py)
# ════════════════════════════════════════════════════════════════════════
def fire_heat_w(case: dict) -> float:
    if not PSVM.yes(case.get("Fire Case")):
        return 0.0
    od_m = _in_to_m(case.get("Vessel OD") or 0.0)
    tt_m = _ft_to_m(case.get("Tan-Tan Length") or 0.0)
    nll_m = _ft_to_m(case.get("Normal Liquid Level") or 0.0)
    grade_m = _ft_to_m(case.get("Grade Elevation") or 0.0)
    firelim_m = _ft_to_m(case.get("Fire Limit Height") or 0.0)
    extra_m2 = (case.get("Extra Wetted Area") or 0.0) * HYD.FT_TO_M ** 2
    f = PSVM.fire_env_factor(PSVM.yes(case.get("Drainage")),
                             PSVM.yes(case.get("Firefighting")),
                             PSVM.yes(case.get("Insulated")),
                             case.get("Env Factor Override") or None)
    wetted = PSVM.wetted_area_m2(case.get("Vessel Orientation") or "Vertical",
                                 od_m, tt_m, case.get("Head Type") or "2:1",
                                 nll_m, grade_m, firelim_m, extra_m2)
    return PSVM.fire_q_w(wetted, f)


# ════════════════════════════════════════════════════════════════════════
#  Composition feed (reporting only — real K-value flash via hydraulics.py)
# ════════════════════════════════════════════════════════════════════════
def _load_feed(case: dict, base_dir: str):
    stream = case.get("Stream Lookup")
    hmb_file = case.get("HMB File")
    if not stream or not hmb_file:
        return None
    hmb_path = hmb_file if os.path.isabs(hmb_file) else os.path.join(base_dir, hmb_file)
    if not os.path.exists(hmb_path):
        return None
    try:
        is_cs = HYD.is_proii_export(hmb_path)
        return HYD.read_feed(hmb_path, stream, case.get("Case") or "Case 1", None, is_cs)
    except Exception:
        return None


# ════════════════════════════════════════════════════════════════════════
#  Time-march engine
# ════════════════════════════════════════════════════════════════════════
@dataclass
class TimeRow:
    t_s: float
    p_psia: float
    t_f: float
    mdot_lbhr: float
    qvol_ft3hr: float
    builtup_bp_psia: float
    choked: bool
    comp: dict | None = None


@dataclass
class CaseResult:
    rows: list[TimeRow] = field(default_factory=list)
    vessel_vol_ft3: float = 0.0
    line_vol_ft3: float = 0.0
    total_vol_ft3: float = 0.0
    peak_builtup_bp_psia: float = 0.0
    feed_names: list[str] | None = None


def run_case(case: dict, pipeline_rows: list[dict], base_dir: str = ".") -> CaseResult:
    mw = case.get("Vap MW") or 28.96
    z = case.get("Vap Z") or 1.0
    k = max(1.01, case.get("Vap Cp/Cv (k)") or 1.2)
    mu_pas = (case.get("Vap Visc (cP)") or 0.012) * 1e-3

    vessel_vol_ft3 = case.get("Vessel Volume") or 0.0
    line_no = (case.get("Tailpipe Line No") or "").strip() or None
    line_vol_ft3 = line_volume_ft3(pipeline_rows, line_no)
    v_ft3 = vessel_vol_ft3 + line_vol_ft3
    v_m3 = _ft3_to_m3(v_ft3)

    superimposed_bp_pa = _to_pa(case.get("Superimposed BP") or 14.696)
    p_pa = _to_pa(case.get("Initial Pressure") or 0.0)
    t_k = _to_k(case.get("Initial Temperature") or 60.0)
    tref_k = t_k

    cp_si = (PSVM.R_UNIV / mw) * k / (k - 1.0)
    cv_si = cp_si / k
    m_kg = p_pa * v_m3 * mw / (z * PSVM.R_UNIV * t_k)

    dt = case.get("Timestep") or 5.0
    max_t = case.get("Max Duration") or 900.0
    q_w = fire_heat_w(case)

    feed = _load_feed(case, base_dir)

    result = CaseResult(vessel_vol_ft3=vessel_vol_ft3, line_vol_ft3=line_vol_ft3,
                        total_vol_ft3=v_ft3,
                        feed_names=list(feed.names) if feed is not None else None)

    builtup_prev_pa = 0.0
    t = 0.0
    while t <= max_t + 1e-6 and m_kg > 1e-9:
        p_back_pa = superimposed_bp_pa + builtup_prev_pa
        mdot_kgs = relief_mdot_kgs(case, p_pa, p_back_pa, t_k)
        builtup_pa = solve_builtup_bp_pa(pipeline_rows, line_no, mdot_kgs, t_k, mw, z,
                                         mu_pas, superimposed_bp_pa, p_pa)
        p_back_pa = superimposed_bp_pa + builtup_pa
        mdot_kgs = relief_mdot_kgs(case, p_pa, p_back_pa, t_k)   # refine once
        builtup_prev_pa = builtup_pa

        _, choked = PSVM.gas_mass_flux(p_pa, p_back_pa, t_k, mw, z, k)
        rho = max(1e-9, p_pa * mw / (z * PSVM.R_UNIV * t_k))
        qvol_ft3hr = _m3_to_ft3(mdot_kgs / rho) * 3600.0 if mdot_kgs > 0 else 0.0
        builtup_psia = builtup_pa / HYD.PSIA_TO_PA
        result.peak_builtup_bp_psia = max(result.peak_builtup_bp_psia, builtup_psia)

        comp = None
        if feed is not None:
            try:
                fr = HYD.flash(feed, _to_psia(p_pa), _to_f(t_k))
                comp = dict(zip(feed.names, fr.y))
            except Exception:
                comp = None

        result.rows.append(TimeRow(
            t_s=t, p_psia=_to_psia(p_pa), t_f=_to_f(t_k),
            mdot_lbhr=_kgs_to_lbhr(mdot_kgs), qvol_ft3hr=qvol_ft3hr,
            builtup_bp_psia=builtup_psia, choked=choked, comp=comp,
        ))

        if mdot_kgs <= 0.0 or t >= max_t:
            break

        h_t = (PSVM.R_UNIV / mw) * tref_k + cp_si * (t_k - tref_k)
        u_t = cv_si * (t_k - tref_k)
        u1 = m_kg * u_t
        du = -mdot_kgs * dt * h_t + q_w * dt
        m2_kg = max(1e-9, m_kg - mdot_kgs * dt)
        u2 = u1 + du
        t2_k = tref_k + u2 / (m2_kg * cv_si)
        p2_pa = max(1.0, z * (m2_kg / mw) * PSVM.R_UNIV * t2_k / v_m3)

        m_kg, p_pa, t_k = m2_kg, p2_pa, t2_k
        t += dt

    return result


def _interp_p_at_t(rows: list[TimeRow], target_t: float) -> float | None:
    if not rows:
        return None
    if target_t <= rows[0].t_s:
        return rows[0].p_psia
    for a, b in zip(rows[:-1], rows[1:]):
        if a.t_s <= target_t <= b.t_s:
            if b.t_s == a.t_s:
                return a.p_psia
            frac = (target_t - a.t_s) / (b.t_s - a.t_s)
            return a.p_psia + (b.p_psia - a.p_psia) * frac
    return rows[-1].p_psia


def design_checks(case: dict, result: CaseResult) -> list[tuple]:
    checks = []
    for n in (1, 2):
        p_target = case.get(f"Final Pressure {n}")
        t_target = case.get(f"Final Time {n}")
        if p_target is None or t_target is None:
            continue
        p_actual = _interp_p_at_t(result.rows, t_target)
        verdict = None
        if p_actual is not None:
            verdict = "PASS" if p_actual <= p_target else "FAIL"
        checks.append((n, t_target, p_target, p_actual, verdict))
    return checks


# ════════════════════════════════════════════════════════════════════════
#  Output workbook
# ════════════════════════════════════════════════════════════════════════
def write_output(case: dict, usys: "UN.UnitSystem", result: CaseResult,
                 checks: list[tuple], out_path: str) -> str:
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = PROFILE_SHEET
    ws.sheet_view.showGridLines = False

    comp_names = result.feed_names or []
    headers = ["Time (s)", usys.hdr("Pressure", "P"), usys.hdr("Temperature", "T"),
               usys.hdr("Mass Flow", "mflow"), usys.hdr("Vol Flow", "qvol"),
               usys.hdr("Builtup BP", "P"), "Choked"]
    headers += [f"y[{n}] (mol frac)" for n in comp_names]

    ws.merge_cells(f"A1:{get_column_letter(len(headers))}1")
    C(ws, 1, 1, f"DEPRESSURIZATION PROFILE — {case.get('Case ID', '')}  "
       f"(Vessel {result.vessel_vol_ft3:.1f} ft³ + Line {result.line_vol_ft3:.1f} ft³ "
       f"= {result.total_vol_ft3:.1f} ft³ system volume)", bg=NAVY, fg=DGRAY, bold=True)
    for ci, h in enumerate(headers, 1):
        C(ws, 2, ci, h, bg=LGRAY, bold=True)
    ws.freeze_panes = "A3"

    for ri, row in enumerate(result.rows, 3):
        C(ws, ri, 1, round(row.t_s, 3))
        C(ws, ri, 2, round(usys.from_internal("P", row.p_psia), 4))
        C(ws, ri, 3, round(usys.from_internal("T", row.t_f), 4))
        C(ws, ri, 4, round(usys.from_internal("mflow", row.mdot_lbhr), 4))
        C(ws, ri, 5, round(usys.from_internal("qvol", row.qvol_ft3hr), 4))
        C(ws, ri, 6, round(usys.from_internal("P", row.builtup_bp_psia), 4))
        C(ws, ri, 7, "Yes" if row.choked else "No")
        for ci, n in enumerate(comp_names, 8):
            v = (row.comp or {}).get(n)
            C(ws, ri, ci, round(v, 5) if v is not None else None)

    for col, w in zip("ABCDEFG", (10, 12, 12, 14, 14, 12, 8)):
        ws.column_dimensions[col].width = w

    # ── Design_Checks ───────────────────────────────────────────────
    ws2 = wb.create_sheet(CHECKS_SHEET)
    ws2.sheet_view.showGridLines = False
    rows = [("DESIGN CHECKS", None, None, "sec")]
    if not checks:
        rows.append(("No Final Pressure/Time targets set", "—", "", ""))
    for n, t_target, p_target, p_actual, verdict in checks:
        rows.append((f"Final Time {n}", t_target, "s", ""))
        rows.append((f"Final Pressure {n} (target)",
                     round(usys.from_internal("P", p_target), 4), usys.label("P"), ""))
        rows.append((f"Pressure at Final Time {n} (actual)",
                     round(usys.from_internal("P", p_actual), 4) if p_actual is not None else "—",
                     usys.label("P"), "interpolated from the profile"))
        rows.append((f"Check {n}", verdict, "", "PASS = actual pressure at/below target by that time"))
    rows.append(("Peak Builtup BP over the run",
                 round(usys.from_internal("P", result.peak_builtup_bp_psia), 4),
                 usys.label("P"), ""))
    _result_rows(ws2, 2, rows)

    # ── Notes ───────────────────────────────────────────────────────
    ws3 = wb.create_sheet("Notes")
    ws3.sheet_view.showGridLines = False
    ws3.column_dimensions["A"].width = 110
    lines = [
        ("DEPRESSURIZATION DYNAMICS — RUN NOTES", True),
        ("", False),
        (f"Generated: {datetime.now():%d-%b-%Y %H:%M}", False),
        ("", False),
        ("METHOD", True),
        ("  Single well-mixed vapour gas space = Vessel Volume + tailpipe Line volume "
         "(from Pipeline_Input geometry).", False),
        ("  Adiabatic open-system energy balance each timestep, ideal-gas Cp/Cv from "
         "Vap Cp/Cv (k) and Vap MW; Z held constant.", False),
        ("  Relief flow: API 520 orifice/Cv equation, resized fresh each timestep from "
         "current vessel P, T (no resizing of the device itself).", False),
        ("  Builtup backpressure: self-contained Darcy-Weisbach + fittings march along the "
         "tailpipe, bisected so the tailpipe's own ΔP lands its outlet on Superimposed BP.", False),
        ("  Fire Case = Yes: constant API 521 fire heat input added to the energy balance; "
         "wetted area fixed at its initial value (no liquid-level tracking).", False),
        ("  Composition split (if Stream Lookup/HMB File given): real K-value flash at each "
         "timestep's P, T — REPORTING ONLY, does not feed back into the mass/energy balance.", False),
        ("  Final Pressure/Time 1 & 2: PASS/FAIL design-check targets only.", False),
        ("", False),
        ("SIMPLIFICATIONS (v1)", True),
        ("  No liquid-level / two-phase vessel tracking — vapour-only system.", False),
        ("  Z and Cp/Cv held constant through the run (no dynamic EOS re-solve).", False),
        ("  Wetted area for the fire case held constant (no level recession modelled).", False),
    ]
    for i, (t, bold) in enumerate(lines, 1):
        ws3.cell(i, 1).value = t
        ws3.cell(i, 1).font = Font(name="Calibri", size=9, bold=bold,
                                   color=NAVY if bold else DGRAY)

    UN.add_units_sheet(wb, usys.system, position=len(wb.worksheets))
    wb.save(out_path)
    return out_path


def run(input_path: str, out_path: str | None = None) -> str:
    case, usys = read_case_input(input_path)
    pipeline_rows = HYD.read_pipeline_input(input_path)
    base_dir = os.path.dirname(os.path.abspath(input_path))
    result = run_case(case, pipeline_rows, base_dir)
    checks = design_checks(case, result)
    out_path = out_path or (os.path.splitext(os.path.basename(input_path))[0]
                            + "_depressurization_output.xlsx")
    return write_output(case, usys, result, checks, out_path)


# ════════════════════════════════════════════════════════════════════════
#  CLI entry point
# ════════════════════════════════════════════════════════════════════════
def main():
    argv = sys.argv[1:]
    flags = [a for a in argv if a.startswith("-")]
    args = [a for a in argv if not a.startswith("-")]

    if "--template" in flags or "-t" in flags:
        unit_sys = "SI" if any(f.lower() in ("--si", "-si") for f in flags) else "FPS"
        out = create_input_template(unit_system=unit_sys)
        print(f"  Template created: {out}  (units: {unit_sys})")
        print("  Fill in Depressurization_Case (and Pipeline_Input for the tailpipe) "
              "then re-run without --template.")
        return

    input_path = args[0] if len(args) > 0 else "depressurization_input.xlsx"
    if not os.path.exists(input_path):
        print(f"  Input not found: {input_path}")
        print("  Run with --template to create the blank input workbook.")
        sys.exit(1)

    out_path = args[1] if len(args) > 1 else None
    print(f"\n  Depressurization Dynamics  |  input: {os.path.basename(input_path)}")
    written = run(input_path, out_path)
    print(f"\n  Done — {written}")


if __name__ == "__main__":
    main()
