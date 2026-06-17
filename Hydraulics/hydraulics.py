#!/usr/bin/env python3
"""HyCalign Hydraulics — unified engine (single file to run).

This file merges the former  asme_data.py + hydraulics_engine.py +
hydraulics_engine_flash_noiso.py + flash_vle.py + flow_pattern_data.py +
flow_pattern_maps.py  into ONE engine.  Companion modules:

  pms_classes.py        — piping spec / schedule data  (PROJECT-SPECIFIC)
  comp_constants.py     — component property constants (project data)
  ../common/units.py    — FPS/SI units layer (UNITS sheet in the input workbook)

Run:
    python hydraulics.py --template            # create blank FPS input
    python hydraulics.py --template --si       # create blank SI input
    python hydraulics.py pipeline_input_noiso.xlsx HMB.xlsx
    python hydraulics.py input.xlsx HMB.xlsx --isenthalpic
or simply use  run_hydraulics.py  (menu / GUI front end).
"""
from __future__ import annotations



# ══════════════════════════════════════════════════════════════════════════
# ║  SECTION: ASME B36.10 / B36.19 pipe data (former asme_data.py)
# ══════════════════════════════════════════════════════════════════════════
"""
ASME reference data for the PMS / GEMS parser.

Everything here is standard-derived engineering reference data:
  - ASME B36.10M / B36.19M  pipe OD and wall thickness per schedule
  - ASME B16.5              pressure-temperature ratings (per material group / class)
  - ASME B16.5              flange class decoded from the material-class prefix
  - ASME B31.3              allowable stress (S) vs temperature for the MOC

IMPORTANT (read before trusting numbers):
  These tables are transcribed for convenience so the parser can build complete
  size and P-T tables. They MUST be verified against the current code edition the
  project is contracted to before being used for fabrication. The two P-T anchor
  points for Group 3.17 / Class 150 (230 psig @ 100 deg F, 155 psig @ 450 deg F)
  match the ExxonMobil GEMS G1S sheets; intermediate rows follow the standard
  B16.5 shape. Verify-against-code is enforced by VERIFY_NOTE below, which is
  written into the Excel output.
"""


VERIFY_NOTE = (
    "REFERENCE DATA - verify against the contracted ASME code edition "
    "(B36.10M / B16.5 / B31.3) before use in fabrication or design."
)

IN2MM = 25.4
PSI2BAR = 0.0689476
F_TO_C = lambda f: (f - 32.0) * 5.0 / 9.0  # noqa: E731

# ── ASME B36.10M pipe outside diameter (inches) ────────────────────────────
# Keyed by NPS string. OD is fixed regardless of schedule.
PIPE_OD_IN: dict[str, float] = {
    "1/2": 0.840, "3/4": 1.050, "1": 1.315, "1-1/4": 1.660, "1-1/2": 1.900,
    "2": 2.375, "2-1/2": 2.875, "3": 3.500, "3-1/2": 4.000, "4": 4.500,
    "5": 5.563, "6": 6.625, "8": 8.625, "10": 10.750, "12": 12.750,
    "14": 14.000, "16": 16.000, "18": 18.000, "20": 20.000, "24": 24.000,
    "26": 26.000, "28": 28.000, "30": 30.000, "32": 32.000, "34": 34.000,
    "36": 36.000,
}

# Ordered NPS sequence for range expansion.
NPS_ORDER: list[str] = [
    "1/2", "3/4", "1", "1-1/4", "1-1/2", "2", "2-1/2", "3", "3-1/2", "4",
    "5", "6", "8", "10", "12", "14", "16", "18", "20", "24",
    "26", "28", "30", "32", "34", "36",
]

# NPS -> DN (mm) for cross-referencing the DN column on the GEMS sheets.
NPS_TO_DN: dict[str, int] = {
    "1/2": 15, "3/4": 20, "1": 25, "1-1/4": 32, "1-1/2": 40, "2": 50,
    "2-1/2": 65, "3": 80, "3-1/2": 90, "4": 100, "5": 125, "6": 150,
    "8": 200, "10": 250, "12": 300, "14": 350, "16": 400, "18": 450,
    "20": 500, "24": 600, "26": 650, "28": 700, "30": 750, "32": 800,
    "34": 850, "36": 900,
}

# ── ASME B36.10M / B36.19M wall thickness (inches) ─────────────────────────
# Keyed by NPS, then by schedule designation. S-schedules (10S/40S/80S) are the
# stainless/alloy series used by the Alloy 20 GEMS classes. STD/XS/40/80/160/XXS
# are the carbon series. Only commonly stocked schedules are listed.
PIPE_WALL_IN: dict[str, dict[str, float]] = {
    "1/2":   {"10S": 0.083, "40S": 0.109, "STD": 0.109, "80S": 0.147, "XS": 0.147, "160": 0.188, "XXS": 0.294},
    "3/4":   {"10S": 0.083, "40S": 0.113, "STD": 0.113, "80S": 0.154, "XS": 0.154, "160": 0.219, "XXS": 0.308},
    "1":     {"10S": 0.109, "40S": 0.133, "STD": 0.133, "80S": 0.179, "XS": 0.179, "160": 0.250, "XXS": 0.358},
    "1-1/4": {"10S": 0.109, "40S": 0.140, "STD": 0.140, "80S": 0.191, "XS": 0.191, "160": 0.250, "XXS": 0.382},
    "1-1/2": {"10S": 0.109, "40S": 0.145, "STD": 0.145, "80S": 0.200, "XS": 0.200, "160": 0.281, "XXS": 0.400},
    "2":     {"10S": 0.109, "40S": 0.154, "STD": 0.154, "80S": 0.218, "XS": 0.218, "160": 0.344, "XXS": 0.436},
    "2-1/2": {"10S": 0.120, "40S": 0.203, "STD": 0.203, "80S": 0.276, "XS": 0.276, "160": 0.375, "XXS": 0.552},
    "3":     {"10S": 0.120, "40S": 0.216, "STD": 0.216, "80S": 0.300, "XS": 0.300, "160": 0.438, "XXS": 0.600},
    "3-1/2": {"10S": 0.120, "40S": 0.226, "STD": 0.226, "80S": 0.318, "XS": 0.318},
    "4":     {"10S": 0.120, "40S": 0.237, "STD": 0.237, "80S": 0.337, "XS": 0.337, "120": 0.438, "160": 0.531, "XXS": 0.674},
    "5":     {"10S": 0.134, "40S": 0.258, "STD": 0.258, "80S": 0.375, "XS": 0.375, "160": 0.625, "XXS": 0.750},
    "6":     {"10S": 0.134, "40S": 0.280, "STD": 0.280, "80S": 0.432, "XS": 0.432, "160": 0.719, "XXS": 0.864},
    "8":     {"10S": 0.148, "40S": 0.322, "STD": 0.322, "40": 0.322, "80S": 0.500, "XS": 0.500, "80": 0.500, "160": 0.906, "XXS": 0.875},
    "10":    {"10S": 0.165, "40S": 0.365, "STD": 0.365, "40": 0.365, "80S": 0.500, "XS": 0.500, "80": 0.594, "160": 1.125},
    "12":    {"10S": 0.180, "40S": 0.375, "STD": 0.375, "40": 0.406, "80S": 0.500, "XS": 0.500, "80": 0.688, "160": 1.312},
    "14":    {"10S": 0.250, "STD": 0.375, "40": 0.438, "XS": 0.500, "80": 0.750, "160": 1.406},
    "16":    {"10S": 0.250, "STD": 0.375, "40": 0.500, "XS": 0.500, "80": 0.844, "160": 1.594},
    "18":    {"10S": 0.250, "STD": 0.375, "40": 0.562, "XS": 0.500, "80": 0.938, "160": 1.781},
    "20":    {"10S": 0.250, "STD": 0.375, "40": 0.594, "XS": 0.500, "80": 1.031},
    "24":    {"10S": 0.250, "STD": 0.375, "40": 0.688, "XS": 0.500, "80": 1.219},
    # 26"-36": only STD/XS are verified (ASME B36.10 holds these constant
    # through the large-bore range); other schedules intentionally omitted
    # rather than guessed.
    "26":    {"STD": 0.375, "XS": 0.500},
    "28":    {"STD": 0.375, "XS": 0.500},
    "30":    {"STD": 0.375, "XS": 0.500},
    "32":    {"STD": 0.375, "XS": 0.500},
    "34":    {"STD": 0.375, "XS": 0.500},
    "36":    {"STD": 0.375, "XS": 0.500},
}

# Commercial schedules ordered thinnest -> thickest, used to round a computed
# (CAL) required thickness UP to the next available wall.
SCHEDULE_LADDER: list[str] = ["10S", "STD", "40", "XS", "80", "120", "160", "XXS"]

# Equivalent schedule designations to try when an exact match is absent. The
# stainless S-series equals the carbon series up to mid sizes; this only fills a
# lookup, it does not assert dimensional equality at every NPS.
SCHEDULE_ALIASES: dict[str, list[str]] = {
    "40": ["40", "STD", "40S"], "STD": ["STD", "40", "40S"], "40S": ["40S", "STD", "40"],
    "80": ["80", "XS", "80S"], "XS": ["XS", "80", "80S"], "80S": ["80S", "XS", "80"],
    "10S": ["10S", "10"], "10": ["10", "10S"],
}

# ── ASME B16.5 flange class decoded from material-class prefix ──────────────
# Class name carries the flange rating in the leading digits after 'G':
#   G1.. -> 150,  G3.. -> 300,  G6.. -> 600,  G9.. -> 900,  G15.. -> 1500,
#   G25.. -> 2500.  Example: "G1S-2" -> 150, "G25S-1" -> 2500.
FLANGE_CLASS_BY_PREFIX: dict[int, int] = {
    1: 150, 3: 300, 6: 600, 9: 900, 15: 1500, 25: 2500,
}

# ── ASME B16.5 pressure-temperature ratings ────────────────────────────────
# Keyed by material group, then flange class, then a list of (temp_F, psig).
# Group 3.17 is the GEMS Alloy 20 (UNS N08020) group. Anchor rows (100/450 F)
# match the printed GEMS limits; remaining rows follow the standard B16.5 curve.
B16_5_PT_RATINGS: dict[str, dict[int, list[tuple[float, float]]]] = {
    "3.17": {
        150: [
            (-20, 230), (100, 230), (200, 195), (300, 175), (400, 160),
            (450, 155), (500, 150), (600, 140), (650, 125), (700, 110),
            (750, 95), (800, 80),
        ],
        300: [
            (-20, 600), (100, 600), (200, 540), (300, 495), (400, 460),
            (450, 445), (500, 430), (600, 400), (650, 385), (700, 360),
            (750, 320), (800, 275),
        ],
    },
}

# ── ASME B31.3 Table A-1 basic allowable stress S (ksi) vs temperature (F) ──
# Keyed by MOC tag, then list of (temp_F, S_ksi). N08020 = Alloy 20 seamless
# pipe (ASME SB-729). Representative published values - verify against code.
ALLOWABLE_STRESS_KSI: dict[str, list[tuple[float, float]]] = {
    "N08020": [
        (100, 23.3), (200, 21.5), (300, 19.6), (400, 18.6),
        (500, 17.8), (600, 17.3), (700, 17.0), (800, 16.7),
    ],
}


# ── helpers ────────────────────────────────────────────────────────────────
def interp(points: list[tuple[float, float]], x: float) -> tuple[float, bool]:
    """Linear interpolate y at x over (x,y) points sorted by x.

    Returns (value, extrapolated). Clamps and flags when x is outside the
    table range so callers can mark extrapolated rows.
    """
    pts = sorted(points, key=lambda p: p[0])
    if x <= pts[0][0]:
        return pts[0][1], x < pts[0][0]
    if x >= pts[-1][0]:
        return pts[-1][1], x > pts[-1][0]
    for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
        if x0 <= x <= x1:
            if x1 == x0:
                return y0, False
            frac = (x - x0) / (x1 - x0)
            return y0 + frac * (y1 - y0), False
    return pts[-1][1], True


def expand_nps_range(low: str, high: str) -> list[str]:
    """Return the ordered NPS list from low..high inclusive (B36.10 order)."""
    if low not in NPS_ORDER or high not in NPS_ORDER:
        return []
    i, j = NPS_ORDER.index(low), NPS_ORDER.index(high)
    if i > j:
        i, j = j, i
    return NPS_ORDER[i:j + 1]


def wall_for(nps: str, schedule: str) -> float | None:
    """Look up wall thickness (in) for an NPS + schedule, or None if absent.

    Tries the exact schedule first, then equivalent designations so common
    carbon/stainless aliases (40/STD/40S, 80/XS/80S) resolve.
    """
    walls = PIPE_WALL_IN.get(nps, {})
    if schedule in walls:
        return walls[schedule]
    for alt in SCHEDULE_ALIASES.get(schedule, []):
        if alt in walls:
            return walls[alt]
    return None


def next_schedule_wall(nps: str, required_in: float) -> tuple[str, float] | None:
    """Smallest commercial schedule whose wall >= required_in for this NPS."""
    walls = PIPE_WALL_IN.get(nps, {})
    best: tuple[str, float] | None = None
    for sched in SCHEDULE_LADDER:
        w = walls.get(sched)
        if w is None or w < required_in:
            continue
        if best is None or w < best[1]:
            best = (sched, w)
    return best


def flange_class_for(class_name: str) -> int | None:
    """Decode ASME B16.5 flange class from the material-class name prefix."""
    import re
    m = re.match(r"^G(\d+)", class_name.strip().upper())
    if not m:
        return None
    return FLANGE_CLASS_BY_PREFIX.get(int(m.group(1)))


# ══════════════════════════════════════════════════════════════════════════
# ║  SECTION: Core hydraulics engine (former hydraulics_engine.py)
# ══════════════════════════════════════════════════════════════════════════
#!/usr/bin/env python3
"""
HyCalign Hydraulics Engine  v3  —  Pressure Profile
===================================================
Pipeline:
  1. Parse an ISOGEN .pcf with the project PCF parser (pcf_parser.py).
  2. Resolve every component's INTERNAL DIAMETER from the piping spec
     (e.g. G1A-5) + bore, using the project PMS catalogue (pms_classes.py)
     and ASME B36.10/B36.19 tables (asme_data.py).  No manual ID entry.
  3. Pull fluid properties for one stream from an HMB workbook (HMB.xlsx).
  4. March a PRESSURE PROFILE along the ordered components.

Phase isolation (per project requirement):
  The marcher detects the stream phase ONCE (vapor / liquid / two-phase) and
  dispatches to exactly ONE of three independent calculators:
      _dp_vapor      compressible gas, density updated with local pressure
      _dp_liquid     incompressible Darcy-Weisbach, constant density
      _dp_twophase   homogeneous no-slip (McAdams viscosity) two-phase model
  The three paths share no state and never fall through into one another, so a
  vapor calc can never be contaminated by liquid logic (or vice-versa).

When run, the engine ALWAYS prompts for the line number and the stream number
to map.  Example mapping: line ER-162  <->  stream T801-OH.

Output: one workbook with
  Pressure_Profile   station table (cum length, ID, velocity, Re, f, dP, P) + chart
  Component_Detail   full per-component hydraulics
  Stream_Props       extracted HMB properties (with units + phase decision)
  PCF_Components      the parsed ER-162 line (pcf_parser house style)
  README             method + assumptions

Usage:
    python3 hydraulics_engine.py                 # prompts for everything
    python3 hydraulics_engine.py ER-162-*.pcf HMB.xlsx
"""

import os, re, sys, math, glob
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference, Series

import pms_classes as PMS
import sys as _sys
A = _sys.modules[__name__]   # asme_data merged into this file

# ── Shared Procalc units layer (../common/units.py) ────────────────────────
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "common"))
import units as UN

_U = UN.UnitSystem("FPS")    # module-level display-units context


def set_units(usys: "UN.UnitSystem") -> None:
    """Set the display unit system for all sheet builders in this module."""
    global _U
    _U = usys or UN.UnitSystem("FPS")

# ── Unit conversions (to SI unless noted) ──────────────────────────────────
G_SI          = 9.80665            # m/s^2
IN_TO_M       = 0.0254
FT_TO_M       = 0.3048
LBHR_TO_KGS   = 0.45359237 / 3600.0
LBFT3_TO_KGM3 = 16.018463
CP_TO_PAS     = 1e-3
PA_TO_PSI     = 1.0 / 6894.757293
PSIA_TO_PA    = 6894.757293
ROUGH_STEEL_M = 0.0000457           # 0.0018 in commercial steel, in metres
R_UNIV        = 8.314462618         # J/(mol*K)
F_TO_K_OFFSET = 459.67              # degF -> degR additive; K = degR / 1.8

# ── Style ──────────────────────────────────────────────────────────────────
NAVY="1F4973"; STEEL="2E75B6"; WHITE="FFFFFF"; LGRAY="F2F2F2"; DGRAY="404040"
ORANGE="C55A11"; AMBER="FFF2CC"; GREEN="E2EFDA"; GRNHDR="375623"; RED="FCE4D6"
LTBLUE="DEEAF1"; PURPLE="7030A0"; MGRAY="A0A0A0"

def _fill(h): return PatternFill("solid", fgColor=h)
def _thin():
    s = Side(style="thin", color="D0D0D0")
    return Border(left=s, right=s, top=s, bottom=s)

def _c(ws, r, c, v=None, bg=None, fg=DGRAY, sz=9, bold=False, italic=False,
       wrap=False, ha="left", va="center", border=True):
    from openpyxl.cell.cell import MergedCell
    cell = ws.cell(row=r, column=c)
    if isinstance(cell, MergedCell):
        return cell
    if v is not None:
        cell.value = v
    cell.font = Font(name="Calibri", size=sz, bold=bold, italic=italic, color=fg)
    if bg:
        cell.fill = _fill(bg)
    cell.alignment = Alignment(horizontal=ha, vertical=va, wrap_text=wrap)
    if border:
        cell.border = _thin()
    return cell

def cw(ws, col, w):
    ws.column_dimensions[get_column_letter(col) if isinstance(col, int) else col].width = w

def num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return None if (isinstance(v, float) and math.isnan(v)) else float(v)
    s = str(v).strip()
    if not s or s.upper() in {"N/A", "NA", "NONE", "MISSING"}:
        return None
    m = re.search(r"[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?", s)
    return float(m.group(0)) if m else None

def txt(v):
    if v is None:
        return None
    s = str(v).strip()
    return s or None


# ════════════════════════════════════════════════════════════════════════
#  1.  INTERNAL DIAMETER from piping spec + bore (PMS catalogue + ASME)
# ════════════════════════════════════════════════════════════════════════
_CLASS_BY_NAME = {c.name.upper(): c for c in PMS.CLASSES}

# Map a numeric bore to the canonical NPS string used by asme_data.
_NPS_BY_FLOAT = {
    0.5: "1/2", 0.75: "3/4", 1.0: "1", 1.25: "1-1/4", 1.5: "1-1/2", 2.0: "2",
    2.5: "2-1/2", 3.0: "3", 3.5: "3-1/2", 4.0: "4", 5.0: "5", 6.0: "6",
    8.0: "8", 10.0: "10", 12.0: "12", 14.0: "14", 16.0: "16", 18.0: "18",
    20.0: "20", 24.0: "24", 26.0: "26", 28.0: "28", 30.0: "30", 32.0: "32",
    34.0: "34", 36.0: "36",
}

def _nps_string(bore) -> str | None:
    if bore is None:
        return None
    s = str(bore).strip()
    if s in A.NPS_ORDER:
        return s
    f = num(s)
    if f is None:
        return None
    best = min(_NPS_BY_FLOAT, key=lambda k: abs(k - f))
    return _NPS_BY_FLOAT[best] if abs(best - f) < 0.13 else None


@dataclass
class IDResult:
    id_in: float | None
    od_in: float | None
    wall_in: float | None
    schedule: str | None
    nps: str | None
    basis: str


def resolve_id(spec: str | None, bore) -> IDResult:
    """Internal diameter (in) for a component from its piping spec + bore.

    schedule comes from the PMS class pipe-rule whose NPS range covers the
    bore; wall from ASME B36.10/B36.19.  ID = OD - 2*wall.  CAL rules fall
    back to their min_schedule wall (handling/structural floor).
    """
    nps = _nps_string(bore)
    if nps is None or nps not in A.PIPE_OD_IN:
        return IDResult(None, None, None, None, nps, "bore not recognised")
    od = A.PIPE_OD_IN[nps]

    cls = _CLASS_BY_NAME.get((spec or "").upper())
    sched = None
    if cls:
        for rule in cls.pipe_rules:
            if nps in A.expand_nps_range(rule.nps_low, rule.nps_high):
                sched = rule.schedule
                if str(sched).upper() == "CAL":
                    sched = (rule.min_schedule or "STD")
                break

    basis = f"{spec} {nps}\""
    if sched is None:
        sched = "STD"
        basis += " (spec n/a -> STD)"
    wall = A.wall_for(nps, str(sched).upper())
    if wall is None:
        wall = A.wall_for(nps, "STD")
        basis += " (sched n/a -> STD)"
    if wall is None:
        return IDResult(None, od, None, sched, nps, basis + " (no wall)")
    return IDResult(round(od - 2 * wall, 4), od, wall, str(sched).upper(), nps, basis)


# ════════════════════════════════════════════════════════════════════════
#  2.  HMB stream extraction  (section- and unit-aware)
# ════════════════════════════════════════════════════════════════════════
from openpyxl import load_workbook

try:
    import hmb_proii_reader as HMBPROII   # optional — PRO/II case-sheet exports
except ImportError:                        # pragma: no cover
    HMBPROII = None
try:
    import stream_map as SMAP              # optional — line→stream mapping
except ImportError:                        # pragma: no cover
    SMAP = None

SKIP_SHEETS = {"UNITS", "COMPONENTS", "INPUT", "OUTPUT", "Summary",
               "Component Index", "Legend", "CASES", "README"}


def is_proii_export(path: str | Path) -> bool:
    """True only for the transposed PRO/II export / clean built HMB (Case N sheets).

    Content-based — the filename is NOT used.  A td_parser per-stream dump (which
    may also be named ``HMB_..._from_proii_*.xlsx``) has UNITS/COMPONENTS/INPUT/
    OUTPUT + one sheet per stream and NO ``Case 1`` sheet, so it correctly routes
    to the per-stream ``extract_stream`` reader instead.
    """
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        sheets = set(wb.sheetnames)
        wb.close()
        return "Case 1" in sheets or "Case 2" in sheets
    except Exception:
        return False


@dataclass
class StreamProps:
    stream: str
    phase: str | None = None
    temp_f: float | None = None
    pres_psia: float | None = None
    total_mass: float | None = None     # LB/HR
    vap_mass: float | None = None       # LB/HR
    liq_mass: float | None = None       # LB/HR
    mol_weight: float | None = None
    vap_density: float | None = None    # LB/FT3
    liq_density: float | None = None    # LB/FT3
    vap_visc: float | None = None       # cP
    liq_visc: float | None = None       # cP
    vap_z: float | None = None
    liq_surf_tens: float | None = None  # dyn/cm
    source_sheet: str | None = None
    warned: bool = False
    # ── extended (HMB_from_proii) for the expanded pressure profile ──────
    total_density: float | None = None  # LB/FT3 (total actual)
    total_z: float | None = None
    vap_mw: float | None = None
    liq_mw: float | None = None
    vap_cp: float | None = None         # BTU/LB-F
    liq_cp: float | None = None
    vap_cp_cv: float | None = None      # Cp/Cv ratio (actual gamma)
    vap_therm_cond: float | None = None  # BTU/HR-FT-F
    liq_therm_cond: float | None = None
    vap_sp_enthalpy: float | None = None  # BTU/LB
    liq_sp_enthalpy: float | None = None
    total_std_liq: float | None = None  # FT3/HR
    total_std_vap: float | None = None  # FT3/HR
    vap_act_vol: float | None = None    # FT3/HR (gas volumetric flow)
    liq_act_rate: float | None = None   # FT3/HR (liquid volumetric flow)
    acentric: float | None = None
    tc_f: float | None = None
    pc_psia: float | None = None

    def total_viscosity_cp(self) -> float | None:
        """Single value for the profile: phase visc, or mass-weighted for 2-φ."""
        vm, lm = self.vap_mass or 0.0, self.liq_mass or 0.0
        if vm > 0 and lm > 0 and self.vap_visc and self.liq_visc:
            tot = vm + lm
            return (vm * self.vap_visc + lm * self.liq_visc) / tot
        return self.vap_visc if (vm >= lm) else self.liq_visc

    def gamma_estimate(self) -> float | None:
        """Ideal-gas Cp/Cv from Cp and MW: γ = Cp/(Cp - R/MW)."""
        cp, mw = self.vap_cp, self.vap_mw or self.mol_weight
        if cp and mw and mw > 0:
            r_btu = 1.98588 / mw          # R = 1.98588 BTU/(lb-mol·°F) -> per lb
            denom = cp - r_btu
            if denom > 0:
                return cp / denom
        return None


def streamprops_from_hmb(hs) -> StreamProps:
    """Build a StreamProps from a hmb_proii_reader.HMBStream (authoritative)."""
    g = hs.props.get
    phase = hs.phase or ""
    sp = StreamProps(
        stream=hs.name,
        phase=phase,
        source_sheet=f"HMB_from_proii[{hs.case}]",
        temp_f=g("temp_f"),
        pres_psia=g("pres_psia"),
        total_mass=g("total_mass"),
        vap_mass=g("vap_mass"),
        liq_mass=g("liq_mass"),
        mol_weight=g("mol_weight"),
        vap_density=g("vap_density"),
        liq_density=g("liq_density"),
        vap_visc=g("vap_visc"),
        liq_visc=g("liq_visc"),
        vap_z=g("vap_z") or g("total_z"),
        liq_surf_tens=g("liq_surf_tens"),
        total_density=g("total_density"),
        total_z=g("total_z"),
        vap_mw=g("vap_mw"),
        liq_mw=g("liq_mw"),
        vap_cp=g("vap_cp"),
        liq_cp=g("liq_cp"),
        vap_cp_cv=g("vap_cp_cv"),
        vap_therm_cond=g("vap_therm_cond"),
        liq_therm_cond=g("liq_therm_cond"),
        vap_sp_enthalpy=g("vap_sp_enthalpy"),
        liq_sp_enthalpy=g("liq_sp_enthalpy"),
        total_std_liq=g("total_std_liq"),
        total_std_vap=g("total_std_vap"),
        vap_act_vol=g("vap_act_vol"),
        liq_act_rate=g("liq_act_rate"),
        acentric=g("acentric"),
        tc_f=g("tc_f"),
        pc_psia=g("pc_psia"),
    )
    # Split total mass by phase if a phase mass is missing (single-phase export).
    if sp.total_mass is not None:
        p = phase.upper()
        if sp.vap_mass is None and sp.liq_mass is None:
            if "LIQUID" in p and "VAPOR" not in p:
                sp.liq_mass, sp.vap_mass = sp.total_mass, 0.0
            else:
                sp.vap_mass, sp.liq_mass = sp.total_mass, 0.0
        elif sp.vap_mass is not None and sp.liq_mass is None:
            sp.liq_mass = 0.0
        elif sp.liq_mass is not None and sp.vap_mass is None:
            sp.vap_mass = 0.0
    return sp


def _conv_density(val, unit):
    """Return LB/FT3.  HMB prints vapor density as 'LB/M FT3' = per 1000 ft3."""
    if val is None:
        return None
    u = (unit or "").upper().replace(" ", "")
    if "M FT3".replace(" ", "") in u or "/MFT3" in u:
        return val / 1000.0
    return val


def _conv_massflow(val, unit):
    """Return LB/HR.  HMB prints mass flow as 'M LB/HR' = thousands of LB/HR."""
    if val is None:
        return None
    u = (unit or "").upper()
    if "M LB" in u:
        return val * 1000.0
    return val


def _stream_title(ws):
    t = txt(ws["B1"].value) or ""
    m = re.search(r"STREAM:\s*([^|]+)", t)
    return (m.group(1).strip() if m else None), t


def list_streams(path: Path) -> list[str]:
    wb = load_workbook(path, read_only=True, data_only=True)
    out = []
    for ws in wb.worksheets:
        if ws.title in SKIP_SHEETS:
            continue
        name, _ = _stream_title(ws)
        if name:
            out.append(name)
    wb.close()
    return out


def extract_stream(path: Path, stream_name: str) -> StreamProps | None:
    wb = load_workbook(path, read_only=True, data_only=True)
    target = None
    for ws in wb.worksheets:
        if ws.title in SKIP_SHEETS:
            continue
        name, _ = _stream_title(ws)
        if name and name.upper() == stream_name.upper():
            target = ws
            break
    if target is None:
        wb.close()
        return None
    sp = _extract_from_sheet(target)
    wb.close()
    return sp


def _extract_from_sheet(ws) -> StreamProps:
    name, title = _stream_title(ws)
    ph = re.search(r"Phase:\s*([^|]+)", title)
    sp = StreamProps(stream=name or ws.title, source_sheet=ws.title,
                     phase=ph.group(1).strip() if ph and ph.group(1).strip() else None,
                     warned=("UNCONVERGED" in title.upper()))

    rows = list(ws.iter_rows(min_row=2, values_only=True))
    if not rows:
        return sp
    header = rows[0]
    # Phase columns start at index 3 (TOTAL, VAPOR, LIQUID, ...). Index 2 = Unit.
    col_of = {}
    for ci, h in enumerate(header):
        if ci >= 3 and h:
            col_of[str(h).strip().upper()] = ci
    TOT, VAP, LIQ = ["TOTAL"], ["VAPOR"], ["LIQUID"]

    def cell(rv, prefer):
        for key in prefer:
            ci = col_of.get(key)
            if ci is not None and ci < len(rv):
                v = num(rv[ci])
                if v is not None:
                    return v
        return None

    section = ""
    for rv in rows[1:]:
        if not rv or len(rv) < 2 or rv[1] is None:
            continue
        label = str(rv[1]).strip()
        unit = str(rv[2]).strip() if len(rv) > 2 and rv[2] is not None else ""
        if re.match(r"^\d+[.\s]", label):
            section = label.upper()
            continue
        ll = label.lower()

        if "FLOW RATE" in section:
            # td_parser puts mass as one "Mass Rate" row with TOTAL/VAPOR/LIQUID
            # columns; older dumps use separate "... mass flow" rows.
            if ll == "mass rate" or "total mass flow" in ll:
                sp.total_mass = _conv_massflow(cell(rv, TOT), unit)
                v = _conv_massflow(cell(rv, VAP), unit)
                if v is not None:
                    sp.vap_mass = v
                lq = _conv_massflow(cell(rv, LIQ), unit)
                if lq is not None:
                    sp.liq_mass = lq
            elif "vapor mass" in ll or ("vapor" in ll and "mass flow" in ll):
                sp.vap_mass = _conv_massflow(cell(rv, VAP + TOT), unit)
            elif "liquid mass" in ll or ("liquid" in ll and "mass flow" in ll):
                sp.liq_mass = _conv_massflow(cell(rv, LIQ + TOT), unit)
            elif "actual vol" in ll:
                if sp.vap_act_vol is None:
                    sp.vap_act_vol = cell(rv, VAP)
                if sp.liq_act_rate is None:
                    sp.liq_act_rate = cell(rv, LIQ)
            elif "std liq" in ll and sp.total_std_liq is None:
                sp.total_std_liq = cell(rv, TOT)
            elif "std vap" in ll and sp.total_std_vap is None:
                sp.total_std_vap = cell(rv, TOT)
        elif "CONDITION" in section:
            if "temperature" in ll and sp.temp_f is None:
                sp.temp_f = cell(rv, TOT)
            elif "pressure" in ll and sp.pres_psia is None:
                sp.pres_psia = cell(rv, TOT)
            elif "molecular weight" in ll:
                if sp.mol_weight is None:
                    sp.mol_weight = cell(rv, TOT + VAP)
                if sp.vap_mw is None:
                    sp.vap_mw = cell(rv, VAP)
                if sp.liq_mw is None:
                    sp.liq_mw = cell(rv, LIQ)
            elif "specific enthalpy" in ll:
                if sp.vap_sp_enthalpy is None:
                    sp.vap_sp_enthalpy = cell(rv, VAP)
                if sp.liq_sp_enthalpy is None:
                    sp.liq_sp_enthalpy = cell(rv, LIQ)
        elif "VAPOR PHASE" in section:
            if "actual density" in ll or (ll == "density" and sp.vap_density is None):
                sp.vap_density = _conv_density(cell(rv, VAP + TOT), unit)
            elif "viscosity" in ll and "kin" not in ll:
                sp.vap_visc = cell(rv, VAP + TOT)
            elif "z-factor" in ll or "z factor" in ll:
                if sp.vap_z is None:
                    sp.vap_z = cell(rv, VAP + TOT)
                if sp.total_z is None:
                    sp.total_z = cell(rv, TOT)
            elif "thermal cond" in ll and sp.vap_therm_cond is None:
                sp.vap_therm_cond = cell(rv, VAP + TOT)
            elif "cp/cv" in ll and sp.vap_cp_cv is None:
                sp.vap_cp_cv = cell(rv, VAP + TOT)
            elif ll == "cp" and sp.vap_cp is None:
                sp.vap_cp = cell(rv, VAP + TOT)
        elif "LIQUID PHASE" in section:
            if "actual density" in ll or (ll == "density" and sp.liq_density is None):
                sp.liq_density = _conv_density(cell(rv, LIQ + TOT), unit)
            elif "viscosity" in ll and "kin" not in ll:
                sp.liq_visc = cell(rv, LIQ + TOT)
            elif "surface tension" in ll:
                sp.liq_surf_tens = cell(rv, LIQ + TOT)
            elif "thermal cond" in ll and sp.liq_therm_cond is None:
                sp.liq_therm_cond = cell(rv, LIQ + TOT)
            elif ll == "cp" and sp.liq_cp is None:
                sp.liq_cp = cell(rv, LIQ + TOT)
        elif "CRITICAL" in section:
            if "critical temperature" in ll and "kay" not in ll and sp.tc_f is None:
                sp.tc_f = cell(rv, TOT)
            elif "critical pressure" in ll and "kay" not in ll and sp.pc_psia is None:
                sp.pc_psia = cell(rv, TOT)
            elif "acentric" in ll and sp.acentric is None:
                sp.acentric = cell(rv, TOT)
            elif "actual density" in ll and sp.total_density is None:
                sp.total_density = _conv_density(cell(rv, TOT), unit)
            elif "molecular weight" in ll and sp.mol_weight is None:
                sp.mol_weight = cell(rv, TOT)

    # Split total mass by phase if only the total is printed.
    if sp.total_mass is not None:
        p = (sp.phase or "").upper()
        if sp.vap_mass is None and sp.liq_mass is None:
            if "LIQUID" in p and "VAPOR" not in p:
                sp.liq_mass, sp.vap_mass = sp.total_mass, 0.0
            else:
                sp.vap_mass, sp.liq_mass = sp.total_mass, 0.0
        elif sp.vap_mass is not None and sp.liq_mass is None:
            sp.liq_mass = max(0.0, sp.total_mass - sp.vap_mass)
        elif sp.liq_mass is not None and sp.vap_mass is None:
            sp.vap_mass = max(0.0, sp.total_mass - sp.liq_mass)
    return sp


# ════════════════════════════════════════════════════════════════════════
#  3.  PHASE DECISION
# ════════════════════════════════════════════════════════════════════════
def classify_phase(sp: StreamProps) -> str:
    """Return 'VAPOR', 'LIQUID', or 'TWO-PHASE' — drives the calc dispatch."""
    vm = sp.vap_mass or 0.0
    lm = sp.liq_mass or 0.0
    tot = vm + lm
    have_v = sp.vap_density is not None
    have_l = sp.liq_density is not None
    if tot > 0:
        x = vm / tot
        if have_v and have_l and 0.005 < x < 0.995:
            return "TWO-PHASE"
        if x >= 0.5:
            return "VAPOR" if have_v else ("LIQUID" if have_l else "VAPOR")
        return "LIQUID" if have_l else ("VAPOR" if have_v else "LIQUID")
    p = (sp.phase or "").upper()
    if "LIQUID" in p and "VAPOR" not in p:
        return "LIQUID"
    if "MIX" in p or ("VAPOR" in p and "LIQUID" in p):
        return "TWO-PHASE"
    return "VAPOR"


# ════════════════════════════════════════════════════════════════════════
#  4.  HYDRAULICS — three independent phase calculators + marcher
# ════════════════════════════════════════════════════════════════════════
def darcy_f(Re: float, eps_over_d: float) -> float:
    """Darcy friction factor: laminar 64/Re, else Swamee-Jain."""
    if Re <= 0:
        return 0.0
    if Re < 2000:
        return 64.0 / Re
    return 0.25 / (math.log10(eps_over_d / 3.7 + 5.74 / Re ** 0.9)) ** 2


def k_fitting(name: str | None) -> float:
    f = (name or "").lower()
    if "pipe" in f or "weld" in f or "support" in f:
        return 0.0
    if "elbow" in f:
        return 0.30        # LR 90 (a 45 is less; conservative single value)
    if "tee" in f or "stub" in f:
        return 1.0
    if "valve" in f:
        return 3.0         # gate valve open ~0.15; full-port relief path conservative
    if "reducer" in f:
        return 0.30
    if "flange" in f:
        return 0.0
    return 0.0


@dataclass
class Geom:
    id_m: float
    area_m2: float
    length_m: float
    dz_m: float
    k: float
    eps_over_d: float


def _common_geom(idr: IDResult, length_ft, dz_ft, fitting) -> Geom | None:
    if idr.id_in is None or idr.id_in <= 0:
        return None
    d = idr.id_in * IN_TO_M
    area = math.pi / 4.0 * d * d
    L = (length_ft or 0.0) * FT_TO_M
    dz = (dz_ft or 0.0) * FT_TO_M
    return Geom(d, area, L, dz, k_fitting(fitting), ROUGH_STEEL_M / d)


def vapor_density_kgm3(sp: StreamProps, p_psia: float) -> float:
    """Real-gas density at the LOCAL pressure: rho = P*MW / (Z*R*T).

    Z and T come from the mapped HMB stream; density is recomputed from the
    equation of state at each station so it tracks the pressure drop.  Falls
    back to isothermal scaling of the HMB-reported density only when MW or T
    is unavailable.
    """
    mw = sp.mol_weight
    t_f = sp.temp_f
    if mw and t_f is not None:
        z = sp.vap_z or 1.0
        t_k = (t_f + F_TO_K_OFFSET) / 1.8
        p_pa = p_psia * PSIA_TO_PA
        mw_kg_per_mol = mw / 1000.0
        return p_pa * mw_kg_per_mol / (z * R_UNIV * t_k)
    # fallback: scale HMB density isothermally with local pressure
    rho_ref = (sp.vap_density or 0.0) * LBFT3_TO_KGM3
    p_ref = sp.pres_psia or p_psia
    return rho_ref * (p_psia / p_ref) if p_ref else rho_ref


def _dp_liquid(g: Geom, sp: StreamProps, p_psia: float) -> dict:
    """Incompressible Darcy-Weisbach.  Density constant; pressure-independent."""
    rho = (sp.liq_density or 0.0) * LBFT3_TO_KGM3
    mu = (sp.liq_visc or 0.0) * CP_TO_PAS
    mdot = ((sp.liq_mass or 0.0) + (sp.vap_mass or 0.0)) * LBHR_TO_KGS
    if rho <= 0 or g.area_m2 <= 0:
        return _zero_dp(rho)
    v = mdot / (rho * g.area_m2)
    Re = rho * v * g.id_m / mu if mu > 0 else 0.0
    f = darcy_f(Re, g.eps_over_d)
    dyn = 0.5 * rho * v * v
    dp_f = f * (g.length_m / g.id_m) * dyn if g.id_m > 0 else 0.0
    dp_k = g.k * dyn
    dp_z = rho * G_SI * g.dz_m
    return _pack(rho, mu, v, Re, f, dp_f, dp_k, dp_z)


def _dp_vapor(g: Geom, sp: StreamProps, p_psia: float) -> dict:
    """Compressible gas.  Density auto-computed from EOS at LOCAL pressure.

    rho = P*MW / (Z*R*T) recomputed per station; velocity rises as P falls.
    """
    rho = vapor_density_kgm3(sp, p_psia)
    mu = (sp.vap_visc or 0.0) * CP_TO_PAS
    mdot = ((sp.vap_mass or 0.0) + (sp.liq_mass or 0.0)) * LBHR_TO_KGS
    if rho <= 0 or g.area_m2 <= 0:
        return _zero_dp(rho)
    v = mdot / (rho * g.area_m2)
    Re = rho * v * g.id_m / mu if mu > 0 else 0.0
    f = darcy_f(Re, g.eps_over_d)
    dyn = 0.5 * rho * v * v
    dp_f = f * (g.length_m / g.id_m) * dyn if g.id_m > 0 else 0.0
    dp_k = g.k * dyn
    dp_z = rho * G_SI * g.dz_m          # small for gas but kept
    return _pack(rho, mu, v, Re, f, dp_f, dp_k, dp_z)


def _dp_twophase(g: Geom, sp: StreamProps, p_psia: float) -> dict:
    """Homogeneous no-slip two-phase model (McAdams viscosity blend).

    rho_ns = 1/(x/rho_g + (1-x)/rho_l) ; 1/mu_ns = x/mu_g + (1-x)/mu_l.
    Uses mass flux G; gas density auto-computed from EOS at local pressure.
    """
    rho_g = vapor_density_kgm3(sp, p_psia)
    rho_l = (sp.liq_density or 0.0) * LBFT3_TO_KGM3
    mu_g = (sp.vap_visc or 0.0) * CP_TO_PAS
    mu_l = (sp.liq_visc or 0.0) * CP_TO_PAS
    vm = (sp.vap_mass or 0.0) * LBHR_TO_KGS
    lm = (sp.liq_mass or 0.0) * LBHR_TO_KGS
    mdot = vm + lm
    if mdot <= 0 or g.area_m2 <= 0 or rho_g <= 0 or rho_l <= 0:
        return _zero_dp(rho_g)
    x = vm / mdot
    rho_ns = 1.0 / (x / rho_g + (1.0 - x) / rho_l)
    mu_ns = 1.0 / (x / mu_g + (1.0 - x) / mu_l) if (mu_g > 0 and mu_l > 0) else (mu_g or mu_l)
    Gflux = mdot / g.area_m2
    v = Gflux / rho_ns
    Re = Gflux * g.id_m / mu_ns if mu_ns > 0 else 0.0
    f = darcy_f(Re, g.eps_over_d)
    dp_f = f * (g.length_m / g.id_m) * (Gflux * Gflux) / (2.0 * rho_ns) if g.id_m > 0 else 0.0
    dp_k = g.k * (Gflux * Gflux) / (2.0 * rho_ns)
    dp_z = rho_ns * G_SI * g.dz_m
    d = _pack(rho_ns, mu_ns, v, Re, f, dp_f, dp_k, dp_z)
    d["quality"] = x
    return d


_DISPATCH = {"VAPOR": _dp_vapor, "LIQUID": _dp_liquid, "TWO-PHASE": _dp_twophase}


def _zero_dp(rho):
    return dict(rho=rho, mu=0.0, v=0.0, Re=0.0, f=0.0,
               dp_f=0.0, dp_k=0.0, dp_z=0.0, dp_tot=0.0, quality=None)


def _pack(rho, mu, v, Re, f, dp_f, dp_k, dp_z):
    return dict(rho=rho, mu=mu, v=v, Re=Re, f=f,
               dp_f=dp_f, dp_k=dp_k, dp_z=dp_z,
               dp_tot=dp_f + dp_k + dp_z, quality=None)


@dataclass
class Station:
    seq: int
    comp_id: str | None
    fitting: str | None
    spec: str | None
    nps: str | None
    schedule: str | None
    id_in: float | None
    od_in: float | None
    wall_in: float | None
    length_ft: float | None
    dz_ft: float | None
    cum_ft: float
    p_in_psia: float
    p_out_psia: float
    v_fts: float | None
    Re: float | None
    f: float | None
    rho_lbft3: float | None
    dp_fric_psi: float
    dp_fit_psi: float
    dp_elev_psi: float
    dp_total_psi: float
    note: str = ""


def build_profile(rows, sp, phase, p_start_psia) -> list[Station]:
    """March cumulative pressure along the ordered components for ONE phase."""
    calc = _DISPATCH[phase]
    stations: list[Station] = []
    p = p_start_psia
    cum = 0.0
    last_bore = None
    for row in rows:
        fitting = row.get("Fitting Name")
        spec = row.get("Piping Spec")
        bore = row.get("Bore (in)") or last_bore
        if bore:
            last_bore = bore
        length_ft = num(row.get("Length (ft)")) or 0.0
        dz_ft = num(row.get("Elev Change (ft)")) or 0.0

        idr = resolve_id(spec, bore)
        g = _common_geom(idr, length_ft, dz_ft, fitting)
        if g is None:
            res = _zero_dp(0.0)
            note = idr.basis
        else:
            res = calc(g, sp, p)
            note = idr.basis

        dp_f = res["dp_f"] * PA_TO_PSI
        dp_k = res["dp_k"] * PA_TO_PSI
        dp_z = res["dp_z"] * PA_TO_PSI
        dp_t = dp_f + dp_k + dp_z
        p_out = p - dp_t
        cum += length_ft

        stations.append(Station(
            seq=int(num(row.get("Seq")) or 0),
            comp_id=txt(row.get("Comp ID")),
            fitting=fitting, spec=spec, nps=idr.nps, schedule=idr.schedule,
            id_in=idr.id_in, od_in=idr.od_in, wall_in=idr.wall_in,
            length_ft=round(length_ft, 3), dz_ft=round(dz_ft, 3),
            cum_ft=round(cum, 3), p_in_psia=round(p, 4), p_out_psia=round(p_out, 4),
            v_fts=round(res["v"] / FT_TO_M, 3) if res["v"] else 0.0,
            Re=round(res["Re"], 0) if res["Re"] else 0.0,
            f=round(res["f"], 5) if res["f"] else 0.0,
            rho_lbft3=round(res["rho"] / LBFT3_TO_KGM3, 5) if res["rho"] else 0.0,
            dp_fric_psi=round(dp_f, 5), dp_fit_psi=round(dp_k, 5),
            dp_elev_psi=round(dp_z, 5), dp_total_psi=round(dp_t, 5),
            note=note,
        ))
        p = p_out
    return stations


# ════════════════════════════════════════════════════════════════════════
#  5.  EXCEL OUTPUT
# ════════════════════════════════════════════════════════════════════════
def _hdr(ws, headers, row=1, bg=NAVY):
    for ci, h in enumerate(headers, 1):
        _c(ws, row, ci, h, bg=bg, fg=WHITE, sz=9, bold=True, ha="center", wrap=True)
    ws.row_dimensions[row].height = 32


#  Expanded per-station property columns (user request).  Pressure-dependent
#  values (Total Density, Z, gas volumetric flow) are computed at the local
#  station pressure; the rest are stream constants repeated per row.
EXP_HEADERS = [
    "Temperature (°F)", "Local Pressure (psia)",
    "Total Density (lb/ft³)", "Vapor Density (lb/ft³)", "Liquid Density (lb/ft³)",
    "Vapor Mass Frac", "Liquid Mass Frac", "GVF (local, vol)",
    "Compress. Z", "Cp/Cv (actual)", "γ est (ideal)", "Surf Tens (dyn/cm)",
    "Mass Flow Total (lb/hr)", "Mass Flow Gas (lb/hr)", "Mass Flow Liq (lb/hr)",
    "Vol Flow Total (ft³/hr)", "Vol Flow Gas (ft³/hr)", "Vol Flow Liq (ft³/hr)",
    "Std Vap Rate (ft³/hr)", "Std Liq Rate (ft³/hr)",
    "Total Visc (cP)", "Vapor Visc (cP)", "Liquid Visc (cP)",
    "MW Total", "MW Vapor", "MW Liquid",
    "Therm Cond Total (BTU/hr-ft-°F)", "Therm Cond Vap (BTU/hr-ft-°F)",
    "Therm Cond Liq (BTU/hr-ft-°F)",
    "Enthalpy Total (BTU/lb)", "Enthalpy Vap (BTU/lb)", "Enthalpy Liq (BTU/lb)",
]
EXP_WIDTHS = [max(11, min(22, int(len(h) * 0.95))) for h in EXP_HEADERS]


def _mass_blend(vm, lm, v_val, l_val):
    """Mass-weighted blend of a vapor/liquid property; tolerant of None/0."""
    vm = vm or 0.0
    lm = lm or 0.0
    tot = vm + lm
    if tot <= 0:
        return None
    if v_val is None and l_val is None:
        return None
    if v_val is None:
        return l_val
    if l_val is None:
        return v_val
    return (vm * v_val + lm * l_val) / tot


def _expanded_vals(sp, station):
    """Per-station expanded property row.

    Pressure-dependent quantities (vapor density, total density, gas/total
    volumetric flow, GVF) are recomputed at the LOCAL station pressure so they
    track the pressure drop.  Composition/temperature-driven quantities (MW,
    viscosity, thermal conductivity, enthalpy) are isothermal stream values.
    """
    p = station.p_in_psia
    vm = sp.vap_mass or 0.0
    lm = sp.liq_mass or 0.0
    tot = vm + lm

    rho_gas = (vapor_density_kgm3(sp, p) / LBFT3_TO_KGM3) if sp.mol_weight \
        else sp.vap_density
    rho_liq = sp.liq_density
    rho_tot = station.rho_lbft3                  # marcher's local total density

    q_gas = (vm / rho_gas) if (rho_gas and rho_gas > 0) else None
    q_liq = sp.liq_act_rate
    if q_liq is None and lm and rho_liq:
        q_liq = lm / rho_liq
    q_tot = None
    if q_gas is not None or q_liq is not None:
        q_tot = (q_gas or 0.0) + (q_liq or 0.0)
    gvf = (q_gas / q_tot) if (q_gas is not None and q_tot) else None

    vfrac = (vm / tot) if tot > 0 else None
    lfrac = (lm / tot) if tot > 0 else None
    tc_tot = _mass_blend(vm, lm, sp.vap_therm_cond, sp.liq_therm_cond)
    h_tot = _mass_blend(vm, lm, sp.vap_sp_enthalpy, sp.liq_sp_enthalpy)

    def _r(x, n):
        return round(x, n) if isinstance(x, (int, float)) else None

    return [
        sp.temp_f,
        _r(p, 4),
        rho_tot,
        _r(rho_gas, 6),
        rho_liq,
        _r(vfrac, 4),
        _r(lfrac, 4),
        _r(gvf, 4),
        sp.total_z or sp.vap_z,
        sp.vap_cp_cv,
        sp.gamma_estimate(),
        sp.liq_surf_tens,
        sp.total_mass,
        sp.vap_mass,
        sp.liq_mass,
        _r(q_tot, 2),
        _r(q_gas, 2),
        _r(q_liq, 2),
        sp.total_std_vap,
        sp.total_std_liq,
        sp.total_viscosity_cp(),
        sp.vap_visc,
        sp.liq_visc,
        sp.mol_weight,
        sp.vap_mw,
        sp.liq_mw,
        _r(tc_tot, 6),
        sp.vap_therm_cond,
        sp.liq_therm_cond,
        _r(h_tot, 3),
        sp.vap_sp_enthalpy,
        sp.liq_sp_enthalpy,
    ]


def build_profile_sheet(wb, stations, line_no, stream_name, phase, sp=None):
    ws = wb.active
    ws.title = "Pressure_Profile"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = STEEL

    base_headers = ["Seq", "Comp ID", "Fitting", "Spec", "NPS", "Sched", "ID (in)",
                    "Length (ft)", "Cum Length (ft)", "Elev Δ (ft)",
                    "Velocity (ft/s)", "Reynolds", "Friction f", "Density (lb/ft³)",
                    "dP Fric (psi)", "dP Fitting (psi)", "dP Elev (psi)",
                    "dP Total (psi)", "P In (psia)", "P Out (psia)"]
    # Expanded property columns are inserted after P Out (col 20); Note last.
    headers = base_headers + EXP_HEADERS + ["Basis / Note"]
    n_cols = len(headers)
    note_col = n_cols

    ws.merge_cells(f"A1:{get_column_letter(n_cols)}1")
    _c(ws, 1, 1,
       f"PRESSURE PROFILE   |   Line {line_no}  <->  Stream {stream_name}   |   "
       f"Phase: {phase}   |   {datetime.now():%d-%b-%Y %H:%M}",
       bg=NAVY, fg=WHITE, sz=11, bold=True)
    ws.row_dimensions[1].height = 22

    _hdr(ws, headers, row=2)
    # tint the expanded-property header block
    for ci in range(21, 21 + len(EXP_HEADERS)):
        ws.cell(2, ci).fill = PatternFill("solid", fgColor=GRNHDR)
    ws.freeze_panes = "C3"
    ws.auto_filter.ref = f"A2:{get_column_letter(n_cols)}2"

    for i, s in enumerate(stations):
        r = i + 3
        bg = LGRAY if i % 2 else WHITE
        base_vals = [s.seq, s.comp_id, s.fitting, s.spec, s.nps, s.schedule, s.id_in,
                     s.length_ft, s.cum_ft, s.dz_ft, s.v_fts, s.Re, s.f, s.rho_lbft3,
                     s.dp_fric_psi, s.dp_fit_psi, s.dp_elev_psi, s.dp_total_psi,
                     s.p_in_psia, s.p_out_psia]
        exp_vals = _expanded_vals(sp, s) if sp is not None else [None] * len(EXP_HEADERS)
        vals = base_vals + exp_vals + [s.note]
        for ci, v in enumerate(vals, 1):
            _c(ws, r, ci, "" if v is None else v, bg=bg, sz=9,
               ha="right" if isinstance(v, (int, float)) else "left")
        # Highlight P Out
        _c(ws, r, 20, s.p_out_psia, bg=AMBER, sz=9, bold=True, ha="right")

    widths = ([5, 8, 16, 9, 6, 7, 9, 11, 14, 10, 13, 12, 10, 14,
               12, 13, 12, 13, 12, 12] + EXP_WIDTHS + [34])
    for ci, w in enumerate(widths, 1):
        cw(ws, ci, w)

    # ── Summary block + chart ──────────────────────────────────────────
    last = len(stations) + 2
    sr = last + 2
    total_dp = round(stations[0].p_in_psia - stations[-1].p_out_psia, 4) if stations else 0
    total_len = stations[-1].cum_ft if stations else 0
    _c(ws, sr, 1, "TOTAL", bg=GRNHDR, fg=WHITE, bold=True)
    _c(ws, sr, 2, f"ΔP = {total_dp} psi over {total_len} ft "
                  f"({len(stations)} components)", bg=GREEN, bold=True)

    if len(stations) >= 2:
        chart = LineChart()
        chart.title = f"Pressure Profile — Line {line_no} ({phase})"
        chart.style = 12
        chart.x_axis.title = "Cumulative Length (ft)"
        chart.y_axis.title = "Pressure (psia)"
        chart.height = 9
        chart.width = 22
        data = Reference(ws, min_col=20, min_row=2, max_row=last)  # P Out + header
        cats = Reference(ws, min_col=9, min_row=3, max_row=last)   # cum length
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart, f"A{sr + 2}")


def build_detail_sheet(wb, stations, phase):
    ws = wb.create_sheet("Component_Detail")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = PURPLE
    note = {
        "VAPOR": "Compressible gas: ρ scaled to local P (isothermal, const Z). "
                 "Darcy f·(L/D)·½ρv² + K·½ρv² + ρgΔz.",
        "LIQUID": "Incompressible Darcy-Weisbach: constant ρ. "
                  "f·(L/D)·½ρv² + K·½ρv² + ρgΔz.",
        "TWO-PHASE": "Homogeneous no-slip: ρns=1/(x/ρg+(1-x)/ρl), McAdams μ, "
                     "mass-flux G. f·(L/D)·G²/(2ρns) + K·G²/(2ρns) + ρns·gΔz.",
    }[phase]
    ws.merge_cells("A1:K1")
    _c(ws, 1, 1, f"COMPONENT DETAIL   |   Phase model: {phase}   |   {note}",
       bg=NAVY, fg=WHITE, sz=10, bold=True, wrap=True)
    ws.row_dimensions[1].height = 30

    headers = ["Seq", "Comp ID", "Fitting", _U.hdr("ID", "Lin"),
               _U.hdr("Length", "L"), _U.hdr("Velocity", "v"),
               "Reynolds", "Friction f", _U.hdr("Density", "rho"),
               _U.hdr("dP Total", "dP"), _U.hdr("P Out", "P")]
    _hdr(ws, headers, row=2)
    ws.freeze_panes = "A3"
    for i, s in enumerate(stations):
        r = i + 3
        bg = LGRAY if i % 2 else WHITE
        vals = [s.seq, s.comp_id, s.fitting,
                _U.disp("Lin", s.id_in), _U.disp("L", s.length_ft),
                _U.disp("v", s.v_fts), s.Re, s.f,
                _U.disp("rho", s.rho_lbft3, 5),
                _U.disp("dP", s.dp_total_psi, 5), _U.disp("P", s.p_out_psia)]
        for ci, v in enumerate(vals, 1):
            _c(ws, r, ci, "" if v is None else v, bg=bg, sz=9,
               ha="right" if isinstance(v, (int, float)) else "left")
    for ci, w in enumerate([5, 8, 16, 9, 11, 13, 12, 10, 14, 13, 12], 1):
        cw(ws, ci, w)


def build_stream_sheet(wb, sp, phase):
    ws = wb.create_sheet("Stream_Props")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = GRNHDR
    ws.merge_cells("A1:C1")
    _c(ws, 1, 1, f"HMB STREAM — {sp.stream}", bg=NAVY, fg=WHITE, sz=11, bold=True)
    ws.row_dimensions[1].height = 20

    rows = [
        ("Phase (HMB)", sp.phase, ""),
        ("Phase model chosen", phase, "drives calc dispatch"),
        ("Temperature", _U.disp("T", sp.temp_f, 2), _U.label("T")),
        ("Pressure (inlet)", _U.disp("P", sp.pres_psia), _U.label("P")),
        ("Total mass flow", _U.disp("mflow", sp.total_mass, 1), _U.label("mflow")),
        ("Vapor mass flow", _U.disp("mflow", sp.vap_mass, 1), _U.label("mflow")),
        ("Liquid mass flow", _U.disp("mflow", sp.liq_mass, 1), _U.label("mflow")),
        ("Molecular weight", sp.mol_weight, "-"),
        ("Vapor density", _U.disp("rho", sp.vap_density, 5), _U.label("rho")),
        ("Liquid density", _U.disp("rho", sp.liq_density, 4), _U.label("rho")),
        ("Vapor viscosity", sp.vap_visc, "cP"),
        ("Liquid viscosity", sp.liq_visc, "cP"),
        ("Vapor Z-factor", sp.vap_z, "-"),
        ("Liquid surface tension", sp.liq_surf_tens, "dyn/cm"),
        ("Source sheet", sp.source_sheet, ""),
        ("Unconverged warning", "YES" if sp.warned else "No", ""),
    ]
    _hdr(ws, ["Property", "Value", "Unit / Note"], row=2)
    for i, (k, v, u) in enumerate(rows):
        r = i + 3
        bg = LGRAY if i % 2 else WHITE
        _c(ws, r, 1, k, bg=bg, sz=9, bold=True)
        _c(ws, r, 2, "" if v is None else v, bg=LTBLUE, sz=9,
           ha="right" if isinstance(v, (int, float)) else "left")
        _c(ws, r, 3, u, bg=bg, sz=9)
    cw(ws, 1, 24); cw(ws, 2, 18); cw(ws, 3, 24)






# ════════════════════════════════════════════════════════════════════════
#  6.  PCF EXCEL (via pcf_parser) + interactive main
# ════════════════════════════════════════════════════════════════════════




def load_stream_props(hmb_path, sim_stream, case="Case 1") -> StreamProps | None:
    """Read one stream as StreamProps from either the PRO/II export or a
    legacy per-stream HMB workbook."""
    if is_proii_export(hmb_path):
        if HMBPROII is None:
            return None
        hs = HMBPROII.get_stream(sim_stream, path=hmb_path, case=case)
        return streamprops_from_hmb(hs) if hs else None
    return extract_stream(Path(hmb_path), sim_stream)


# ════════════════════════════════════════════════════════════════════════
#  FIV — Energy Institute T2.2 screening (inlined; SI; D=OD mm, T=wall mm)
# ════════════════════════════════════════════════════════════════════════
P_REF_W = 1e-12          # reference acoustic power (W)


def fcf(gvf: float, mu_gas_pas: float | None) -> float:
    """Fluid Correction Factor (EI T2.2), piecewise in GVF."""
    if gvf < 0.2:
        return 0.2 + 4.0 * gvf
    if gvf < 0.88:
        return 1.0
    if gvf < 0.99:
        return -27.882 * gvf ** 2 + 45.545 * gvf - 17.495
    mu = mu_gas_pas if (mu_gas_pas and mu_gas_pas > 0) else 1.2e-3
    return math.sqrt(mu / 1.2e-3)


def fv_stiff(D: float, T: float) -> float:
    if D < 812:
        a = 446187 + 646 * D + 9.17e-4 * D ** 3
        b = 0.1 * math.log(D) - 1.374
    else:
        a = 419373.9 + 859.4 * D + 0.4803 * D ** 2 - 1.625e-4 * D ** 3
        b = 0.1001 * math.log(D) - 5.309e-5 * D - 1.358
    return a * (D / T) ** b


def fv_medstiff(D: float, T: float) -> float:
    if D < 812:
        a = 283921 + 370 * D
        b = 0.1106 * math.log(D) - 1.501
    else:
        a = 246340 + 502.2 * D + 0.1218 * D ** 2 - 3.32e-5 * D ** 3
        b = 0.0916 * math.log(D) - 1.25e-5 * D - 1.38
    return a * (D / T) ** b


def fv_medium(D: float, T: float) -> float:
    if D < 812:
        if D <= 219:
            return math.exp((13.1 - 4.75e-3 * D + 1.41e-5 * D ** 2)
                            * (D / T) ** (-0.132 + 2.28e-4 * D - 3.72e-7 * D ** 2))
        a = 150412 + 209 * D
        b = 0.0815 * math.log(D) - 1.327
        return a * (D / T) ** b
    a = 179.1 * D + 195015
    b = 0.1533 * math.log(D) - 7e-5 * D - 1.768
    return a * (D / T) ** b


def fv_flexible(D: float, T: float) -> float:
    if D < 812:
        if D <= 219:
            return math.exp((1.32e-5 * D ** 2 - 4.42e-3 * D + 12.22)
                            * (D / T) ** (2.84e-4 * D - 4.62e-7 * D ** 2 - 0.164))
        a = 41.21 * D + 49397
        b = 0.0815 * math.log(D) - 1.384
        return a * (D / T) ** b
    a = 42.21 * D + 52778
    b = 0.0697 * math.log(D) - 1.323
    return a * (D / T) ** b


def span_stiff(D: float, material="Steel") -> float:
    if material.upper() == "GRE":
        return 0.3274 * D ** 0.4332
    if D <= 762:
        return -1.235e-5 * D ** 2 + 0.02 * D + 2.056
    return 0.4056 * D ** 0.4927


def span_medstiff(D: float, material="Steel") -> float:
    if material.upper() == "GRE":
        return 0.4643 * D ** 0.4332
    if D <= 762:
        return -1.189e-5 * D ** 2 + 0.0253 * D + 3.360
    return 0.5623 * D ** 0.5034


def span_medium(D: float, material="Steel") -> float:
    if material.upper() == "GRE":
        return 0.6132 * D ** 0.4332
    if D <= 762:
        return -1.597e-5 * D ** 2 + 0.0336 * D + 4.429
    return 0.7443 * D ** 0.5033


# Most-flexible arrangement first (cheapest support that still meets the limit).
_ARRANGE = [
    ("Flexible", fv_flexible, None),
    ("Medium", fv_medium, span_medium),
    ("Medium Stiff", fv_medstiff, span_medstiff),
    ("Stiff", fv_stiff, span_stiff),
]


def assess_span(rho_v2: float, fcf_val: float, D: float, T: float,
                limit: float, material="Steel") -> dict:
    """Least-stiff support arrangement (and its max span, m) whose Fv keeps
    LOF <= limit.  Flexible => no EI upper span limit."""
    fv_req = rho_v2 * fcf_val / limit if limit else float("inf")
    for name, fv_fn, span_fn in _ARRANGE:
        if fv_fn(D, T) >= fv_req:
            span_m = None if span_fn is None else span_fn(D, material)
            return {"fv_req": fv_req, "arrangement": name, "span_m": span_m}
    return {"fv_req": fv_req, "arrangement": "Not achievable", "span_m": None}


def fiv_station_inputs(station, sp, gas_density_fn):
    """rho*v^2 (SI), GVF and FCF for one station; None if geometry missing."""
    D = (station.od_in or 0.0) * 25.4          # mm
    T = (station.wall_in or 0.0) * 25.4        # mm
    if D <= 0 or T <= 0:
        return None
    rho_si = (station.rho_lbft3 or 0.0) * LBFT3_TO_KGM3
    v_si = (station.v_fts or 0.0) * FT_TO_M
    rho_v2 = rho_si * v_si * v_si
    rho_gas_si = (gas_density_fn(station.p_in_psia) or 0.0) * LBFT3_TO_KGM3
    vm = (sp.vap_mass or 0.0) * LBHR_TO_KGS
    lm = (sp.liq_mass or 0.0) * LBHR_TO_KGS
    q_gas = vm / rho_gas_si if rho_gas_si > 0 else 0.0
    rho_l_si = (sp.liq_density or 0.0) * LBFT3_TO_KGM3
    q_liq = lm / rho_l_si if rho_l_si > 0 else 0.0
    gvf = q_gas / (q_gas + q_liq) if (q_gas + q_liq) > 0 else 1.0
    mu_gas = (sp.vap_visc or 0.0) * CP_TO_PAS
    return {"D": D, "T": T, "rho_v2": rho_v2, "gvf": gvf, "fcf": fcf(gvf, mu_gas)}


def _fiv_support_color(arrangement: str) -> str:
    """Background colour for a support-arrangement cell."""
    if arrangement in ("Stiff", "Not achievable"):
        return RED
    if arrangement == "Medium Stiff":
        return AMBER
    return GREEN


def build_fiv_sheet(wb, stations, sp, line_no, stream_name,
                    gas_density_fn, limits=(0.15, 0.25, 0.35, 0.45), material="Steel",
                    station_lines=None, line_colors=None):
    ws = wb.create_sheet("FIV_EI_T2.2")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = ORANGE

    limits = list(limits)
    base_headers = [
        "Seq", "Comp ID", "Fitting", "NPS", "OD (mm)", "Wall (mm)",
        _U.hdr("ID", "Lin"), _U.hdr("Velocity", "v"), _U.hdr("ρ", "rho"),
        "ρv² (kg/m·s²)", "GVF", "FCF",
        "Fv Stiff", "Fv Med-Stiff", "Fv Medium", "Fv Flexible",
    ]
    n_base = len(base_headers)               # 16
    # Per-LOF block: Support type, Span (m), Span (ft).
    lim_headers = []
    for lim in limits:
        lim_headers += [f"Support (LOF≤{lim})", f"Span @{lim} (m)",
                        f"Span @{lim} (ft)"]
    headers = base_headers + lim_headers
    ncol = len(headers)

    ws.merge_cells(f"A1:{get_column_letter(ncol)}1")
    lim_txt = ", ".join(str(l) for l in limits)
    _c(ws, 1, 1,
       f"FIV — EI T2.2   |   Line {line_no} <-> Stream {stream_name}   |   "
       f"LOF limits = {lim_txt}   |   {material}   |   "
       f"{datetime.now():%d-%b-%Y %H:%M}",
       bg=NAVY, fg=WHITE, sz=11, bold=True)
    ws.row_dimensions[1].height = 22

    _hdr(ws, headers, row=2)
    for ci in range(n_base + 1, ncol + 1):   # tint the LOF blocks
        ws.cell(2, ci).fill = _fill(GRNHDR)
    ws.freeze_panes = "C3"

    def _span_pair(a):
        if a["arrangement"] == "Not achievable":
            return "Not achievable", ""
        if a["span_m"] is None:
            return "No EI span limit", ""
        return round(a["span_m"], 3), round(a["span_m"] / FT_TO_M, 2)

    multi = bool(station_lines and line_colors and len(set(station_lines)) > 1)
    r = 3
    worst = None
    prev_line = None
    for idx, s in enumerate(stations):
        inp = fiv_station_inputs(s, sp, gas_density_fn)
        if inp is None:
            continue
        cur_line = (station_lines[idx]
                    if (station_lines and idx < len(station_lines)) else None)

        # ── Line-transition divider (multi-line circuit) ─────────────────
        if multi and cur_line and cur_line != prev_line:
            ws.merge_cells(f"A{r}:{get_column_letter(ncol)}{r}")
            _c(ws, r, 1, f"  ▶  LINE: {cur_line}",
               bg=line_colors.get(cur_line, LGRAY), fg=NAVY, sz=9, bold=True)
            ws.row_dimensions[r].height = 14
            prev_line = cur_line
            r += 1

        D, T = inp["D"], inp["T"]
        rho_v2, gvf, fcf_v = inp["rho_v2"], inp["gvf"], inp["fcf"]
        assessments = [assess_span(rho_v2, fcf_v, D, T, lim, material)
                       for lim in limits]

        bg = (line_colors.get(cur_line, LGRAY) if (multi and cur_line)
              else (LGRAY if (r % 2 == 0) else WHITE))
        base_vals = [
            s.seq, s.comp_id, s.fitting, s.nps, round(D, 2), round(T, 2),
            _U.disp("Lin", s.id_in), _U.disp("v", s.v_fts),
            _U.disp("rho", s.rho_lbft3, 5),
            round(rho_v2, 1), round(gvf, 4), round(fcf_v, 4),
            round(fv_stiff(D, T), 1), round(fv_medstiff(D, T), 1),
            round(fv_medium(D, T), 1), round(fv_flexible(D, T), 1),
        ]
        for ci, v in enumerate(base_vals, 1):
            _c(ws, r, ci, "" if v is None else v, bg=bg, sz=9,
               ha="right" if isinstance(v, (int, float)) else "left")
        # Per-LOF blocks (support cell colour-coded by stiffness).
        ci = n_base + 1
        for a in assessments:
            span_m, span_ft = _span_pair(a)
            _c(ws, r, ci, a["arrangement"],
               bg=_fiv_support_color(a["arrangement"]), sz=9, bold=True)
            _c(ws, r, ci + 1, span_m, bg=bg, sz=9,
               ha="right" if isinstance(span_m, (int, float)) else "left")
            _c(ws, r, ci + 2, span_ft, bg=bg, sz=9,
               ha="right" if isinstance(span_ft, (int, float)) else "left")
            ci += 3
        if worst is None or rho_v2 > worst[0]:
            worst = (rho_v2, s, gvf, fcf_v, assessments)
        r += 1

    base_widths = [5, 9, 16, 6, 9, 9, 8, 12, 11, 14, 8, 8, 11, 12, 11, 12]
    widths = base_widths + [16, 13, 13] * len(limits)
    for ci, w in enumerate(widths, 1):
        cw(ws, ci, w)

    rr = r + 2
    if worst:
        _, s, gvf, fcf_v, assessments = worst
        _c(ws, rr, 1, "CONTROLLING (max ρv²)", bg=GRNHDR, fg=WHITE, bold=True)
        summ = "  ".join(f"LOF≤{lim}: {a['arrangement']}"
                         for lim, a in zip(limits, assessments))
        _c(ws, rr, 3, f"{s.comp_id} / {s.fitting}  |  ρv²={worst[0]:.0f}  "
                      f"GVF={gvf:.3f}  FCF={fcf_v:.3f}  |  {summ}",
           bg=GREEN, bold=True)
        ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=ncol)
    notes = [
        "",
        "METHOD — Energy Institute T2.2 screening (SI units).",
        "  ρv² = ρ·v²  (effective mixture density × effective velocity²).",
        "  FCF: GVF<0.2 → 0.2+4·GVF;  0.2–0.88 → 1;  0.88–0.99 → "
        "−27.882·GVF²+45.545·GVF−17.495;  ≥0.99 → √(µgas/1.2e-3).",
        "  Fv per support arrangement = α·(D/T)^β  (Medium/Flexible ≤DN200 use "
        "exp(α·(D/T)^β)).  D=OD mm, T=wall mm.",
        "  LOF = ρv²·FCF / Fv.   For each LOF limit the least-stiff support "
        "arrangement whose Fv keeps LOF ≤ limit is reported, with its max span.",
        f"  LOF limits screened: {lim_txt}.",
        "  'No EI span limit' = even a Flexible arrangement satisfies the limit. "
        "'Not achievable' = even Stiff cannot.",
        "VERIFY against the contracted EI guideline edition before design use.",
    ]
    for i, t in enumerate(notes):
        _c(ws, rr + 2 + i, 1, t, sz=9,
           bold=t.startswith("METHOD"), fg=NAVY if t.startswith("METHOD") else DGRAY)


# ════════════════════════════════════════════════════════════════════════
#  AIV — acoustic-induced vibration with source attenuation (inlined)
# ════════════════════════════════════════════════════════════════════════
_ATT_STRAIGHT = 0.2     # dB/m
_ATT_ELBOW = 2.0        # dB each
_ATT_TEE = 5.0
_ATT_REDUCER = 2.0


def hydraulic_power(mdot_kgs: float, dp_pa: float, rho_kgm3: float) -> float:
    """W dissipated across the pressure letdown."""
    return mdot_kgs * dp_pa / rho_kgm3 if rho_kgm3 > 0 else 0.0


def acoustic_power(mdot_kgs, dp_pa, rho_kgm3, efficiency=1e-4) -> float:
    return efficiency * hydraulic_power(mdot_kgs, dp_pa, rho_kgm3)


def sound_power_level(power_w: float) -> float:
    return 10.0 * math.log10(power_w / P_REF_W) if power_w > 0 else 0.0


def _fitting_loss(fitting: str | None) -> float:
    f = (fitting or "").lower()
    if "elbow" in f or "bend" in f:
        return _ATT_ELBOW
    if "tee" in f or "stub" in f or "branch" in f:
        return _ATT_TEE
    if "reducer" in f or "swage" in f:
        return _ATT_REDUCER
    return 0.0


def carucci_mueller_param(lw: float, d_mm: float) -> float:
    return lw + 10.0 * math.log10(d_mm) if d_mm > 0 else lw


def cm_screen(param: float) -> str:
    return "LOW" if param < 145 else "MEDIUM" if param < 155 else "HIGH"


def ei_likelihood(lw: float, d_mm: float, t_mm: float) -> float:
    return lw + 20.0 * math.log10(d_mm / t_mm) if (d_mm > 0 and t_mm > 0) else lw


def ei_risk(factor: float) -> str:
    return "LOW" if factor < 160 else "MEDIUM" if factor < 170 else "HIGH"


def build_aiv_sheet(wb, stations, sp, line_no, stream_name, efficiency=1e-4,
                    station_lines=None, line_colors=None):
    ws = wb.create_sheet("AIV")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = PURPLE

    src_idx = (max(range(len(stations)), key=lambda i: stations[i].dp_total_psi)
               if stations else None)
    mdot = (sp.total_mass or 0.0) * LBHR_TO_KGS
    src = stations[src_idx] if src_idx is not None else None
    if src is not None:
        dp_src_pa = src.dp_total_psi * PSIA_TO_PA
        rho_src = (src.rho_lbft3 or 0.0) * LBFT3_TO_KGM3
        W_ac = acoustic_power(mdot, dp_src_pa, rho_src, efficiency)
        lw_src = sound_power_level(W_ac)
    else:
        dp_src_pa = rho_src = W_ac = lw_src = 0.0

    ncol = 13
    ws.merge_cells(f"A1:{get_column_letter(ncol)}1")
    _c(ws, 1, 1,
       f"AIV — source attenuation   |   Line {line_no} <-> Stream {stream_name}   |   "
       f"η={efficiency:g}   |   {datetime.now():%d-%b-%Y %H:%M}",
       bg=NAVY, fg=WHITE, sz=11, bold=True)
    ws.row_dimensions[1].height = 22

    _c(ws, 2, 1, "SOURCE", bg=GRNHDR, fg=WHITE, bold=True)
    src_txt = (f"{src.comp_id} / {src.fitting}  |  "
               f"Δp={_U.disp('dP', src.dp_total_psi, 4)} {_U.label('dP')}  "
               f"|  W_acoustic={W_ac:.3g} W  |  PWL={lw_src:.1f} dB"
               if src else "no stations")
    ws.merge_cells("B2:M2")
    _c(ws, 2, 2, src_txt, bg=GREEN, bold=True)

    headers = ["Seq", "Comp ID", "Fitting", "OD (mm)", "Wall (mm)",
               "Dist from src (m)", "Cum. atten (dB)", "Local PWL (dB)",
               "C-M param", "C-M screen", "EI likelihood", "EI risk",
               "Dyn. stress idx (MPa)"]
    _hdr(ws, headers, row=3)
    ws.freeze_panes = "C4"

    multi = bool(station_lines and line_colors and len(set(station_lines)) > 1)
    r = 4
    cum_att = 0.0
    started = False
    prev_cum_ft = None
    prev_line = None
    for i, s in enumerate(stations):
        D = (s.od_in or 0.0) * 25.4
        T = (s.wall_in or 0.0) * 25.4
        if src_idx is not None and i == src_idx:
            started = True
            prev_cum_ft = s.cum_ft
            cum_att = 0.0
        if not started:
            continue
        cur_line = (station_lines[i]
                    if (station_lines and i < len(station_lines)) else None)

        # ── Line-transition divider (multi-line circuit) ─────────────────
        if multi and cur_line and cur_line != prev_line:
            ws.merge_cells(f"A{r}:{get_column_letter(ncol)}{r}")
            _c(ws, r, 1, f"  ▶  LINE: {cur_line}",
               bg=line_colors.get(cur_line, LGRAY), fg=NAVY, sz=9, bold=True)
            ws.row_dimensions[r].height = 14
            prev_line = cur_line
            r += 1

        if prev_cum_ft is not None:
            d_ft = max(0.0, (s.cum_ft or 0.0) - prev_cum_ft)
            cum_att += _ATT_STRAIGHT * d_ft * FT_TO_M
        cum_att += _fitting_loss(s.fitting)
        prev_cum_ft = s.cum_ft
        dist_m = max(0.0, (s.cum_ft or 0.0) - (src.cum_ft or 0.0)) * FT_TO_M
        lw_local = lw_src - cum_att
        cmp = carucci_mueller_param(lw_local, D)
        eil = ei_likelihood(lw_local, D, T)
        p_pa = s.p_in_psia * PSIA_TO_PA
        dyn = (p_pa * (D / 1000.0) / (2.0 * (T / 1000.0))) / 1e6 if T > 0 else 0.0
        bg = (line_colors.get(cur_line, LGRAY) if (multi and cur_line)
              else (LGRAY if (r % 2 == 0) else WHITE))
        vals = [s.seq, s.comp_id, s.fitting, round(D, 1), round(T, 2),
                round(dist_m, 2), round(cum_att, 2), round(lw_local, 1),
                round(cmp, 1), cm_screen(cmp), round(eil, 1), ei_risk(eil),
                round(dyn, 1)]
        for ci, v in enumerate(vals, 1):
            _c(ws, r, ci, "" if v is None else v, bg=bg, sz=9,
               ha="right" if isinstance(v, (int, float)) else "left")
        for ci, risk in ((10, cm_screen(cmp)), (12, ei_risk(eil))):
            col = GREEN if risk == "LOW" else AMBER if risk == "MEDIUM" else RED
            _c(ws, r, ci, risk, bg=col, sz=9, bold=True)
        r += 1

    for ci, w in enumerate([5, 9, 16, 9, 9, 15, 14, 13, 11, 11, 13, 10, 18], 1):
        cw(ws, ci, w)

    rr = r + 2
    notes = [
        "METHOD — AIV with source attenuation.",
        "  Source = the component with the largest Δp (pressure letdown).",
        "  Acoustic power W = η · ṁ·Δp/ρ ;  PWL = 10·log10(W/1e-12).",
        "  Attenuation from source: 0.2 dB/m straight pipe + 2 dB/elbow, "
        "5 dB/tee, 2 dB/reducer.  Local PWL = source PWL − cumulative attenuation.",
        "  Carucci-Mueller: param = PWL + 10·log10(D_mm);  <145 LOW, <155 MED, "
        "≥155 HIGH.",
        "  EI likelihood: PWL + 20·log10(D/t);  <160 LOW, <170 MED, ≥170 HIGH.",
        f"  η (acoustic efficiency) = {efficiency:g}; calibrate to the device.",
        "VERIFY source location and η against the design basis before use.",
    ]
    for i, t in enumerate(notes):
        _c(ws, rr + i, 1, t, sz=9, bold=t.startswith("METHOD"),
           fg=NAVY if t.startswith("METHOD") else DGRAY)


# ════════════════════════════════════════════════════════════════════════
#  Two-phase flow-regime mapping (inlined; SI).  Six independent flow-pattern
#  correlations are run per station so transitions (regime disagreement) are
#  visible, plus slug frequency / holdup / bend force.
# ════════════════════════════════════════════════════════════════════════
G = G_SI                   # gravity alias for the two-phase correlations
DYNCM_TO_NM = 1e-3         # dyne/cm -> N/m
RHO_AIR = 1.225            # kg/m³  (Baker reference)
RHO_WATER = 1000.0         # kg/m³
MU_WATER = 1.0e-3          # Pa·s
SIGMA_WATER = 0.072        # N/m

# colour by regime family (for quick visual scan)
_REGIME_BG = {
    "Stratified Smooth": GREEN, "Stratified Wavy": GREEN, "Stratified": GREEN,
    "Wave": GREEN, "Smooth Stratified": GREEN,
    "Bubble": "DDEBF7", "Bubbly": "DDEBF7", "Dispersed Bubble": "DDEBF7",
    "Froth": "DDEBF7", "Plug": AMBER, "Elongated Bubble": AMBER,
    "Intermittent": AMBER, "Slug": AMBER, "Bubbly-Slug": AMBER, "Slug/Churn": AMBER,
    "Churn": ORANGE, "Wispy Annular": "D9D2E9",
    "Annular": RED, "Annular-Mist": RED, "Annular/Mist": RED, "Mist": RED,
    "Spray": RED, "Dispersed": "D9D2E9", "Distributed": "DDEBF7",
    "Transition": "FFF2CC",
    "Single-phase gas": LGRAY, "Single-phase liquid": LGRAY, "n/a": LGRAY,
}


@dataclass(frozen=True)
class TPInputs:
    """SI two-phase inputs for one station."""
    vsg: float          # superficial gas velocity (m/s)
    vsl: float          # superficial liquid velocity (m/s)
    vm: float           # mixture velocity (m/s)
    rho_g: float        # gas density (kg/m³, local)
    rho_l: float        # liquid density (kg/m³)
    mu_g: float         # gas viscosity (Pa·s)
    mu_l: float         # liquid viscosity (Pa·s)
    sigma: float        # surface tension (N/m)
    d: float            # pipe ID (m)
    theta_deg: float    # inclination (+ up, − down)
    gvf: float          # gas volume fraction (input, no-slip)


def station_tp_inputs(station, sp, gas_density_fn) -> "TPInputs | None":
    """Build SI two-phase inputs for a station, or None if geometry missing."""
    d_in = station.id_in
    if not d_in:
        return None
    d = d_in * IN_TO_M
    area = math.pi * 0.25 * d * d

    rho_g = (gas_density_fn(station.p_in_psia) or 0.0) * LBFT3_TO_KGM3
    rho_l = (sp.liq_density or 0.0) * LBFT3_TO_KGM3
    mu_g = (sp.vap_visc or 0.01) * CP_TO_PAS
    mu_l = (sp.liq_visc or 0.5) * CP_TO_PAS
    sigma = (sp.liq_surf_tens or 20.0) * DYNCM_TO_NM

    mg = (sp.vap_mass or 0.0) * LBHR_TO_KGS      # kg/s
    ml = (sp.liq_mass or 0.0) * LBHR_TO_KGS
    qg = mg / rho_g if rho_g > 0 else 0.0
    ql = ml / rho_l if rho_l > 0 else 0.0
    vsg = qg / area
    vsl = ql / area
    vm = vsg + vsl
    gvf = qg / (qg + ql) if (qg + ql) > 0 else 1.0

    length = (station.length_ft or 0.0)
    dz = (station.dz_ft or 0.0)
    theta = (math.degrees(math.asin(max(-1.0, min(1.0, dz / length))))
             if length > 1e-9 else 0.0)
    return TPInputs(vsg, vsl, vm, rho_g, rho_l, mu_g, mu_l, sigma, d, theta, gvf)


def _is_single_phase(t: TPInputs) -> "str | None":
    if t.vsl <= 1e-9 or t.gvf >= 0.99999:
        return "Single-phase gas"
    if t.vsg <= 1e-9 or t.gvf <= 1e-5:
        return "Single-phase liquid"
    return None


def _td_geom(hd: float):
    """Dimensionless stratified geometry vs h̃L = hL/D ∈ (0,1)."""
    g = max(-0.999999, min(0.999999, 2.0 * hd - 1.0))
    ac = math.acos(g)
    root = math.sqrt(1.0 - g * g)
    AL = 0.25 * (math.pi - ac + g * root)
    AG = 0.25 * (ac - g * root)
    SL = math.pi - ac
    SG = ac
    Si = root
    return AL, AG, SL, SG, Si


def _td_equilibrium(x2: float, n: float = 0.2, m: float = 0.2) -> float:
    """Solve Taitel-Dukler combined-momentum for h̃L given X² (horizontal)."""
    A = math.pi / 4.0

    def f(hd):
        AL, AG, SL, SG, Si = _td_geom(hd)
        if AL < 1e-9 or AG < 1e-9:
            return None
        UL = A / AL
        UG = A / AG
        DL = 4.0 * AL / SL
        DG = 4.0 * AG / (SG + Si)
        liq = x2 * (UL * DL) ** (-n) * UL ** 2 * (SL / AL)
        gas = (UG * DG) ** (-m) * UG ** 2 * (SG / AG + Si / AL + Si / AG)
        return liq - gas

    lo, hi = 1e-4, 1.0 - 1e-4
    flo, fhi = f(lo), f(hi)
    if flo is None or fhi is None or flo * fhi > 0:
        best, bestabs = 0.5, 1e30
        for i in range(1, 1000):
            hd = i / 1000.0
            v = f(hd)
            if v is not None and abs(v) < bestabs:
                bestabs, best = abs(v), hd
        return best
    for _ in range(80):
        mid = 0.5 * (lo + hi)
        fm = f(mid)
        if fm is None:
            lo = mid
            continue
        if flo * fm <= 0:
            hi, fhi = mid, fm
        else:
            lo, flo = mid, fm
    return 0.5 * (lo + hi)


def taitel_dukler(t: TPInputs) -> dict:
    """{'regime','hL_D','F','K','T','X'} — horizontal/near-horizontal TD."""
    cos = max(0.05, math.cos(math.radians(t.theta_deg)))

    def dpdl(v, rho, mu):
        if v <= 0:
            return 1e-12
        Re = rho * abs(v) * t.d / mu
        fF = 0.046 * Re ** -0.2 if Re > 0 else 0.046
        return 4.0 * fF * rho * v * v / (2.0 * t.d)

    dpl = dpdl(t.vsl, t.rho_l, t.mu_l)
    dpg = dpdl(t.vsg, t.rho_g, t.mu_g)
    X = math.sqrt(dpl / dpg) if dpg > 0 else 1e6
    F = math.sqrt(t.rho_g / max(1e-9, t.rho_l - t.rho_g)) * t.vsg / math.sqrt(t.d * G * cos)
    nu_l = t.mu_l / t.rho_l
    Re_sl = t.vsl * t.d / nu_l if nu_l > 0 else 0.0
    K = F * math.sqrt(max(0.0, Re_sl))
    Tg = math.sqrt(abs(dpl) / max(1e-9, (t.rho_l - t.rho_g) * G * cos))

    hd = _td_equilibrium(X * X)
    AL, AG, SL, SG, Si = _td_geom(hd)
    A = math.pi / 4.0
    UG = A / AG
    UL = A / AL

    dAL_dh = Si  # d(ÃL)/d(h̃) ≈ S̃i
    F2_crit = (1.0 - hd) ** 2 * AG / (UG ** 2 * (dAL_dh / A))
    stratified = F * F <= F2_crit

    if stratified:
        s = 0.01
        K_crit = (math.sqrt(4.0 * nu_l * (t.rho_l - t.rho_g) * G * cos
                            / (s * t.rho_g * max(1e-9, t.vsg) ** 2))
                  if t.vsg > 0 else 1e9)
        regime = "Stratified Wavy" if K >= K_crit else "Stratified Smooth"
    else:
        if hd >= 0.35:
            T_crit = math.sqrt(8.0 * AG / (Si * UL ** 2 * (UL * 4 * AL / SL) ** -0.2))
            regime = "Dispersed Bubble" if Tg >= T_crit else "Intermittent"
        else:
            regime = "Annular"
    return {"regime": regime, "hL_D": hd, "F": F, "K": K, "T": Tg, "X": X}


def taitel_barnea_vertical(t: TPInputs) -> str:
    """Vertical / steep up-flow regime (Taitel-Barnea 1980)."""
    dr = max(1e-9, t.rho_l - t.rho_g)
    u0 = 1.53 * (G * t.sigma * dr / t.rho_l ** 2) ** 0.25
    vsg_ann = 3.1 * (G * t.sigma * dr / t.rho_g ** 2) ** 0.25
    if t.vsg >= vsg_ann:
        return "Annular"
    vm_db = (4.0 * (t.d * G * dr / t.rho_l) ** 0.5
             * (t.sigma / (dr * G * t.d ** 2)) ** 0.1) if t.d > 0 else 1e9
    if t.vm >= vm_db and t.gvf < 0.52:
        return "Dispersed Bubble"
    vsl_bs = 3.0 * t.vsg - 1.15 * u0
    if t.vsl > vsl_bs and t.gvf < 0.25:
        return "Bubble"
    if t.vsg > 0.5 and t.gvf > 0.35:
        return "Churn"
    return "Slug"


def baker(t: TPInputs) -> dict:
    """Baker (1954) horizontal mass-flux map."""
    Gg = t.rho_g * t.vsg
    Gl = t.rho_l * t.vsl
    lam = math.sqrt((t.rho_g / RHO_AIR) * (t.rho_l / RHO_WATER))
    psi = (SIGMA_WATER / t.sigma) * ((t.mu_l / MU_WATER) * (RHO_WATER / t.rho_l) ** 2) ** (1.0 / 3.0)
    By = Gg / lam if lam > 0 else 0.0
    Bx = Gl / Gg * lam * psi if Gg > 0 else 1e9
    if By > 80.0:
        regime = "Dispersed" if Bx < 1.0 else "Annular"
    elif By > 3.0:
        regime = "Annular" if Bx < 0.3 else "Slug" if Bx < 6.0 else "Bubble"
    else:
        if Bx < 0.5:
            regime = "Stratified"
        elif Bx < 5.0:
            regime = "Wave"
        elif Bx < 50.0:
            regime = "Plug"
        else:
            regime = "Bubble"
    return {"regime": regime, "Bx": Bx, "By": By, "lambda": lam, "psi": psi}


def hewitt_roberts(t: TPInputs) -> dict:
    """Hewitt-Roberts (1969) vertical momentum-flux map."""
    jl = t.rho_l * t.vsl ** 2
    jg = t.rho_g * t.vsg ** 2
    if jg > 1e3 and jl > 1e4:
        regime = "Wispy Annular"
    elif jg > 1e3:
        regime = "Annular"
    elif jg > 1e1:
        regime = "Churn" if jl < 1e3 else "Bubbly-Slug"
    elif jl > 1e3:
        regime = "Bubbly"
    else:
        regime = "Slug"
    return {"regime": regime, "rho_l_vsl2": jl, "rho_g_vsg2": jg}


def barnea_unified(t: TPInputs) -> str:
    """Barnea (1987) inclination-aware unified classification."""
    ang = abs(t.theta_deg)
    dr = max(1e-9, t.rho_l - t.rho_g)
    fF = 0.046 * max(1.0, t.rho_l * t.vm * t.d / t.mu_l) ** -0.2
    vm_db = (2.0 * (0.4 * t.sigma / (dr * G)) ** 0.5
             * (t.rho_l / t.sigma) ** 0.6 * (2.0 * fF / t.d) ** 0.4) if t.d > 0 else 1e9
    if t.vm >= vm_db and t.gvf < 0.52:
        return "Dispersed Bubble"
    if ang >= 60.0:
        return taitel_barnea_vertical(t)
    return taitel_dukler(t)["regime"]


def shell_beggs_brill(t: TPInputs) -> str:
    """Shell DEP 31.22.05.11 / Beggs-Brill horizontal pattern map."""
    lam = max(1e-6, min(1.0, t.vsl / t.vm if t.vm > 0 else 0.0))
    Fr = t.vm ** 2 / (G * t.d) if t.d > 0 else 0.0
    L1 = 316.0 * lam ** 0.302
    L2 = 0.0009252 * lam ** -2.4684
    L3 = 0.10 * lam ** -1.4516
    L4 = 0.5 * lam ** -6.738
    if (lam < 0.01 and Fr < L1) or (lam >= 0.01 and Fr < L2):
        return "Stratified"
    if lam >= 0.01 and L2 <= Fr <= L3:
        return "Transition"
    if (0.01 <= lam < 0.4 and L3 < Fr <= L1) or (lam >= 0.4 and L3 < Fr <= L4):
        return "Intermittent"
    return "Distributed"


def slug_holdup_gregory(vm: float) -> float:
    """Gregory et al. (1978) slug-body liquid holdup."""
    return 1.0 / (1.0 + (max(0.0, vm) / 8.66) ** 1.39)


def slug_freq_gregory_scott(vsl: float, vm: float, d: float) -> float:
    """Gregory-Scott (1969) slug frequency (Hz)."""
    if d <= 0 or vm <= 0:
        return 0.0
    inner = vsl / (G * d) * (19.75 / vm + vm)
    return 0.0226 * inner ** 1.2 if inner > 0 else 0.0


def slug_force_bend(rho_s: float, vm: float, d: float) -> float:
    """Resultant momentum reaction of a slug at a 90° bend (N)."""
    area = math.pi * 0.25 * d * d
    return math.sqrt(2.0) * rho_s * vm * vm * area


_TPR_HEADERS = [
    "Seq", "Comp ID", "Fitting", "ID (in)", "Incl (°)",
    "Vsg (m/s)", "Vsl (m/s)", "Vm (m/s)", "GVF",
    "Taitel-Dukler (horiz)", "Taitel-Barnea (vert)", "Baker (horiz)",
    "Hewitt-Roberts (vert)", "Barnea unified", "Shell DEP / Beggs-Brill",
    "Slug freq (/min)", "Slug holdup", "Slug force @bend (N)",
]
_TPR_WIDTHS = [5, 9, 12, 8, 8, 10, 10, 10, 8, 19, 18, 14, 16, 15, 18, 13, 11, 16]


def _regime_cell(ws, r, ci, regime):
    _c(ws, r, ci, regime, bg=_REGIME_BG.get(regime, WHITE), ha="center")


def build_regime_sheet(wb, stations, sp, line_no, stream_name, gas_density_fn):
    """Two-phase flow-regime map across all stations + slug summary."""
    ws = wb.create_sheet("Two_Phase_Regime")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = PURPLE

    title = (f"Two-Phase Flow Regime Map   |   Line {line_no} <-> Stream "
             f"{stream_name}   |   {datetime.now():%d-%b-%Y %H:%M}")
    _c(ws, 1, 1, title, fg=NAVY, sz=11, bold=True)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(_TPR_HEADERS))

    _hdr(ws, _TPR_HEADERS, row=2)
    _c(ws, 2, 4, _U.hdr("ID", "Lin"), bg=NAVY, fg=WHITE, sz=9, bold=True,
       ha="center", wrap=True)
    for ci, w in enumerate(_TPR_WIDTHS, 1):
        cw(ws, ci, w)

    r = 3
    regime_counts: dict[str, int] = {}
    max_force = (0.0, None)
    single_phase_label = None
    for s in stations:
        t = station_tp_inputs(s, sp, gas_density_fn)
        _c(ws, r, 1, s.seq, ha="center")
        _c(ws, r, 2, s.comp_id or "", ha="center")
        _c(ws, r, 3, s.fitting or "", ha="center")
        _c(ws, r, 4, round(_U.disp("Lin", s.id_in or 0.0, None), 3), ha="center")
        if t is None:
            _c(ws, r, 5, "geom?", ha="center", bg=LGRAY)
            r += 1
            continue
        sp_label = _is_single_phase(t)
        _c(ws, r, 5, round(t.theta_deg, 1), ha="center")
        _c(ws, r, 6, round(t.vsg, 3), ha="center")
        _c(ws, r, 7, round(t.vsl, 4), ha="center")
        _c(ws, r, 8, round(t.vm, 3), ha="center")
        _c(ws, r, 9, round(t.gvf, 4), ha="center")

        if sp_label:
            single_phase_label = sp_label
            for ci in range(10, 16):
                _c(ws, r, ci, sp_label, bg=LGRAY, ha="center")
            for ci in (16, 17, 18):
                _c(ws, r, ci, "—", ha="center")
            r += 1
            continue

        td = taitel_dukler(t)
        tbv = taitel_barnea_vertical(t)
        bk = baker(t)
        hr = hewitt_roberts(t)
        bu = barnea_unified(t)
        sb = shell_beggs_brill(t)
        _regime_cell(ws, r, 10, td["regime"])
        _regime_cell(ws, r, 11, tbv)
        _regime_cell(ws, r, 12, bk["regime"])
        _regime_cell(ws, r, 13, hr["regime"])
        _regime_cell(ws, r, 14, bu)
        _regime_cell(ws, r, 15, sb)
        for reg in (td["regime"], tbv, bk["regime"], hr["regime"], bu, sb):
            regime_counts[reg] = regime_counts.get(reg, 0) + 1

        freq_hz = slug_freq_gregory_scott(t.vsl, t.vm, t.d)
        hls = slug_holdup_gregory(t.vm)
        rho_s = t.rho_l * hls + t.rho_g * (1.0 - hls)
        force = slug_force_bend(rho_s, t.vm, t.d)
        if force > max_force[0]:
            max_force = (force, s)
        _c(ws, r, 16, round(freq_hz * 60.0, 2), ha="center")
        _c(ws, r, 17, round(hls, 3), ha="center")
        _c(ws, r, 18, round(force, 1), ha="center")
        r += 1

    # ── summary ──
    r += 1
    for ci in range(1, len(_TPR_HEADERS) + 1):
        _c(ws, r, ci, "SUMMARY" if ci == 1 else "", bg=NAVY,
           fg=WHITE, bold=True)
    r += 1
    if single_phase_label and not regime_counts:
        _c(ws, r, 1, f"All stations {single_phase_label.lower()} — no two-phase "
                     f"regime applies (GVF≈1).", fg=DGRAY, sz=9)
        r += 1
    else:
        if max_force[1] is not None:
            s = max_force[1]
            _c(ws, r, 1, "Max slug force @bend", bold=True)
            _c(ws, r, 3, f"seq {s.seq} / {s.fitting}  |  {max_force[0]:.0f} N "
                         f"({max_force[0] / 4.448:.0f} lbf)", fg=PURPLE)
            r += 1
        if regime_counts:
            top = sorted(regime_counts.items(), key=lambda kv: -kv[1])
            _c(ws, r, 1, "Regime votes (all methods)", bold=True)
            _c(ws, r, 3, ", ".join(f"{k}×{v}" for k, v in top), fg=DGRAY)
            r += 1

    # ── method legend ──
    r += 1
    notes = [
        "METHODS — six independent two-phase flow-pattern correlations are run "
        "per station so transitions (regime disagreement) are visible.",
        "  Taitel-Dukler (1976): mechanistic, horizontal/near-horizontal. "
        "Solves equilibrium liquid level h̃L from Martinelli X, then transitions "
        "A (stratified→intermittent/annular, via F), B (smooth→wavy, via K), "
        "C (→dispersed bubble, via T), D (intermittent↔annular at hL/D≈0.35).",
        "  Taitel-Barnea (1980): mechanistic vertical/steep up-flow — bubble, "
        "slug, churn, annular, dispersed-bubble (bubble-rise & annular vsg).",
        "  Baker (1954): empirical horizontal mass-flux map (Bx, By with λ, ψ "
        "property groups). Boundaries are an approximate chart digitisation.",
        "  Hewitt-Roberts (1969): empirical vertical momentum-flux map "
        "(ρL·vsL² vs ρG·vsG²).",
        "  Barnea (1987) unified: inclination-aware; adds the dispersed-bubble & "
        "annular criteria and selects horizontal-TD or vertical-Barnea by angle.",
        "  Shell DEP / Beggs-Brill: the λL–Froude L1..L4 pattern map referenced "
        "by Shell DEP 31.22.05.11 for horizontal lines.",
        "SLUG — frequency: Gregory-Scott (1969) → slugs/min;  holdup: Gregory "
        "(1978);  force: √2·ρ_slug·vm²·A = resultant momentum reaction at a 90° "
        "bend (ρ_slug = ρL·HLS + ρG·(1−HLS)).",
        "VERIFY mechanistic transitions against a dedicated multiphase tool "
        "(e.g. OLGA / LedaFlow) before design use.",
    ]
    for n in notes:
        _c(ws, r, 1, n, fg=DGRAY, sz=8, wrap=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(_TPR_HEADERS))
        r += 1

    ws.freeze_panes = "D3"
    return ws


def run(pcf_path, hmb_path, line_no=None, stream_name=None, out_path=None,
        interactive=True, case="Case 1", map_path=None):
    pipeline, components = PCF.parse_pcf(pcf_path)
    rows = PCF.build_rows(pipeline, components)
    parsed_line = pipeline["line_no"]
    if line_no is None:
        line_no = parsed_line
    material_class = pipeline["attrs"].get("PIPING-SPEC")

    # ── Auto-resolution via stream_map.xlsx (map-only; error if missing) ──
    map_row = None
    if stream_name is None:
        map_row = SMAP.require(line_no, material_class, map_path)
        stream_name = map_row.lookup_key
        print(f"  stream_map: {line_no} / {material_class}  ->  "
              f"'{map_row.stream_number}'  (sim '{map_row.sim_stream_number}')")

    sp = load_stream_props(hmb_path, stream_name, case)
    if sp is None:
        raise SystemExit(
            f"Stream '{stream_name}' not found in {hmb_path} "
            f"(case '{case}'). Check the Simulation Stream Number in stream_map.")

    phase = classify_phase(sp)
    p_start = sp.pres_psia or num(pipeline["attrs"].get("ATTRIBUTE1")) or 100.0
    print(f"  Phase: {sp.phase} -> '{phase}'  |  P_in {p_start} psia  |  "
          f"vap ρ {sp.vap_density} liq ρ {sp.liq_density} lb/ft3")

    stations = build_profile(rows, sp, phase, p_start)

    wb = Workbook()
    build_profile_sheet(wb, stations, line_no, stream_name, phase, sp)
    build_detail_sheet(wb, stations, phase)
    build_stream_sheet(wb, sp, phase)
    build_fiv_sheet(
        wb, stations, sp, line_no, stream_name,
        gas_density_fn=lambda p: vapor_density_kgm3(sp, p) / LBFT3_TO_KGM3
        if sp.mol_weight else (sp.vap_density or 0.0))
    build_aiv_sheet(wb, stations, sp, line_no, stream_name)
    build_regime_sheet(
        wb, stations, sp, line_no, stream_name,
        gas_density_fn=lambda p: vapor_density_kgm3(sp, p) / LBFT3_TO_KGM3
        if sp.mol_weight else (sp.vap_density or 0.0))
    fp_sheets = []
    try:
        import flow_pattern_maps as FPM
        fp_sheets = FPM.build_flow_pattern_maps(
            wb, stations, sp, line_no, stream_name,
            gas_density_fn=lambda p: vapor_density_kgm3(sp, p) / LBFT3_TO_KGM3
            if sp.mol_weight else (sp.vap_density or 0.0))
    except Exception as exc:                      # never block the core output
        print(f"  (flow-pattern maps skipped: {exc})")
    build_pcf_sheet(wb, pipeline, rows)
    build_readme(wb, line_no, stream_name, sp, phase, pcf_path, hmb_path)
    order = (["Pressure_Profile", "Component_Detail", "Stream_Props",
              "FIV_EI_T2.2", "AIV", "Two_Phase_Regime"] + fp_sheets
             + ["PCF_Components", "README"])
    for name in reversed(order):
        if name in [s.title for s in wb.worksheets]:
            wb.move_sheet(name, offset=-len(wb.worksheets))

    if out_path is None:
        safe_l = re.sub(r"[^A-Za-z0-9_-]", "_", line_no)
        safe_s = re.sub(r"[^A-Za-z0-9_-]", "_", stream_name)
        out_path = f"hydraulics_{safe_l}_{safe_s}.xlsx"
    wb.save(out_path)
    return out_path, stations, sp, phase








# ══════════════════════════════════════════════════════════════════════════
# ║  SECTION: No-ISO flash orchestration (former hydraulics_engine_flash_noiso.py)
# ══════════════════════════════════════════════════════════════════════════


# ═══ aliases: engine + flash modules are merged into this file ═══
HE = _sys.modules[__name__]
FV = _sys.modules[__name__]   # flash VLE merged into this file

#!/usr/bin/env python3
"""Flash-coupled hydraulics engine — NO-ISO mode (manual pipeline input).

Instead of parsing an ISOGEN .pcf the user builds the pipeline row-by-row in a
structured Excel workbook (``pipeline_input_noiso.xlsx``).  Each row specifies:

  Fitting name  — drop-down from the project Fittings.md catalogue
  Bore (in), Piping Spec, Length (ft), Elevation Change (ft)
  Fixed K       — K coefficient override for the ``Fix K`` fitting
  Fixed dP      — direct pressure drop for the ``Fix Pressure`` fitting
  Instr Type / Instr Tag / Instr dP — Flow Orifices (FO) with manual dP

Multiple Line Numbers live in the same workbook (one row per component).
Each line's stream is resolved through ``stream_map.xlsx`` exactly as the ISO
flash engine.  Main and Branch run types are both supported; each Branch block
starts at its own user-supplied ``Start P (psia)``.

The pressure marcher and VLE flash logic are identical to
``hydraulics_engine_flash.py`` (isothermal by default; add ``--isenthalpic``
for adiabatic / JT mode).

Usage:
    python hydraulics_engine_flash_noiso.py --template
    python hydraulics_engine_flash_noiso.py pipeline_input_noiso.xlsx HMB.xlsx
    python hydraulics_engine_flash_noiso.py pipeline_input_noiso.xlsx HMB.xlsx --isenthalpic
"""

import math
import os
import re
import sys
from dataclasses import replace
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation




def _u() -> "UN.UnitSystem":
    """Live display-units context (set per run from the input workbook)."""
    return _U

# ── Re-use all styling / conversion constants from the base engine ─────────
_c, _hdr, cw, num, txt = HE._c, HE._hdr, HE.cw, HE.num, HE.txt


# ════════════════════════════════════════════════════════════════════════
#  Fitting catalogue  (Fittings.md)
# ════════════════════════════════════════════════════════════════════════
FITTING_NAMES: list[str] = [
    # Pipeline Entrances
    "Pipeline Entrance - Bell Mouth",
    "Pipeline Entrance - Inward Projecting",
    "Pipeline Entrance - Sharp Edge",
    "Pipeline Entrance - Slightly Rounded",
    "Pipeline Entrance - Well Rounded",
    # Pipeline Exits
    "Pipeline Exit - Outward Projecting",
    "Pipeline Exit - Rounded",
    "Pipeline Exit - Sharp Edge",
    # Elbows
    "Elbow 45 Long",
    "Elbow 45 Short",
    "Elbow 90 Long",
    "Elbow 90 Short",
    "Elbow Miter",
    "Elbow Long Return",
    "Elbow Short Return",
    # Gate valves
    "Gate Half Open",
    "Gate One Quarter Open",
    "Gate Three Quarter Open",
    "Gate Full Open",
    # Globe valves
    "Globe Half Open",
    "Globe One Quarter Open",
    "Globe Three Quarter Open",
    "Globe Full Open",
    # Ball valves
    "Ball Half Open",
    "Ball One Quarter Open",
    "Ball Three Quarter Open",
    "Ball Full Open",
    # Butterfly valves
    "Butterfly Half Open",
    "Butterfly One Quarter Open",
    "Butterfly Three Quarter Open",
    "Butterfly Full Open",
    # Angle valves
    "Angle No Obstruction",
    "Angle Wing Pin",
    # Check valves
    "Check Angle Lift Stop",
    "Check Clear Way Swing",
    "Check Conventional Swing",
    "Check Global Lift Stop",
    "Check Inline Ball",
    "Check T or Y",
    # Size changes
    "Expander",
    "Reducer",
    "Sudden Contraction",
    "Sudden Enlargement",
    # Tees
    "Tee Join Branch Flow 1",
    "Tee Join Branch Flow 2",
    "Tee Join Line Flow",
    "Tee Split Branch Flow 1",
    "Tee Split Branch Flow 2",
    "Tee Split Line Flow",
    # Special / instruments
    "Fix Pressure",
    "Fix K",
    "Nozzle",
    "Orifice",
    "Straight Pipeline",
    "Venturi",
    # Boundaries / control devices
    "Source",
    "Control Valve",
    "Destination",
]

# K values — Crane TP-410 / Darby representative values
_K_TABLE: dict[str, float] = {
    "pipeline entrance - bell mouth":         0.04,
    "pipeline entrance - inward projecting":  0.78,
    "pipeline entrance - sharp edge":         0.50,
    "pipeline entrance - slightly rounded":   0.23,
    "pipeline entrance - well rounded":       0.04,
    "pipeline exit - outward projecting":     1.00,
    "pipeline exit - rounded":                1.00,
    "pipeline exit - sharp edge":             1.00,
    "elbow 45 long":                          0.20,
    "elbow 45 short":                         0.40,
    "elbow 90 long":                          0.30,
    "elbow 90 short":                         0.75,
    "elbow miter":                            1.10,
    "elbow long return":                      0.30,
    "elbow short return":                     0.75,
    "gate half open":                         5.60,
    "gate one quarter open":                 24.00,
    "gate three quarter open":                1.00,
    "gate full open":                         0.15,
    "globe half open":                        9.50,
    "globe one quarter open":                40.00,
    "globe three quarter open":               3.70,
    "globe full open":                        6.40,
    "ball half open":                         5.60,
    "ball one quarter open":                 24.00,
    "ball three quarter open":                1.00,
    "ball full open":                         0.05,
    "butterfly half open":                   10.00,
    "butterfly one quarter open":            45.00,
    "butterfly three quarter open":           2.00,
    "butterfly full open":                    0.60,
    "angle no obstruction":                   2.00,
    "angle wing pin":                         5.80,
    "check angle lift stop":                  8.50,
    "check clear way swing":                  0.75,
    "check conventional swing":               2.30,
    "check global lift stop":                 6.00,
    "check inline ball":                      3.00,
    "check t or y":                           1.80,
    "expander":                               0.30,
    "reducer":                                0.10,
    "sudden contraction":                     0.50,
    "sudden enlargement":                     1.00,
    "tee join branch flow 1":                 1.80,
    "tee join branch flow 2":                 1.80,
    "tee join line flow":                     0.40,
    "tee split branch flow 1":                1.80,
    "tee split branch flow 2":                1.80,
    "tee split line flow":                    0.40,
    "nozzle":                                 0.04,
    "orifice":                                0.00,   # dP from Instr dP column
    "straight pipeline":                      0.00,
    "venturi":                                0.10,
    "fix pressure":                           0.00,   # dP from Fixed dP column
    "fix k":                                  0.00,   # K from Fixed K column
    "source":                                 0.00,   # boundary — anchors Set P
    "control valve":                          0.00,   # dP from control type / Fix K / Fix dP
    "destination":                            0.00,   # boundary — target Set P
}


# ── PCF palette (matches PCF_Parsed_updated.xlsx / Legend sheet) ──────────
_CLR_PIPE      = "F2F2F2"  # grey  — straight pipe
_CLR_ELBOW     = "DEEAF1"  # blue  — elbow / return / bend
_CLR_VALVE     = "FFF2CC"  # amber — any valve
_CLR_FLANGE    = "E2EFDA"  # green — flange / Main run-type label
_CLR_REDUCER   = "FCE4D6"  # red   — reducer / expander / contraction
_CLR_TEE       = "EAD1DC"  # pink  — tee / stub
_CLR_INSTR     = "E8D5F5"  # purple — instrument / orifice / venturi / nozzle
_CLR_CV        = "C9B8E8"  # dark-purple — Fix K (control-valve resistance)
_CLR_PSV       = "F4CCCC"  # salmon — Fix Pressure (relief / safety device)
_CLR_ENTRANCE  = "D9E2F3"  # steel-blue — entrance / exit
_CLR_UP        = "C6EFCE"  # light-green — UP direction
_CLR_DOWN      = "FFC7CE"  # light-red   — DOWN direction
_CLR_BRANCH    = "DEEAF1"  # same as elbow — Branch run-type label


_CLR_SOURCE    = "C6EFCE"  # light-green — Source boundary (circuit inlet)
_CLR_DEST      = "FFC7CE"  # light-red   — Destination boundary (circuit outlet)


def fitting_bg(name: str | None) -> str:
    """Return the PCF-palette hex fill for a fitting name."""
    f = (name or "").lower()
    # Control / boundary devices first (before the generic 'valve' match,
    # since 'control valve' also contains 'valve').
    if "control valve" in f:                     return _CLR_CV
    if "source" in f:                            return _CLR_SOURCE
    if "destination" in f:                       return _CLR_DEST
    if "entrance" in f or "exit" in f:           return _CLR_ENTRANCE
    if "elbow" in f or "return" in f:            return _CLR_ELBOW
    if ("gate" in f or "globe" in f or "ball" in f
            or "butterfly" in f or "angle" in f
            or "check" in f or "valve" in f):    return _CLR_VALVE
    if "flange" in f:                            return _CLR_FLANGE
    if ("reducer" in f or "expander" in f
            or "contraction" in f or "enlargement" in f): return _CLR_REDUCER
    if "tee" in f or "stub" in f:                return _CLR_TEE
    if "orifice" in f or "venturi" in f or "nozzle" in f: return _CLR_INSTR
    if "fix k" in f:                             return _CLR_CV
    if "fix pressure" in f:                      return _CLR_PSV
    if "pipe" in f or "straight" in f or "weld" in f: return _CLR_PIPE
    return WHITE


def direction_bg(direction: str | None) -> str | None:
    """Return highlight colour for UP/DOWN direction cells, else None."""
    d = (direction or "").upper()
    if d == "UP":   return _CLR_UP
    if d == "DOWN": return _CLR_DOWN
    return None


def fitting_k(name: str | None) -> float:
    return _K_TABLE.get((name or "").lower().strip(), 0.0)


def _is_fix_pressure(fitting: str | None) -> bool:
    return (fitting or "").lower().strip() == "fix pressure"


def _is_fix_k(fitting: str | None) -> bool:
    return (fitting or "").lower().strip() == "fix k"


def _is_orifice(fitting: str | None, instr_type: str | None) -> bool:
    f  = (fitting or "").lower().strip()
    it = (instr_type or "").upper().strip()
    return f == "orifice" or it in ("FO", "ORIFICE", "FLOW ORIFICE")


def _is_source(fitting: str | None) -> bool:
    return (fitting or "").lower().strip() == "source"


def _is_destination(fitting: str | None) -> bool:
    return (fitting or "").lower().strip() == "destination"


def _is_control_valve(fitting: str | None) -> bool:
    return (fitting or "").lower().strip() == "control valve"


def _cv_control_type(row: dict) -> str:
    """Normalise the per-row 'Control Valve Type' to one of P/T/F/L.

    P = pressure control, T = temperature, F = flow, L = level.  Blank or
    unrecognised defaults to F (flow) — the only mode whose ΔP floats with the
    fixed downstream Destination pressure; P/T/L use a fixed design ΔP."""
    v = (txt(row.get("Control Valve Type")) or "").strip().upper()
    return v[0] if v and v[0] in ("P", "T", "F", "L") else "F"


def _downstream_dest_p(block_rows: list[dict], start_idx: int) -> float | None:
    """First downstream ``Destination`` row's ``Set P`` (internal psia), or None.

    Used by a flow-control ``Control Valve`` to size its ΔP so the running
    pressure lands on the fixed Destination pressure (typical topology places
    the valve immediately upstream of the Destination)."""
    for r in block_rows[start_idx + 1:]:
        if _is_destination(r.get("Fitting Name")):
            sp = num(r.get("Set P (psia)"))
            if sp is not None:
                return sp
    return None


# ════════════════════════════════════════════════════════════════════════
#  Input-template column definitions
# ════════════════════════════════════════════════════════════════════════
INPUT_HEADERS: list[str] = [
    "Circuit",                # hydraulic circuit ID (e.g. C1).  All rows with the same
    #                           Circuit ID are chained in series into ONE workbook.
    "Line No",                # line identifier (e.g. ER-162)
    "HMB File",               # per-row HMB workbook override; blank → CLI default
    "Case",                   # per-row HMB case override; blank → CLI default
    "Stream Lookup",          # HMB stream key; blank → use manual property cells below
    "Run Type",               # Main | Branch
    "Seq",                    # component sequence (auto-incremented if blank)
    "Start P (psia)",         # starting pressure; first row of the circuit / branch
    "Flow Fraction from Main",  # signed fraction of Main flow at a Tee (+merge / -split)
    "Mass Vapor Fraction Carry Over",   # MASS fraction of vapour carried (blank=1; 0=dump)
    "Mass Liquid Fraction Carry Over",  # MASS fraction of liquid carried (blank=1; 0=dump)
    # ── Manual stream properties (used when Stream Lookup is blank) ──────────
    "Temp (degF)",
    "Vapor Mass Flow",        # lb/hr
    "Vap MW",
    "Vap Visc (cP)",
    "Vap Z",
    "Vap Cp/Cv",
    "Vap Density (lb/ft3)",
    "Liq Mass Flow (lb/h)",
    "Liq Density (lb/ft3)",
    "Liq Visc (cP)",
    # ── Component / geometry ────────────────────────────────────────────────
    "Comp ID",                # identifier / tag label
    "Fitting Name",           # from Fittings.md drop-down
    "Bore (in)",              # pipe bore NPS in inches (e.g. 8)
    "Piping Spec",            # piping class (e.g. G1A-5)
    "Length (ft)",            # pipe run or equivalent length
    "Elev Change (ft)",       # elevation rise (+) or fall (-) in ft
    "Direction",              # N / S / E / W / UP / DOWN (informational)
    "Fixed K",                # resistance K override (Fix K / Control Valve)
    "Fixed dP (psi)",         # fixed pressure drop (Fix Pressure / Control Valve)
    "Instr Type",             # FO = Flow Orifice; blank for non-instruments
    "Instr Tag",              # instrument tag number (e.g. FI-101)
    "Instr dP (psi)",         # rated / manual pressure drop for instrument
    # ── Boundaries / control valves ─────────────────────────────────────────
    "Set P (psia)",           # boundary pressure for Source / Destination rows
    "Control Valve Type",     # P | T | F | L  (Control Valve rows only)
    "Exch Max Allow dP (psi)",  # T-control floor: max allowable dP across exchanger
    "Notes",                  # free notes
]
_NCOL = len(INPUT_HEADERS)   # 37

# Manual stream-property columns (styled yellow in the template; used when no stream)
_MANUAL_PROP_COLS = [
    "Temp (degF)", "Vapor Mass Flow", "Vap MW", "Vap Visc (cP)", "Vap Z",
    "Vap Cp/Cv", "Vap Density (lb/ft3)", "Liq Mass Flow (lb/h)",
    "Liq Density (lb/ft3)", "Liq Visc (cP)",
]

# ── Units-layer metadata per canonical column ───────────────────────────────
# canonical key → (display base name, quantity code or None)
# Quantity None ⇒ value passes through unchanged (IDs, text, NPS inches, K, Z…).
_FIELD_META: dict[str, tuple[str, str | None]] = {
    "Circuit": ("Circuit", None), "Line No": ("Line No", None),
    "HMB File": ("HMB File", None), "Case": ("Case", None),
    "Stream Lookup": ("Stream Lookup", None), "Run Type": ("Run Type", None),
    "Seq": ("Seq", None),
    "Start P (psia)": ("Start P", "P"),
    "Flow Fraction from Main": ("Flow Fraction from Main", None),
    "Mass Vapor Fraction Carry Over": ("Mass Vapor Fraction Carry Over", None),
    "Mass Liquid Fraction Carry Over": ("Mass Liquid Fraction Carry Over", None),
    "Temp (degF)": ("Temp", "T"),
    "Vapor Mass Flow": ("Vapor Mass Flow", "mflow"),
    "Vap MW": ("Vap MW", None), "Vap Visc (cP)": ("Vap Visc (cP)", None),
    "Vap Z": ("Vap Z", None), "Vap Cp/Cv": ("Vap Cp/Cv", None),
    "Vap Density (lb/ft3)": ("Vap Density", "rho"),
    "Liq Mass Flow (lb/h)": ("Liq Mass Flow", "mflow"),
    "Liq Density (lb/ft3)": ("Liq Density", "rho"),
    "Liq Visc (cP)": ("Liq Visc (cP)", None),
    "Comp ID": ("Comp ID", None), "Fitting Name": ("Fitting Name", None),
    "Bore (in)": ("Bore (in)", None),    # NPS — always inches (B36.10 convention)
    "Piping Spec": ("Piping Spec", None),
    "Length (ft)": ("Length", "L"),
    "Elev Change (ft)": ("Elev Change", "L"),
    "Direction": ("Direction", None), "Fixed K": ("Fixed K", None),
    "Fixed dP (psi)": ("Fixed dP", "dP"),
    "Instr Type": ("Instr Type", None), "Instr Tag": ("Instr Tag", None),
    "Instr dP (psi)": ("Instr dP", "dP"),
    "Set P (psia)": ("Set P", "P"),
    "Control Valve Type": ("Control Valve Type", None),
    "Exch Max Allow dP (psi)": ("Exch Max Allow dP", "dP"),
    "Notes": ("Notes", None),
}
# display base name → canonical key (for unit-suffix-tolerant header matching)
_BASE_TO_CANON = {meta[0]: k for k, meta in _FIELD_META.items()}


def _display_header(canon: str, usys: "UN.UnitSystem") -> str:
    base, qty = _FIELD_META.get(canon, (canon, None))
    return usys.hdr(base, qty) if qty else base


def _canon_of_header(h: str) -> str | None:
    """Map a sheet header (any unit suffix) back to its canonical key."""
    s = str(h).strip()
    if s in _FIELD_META:
        return s
    base = re.sub(r"\s*\([^)]*\)\s*$", "", s).strip()
    if base in _BASE_TO_CANON:
        return _BASE_TO_CANON[base]
    # legacy carry-over header names handled separately by the reader
    return None


def _hcol(name: str) -> int:
    """1-based column index of an INPUT_HEADERS name."""
    return INPUT_HEADERS.index(name) + 1


def _hcol_letter(name: str) -> str:
    return get_column_letter(_hcol(name))

# ── Per-line colour palette for multi-line circuits ────────────────────────
_LINE_PALETTE = [
    "DEEAF1",  # soft blue
    "E2EFDA",  # soft green
    "FFF2CC",  # soft amber
    "EAD1DC",  # soft pink
    "D9E2F3",  # steel blue
    "E8D5F5",  # soft purple
    "C6EFCE",  # mint
    "FCE4D6",  # soft salmon
    "F3E5F5",  # lavender
    "FFFDE7",  # soft yellow
]

def _line_color_map(line_nos_ordered: list[str]) -> dict[str, str]:
    unique = list(dict.fromkeys(line_nos_ordered))
    return {ln: _LINE_PALETTE[i % len(_LINE_PALETTE)] for i, ln in enumerate(unique)}


def _parse_hmb_filename(path: str | None) -> dict:
    """Parse an HMB filename of the form ``Unit_Case_Description_Date.xlsx``.

    e.g. ``RHC_Case1_EXHA_13Jan26.xlsx`` ->
         {unit: RHC, case_no: Case1, description: EXHA, date: 13Jan26, file: ...}

    Degrades gracefully: any missing token is left blank.  Returns the raw
    filename under ``file`` regardless.  Used for DISPLAY ONLY in Input_Pipeline.
    """
    meta = {"unit": "", "case_no": "", "description": "", "date": "", "file": ""}
    if not path:
        return meta
    base = os.path.splitext(os.path.basename(str(path)))[0]
    meta["file"] = os.path.basename(str(path))
    parts = base.split("_")
    fields = ("unit", "case_no", "description", "date")
    for i, fld in enumerate(fields):
        if i < len(parts):
            meta[fld] = parts[i].strip()
    return meta


# ════════════════════════════════════════════════════════════════════════
#  Input-template creator
# ════════════════════════════════════════════════════════════════════════
def create_input_template(out_path: str = "pipeline_input_noiso.xlsx",
                          unit_system: str = "FPS") -> str:
    """Write a blank, styled input workbook with Circuit-aware column layout.

    ``unit_system`` ("FPS" | "SI") drives the column-header unit labels and is
    recorded on the UNITS sheet; the engine reads that sheet to interpret the
    workbook, so the user may also flip it (or override per quantity) later.
    """
    usys = UN.UnitSystem(unit_system)
    wb = Workbook()
    ws = wb.active
    ws.title = "Pipeline_Input"
    ws.sheet_view.showGridLines = False

    ws.merge_cells(f"A1:{get_column_letter(_NCOL)}1")
    _c(ws, 1, 1,
       "NO-ISO PIPELINE INPUT  —  Group rows into hydraulic circuits using column A (Circuit).  "
       "All rows with the same Circuit ID are processed in series → one workbook per circuit.  "
       f"Units: {usys.system} (see UNITS sheet).",
       bg=NAVY, fg=WHITE, sz=10, bold=True, wrap=True)
    ws.row_dimensions[1].height = 28

    _hdr(ws, [_display_header(h, usys) for h in INPUT_HEADERS], row=2)
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(_NCOL)}2"

    width_map = {
        "Circuit": 9, "Line No": 12, "HMB File": 16, "Case": 10,
        "Stream Lookup": 14, "Run Type": 10, "Seq": 6, "Start P (psia)": 12,
        "Flow Fraction from Main": 14, "Temp (degF)": 11, "Vapor Mass Flow": 14,
        "Vap MW": 9, "Vap Visc (cP)": 11, "Vap Z": 8, "Vap Cp/Cv": 10,
        "Vap Density (lb/ft3)": 14, "Liq Mass Flow (lb/h)": 14,
        "Liq Density (lb/ft3)": 14, "Liq Visc (cP)": 10, "Comp ID": 12,
        "Fitting Name": 32, "Bore (in)": 9, "Piping Spec": 11, "Length (ft)": 11,
        "Elev Change (ft)": 13, "Direction": 10, "Fixed K": 9,
        "Fixed dP (psi)": 12, "Instr Type": 10, "Instr Tag": 12,
        "Instr dP (psi)": 12, "Set P (psia)": 12, "Control Valve Type": 14,
        "Exch Max Allow dP (psi)": 16, "Notes": 30,
    }
    for ci, name in enumerate(INPUT_HEADERS, 1):
        cw(ws, ci, width_map.get(name, 12))

    # ── Drop-down: Run Type ───────────────────────────────────────────
    rt_col = _hcol_letter("Run Type")
    dv_rt = DataValidation(type="list", formula1='"Main,Branch"', allow_blank=True)
    ws.add_data_validation(dv_rt)
    dv_rt.add(f"{rt_col}3:{rt_col}2000")

    # ── Drop-down: Control Valve Type (P/T/F/L) ───────────────────────
    cvt_col = _hcol_letter("Control Valve Type")
    dv_cvt = DataValidation(
        type="list", formula1='"F,P,T,L"', allow_blank=True,
        showErrorMessage=True,
        error="F=flow, P=pressure, T=temperature, L=level control",
        errorTitle="Control Valve Type")
    ws.add_data_validation(dv_cvt)
    dv_cvt.add(f"{cvt_col}3:{cvt_col}2000")

    # ── Drop-down: Fitting Name — hidden list sheet ───────────────────
    ws_fit = wb.create_sheet("_FittingList")
    ws_fit.sheet_state = "hidden"
    for i, fn in enumerate(FITTING_NAMES, 1):
        ws_fit.cell(row=i, column=1, value=fn)
    n_fit = len(FITTING_NAMES)
    fit_col = _hcol_letter("Fitting Name")
    dv_fit = DataValidation(
        type="list",
        formula1=f"_FittingList!$A$1:$A${n_fit}",
        allow_blank=True,
        showErrorMessage=True,
        error="Choose a fitting from the list",
        errorTitle="Invalid Fitting",
    )
    ws.add_data_validation(dv_fit)
    dv_fit.add(f"{fit_col}3:{fit_col}2000")

    # ── Example rows (keyed by header so column order is robust) ───────
    _ex_dicts = [
        # ── Circuit C1 — Line ER-162 (HMB stream) ─────────────────────
        {"Circuit":"C1","Line No":"ER-162","Stream Lookup":"T801-OH","Run Type":"Main","Seq":1,"Start P (psia)":100.0,"Comp ID":"ENT-01","Fitting Name":"Pipeline Entrance - Sharp Edge","Bore (in)":8,"Piping Spec":"G1A-5","Length (ft)":0,"Elev Change (ft)":0,"Notes":"Entrance"},
        {"Circuit":"C1","Line No":"ER-162","Run Type":"Main","Seq":2,"Comp ID":"P-001","Fitting Name":"Straight Pipeline","Bore (in)":8,"Piping Spec":"G1A-5","Length (ft)":50,"Elev Change (ft)":-2,"Direction":"DOWN","Notes":"Drop leg"},
        {"Circuit":"C1","Line No":"ER-162","Run Type":"Main","Seq":3,"Comp ID":"EL-001","Fitting Name":"Elbow 90 Long","Bore (in)":8,"Piping Spec":"G1A-5","Length (ft)":0,"Elev Change (ft)":0},
        {"Circuit":"C1","Line No":"ER-162","Run Type":"Main","Seq":4,"Comp ID":"TEE-01","Fitting Name":"Tee Split Line Flow","Bore (in)":8,"Piping Spec":"G1A-5","Flow Fraction from Main":-0.4,"Length (ft)":0,"Elev Change (ft)":0,"Notes":"40% splits to branch"},
        {"Circuit":"C1","Line No":"ER-162","Run Type":"Main","Seq":5,"Comp ID":"PSV-01","Fitting Name":"Fix Pressure","Bore (in)":8,"Piping Spec":"G1A-5","Fixed dP (psi)":3.5,"Length (ft)":0,"Elev Change (ft)":0,"Notes":"PSV seat loss"},
        # ── Circuit C1 — Line H-284 (continuation, same circuit) ──────
        {"Circuit":"C1","Line No":"H-284","Run Type":"Main","Seq":1,"Comp ID":"P-002","Fitting Name":"Straight Pipeline","Bore (in)":8,"Piping Spec":"G1A-5","Length (ft)":30,"Elev Change (ft)":0,"Notes":"Header run"},
        {"Circuit":"C1","Line No":"H-284","Run Type":"Main","Seq":2,"Comp ID":"EL-002","Fitting Name":"Elbow 90 Long","Bore (in)":8,"Piping Spec":"G1A-5","Length (ft)":0,"Elev Change (ft)":0},
        {"Circuit":"C1","Line No":"H-284","Run Type":"Main","Seq":3,"Comp ID":"VLV-01","Fitting Name":"Fix K","Bore (in)":8,"Piping Spec":"G1A-5","Fixed K":7.2,"Length (ft)":0,"Elev Change (ft)":0,"Notes":"CV K=7.2"},
        # ── Circuit C2 — manual properties (no HMB stream) ────────────
        {"Circuit":"C2","Line No":"H-285","Run Type":"Main","Seq":1,"Start P (psia)":80.0,"Temp (degF)":250.0,"Vapor Mass Flow":12000.0,"Vap MW":24.0,"Vap Visc (cP)":0.012,"Vap Z":0.95,"Vap Cp/Cv":1.28,"Vap Density (lb/ft3)":0.45,"Comp ID":"P-003","Fitting Name":"Straight Pipeline","Bore (in)":6,"Piping Spec":"G1A-5","Length (ft)":20,"Elev Change (ft)":0,"Notes":"Manual props (no stream)"},
        {"Circuit":"C2","Line No":"H-285","Run Type":"Main","Seq":2,"Comp ID":"EL-003","Fitting Name":"Elbow 90 Long","Bore (in)":6,"Piping Spec":"G1A-5","Length (ft)":0,"Elev Change (ft)":0},
        # ── Circuit C3 — Source → line → Control Valve (flow control) → Destination ──
        {"Circuit":"C3","Line No":"L-301","Run Type":"Main","Seq":1,"Stream Lookup":"T801-OH","Comp ID":"SRC-01","Fitting Name":"Source","Bore (in)":6,"Piping Spec":"G1A-5","Set P (psia)":150.0,"Length (ft)":0,"Elev Change (ft)":0,"Notes":"Upstream source pressure"},
        {"Circuit":"C3","Line No":"L-301","Run Type":"Main","Seq":2,"Comp ID":"P-301","Fitting Name":"Straight Pipeline","Bore (in)":6,"Piping Spec":"G1A-5","Length (ft)":40,"Elev Change (ft)":0,"Notes":"Run to valve"},
        {"Circuit":"C3","Line No":"L-301","Run Type":"Main","Seq":3,"Comp ID":"FCV-301","Fitting Name":"Control Valve","Bore (in)":4,"Piping Spec":"G1A-5","Control Valve Type":"F","Length (ft)":0,"Elev Change (ft)":0,"Notes":"Flow control — ΔP floats to Destination P"},
        {"Circuit":"C3","Line No":"L-301","Run Type":"Main","Seq":4,"Comp ID":"DST-01","Fitting Name":"Destination","Bore (in)":6,"Piping Spec":"G1A-5","Set P (psia)":60.0,"Length (ft)":0,"Elev Change (ft)":0,"Notes":"Downstream destination pressure"},
        # ── Circuit C4 — two parallel control valves (Tee split, ratio control) ──
        {"Circuit":"C4","Line No":"L-401","Run Type":"Main","Seq":1,"Stream Lookup":"T801-OH","Comp ID":"SRC-02","Fitting Name":"Source","Bore (in)":8,"Piping Spec":"G1A-5","Set P (psia)":200.0,"Length (ft)":0,"Elev Change (ft)":0,"Notes":"Header source"},
        {"Circuit":"C4","Line No":"L-401","Run Type":"Main","Seq":2,"Comp ID":"TEE-40","Fitting Name":"Tee Split Branch Flow 1","Bore (in)":8,"Piping Spec":"G1A-5","Flow Fraction from Main":-0.5,"Length (ft)":0,"Elev Change (ft)":0,"Notes":"50% splits to parallel branch"},
        {"Circuit":"C4","Line No":"L-401","Run Type":"Main","Seq":3,"Comp ID":"FCV-40A","Fitting Name":"Control Valve","Bore (in)":4,"Piping Spec":"G1A-5","Control Valve Type":"F","Fixed dP (psi)":25.0,"Length (ft)":0,"Elev Change (ft)":0,"Notes":"Main-leg valve (50% flow)"},
        {"Circuit":"C4","Line No":"L-401","Run Type":"Branch","Seq":1,"Start P (psia)":200.0,"Comp ID":"FCV-40B","Fitting Name":"Control Valve","Bore (in)":4,"Piping Spec":"G1A-5","Control Valve Type":"F","Fixed dP (psi)":25.0,"Length (ft)":0,"Elev Change (ft)":0,"Notes":"Parallel-leg valve (other 50% flow)"},
    ]
    manual_idx = {_hcol(n) for n in _MANUAL_PROP_COLS}
    for ri, d in enumerate(_ex_dicts, 3):
        for ci, name in enumerate(INPUT_HEADERS, 1):
            v = d.get(name, "")
            qty = _FIELD_META.get(name, (name, None))[1]
            if qty and isinstance(v, (int, float)):
                v = round(usys.from_internal(qty, float(v)), 4)
            cell = ws.cell(row=ri, column=ci, value=(v if v != "" else None))
            cell.font = Font(name="Calibri", size=9, italic=True, color="808080")
            cell.fill = PatternFill("solid", fgColor=(AMBER if ci in manual_idx else WHITE))
            cell.alignment = Alignment(horizontal="left", vertical="center")

    # ── Notes sheet ──────────────────────────────────────────────────
    ws_n = wb.create_sheet("Notes")
    ws_n.sheet_view.showGridLines = False
    cw(ws_n, 1, 110)
    note_lines = [
        ("NO-ISO PIPELINE INPUT — QUICK REFERENCE", True),
        ("", False),
        ("CIRCUIT (column A)", True),
        ("  All rows with the same Circuit ID (e.g. C1) are treated as ONE hydraulic series path.", False),
        ("  Lines are chained in the ORDER they appear in this sheet.", False),
        ("  The engine produces ONE workbook per circuit.", False),
        ("  Stream Lookup, HMB File, Case and Start P need only be filled on the FIRST row of the circuit.", False),
        ("", False),
        ("HMB FILE & CASE", True),
        ("  HMB File and Case override the command-line defaults per row (carried forward within a circuit).", False),
        ("  This lets ONE circuit mix multiple streams / cases / HMB workbooks.", False),
        ("  CLI default case is set with  --case \"Case 1\".", False),
        ("", False),
        ("MANUAL STREAM PROPERTIES (yellow cells)", True),
        ("  Leave Stream Lookup BLANK to enter properties by hand in the yellow columns.", False),
        ("  Fill Temp, Vapor Mass Flow, Vap MW/Visc/Z/Cp-Cv/Density and (optionally) the Liq columns.", False),
        ("  The engine synthesises a frozen flash split from these numbers — no HMB needed.", False),
        ("  OVERRIDE MODE: with Stream Lookup filled (pulling an HMB stream), filling a yellow cell on", False),
        ("  ANY row of that stream makes the typed value win over the HMB property for that field —", False),
        ("  e.g. type a Vapor Mass Flow to test a different rate without touching the HMB file.", False),
        ("  The override persists on every row downstream until changed again or a new stream loads.", False),
        ("", False),
        ("FLOW FRACTION FROM MAIN (Tee split / merge)", True),
        ("  On a Tee fitting row, enter a SIGNED fraction of the Main flow:", False),
        ("    negative  = SPLIT   (that fraction of the CURRENT flow leaves the line)", False),
        ("    positive  = MERGE   (running flow becomes that fraction of the ORIGINAL Main flow —", False),
        ("                         an absolute target, e.g. 0.125 → 0.25 → 0.5 → 1.0 builds a header", False),
        ("                         up in stages as branches join it)", False),
        ("  The running vapour & liquid mass flows are scaled in Seq order; downstream dP follows.", False),
        ("", False),
        ("MULTIPLE LINES IN SERIES", True),
        ("  Change Line No to mark a new pipe line.  Keep Circuit the same.", False),
        ("  The outlet pressure of the last component of Line N becomes the inlet pressure of Line N+1.", False),
        ("", False),
        ("FITTINGS", True),
        ("  Select from the Fitting Name drop-down.  K values applied automatically.", False),
        ("", False),
        ("FIX K  (Fitting = 'Fix K')", True),
        ("  Enter K in the Fixed K column.  Darcy-Weisbach friction and elevation terms still apply.", False),
        ("", False),
        ("FIX PRESSURE  (Fitting = 'Fix Pressure')", True),
        ("  Enter dP (psi) in Fixed dP.  All hydraulics skipped; dP subtracted directly.", False),
        ("", False),
        ("FLOW ORIFICE  (Instr Type = FO)", True),
        ("  Set Instr Type = FO.  Enter rated dP in Instr dP (psi).", False),
        ("", False),
        ("SOURCE / DESTINATION  (boundary fittings)", True),
        ("  Source: enter the upstream pressure in 'Set P (psia)' — it anchors the circuit inlet pressure.", False),
        ("  Destination: enter the downstream pressure in 'Set P (psia)' — the target the circuit must reach.", False),
        ("  If BOTH are filled, a Control Valve placed between them absorbs the pressure difference.", False),
        ("  If only one is filled, the pressure is taken from that single value.", False),
        ("", False),
        ("CONTROL VALVE  (Fitting = 'Control Valve')", True),
        ("  Set 'Control Valve Type' = F (flow), P (pressure), T (temperature) or L (level).", False),
        ("  F (flow): ΔP floats so the running pressure lands on the downstream Destination 'Set P'.", False),
        ("  P / L: fixed design ΔP — enter it in 'Fixed dP (psi)' (or a resistance in 'Fixed K').", False),
        ("  T (temperature): like a fixed ΔP, but floored at 'Exch Max Allow dP (psi)' (exchanger limit).", False),
        ("  β ratio = valve Bore (in) ÷ upstream line bore is reported in the Notes column.", False),
        ("  SERIES valves: place several Control Valve rows in Seq order on the same run.", False),
        ("  PARALLEL valves: split flow at a Tee (negative 'Flow Fraction from Main'), put one Control", False),
        ("    Valve on the Main leg and one on a Branch (Run Type = Branch, with its own Start P).", False),
        ("", False),
        ("BRANCHES", True),
        ("  Set Run Type = Branch.  Fill Start P on the first Branch row.", False),
        ("  Branches are marched independently.  FIV/AIV/regime sheets use the Main run.", False),
    ]
    note_lines += [
        ("", False),
        ("UNITS", True),
        (f"  This workbook is in {usys.system} units (see the UNITS sheet).", False),
        ("  Change the Unit System cell (or per-quantity overrides) on the UNITS "
         "sheet and enter values accordingly —", False),
        ("  the engine reads the UNITS sheet and presents all results in the "
         "same units.  Bore/NPS is ALWAYS inches.", False),
    ]
    for i, (t, bold) in enumerate(note_lines, 1):
        ws_n.cell(i, 1).value = t
        ws_n.cell(i, 1).font = Font(name="Calibri", size=9, bold=bold,
                                     color=NAVY if bold else DGRAY)

    UN.add_units_sheet(wb, usys.system, position=len(wb.worksheets))
    wb.save(out_path)
    return out_path


# ════════════════════════════════════════════════════════════════════════
#  Input reader
# ════════════════════════════════════════════════════════════════════════
def read_pipeline_input(xlsx_path: str) -> list[dict]:
    """Return list of row dicts from the Pipeline_Input sheet.

    The ``Circuit`` column (col A) groups rows into hydraulic series paths.
    ``Stream Lookup`` and ``Start P`` are carried forward within a circuit.
    """
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = next(
        (s for s in wb.worksheets if s.title == "Pipeline_Input"),
        next((s for s in wb.worksheets if s.title not in ("Notes", "_FittingList")),
             None),
    )
    if ws is None:
        wb.close()
        return []

    usys = UN.UnitSystem.from_workbook(wb)

    hdr_vals = list(ws.iter_rows(min_row=2, max_row=2, values_only=True))[0]
    hdr: dict[str, int] = {}
    for i, h in enumerate(hdr_vals):
        if h is None:
            continue
        raw = str(h).strip()
        hdr.setdefault(raw, i)                  # raw name (legacy headers too)
        canon = _canon_of_header(raw)
        if canon:
            hdr.setdefault(canon, i)            # canonical key, any unit suffix

    def g(rv, key):
        ci = hdr.get(key)
        return rv[ci] if (ci is not None and ci < len(rv)) else None

    def gnum(rv, key):
        """num() + user-units → internal-FPS conversion per _FIELD_META."""
        v = num(g(rv, key))
        qty = _FIELD_META.get(key, (key, None))[1]
        return usys.to_internal(qty, v) if (qty and v is not None) else v

    rows: list[dict] = []
    seq_ctr: dict[tuple, int] = {}
    current_circuit: str | None = None
    current_stream:  str | None = None   # stream lookup carried within circuit
    current_start_p: float | None = None # start P carried within circuit
    current_hmb:     str | None = None   # HMB File carried within circuit
    current_case:    str | None = None   # Case carried within circuit

    for rv in ws.iter_rows(min_row=3, values_only=True):
        if not rv or all(v is None for v in rv):
            continue
        line_no = txt(g(rv, "Line No"))
        if not line_no:
            continue

        # ── Skip an intentionally blank/contentless component row ──────────
        # A row with neither a Comp ID nor a Fitting Name carries no component;
        # ignore it completely (do not number it, do not disturb carries) so the
        # series keeps reading.  A row that only sets Start P / stream still
        # counts as real and is handled below.
        if (not txt(g(rv, "Comp ID")) and not txt(g(rv, "Fitting Name"))
                and num(g(rv, "Start P (psia)")) is None
                and not txt(g(rv, "Stream Lookup"))
                and not txt(g(rv, "HMB File")) and not txt(g(rv, "Case"))):
            continue

        circuit = txt(g(rv, "Circuit")) or current_circuit or "DEFAULT"
        # Reset carries when circuit changes
        if circuit != current_circuit:
            current_circuit = circuit
            current_stream  = None
            current_start_p = None
            current_hmb     = None
            current_case    = None

        sl = txt(g(rv, "Stream Lookup"))
        if sl:
            current_stream = sl
        hmb_f = txt(g(rv, "HMB File"))
        if hmb_f:
            current_hmb = hmb_f
        case_v = txt(g(rv, "Case"))
        if case_v:
            current_case = case_v
        sp_val = gnum(rv, "Start P (psia)")
        if sp_val is not None:
            current_start_p = sp_val

        run_type = txt(g(rv, "Run Type")) or "Main"
        seq_key  = (circuit, line_no, run_type)
        seq_raw  = num(g(rv, "Seq"))
        if seq_raw is not None:
            seq = int(seq_raw)
        else:
            seq_ctr[seq_key] = seq_ctr.get(seq_key, 0) + 1
            seq = seq_ctr[seq_key]

        rows.append({
            "Circuit":          circuit,
            "Line No":          line_no,
            "HMB File":         current_hmb,
            "Case":             current_case,
            "Stream Lookup":    current_stream,
            "Start P (psia)":   sp_val,          # raw (only non-None on explicit rows)
            "Run Type":         run_type,
            "Seq":              seq,
            "Flow Fraction from Main":      num(g(rv, "Flow Fraction from Main")),
            # MASS-basis phase carry-over (back-compat: also read the old header names)
            "Mass Vapor Fraction Carry Over":
                num(g(rv, "Mass Vapor Fraction Carry Over"))
                if g(rv, "Mass Vapor Fraction Carry Over") is not None
                else num(g(rv, "Vapor Fraction Carry Forward")),
            "Mass Liquid Fraction Carry Over":
                num(g(rv, "Mass Liquid Fraction Carry Over"))
                if g(rv, "Mass Liquid Fraction Carry Over") is not None
                else num(g(rv, "Liquid Fraction Carry Forward")),
            # ── Manual stream properties (None when blank) ──────────────
            "Temp (degF)":          gnum(rv, "Temp (degF)"),
            "Vapor Mass Flow":      gnum(rv, "Vapor Mass Flow"),
            "Vap MW":               num(g(rv, "Vap MW")),
            "Vap Visc (cP)":        num(g(rv, "Vap Visc (cP)")),
            "Vap Z":                num(g(rv, "Vap Z")),
            "Vap Cp/Cv":            num(g(rv, "Vap Cp/Cv")),
            "Vap Density (lb/ft3)": gnum(rv, "Vap Density (lb/ft3)"),
            "Liq Mass Flow (lb/h)": gnum(rv, "Liq Mass Flow (lb/h)"),
            "Liq Density (lb/ft3)": gnum(rv, "Liq Density (lb/ft3)"),
            "Liq Visc (cP)":        num(g(rv, "Liq Visc (cP)")),
            # ── Geometry / component ────────────────────────────────────
            "Comp ID":          txt(g(rv, "Comp ID")),
            "Fitting Name":     txt(g(rv, "Fitting Name")),
            "Bore (in)":        num(g(rv, "Bore (in)")),
            "Piping Spec":      txt(g(rv, "Piping Spec")),
            "Length (ft)":      gnum(rv, "Length (ft)") or 0.0,
            "Elev Change (ft)": gnum(rv, "Elev Change (ft)") or 0.0,
            "Direction":        txt(g(rv, "Direction")),
            "Fixed K":          num(g(rv, "Fixed K")),
            "Fixed dP (psi)":   gnum(rv, "Fixed dP (psi)"),
            "Instr Type":       txt(g(rv, "Instr Type")),
            "Instr Tag":        txt(g(rv, "Instr Tag")),
            "Instr dP (psi)":   gnum(rv, "Instr dP (psi)"),
            # ── Boundaries / control valves ─────────────────────────────
            "Set P (psia)":     gnum(rv, "Set P (psia)"),
            "Control Valve Type": txt(g(rv, "Control Valve Type")),
            "Exch Max Allow dP (psi)": gnum(rv, "Exch Max Allow dP (psi)"),
            "Notes":            txt(g(rv, "Notes")),
        })
    wb.close()
    return rows


def _group_circuits(rows: list[dict]) -> dict[str, dict]:
    """
    Return {circuit_id: {"main_rows": [...], "branch_blocks": [[...], ...],
                          "stream_lookup": str, "start_p": float,
                          "line_order": [line_nos in appearance order]}}.

    ``main_rows``    — all Main-run rows across all lines, in order (for series chain).
    ``branch_blocks``— list of Branch row-blocks per line, each starting at its own P.
    """
    circuits: dict[str, dict] = {}

    for r in rows:
        cid = r["Circuit"]
        if cid not in circuits:
            circuits[cid] = {
                "main_rows":     [],
                "branch_blocks": [],
                "stream_lookup": None,
                "hmb_file":      None,
                "case":          None,
                "start_p":       None,
                "line_order":    [],
            }
        c = circuits[cid]

        # Capture stream / HMB / case and start P from first non-None occurrence
        if c["stream_lookup"] is None and r.get("Stream Lookup"):
            c["stream_lookup"] = r["Stream Lookup"]
        if c["hmb_file"] is None and r.get("HMB File"):
            c["hmb_file"] = r["HMB File"]
        if c["case"] is None and r.get("Case"):
            c["case"] = r["Case"]
        if c["start_p"] is None and r.get("Start P (psia)") is not None:
            c["start_p"] = r["Start P (psia)"]
        # A Source boundary's Set P also seeds the circuit start pressure.
        if (c["start_p"] is None and _is_source(r.get("Fitting Name"))
                and r.get("Set P (psia)") is not None):
            c["start_p"] = r["Set P (psia)"]

        ln = r["Line No"]
        if ln not in c["line_order"]:
            c["line_order"].append(ln)

        if r["Run Type"] == "Main":
            c["main_rows"].append(r)
        else:
            # Group consecutive Branch rows into blocks (split on new Start P)
            if (not c["branch_blocks"]
                    or (r.get("Start P (psia)") is not None
                        and c["branch_blocks"][-1])):
                c["branch_blocks"].append([r])
            else:
                c["branch_blocks"][-1].append(r)

    return circuits


def _start_p_of_block(block: list[dict]) -> float | None:
    for r in block:
        if r.get("Start P (psia)") is not None:
            return r["Start P (psia)"]
        if (_is_source(r.get("Fitting Name"))
                and r.get("Set P (psia)") is not None):
            return r["Set P (psia)"]
    return None


# ════════════════════════════════════════════════════════════════════════
#  Flash VLE helpers  (duplicated from hydraulics_engine_flash.py so
#  this file is self-contained when flash.py changes)
# ════════════════════════════════════════════════════════════════════════
def _gas_rho_kgm3(vap_mw, z, t_f, p_psia) -> float:
    if not (vap_mw and t_f is not None):
        return 0.0
    t_k = (t_f + F_TO_K_OFFSET) / 1.8
    return p_psia * PSIA_TO_PA * (vap_mw / 1000.0) / ((z or 1.0) * R_UNIV * t_k)


def _phase_of(quality: float) -> str:
    if quality <= 1e-5:
        return "LIQUID"
    if quality >= 1.0 - 1e-5:
        return "VAPOR"
    return "TWO-PHASE"


def _dp_flashed(g: HE.Geom, sp: HE.StreamProps, fr) -> dict:
    """Δp using flashed vapor/liquid split (homogeneous model)."""
    rho_g = _gas_rho_kgm3(fr.vap_mw, sp.vap_z, sp.temp_f, fr.p_psia)
    rho_l = (sp.liq_density or 0.0) * LBFT3_TO_KGM3
    mu_g  = (sp.vap_visc or 0.0) * CP_TO_PAS
    mu_l  = (sp.liq_visc or 0.0) * CP_TO_PAS
    vm    = fr.vap_mass * LBHR_TO_KGS
    lm    = fr.liq_mass * LBHR_TO_KGS
    mdot  = vm + lm
    if mdot <= 0 or g.area_m2 <= 0:
        return dict(rho=rho_g, v=0.0, Re=0.0, f=0.0, dp_f=0.0, dp_k=0.0, dp_z=0.0)
    x = vm / mdot
    if x <= 1e-5 or rho_g <= 0:
        rho, mu = rho_l, mu_l
    elif x >= 1.0 - 1e-5 or rho_l <= 0:
        rho, mu = rho_g, mu_g
    else:
        rho = 1.0 / (x / rho_g + (1.0 - x) / rho_l)
        mu  = (1.0 / (x / mu_g + (1.0 - x) / mu_l)
               if (mu_g > 0 and mu_l > 0) else (mu_g or mu_l))
    Gflux = mdot / g.area_m2
    v     = Gflux / rho if rho > 0 else 0.0
    Re    = Gflux * g.id_m / mu if mu > 0 else 0.0
    f     = HE.darcy_f(Re, g.eps_over_d)
    dyn   = (Gflux * Gflux) / (2.0 * rho) if rho > 0 else 0.0
    dp_f  = f * (g.length_m / g.id_m) * dyn if g.id_m > 0 else 0.0
    dp_k  = g.k * dyn
    dp_z  = rho * G_SI * g.dz_m
    return dict(rho=rho, v=v, Re=Re, f=f, dp_f=dp_f, dp_k=dp_k, dp_z=dp_z)


# ════════════════════════════════════════════════════════════════════════
#  Manual-stream support (used when no HMB Stream Lookup is supplied)
# ════════════════════════════════════════════════════════════════════════
def _streamprops_from_manual(row: dict) -> HE.StreamProps | None:
    """Build a StreamProps from the yellow manual-input cells of a row.

    Returns None when the row carries no usable manual stream data.
    """
    vap_mass = num(row.get("Vapor Mass Flow"))
    liq_mass = num(row.get("Liq Mass Flow (lb/h)"))
    temp_f   = num(row.get("Temp (degF)"))
    vap_mw   = num(row.get("Vap MW"))
    vap_den  = num(row.get("Vap Density (lb/ft3)"))
    liq_den  = num(row.get("Liq Density (lb/ft3)"))

    # Need at least some flow and a way to get a density to be useful.
    if not any(v for v in (vap_mass, liq_mass)):
        return None

    vm = vap_mass or 0.0
    lm = liq_mass or 0.0
    if vm > 0 and lm > 0:
        phase = "TWO-PHASE"
    elif vm > 0:
        phase = "VAPOR"
    else:
        phase = "LIQUID"

    return HE.StreamProps(
        stream      = "(manual)",
        phase       = phase,
        source_sheet= "manual-input",
        temp_f      = temp_f,
        pres_psia   = num(row.get("Start P (psia)")),
        total_mass  = vm + lm,
        vap_mass    = vap_mass,
        liq_mass    = liq_mass,
        mol_weight  = vap_mw,
        vap_mw      = vap_mw,
        vap_density = vap_den,
        liq_density = liq_den,
        vap_visc    = num(row.get("Vap Visc (cP)")),
        liq_visc    = num(row.get("Liq Visc (cP)")),
        vap_z       = num(row.get("Vap Z")),
        vap_cp_cv   = num(row.get("Vap Cp/Cv")),
    )


def _synth_flash_result(sp: HE.StreamProps, p_psia: float):
    """Synthesise a frozen FlashResult from manual StreamProps.

    No composition is available, so the vapour/liquid split is held constant
    (frozen) at the user-supplied mass flows.  This lets the flash-coupled
    marcher and all downstream sheets render in manual mode.
    """
    vap_mass = sp.vap_mass or 0.0
    liq_mass = sp.liq_mass or 0.0
    vap_mw   = sp.vap_mw or sp.mol_weight or 1.0
    liq_mw   = sp.liq_mw or sp.mol_weight or vap_mw or 1.0
    tot      = vap_mass + liq_mass
    quality  = (vap_mass / tot) if tot > 0 else (1.0 if vap_mass > 0 else 0.0)
    vap_moles = (vap_mass / vap_mw) if vap_mw else 0.0
    liq_moles = (liq_mass / liq_mw) if liq_mw else 0.0
    tot_moles = vap_moles + liq_moles
    beta      = (vap_moles / tot_moles) if tot_moles > 0 else quality
    return FV.FlashResult(
        p_psia    = p_psia,
        beta      = beta,
        quality   = quality,
        vap_moles = vap_moles,
        liq_moles = liq_moles,
        vap_mass  = vap_mass,
        liq_mass  = liq_mass,
        vap_mw    = vap_mw,
        liq_mw    = liq_mw,
        temp_f    = sp.temp_f,
    )


def _scale_flash_result(fr, vap_mult: float, liq_mult: float | None = None):
    """Return a copy of a FlashResult with vapour & liquid flows scaled
    **independently**.

    ``vap_mult`` scales the vapour phase, ``liq_mult`` the liquid phase
    (defaults to ``vap_mult`` when omitted → uniform scaling).  Quality and beta
    are recomputed from the scaled masses/moles so a phase-specific change (e.g.
    dumping the liquid) is reflected correctly downstream.
    """
    if fr is None:
        return fr
    if liq_mult is None:
        liq_mult = vap_mult
    vs = max(0.0, vap_mult)
    ls = max(0.0, liq_mult)
    if vs == 1.0 and ls == 1.0:
        return fr
    vm = fr.vap_mass * vs
    lm = fr.liq_mass * ls
    vmol = fr.vap_moles * vs
    lmol = fr.liq_moles * ls
    tot_m   = vm + lm
    tot_mol = vmol + lmol
    return replace(
        fr,
        vap_mass  = vm,
        liq_mass  = lm,
        vap_moles = vmol,
        liq_moles = lmol,
        quality   = (vm / tot_m) if tot_m > 0 else 0.0,
        beta      = (vmol / tot_mol) if tot_mol > 0 else 0.0,
    )


def _is_tee(fitting: str | None) -> bool:
    f = (fitting or "").lower()
    return "tee" in f or "stub" in f or "branch" in f


def _is_tee_join(fitting: str | None) -> int | None:
    """If ``fitting`` is a 'Tee Join Branch Flow N', return N; else None."""
    f = (fitting or "").lower()
    if "tee" in f and "join" in f and "branch" in f:
        m = re.search(r"(\d+)\s*$", f)
        return int(m.group(1)) if m else 1
    return None


def _feed_comp_flows(feed, fr, vap_mult: float = 1.0, liq_mult: float = 1.0) -> dict:
    """Per-component molar flow (lb-mol/hr) for a station's (scaled) flash.

    Uses the conserved per-component total = feed.z_i · total_moles, applying the
    vapour/liquid multipliers via the phase split so dumping a phase changes the
    reported composition.  Returns {} when no feed/composition is available.
    """
    if feed is None or fr is None or not getattr(feed, "names", None):
        return {}
    # Per-phase moles already reflect any scaling applied to ``fr``; if explicit
    # multipliers are passed (un-scaled fr), apply them here instead.
    vmol = fr.vap_moles * vap_mult
    lmol = fr.liq_moles * liq_mult
    y = fr.y or []
    x = fr.x or []
    out: dict[str, float] = {}
    for i, name in enumerate(feed.names):
        yi = y[i] if i < len(y) else 0.0
        xi = x[i] if i < len(x) else 0.0
        flow = vmol * yi + lmol * xi
        if flow:
            out[name] = out.get(name, 0.0) + flow
    return out


def _feed_props_lookup(feed, p_ref_target: float) -> dict:
    """name → {mw, k_ref, tc_r, omega, pc} with k_ref normalised to a target p_ref.

    K_i(P) = k_ref_i · (p_ref/P) is composition-independent, so a feed's k_ref
    referenced at its own ``p_ref`` is converted to the blended feed's reference
    by k_ref' = k_ref · (p_ref_feed / p_ref_target).
    """
    out: dict = {}
    if feed is None or not getattr(feed, "names", None):
        return out
    fac = (feed.p_ref_psia / p_ref_target) if p_ref_target else 1.0
    tc = feed.tc_r or [None] * feed.n
    om = feed.omega or [None] * feed.n
    pc = feed.pc_psia or [None] * feed.n
    for i, nm in enumerate(feed.names):
        out[nm] = {
            "mw":    feed.mw[i] if i < len(feed.mw) else None,
            "k_ref": (feed.k_ref[i] * fac) if i < len(feed.k_ref) else 1.0,
            "tc_r":  tc[i] if i < len(tc) else None,
            "omega": om[i] if i < len(om) else None,
            "pc":    pc[i] if i < len(pc) else None,
        }
    return out


def _build_feed_from_moles(moles: dict, props: dict, p_ref: float,
                           temp_f, k_basis: str = "yx", proto=None):
    """Construct a FlashFeed from blended component moles + a property lookup.

    ``moles``  : {name: lb-mol/hr}  (combined mixture)
    ``props``  : {name: {mw,k_ref,tc_r,omega,pc}} (mole-weighted where shared)
    Enthalpy reference scalars are inherited from ``proto`` (the main feed).
    """
    names = [n for n, m in moles.items() if m and m > 0]
    if len(names) < 2 or not p_ref:
        return None
    F = sum(moles[n] for n in names)
    z, mw, k_ref, tc_r, omega, pc_l = [], [], [], [], [], []
    for n in names:
        pr = props.get(n, {})
        z.append(moles[n] / F)
        mw.append(pr.get("mw") or 50.0)
        k_ref.append(min(1e10, max(1e-10, pr.get("k_ref") or 1.0)))
        tc_r.append(pr.get("tc_r")); omega.append(pr.get("omega")); pc_l.append(pr.get("pc"))
    total_mw   = sum(zi * mi for zi, mi in zip(z, mw))
    total_mass = sum(moles[n] * (props.get(n, {}).get("mw") or 50.0) for n in names)
    feed = FV.FlashFeed(
        names=names, z=z, k_ref=k_ref, mw=mw,
        total_mass_lbhr=total_mass, p_ref_psia=p_ref, temp_f=temp_f,
        total_mw=total_mw, total_moles_lbmolhr=F,
        tc_r=tc_r, omega=omega, pc_psia=pc_l,
        vap_h_ref=getattr(proto, "vap_h_ref", None),
        liq_h_ref=getattr(proto, "liq_h_ref", None),
        vap_cp=getattr(proto, "vap_cp", None),
        liq_cp=getattr(proto, "liq_cp", None),
        k_basis=k_basis)
    feed.beta_ref = FV.rachford_rice(z, k_ref)
    return feed


def _blend_feeds(feed_a, moles_a: dict, feed_b, moles_b: dict):
    """Blend two mixtures (by component moles) into ONE new FlashFeed.

    The combined mixture is referenced to ``feed_a``'s p_ref/temperature and is
    re-flashed downstream as a brand-new feed.  Shared components mole-weight
    their (p_ref-normalised) k_ref.
    """
    if feed_a is None:
        return None
    p_ref = feed_a.p_ref_psia
    pa = _feed_props_lookup(feed_a, p_ref)
    pb = _feed_props_lookup(feed_b, p_ref) if feed_b is not None else {}
    moles: dict = {}
    for n, m in moles_a.items():
        if m:
            moles[n] = moles.get(n, 0.0) + m
    for n, m in moles_b.items():
        if m:
            moles[n] = moles.get(n, 0.0) + m
    props: dict = {}
    for n in moles:
        a, b = pa.get(n), pb.get(n)
        if a and b:
            wa, wb = moles_a.get(n, 0.0), moles_b.get(n, 0.0)
            tot = (wa + wb) or 1.0
            props[n] = {
                "mw":    a["mw"] or b["mw"],
                "k_ref": (wa * (a["k_ref"] or 1.0) + wb * (b["k_ref"] or 1.0)) / tot,
                "tc_r":  a["tc_r"] if a["tc_r"] is not None else b["tc_r"],
                "omega": a["omega"] if a["omega"] is not None else b["omega"],
                "pc":    a["pc"] if a["pc"] is not None else b["pc"],
            }
        else:
            props[n] = a or b or {}
    return _build_feed_from_moles(moles, props, p_ref, feed_a.temp_f,
                                  k_basis=feed_a.k_basis, proto=feed_a)


def _blend_sp(sp_a, sp_b, mass_a: float, mass_b: float):
    """Mass-weighted blend of the StreamProps scalars used for ΔP (densities,
    viscosities, Z, MW).  Returns a copy of ``sp_a`` with blended values."""
    if sp_a is None:
        return sp_b
    if sp_b is None or (mass_a + mass_b) <= 0:
        return sp_a
    wa = mass_a / (mass_a + mass_b)
    wb = 1.0 - wa

    def mix(av, bv):
        if av is None:
            return bv
        if bv is None:
            return av
        return wa * av + wb * bv

    phase = "TWO-PHASE" if ((sp_a.vap_mass or sp_b.vap_mass)
                            and (sp_a.liq_mass or sp_b.liq_mass)) else sp_a.phase
    return replace(
        sp_a,
        phase       = phase,
        vap_density = mix(sp_a.vap_density, sp_b.vap_density),
        liq_density = mix(sp_a.liq_density, sp_b.liq_density),
        vap_visc    = mix(sp_a.vap_visc,    sp_b.vap_visc),
        liq_visc    = mix(sp_a.liq_visc,    sp_b.liq_visc),
        vap_z       = mix(sp_a.vap_z,       sp_b.vap_z),
        vap_mw      = mix(sp_a.vap_mw,      sp_b.vap_mw),
        total_mass  = (sp_a.total_mass or 0.0) + (sp_b.total_mass or 0.0),
    )


# ════════════════════════════════════════════════════════════════════════
#  No-ISO pressure marcher
# ════════════════════════════════════════════════════════════════════════
# ── Yellow-cell property overrides (win over a resolved HMB stream) ──────
# Maps an input-sheet column to the StreamProps/FlashResult field it overrides.
_YELLOW_OVERRIDE_MAP = {
    "Temp (degF)":          "temp_f",
    "Vapor Mass Flow":      "vap_mass",
    "Vap MW":               "vap_mw",
    "Vap Visc (cP)":        "vap_visc",
    "Vap Z":                "vap_z",
    "Vap Cp/Cv":            "vap_cp_cv",
    "Vap Density (lb/ft3)": "vap_density",
    "Liq Mass Flow (lb/h)": "liq_mass",
    "Liq Density (lb/ft3)": "liq_density",
    "Liq Visc (cP)":        "liq_visc",
}


def _row_yellow_overrides(row: dict) -> dict:
    """Non-blank yellow manual-input cells on this row, keyed by field name."""
    out = {}
    for col, field_name in _YELLOW_OVERRIDE_MAP.items():
        v = num(row.get(col))
        if v is not None:
            out[field_name] = v
    return out


def _apply_property_overrides(sp, fr, overrides: dict):
    """Overlay persisted yellow-cell overrides onto a resolved sp/fr pair.

    Lets any filled yellow cell win over the matching property pulled from
    the HMB stream, without having to clear Stream Lookup and re-enter every
    property by hand. Mass-flow overrides also re-derive quality/beta/moles
    on the FlashResult so downstream phase split & ΔP stay consistent.
    """
    if not overrides or sp is None:
        return sp, fr
    sp_fields = dict(overrides)
    if "vap_mw" in sp_fields:
        sp_fields["mol_weight"] = sp_fields["vap_mw"]
    if "vap_mass" in sp_fields or "liq_mass" in sp_fields:
        sp_fields["total_mass"] = ((sp_fields.get("vap_mass", sp.vap_mass) or 0.0)
                                    + (sp_fields.get("liq_mass", sp.liq_mass) or 0.0))
    sp = replace(sp, **sp_fields)

    if fr is not None:
        fr_fields = {k: overrides[k] for k in ("vap_mass", "liq_mass", "vap_mw", "temp_f")
                     if k in overrides}
        if fr_fields:
            vm  = fr_fields.get("vap_mass", fr.vap_mass)
            lm  = fr_fields.get("liq_mass", fr.liq_mass)
            vmw = fr_fields.get("vap_mw", fr.vap_mw) or 1.0
            lmw = fr.liq_mw or vmw
            tot = vm + lm
            vmol, lmol = (vm / vmw if vmw else 0.0), (lm / lmw if lmw else 0.0)
            tmol = vmol + lmol
            fr_fields["quality"]   = (vm / tot) if tot > 0 else (1.0 if vm > 0 else 0.0)
            fr_fields["vap_moles"] = vmol
            fr_fields["liq_moles"] = lmol
            fr_fields["beta"]      = (vmol / tmol) if tmol > 0 else fr_fields["quality"]
            fr = replace(fr, **fr_fields)
    return sp, fr


def build_profile_flash_noiso(
        block_rows: list[dict],
        resolver,
        p_start: float,
        p_floor: float = 0.05,
        flash_mode: str = "isothermal",
        default_sp: HE.StreamProps | None = None,
        injections: dict | None = None,
        cv_overrides: dict | None = None,
) -> tuple[list[HE.Station], list, list, dict]:
    """
    March pressure along ``block_rows`` with flash-coupled VLE.

    Special fittings handled:
      Fix Pressure — subtract ``Fixed dP (psi)`` directly; no velocity calc
      Fix K        — Darcy-Weisbach with user ``Fixed K`` value; friction + elev normal
      Orifice/FO   — subtract ``Instr dP (psi)`` directly; no velocity calc

    Tee split / merge:
      A ``Flow Fraction from Main`` on a Tee fitting row scales the running
      vapour & liquid mass flows (negative = split, positive = merge).

    Mass Vapor / Liquid carry-over:
      ``Mass Vapor/Liquid Fraction Carry Over`` scale the running vapour and
      liquid MASS (and moles equally) independently and persist (compound)
      downstream.  ``1,1`` = unchanged; ``1,0`` = vapour only (dump liquid mass);
      ``0,0`` = stream terminates.

    Branch→Main join:
      ``injections`` maps a block-row index → a branch outlet payload
      (absolute vap/liq mass & moles + component vector).  At that row the branch
      flow is merged into the running stream and persists downstream.

    Multi-stream / manual modes work as before.

    Returns (stations, flashes, comps, mw_map) — ``comps[i]`` is a
    {component: lb-mol/hr} dict for station ``i`` (empty when no composition /
    manual stream); ``mw_map`` is {component: MW} for converting moles → mass.
    """
    use_isoH = str(flash_mode).lower().startswith("isenth")
    injections = injections or {}
    stations: list[HE.Station] = []
    flashes: list = []
    comps: list[dict] = []
    p       = max(p_start, p_floor)
    cum     = 0.0
    last_bore = None
    flow_scale = 1.0          # running Tee split/merge multiplier (both phases)
    vap_cf = 1.0              # running vapour carry-forward multiplier
    liq_cf = 1.0              # running liquid carry-forward multiplier
    cur_feed = None           # ACTIVE mixture feed (a blend after a Tee join)
    cur_sp:  HE.StreamProps | None = None   # active mixture StreamProps
    last_key = None           # (stream, hmb, case) to detect a genuine new stream
    mw_map: dict[str, float] = {}   # component → MW (for mole→mass in the sheet)
    override_state: dict = {}       # persisted yellow-cell overrides (this stream)

    def _flash(fd, pp):
        return (FV.flash_isenthalpic(fd, pp) if use_isoH else FV.flash(fd, pp))

    for ridx, row in enumerate(block_rows):
        row_sp, row_feed = resolver(row)
        key = (row.get("Stream Lookup"), row.get("HMB File"), row.get("Case"))
        # First row, or a genuinely NEW stream → (re)set the active mixture.
        if last_key is None or key != last_key:
            cur_feed = row_feed
            cur_sp   = row_sp if row_sp is not None else default_sp
            last_key = key
            flow_scale = vap_cf = liq_cf = 1.0
            override_state = {}
        elif cur_feed is None and row_sp is not None:
            cur_sp = row_sp                       # manual rows refresh sp
        override_state.update(_row_yellow_overrides(row))
        a_feed, a_sp = cur_feed, (cur_sp or default_sp)

        fitting    = row.get("Fitting Name")
        spec       = row.get("Piping Spec")
        line_bore_before = last_bore        # upstream line bore (for CV β ratio)
        bore       = row.get("Bore (in)") or last_bore
        # A Control Valve's Bore is its own port/seat size, not the line bore —
        # don't let it overwrite the running line bore used downstream.
        if bore and not _is_control_valve(fitting):
            last_bore = bore
        length_ft  = num(row.get("Length (ft)")) or 0.0
        dz_ft      = num(row.get("Elev Change (ft)")) or 0.0
        instr_type = row.get("Instr Type")
        instr_tag  = row.get("Instr Tag")
        instr_dp   = num(row.get("Instr dP (psi)")) or 0.0
        fixed_dp   = num(row.get("Fixed dP (psi)")) or 0.0
        fixed_k    = num(row.get("Fixed K"))
        set_p      = num(row.get("Set P (psia)"))
        exch_dp    = num(row.get("Exch Max Allow dP (psi)"))

        idr = HE.resolve_id(spec, bore)
        p_eval = max(p, p_floor)
        scale_tag = ""

        # ── Tee split / merge: update running flow scale ─────────────────
        # negative = SPLIT: a delta fraction of the CURRENT flow leaves the line.
        # positive = MERGE: the running flow becomes that fraction of the
        # original Main flow (an ABSOLUTE target), e.g. 0.125 → 0.25 → 0.5 → 1.0
        # builds a header up in stages as branches join it.
        frac = num(row.get("Flow Fraction from Main"))
        if frac is not None and _is_tee(fitting):
            new_scale = frac if frac > 0 else max(0.0, flow_scale + frac)
            verb = "merge" if frac > 0 else "split"
            scale_tag += f" | Tee {verb} {frac:+.3f} → flow×{new_scale:.3f}"
            flow_scale = new_scale

        # ── Mass Vapor/Liquid carry-over (phase MASS fraction, compounding) ──
        # f scales the phase's MASS (and its moles equally, MW unchanged), so
        # f=0 dumps all of that phase's mass; f=1 carries it all forward.
        v_cf = num(row.get("Mass Vapor Fraction Carry Over"))
        l_cf = num(row.get("Mass Liquid Fraction Carry Over"))
        if v_cf is not None:
            vap_cf *= max(0.0, v_cf)
        if l_cf is not None:
            liq_cf *= max(0.0, l_cf)
        if v_cf is not None or l_cf is not None:
            scale_tag += f" | carry-over mass V×{vap_cf:.3f} L×{liq_cf:.3f}"

        # ── Branch→Main join: blend the mixtures and RE-FLASH a new feed ──
        if ridx in injections:
            inj = injections[ridx]
            b_feed, b_comp = inj.get("feed"), inj.get("comp", {})
            b_mass = inj.get("vap_mass", 0.0) + inj.get("liq_mass", 0.0)
            if a_feed is not None and b_feed is not None and b_comp:
                fr0 = _flash(a_feed, p_eval)
                main_moles = _feed_comp_flows(a_feed, fr0,
                                              flow_scale * vap_cf, flow_scale * liq_cf)
                blended = _blend_feeds(a_feed, main_moles, b_feed, b_comp)
                if blended is not None:
                    mwmap = {n: a_feed.mw[i] for i, n in enumerate(a_feed.names)}
                    main_mass = sum(m * mwmap.get(n, 50.0)
                                    for n, m in main_moles.items())
                    cur_feed = blended
                    cur_sp   = _blend_sp(a_sp, inj.get("sp"), main_mass, b_mass)
                    a_feed, a_sp = cur_feed, cur_sp
                    flow_scale = vap_cf = liq_cf = 1.0   # blend = new full flow
                    scale_tag += (f" | Tee join {inj.get('label','Branch')} "
                                  f"→ re-flash blended feed (+{b_mass:.1f} lb/hr)")
            else:
                scale_tag += (f" | Tee join {inj.get('label','Branch')} "
                              f"(skipped — needs HMB streams on both runs)")

        # ── Flash the ACTIVE mixture at this pressure, then scale ────────
        vap_eff = flow_scale * vap_cf
        liq_eff = flow_scale * liq_cf
        comp_vec: dict = {}
        if a_feed is not None:
            for i, n in enumerate(a_feed.names):
                if n not in mw_map and i < len(a_feed.mw):
                    mw_map[n] = a_feed.mw[i]
            fr_full = _flash(a_feed, p_eval)
            fr = _scale_flash_result(fr_full, vap_eff, liq_eff)
            comp_vec = _feed_comp_flows(a_feed, fr_full, vap_eff, liq_eff)
        elif a_sp is not None and (a_sp.vap_mass or a_sp.liq_mass):
            fr = _scale_flash_result(_synth_flash_result(a_sp, p_eval),
                                     vap_eff, liq_eff)        # manual mode
        else:
            fr = None
        sp = a_sp                       # used by the ΔP dispatch below
        if override_state:
            sp, fr = _apply_property_overrides(sp, fr, override_state)
            scale_tag += (" | manual override: "
                          + ", ".join(sorted(override_state)))

        if vap_eff <= 1e-9 and liq_eff <= 1e-9:
            scale_tag += ("  | stream terminated (carry-forward 0)"
                          if (v_cf == 0 or l_cf == 0) else
                          "  | flow≈0 (fully split off)")

        phase_str = (_phase_of(fr.quality) if fr else (sp.phase if sp else ""))
        flash_tag = (f" (flash β={fr.beta:.3f})" if fr else "") + scale_tag

        # ── Dispatch by special fitting type ────────────────────────────
        if _is_fix_pressure(fitting):
            dp_t  = fixed_dp
            dp_f  = dp_k = dp_z = 0.0
            res   = dict(rho=0.0, v=0.0, Re=0.0, f=0.0, dp_f=0.0, dp_k=0.0, dp_z=0.0)
            note  = f"Fix Pressure: −{fixed_dp:.3f} psi{flash_tag}"

        elif _is_orifice(fitting, instr_type):
            dp_t  = instr_dp
            dp_f  = dp_k = dp_z = 0.0
            res   = dict(rho=0.0, v=0.0, Re=0.0, f=0.0, dp_f=0.0, dp_k=0.0, dp_z=0.0)
            note  = (f"Orifice {instr_tag or ''}:  −{instr_dp:.3f} psi"
                     f"{flash_tag}")

        elif _is_source(fitting):
            # Boundary: (re)anchor the running pressure to the Source Set P.
            tgt   = set_p if set_p is not None else p
            dp_t  = p - tgt          # negative ⇒ pressure rises to the source
            dp_f  = dp_k = dp_z = 0.0
            res   = dict(rho=0.0, v=0.0, Re=0.0, f=0.0, dp_f=0.0, dp_k=0.0, dp_z=0.0)
            note  = (f"Source: anchor P = {tgt:.3f} psia{flash_tag}"
                     if set_p is not None else
                     f"Source (no Set P — using running P){flash_tag}")

        elif _is_destination(fitting):
            # Boundary: observation point.  Pressure unchanged; report arrival
            # against the target Set P (does not clamp).
            dp_t  = 0.0
            dp_f  = dp_k = dp_z = 0.0
            res   = dict(rho=0.0, v=0.0, Re=0.0, f=0.0, dp_f=0.0, dp_k=0.0, dp_z=0.0)
            if set_p is not None:
                margin = p - set_p
                note  = (f"Destination: target {set_p:.3f} psia, arrival "
                         f"{p:.3f} psia (margin {margin:+.3f} psi){flash_tag}")
            else:
                note  = f"Destination (no target Set P){flash_tag}"

        elif _is_control_valve(fitting):
            ctype  = _cv_control_type(row)
            dest_p = (_downstream_dest_p(block_rows, ridx)
                      if ctype == "F" else None)
            # Base ΔP: flow control floats to the downstream Destination P;
            # otherwise a fixed design ΔP (Fixed dP, else Darcy with Fixed K).
            if dest_p is not None:
                ov = (cv_overrides or {}).get(ridx)
                if ov is not None:
                    base_dp = ov
                    src_tag = (f"flow control → ΔP solved for arrival at "
                               f"Destination {dest_p:.1f} psia (incl. downstream "
                               f"losses)")
                else:
                    base_dp = max(0.0, p - dest_p)
                    src_tag = (f"flow control → ΔP floats to Destination "
                               f"{dest_p:.1f} psia")
            elif fixed_dp:
                base_dp = fixed_dp
                src_tag = f"fixed design ΔP = {fixed_dp:.3f} psi"
            elif fixed_k is not None and idr.id_in and idr.id_in > 0 and fr is not None:
                d     = idr.id_in * IN_TO_M
                area  = math.pi / 4.0 * d * d
                g_obj = HE.Geom(d, area, 0.0, 0.0, fixed_k, ROUGH_STEEL_M / d)
                rk    = _dp_flashed(g_obj, sp, fr)
                base_dp = (rk["dp_f"] + rk["dp_k"] + rk["dp_z"]) * PA_TO_PSI
                src_tag = f"K = {fixed_k} resistance"
            else:
                base_dp = 0.0
                src_tag = "no ΔP basis (set Control Valve Type / Fixed dP / Fixed K)"
            # Temperature control: ΔP floored at the exchanger max-allowable ΔP.
            floor_tag = ""
            if ctype == "T" and exch_dp is not None and exch_dp > base_dp:
                floor_tag = (f"; T-control floor → exchanger max-allow "
                             f"ΔP {exch_dp:.3f} psi")
                base_dp = exch_dp
            # β ratio = valve port bore / upstream line bore
            v_bore = num(row.get("Bore (in)"))
            beta_tag = ""
            if v_bore and line_bore_before and line_bore_before > 0:
                beta_tag = f"; β = {v_bore / line_bore_before:.3f}"
            dp_t  = base_dp
            dp_f  = dp_k = dp_z = 0.0
            res   = dict(rho=0.0, v=0.0, Re=0.0, f=0.0, dp_f=0.0, dp_k=0.0, dp_z=0.0)
            note  = (f"Control Valve [{ctype}]: −{dp_t:.3f} psi "
                     f"({src_tag}{floor_tag}{beta_tag}){flash_tag}")

        else:
            # Standard Darcy-Weisbach  (includes Fix K)
            if idr.id_in is None or idr.id_in <= 0:
                res  = dict(rho=0.0, v=0.0, Re=0.0, f=0.0, dp_f=0.0, dp_k=0.0, dp_z=0.0)
                note = idr.basis
            else:
                d    = idr.id_in * IN_TO_M
                area = math.pi / 4.0 * d * d
                L    = length_ft * FT_TO_M
                dz   = dz_ft * FT_TO_M
                # Fix K overrides the table K; all other terms stay
                k_val = (fixed_k if (_is_fix_k(fitting) and fixed_k is not None)
                         else fitting_k(fitting))
                g_obj = HE.Geom(d, area, L, dz, k_val, ROUGH_STEEL_M / d)
                res   = (_dp_flashed(g_obj, sp, fr) if fr is not None
                         else dict(rho=0.0, v=0.0, Re=0.0, f=0.0,
                                   dp_f=0.0, dp_k=0.0, dp_z=0.0))
                k_tag = f" | K={k_val}" if _is_fix_k(fitting) else ""
                note  = idr.basis + k_tag + flash_tag

            dp_f = res["dp_f"] * PA_TO_PSI
            dp_k = res["dp_k"] * PA_TO_PSI
            dp_z = res["dp_z"] * PA_TO_PSI
            dp_t = dp_f + dp_k + dp_z

        # Choke guard
        choked = (p - dp_t) < p_floor
        if choked:
            dp_t = max(0.0, p - p_floor)
            note += "  | CHOKED (Δp capped)"

        p_out = p - dp_t
        cum  += length_ft

        stations.append(HE.Station(
            seq          = int(num(row.get("Seq")) or 0),
            comp_id      = txt(row.get("Comp ID")),
            fitting      = fitting,
            spec         = spec,
            nps          = idr.nps,
            schedule     = idr.schedule,
            id_in        = idr.id_in,
            od_in        = idr.od_in,
            wall_in      = idr.wall_in,
            length_ft    = round(length_ft, 3),
            dz_ft        = round(dz_ft, 3),
            cum_ft       = round(cum, 3),
            p_in_psia    = round(p, 4),
            p_out_psia   = round(p_out, 4),
            v_fts        = round(res["v"] / FT_TO_M, 3) if res.get("v") else 0.0,
            Re           = round(res["Re"], 0)  if res.get("Re") else 0.0,
            f            = round(res["f"], 5)   if res.get("f") else 0.0,
            rho_lbft3    = round(res["rho"] / LBFT3_TO_KGM3, 5) if res.get("rho") else 0.0,
            dp_fric_psi  = round(dp_f, 5),
            dp_fit_psi   = round(dp_k, 5),
            dp_elev_psi  = round(dp_z, 5),
            dp_total_psi = round(dp_t, 5),
            note         = note,
        ))
        flashes.append(fr)
        comps.append(comp_vec)
        p = p_out

    return stations, flashes, comps, mw_map


def build_profile_solved(
        block_rows: list[dict],
        resolver,
        p_start: float,
        p_floor: float = 0.05,
        flash_mode: str = "isothermal",
        default_sp: HE.StreamProps | None = None,
        injections: dict | None = None,
        max_pass: int = 24,
        tol: float = 0.01,
) -> tuple[list[HE.Station], list, list, dict]:
    """March a block, solving flow-control Control Valves to their Destination.

    A flow-control valve must absorb only the *residual* ΔP so that, AFTER the
    downstream line/fitting losses between the valve and its Destination, the
    pressure arrives at the target.  Those downstream losses depend (via
    density / flash) on the valve's outlet pressure, so we iterate: march,
    recompute each valve's required ΔP from the realised downstream losses,
    re-march, until the ΔP settles (< ``tol`` psi) or ``max_pass`` is reached.

    With no flow-control valve that has a downstream Destination this is a
    single pass — identical to ``build_profile_flash_noiso`` directly."""
    # Pre-scan two-pass targets: valve row index → (Destination index, dest P).
    targets: dict[int, tuple[int, float]] = {}
    for ri, row in enumerate(block_rows):
        if (_is_control_valve(row.get("Fitting Name"))
                and _cv_control_type(row) == "F"):
            di = next((j for j in range(ri + 1, len(block_rows))
                       if _is_destination(block_rows[j].get("Fitting Name"))
                       and num(block_rows[j].get("Set P (psia)")) is not None),
                      None)
            if di is not None:
                targets[ri] = (di, num(block_rows[di].get("Set P (psia)")))

    common = dict(p_floor=p_floor, flash_mode=flash_mode,
                  default_sp=default_sp, injections=injections)
    if not targets:
        return build_profile_flash_noiso(block_rows, resolver, p_start, **common)

    # Pass 0: valves drop straight to their Destination P (no override) — this
    # over-drops by the downstream losses, giving the lower bracket on arrival.
    result = build_profile_flash_noiso(block_rows, resolver, p_start, **common)
    stations = result[0]

    # The arrival pressure at a Destination is monotonically DECREASING in its
    # valve's ΔP (more drop → lower outlet AND higher downstream loss).  So
    # bisect each valve's ΔP between 0 (fully open) and p_in − dest_p (drop
    # straight to target) — robust even when the compressible downstream loss
    # is strongly pressure-dependent (plain substitution oscillates there).
    overrides: dict[int, float] = {}
    brackets: dict[int, list[float]] = {}
    for vri, (di, dest_p) in targets.items():
        hi = max(0.0, stations[vri].p_in_psia - dest_p)
        brackets[vri] = [0.0, hi]
        overrides[vri] = hi / 2.0

    for _ in range(max_pass):
        result = build_profile_flash_noiso(block_rows, resolver, p_start,
                                           cv_overrides=overrides, **common)
        stations = result[0]
        converged = True
        for vri, (di, dest_p) in targets.items():
            arrival = stations[di].p_out_psia      # Destination ΔP = 0 ⇒ arrival
            lo, hi = brackets[vri]
            if abs(arrival - dest_p) > tol and (hi - lo) > tol:
                converged = False
                if arrival > dest_p:               # under-dropped → need more ΔP
                    lo = overrides[vri]
                else:                              # over-dropped → need less ΔP
                    hi = overrides[vri]
                brackets[vri] = [lo, hi]
                overrides[vri] = 0.5 * (lo + hi)
        if converged:
            break
    return result


# ════════════════════════════════════════════════════════════════════════
#  Pressure_Profile sheet  (no-ISO version — same layout as flash engine)
# ════════════════════════════════════════════════════════════════════════
def _pp_headers(u: "UN.UnitSystem") -> list[str]:
    return [
        "Seq", "Comp ID", "Fitting", "NPS", "Sched", u.hdr("ID", "Lin"),
        u.hdr("Length", "L"), u.hdr("Cum Length", "L"), u.hdr("Elev Δ", "L"),
        "Phase (flash)", u.hdr("Velocity", "v"), "Reynolds", "Friction f",
        u.hdr("Mix Density", "rho"), u.hdr("dP Fric", "dP"),
        u.hdr("dP Fitting", "dP"), u.hdr("dP Elev", "dP"),
        u.hdr("dP Total", "dP"), u.hdr("P In", "P"), u.hdr("P Out", "P"),
        "β (molar vap)", "Quality (mass)", u.hdr("Vap Mass Flow", "mflow"),
        u.hdr("Liq Mass Flow", "mflow"), "MW Vapor", "MW Liquid",
        u.hdr("Vap Density", "rho"), u.hdr("Liq Density", "rho"),
        "GVF (local, vol)", "Note",
    ]


_PP_NCOL = 30


def _build_pressure_profile_sheet(
        wb: Workbook,
        stations: list[HE.Station],
        flashes: list,
        sp: HE.StreamProps,
        circuit_id: str,
        stream_name: str,
        run_label: str = "Main",
        flash_mode: str = "isothermal",
        station_lines: list[str] | None = None,   # parallel list: line_no per station
        line_colors: dict[str, str] | None = None, # line_no -> palette color
) -> str:
    """Create a Pressure_Profile sheet.

    When ``station_lines`` is supplied (multi-line circuit) the sheet inserts a
    bold divider row at every line transition and colour-codes each data row with
    the per-line palette colour instead of the fitting-type palette.
    """
    sheet_title = f"Pressure_Profile_{run_label}" if run_label != "Main" else "Pressure_Profile"
    ws = wb.create_sheet(sheet_title)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = STEEL

    multi = bool(station_lines and line_colors and len(set(station_lines)) > 1)
    line_nos_str = (", ".join(dict.fromkeys(station_lines))
                   if station_lines else circuit_id)

    u = _u()
    ncol = _PP_NCOL
    ws.merge_cells(f"A1:{get_column_letter(ncol)}1")
    _c(ws, 1, 1,
       f"FLASH-COUPLED PRESSURE PROFILE  [{run_label}]   |   "
       f"Circuit {circuit_id}  ·  Lines: {line_nos_str}   |   "
       f"Stream: {stream_name}   |   {flash_mode}   |   units: {u.system}   |   "
       f"{datetime.now():%d-%b-%Y %H:%M}",
       bg=NAVY, fg=WHITE, sz=10, bold=True)
    ws.row_dimensions[1].height = 22

    _hdr(ws, _pp_headers(u), row=2)
    for ci in range(21, ncol + 1):
        ws.cell(2, ci).fill = PatternFill("solid", fgColor=GRNHDR)
    ws.freeze_panes = "C3"
    ws.auto_filter.ref = f"A2:{get_column_letter(ncol)}2"

    rho_l_lbft3  = sp.liq_density
    data_row      = 3          # current Excel row for writing
    prev_line     = None
    chart_data_rows: list[tuple[int, float, float]] = []   # (excel_row, cum_ft, p_out)

    for i, (s, fr) in enumerate(zip(stations, flashes)):
        cur_line = (station_lines[i] if station_lines else None)

        # ── Line-transition divider ──────────────────────────────────────
        if multi and cur_line and cur_line != prev_line:
            ws.merge_cells(f"A{data_row}:{get_column_letter(ncol)}{data_row}")
            lc = (line_colors or {}).get(cur_line, LGRAY)
            _c(ws, data_row, 1,
               f"  ▶  LINE: {cur_line}",
               bg=lc, fg=NAVY, sz=9, bold=True)
            ws.row_dimensions[data_row].height = 14
            data_row += 1
            prev_line = cur_line

        # ── Row background ───────────────────────────────────────────────
        if multi and cur_line and line_colors:
            row_bg = line_colors.get(cur_line, LGRAY)
        else:
            row_bg = fitting_bg(s.fitting)

        vap_rho = (_gas_rho_kgm3(fr.vap_mw, sp.vap_z, sp.temp_f, s.p_in_psia)
                   / LBFT3_TO_KGM3) if fr else None
        q_gas  = (fr.vap_mass / vap_rho) if (fr and vap_rho) else None
        q_liq  = (fr.liq_mass / rho_l_lbft3) if (fr and rho_l_lbft3) else None
        q_tot  = (q_gas or 0.0) + (q_liq or 0.0)
        gvf    = (q_gas / q_tot) if (q_gas and q_tot) else None

        vals = [
            s.seq, s.comp_id, s.fitting, s.nps, s.schedule,
            u.disp("Lin", s.id_in),
            u.disp("L", s.length_ft), u.disp("L", s.cum_ft, 3),
            u.disp("L", s.dz_ft),
            _phase_of(fr.quality) if fr else (sp.phase or ""),
            u.disp("v", s.v_fts), s.Re, s.f, u.disp("rho", s.rho_lbft3, 5),
            u.disp("dP", s.dp_fric_psi, 5), u.disp("dP", s.dp_fit_psi, 5),
            u.disp("dP", s.dp_elev_psi, 5), u.disp("dP", s.dp_total_psi, 5),
            u.disp("P", s.p_in_psia), u.disp("P", s.p_out_psia),
            round(fr.beta,    4) if fr else None,
            round(fr.quality, 4) if fr else None,
            u.disp("mflow", fr.vap_mass, 1) if fr else None,
            u.disp("mflow", fr.liq_mass, 1) if fr else None,
            round(fr.vap_mw, 2)   if fr else None,
            round(fr.liq_mw, 2)   if fr else None,
            u.disp("rho", vap_rho, 6)     if vap_rho else None,
            u.disp("rho", rho_l_lbft3, 4) if rho_l_lbft3 else None,
            round(gvf, 4)         if gvf is not None else None,
            s.note,
        ]
        for ci, v in enumerate(vals, 1):
            _c(ws, data_row, ci, "" if v is None else v, bg=row_bg, sz=9,
               ha="right" if isinstance(v, (int, float)) else "left")
        _c(ws, data_row, 20, u.disp("P", s.p_out_psia), bg=AMBER, sz=9,
           bold=True, ha="right")

        chart_data_rows.append((data_row, s.cum_ft, s.p_out_psia))
        data_row += 1

    widths = ([5, 10, 28, 6, 7, 9, 11, 14, 10, 16, 13, 12, 10, 16,
               12, 13, 12, 13, 12, 12]
              + [12, 12, 18, 18, 10, 10, 16, 16, 14, 36])
    for ci, w in enumerate(widths, 1):
        cw(ws, ci, w)

    sr = data_row + 1
    if stations:
        tot_dp = u.disp("dP", stations[0].p_in_psia - stations[-1].p_out_psia, 4)
        tot_len = u.disp("L", stations[-1].cum_ft, 2)
        _c(ws, sr, 1, "TOTAL", bg=GRNHDR, fg=WHITE, bold=True)
        _c(ws, sr, 2, f"ΔP = {tot_dp} {u.label('dP')} over {tot_len} {u.label('L')} "
                      f"({len(stations)} components  ·  {len(set(station_lines or []))} lines)",
           bg=GREEN, bold=True)

    if len(chart_data_rows) >= 2:
        chart = LineChart()
        chart.title = f"Flash Pressure Profile — Circuit {circuit_id} [{run_label}]"
        chart.x_axis.title = f"Cumulative Length ({u.label('L')})"
        chart.y_axis.title = f"Pressure ({u.label('P')})"
        chart.height, chart.width = 9, 24
        # Build the chart from column 20 (P Out) across only data rows
        first_dr = chart_data_rows[0][0]
        last_dr  = chart_data_rows[-1][0]
        data  = Reference(ws, min_col=20, min_row=first_dr - 1, max_row=last_dr)
        cats  = Reference(ws, min_col=8,  min_row=first_dr,     max_row=last_dr)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart, f"A{sr + 2}")

    return sheet_title


# ════════════════════════════════════════════════════════════════════════
#  Flash_Profile detail sheet
# ════════════════════════════════════════════════════════════════════════
def _build_flash_detail_sheet(
        wb: Workbook,
        stations: list[HE.Station],
        flashes: list,
        feed,
        sp: HE.StreamProps,
        run_label: str = "Main",
) -> str:
    sheet_title = f"Flash_Profile_{run_label}" if run_label != "Main" else "Flash_Profile"
    ws = wb.create_sheet(sheet_title)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = PURPLE
    ws.merge_cells("A1:L1")
    _c(ws, 1, 1, f"RIGOROUS FLASH PROFILE [{run_label}]  —  isothermal VLE vs pressure",
       bg=NAVY, fg=WHITE, sz=11, bold=True)
    ws.row_dimensions[1].height = 20

    if feed is None:
        _c(ws, 3, 1, "No composition available — flash disabled.", sz=9, fg=ORANGE)
        return sheet_title

    u = _u()
    info = [
        ("Feed components",           feed.n,                      ""),
        ("Total molar rate",
         u.disp("molflow", feed.total_moles_lbmolhr, 2), u.label("molflow")),
        ("Total mass rate",
         u.disp("mflow", feed.total_mass_lbhr, 1), u.label("mflow")),
        ("Total MW",                  round(feed.total_mw, 3),             ""),
        ("Reference pressure",        u.disp("P", feed.p_ref_psia),  u.label("P")),
        ("Reference temperature",
         u.disp("T", sp.temp_f, 2), f"{u.label('T')} (isothermal)"),
        ("Reference molar vap frac β₀", round(feed.beta_ref, 4),    "anchored to HMB"),
    ]
    r = 3
    for k, v, unit_note in info:
        _c(ws, r, 1, k, bg=LGRAY, sz=9, bold=True)
        _c(ws, r, 2, v, sz=9, ha="right")
        _c(ws, r, 3, unit_note, sz=9, fg="808080")
        r += 1
    r += 1

    hdr = ["Seq", "Comp ID", u.hdr("P In", "P"), "β (molar)", "Quality (mass)",
           u.hdr("Vap Moles", "molflow"), u.hdr("Liq Moles", "molflow"),
           u.hdr("Vap Mass", "mflow"), u.hdr("Liq Mass", "mflow"),
           "MW Vapor", "MW Liquid"]
    _hdr(ws, hdr, row=r)
    head = r
    ws.freeze_panes = f"A{r + 1}"
    for i, (s, fr) in enumerate(zip(stations, flashes)):
        rr = head + 1 + i
        bg = LGRAY if i % 2 else WHITE
        vals = [s.seq, s.comp_id, u.disp("P", s.p_in_psia),
                round(fr.beta, 4), round(fr.quality, 4),
                u.disp("molflow", fr.vap_moles, 3),
                u.disp("molflow", fr.liq_moles, 3),
                u.disp("mflow", fr.vap_mass, 1),
                u.disp("mflow", fr.liq_mass, 1),
                round(fr.vap_mw, 2), round(fr.liq_mw, 2)] if fr else [
            s.seq, s.comp_id, u.disp("P", s.p_in_psia), None, None,
            None, None, None, None, None, None]
        for ci, v in enumerate(vals, 1):
            _c(ws, rr, ci, "" if v is None else v, bg=bg, sz=9,
               ha="right" if isinstance(v, (int, float)) else "left")
    for ci, w in enumerate([5, 9, 12, 11, 13, 20, 20, 16, 16, 10, 10], 1):
        cw(ws, ci, w)
    return sheet_title


# ════════════════════════════════════════════════════════════════════════
#  Composition Splits sheet
# ════════════════════════════════════════════════════════════════════════
def _build_composition_splits_sheet(
        wb: Workbook,
        stations: list,
        comps: list[dict],
        mw_map: dict,
        circuit_id: str,
        stream_name: str,
        station_lines: list[str] | None = None,
        line_colors: dict[str, str] | None = None,
) -> str:
    """Per-component split across the marching sequence (columns left→right).

    Four stacked blocks — mass flow (lb/hr), molar flow (lb-mol/hr), mass
    fraction (mass %) and mole fraction (mol %).  Component moles come straight
    from ``comps`` and are converted to mass with ``mw_map``.  Built only when at
    least one station carries composition; else a note is left.
    """
    ws = wb.create_sheet("Composition Splits")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = PURPLE

    ncol = 1 + len(stations)
    last_col = get_column_letter(max(ncol, 2))
    ws.merge_cells(f"A1:{last_col}1")
    _c(ws, 1, 1,
       f"COMPOSITION SPLITS (mass & molar)   |   Circuit {circuit_id}   |   "
       f"Stream {stream_name}   |   {datetime.now():%d-%b-%Y %H:%M}",
       bg=NAVY, fg=WHITE, sz=11, bold=True)
    ws.row_dimensions[1].height = 22

    u = _u()
    # Convert each station's molar vector → mass (lb/hr) using mw_map.
    mass_cols: list[dict] = []
    for cv in comps:
        mc = {}
        for nm, mol in cv.items():
            mw = mw_map.get(nm)
            if mol and mw:
                mc[nm] = mol * mw
        mass_cols.append(mc)

    # Ordered union of component names (first appearance order)
    names: list[str] = []
    seen: set = set()
    for mc in mass_cols:
        for nm in mc:
            if nm not in seen:
                seen.add(nm)
                names.append(nm)

    if not names:
        _c(ws, 3, 1,
           "No composition available — Stream Lookup blank (manual stream) "
           "or no component data in the HMB.", sz=9, fg=ORANGE)
        cw(ws, 1, 60)
        return "Composition Splits"

    def _station_header(r0: int):
        _c(ws, r0, 1, "Component", bg=NAVY, fg=WHITE, sz=9, bold=True, ha="left")
        for j, s in enumerate(stations):
            col = 2 + j
            ln = station_lines[j] if (station_lines and j < len(station_lines)) else None
            bg = (line_colors or {}).get(ln, STEEL) if line_colors else STEEL
            _c(ws, r0, col, f"{s.seq} · {s.comp_id or ''}".strip(),
               bg=bg, fg=NAVY, sz=8, bold=True, ha="center", wrap=True)
        ws.row_dimensions[r0].height = 26

    col_tot  = [sum(mc.values()) for mc in mass_cols]   # mass totals per station
    mole_tot = [sum(cv.values()) for cv in comps]       # molar totals per station

    # ── Block 1: mass flow ───────────────────────────────────────────────
    r = 3
    ws.merge_cells(f"A{r}:{last_col}{r}")
    _c(ws, r, 1, f"MASS FLOW  ({u.label('mflow')})",
       bg=GRNHDR, fg=WHITE, sz=9, bold=True)
    r += 1
    _station_header(r)
    r += 1
    for nm in names:
        _c(ws, r, 1, nm, bg=LGRAY, sz=8, ha="left")
        for j, mc in enumerate(mass_cols):
            val = mc.get(nm)
            _c(ws, r, 2 + j, u.disp("mflow", val, 2) if val else "",
               sz=8, ha="right")
        r += 1
    _c(ws, r, 1, f"TOTAL ({u.label('mflow')})", bg=GREEN, sz=8, bold=True)
    for j, tot in enumerate(col_tot):
        _c(ws, r, 2 + j, u.disp("mflow", tot, 2) if tot else "", bg=GREEN, sz=8,
           bold=True, ha="right")
    r += 2

    # ── Block 2: molar flow ──────────────────────────────────────────────
    ws.merge_cells(f"A{r}:{last_col}{r}")
    _c(ws, r, 1, f"MOLAR FLOW  ({u.label('molflow')})",
       bg=GRNHDR, fg=WHITE, sz=9, bold=True)
    r += 1
    _station_header(r)
    r += 1
    for nm in names:
        _c(ws, r, 1, nm, bg=LGRAY, sz=8, ha="left")
        for j, cv in enumerate(comps):
            val = cv.get(nm)
            _c(ws, r, 2 + j, u.disp("molflow", val, 4) if val else "",
               sz=8, ha="right")
        r += 1
    _c(ws, r, 1, f"TOTAL ({u.label('molflow')})", bg=GREEN, sz=8, bold=True)
    for j, tot in enumerate(mole_tot):
        _c(ws, r, 2 + j, u.disp("molflow", tot, 4) if tot else "", bg=GREEN, sz=8,
           bold=True, ha="right")
    r += 2

    # ── Block 3: mass fraction (mass %) ──────────────────────────────────
    ws.merge_cells(f"A{r}:{last_col}{r}")
    _c(ws, r, 1, "MASS FRACTION  (mass %)", bg=GRNHDR, fg=WHITE, sz=9, bold=True)
    r += 1
    _station_header(r)
    r += 1
    for nm in names:
        _c(ws, r, 1, nm, bg=LGRAY, sz=8, ha="left")
        for j, mc in enumerate(mass_cols):
            tot = col_tot[j]
            val = mc.get(nm)
            pct = (100.0 * val / tot) if (val and tot) else None
            _c(ws, r, 2 + j, round(pct, 3) if pct else "", sz=8, ha="right")
        r += 1
    _c(ws, r, 1, "TOTAL (mass %)", bg=GREEN, sz=8, bold=True)
    for j in range(len(mass_cols)):
        _c(ws, r, 2 + j, 100.0 if col_tot[j] else "", bg=GREEN, sz=8,
           bold=True, ha="right")
    r += 2

    # ── Block 4: mole fraction (mol %) ───────────────────────────────────
    ws.merge_cells(f"A{r}:{last_col}{r}")
    _c(ws, r, 1, "MOLE FRACTION  (mol %)", bg=GRNHDR, fg=WHITE, sz=9, bold=True)
    r += 1
    _station_header(r)
    r += 1
    for nm in names:
        _c(ws, r, 1, nm, bg=LGRAY, sz=8, ha="left")
        for j, cv in enumerate(comps):
            tot = mole_tot[j]
            val = cv.get(nm)
            pct = (100.0 * val / tot) if (val and tot) else None
            _c(ws, r, 2 + j, round(pct, 3) if pct else "", sz=8, ha="right")
        r += 1
    _c(ws, r, 1, "TOTAL (mol %)", bg=GREEN, sz=8, bold=True)
    for j in range(len(comps)):
        _c(ws, r, 2 + j, 100.0 if mole_tot[j] else "", bg=GREEN, sz=8,
           bold=True, ha="right")

    # Widths + freeze
    cw(ws, 1, 26)
    for j in range(len(stations)):
        cw(ws, 2 + j, 12)
    ws.freeze_panes = "B4"
    return "Composition Splits"


# ════════════════════════════════════════════════════════════════════════
#  Input_Pipeline sheet  — styled identically to PCF_Parsed_updated.xlsx
# ════════════════════════════════════════════════════════════════════════
# Column layout mirrors PCF_Parsed_updated.xlsx; no-ISO extras appended at end.
_PCF_HEADERS = [
    "Seq",              # 1
    "Comp ID",          # 2
    "Run Type",         # 3   — color-coded green(Main) / blue(Branch)
    "Fitting Name",     # 4   — fitting-type palette
    "Tag",              # 5   — instrument tag (= Instr Tag for instruments)
    "Bore In (in)",     # 6
    "Bore Out (in)",    # 7
    "Piping Spec",      # 8
    "Length (ft)",      # 9
    "Elev Change (ft)", # 10  — green if +, red if −
    "Direction",        # 11  — UP=green, DOWN=red
    "Start P (psia)",   # 12  — filled only on first row of each block
    # no-ISO specific extras
    "Fixed K",          # 13
    "Fixed dP (psi)",   # 14
    "Instr Type",       # 15
    "Instr Tag",        # 16
    "Instr dP (psi)",   # 17
    "Notes",            # 18
    # ── stream / case provenance + flow split ───────────────────────────
    "Stream",           # 19
    "HMB File",         # 20
    "Case",             # 21
    "Flow Frac",        # 22  — signed Tee split(−)/merge(+) fraction of Main
    "Mass Vap CF",      # 23  — Mass Vapor Fraction Carry Over
    "Mass Liq CF",      # 24  — Mass Liquid Fraction Carry Over
]
_PCF_NCOL = len(_PCF_HEADERS)


def _build_input_sheet(
        wb: Workbook,
        circuit_rows: list[dict],
        circuit_id: str,
        stream_name: str = "",
        start_p: float | None = None,
        flash_mode: str = "isothermal",
        line_colors: dict[str, str] | None = None,
        hmb_meta: dict | None = None,
) -> str:
    """PCF-style pipeline component sheet for a whole circuit (all lines)."""
    ws = wb.create_sheet("Input_Pipeline")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = STEEL
    last_col = get_column_letter(_PCF_NCOL)

    line_nos = list(dict.fromkeys(r["Line No"] for r in circuit_rows))

    # ── Row 1: title bar ─────────────────────────────────────────────────
    ws.merge_cells(f"A1:{last_col}1")
    title = (f"CIRCUIT:  {circuit_id}   |   ISO: (manual — no ISO)   |   "
             f"Lines: {', '.join(line_nos)}   |   Stream: {stream_name or '—'}")
    _c(ws, 1, 1, title, bg=NAVY, fg=WHITE, sz=10, bold=True)
    ws.row_dimensions[1].height = 20

    # ── Row 2: subtitle ──────────────────────────────────────────────────
    ws.merge_cells(f"A2:{last_col}2")
    n_main   = sum(1 for r in circuit_rows if r.get("Run Type", "Main") == "Main")
    n_branch = sum(1 for r in circuit_rows if r.get("Run Type") == "Branch")
    subtitle = (f"Start P: {_u().disp('P', start_p)} {_u().label('P')}   |   "
                f"Flash mode: {flash_mode}   |   "
                f"Components: {n_main} Main  /  {n_branch} Branch  |  "
                f"{len(line_nos)} line(s)")
    _c(ws, 2, 1, subtitle, bg=STEEL, fg=WHITE, sz=9)
    ws.row_dimensions[2].height = 16

    # ── Row 3: HMB provenance (parsed from the HMB filename) — Input only ──
    hdr_row = 3
    if hmb_meta and hmb_meta.get("file"):
        ws.merge_cells(f"A3:{last_col}3")
        meta = (f"HMB: {hmb_meta.get('file','')}"
                f"   |   Unit: {hmb_meta.get('unit','') or '—'}"
                f"   |   Case: {hmb_meta.get('case_no','') or '—'}"
                f"   |   Desc: {hmb_meta.get('description','') or '—'}"
                f"   |   Issued: {hmb_meta.get('date','') or '—'}")
        _c(ws, 3, 1, meta, bg=AMBER, fg=DGRAY, sz=9, bold=True)
        ws.row_dimensions[3].height = 16
        hdr_row = 4

    # Redefine line_rows to use all circuit rows
    line_rows = circuit_rows

    # ── Column headers (Navy, white, bold; unit-aware) ────────────────────
    _uh = _u()
    _pcf_hdr_units = {
        "Length (ft)": _uh.hdr("Length", "L"),
        "Elev Change (ft)": _uh.hdr("Elev Change", "L"),
        "Start P (psia)": _uh.hdr("Start P", "P"),
        "Fixed dP (psi)": _uh.hdr("Fixed dP", "dP"),
        "Instr dP (psi)": _uh.hdr("Instr dP", "dP"),
    }
    for ci, h in enumerate(_PCF_HEADERS, 1):
        _c(ws, hdr_row, ci, _pcf_hdr_units.get(h, h), bg=NAVY, fg=WHITE, sz=9,
           bold=True, ha="center", wrap=True)
    ws.row_dimensions[hdr_row].height = 30
    ws.freeze_panes = f"A{hdr_row + 1}"
    ws.auto_filter.ref = f"A{hdr_row}:{last_col}{hdr_row}"

    # ── Data rows ────────────────────────────────────────────────────────
    rn = hdr_row + 1
    prev_ln = None
    for row in line_rows:
        fitting = row.get("Fitting Name")
        rt      = row.get("Run Type", "Main")
        elev_ch = row.get("Elev Change (ft)")
        dirn    = row.get("Direction")
        cur_ln  = row.get("Line No")

        # Insert a line-transition banner when the line number changes
        if cur_ln and cur_ln != prev_ln:
            ws.merge_cells(f"A{rn}:{get_column_letter(_PCF_NCOL)}{rn}")
            lc = (line_colors or {}).get(cur_ln, LGRAY) if line_colors else LGRAY
            _c(ws, rn, 1, f"  ▶  LINE: {cur_ln}", bg=lc, fg=NAVY, sz=9, bold=True)
            ws.row_dimensions[rn].height = 13
            rn += 1
            prev_ln = cur_ln

        # Base row colour: fitting-type palette (PCF_Parsed_updated.xlsx convention)
        row_bg = fitting_bg(fitting)

        u = _u()
        vals = [
            row.get("Seq"),
            row.get("Comp ID"),
            rt,
            fitting,
            row.get("Instr Tag"),          # Tag (instrument tag if present)
            row.get("Bore (in)"),          # Bore In
            row.get("Bore (in)"),          # Bore Out (same for non-reducers)
            row.get("Piping Spec"),
            u.disp("L", row.get("Length (ft)")),
            u.disp("L", elev_ch),
            dirn,
            u.disp("P", row.get("Start P (psia)")),
            row.get("Fixed K"),
            u.disp("dP", row.get("Fixed dP (psi)"), 5),
            row.get("Instr Type"),
            row.get("Instr Tag"),
            u.disp("dP", row.get("Instr dP (psi)"), 5),
            row.get("Notes"),
            row.get("Stream Lookup"),
            row.get("HMB File"),
            row.get("Case"),
            row.get("Flow Fraction from Main"),
            row.get("Mass Vapor Fraction Carry Over"),
            row.get("Mass Liquid Fraction Carry Over"),
        ]

        for ci, v in enumerate(vals, 1):
            cell_bg = row_bg
            if ci == 3:
                cell_bg = _CLR_FLANGE if rt == "Main" else _CLR_BRANCH
            elif ci == 10 and v is not None:
                try:
                    cell_bg = _CLR_UP if float(v) > 0 else (_CLR_DOWN if float(v) < 0 else row_bg)
                except (TypeError, ValueError):
                    pass
            elif ci == 11:
                d_bg = direction_bg(v)
                if d_bg:
                    cell_bg = d_bg
            _c(ws, rn, ci, "" if v is None else v, bg=cell_bg, sz=9,
               ha="right" if isinstance(v, (int, float)) else "left")
        rn += 1

    # ── Column widths ─────────────────────────────────────────────────────
    col_widths = [5, 12, 9, 28, 12, 10, 10, 12, 11, 14, 10, 12, 10, 13, 11, 12, 14, 30,
                  14, 16, 10, 10, 9, 9]
    for ci, w in enumerate(col_widths, 1):
        cw(ws, ci, w)

    return "Input_Pipeline"


# ════════════════════════════════════════════════════════════════════════
#  Orchestration — one workbook per CIRCUIT
# ════════════════════════════════════════════════════════════════════════
def _make_stream_resolver(default_hmb: str, default_case: str, map_path):
    """Return a memoised ``resolver(row) -> (StreamProps|None, feed|None)``.

    Resolution order per row:
      1. ``Stream Lookup`` set  → load from (row HMB File | CLI default,
         row Case | CLI default).  Composition feed loaded for real flash.
      2. ``Stream Lookup`` blank → build manual StreamProps from yellow cells
         (no feed → frozen synthetic flash downstream).
    """
    cache: dict[tuple, tuple] = {}

    def resolver(row):
        stream = row.get("Stream Lookup")
        if not stream:
            return _streamprops_from_manual(row), None

        hmb  = row.get("HMB File") or default_hmb
        case = row.get("Case") or default_case
        key  = (hmb, case, stream)
        if key in cache:
            return cache[key]

        sp = HE.load_stream_props(hmb, stream, case)
        feed = None
        if sp is not None:
            try:
                is_case = HE.is_proii_export(hmb)
                feed = FV.read_feed(hmb, stream, case, sp, is_casesheet=is_case)
            except Exception as exc:
                print(f"  (feed load failed for {stream}@{case}: {exc})")
        cache[key] = (sp, feed)
        return cache[key]

    return resolver


def run_noiso(
        input_path: str,
        hmb_path: str,
        circuit_id: str | None = None,   # process one circuit; None = all
        stream_name: str | None = None,
        out_path: str | None = None,
        case: str = "Case 1",
        map_path: str | None = None,
        flash_mode: str = "isothermal",
) -> list[tuple[str, list[HE.Station], list, HE.StreamProps]]:
    """
    Process hydraulic circuits from ``input_path``.

    Each circuit (col A) is a series chain of lines.  All Main-run rows across
    all lines are marched sequentially as ONE continuous pressure profile.
    Branch rows per line are marched independently, each from its own Start P.

    Returns list of (out_xlsx, main_stations, main_flashes, sp) — one per circuit.
    """
    # ── Display-units context: read the UNITS sheet of the input workbook ──
    try:
        _uwb = load_workbook(input_path, read_only=True, data_only=True)
        set_units(UN.UnitSystem.from_workbook(_uwb))
        _uwb.close()
    except Exception:
        set_units(UN.UnitSystem("FPS"))
    print(f"  Units: {_u().system}"
          + ("" if _u().is_fps else "  (engine internals FPS; I/O converted)"))

    all_rows = read_pipeline_input(input_path)
    if not all_rows:
        raise SystemExit(f"No data rows found in {input_path}. "
                         "Run with --template to create the blank workbook.")

    circuits = _group_circuits(all_rows)
    target_ids = [circuit_id] if circuit_id else list(circuits.keys())
    results = []

    for cid in target_ids:
        if cid not in circuits:
            print(f"  WARNING: circuit '{cid}' not found — skipped.")
            continue

        c   = circuits[cid]
        sk  = stream_name or c["stream_lookup"]
        p0  = c["start_p"]
        lns = c["line_order"]

        # ── Per-row stream resolver (HMB File / Case / manual) ────────────
        def_hmb  = c["hmb_file"] or hmb_path
        def_case = c["case"] or case
        resolver = _make_stream_resolver(def_hmb, def_case, map_path)

        # ── Circuit default StreamProps (drives the screening sheets) ─────
        feed = None
        if sk:
            sp = HE.load_stream_props(def_hmb, sk, def_case)
            if sp is None:
                try:
                    map_row = SMAP.resolve_from_line(lns[0] if lns else "", "", map_path)
                    if map_row:
                        sk = map_row.lookup_key
                        sp = HE.load_stream_props(def_hmb, sk, def_case)
                except Exception:
                    pass
            if sp is None:
                print(f"  WARNING: stream '{sk}' not found in {def_hmb} "
                      f"(case '{def_case}') — circuit '{cid}' skipped.")
                continue
            is_case = HE.is_proii_export(def_hmb)
            feed = FV.read_feed(def_hmb, sk, def_case, sp, is_casesheet=is_case)
            if feed is None:
                print("  WARNING: no composition for default stream — flash uses "
                      "frozen split where applicable.")
            else:
                print(f"  Flash feed: {feed.n} components | β₀={feed.beta_ref:.4f} "
                      f"| P_ref={feed.p_ref_psia:.2f} psia | mode: {flash_mode}")
        else:
            # No HMB stream at circuit level → manual mode.  Use the first main
            # row carrying manual data as the representative StreamProps.
            sp = next((m for m in (_streamprops_from_manual(r)
                                   for r in c["main_rows"]) if m), None)
            if sp is None:
                print(f"  WARNING: circuit '{cid}' has no Stream Lookup and no "
                      f"manual properties — skipped.")
                continue
            sk = "(manual)"
            print("  Manual stream mode — properties from yellow input cells.")

        print(f"\n  Circuit {cid}  |  Lines: {', '.join(lns)}  |  "
              f"Stream {sk}  |  {sp.phase}  "
              f"T={_u().disp('T', sp.temp_f, 2)}{_u().label('T')}  "
              f"P={_u().disp('P', sp.pres_psia)} {_u().label('P')}")

        phase = HE.classify_phase(sp)
        lc    = _line_color_map(lns)

        def gdf(p, _sp=sp):
            return (HE.vapor_density_kgm3(_sp, p) / LBFT3_TO_KGM3
                    if _sp.mol_weight else (_sp.vap_density or 0.0))

        main_rows     = c["main_rows"]
        branch_blocks = c["branch_blocks"]
        p_start = p0 or sp.pres_psia or 100.0

        # ── March branch blocks FIRST (independent sub-runs) ─────────────
        # A branch whose outlet matches a Main 'Tee Join Branch Flow N' row on
        # the same line is INJECTED into the Main at that row (its outlet flow +
        # composition merge in); others stay standalone Branch sheets.
        branch_runs: list = []            # (label, sts, fls, cmps, b_lns)
        injections: dict[int, dict] = {}  # main_rows index → branch outlet
        line_branch_count: dict[str, int] = {}
        for bi, block in enumerate(branch_blocks):
            bp = _start_p_of_block(block) or p_start
            sts, fls, cmps, _bmw = build_profile_solved(
                block, resolver, bp, flash_mode=flash_mode, default_sp=sp)
            b_lns = [r["Line No"] for r in block]
            label = f"Branch_{bi+1}" if len(branch_blocks) > 1 else "Branch"
            branch_runs.append((label, sts, fls, cmps, b_lns))

            bline  = block[0].get("Line No") if block else None
            line_branch_count[bline] = line_branch_count.get(bline, 0) + 1
            join_n = line_branch_count[bline]
            join_idx = next(
                (mi for mi, mr in enumerate(main_rows)
                 if mr.get("Line No") == bline
                 and _is_tee_join(mr.get("Fitting Name")) == join_n),
                None)
            if join_idx is not None and fls and fls[-1] is not None:
                out_fr = fls[-1]
                b_sp, b_feed = resolver(block[0]) if block else (None, None)
                injections[join_idx] = {
                    "vap_mass": out_fr.vap_mass,
                    "liq_mass": out_fr.liq_mass,
                    "vmol":     out_fr.vap_moles,
                    "lmol":     out_fr.liq_moles,
                    "comp":     cmps[-1] if cmps else {},
                    "feed":     b_feed,
                    "sp":       b_sp,
                    "label":    label,
                }
                print(f"  Branch '{label}' (line {bline}) joins Main at "
                      f"Tee Join Branch Flow {join_n}.")

        # ── March the Main with branch injections ────────────────────────
        main_stations, main_flashes, main_comps, main_mw = build_profile_solved(
            main_rows, resolver, p_start, flash_mode=flash_mode,
            default_sp=sp, injections=injections)

        # station_lines: map each station back to its source Line No
        station_lines = [r["Line No"] for r in main_rows
                         if r["Run Type"] == "Main"]

        wb = Workbook()
        pp_sheets, fp_sheets = [], []

        pp_sheets.append(_build_pressure_profile_sheet(
            wb, main_stations, main_flashes, sp,
            circuit_id=cid, stream_name=sk,
            run_label="Main", flash_mode=flash_mode,
            station_lines=station_lines, line_colors=lc))
        fp_sheets.append(_build_flash_detail_sheet(
            wb, main_stations, main_flashes, feed, sp, run_label="Main"))

        # ── Composition Splits (component MASS split across the sequence) ──
        _build_composition_splits_sheet(
            wb, main_stations, main_comps, main_mw, cid, sk,
            station_lines=station_lines, line_colors=lc)

        # ── Branch sheets ────────────────────────────────────────────────
        for (label, sts, fls, cmps, b_lns) in branch_runs:
            pp_sheets.append(_build_pressure_profile_sheet(
                wb, sts, fls, sp,
                circuit_id=cid, stream_name=sk,
                run_label=label, flash_mode=flash_mode,
                station_lines=b_lns, line_colors=lc))
            fp_sheets.append(_build_flash_detail_sheet(
                wb, sts, fls, feed, sp, run_label=label))

        # ── Screening / analysis sheets (all main stations, whole circuit)
        circuit_label = f"Circuit {cid}  ({', '.join(lns)})"
        if main_stations:
            HE.build_detail_sheet(wb, main_stations, phase)
            HE.build_stream_sheet(wb, sp, phase)
            HE.build_fiv_sheet(wb, main_stations, sp,
                                circuit_label, sk, gas_density_fn=gdf,
                                station_lines=station_lines, line_colors=lc)
            HE.build_aiv_sheet(wb, main_stations, sp,
                                circuit_label, sk,
                                station_lines=station_lines, line_colors=lc)
            HE.build_regime_sheet(wb, main_stations, sp,
                                   circuit_label, sk, gas_density_fn=gdf)
        else:
            wb.create_sheet("Component_Detail")
            HE.build_stream_sheet(wb, sp, phase)

        fp_map_sheets = []
        try:
            FPM = HE                      # flow-pattern maps merged into this file
            fp_map_sheets = FPM.build_flow_pattern_maps(
                wb, main_stations, sp, circuit_label, sk, gas_density_fn=gdf)
        except Exception as exc:
            print(f"  (flow-pattern maps skipped: {exc})")

        # ── Input_Pipeline and README ─────────────────────────────────────
        circuit_rows = [r for r in all_rows if r["Circuit"] == cid]
        hmb_meta = _parse_hmb_filename(def_hmb)
        _build_input_sheet(wb, circuit_rows, cid,
                           stream_name=sk,
                           start_p=p_start,
                           flash_mode=flash_mode,
                           line_colors=lc,
                           hmb_meta=hmb_meta)
        _build_readme_noiso(wb, circuit_label, sk, sp, phase,
                            input_path, hmb_path, flash_mode)

        # ── Sheet order ──────────────────────────────────────────────────
        order = (pp_sheets + fp_sheets +
                 ["Composition Splits",
                  "Component_Detail", "Stream_Props",
                  "FIV_EI_T2.2", "AIV", "Two_Phase_Regime"]
                 + fp_map_sheets
                 + ["Input_Pipeline", "README"])
        for name in reversed(order):
            if name in [s.title for s in wb.worksheets]:
                wb.move_sheet(name, offset=-len(wb.worksheets))
        if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
            del wb["Sheet"]

        if out_path is None:
            safe_c = re.sub(r"[^A-Za-z0-9_-]", "_", str(cid))
            safe_s = re.sub(r"[^A-Za-z0-9_-]", "_", str(sk))
            op = f"hydraulics_flash_noiso_{safe_c}_{safe_s}.xlsx"
        else:
            op = out_path

        wb.save(op)
        results.append((op, main_stations, main_flashes, sp))

        if main_stations:
            u = _u()
            tot_dp = u.disp(
                "dP", main_stations[0].p_in_psia - main_stations[-1].p_out_psia, 4)
            print(f"  Saved : {op}")
            print(f"  Phase : {phase}   Lines: {len(lns)}   "
                  f"Components (Main): {len(main_stations)}")
            print(f"  Total ΔP : {tot_dp} {u.label('dP')}  "
                  f"({u.disp('P', main_stations[0].p_in_psia)} → "
                  f"{u.disp('P', main_stations[-1].p_out_psia)} {u.label('P')})")

    return results


# ════════════════════════════════════════════════════════════════════════
#  README sheet
# ════════════════════════════════════════════════════════════════════════
def _build_readme_noiso(wb, line_no, stream_name, sp, phase,
                         input_path, hmb_path, flash_mode):
    ws = wb.create_sheet("README")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = MGRAY
    cw(ws, 1, 120)
    ws["A1"] = f"HyCalign Hydraulics — NO-ISO Flash Mode — Line {line_no}"
    ws["A1"].font = Font(name="Calibri", bold=True, size=14, color=NAVY)

    lines = [
        "",
        f"Line:         {line_no}",
        f"Stream:       {stream_name}   (HMB: {os.path.basename(hmb_path)})",
        f"Input file:   {os.path.basename(input_path)}",
        f"Phase (HMB):  {sp.phase}  ->  calc path '{phase}'",
        f"Flash mode:   {flash_mode}",
        f"Units:        {_u().system}  (input & results; engine internals FPS)",
        f"Generated:    {datetime.now():%d-%b-%Y %H:%M}",
        "",
        "PIPELINE INPUT",
        "  Components entered manually in pipeline_input_noiso.xlsx.",
        "  Internal diameter resolved from Piping Spec + Bore via pms_classes.py / asme_data.py.",
        "  No ISOGEN .pcf required.",
        "",
        "SPECIAL FITTINGS",
        "  Fix K        — user-supplied K coefficient; Darcy-Weisbach friction + elevation still applied.",
        f"  Fix Pressure — user-supplied dP ({_u().label('dP')}) subtracted directly; all hydraulics skipped.",
        f"  Orifice / FO — user-supplied Instr dP ({_u().label('dP')}) subtracted directly; no velocity calculation.",
        "",
        "FLASH VLE",
        f"  Method: {flash_mode} flash at each station's local pressure.",
        "  Reference K-values anchored to HMB molar vapor fraction at P_ref.",
        "  Vapor mass flow and quality evolve as pressure falls (genuine flashing).",
        "  Falls back to frozen HMB split when no composition is available.",
        "",
        "FORMULAS",
        "  Friction:  Darcy f·(L/D)·½ρv²   (64/Re laminar; Swamee-Jain turbulent)",
        "  Fittings:  K·½ρv²   (K from Crane TP-410 table; Fix K uses user value)",
        "  Elevation: ρ·g·Δz",
        "  Roughness: 0.0018 in commercial steel.",
        "",
        "VERIFY: schedule/wall from asme_data.py — confirm before design use.",
    ]
    for i, t in enumerate(lines, 2):
        ws.cell(i, 1).value = t
        ws.cell(i, 1).font = Font(
            name="Calibri", size=9,
            bold=(t.isupper() and len(t) > 3),
            color=NAVY if (t.isupper() and len(t) > 3) else DGRAY,
        )


# ════════════════════════════════════════════════════════════════════════
#  CLI entry point
# ════════════════════════════════════════════════════════════════════════
def _extract_opt(argv: list[str], name: str) -> tuple[str | None, list[str]]:
    """Pull ``--name value`` or ``--name=value`` out of argv (returns value, rest)."""
    val: str | None = None
    rest: list[str] = []
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == name and i + 1 < len(argv):
            val = argv[i + 1]
            i += 2
            continue
        if a.startswith(name + "="):
            val = a.split("=", 1)[1]
            i += 1
            continue
        rest.append(a)
        i += 1
    return val, rest


def main():
    argv = sys.argv[1:]
    case, argv = _extract_opt(argv, "--case")      # e.g. --case "Case 2"
    case = case or "Case 1"

    flags = [a for a in argv if a.startswith("-")]
    args  = [a for a in argv if not a.startswith("-")]

    if "--template" in flags or "-t" in flags:
        unit_sys = "SI" if any(f.lower() in ("--si", "-si") for f in flags) else "FPS"
        out = create_input_template(unit_system=unit_sys)
        print(f"  Template created: {out}  (units: {unit_sys})")
        print("  Fill in the Pipeline_Input sheet then re-run without --template.")
        print("  Flip FPS/SI any time on the workbook's UNITS sheet.")
        return

    flash_mode = "isenthalpic" if any(
        f.lower() in ("--adiabatic", "--isenthalpic", "-h2", "--jt")
        for f in flags
    ) else "isothermal"

    input_path = args[0] if len(args) > 0 else "pipeline_input_noiso.xlsx"
    hmb_path   = args[1] if len(args) > 1 else None
    circuit_id = args[2] if len(args) > 2 else None   # optional: run only this circuit

    if not os.path.exists(input_path):
        print(f"  Input not found: {input_path}")
        print("  Run with --template to create the blank input workbook.")
        sys.exit(1)

    if hmb_path is None:
        for cand in ("HMB_from_proii.xlsx", "HMB.xlsx"):
            if os.path.exists(cand):
                hmb_path = cand
                break
    if not hmb_path or not os.path.exists(hmb_path):
        raise SystemExit(f"HMB not found: {hmb_path}")

    print(f"\n  No-ISO Flash Hydraulics  |  input: {os.path.basename(input_path)}"
          f"  |  HMB: {os.path.basename(hmb_path)}  |  case: {case}"
          f"  |  mode: {flash_mode}")

    results = run_noiso(input_path, hmb_path,
                        circuit_id=circuit_id, case=case, flash_mode=flash_mode)
    print(f"\n  Done — {len(results)} workbook(s) written.")




# ══════════════════════════════════════════════════════════════════════════
# ║  SECTION: flash VLE + flow-pattern maps (former flash_flowpatterns.py)
# ══════════════════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════════════════
# ║  SECTION: former flash_vle.py
# ══════════════════════════════════════════════════════════════════════════
#!/usr/bin/env python3
"""Rigorous flash (VLE) for the pressure-marched hydraulics engine.

The HMB carries everything a component flash needs:
  * per-component TOTAL / VAPOR / LIQUID mole fractions  ->  feed z_i and the
    *true* reference equilibrium ratio  K_ref,i = y_i / x_i  at the HMB (T, P_ref);
  * per-component MOLECULAR WEIGHT  MW_i = (mass rate_i) / (molar rate_i);
  * (optionally) a COMP_CONSTANTS sheet with per-component Tc / Pc / acentric
    (library values for pure components, Lee-Kesler + Edmister estimates for
    pseudo-fractions) used for the Wilson temperature correction.

Reference equilibrium ratios
----------------------------
When reliable vapor (y) and liquid (x) mole fractions are present, the true
ratio  K_ref,i = y_i / x_i  is used directly (exact at the HMB point).  Where a
stream's liquid basis is inconsistent (3-phase / free-water), the code falls
back to the robust  K_raw,i = z_i / x_i  proxy with a single multiplier λ solved
so Rachford-Rice reproduces the HMB's authoritative molar vapor fraction β₀.

Pressure / temperature dependence
---------------------------------
Isothermal (default):   K_i(P) = K_ref,i · (P_ref / P)        (Raoult / ideal-K)

Non-isothermal:         K_i(P,T) = K_ref,i · (P_ref / P) · W_i(T)
                        W_i(T) = exp[ 5.373·(1+ω_i)·Tc_i·(1/T_ref − 1/T) ]   (°R)
i.e. the Wilson temperature correction layered on the true reference K.  W_i = 1
at T = T_ref, so the isothermal result is recovered exactly.  Components without
Tc/ω get W_i = 1 (graceful).

Isenthalpic flash
-----------------
flash_isenthalpic() solves T and β jointly so total enthalpy is conserved across
a segment (Joule-Thomson / latent-heat cooling as the fluid flashes), using the
HMB phase specific-enthalpies and Cp.  Use it across valves / large ΔP.

This module is import-only side-effect free; the flash engine drives it.
"""


import math
import re
from dataclasses import dataclass, field

F_TO_R = 459.67   # °F + 459.67 = °R
WILSON_C = 5.373  # Wilson (1969) acentric coefficient


# ════════════════════════════════════════════════════════════════════════
#  Feed definition
# ════════════════════════════════════════════════════════════════════════
@dataclass
class FlashFeed:
    names: list[str]
    z: list[float]            # feed mole fractions (Σ = 1)
    k_ref: list[float]        # reference equilibrium ratios at p_ref (and t_ref)
    mw: list[float]           # component molecular weights (lb/lb-mol)
    total_mass_lbhr: float    # total mass flow (constant along the line)
    p_ref_psia: float         # reference (HMB) pressure
    temp_f: float | None = None
    total_mw: float = 0.0
    total_moles_lbmolhr: float = 0.0
    beta_ref: float = 0.0     # reference molar vapor fraction (sanity anchor)
    # Per-component criticals for the Wilson temperature correction (None if
    # unavailable for a given component → that component's W_i = 1).
    tc_r: list[float] | None = None     # true critical temperature, °R
    omega: list[float] | None = None    # acentric factor
    pc_psia: list[float] | None = None  # critical pressure, psia (Wilson basis)
    # Reference-state enthalpy data for the isenthalpic march (BTU/LB, BTU/LB-°F)
    vap_h_ref: float | None = None
    liq_h_ref: float | None = None
    vap_cp: float | None = None
    liq_cp: float | None = None
    k_basis: str = "yx"       # 'yx' (true), 'zx-lambda' (proxy), or 'wilson'

    @property
    def n(self) -> int:
        return len(self.z)

    def has_constants(self) -> bool:
        return bool(self.tc_r) and any(t is not None for t in self.tc_r)


@dataclass
class FlashResult:
    p_psia: float
    beta: float               # molar vapor fraction
    quality: float            # mass vapor fraction (= vap_mass / total)
    vap_moles: float
    liq_moles: float
    vap_mass: float           # lb/hr
    liq_mass: float           # lb/hr
    vap_mw: float
    liq_mw: float
    temp_f: float | None = None    # station temperature (= feed T if isothermal)
    y: list[float] = field(default_factory=list)
    x: list[float] = field(default_factory=list)
    k: list[float] = field(default_factory=list)


# ════════════════════════════════════════════════════════════════════════
#  Rachford-Rice + flash
# ════════════════════════════════════════════════════════════════════════
def rachford_rice(z, k, tol: float = 1e-12, it_max: int = 200) -> float:
    """Molar vapor fraction beta in [0, 1].  g(beta) is monotone-decreasing."""
    def g(b):
        s = 0.0
        for zi, ki in zip(z, k):
            s += zi * (ki - 1.0) / (1.0 + b * (ki - 1.0))
        return s

    if g(0.0) <= 0.0:
        return 0.0            # subcooled liquid (no vapor)
    if g(1.0) >= 0.0:
        return 1.0            # superheated vapor (no liquid)
    lo, hi = 0.0, 1.0
    for _ in range(it_max):
        mid = 0.5 * (lo + hi)
        gm = g(mid)
        if abs(gm) < tol:
            return mid
        if gm > 0.0:
            lo = mid
        else:
            hi = mid
    return 0.5 * (lo + hi)


def _wilson_tcorr(feed: FlashFeed, temp_f: float) -> list[float]:
    """Per-component Wilson temperature multiplier W_i(T) relative to T_ref.

        W_i = exp[ 5.373·(1+ω_i)·Tc_i·(1/T_ref − 1/T) ]      (T in °R)

    W_i = 1 where Tc_i or ω_i is unavailable, or when T == T_ref.
    """
    n = feed.n
    if not feed.has_constants() or feed.temp_f is None:
        return [1.0] * n
    t_r     = temp_f + F_TO_R
    t_ref_r = feed.temp_f + F_TO_R
    if t_r <= 0 or t_ref_r <= 0 or abs(t_r - t_ref_r) < 1e-9:
        return [1.0] * n
    inv_diff = (1.0 / t_ref_r) - (1.0 / t_r)
    out = []
    tc_list = feed.tc_r or [None] * n
    om_list = feed.omega or [None] * n
    for tc, om in zip(tc_list, om_list):
        if tc is None or om is None or tc <= 0:
            out.append(1.0)
        else:
            expo = WILSON_C * (1.0 + om) * tc * inv_diff
            expo = max(-50.0, min(50.0, expo))   # overflow guard
            out.append(math.exp(expo))
    return out


def flash(feed: FlashFeed, p_psia: float,
          temp_f: float | None = None) -> FlashResult:
    """Flash at local pressure p_psia.

    Isothermal (temp_f is None or == feed.temp_f):
        K_i(P) = K_ref,i · (P_ref / P)                       (ideal-K)
    Non-isothermal (temp_f given and differs, constants present):
        K_i(P,T) = K_ref,i · (P_ref / P) · W_i(T)            (Wilson T-correction)
    """
    p = max(p_psia, 1e-6)
    ratio = feed.p_ref_psia / p
    t_used = temp_f if temp_f is not None else feed.temp_f
    wcorr = _wilson_tcorr(feed, t_used) if (t_used is not None) else [1.0] * feed.n

    k = [min(1e12, max(1e-12, kr * ratio * w))
         for kr, w in zip(feed.k_ref, wcorr)]
    beta = rachford_rice(feed.z, k)

    x, y = [], []
    for zi, ki in zip(feed.z, k):
        xi = zi / (1.0 + beta * (ki - 1.0))
        x.append(xi)
        y.append(ki * xi)
    sx = sum(x) or 1.0
    sy = sum(y) or 1.0
    x = [xi / sx for xi in x]
    y = [yi / sy for yi in y]

    vap_mw = sum(yi * mi for yi, mi in zip(y, feed.mw))
    liq_mw = sum(xi * mi for xi, mi in zip(x, feed.mw))
    F = feed.total_moles_lbmolhr
    vap_moles = beta * F
    liq_moles = (1.0 - beta) * F
    vap_mass = vap_moles * vap_mw
    liq_mass = liq_moles * liq_mw
    tot = vap_mass + liq_mass
    quality = vap_mass / tot if tot > 0 else 0.0
    return FlashResult(p_psia=p, beta=beta, quality=quality,
                       vap_moles=vap_moles, liq_moles=liq_moles,
                       vap_mass=vap_mass, liq_mass=liq_mass,
                       vap_mw=vap_mw, liq_mw=liq_mw,
                       temp_f=t_used, y=y, x=x, k=k)


def _feed_enthalpy_ref(feed: FlashFeed) -> float | None:
    """Reference total specific enthalpy (BTU/LB) at (T_ref, P_ref) from the HMB
    phase enthalpies weighted by the reference mass split."""
    if feed.vap_h_ref is None and feed.liq_h_ref is None:
        return None
    # mass-weighted using reference quality
    fr0 = flash(feed, feed.p_ref_psia)            # reproduces HMB split
    q0  = fr0.quality
    hv  = feed.vap_h_ref if feed.vap_h_ref is not None else feed.liq_h_ref
    hl  = feed.liq_h_ref if feed.liq_h_ref is not None else feed.vap_h_ref
    return q0 * hv + (1.0 - q0) * hl


def _total_enthalpy_at(feed: FlashFeed, fr: FlashResult, temp_f: float) -> float:
    """Total specific enthalpy (BTU/LB) of the flashed mixture at temp_f,
    Cp-linearised from the HMB reference state."""
    dT = temp_f - (feed.temp_f if feed.temp_f is not None else temp_f)
    hv0 = feed.vap_h_ref if feed.vap_h_ref is not None else (feed.liq_h_ref or 0.0)
    hl0 = feed.liq_h_ref if feed.liq_h_ref is not None else (feed.vap_h_ref or 0.0)
    cpv = feed.vap_cp or 0.0
    cpl = feed.liq_cp or 0.0
    hv = hv0 + cpv * dT
    hl = hl0 + cpl * dT
    return fr.quality * hv + (1.0 - fr.quality) * hl


def flash_isenthalpic(feed: FlashFeed, p_psia: float,
                      h_target: float | None = None,
                      t_lo: float | None = None, t_hi: float | None = None,
                      it_max: int = 60) -> FlashResult:
    """Adiabatic (isenthalpic) flash: solve T and β jointly so the mixture
    enthalpy equals h_target (default = HMB reference enthalpy).  Captures
    Joule-Thomson / latent-heat cooling as the fluid flashes across a large ΔP.

    Falls back to an isothermal flash if no enthalpy data is available.
    """
    if feed.vap_h_ref is None and feed.liq_h_ref is None:
        return flash(feed, p_psia)                # no enthalpy data → isothermal

    if h_target is None:
        h_target = _feed_enthalpy_ref(feed)
        if h_target is None:
            return flash(feed, p_psia)

    t_ref = feed.temp_f if feed.temp_f is not None else 100.0
    lo = t_lo if t_lo is not None else t_ref - 200.0
    hi = t_hi if t_hi is not None else t_ref + 50.0

    def resid(t):
        fr = flash(feed, p_psia, temp_f=t)
        return _total_enthalpy_at(feed, fr, t) - h_target, fr

    r_lo, _ = resid(lo)
    r_hi, _ = resid(hi)
    # If not bracketed, expand once, else return best-effort isothermal
    if r_lo * r_hi > 0:
        fr = flash(feed, p_psia, temp_f=t_ref)
        return fr

    fr = None
    for _ in range(it_max):
        mid = 0.5 * (lo + hi)
        rm, fr = resid(mid)
        if abs(rm) < 1e-4:
            return fr
        if r_lo * rm < 0:
            hi = mid; r_hi = rm
        else:
            lo = mid; r_lo = rm
    return fr if fr is not None else flash(feed, p_psia, temp_f=0.5 * (lo + hi))


# ════════════════════════════════════════════════════════════════════════
#  Composition readers  (transposed Case-sheet HMB  and  td-dump per-stream)
# ════════════════════════════════════════════════════════════════════════
def _ref_beta(sp) -> float:
    """Reference molar vapor fraction from the HMB phase mass flows + phase MWs."""
    vm, lm = (sp.vap_mass or 0.0), (sp.liq_mass or 0.0)
    vmw = sp.vap_mw or sp.mol_weight
    lmw = sp.liq_mw or sp.mol_weight
    vmol = (vm / vmw) if vmw else 0.0
    lmol = (lm / lmw) if lmw else 0.0
    tot = vmol + lmol
    return (vmol / tot) if tot > 0 else 0.0


def _solve_lambda(z, k_raw, beta_target, it_max: int = 200) -> float:
    """Multiplier λ such that Rachford-Rice(z, λ·k_raw) == beta_target.

    β increases monotonically with λ, so a geometric bisection converges.  This
    anchors the reference flash to the HMB's authoritative molar vapor fraction
    while preserving the relative-volatility ordering carried by k_raw.
    """
    def beta_of(lam):
        return rachford_rice(z, [min(1e12, max(1e-12, lam * k)) for k in k_raw])

    lo, hi = 1e-9, 1e9
    if beta_of(lo) >= beta_target:
        return lo
    if beta_of(hi) <= beta_target:
        return hi
    for _ in range(it_max):
        mid = math.sqrt(lo * hi)
        b = beta_of(mid)
        if abs(b - beta_target) < 1e-9:
            return mid
        if b < beta_target:
            lo = mid
        else:
            hi = mid
    return math.sqrt(lo * hi)


_CONST_CACHE: dict[str, dict] = {}

def _read_comp_constants(hmb_path) -> dict:
    """Read the COMP_CONSTANTS sheet embedded in the HMB (if present) →
    {NAME_UPPER: {'tc_f','pc_psia','omega','mw'}}.  Cached per file.

    Falls back to comp_constants.load + Constants.xlsx auto-locate if the sheet
    is absent.  Returns {} if nothing is available (flash stays isothermal-y/x).
    """
    key = str(hmb_path)
    if key in _CONST_CACHE:
        return _CONST_CACHE[key]

    out: dict = {}
    # 1) embedded COMP_CONSTANTS sheet
    try:
        from openpyxl import load_workbook
        wb = load_workbook(key, read_only=True, data_only=True)
        if "COMP_CONSTANTS" in wb.sheetnames:
            ws = wb["COMP_CONSTANTS"]
            # headers on row 2; data row 3+. cols: B=2 Name, E=5 MW,
            # I=9 Tc(F), J=10 Pc(psia), M=13 omega
            for row in ws.iter_rows(min_row=3, values_only=True):
                if not row or len(row) < 13 or row[1] is None:
                    continue
                name = str(row[1]).strip()
                if not name or name.upper() in ("PURE", "PSEUDO", "ESTIMATED"):
                    continue
                if name.upper().startswith("TOTAL:") or "  |  " in name:
                    continue          # summary / legend rows
                def _f(v):
                    try:    return float(v)
                    except (TypeError, ValueError): return None
                out[name.upper()] = {
                    "mw":      _f(row[4]),    # E
                    "tc_f":    _f(row[8]),    # I
                    "pc_psia": _f(row[9]),    # J
                    "omega":   _f(row[12]),   # M
                }
        wb.close()
    except Exception:
        pass

    # 2) fallback: comp_constants + Constants.xlsx
    if not out:
        try:
            import comp_constants as _CC
            from pathlib import Path
            near = Path(key).resolve().parent
            cpath = None
            for d in (near, Path.cwd()):
                hits = sorted(d.glob("Constants*.xlsx")) + sorted(d.glob("*onstants*.xlsx"))
                if hits:
                    cpath = hits[0]; break
            if cpath is not None:
                for c in _CC.read_constants(str(cpath)):
                    out[c["name"].upper()] = {
                        "mw":      c.get("MW"),
                        "tc_f":    c.get("Tc_F"),
                        "pc_psia": c.get("Pc_psia"),
                        "omega":   c.get("omega"),
                    }
        except Exception:
            pass

    _CONST_CACHE[key] = out
    return out


def _build_feed(z_d, x_d, mass_d, mole_d, sp,
                y_d: dict | None = None,
                const_map: dict | None = None) -> FlashFeed | None:
    """Assemble a FlashFeed from per-component dicts + the stream's StreamProps.

    Reference equilibrium ratios:
      * if reliable vapor (y) and liquid (x) mole fractions are present, the
        true ratio K_ref,i = y_i / x_i is used (exact at the HMB point);
      * otherwise the robust proxy K_raw,i = z_i / x_i is used with a single
        multiplier λ solved so Rachford-Rice reproduces the HMB molar vapor
        fraction β₀ at P_ref.

    Per-component Tc / Pc / ω (from const_map, keyed by upper-case name) are
    attached for the Wilson temperature correction in non-isothermal marches.
    """
    beta0 = _ref_beta(sp)
    beta0 = min(max(beta0, 1e-4), 1.0 - 1e-4)     # keep numerically usable
    y_d = y_d or {}

    # Decide whether the y/x basis is usable: need a meaningful number of
    # components with both y and x present and a non-degenerate spread.
    y_ok = 0
    for n, zi in z_d.items():
        if (zi or 0) > 1e-12 and (y_d.get(n) or 0) > 0 and (x_d.get(n) or 0) > 1e-12:
            y_ok += 1
    use_yx = y_ok >= 2

    names, z, k_raw, mw = [], [], [], []
    tc_r, omega, pc_l = [], [], []
    for n, zi in z_d.items():
        if zi is None or zi <= 1e-12:
            continue
        xi = x_d.get(n, 0.0) or 0.0
        yi = y_d.get(n, 0.0) or 0.0
        mi = mass_d.get(n)
        mo = mole_d.get(n)
        mwi = (mi / mo) if (mi and mo and mo > 0) else None
        names.append(n)
        z.append(float(zi))
        if use_yx and yi > 0 and xi > 1e-12:
            k_raw.append(yi / xi)               # true equilibrium ratio
        else:
            k_raw.append((float(zi) / xi) if xi > 1e-12 else 1e6)
        mw.append(mwi)
        # criticals
        cm = (const_map or {}).get(str(n).strip().upper())
        if cm and cm.get("tc_f") is not None:
            tc_r.append(cm["tc_f"] + F_TO_R)
            omega.append(cm.get("omega"))
            pc_l.append(cm.get("pc_psia"))
        else:
            tc_r.append(None); omega.append(None); pc_l.append(None)

    if len(z) < 2:
        return None
    known = [m for m in mw if m]
    avg = (sum(known) / len(known)) if known else (sp.mol_weight or 50.0)
    mw = [m if m else avg for m in mw]
    s = sum(z) or 1.0
    z = [zi / s for zi in z]

    if use_yx:
        # true K_ref already material-balance correct; gentle λ nudge to β₀
        lam = _solve_lambda(z, k_raw, beta0)
        basis = "yx"
    else:
        lam = _solve_lambda(z, k_raw, beta0)
        basis = "zx-lambda"
    k_ref = [min(1e10, max(1e-10, lam * kr)) for kr in k_raw]

    total_mw = sum(zi * mi for zi, mi in zip(z, mw))
    total_mass = (sp.vap_mass or 0.0) + (sp.liq_mass or 0.0)
    if total_mass <= 0:
        total_mass = sp.total_mass or 0.0
    total_moles = (total_mass / total_mw) if total_mw > 0 else 0.0
    p_ref = sp.pres_psia or 0.0
    if p_ref <= 0:
        return None

    feed = FlashFeed(names=names, z=z, k_ref=k_ref, mw=mw,
                     total_mass_lbhr=total_mass, p_ref_psia=p_ref,
                     temp_f=sp.temp_f, total_mw=total_mw,
                     total_moles_lbmolhr=total_moles,
                     tc_r=tc_r, omega=omega, pc_psia=pc_l,
                     vap_h_ref=getattr(sp, "vap_sp_enthalpy", None),
                     liq_h_ref=getattr(sp, "liq_sp_enthalpy", None),
                     vap_cp=getattr(sp, "vap_cp", None),
                     liq_cp=getattr(sp, "liq_cp", None),
                     k_basis=basis)
    feed.beta_ref = rachford_rice(z, k_ref)
    return feed


def _from_casesheet(hmb_path, stream, case, sp) -> FlashFeed | None:
    """Read compositions from a transposed PRO/II export via td_parser."""
    import td_parser
    _streams, comp_data, names = td_parser.parse_proii(hmb_path, case)
    cd = comp_data.get(stream)
    if cd is None:                       # case-insensitive fallback
        up = str(stream).upper()
        for nm in names:
            if nm.upper() == up:
                cd = comp_data.get(nm)
                break
    if not cd:
        return None
    const_map = _read_comp_constants(hmb_path)
    return _build_feed(cd.get("MOLE_FRAC", {}), cd.get("LIQ_MOLE_FRAC", {}),
                       cd.get("MASS_FLOW", {}), cd.get("MOLE_FLOW", {}), sp,
                       y_d=cd.get("VAP_MOLE_FRAC", {}), const_map=const_map)


_SECT = [
    ("MOLE FLOW", "mole"), ("MASS FLOW", "mass"),
    ("VAPOR MOLE FRACTIONS", "y"), ("LIQUID MOLE FRACTIONS", "x"),
    ("MOLE FRACTIONS", "z"),          # total mole fractions (checked last)
]


def _section_key(label_up: str) -> str | None:
    if "VAPOR MOLE FRACTIONS" in label_up:
        return "y"
    if "LIQUID MOLE FRACTIONS" in label_up:
        return "x"
    if "MOLE FLOW" in label_up:
        return "mole"
    if "MASS FLOW" in label_up:
        return "mass"
    if "MOLE FRACTIONS" in label_up:    # total
        return "z"
    return None


def _from_td_dump(hmb_path, stream, sp) -> FlashFeed | None:
    """Read compositions from a td_parser per-stream sheet (Component|Value)."""
    from openpyxl import load_workbook
    wb = load_workbook(hmb_path, read_only=True, data_only=True)
    target = None
    up = str(stream).upper()
    for ws in wb.worksheets:
        v = ws["B1"].value
        m = re.search(r"STREAM:\s*([^|]+)", str(v) if v is not None else "")
        if m and m.group(1).strip().upper() == up:
            target = ws
            break
        if ws.title.upper() == up:
            target = ws
    if target is None:
        wb.close()
        return None

    buckets: dict[str, dict[str, float]] = {}
    cur = None
    for row in target.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 2 or row[1] is None:
            continue
        label = str(row[1]).strip()
        if re.match(r"^\d+\.\s", label):          # section header
            cur = _section_key(label.upper())
            continue
        if cur is None or label.lower() == "component":
            continue
        val = row[3] if len(row) > 3 else None     # column D = Value
        if isinstance(val, (int, float)):
            buckets.setdefault(cur, {})[label] = float(val)
    wb.close()
    if "z" not in buckets:
        return None
    const_map = _read_comp_constants(hmb_path)
    return _build_feed(buckets.get("z", {}), buckets.get("x", {}),
                       buckets.get("mass", {}), buckets.get("mole", {}), sp,
                       y_d=buckets.get("y", {}), const_map=const_map)


def read_feed(hmb_path, stream, case, sp, is_casesheet: bool) -> FlashFeed | None:
    """Build a FlashFeed for `stream` from whichever HMB layout is supplied."""
    try:
        if is_casesheet:
            return _from_casesheet(hmb_path, stream, case, sp)
        return _from_td_dump(hmb_path, stream, sp)
    except Exception:
        return None


# ══════════════════════════════════════════════════════════════════════════
# ║  SECTION: former flow_pattern_data.py
# ══════════════════════════════════════════════════════════════════════════
"""Flow-pattern map boundary polylines, extracted verbatim from
TECHNIP 'Slug flow.xlsm' Data sheet. Each value is a list of polylines;
each polyline is a list of (x, y) points in the map's own coordinates."""

BOUNDARIES = {
    'GF_V': [[(0.12, 0.01), (0.8, 20.0)], [(2.5, 0.01), (2.5, 0.22)], [(2.5, 0.22), (8.0, 1.2)], [(8.0, 1.2), (8.0, 20.0)], [(8.0, 1.2), (28.0, 0.01)], [(0.01, 20.0), (10000.0, 20.0)]],
    'SHELL_V': [[(10.0, 0.9), (3.0, 0.26), (1.8, 0.01)], [(3.0, 2.0), (2.0, 1.0), (0.9, 0.01)], [(0.11, 0.01), (0.08, 2.0), (0.045, 16.0)], [(0.09, 0.01), (0.04, 3.8), (0.02, 13.0)], [(2.0, 4.0), (0.6, 6.0), (0.3, 10.0)], [(0.045, 16.0), (0.02, 13.0), (0.004, 2.0)], [(0.3, 10.0), (0.2, 12.0), (0.045, 16.0)]],
    'SHELL2_V': [[(0.004, 1e-05), (0.004, 0.04)], [(0.004, 0.04), (0.2, 6.0)], [(0.2, 1e-05), (0.2, 1.5)], [(0.2, 6.0), (0.8, 6.0)], [(0.2, 1.5), (9.0, 60.0)], [(1.0, 1e-05), (40.0, 0.0005)]],
    'HTRI_V': [[(0.18, 1000.0), (3.0, 10.0)], [(0.55, 1000.0), (19.0, 10.0)], [(3.0, 1000.0), (100.0, 22.0)]],
    'HEDH_V': [[(1.0, 140.0), (10.0, 110.0), (100000.0, 100.0)], [(6000.0, 10.0), (6000.0, 100.0)], [(6000.0, 10.0), (10000.0, 2.0), (30000.0, 1.0)], [(10.0, 0.3), (100.0, 21.877616239495527), (6000.0, 70.0)], [(1000.0, 100.0), (1000.0, 1000.0)], [(1000.0, 1000.0), (2000.0, 4168.693834703358), (6000.0, 10000.0)]],
    'HEDH_DN_V': [[(0.2, 0.0), (1.8, 2.0)], [(0.4, 0.0), (2.0, 1.6)]],
    'OSH_UP_V': [[(0.11, 1.2), (0.8, 1.5)], [(0.8, 1.5), (3.0, 1.2), (7.0, 0.5)], [(0.4, 0.18), (7.0, 0.5), (30.0, 0.6)], [(5.0, 18.0), (10.0, 2.0), (30.0, 0.6)], [(19.0, 25.0), (100.0, 8.0), (280.0, 1.8)], [(150.0, 30.0), (1500.0, 4.0)]],
    'OSH_DN_V': [[(1.8, 0.3), (7.0, 0.5), (35.0, 0.58)], [(0.38, 1.5), (7.0, 1.9), (40.0, 0.7)], [(28.0, 25.0), (45.0, 17.0), (400.0, 5.8)], [(400.0, 5.8), (1000.0, 5.0)], [(7.0, 1.9), (5.942921586155727, 6.854882264526617), (45.0, 17.0)], [(18.0, 1.2), (30.0, 2.2)], [(30.0, 2.2), (100.0, 3.0), (400.0, 5.8)]],
    'GF_H': [[(0.01, 0.16), (1.6, 0.16)], [(0.01, 4.5), (66.0, 4.5)], [(160.0, 20.0), (13.0, 0.3)], [(0.9, 4.5), (0.8, 0.3)], [(0.8, 0.3), (2.5, 0.11)], [(2.5, 0.11), (8.5, 0.01)], [(2.5, 0.11), (15.0, 0.11)], [(15.0, 0.11), (20.0, 0.01)], [(13.0, 0.3), (15.0, 0.11)]],
    'SHELL_H': [[(10.0, 0.22), (0.52, 0.22)], [(1.7, 0.01), (0.2, 1.0), (0.001, 51.0)], [(10.0, 0.9), (0.1, 80.0), (0.001, 1500.0)]],
    'SHELL2_H': [[(0.001, 0.1), (0.03, 0.04), (0.08, 1e-05)], [(0.001, 0.1), (0.3, 0.6)], [(0.3, 0.6), (2.0, 0.09)], [(2.0, 0.09), (6.0, 7e-05)], [(1.0, 1e-05), (150.0, 0.002)], [(0.001, 2.187761623949552), (0.2, 12.0)], [(0.2, 12.0), (3.5, 18.0)], [(0.3, 0.6), (3.5, 18.0)], [(3.5, 18.0), (10.0, 50.0)]],
    'HTRI_H': [[(0.0001, 0.7), (1.0, 0.7)], [(0.0001, 0.3), (0.5, 0.3)], [(0.0001, 0.1), (0.1, 0.1)], [(0.01, 0.04), (0.01, 0.1)], [(0.06, 0.04), (0.06, 0.1)], [(0.1, 0.04), (0.1, 0.1)], [(0.06, 0.1), (0.9, 0.7)], [(0.1, 0.1), (1.0, 0.5)], [(0.005, 0.3), (0.3, 3.0)], [(0.015, 0.3), (0.5, 3.2)], [(0.3, 3.0), (1.0, 9.0)], [(0.5, 3.2), (1.0, 9.0)]],
    'HEDH_H': [[(1.0, 9.0), (50.0, 4.0), (300.0, 0.1)], [(10.0, 100.0), (50.0, 4.0)], [(120.0, 1.2), (2500.0, 0.5)], [(3000.0, 0.1), (2000.0, 5.0)], [(2000.0, 5.0), (3083.1879502493534, 22.0), (15000.0, 100.0)], [(3083.1879502493534, 22.0), (500.0, 7.0)], [(26.0, 15.0), (120.0, 6.0), (500.0, 7.0)]],
    'TD_A': [[(0.001, 2.0), (1.0, 0.3), (10.0, 0.04), (200.0, 0.001)], [(1.8, 0.23), (1.8, 1000.0)]],
    'TD_K': [[(0.003, 1.0), (0.1, 5.5), (4.0, 5.5), (60.0, 1.0)]],
    'TD_T': [[(1.8, 1.0), (100.0, 0.6), (10000.0, 0.1)]],
}


# ══════════════════════════════════════════════════════════════════════════
# ║  SECTION: former flow_pattern_maps.py
# ══════════════════════════════════════════════════════════════════════════
#!/usr/bin/env python3
"""Gas-liquid flow-pattern maps — exact replica of TECHNIP 'Slug flow.xlsm'.

Reproduces every flow-pattern map in the reference workbook (horizontal and
vertical correlations: Gregory & Fogarasi, Shell DEP, Shell DEP2, HTRI, HEDH,
Oshinowo-Charles, Taitel & Dukler) as native Excel scatter charts:

  * boundary curves taken verbatim from the workbook ``Data`` sheet
    (``flow_pattern_data.BOUNDARIES``),
  * the dimensionless operating parameters (Froude numbers, Lockhart-Martinelli
    X, Taitel-Dukler K & T, Baker lambda/psi, mass flux, ...) computed per pipe
    station with the SAME formulas as the workbook ``Calculation`` sheet, at the
    LOCAL station pressure so the operating path migrates with the pressure drop,
  * each station plotted as a point (the full operating path along the line)
    AND the controlling (max rho*v^2) station highlighted distinctly.

All inputs are SI, matching the reference sheet.
"""


import math
from dataclasses import dataclass

from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.chart.marker import Marker
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── unit conversions ──────────────────────────────────────────────────────
LBHR_TO_KGH = 0.45359237
IN_TO_MM = 25.4
G_FP = 9.81   # NOTE: flow-pattern replica uses 9.81 (workbook basis), engine uses G_SI=9.80665                      # the workbook uses 9.81 / 9.814

# ── style ─────────────────────────────────────────────────────────────────
NAVY, STEEL, WHITE = "1F4973", "2E75B6", "FFFFFF"
LGRAY, DGRAY, GRNHDR = "F2F2F2", "404040", "375623"
AMBER, GREEN = "FFF2CC", "E2EFDA"




def _cell(ws, r, c, v=None, bg=None, fg=DGRAY, sz=9, bold=False,
          ha="left", wrap=False, border=True):
    cl = ws.cell(r, c)
    if v is not None:
        cl.value = v
    cl.font = Font(name="Calibri", size=sz, bold=bold, color=fg)
    if bg:
        cl.fill = _fill(bg)
    cl.alignment = Alignment(horizontal=ha, vertical="center", wrap_text=wrap)
    if border:
        cl.border = _thin()
    return cl


# ════════════════════════════════════════════════════════════════════════
#  1.  PARAMETERS  (verbatim port of Calculation-sheet formulas)
# ════════════════════════════════════════════════════════════════════════
@dataclass
class FPParams:
    Vsg: float | None = None
    Vsl: float | None = None
    FrG: float | None = None
    FrL: float | None = None
    Fi: float | None = None
    Gt: float | None = None
    y: float | None = None
    Xtt: float | None = None
    Cgt: float | None = None
    ReL: float | None = None
    GL: float | None = None
    fL: float | None = None
    T: float | None = None
    K: float | None = None
    R: float | None = None
    rlvl2: float | None = None
    rgvg2: float | None = None
    lb: float | None = None
    Fb: float | None = None
    Qg_lb: float | None = None
    Ql_Fb: float | None = None
    QgQl: float | None = None
    Frtp: float | None = None
    invX: float | None = None        # 1/Xtt (HTRI vertical map abscissa)

    def get(self, key):
        return getattr(self, key, None)


def compute_params(mg_kgh, rg, mug_cp, ml_kgh, rl, mul_cp, sigma_nm, d_mm):
    """All flow-pattern parameters (SI), per the workbook Calculation sheet.

    mg/ml  = gas/liquid mass flow (kg/h);  rg/rl = density (kg/m3);
    mug/mul = viscosity (cP);  sigma = surface tension (N/m);  d_mm = ID (mm).
    Any non-finite / impossible result is returned as None.
    """
    p = FPParams()
    if not (d_mm and rg and rl and rl > rg):
        return p
    d = d_mm / 1000.0
    area = math.pi / 4.0 * d * d
    mug = mug_cp / 1000.0            # Pa·s
    mul = mul_cp / 1000.0
    mg = mg_kgh or 0.0
    ml = ml_kgh or 0.0

    def _safe(fn):
        try:
            v = fn()
            return v if (isinstance(v, (int, float)) and math.isfinite(v)) else None
        except Exception:
            return None

    p.Vsg = _safe(lambda: mg / rg / (area * 3600.0))
    p.Vsl = _safe(lambda: ml / rl / (area * 3600.0))
    if p.Vsg is not None:
        p.FrG = _safe(lambda: p.Vsg * (rg * 1000.0 / (rl - rg) / G_FP / d_mm) ** 0.5)
        p.rgvg2 = _safe(lambda: rg * p.Vsg ** 2)
    if p.Vsl is not None:
        p.FrL = _safe(lambda: p.Vsl * (rl * 1000.0 / (rl - rg) / G_FP / d_mm) ** 0.5)
        p.rlvl2 = _safe(lambda: rl * p.Vsl ** 2)
    if p.Vsg and p.Vsl:
        p.Fi = _safe(lambda: p.Vsl * rl ** 0.5 / (p.Vsg * rg ** 0.5))
    p.Gt = _safe(lambda: (mg + ml) / 3600.0 / area)
    p.y = _safe(lambda: mg / (mg + ml)) if (mg + ml) else None
    if p.y:
        p.Xtt = _safe(lambda: ((1 - p.y) / p.y) ** 0.9 * (rg / rl) ** 0.5
                      * (mul / mug) ** 0.1)
        if p.Gt:
            p.Cgt = _safe(lambda: (1 / p.Gt)
                          * (d * G_FP * rg * (rl - rg) * ((1 - p.y) / p.y)) ** 0.5)
    if p.Xtt:
        p.invX = _safe(lambda: 1.0 / p.Xtt)
    p.GL = _safe(lambda: ml / 3600.0 / area)
    if p.GL:
        p.ReL = _safe(lambda: p.GL * d / mul)
    if p.ReL:
        p.fL = (0.079 / p.ReL ** 0.25) if p.ReL > 2000 else (16.0 / p.ReL)
        dpdzL = _safe(lambda: -2 * p.fL * p.GL ** 2 / rl / d)
        if dpdzL is not None:
            p.T = _safe(lambda: (abs(dpdzL) / 9.814 / (rl - rg)) ** 0.5)
    if p.FrG and p.ReL:
        p.K = _safe(lambda: p.FrG * p.ReL ** 0.5)
    if mg and ml and rg and rl:
        p.R = _safe(lambda: (ml / rl) / ((ml / rl) + (mg / rg)))
        p.QgQl = _safe(lambda: ((mg / rg) / (ml / rl)) ** 0.5)
    p.lb = _safe(lambda: (rg * rl / (1.205 * 1000.0)) ** 0.5)
    if sigma_nm:
        p.Fb = _safe(lambda: (0.0728 / sigma_nm)
                     * ((mul / 0.001) * (1000.0 / rl) ** 2) ** (1.0 / 3.0))
    if p.Vsg and p.lb:
        p.Qg_lb = _safe(lambda: p.Vsg * rg / p.lb)
    if p.Vsl and p.Fb:
        p.Ql_Fb = _safe(lambda: p.Vsl * rl * p.Fb)
    if sigma_nm and mg and ml:
        p.Frtp = _safe(lambda: 1452 * ((35.3 * (mg / rg + ml / rl) / 3600.0) ** 2)
                       * ((((0.062 * rl) ** 0.5) * (((10 ** 7) * sigma_nm) ** 1.5)) ** 0.25)
                       / ((0.03937 * d_mm) ** 5)
                       / ((((10 ** 3) * mul) ** 2) ** 0.25))
    return p


def params_from_station(sp, station, gas_density_fn):
    """FPParams for one pipe station, using LOCAL (pressure-dependent) gas ρ."""
    rg_lbft3 = gas_density_fn(station.p_in_psia)
    rg = (rg_lbft3 or 0.0) * LBFT3_TO_KGM3
    rl = (sp.liq_density or 0.0) * LBFT3_TO_KGM3
    mg = (sp.vap_mass or 0.0) * LBHR_TO_KGH
    ml = (sp.liq_mass or 0.0) * LBHR_TO_KGH
    sigma = (sp.liq_surf_tens or 0.0) * DYNCM_TO_NM
    d_mm = (station.id_in or 0.0) * IN_TO_MM
    return compute_params(mg, rg, sp.vap_visc or 0.0, ml, rl,
                          sp.liq_visc or 0.0, sigma, d_mm)


# ════════════════════════════════════════════════════════════════════════
#  2.  MAP DEFINITIONS  (chart -> param pair + axis spec + boundary block)
# ════════════════════════════════════════════════════════════════════════
@dataclass
class MapDef:
    sheet: str
    title: str
    xkey: str
    xlabel: str
    xmin: float
    xmax: float
    ykey: str
    ylabel: str
    ymin: float
    ymax: float
    bnd: str
    logx: bool = True
    logy: bool = True


# ──────────────────────────────────────────────────────────────────────────
#  THE 16 MAPS  (one MapDef == one chart in 'Slug flow.xlsm')
#
#  Each MapDef maps directly onto a chart in the reference workbook.  The
#  ``# chartN`` tag is the workbook's chart number (xl/charts/chartN.xml); the
#  axes / log-scaling / ranges and the boundary block (``bnd`` -> BOUNDARIES
#  key, taken from the Data sheet) are reproduced verbatim.  Output sheets:
#  the three Taitel-Dukler charts share one sheet (as in the workbook); every
#  other map gets its own sheet.
#
#  HORIZONTAL (8 charts on 6 sheets)
#    chart9   Taitel & Dukler   X   vs FrG   (main flow-pattern map, Fig. 8)
#    chart10  Taitel & Dukler   X   vs K     (stratified transition curve C)
#    chart11  Taitel & Dukler   X   vs T     (bubbly/intermittent curve D)
#    chart12  Gregory & Fogarasi  Vsg vs Vsl (apparent velocities)
#    chart13  Shell DEP           FrG vs Fi  (gas Froude vs liquid load)
#    chart14  Shell DEP2          FrG vs FrL (gas vs liquid Froude)
#    chart15  HTRI                R   vs Cgt (homogeneous liq vol frac)
#    chart16  Baker (HEDH block)  Ql·Fb vs Qg/λb  (Baker map, Fig. 6)
#  VERTICAL (8 charts on 8 sheets)
#    chart1   Gregory & Fogarasi  Vsg vs Vsl
#    chart2   Shell DEP2          FrG vs FrL
#    chart3   Shell DEP           FrG vs Fi
#    chart4   HTRI (intube boil)  1/X vs Gt
#    chart5   HEDH upward         ρl·vl² vs ρg·vg²
#    chart6   Oshinowo-Charles up Frtp/√L vs (Qg/Ql)^0.5
#    chart7   HEDH downward       Vsl vs Vsg  (LINEAR axes)
#    chart8   Oshinowo-Charles dn Frtp/√L vs (Qg/Ql)^0.5
#
#  To add/modify a map later: append a MapDef (param keys must exist on
#  FPParams), and add its boundary polylines under a new BOUNDARIES key.
# ──────────────────────────────────────────────────────────────────────────
MAPS = [
    # ── horizontal ──────────────────────────────────────────────────────
    MapDef("H Taitel-Dukler", "TAITEL & DUKLER — HORIZONTAL (X vs Fr)",   # chart9
           "Xtt", "X (Lockhart-Martinelli)", 0.01, 1000,
           "FrG", "Gas Froude Number", 1e-3, 1e4, "TD_A"),
    MapDef("H Taitel-Dukler", "TAITEL & DUKLER — HORIZONTAL (X vs K)",    # chart10
           "Xtt", "X", 1e-3, 100, "K", "K", 1, 100, "TD_K"),
    MapDef("H Taitel-Dukler", "TAITEL & DUKLER — HORIZONTAL (X vs T)",    # chart11
           "Xtt", "X", 1, 1000, "T", "T", 0.01, 10, "TD_T"),
    MapDef("H Gregory-Fogarasi", "GREGORY & FOGARASI — HORIZONTAL",       # chart12
           "Vsg", "Gas apparent velocity (m/s)", 0.01, 1000,
           "Vsl", "Liquid apparent velocity (m/s)", 1e-3, 100, "GF_H"),
    MapDef("H Shell DEP", "SHELL DEP — HORIZONTAL",                       # chart13
           "FrG", "Gas Froude Number", 1e-3, 100,
           "Fi", "Liquid load Fi", 1e-3, 1e4, "SHELL_H"),
    MapDef("H Shell DEP2", "SHELL DEP2 — HORIZONTAL",                     # chart14
           "FrG", "Gas Froude Number", 1e-3, 1000,
           "FrL", "Liquid Froude Number", 1e-5, 100, "SHELL2_H"),
    MapDef("H HTRI", "HTRI — HORIZONTAL TUBE",                            # chart15
           "R", "Homogeneous Liquid Volume Fraction, R", 1e-4, 1,
           "Cgt", "Flow regime parameter, C gt", 0.01, 10, "HTRI_H"),
    MapDef("H Baker (HEDH)", "BAKER MAP — HORIZONTAL TWO-PHASE",          # chart16
           "Ql_Fb", "Ql·Fib  (kg m⁻² s⁻¹)", 1, 1e5,
           "Qg_lb", "Qg/λb  (kg m⁻² s⁻¹)", 0.1, 1000, "HEDH_H"),
    # ── vertical ────────────────────────────────────────────────────────
    MapDef("V Gregory-Fogarasi", "GREGORY & FOGARASI — VERTICAL",        # chart1
           "Vsg", "Gas apparent velocity (m/s)", 0.01, 1000,
           "Vsl", "Liquid apparent velocity (m/s)", 0.01, 100, "GF_V"),
    MapDef("V Shell DEP2", "SHELL DEP2 — VERTICAL",                      # chart2
           "FrG", "Gas Froude Number", 1e-3, 1000,
           "FrL", "Liquid Froude number", 1e-5, 100, "SHELL2_V"),
    MapDef("V Shell DEP", "SHELL DEP — VERTICAL",                        # chart3
           "FrG", "Gas Froude Number", 1e-3, 10,
           "Fi", "Liquid load Fi", 0.01, 1e4, "SHELL_V"),
    MapDef("V HTRI", "HTRI — VERTICAL INTUBE BOILING",                   # chart4
           "invX", "1/X", 0.01, 100,
           "Gt", "Gt, Total mass velocity (kg/s m²)", 10, 1000, "HTRI_V"),
    MapDef("V HEDH up", "HEDH — VERTICAL UPWARD",                        # chart5
           "rlvl2", "ρl·vl²  (kg m⁻¹ s⁻²)", 1, 1e5,
           "rgvg2", "ρg·vg²  (kg m⁻¹ s⁻²)", 0.01, 1e4, "HEDH_V"),
    MapDef("V Oshinowo up", "OSHINOWO-CHARLES — VERTICAL UPWARD",        # chart6
           "Frtp", "Frtp/√L", 0.1, 1e4,
           "QgQl", "(Qg/Ql)^0.5", 0.1, 100, "OSH_UP_V"),
    MapDef("V HEDH down", "HEDH — VERTICAL DOWNWARD",                    # chart7
           "Vsl", "Liquid apparent velocity (m/s)", 0, 2,
           "Vsg", "Gas apparent velocity (m/s)", 0, 2, "HEDH_DN_V",
           logx=False, logy=False),
    MapDef("V Oshinowo down", "OSHINOWO-CHARLES — VERTICAL DOWNWARD",    # chart8
           "Frtp", "Frtp/√L", 0.1, 1e4,
           "QgQl", "(Qg/Ql)^0.5", 0.1, 100, "OSH_DN_V"),
]


# ════════════════════════════════════════════════════════════════════════
#  3.  PARAMETER TABLE SHEET  ("Froude number and corresponding data")
# ════════════════════════════════════════════════════════════════════════
_TABLE_COLS = [
    ("Seq", "seq"), ("Comp ID", "comp_id"), ("Fitting", "fitting"),
    ("ID (in)", "id_in"), ("P (psia)", "p"),
    ("Vsg (m/s)", "Vsg"), ("Vsl (m/s)", "Vsl"),
    ("Fr Gas", "FrG"), ("Fr Liq", "FrL"), ("Liquid load Fi", "Fi"),
    ("X (L-M)", "Xtt"), ("1/X", "invX"), ("K (T-D)", "K"), ("T (T-D)", "T"),
    ("Gt (kg/s·m²)", "Gt"), ("Vap frac y", "y"),
    ("R (homog)", "R"), ("C gt", "Cgt"),
    ("ρl·vl²", "rlvl2"), ("ρg·vg²", "rgvg2"),
    ("λb", "lb"), ("Fb", "Fb"), ("Qg/λb", "Qg_lb"), ("Ql·Fb", "Ql_Fb"),
    ("(Qg/Ql)^0.5", "QgQl"), ("Frtp/√L", "Frtp"),
]


def _fmt(v):
    if v is None or isinstance(v, str):
        return v
    if isinstance(v, float):
        if not math.isfinite(v):
            return ""
        a = abs(v)
        if a != 0 and (a < 1e-3 or a >= 1e6):
            return float(f"{v:.4E}")
        return round(v, 5)
    return v


def build_param_table(wb, rows, line_no, stream_name, ctrl_seq):
    ws = wb.create_sheet("Flow_Pattern_Data")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = STEEL
    ncol = len(_TABLE_COLS)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    _cell(ws, 1, 1,
          f"FLOW-PATTERN PARAMETERS  |  Line {line_no} <-> Stream {stream_name}"
          "   |   per-station, local pressure  (replica of Slug-flow Calculation)",
          bg=NAVY, fg=WHITE, sz=11, bold=True)
    ws.row_dimensions[1].height = 22
    _u = _U
    for ci, (h, _) in enumerate(_TABLE_COLS, 1):
        if _u is not None:
            if h == "ID (in)":
                h = _u.hdr("ID", "Lin")
            elif h == "P (psia)":
                h = _u.hdr("P", "P")
        _cell(ws, 2, ci, h, bg=NAVY, fg=WHITE, sz=9, bold=True, ha="center", wrap=True)
    ws.row_dimensions[2].height = 30
    ws.freeze_panes = "C3"
    for i, (st, pr) in enumerate(rows):
        r = i + 3
        is_ctrl = (st.seq == ctrl_seq)
        bg = AMBER if is_ctrl else (LGRAY if i % 2 else WHITE)
        for ci, (_, key) in enumerate(_TABLE_COLS, 1):
            if key == "id_in":
                v = getattr(st, key, None)
                if _u is not None:
                    v = _u.disp("Lin", v, None)
            elif key in ("seq", "comp_id", "fitting"):
                v = getattr(st, key, None)
            elif key == "p":
                v = (_u.disp("P", st.p_in_psia, None)
                     if _u is not None else st.p_in_psia)
            else:
                v = pr.get(key)
            _cell(ws, r, ci, _fmt(v), bg=bg, sz=9,
                  bold=is_ctrl, ha="right" if ci > 3 else "left")
    for ci, (h, _) in enumerate(_TABLE_COLS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = max(9, min(15, len(h) + 1))
    return ws


# ════════════════════════════════════════════════════════════════════════
#  4.  MAP SHEETS  (boundary curves + per-station path + controlling point)
# ════════════════════════════════════════════════════════════════════════
def _write_series_cols(ws, col, points, header):
    """Write (x, y) points into two columns; return (xcol, ycol, r0, r1)."""
    _cell(ws, 1, col, f"{header} X", bg=GRNHDR, fg=WHITE, sz=8, border=False)
    _cell(ws, 1, col + 1, f"{header} Y", bg=GRNHDR, fg=WHITE, sz=8, border=False)
    r = 2
    for x, y in points:
        if x is None or y is None:
            continue
        ws.cell(r, col).value = x
        ws.cell(r, col + 1).value = y
        r += 1
    return col, col + 1, 2, max(2, r - 1)


def _add_scatter(ws, mp, data_col, op_points, ctrl_point):
    """Build one ScatterChart for map mp; data laid out from column data_col."""
    chart = ScatterChart()
    chart.title = mp.title
    chart.height = 11
    chart.width = 18
    chart.x_axis.title = mp.xlabel
    chart.y_axis.title = mp.ylabel
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.majorGridlines = None
    if mp.logx:
        chart.x_axis.scaling.logBase = 10
    if mp.logy:
        chart.y_axis.scaling.logBase = 10
    chart.x_axis.scaling.min, chart.x_axis.scaling.max = mp.xmin, mp.xmax
    chart.y_axis.scaling.min, chart.y_axis.scaling.max = mp.ymin, mp.ymax

    col = data_col
    # boundary polylines (grey lines, no markers)
    for i, poly in enumerate(BOUNDARIES.get(mp.bnd, [])):
        xc, yc, r0, r1 = _write_series_cols(ws, col, poly, f"bnd{i}")
        xref = Reference(ws, min_col=xc, min_row=r0, max_row=r1)
        yref = Reference(ws, min_col=yc, min_row=r0, max_row=r1)
        s = Series(yref, xref, title=("regime boundary" if i == 0 else None))
        s.marker = Marker(symbol="none")
        s.graphicalProperties.line.solidFill = "808080"
        s.graphicalProperties.line.width = 14000
        chart.series.append(s)
        col += 3

    # operating path (blue line + small markers)
    xc, yc, r0, r1 = _write_series_cols(ws, col, op_points, "path")
    if r1 >= r0:
        xref = Reference(ws, min_col=xc, min_row=r0, max_row=r1)
        yref = Reference(ws, min_col=yc, min_row=r0, max_row=r1)
        s = Series(yref, xref, title="operating path")
        s.marker = Marker(symbol="circle", size=5)
        s.graphicalProperties.line.solidFill = "2E75B6"
        s.graphicalProperties.line.width = 12000
        chart.series.append(s)
    col += 3

    # controlling point (large red diamond)
    if ctrl_point and ctrl_point[0] is not None and ctrl_point[1] is not None:
        xc, yc, r0, r1 = _write_series_cols(ws, col, [ctrl_point], "ctrl")
        xref = Reference(ws, min_col=xc, min_row=r0, max_row=r1)
        yref = Reference(ws, min_col=yc, min_row=r0, max_row=r1)
        s = Series(yref, xref, title="controlling (max ρv²)")
        s.marker = Marker(symbol="diamond", size=11)
        s.graphicalProperties.line.noFill = True
        chart.series.append(s)
        col += 3
    return chart, col


def build_map_sheets(wb, rows, ctrl_seq):
    """rows = [(station, FPParams)]; build one sheet per map-group with charts."""
    # group maps by output sheet (T&D has three charts on one sheet)
    by_sheet: dict[str, list[MapDef]] = {}
    for mp in MAPS:
        by_sheet.setdefault(mp.sheet, []).append(mp)

    op_seq = [st for st, _ in rows]
    ctrl_st = next((st for st in op_seq if st.seq == ctrl_seq), None)

    for sheet, maps in by_sheet.items():
        ws = wb.create_sheet(sheet[:31])
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = GRNHDR
        data_col = 30                       # data lives far to the right
        for j, mp in enumerate(maps):
            op_points = [(pr.get(mp.xkey), pr.get(mp.ykey)) for _, pr in rows]
            ctrl_pr = next((pr for st, pr in rows if st.seq == ctrl_seq), None)
            ctrl_point = (ctrl_pr.get(mp.xkey), ctrl_pr.get(mp.ykey)) if ctrl_pr else None
            chart, data_col = _add_scatter(ws, mp, data_col, op_points, ctrl_point)
            anchor = f"A{2 + j * 23}"
            ws.add_chart(chart, anchor)
    return list(by_sheet)


# ════════════════════════════════════════════════════════════════════════
#  5.  ENTRY POINT
# ════════════════════════════════════════════════════════════════════════
def build_flow_pattern_maps(wb, stations, sp, line_no, stream_name, gas_density_fn):
    """Add the parameter table + all flow-pattern map sheets to workbook wb."""
    rows = [(st, params_from_station(sp, st, gas_density_fn)) for st in stations]
    if not rows:
        return []
    # controlling station = max ρg·vg² (gas momentum flux), like the FIV sheet
    ctrl_seq = max(
        rows, key=lambda rp: (rp[1].rgvg2 or 0.0))[0].seq
    build_param_table(wb, rows, line_no, stream_name, ctrl_seq)
    sheets = build_map_sheets(wb, rows, ctrl_seq)
    return ["Flow_Pattern_Data"] + sheets


if __name__ == "__main__":
    main()
