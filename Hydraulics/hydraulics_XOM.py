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

import os, re, sys, math, glob, json, gzip, base64
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
    cv_raw: dict | None = None      # control-valve sizing capture (CV datasheet)


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


def station_tp_inputs(station, sp, gas_density_fn, flash_result=None) -> "TPInputs | None":
    """Build SI two-phase inputs for a station, or None if geometry missing.

    ``flash_result``, when given, is this station's own locally-flashed
    FlashResult (same object the pressure-march / Flash_Profile sheet uses) —
    its vap_mass/liq_mass (re-equilibrated at THIS station's local pressure)
    drive the phase split, exactly as ``_dp_flashed`` already does for the
    Δp calc.  Without it (e.g. the no-flash legacy ``run()`` path) the static
    feed-level ``sp.vap_mass``/``sp.liq_mass`` is used as before — correct
    only when the phase split truly does not change along the line.
    """
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

    if flash_result is not None:
        vap_mass = flash_result.vap_mass or 0.0
        liq_mass = flash_result.liq_mass or 0.0
    else:
        vap_mass = sp.vap_mass or 0.0
        liq_mass = sp.liq_mass or 0.0
    mg = vap_mass * LBHR_TO_KGS      # kg/s
    ml = liq_mass * LBHR_TO_KGS
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


def build_regime_sheet(wb, stations, sp, line_no, stream_name, gas_density_fn,
                        flashes=None):
    """Two-phase flow-regime map across all stations + slug summary.

    ``flashes``, when given, is the per-station list of FlashResult objects
    from the rigorous pressure march (same list Flash_Profile/Pressure_Profile
    use) — station i's local vap_mass/liq_mass drive its GVF, instead of the
    static feed-level ``sp`` used for every station.  Without it, the sheet
    falls back to the old feed-level behaviour (correct only when ``sp`` is
    known not to vary along the line, e.g. the no-flash legacy path).
    """
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
    for i, s in enumerate(stations):
        fr = flashes[i] if flashes is not None and i < len(flashes) else None
        t = station_tp_inputs(s, sp, gas_density_fn, fr)
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
    "Tee Join Branch Flow 3",
    "Tee Join Branch Flow 4",
    "Tee Join Branch Flow 5",
    "Tee Join Branch Flow 6",
    "Tee Join Line Flow",
    "Tee Split Branch Flow 1",
    "Tee Split Branch Flow 2",
    "Tee Split Branch Flow 3",
    "Tee Split Branch Flow 4",
    "Tee Split Branch Flow 5",
    "Tee Split Branch Flow 6",
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
    "tee join branch flow 3":                 1.80,
    "tee join branch flow 4":                 1.80,
    "tee join branch flow 5":                 1.80,
    "tee join branch flow 6":                 1.80,
    "tee join line flow":                     0.40,
    "tee split branch flow 1":                1.80,
    "tee split branch flow 2":                1.80,
    "tee split branch flow 3":                1.80,
    "tee split branch flow 4":                1.80,
    "tee split branch flow 5":                1.80,
    "tee split branch flow 6":                1.80,
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
    "Upstream Line No",       # Branch only: Main Line No it splits from; blank → previous Main line
    "Upstream Seq",           # Branch only: exact Main Seq of the 'Tee Split Branch Flow N' row
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
    "Exch Max Allow dP (psi)",  # T-control floor / P & T min-dP floor
    # ── Control-valve sizing / datasheet (Control Valve rows) ────────────────
    "Valve Body Style",       # Globe | Angle | Ball | Butterfly | Eccentric
    "Valve Characteristic",   # Linear | Equal% | Quick-Open  (F & L)
    "Design Opening %",       # target max-flow % travel for auto rated-Cv pick
    "Rated Cv",               # existing valve rated Cv100 (blank ⇒ engine sizes)
    "Min Flow Mult",          # turndown: Min-flow multiple of normal (default 0.35)
    "Max Flow Mult",          # turndown: Max-flow multiple of normal (default 1.20)
    "Noise Limit dBA",        # project noise limit (default 85)
    "Seat Leakage Class",     # II | III | IV | V | VI (default IV)
    "Notes",                  # free notes
]
_NCOL = len(INPUT_HEADERS)   # 45

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
    "Upstream Line No": ("Upstream Line No", None),
    "Upstream Seq": ("Upstream Seq", None),
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
    "Valve Body Style": ("Valve Body Style", None),
    "Valve Characteristic": ("Valve Characteristic", None),
    "Design Opening %": ("Design Opening %", None),
    "Rated Cv": ("Rated Cv", None),
    "Min Flow Mult": ("Min Flow Mult", None),
    "Max Flow Mult": ("Max Flow Mult", None),
    "Noise Limit dBA": ("Noise Limit dBA", None),
    "Seat Leakage Class": ("Seat Leakage Class", None),
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
        "Upstream Line No": 14, "Upstream Seq": 11,
        "Flow Fraction from Main": 14, "Temp (degF)": 11, "Vapor Mass Flow": 14,
        "Vap MW": 9, "Vap Visc (cP)": 11, "Vap Z": 8, "Vap Cp/Cv": 10,
        "Vap Density (lb/ft3)": 14, "Liq Mass Flow (lb/h)": 14,
        "Liq Density (lb/ft3)": 14, "Liq Visc (cP)": 10, "Comp ID": 12,
        "Fitting Name": 32, "Bore (in)": 9, "Piping Spec": 11, "Length (ft)": 11,
        "Elev Change (ft)": 13, "Direction": 10, "Fixed K": 9,
        "Fixed dP (psi)": 12, "Instr Type": 10, "Instr Tag": 12,
        "Instr dP (psi)": 12, "Set P (psia)": 12, "Control Valve Type": 14,
        "Exch Max Allow dP (psi)": 16,
        "Valve Body Style": 15, "Valve Characteristic": 16,
        "Design Opening %": 13, "Rated Cv": 10, "Min Flow Mult": 12,
        "Max Flow Mult": 12, "Noise Limit dBA": 13, "Seat Leakage Class": 15,
        "Notes": 30,
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

    # ── Drop-downs: control-valve datasheet columns ───────────────────
    for hdr, opts, err in [
        ("Valve Body Style", "Globe,Angle,Ball,Butterfly,Eccentric",
         "Valve body style"),
        ("Valve Characteristic", "Linear,Equal%,Quick-Open",
         "Inherent trim characteristic (F & L valves)"),
        ("Seat Leakage Class", "II,III,IV,V,VI",
         "FCI 70-2 / IEC 60534-4 seat leakage class"),
    ]:
        col = _hcol_letter(hdr)
        dv = DataValidation(type="list", formula1=f'"{opts}"', allow_blank=True,
                            showErrorMessage=True, error=err, errorTitle=hdr)
        ws.add_data_validation(dv)
        dv.add(f"{col}3:{col}2000")

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
        {"Circuit":"C3","Line No":"L-301","Run Type":"Main","Seq":3,"Comp ID":"FCV-301","Fitting Name":"Control Valve","Bore (in)":4,"Piping Spec":"G1A-5","Control Valve Type":"F","Valve Body Style":"Globe","Valve Characteristic":"Equal%","Design Opening %":80,"Min Flow Mult":0.35,"Max Flow Mult":1.2,"Noise Limit dBA":85,"Seat Leakage Class":"IV","Length (ft)":0,"Elev Change (ft)":0,"Notes":"Flow control — datasheet auto-sizes Rated Cv (leave Rated Cv blank)"},
        {"Circuit":"C3","Line No":"L-301","Run Type":"Main","Seq":4,"Comp ID":"DST-01","Fitting Name":"Destination","Bore (in)":6,"Piping Spec":"G1A-5","Set P (psia)":60.0,"Length (ft)":0,"Elev Change (ft)":0,"Notes":"Downstream destination pressure"},
        # ── Circuit C4 — two parallel control valves (Tee split, ratio control) ──
        {"Circuit":"C4","Line No":"L-401","Run Type":"Main","Seq":1,"Stream Lookup":"T801-OH","Comp ID":"SRC-02","Fitting Name":"Source","Bore (in)":8,"Piping Spec":"G1A-5","Set P (psia)":200.0,"Length (ft)":0,"Elev Change (ft)":0,"Notes":"Header source"},
        {"Circuit":"C4","Line No":"L-401","Run Type":"Main","Seq":2,"Comp ID":"TEE-40","Fitting Name":"Tee Split Branch Flow 1","Bore (in)":8,"Piping Spec":"G1A-5","Flow Fraction from Main":-0.5,"Length (ft)":0,"Elev Change (ft)":0,"Notes":"50% splits to parallel branch"},
        {"Circuit":"C4","Line No":"L-401","Run Type":"Main","Seq":3,"Comp ID":"FCV-40A","Fitting Name":"Control Valve","Bore (in)":4,"Piping Spec":"G1A-5","Control Valve Type":"F","Fixed dP (psi)":25.0,"Valve Body Style":"Globe","Valve Characteristic":"Equal%","Rated Cv":50,"Length (ft)":0,"Elev Change (ft)":0,"Notes":"Existing valve — Rated Cv 50 given → adequacy check"},
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
        ("  P & T both honour 'Exch Max Allow dP (psi)' as a MINIMUM ΔP floor across the valve.", False),
        ("  β ratio = valve Bore (in) ÷ upstream line bore is reported in the Notes column.", False),
        ("  SERIES valves: place several Control Valve rows in Seq order on the same run.", False),
        ("  PARALLEL valves: split flow at a Tee (negative 'Flow Fraction from Main'), put one Control", False),
        ("    Valve on the Main leg and one on a Branch (Run Type = Branch, with its own Start P).", False),
        ("", False),
        ("CONTROL VALVE DATASHEET  (one sheet per Control Valve, IEC 60534)", True),
        ("  Each Control Valve row produces a 'CV_<tag>' datasheet: Min/Norm/Max sizing, cavitation,", False),
        ("  seat leakage, β ratio and IEC 60534-8-3/-8-4 dB(A) noise vs the limit.", False),
        ("  Valve Body Style : Globe / Angle / Ball / Butterfly / Eccentric (sets typical FL/xT/Fd).", False),
        ("  Valve Characteristic : Linear / Equal% / Quick-Open (F & L valves) — drives % travel.", False),
        ("  Rated Cv : enter the EXISTING valve's Cv100 → ADEQUACY CHECK mode (verifies that valve).", False),
        ("             Leave BLANK → SIZING/SELECTION mode (engine picks a generic Rated Cv so the", False),
        ("             required Cv, travel window and dB(A) limit are met where physically possible).", False),
        ("  Design Opening % : target max-flow % travel used when the engine auto-selects Rated Cv.", False),
        ("  Min/Max Flow Mult : turndown multiples of the marched Normal flow (defaults 0.35 / 1.20).", False),
        ("  Noise Limit dBA : project sound limit (default 85).  Exceedance is flagged + mitigations", False),
        ("             recommended (Rated-Cv choice cannot change service ΔP/noise — use low-noise /", False),
        ("             multistage trim, a larger body, or split the ΔP across two valves).", False),
        ("  Seat Leakage Class : II/III/IV/V/VI (FCI 70-2 / IEC 60534-4); default IV.", False),
        ("", False),
        ("BRANCHES", True),
        ("  Set Run Type = Branch.  Fill Start P on the first Branch row.", False),
        ("  Branches are marched independently.  FIV/AIV/regime sheets use the Main run.", False),
        ("", False),
        ("AUTO-LINKED SPLITS & JOINS (Main ↔ Branch)", True),
        ("  JOIN (Branch → Main): on a Main row, set Fitting = 'Tee Join Branch Flow N'.", False),
        ("  The Nth Branch block sharing that row's Line No is auto-matched by ordinal count;", False),
        ("  its real flashed outlet (flow + composition) is blended into Main at that row — no", False),
        ("  manual entry needed on the Branch side.", False),
        ("  SPLIT (Main → Branch): on a Main row, set Fitting = 'Tee Split Branch Flow N' with a", False),
        ("  negative 'Flow Fraction from Main'.  The exact flow + composition that leaves Main at", False),
        ("  that row is captured and auto-fed as the Branch's inlet — the Branch needs no Stream", False),
        ("  Lookup or manual properties of its own.", False),
        ("  Upstream Line No (Branch, optional): which Main Line No this Branch splits from.", False),
        ("    Blank → defaults to the most recently seen Main Line No.", False),
        ("  Upstream Seq (Branch, optional): the exact Seq of the Main 'Tee Split Branch Flow N' row.", False),
        ("    Blank → matched by ordinal count, same as the Join side (N = 1, 2, 3 … up to 6).", False),
        ("  A Branch with no Upstream Line No match falls back to standalone marching (own Stream", False),
        ("  Lookup / manual properties / circuit default).", False),
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
    current_main_line: str | None = None # most recent Main Line No (default Upstream Line No)

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
            current_main_line = None

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

        # Upstream Line No / Seq (Branch only): where this branch splits off
        # Main.  Blank Upstream Line No → most recent Main Line No seen so far.
        ul_raw = txt(g(rv, "Upstream Line No"))
        upstream_line = ul_raw or (current_main_line if run_type == "Branch" else None)
        upstream_seq  = num(g(rv, "Upstream Seq"))
        if run_type == "Main":
            current_main_line = line_no

        rows.append({
            "Circuit":          circuit,
            "Line No":          line_no,
            "HMB File":         current_hmb,
            "Case":             current_case,
            "Stream Lookup":    current_stream,
            "Start P (psia)":   sp_val,          # raw (only non-None on explicit rows)
            "Upstream Line No": upstream_line,
            "Upstream Seq":     int(upstream_seq) if upstream_seq is not None else None,
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
            # ── Control-valve sizing / datasheet ────────────────────────
            "Valve Body Style":     txt(g(rv, "Valve Body Style")),
            "Valve Characteristic": txt(g(rv, "Valve Characteristic")),
            "Design Opening %":     num(g(rv, "Design Opening %")),
            "Rated Cv":             num(g(rv, "Rated Cv")),
            "Min Flow Mult":        num(g(rv, "Min Flow Mult")),
            "Max Flow Mult":        num(g(rv, "Max Flow Mult")),
            "Noise Limit dBA":      num(g(rv, "Noise Limit dBA")),
            "Seat Leakage Class":   txt(g(rv, "Seat Leakage Class")),
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


def _is_tee_split_branch(fitting: str | None) -> int | None:
    """If ``fitting`` is a 'Tee Split Branch Flow N', return N; else None."""
    f = (fitting or "").lower()
    if "tee" in f and "split" in f and "branch" in f:
        m = re.search(r"(\d+)\s*$", f)
        return int(m.group(1)) if m else 1
    return None


def _feed_comp_mole_fracs(feed, fr) -> dict:
    """{component: (y_i, x_i)} vapor/liquid mole fractions for one station's
    flash.  Mole fractions are scale-independent, so this always reflects the
    feed's own composition regardless of any mass-flow override/carry-over.
    Returns {} when no feed/composition is available."""
    if feed is None or fr is None or not getattr(feed, "names", None):
        return {}
    y = fr.y or []
    x = fr.x or []
    return {
        name: (y[i] if i < len(y) else 0.0, x[i] if i < len(x) else 0.0)
        for i, name in enumerate(feed.names)
    }


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


def _apply_property_overrides(sp, fr, overrides: dict, feed_mode: bool = False):
    """Overlay persisted yellow-cell overrides onto a resolved sp/fr pair.

    Lets any filled yellow cell win over the matching property pulled from
    the HMB stream, without having to clear Stream Lookup and re-enter every
    property by hand. In MANUAL mode (no composition feed) mass-flow
    overrides also re-derive quality/beta/moles on the FlashResult directly,
    since ``fr`` there has no other source of truth.  In FEED mode (a
    composition feed flashed via Rachford-Rice), vap_mass/liq_mass overrides
    are instead applied upstream as a persistent mass-scale ratio on the
    whole feed-flash (see ``build_profile_flash_noiso``) — they are dropped
    here to avoid double-counting the override against the feed's own
    un-scaled natural flow.
    """
    if not overrides or sp is None:
        return sp, fr
    if feed_mode:
        overrides = {k: v for k, v in overrides.items()
                     if k not in ("vap_mass", "liq_mass")}
        if not overrides:
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
        default_feed=None,
        default_vap_cf: float = 1.0,
        default_liq_cf: float = 1.0,
        splits_out: dict | None = None,
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

    Main→Branch split (the reverse of a Branch→Main join):
      ``default_feed``/``default_vap_cf``/``default_liq_cf`` seed the FIRST
      row of a Branch block — used to auto-feed a Branch with the actual
      mixture + mass flow that left Main at a matching 'Tee Split Branch
      Flow N' row (see ``splits_out`` below and ``run_noiso``'s Upstream
      Line No / Upstream Seq matching).

    ``splits_out``, if given, is populated with one entry per encountered
    'Tee Split Branch Flow N' row, keyed by (Line No, Seq) of that row:
    ``{"n": N, "feed": ..., "sp": ..., "vap_cf": ..., "liq_cf": ...}`` — the
    composition/properties feed and mass-fraction multipliers of the flow
    that left Main there, ready to be passed back in as a downstream
    Branch's ``default_feed``/``default_vap_cf``/``default_liq_cf``.

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
    comp_fracs: list[dict] = []   # comp_fracs[i] = {component: (y_i, x_i)}
    p       = max(p_start, p_floor)
    cum     = 0.0
    last_bore = None
    flow_scale = 1.0          # running Tee split/merge multiplier (both phases)
    vap_cf = 1.0              # running vapour carry-forward multiplier
    liq_cf = 1.0              # running liquid carry-forward multiplier
    mass_scale = 1.0          # feed-mode override: rescales the whole feed-flash
    mass_scale_set = False    # to the user's real total flow (fixed once per stream)
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
            cur_feed = row_feed if row_feed is not None else (
                default_feed if ridx == 0 else None)
            cur_sp   = row_sp if row_sp is not None else default_sp
            last_key = key
            flow_scale = 1.0
            vap_cf = default_vap_cf if ridx == 0 else 1.0
            liq_cf = default_liq_cf if ridx == 0 else 1.0
            mass_scale = 1.0
            mass_scale_set = False
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
        split_n = _is_tee_split_branch(fitting)
        pending_split_scale = None    # filled below if this row splits off a branch
        if frac is not None and _is_tee(fitting):
            new_scale = frac if frac > 0 else max(0.0, flow_scale + frac)
            verb = "merge" if frac > 0 else "split"
            scale_tag += f" | Tee {verb} {frac:+.3f} → flow×{new_scale:.3f}"
            if frac < 0 and split_n is not None:
                pending_split_scale = flow_scale - new_scale   # magnitude removed
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
        comp_vec: dict = {}
        fr_full = None
        if a_feed is not None:
            for i, n in enumerate(a_feed.names):
                if n not in mw_map and i < len(a_feed.mw):
                    mw_map[n] = a_feed.mw[i]
            fr_full = _flash(a_feed, p_eval)
            # A Vapor/Liquid Mass Flow override on a Stream-Lookup row means
            # "this composition's REAL total flow is X" — fix a single ratio,
            # once per stream, between that target and the feed's own
            # natural (un-scaled) total, and apply it like any other
            # carry-forward multiplier.  This preserves the feed's rigorous
            # VLE phase split at each pressure instead of substituting one
            # phase's absolute override value and adding it to the other
            # phase's un-rescaled natural value (which double-counts mass
            # from two different bases).
            if not mass_scale_set and ("vap_mass" in override_state
                                        or "liq_mass" in override_state):
                nat_total = fr_full.vap_mass + fr_full.liq_mass
                tgt_total = (override_state.get("vap_mass", fr_full.vap_mass)
                             + override_state.get("liq_mass", fr_full.liq_mass))
                if nat_total > 0:
                    mass_scale = tgt_total / nat_total
                mass_scale_set = True
            vap_eff = flow_scale * vap_cf * mass_scale
            liq_eff = flow_scale * liq_cf * mass_scale
            fr = _scale_flash_result(fr_full, vap_eff, liq_eff)
            comp_vec = _feed_comp_flows(a_feed, fr_full, vap_eff, liq_eff)
        elif a_sp is not None and (a_sp.vap_mass or a_sp.liq_mass):
            vap_eff = flow_scale * vap_cf
            liq_eff = flow_scale * liq_cf
            fr_full = _synth_flash_result(a_sp, p_eval)
            fr = _scale_flash_result(fr_full, vap_eff, liq_eff)  # manual mode
        else:
            vap_eff = flow_scale * vap_cf
            liq_eff = flow_scale * liq_cf
            fr = None
        sp = a_sp                       # used by the ΔP dispatch below

        # ── Main→Branch split capture: stash the flow that LEFT here ─────
        if pending_split_scale is not None and splits_out is not None and fr_full is not None:
            splits_out[(row.get("Line No"), int(num(row.get("Seq")) or 0))] = {
                "n":      split_n,
                "feed":   a_feed,
                "sp":     a_sp,
                "vap_cf": pending_split_scale * vap_cf,
                "liq_cf": pending_split_scale * liq_cf,
            }
        if override_state:
            sp, fr = _apply_property_overrides(sp, fr, override_state,
                                                feed_mode=(a_feed is not None))
            scale_tag += (" | manual override: "
                          + ", ".join(sorted(override_state)))

        if vap_eff <= 1e-9 and liq_eff <= 1e-9:
            scale_tag += ("  | stream terminated (carry-forward 0)"
                          if (v_cf == 0 or l_cf == 0) else
                          "  | flow≈0 (fully split off)")

        phase_str = (_phase_of(fr.quality) if fr else (sp.phase if sp else ""))
        flash_tag = (f" (flash β={fr.beta:.3f})" if fr else "") + scale_tag
        cv_raw = None                      # set only on Control Valve rows

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
            # P & T control: ΔP floored at the min-allowable ΔP in col AJ
            # (exchanger max-allowable for T; specified min drop for P).
            floor_tag = ""
            if ctype in ("P", "T") and exch_dp is not None and exch_dp > base_dp:
                lbl = "exchanger max-allow" if ctype == "T" else "min-allow"
                floor_tag = f"; {ctype}-control floor → {lbl} ΔP {exch_dp:.3f} psi"
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
            # capture everything the CV datasheet needs (sized later, outside
            # the march, from the authoritative post-choke station pressures)
            _cv_sp, _cv_fr = a_sp, fr
            cv_raw = {
                "row": dict(row), "ctype": ctype, "dest_p": dest_p,
                "p_src": p_start, "line_bore_in": line_bore_before,
                "bore_in": v_bore, "exch_dp": exch_dp, "fixed_dp": fixed_dp,
                "feed": a_feed,
                "temp_f": getattr(_cv_sp, "temp_f", None),
                "phase": getattr(_cv_sp, "phase", None),
                "pc_psia": getattr(_cv_sp, "pc_psia", None),
                "mol_weight": getattr(_cv_sp, "mol_weight", None),
                "vap_mw": getattr(_cv_fr, "vap_mw", None) if _cv_fr else None,
                "liq_density": getattr(_cv_sp, "liq_density", None),
                "vap_density": getattr(_cv_sp, "vap_density", None),
                "liq_visc": getattr(_cv_sp, "liq_visc", None),
                "vap_visc": getattr(_cv_sp, "vap_visc", None),
                "vap_z": getattr(_cv_sp, "vap_z", None),
                "gamma": (getattr(_cv_sp, "vap_cp_cv", None)
                          or (_cv_sp.gamma_estimate() if _cv_sp else None)),
                "quality": getattr(_cv_fr, "quality", None) if _cv_fr else None,
                "vap_mass": (getattr(_cv_fr, "vap_mass", None) if _cv_fr
                             else getattr(_cv_sp, "vap_mass", None)),
                "liq_mass": (getattr(_cv_fr, "liq_mass", None) if _cv_fr
                             else getattr(_cv_sp, "liq_mass", None)),
            }

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
            cv_raw       = cv_raw,
        ))
        flashes.append(fr)
        comps.append(comp_vec)
        comp_fracs.append(_feed_comp_mole_fracs(a_feed, fr_full))
        p = p_out

    return stations, flashes, comps, mw_map, comp_fracs


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
        default_feed=None,
        default_vap_cf: float = 1.0,
        default_liq_cf: float = 1.0,
        splits_out: dict | None = None,
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
                  default_sp=default_sp, injections=injections,
                  default_feed=default_feed, default_vap_cf=default_vap_cf,
                  default_liq_cf=default_liq_cf, splits_out=splits_out)
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
#  Stream_Props_Detail sheet — fitting-wise Total/Vapor/Liquid breakdown
# ════════════════════════════════════════════════════════════════════════
def _build_stream_props_detail_sheet(
        wb: Workbook,
        stations: list[HE.Station],
        flashes: list,
        comp_fracs: list[dict],
        sp: HE.StreamProps,
        circuit_id: str,
        stream_name: str,
        run_label: str = "Main",
) -> str:
    """One Property/Total/Vapor/Liquid block per fitting, plus a per-fitting
    component vapor/liquid mole-fraction table.

    Only properties that the engine actually re-evaluates at each fitting's
    (T, P) are shown (flow rates, density, viscosity, Z, mole/mass
    fractions).  Properties only known at the single HMB reference state
    (Cp, thermal conductivity, std density/API gravity, enthalpy, surface
    tension) are NOT carried into this per-fitting table — they don't vary
    with the march and would be misleading presented as if they did.
    """
    sheet_title = (f"Stream_Props_Detail_{run_label}"
                    if run_label != "Main" else "Stream_Props_Detail")
    ws = wb.create_sheet(sheet_title)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = GRNHDR
    ncol = 4
    ws.merge_cells(f"A1:{get_column_letter(ncol)}1")
    _c(ws, 1, 1,
       f"STREAM PROPERTIES BY FITTING  [{run_label}]   |   Circuit {circuit_id}   |   "
       f"Stream: {stream_name}",
       bg=NAVY, fg=WHITE, sz=11, bold=True)
    ws.row_dimensions[1].height = 20

    u = _u()
    r = 3
    for s, fr, fracs in zip(stations, flashes, comp_fracs):
        if fr is None:
            continue
        vap_rho = (_gas_rho_kgm3(fr.vap_mw, sp.vap_z, sp.temp_f, s.p_in_psia)
                   / LBFT3_TO_KGM3) if sp.vap_z else None
        liq_rho = sp.liq_density
        q_vap = (fr.vap_mass / vap_rho) if (vap_rho and fr.vap_mass) else None
        q_liq = (fr.liq_mass / liq_rho) if (liq_rho and fr.liq_mass) else None
        tot_mass  = fr.vap_mass + fr.liq_mass
        tot_moles = fr.vap_moles + fr.liq_moles
        tot_mw    = (tot_mass / tot_moles) if tot_moles else None

        ws.merge_cells(f"A{r}:{get_column_letter(ncol)}{r}")
        _c(ws, r, 1,
           f"Seq {s.seq} · {s.comp_id or ''} · {s.fitting or ''}   "
           f"({u.disp('P', s.p_in_psia)} {u.label('P')})",
           bg=STEEL, fg=WHITE, sz=10, bold=True)
        r += 1
        _hdr(ws, ["Property", "Total", "Vapor", "Liquid"], row=r)
        head = r
        r += 1

        def row(label, tot, vap, liq, bold=False, section=False):
            nonlocal r
            if section:
                ws.merge_cells(f"A{r}:{get_column_letter(ncol)}{r}")
                _c(ws, r, 1, label, bg=DGRAY, fg=WHITE, sz=9, bold=True)
                r += 1
                return
            bg = LGRAY if (r - head) % 2 else WHITE
            _c(ws, r, 1, label, bg=bg, sz=9, bold=bold)
            for ci, v in ((2, tot), (3, vap), (4, liq)):
                _c(ws, r, ci, "" if v is None else v, bg=bg, sz=9,
                   ha="right" if isinstance(v, (int, float)) else "left", bold=bold)
            r += 1

        row("FLOW RATES", None, None, None, section=True)
        row(f"Molar Rate ({u.label('molflow')})",
            u.disp("molflow", tot_moles, 3) if tot_moles else 0,
            u.disp("molflow", fr.vap_moles, 3), u.disp("molflow", fr.liq_moles, 3))
        row(f"Mass Rate ({u.label('mflow')})",
            u.disp("mflow", tot_mass, 1), u.disp("mflow", fr.vap_mass, 1),
            u.disp("mflow", fr.liq_mass, 1))
        row("Actual Vol Rate (ft3/hr)",
            round((q_vap or 0) + (q_liq or 0), 1),
            round(q_vap, 1) if q_vap is not None else None,
            round(q_liq, 1) if q_liq is not None else None)

        row("CONDITIONS", None, None, None, section=True)
        row(f"Temperature ({u.label('T')})", u.disp("T", sp.temp_f, 2), None, None)
        row(f"Pressure ({u.label('P')})", u.disp("P", s.p_in_psia), None, None)
        row("Molecular Weight",
            round(tot_mw, 4) if tot_mw else None,
            round(fr.vap_mw, 4), round(fr.liq_mw, 4))
        row("Vapor Mole Fraction", round(fr.beta, 4), None, None)
        row("Liquid Mole Fraction", round(1.0 - fr.beta, 4), None, None)
        row("Liquid Mass Fraction (quality)", round(1.0 - fr.quality, 4), None, None)

        row("VAPOR PHASE PROPERTIES", None, None, None, section=True)
        row(f"Density ({u.label('rho')})", None,
            u.disp("rho", vap_rho, 6) if vap_rho else None, None)
        row("Viscosity (cP)", None, sp.vap_visc, None)
        row("Z Factor", None, sp.vap_z, None)

        row("LIQUID PHASE PROPERTIES", None, None, None, section=True)
        row(f"Density ({u.label('rho')})", None, None,
            u.disp("rho", liq_rho, 4) if liq_rho else None)
        row("Viscosity (cP)", None, None, sp.liq_visc)
        r += 1

        # ── Component vapor/liquid mole-fraction table for this fitting ──
        if fracs:
            names = sorted(fracs.keys(), key=lambda n: -(fracs[n][0] + fracs[n][1]))
            _c(ws, r, 1, "Component Mole Fraction", bg=DGRAY, fg=WHITE, sz=9, bold=True)
            for ci, name in enumerate(names, 2):
                _c(ws, r, ci, name, bg=GRNHDR, fg=WHITE, sz=8, bold=True)
            r += 1
            _c(ws, r, 1, "Vapor (y)", bg=LGRAY, sz=9, bold=True)
            for ci, name in enumerate(names, 2):
                _c(ws, r, ci, round(fracs[name][0], 6), sz=8, ha="right")
            r += 1
            _c(ws, r, 1, "Liquid (x)", bg=LGRAY, sz=9, bold=True)
            for ci, name in enumerate(names, 2):
                _c(ws, r, ci, round(fracs[name][1], 6), sz=8, ha="right")
            r += 2
        else:
            r += 1

    cw(ws, 1, 30); cw(ws, 2, 16); cw(ws, 3, 16); cw(ws, 4, 16)
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

        # ── Preliminary Main march (no branch injections) ────────────────
        # Captures the flow/composition that leaves Main at every
        # 'Tee Split Branch Flow N' row, so any Branch block that splits off
        # of Main can be auto-seeded below (mirrors the Branch→Main join
        # injection further down, but in the opposite direction).
        splits_out: dict[tuple, dict] = {}
        build_profile_solved(main_rows, resolver, p_start, flash_mode=flash_mode,
                              default_sp=sp, splits_out=splits_out)

        # ── March branch blocks FIRST (independent sub-runs) ─────────────
        # A branch whose outlet matches a Main 'Tee Join Branch Flow N' row on
        # the same line is INJECTED into the Main at that row (its outlet flow +
        # composition merge in); others stay standalone Branch sheets.
        # A branch whose INLET matches a Main 'Tee Split Branch Flow N' row
        # (via "Upstream Line No" / "Upstream Seq") is SEEDED from that row's
        # captured split-off flow/composition instead of marching standalone.
        branch_runs: list = []            # (label, sts, fls, cmps, b_lns, cfracs)
        injections: dict[int, dict] = {}  # main_rows index → branch outlet
        line_branch_count: dict[str, int] = {}
        line_split_count: dict[str, int] = {}
        for bi, block in enumerate(branch_blocks):
            bp = _start_p_of_block(block) or p_start
            b0 = block[0] if block else {}
            up_line = b0.get("Upstream Line No")
            up_seq  = b0.get("Upstream Seq")
            split_cap = None
            if up_line:
                if up_seq is not None:
                    split_cap = splits_out.get((up_line, up_seq))
                else:
                    line_split_count[up_line] = line_split_count.get(up_line, 0) + 1
                    split_n = line_split_count[up_line]
                    split_row = next(
                        (mr for mr in main_rows
                         if mr.get("Line No") == up_line
                         and _is_tee_split_branch(mr.get("Fitting Name")) == split_n),
                        None)
                    if split_row is not None:
                        split_cap = splits_out.get((up_line, split_row.get("Seq")))
                if split_cap is None:
                    print(f"  WARNING: branch (line {b0.get('Line No')}) specifies "
                          f"Upstream Line No '{up_line}' but no matching "
                          f"'Tee Split Branch Flow N' output was found on Main.")

            solve_kwargs = dict(flash_mode=flash_mode, default_sp=sp)
            if split_cap is not None:
                solve_kwargs.update(default_feed=split_cap["feed"],
                                     default_vap_cf=split_cap["vap_cf"],
                                     default_liq_cf=split_cap["liq_cf"])
                if split_cap.get("sp") is not None:
                    solve_kwargs["default_sp"] = split_cap["sp"]
            sts, fls, cmps, _bmw, cfracs = build_profile_solved(block, resolver, bp, **solve_kwargs)
            b_lns = [r["Line No"] for r in block]
            label = f"Branch_{bi+1}" if len(branch_blocks) > 1 else "Branch"
            branch_runs.append((label, sts, fls, cmps, b_lns, cfracs))
            if split_cap is not None:
                print(f"  Branch '{label}' (line {b0.get('Line No')}) splits off "
                      f"Main line {up_line} at Tee Split Branch Flow {split_cap['n']}.")

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
        main_stations, main_flashes, main_comps, main_mw, main_cfracs = build_profile_solved(
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
        sp_sheets = [_build_stream_props_detail_sheet(
            wb, main_stations, main_flashes, main_cfracs, sp,
            circuit_id=cid, stream_name=sk, run_label="Main")]

        for (label, sts, fls, cmps, b_lns, cfracs) in branch_runs:
            pp_sheets.append(_build_pressure_profile_sheet(
                wb, sts, fls, sp,
                circuit_id=cid, stream_name=sk,
                run_label=label, flash_mode=flash_mode,
                station_lines=b_lns, line_colors=lc))
            fp_sheets.append(_build_flash_detail_sheet(
                wb, sts, fls, feed, sp, run_label=label))
            sp_sheets.append(_build_stream_props_detail_sheet(
                wb, sts, fls, cfracs, sp,
                circuit_id=cid, stream_name=sk, run_label=label))

        # ── Control-valve datasheets (one per Control Valve in the circuit) ──
        cv_sheets = _build_cv_datasheets(wb, main_stations, stream_name=sk,
                                         run_label="Main")
        for (label, sts, fls, cmps, b_lns, cfracs) in branch_runs:
            cv_sheets += _build_cv_datasheets(wb, sts, stream_name=sk,
                                              run_label=label)
        if cv_sheets:
            print(f"  Control-valve datasheets: {', '.join(cv_sheets)}")

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
                                   circuit_label, sk, gas_density_fn=gdf,
                                   flashes=main_flashes)
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
                  "Component_Detail", "Stream_Props"]
                 + sp_sheets + cv_sheets +
                 ["FIV_EI_T2.2", "AIV", "Two_Phase_Regime"]
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
    # Components whose K is re-evaluated directly from a physical
    # solubility correlation (Riazi-Vera / AGS) at flash()'s actual (P,T)
    # rather than extrapolated from k_ref via the generic Wilson (Pref/P)
    # shape — those correlations are not the simple multiplicatively
    # separable form Wilson is, so extrapolating their reference-state
    # value with Wilson's shape would silently misrepresent them away from
    # the reference point.  Maps component index -> (gas name, solvent MW).
    k_special: dict[int, tuple[str, float]] | None = None
    k_lambda: float = 1.0     # the single calibration multiplier from _solve_lambda

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

    if feed.k_special:
        for idx, (nm_up, solvent_mw) in feed.k_special.items():
            if nm_up == "H2":
                kv = _ags_k_h2(p, t_used, solvent_mw)
            else:
                kv = _riazi_vera_k(nm_up, p, t_used, solvent_mw)
            if kv is not None:
                k[idx] = min(1e12, max(1e-12, feed.k_lambda * kv))

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


def _wilson_k_full(tc_r, omega, pc_psia, p_psia, temp_f) -> float | None:
    """Full Wilson (1968) K-value correlation, anchored purely to fundamental
    component properties (no HMB phase-composition data required):

        K_i(T,P) = (Pc_i/P) · exp[5.373·(1+ω_i)·(1 − Tc_i/T)]      (T, Tc in °R)

    Used as a fallback reference when the HMB's own y/x (or z/x proxy)
    K-ratios are numerically degenerate — e.g. when the reference molar
    vapor fraction β₀ is extremely close to 0 or 1, the minor phase's
    reported composition carries little real relative-volatility information
    (x_i ≈ z_i for nearly all i), so every component's z_i/x_i ratio collapses
    to ~1 regardless of true volatility.  None if any input is unavailable.
    """
    if tc_r is None or omega is None or pc_psia is None or temp_f is None or p_psia is None:
        return None
    t_r = temp_f + F_TO_R
    if t_r <= 0 or pc_psia <= 0 or p_psia <= 0:
        return None
    expo = WILSON_C * (1.0 + omega) * (1.0 - tc_r / t_r)
    expo = max(-50.0, min(50.0, expo))
    return (pc_psia / p_psia) * math.exp(expo)


# ══════════════════════════════════════════════════════════════════════════
# ║  SECTION: PRO/II Custom Properties library (former proii_props.json)
# ══════════════════════════════════════════════════════════════════════════
# Real component properties exported directly from PRO/II's "Custom
# Properties" library (PROCESS / LNG / REFORMER / HFALKYLATION exports),
# gzip+base64-packed to keep this self-contained as a single .py file.
# 1209 components, keyed by upper-case name. See _proii_lookup() below.
# (1209 components, 641571 raw JSON bytes)
_PROII_PROPS_B64 = (
    'H4sIADz7N2oC/+S9S68cOZYm+FcSsZlKQO7geZLMnVKhrBBaoatSKLJTuZlV7xozjRkMMEBj/vuQNDMazYxmbqS7VUWhu4FC'
    'hjJDurr388Pz+B7/86d/gduvH7//8uNz+D8fv/z4/OfbXz9++efHLx9/+sv//On9h49fvn/79OGnv5g7Adp3P73//fvbp3/9'
    '8un7p7cvP/3FqruDvPvpw5fff/3pL/5u3v3086ePnz9+GP6l/+P/+e//Pf7K17fPH+NvEf7rv31+/9svX98+ffn+018INP3b'
    '//q3t2+/vh9+RzROUOP/8l9/+fv0O/zy4e3Xv/7+2/f8z8W/AECGOf4Lv/zt99/CL3359Wv+330Kf9yvv+Z//Pv7r1/+mv/b'
    'z/Mf8Plt/B+Gr9EYePfTr39/+4yS/t7h9xfnwy/91/iHhb+ve/dT+l2Y3F3Dl5/+QGR/9+EfvoZ/h0jT3/Xr9/xnfXv/4b98'
    '/P59/sefP/3+W/gt7sD2v90g/I9/+1c1f4t/ngdC5PALb59///r+2/uf/uLCr/n4zf/t7fdvH8J38qev394+fPztt//xf/2f'
    '/+P+//73//undz99D3+uio1/7Pf5j/39y6e/vf/wv/9b+JruXl3+hW8//UXuhg2+++nv6S/J8bf/Z/qPqBT+Y/qB3tzd/H/v'
    'AkTu7z9//eX9/R1O/4Huf/34/f39zzd4h+/oFn7cAURrxKDFNWCGL2uAi2uEC1rawIWN93CH02i5gYpXeRVc/F3+283IEi/s'
    'rWS84B3BjoAhR3eZAAMBI6wDYNAFyHt7BjHDP2WwWLFebAkWeyfv+SFWxGL6ZlbBsvjnb9M/j0gBzUhhiwgTWMyAlX9GTNDt'
    '508ffvn89u3thre//v69Uk7Q6xE4+BgcGL/Nzq4A4v3dresJeJI7nweIiDqO/8Ir8DH8YwkO44hwBAfy3ftcTQzccQLH8O+l'
    'WmKtD5/OZmiED0UoB1giw99Vffx5PSojY1neqSMCWGIj/KBI/YQOQj+jI3z+NZcSzPCg8cEJ0Pga8FDDRvwN97GhjYVj+K9L'
    'VFgKf0TDI3MTG/7f3V+FivAxAh5QEcAHilPFYBP/0KFghOeGRlCEghHehI56od6pUImK8X/woFqA231ZjooFCWQ42G2h4IdI'
    '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'
    '5vBD3iyr4ybqkJh5OHP6dpvcis16+HSJaeHQgDce3JWHcWsnvl28wOX5grVgYkamR76ME1Mic3eYTliA9YCh5gRTQrEzKK4g'
    '3FkpDQ51AYq9bLAwENFzLeXam0b91ptmXLWchsS10YGABjN7BmPWUz58xmSZ3ETwrBI9nSjJceJbk+1QNlOnZaUnqRKW17Lh'
    'UBJmsp2zFd3wjIoUJvk2cjDf/Sn+czQw+u17xcCIyPNLY1wqnlah2YRWiXlMpLw0+Ml5O9PvgPnBY4JyGiVEm9wnMLAaQ214'
    'XvlUcNw+Svwyd5TvCol8Od0/CxI3bHvNhIojC2U5HEWAmi2UK8yJUJttyw00tIPeXgcLy95MsLBJZjttMVkLlxpb+K7HTQFS'
    'X1Jcer1KOg05g3Ti8tXpeSauiKIVLnOsMyx+fPv29vnTz5XxlOMl9Zlwc7qvRtN0WfarzfYtfEfbVttw8WRq1SXZdOo8wyia'
    'bLZHllUZMunnnDgFe5rCbSyvXhUAFb+aTI09MZpajwc5xSq61HyE+dfnV2VIuh9bULvKpKV39I5v4Yd6yNs9Poy2U622zlbo'
    'IktdW6hWyhYupVpZ9JnKvaRaha80U62iAmiqGYqczGE69OWeaeXFH9s8f8raCvuY3MWkSro2043AkBIYdY6/cxeTdsNs0hYK'
    'huIsX8jv95i+c5mZ6Wb/GlfLdiHsNLgK7abnLk9d7XlCqEx00Q0YbuGneCgK26YV63N8TF/h9otP+TnniwQRnTI0OhlVbCtV'
    'IjwbWq0S4dmHnM7A4Y+ZSP1q+e57FhhWrDIvy4RYfyayIyZmdZUJWzrgDYyrDTIe+J8BwrUSYx99yFsEg0zRMMZf6EziuSYx'
    'huQ2kOWC2dBIMJ2JegiYjtrPH2NcVoel0Z7CuMTBXhMBr9WNbvOeIJ4Fm96KKCi8cOxQcFiXjYYvE3yOkLTZHBNjdp/tImsH'
    'TK+ouOBPCX4UOktDKRwNf1e7Xxr2CNvqXmpaVFlusmNb6MFYDJ/Rg3n3BAeTOQyz+lAPJmypqgcTP0bGTVYEOsmKPUtnoQhf'
    'kMcVU9v6Mz54pOlz3xFfXlQLYt30FBzg8eGXz2/f3nZfDjbPzKXb9XctnTr8sKFJQSqhLzPX+VMYl8hIQ0MRGgWfCfxe5wSP'
    '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'
)

def _load_proii_props():
    raw = gzip.decompress(base64.b64decode(_PROII_PROPS_B64))
    return json.loads(raw.decode("utf-8"))

PROII_PROPS = _load_proii_props()


def proii_lookup(name: str) -> dict | None:
    """Look up real PRO/II library properties by component name
    (case-insensitive). Returns the raw property dict (MW, TC, PC, VC, ZC,
    ACENTRIC, SOLUPARA, NBP, SG60F, etc. in PRO/II SI units — K, kPa,
    m3/kg-mol, kJ/kg-mol) or None if not found."""
    return PROII_PROPS.get(name.strip().upper())


# Riazi & Vera (2005, Ind. Eng. Chem. Res. 44, 186-192), Table 1 (DIPPR) +
# Table 4 recommended petroleum/crude-oil correction factors alpha.  CO2's
# alpha is the single fitted value from their Table 3 (no fixed petroleum
# recommendation was given).  H2 is handled separately by the more rigorous
# Augmented Grayson-Streed model below (_ags_k_h2) — Torres, de Hemptinne &
# Machin (2013, OGST 68(2), 217-233) verified that model specifically
# against heavy petroleum cuts (LVGO/HVGO/GDAR/ABVB), exactly this use case.
_RIAZI_VERA_GASES = {
    "METHANE": dict(tc_k=190.56, pc_bar=45.99, v1l=52.0,   delta1=11.62,
                     omega=0.0115, alpha=0.94, frol_fixed=None),
    "ETHANE":  dict(tc_k=305.32, pc_bar=78.83, v1l=45.7,   delta1=12.4,
                     omega=0.0995, alpha=1.3,  frol_fixed=None),
    "CO2":     dict(tc_k=304.21, pc_bar=78.83, v1l=37.27,  delta1=14.56,
                     omega=0.225, alpha=1.10, frol_fixed=None),
}


def _rk_vapor_z(a: float, b: float) -> float:
    """Largest real (vapour) root of the Redlich-Kwong compressibility cubic
    Z**3 - Z**2 + (A - B - B**2)*Z - A*B = 0, by Newton-Raphson from the
    ideal-gas seed Z=1 (robust here since Tr is large / Pr modest for H2 in
    this application — the vapour root dominates)."""
    z = 1.0
    for _ in range(100):
        f = z ** 3 - z ** 2 + (a - b - b ** 2) * z - a * b
        fp = 3 * z ** 2 - 2 * z + (a - b - b ** 2)
        if fp == 0:
            break
        z_new = z - f / fp
        z_new = max(z_new, b + 1e-9)
        if abs(z_new - z) < 1e-12:
            z = z_new
            break
        z = z_new
    return z


# Grayson & Streed (1963) Curl-Pitzer coefficients for H2's pure-liquid
# fugacity coefficient (eq 6), as refit by Torres, de Hemptinne & Machin's
# "Augmented Grayson-Streed" (AGS, 2013, OGST 68(2), 217-233, Table 6) — only
# A0/A1 (the dominant temperature-dependence terms) were refit, against
# hydrogen solubility data spanning n-heptane through aromatics/heavy
# petroleum cuts; the rest are carried over unchanged from the original
# Grayson-Streed (1963) H2 row.  Per both papers, only the log(phi^(0)) term
# is used for H2 (the acentric-factor correction term is dropped entirely).
_AGS_H2_COEFFS = dict(A0=1.67380, A1=6.93898, A2=-0.02110, A3=0.00011,
                       A4=0.0, A5=0.008585, A6=0.0, A7=0.0, A8=0.0, A9=0.0)
_AGS_H2_TC_K = 33.4
_AGS_H2_PC_BAR = 13.155      # 1 315 524 Pa
_AGS_H2_V1_CM3MOL = 31.0     # 0.0310 m3/kmol
_AGS_H2_DELTA1 = 6.648       # (J/cm3)^0.5


def _ags_k_h2(p_psia, temp_f, solvent_mw) -> float | None:
    """Augmented Grayson-Streed (AGS) K-value for hydrogen dissolved in a
    heavy hydrocarbon liquid (Torres, de Hemptinne & Machin, 2013, OGST
    68(2), 217-233).  Standard Grayson-Streed structure,
    K1 = phi1^(L*) * gamma1 / phi1^V, with two changes from the original
    1963 correlation that this paper showed cut the AAD on heavy-cut H2
    solubility roughly in half (55% -> 30%, validated against real
    LVGO/HVGO/GDAR/ABVB hydroprocessing feeds):
      * gamma1 includes a Flory entropic term in addition to Hildebrand's
        regular-solution enthalpic term, evaluated at infinite dilution
        (consistent with this implementation's Henry's-law-style K, and
        with how the paper itself reports gamma1^inf in Tables 4/5);
      * phi1^(L*)'s temperature-dependence coefficients (A0, A1) are
        refit for H2 specifically (Table 6) instead of the original 1963
        values, which the paper found systematically under/over-predict
        for n-C16+ and aromatics.
    The solvent is characterized only by its SCN (single-carbon-number)
    average MW (Riazi & Vera, 2005, eq 17/18) for delta2/v2 — the same
    lumped-solvent approach AGS itself found most reliable (Table 14).
    phi1^V uses the Redlich-Kwong EOS for pure H2 vapour (eq 11-17),
    neglecting other vapour-phase species' effect on H2's own fugacity
    coefficient — a minor simplification at the modest reduced pressures
    typical of refinery off-gas.  Returns None if inputs are invalid.
    """
    if not p_psia or p_psia <= 0 or temp_f is None or not solvent_mw or solvent_mw <= 0:
        return None
    t_k = (temp_f - 32.0) * 5.0 / 9.0 + 273.15
    p_bar = p_psia / 14.5038
    if t_k <= 0:
        return None

    m = solvent_mw
    delta2 = 17.5913 - math.exp(3.0076 - 0.549097 * m ** 0.3)          # eq 17
    v2 = m / (1.05 - math.exp(3.80258 - 3.12287 * m ** 0.1))           # eq 18
    if v2 <= 0:
        return None

    v1 = _AGS_H2_V1_CM3MOL
    delta1 = _AGS_H2_DELTA1
    r = 8.314
    # gamma1 at infinite dilution: Hildebrand (enthalpic) + Flory (entropic)
    ln_gamma1 = (v1 * (delta1 - delta2) ** 2 / (r * 298.15)
                 + math.log(v1 / v2) + 1.0 - v1 / v2)
    gamma1 = math.exp(ln_gamma1)

    tr = t_k / _AGS_H2_TC_K
    pr = p_bar / _AGS_H2_PC_BAR
    if tr <= 0 or pr <= 0:
        return None
    c = _AGS_H2_COEFFS
    log_phi0 = (c["A0"] + c["A1"] / tr + c["A2"] * tr + c["A3"] * tr ** 2
                + c["A4"] * tr ** 3
                + (c["A5"] + c["A6"] * tr + c["A7"] * tr ** 2) * pr
                + (c["A8"] + c["A9"] * tr) * pr ** 2
                - math.log10(pr))
    phi1_lstar = 10.0 ** log_phi0

    a_rk = 0.42748 * pr / tr ** 2.5
    b_rk = 0.08664 * pr / tr
    z = _rk_vapor_z(a_rk, b_rk)
    if z <= b_rk:
        return None
    ln_phi_v = (z - 1.0) - math.log(z - b_rk) - (a_rk / b_rk) * math.log(1.0 + b_rk / z)
    phi1_v = math.exp(ln_phi_v)
    if phi1_v <= 0:
        return None

    k = phi1_lstar * gamma1 / phi1_v
    return k if (math.isfinite(k) and k > 0) else None


def _riazi_vera_k(name_upper, p_psia, temp_f, solvent_mw) -> float | None:
    """Riazi & Vera (2005) regular-solution-theory K-value for a light gas
    dissolved in a heavy hydrocarbon liquid, used in place of Wilson's
    corresponding-states shape for the specific species this paper targets
    (H2, CH4, C2H6, CO2) — Wilson is known to badly misjudge H2 (Tc=33 K)
    volatility once it is partly dissolved in a hot heavy liquid, exactly the
    failure mode this paper was built to address.

    From the paper's eq 1, x1 = phi1V·P1/(gamma1·f1L) with P1 = y1·P, so
    K = y1/x1 = gamma1·f1L/(phi1V·P).  gamma1 (eq 2) uses the "SCN"
    single-pseudo-component solvent model (eq 17/18, requiring only the
    solvent's average MW) and approximates delta_mix ~= delta2 (the
    solute's own volume fraction in the liquid is neglected — reasonable at
    the low-to-moderate solubilities this model targets).  Returns None if
    `name_upper` isn't tabulated or inputs are invalid.
    """
    g = _RIAZI_VERA_GASES.get(name_upper)
    if (g is None or not p_psia or p_psia <= 0 or temp_f is None
            or not solvent_mw or solvent_mw <= 0):
        return None
    t_k = (temp_f - 32.0) * 5.0 / 9.0 + 273.15
    p_bar = p_psia / 14.5038
    if t_k <= 0:
        return None

    m = solvent_mw
    delta2 = 17.5913 - math.exp(3.0076 - 0.54907 * m ** 0.3)   # eq 17
    delta1 = g["alpha"] * g["delta1"]
    v1l = g["v1l"]
    r = 8.314
    gamma1 = math.exp(v1l * (delta1 - delta2) ** 2 / (r * 298.15))   # eq 2

    tc_k, pc_bar = g["tc_k"], g["pc_bar"]
    tr = t_k / tc_k
    if tr <= 0:
        return None
    fr_ol = g["frol_fixed"]
    if fr_ol is None:
        fr_ol = math.exp(7.902 - 8.19643 / tr - 3.08 * math.log(tr))   # eq 6
    f1l = fr_ol * pc_bar * math.exp(v1l * (p_bar - 1.013) / (r * t_k))  # eq 5

    pr = p_bar / pc_bar
    omega = g["omega"]
    expo = (pr / tr) * ((0.083 - 0.422 * tr ** -1.6)
                         + omega * (0.139 - 0.122 * tr ** -4.2))        # eq 7
    expo = max(-50.0, min(50.0, expo))
    phi1v = math.exp(expo)

    k = gamma1 * f1l / (phi1v * p_bar)
    return k if (math.isfinite(k) and k > 0) else None


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


_PF_NAME_RE = re.compile(r"^PF(\d+)A(\d+)D(?:_\d+)?$")


# Direct Tb/SG/MW curve fit for Tc/Pc, Watson K 6.76-32.42 — same unified
# regression as comp_constants.py's estimate_pseudo_props (ln(Tc_R)/ln(Pc_psia)
# by OLS against Tb(R), SG, MW; fit once against 220 PRO/II ground-truth
# pseudo-fractions spanning that whole K range; max error 0.23%/1.55%).
# Inlined here (not imported from comp_constants) so this decode path keeps
# its no-pandas-dependency property — comp_constants imports pandas/openpyxl
# at module level for spreadsheet I/O, which isn't needed for the bare
# correlation.
_CF_K_MIN = 6.7607782905215785
_CF_K_MAX = 32.42
_CF_TC_COEF = (5.007883101013148, 0.003450517505203951,
               -5.350561845008301e-08, -1.074024144669105e-09,
               1.6490962015029442, -0.5550533538284068,
               -0.0015136220956506788, 6.726083603733643e-07,
               -0.01062975296484001, -1.8505676528363376e-06,
               1.1442752460598282e-05, -0.002262849729260199,
               -1.5958487089281692e-09, 0.0015542330963586107)
_CF_PC_COEF = (5.701626621482552, -0.0073214380294048045,
               2.2394905038568905e-05, -1.4116473210398495e-08,
               9.73621911814537, -3.3041596315868413,
               -0.00870625333062654, 3.7745792200831124e-06,
               -0.055143811415915074, -1.7498645134846417e-05,
               4.5674217570408565e-05, -0.011989311917111075,
               4.171268477379032e-09, 0.008258835869442238)


def _cf_ln_poly(tb_r, sg, mw, c):
    return (c[0] + c[1]*tb_r + c[2]*tb_r**2 + c[3]*tb_r**3
            + c[4]*sg + c[5]*sg**2 + c[6]*tb_r*sg + c[7]*tb_r**2*sg
            + c[8]*mw + c[9]*mw**2 + c[10]*tb_r*mw + c[11]*sg*mw
            + c[12]*tb_r**2*mw + c[13]*sg**2*mw)


# PRO/II SIMSCI/TWU acentric factor (generalized Frost-Kalkwarf-Thodos vapor
# pressure correlation), back-solved at the NBP boundary condition — same
# omega companion comp_constants.py pairs with the curve fit above.
_FK_A = (10.2005, -10.6317, -5.58058, 2.09167, -2.09167, -1.70214, 0.4312)


def _fk_omega(tb_r, tc_r, pc_psia):
    if tc_r <= tb_r or pc_psia <= 0:
        return None
    a7 = _FK_A[6]
    tr_b, pr_b = tb_r / tc_r, 14.696 / pc_psia
    f0_b = _FK_A[0] + _FK_A[1] / tr_b + _FK_A[2] * math.log(tr_b)
    f1_b = _FK_A[3] + _FK_A[4] / tr_b + _FK_A[5] * math.log(tr_b)
    if f1_b == 0:
        return None
    omega_cap = (math.log(pr_b) - a7 * pr_b / tr_b ** 2 - f0_b) / f1_b

    tr = 0.7
    f0 = _FK_A[0] + _FK_A[1] / tr + _FK_A[2] * math.log(tr)
    f1 = _FK_A[3] + _FK_A[4] / tr + _FK_A[5] * math.log(tr)
    rhs = f0 + omega_cap * f1
    pr = math.exp(rhs)
    for _ in range(50):
        g = math.log(pr) - a7 * pr / tr ** 2 - rhs
        dg = 1.0 / pr - a7 / tr ** 2
        step = g / dg
        pr -= step
        if abs(step) < 1e-12:
            break
    if pr <= 0:
        return None
    return -math.log10(pr) - 1.0


def _decode_pf_pseudo(name_upper: str, mw: float | None = None) -> dict | None:
    """Petroleum-fraction pseudo-component names emitted by the source PRO/II
    model encode their own normal boiling point and API gravity directly,
    e.g. 'PF736A30D_6' = NBP 736°F, API 30 (the trailing '_<n>' is just a
    cut-set index and varies by stream/column, which is why most of these
    don't have an exact-name match in the COMP_CONSTANTS sheet — that sheet
    only enumerates one arbitrarily-chosen cut set per NBP/API pair). Decode
    NBP/SG straight from the name instead of relying on an exact lookup match.

    When the caller has a per-component MW (from the stream's mass/mole flow
    ratio) and the decoded Watson K falls inside the curve fit's verified
    range, Tc/Pc come from the direct Tb/SG/MW curve fit above. Otherwise —
    MW unavailable, or K outside 6.76-32.42 — falls back to the Lee-Kesler
    (1975) Tc/Pc + Edmister (1958) omega correlation, which only needs
    NBP/SG and has no fitted-range restriction."""
    m = _PF_NAME_RE.match(name_upper)
    if not m:
        return None
    nbp_f, api = float(m.group(1)), float(m.group(2))
    sg = 141.5 / (131.5 + api)
    tb_r = nbp_f + 459.67
    if tb_r <= 0 or sg <= 0:
        return None

    if mw:
        k_w = tb_r ** (1 / 3) / sg
        if _CF_K_MIN <= k_w <= _CF_K_MAX:
            tc_r = math.exp(_cf_ln_poly(tb_r, sg, mw, _CF_TC_COEF))
            pc_psia = math.exp(_cf_ln_poly(tb_r, sg, mw, _CF_PC_COEF))
            omega = _fk_omega(tb_r, tc_r, pc_psia)
            return {"tc_f": tc_r - 459.67, "pc_psia": pc_psia, "omega": omega}

    # Lee-Kesler (1975) Tc/Pc + Edmister (1958) omega — same correlation
    # comp_constants.py falls back to (via raw Twu) outside the curve fit's
    # verified range, inlined here so this decode path has no dependency on
    # pandas (comp_constants imports it for spreadsheet I/O, which isn't
    # needed for the bare correlation).
    tc_r = (341.7 + 811.1 * sg
            + (0.4244 + 0.1174 * sg) * tb_r
            + (0.4669 - 3.2623 * sg) * 1e5 / tb_r)
    ln_pc = (8.3634
             - 0.0566 / sg
             - (0.24244 + 2.2898 / sg + 0.11857 / sg ** 2) * 1e-3 * tb_r
             + (1.4685 + 3.648 / sg + 0.47227 / sg ** 2) * 1e-7 * tb_r ** 2
             - (0.42019 + 1.6977 / sg ** 2) * 1e-10 * tb_r ** 3)
    pc_psia = math.exp(ln_pc)
    if pc_psia <= 14.696 or tc_r <= tb_r:
        return None
    theta = tc_r / tb_r - 1.0
    omega = (3.0 / 7.0) * math.log10(pc_psia / 14.696) / theta - 1.0 if theta > 0 else None
    return {"tc_f": tc_r - 459.67, "pc_psia": pc_psia, "omega": omega}


_KPA_TO_PSIA = 0.14503774

# Short PRO/II component IDs (as they appear in stream composition tables)
# that don't match the library's full chemical name verbatim.
_PROII_ALIAS = {
    "H2": "HYDROGEN", "N2": "NITROGEN", "O2": "OXYGEN", "CO": "CARBON MONOXIDE",
    "CO2": "CARBON DIOXIDE", "H2S": "HYDROGEN SULFIDE", "H2O": "WATER",
    "NH3": "AMMONIA", "HCL": "HYDROGEN CHLORIDE",
    "NC4": "N-BUTANE", "IC4": "ISOBUTANE",
    "NC5": "N-PENTANE", "IC5": "ISOPENTANE",
    "NC6": "N-HEXANE", "NC7": "N-HEPTANE", "NC8": "N-OCTANE",
    "NC9": "N-NONANE", "NC10": "N-DECANE",
}

def _proii_cm(n_up: str) -> dict | None:
    """Real Tc/Pc/omega for a named (non-pseudo) component from the embedded
    PRO/II library, converted to the tc_f/pc_psia/omega format the rest of
    _build_feed expects. Returns None if the component isn't in the library
    or is missing Tc/Pc (e.g. some lumps only carry MW/NBP)."""
    p = PROII_PROPS.get(n_up) or PROII_PROPS.get(_PROII_ALIAS.get(n_up, ""))
    if not p:
        return None
    tc_k, pc_kpa = p.get("TC"), p.get("PC")
    if tc_k is None or pc_kpa is None:
        return None
    return {"tc_f": tc_k * 9.0 / 5.0 - 459.67,
            "pc_psia": pc_kpa * _KPA_TO_PSIA,
            "omega": p.get("ACENTRIC")}


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
        # criticals — real PRO/II library values take priority for named
        # (non-pseudo) components; the PF<nbp>A<api>D cut-set pseudos keep
        # the predicted direct-curve-fit/Lee-Kesler-Edmister path (no exact
        # library match is meaningful for those — see _decode_pf_pseudo's
        # docstring).
        n_up = str(n).strip().upper()
        cm = None
        if not _PF_NAME_RE.match(n_up):
            cm = _proii_cm(n_up)
        if cm is None:
            cm = (const_map or {}).get(n_up)
        if cm is None or cm.get("tc_f") is None:
            cm = _decode_pf_pseudo(n_up, mwi)
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

    p_ref = sp.pres_psia or 0.0
    if p_ref <= 0:
        return None

    if use_yx:
        # true K_ref already material-balance correct; gentle λ nudge to β₀
        lam = _solve_lambda(z, k_raw, beta0)
        k_ref = [min(1e10, max(1e-10, lam * kr)) for kr in k_raw]
        basis = "yx"
    else:
        # The z/x proxy degenerates to ~1 for every component whenever β₀ is
        # extremely close to 0 or 1 (x_i ≈ z_i), losing all relative-volatility
        # information — every component then flips from all-liquid to
        # all-vapor together as pressure falls, instead of the light ends
        # flashing off preferentially.  Where fundamental criticals (Tc/Pc/ω)
        # are available — typically the light/volatile species — substitute
        # the physically-grounded full Wilson correlation's SHAPE for those
        # components' raw ratio (still globally λ-scaled together with the
        # rest, so Rachford-Rice still reproduces the HMB's exact β₀ — the
        # absolute Wilson K is not trustworthy enough quantitatively, but its
        # relative ordering across components is far better than the flat
        # proxy's).
        wilson_k = [_wilson_k_full(tc, om, pc, p_ref, sp.temp_f)
                    for tc, om, pc in zip(tc_r, omega, pc_l)]
        used_wilson = any(wk is not None for wk in wilson_k)
        if used_wilson:
            k_raw = [wk if wk is not None else kr for wk, kr in zip(wilson_k, k_raw)]

        # For the specific light gases Riazi & Vera (2005) / the Augmented
        # Grayson-Streed model (Torres et al., 2013) target (H2, CH4, C2H6,
        # CO2), their solubility-parameter K-values are more reliable than
        # Wilson's generic corresponding-states shape — Wilson is known to
        # misjudge H2 once dissolved in a hot heavy liquid, the exact
        # failure mode these correlations were built for.  The "solvent" is
        # the rest of the feed's z-weighted average MW (this component
        # excluded), per the SCN approach both papers settled on as most
        # reliable for lumped heavy-cut characterization.
        total_mw_zw = sum(zi * mi for zi, mi in zip(z, mw))
        used_riazi = False
        k_special: dict[int, tuple[str, float]] = {}
        for idx, nm in enumerate(names):
            nm_up = str(nm).strip().upper()
            if nm_up not in _RIAZI_VERA_GASES and nm_up != "H2":
                continue
            rest_z = 1.0 - z[idx]
            if rest_z <= 1e-9:
                continue
            solvent_mw = (total_mw_zw - z[idx] * mw[idx]) / rest_z
            if nm_up == "H2":
                rv_k = _ags_k_h2(p_ref, sp.temp_f, solvent_mw)
            else:
                rv_k = _riazi_vera_k(nm_up, p_ref, sp.temp_f, solvent_mw)
            if rv_k is not None:
                k_raw[idx] = rv_k
                used_riazi = True
                k_special[idx] = (nm_up, solvent_mw)

        if used_riazi and used_wilson:
            basis = "zx-lambda+wilson+riazivera+ags"
        elif used_riazi:
            basis = "zx-lambda+riazivera+ags"
        elif used_wilson:
            basis = "zx-lambda+wilson"
        else:
            basis = "zx-lambda"
        lam = _solve_lambda(z, k_raw, beta0)
        k_ref = [min(1e10, max(1e-10, lam * kr)) for kr in k_raw]

    if use_yx:
        k_special = {}

    total_mw = sum(zi * mi for zi, mi in zip(z, mw))
    total_mass = (sp.vap_mass or 0.0) + (sp.liq_mass or 0.0)
    if total_mass <= 0:
        total_mass = sp.total_mass or 0.0
    total_moles = (total_mass / total_mw) if total_mw > 0 else 0.0

    feed = FlashFeed(names=names, z=z, k_ref=k_ref, mw=mw,
                     total_mass_lbhr=total_mass, p_ref_psia=p_ref,
                     temp_f=sp.temp_f, total_mw=total_mw,
                     total_moles_lbmolhr=total_moles,
                     tc_r=tc_r, omega=omega, pc_psia=pc_l,
                     vap_h_ref=getattr(sp, "vap_sp_enthalpy", None),
                     liq_h_ref=getattr(sp, "liq_sp_enthalpy", None),
                     vap_cp=getattr(sp, "vap_cp", None),
                     liq_cp=getattr(sp, "liq_cp", None),
                     k_basis=basis,
                     k_special=(k_special or None), k_lambda=lam)
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


# ══════════════════════════════════════════════════════════════════════════
# ║  SECTION: CONTROL VALVE SIZING & DATASHEET  (IEC 60534)
# ══════════════════════════════════════════════════════════════════════════
# Sizing (IEC 60534-2-1 liquid/gas/two-phase), cavitation (ISA RP75.23),
# seat leakage (FCI 70-2 / IEC 60534-4) and % travel are ported from the
# companion CV/cv.py tool so the two stay numerically identical.  Extended
# here with (a) IEC 60534-8-3 (aerodynamic) and 60534-8-4 (hydrodynamic)
# sound-pressure-level (dB(A)) prediction, (b) a generic Rated-Cv selection
# library, and (c) a Min/Norm/Max vendor-style datasheet — all driven
# directly from the control-valve stations of the hydraulic march.
#
# Two operating modes, auto-detected per valve:
#   * Adequacy check   — a Rated Cv is supplied (existing valve): the tool
#                        reports required Cv / % travel / cavitation / noise
#                        and PASS/FAIL verdicts against that valve; on a noise
#                        exceedance it flags and recommends (never silently
#                        re-sizes the user's real valve).
#   * Sizing/selection — no Rated Cv supplied (new valve): the tool selects a
#                        Rated Cv from the generic library so that required Cv,
#                        the travel window AND the dB(A) limit are all met
#                        where physically possible.
#
# US-customary Cv basis (gpm·√SG/√psi liquid; lb/h, psia, lb/ft³ gas) per
# datasheet convention, independent of the workbook's unit system.

_CV_PSI_TO_PA  = 6894.757293
_CV_R_KMOL     = 8314.462618          # J/(kmol*K)  (module R_UNIV is per-mol)
_CV_KGS_TO_LBH = 7936.6414            # kg/s -> lb/h
_CV_M3S_TO_GPM = 15850.323            # m3/s -> US gpm
_CV_ATM_PA     = 101_325.0
_CV_PREF_PA    = 2e-5                 # reference sound pressure (20 uPa)

# Class II/III/IV seat leakage as % of rated Cv (FCI 70-2 / IEC 60534-4)
_CV_CLASS_PCT_CV = {"II": 0.5, "III": 0.1, "IV": 0.01}
# Class VI — max bubbles/min of air at standard test dP, by port dia (in)
_CV_CLASS_VI_TABLE = [
    (1.0, 0.15), (1.5, 0.30), (2.0, 0.45), (2.5, 0.60),
    (3.0, 0.90), (4.0, 1.70), (6.0, 4.00), (8.0, 6.75),
]

# Typical body-style factors (IEC 60534-2-1 Table 2 "typical values"):
#   fl  = liquid pressure-recovery factor FL (flow-to-open, mid-travel)
#   xt  = terminal pressure-drop ratio xT (gas choke)
#   fd  = valve-style modifier Fd (jet diameter, for -8-3/-8-4 noise)
#   cd  = generic full-open rated-Cv coefficient  Cv100 ~= cd * NPS**2
# These are indicative; a real selection confirms them against vendor data.
_CV_BODY = {
    "GLOBE":     dict(label="Globe",              fl=0.90, xt=0.72, fd=0.46, cd=16.0),
    "ANGLE":     dict(label="Angle",              fl=0.85, xt=0.72, fd=0.44, cd=16.0),
    "BALL":      dict(label="Ball (segmented)",   fl=0.66, xt=0.30, fd=0.98, cd=32.0),
    "BUTTERFLY": dict(label="Butterfly (60 deg)", fl=0.68, xt=0.38, fd=0.57, cd=28.0),
    "ECCENTRIC": dict(label="Eccentric rotary",   fl=0.85, xt=0.61, fd=0.42, cd=20.0),
}
_CV_BODY_ALIASES = {
    "GLOBE": "GLOBE", "ANGLE": "ANGLE", "BALL": "BALL",
    "SEGMENTED BALL": "BALL", "V-BALL": "BALL", "VBALL": "BALL",
    "BUTTERFLY": "BUTTERFLY", "BFLY": "BUTTERFLY",
    "ECCENTRIC": "ECCENTRIC", "ECCENTRIC ROTARY": "ECCENTRIC",
    "ROTARY": "ECCENTRIC", "CAMFLEX": "ECCENTRIC",
}
# Standard reduced-trim rated-Cv ladder (indicative ISA-style steps).  The
# selection picks the smallest rung that satisfies capacity/travel/noise,
# capped by the body-style full-open maximum for the valve NPS.
_CV_RATED_LADDER = [
    0.5, 0.8, 1.2, 2.0, 3.0, 4.0, 5.0, 6.0, 8.0, 10.0, 12.0, 15.0, 18.0,
    22.0, 28.0, 36.0, 46.0, 60.0, 75.0, 95.0, 120.0, 150.0, 195.0, 250.0,
    320.0, 400.0, 520.0, 650.0, 850.0, 1100.0,
]

_CV_TRIM_CHARS = {
    "LINEAR": "Linear", "EQUAL%": "Equal Percentage",
    "EQUAL PERCENTAGE": "Equal Percentage", "EQ%": "Equal Percentage",
    "EQUALPCT": "Equal Percentage", "QUICK": "Quick Opening",
    "QUICK-OPEN": "Quick Opening", "QUICK OPENING": "Quick Opening",
    "QO": "Quick Opening",
}


def _cv_body_key(name: str | None) -> str:
    k = (name or "GLOBE").strip().upper()
    return _CV_BODY_ALIASES.get(k, "GLOBE")


def _cv_trim_name(name: str | None) -> str:
    return _CV_TRIM_CHARS.get((name or "").strip().upper(), "Linear")


# ── Core IEC 60534-2-1 sizing (ported verbatim from CV/cv.py) ──────────────
def _cv_ff_factor(pv_pa, pc_pa) -> float:
    """Liquid critical-pressure-ratio factor FF = 0.96 − 0.28·√(Pv/Pc)."""
    if not pv_pa or not pc_pa or pc_pa <= 0:
        return 0.96
    return max(0.0, min(1.0, 0.96 - 0.28 * math.sqrt(max(0.0, pv_pa) / pc_pa)))


def _cv_sigma_index(p1_pa, p2_pa, pv_pa):
    """Service cavitation index σ = (P1 − Pv)/(P1 − P2)  (ISA RP75.23)."""
    dp = p1_pa - p2_pa
    if dp <= 0:
        return None
    return (p1_pa - (pv_pa or 0.0)) / dp


def _cv_liquid_dp_choked(fl, p1_pa, pv_pa, ff) -> float:
    """Choked ΔP for liquids: FL²·(P1 − FF·Pv)."""
    fl = fl or 0.9
    return (fl ** 2) * (p1_pa - ff * (pv_pa or 0.0))


def _cv_required_liquid(w_kgs, rho, dp_eff_pa) -> float:
    """Required Cv (US basis) — liquid, IEC 60534-2-1: Cv = Q[gpm]·√(SG/ΔP[psi])."""
    if w_kgs <= 0 or rho <= 0 or dp_eff_pa <= 0:
        return 0.0
    q_gpm = (w_kgs / rho) * _CV_M3S_TO_GPM
    sg = rho / 999.0
    dp_psi = dp_eff_pa / _CV_PSI_TO_PA
    return q_gpm * math.sqrt(sg / dp_psi)


def _cv_gas_xy(p1_pa, p2_pa, k, xt):
    """Gas ratio x, choked limit xT_eff = Fk·xT, effective x, expansion Y,
    and choked flag — IEC 60534-2-1 §5.5/5.6."""
    if p1_pa <= 0:
        return 0.0, 0.0, 0.0, 1.0, False
    fk = max(0.1, (k or 1.4) / 1.4)
    x = max(0.0, (p1_pa - p2_pa) / p1_pa)
    x_choked = max(1e-6, fk * (xt or 0.7))
    x_eff = min(x, x_choked)
    y = max(2.0 / 3.0, 1.0 - x_eff / (3.0 * x_choked))
    return x, x_choked, x_eff, y, (x >= x_choked)


def _cv_required_gas(w_kgs, p1_pa, rho1, x_eff, y) -> float:
    """Required Cv (US basis) — gas/vapor, mass-flow form:
    Cv = W[lb/h] / (63.3·Y·√(x_eff·P1[psia]·ρ1[lb/ft³]))."""
    if w_kgs <= 0 or p1_pa <= 0 or rho1 <= 0 or x_eff <= 0 or y <= 0:
        return 0.0
    w_lbh = w_kgs * _CV_KGS_TO_LBH
    p1_psia = p1_pa / _CV_PSI_TO_PA
    rho1_lbft3 = rho1 / 16.018463
    return w_lbh / (63.3 * y * math.sqrt(x_eff * p1_psia * rho1_lbft3))


def _cv_required_two_phase(w_kgs, x_quality, rho_l, rho_v, dp_eff_pa):
    """Required Cv — flashing/two-phase, homogeneous-mixture estimate (mixes
    liquid+vapor densities at the mass quality, then the liquid equation).
    For severe flashing/cavitating service confirm with a vendor method."""
    if w_kgs <= 0 or dp_eff_pa <= 0:
        return 0.0, None
    rho_l = rho_l or 999.0
    rho_v = rho_v or 1.2
    x_quality = max(0.0, min(1.0, x_quality or 0.0))
    rho_mix = 1.0 / (x_quality / rho_v + (1.0 - x_quality) / rho_l)
    return _cv_required_liquid(w_kgs, rho_mix, dp_eff_pa), rho_mix


def _cv_travel_pct(cv_req, cv100, char, rangeability=50.0):
    """Estimated % travel from required Cv, rated Cv100 & inherent trim
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
    else:
        travel = 100.0 * ratio
    return max(0.0, min(100.0, travel))


def _cv_beta_ratio(bore_m, pipe_m):
    if not bore_m or not pipe_m or pipe_m <= 0:
        return None
    return bore_m / pipe_m


def _cv_seat_leakage(klass, cv100, dp_pa, seat_dia_mm) -> dict:
    """Seat leakage estimate by class (FCI 70-2 / IEC 60534-4)."""
    k = (klass or "IV").strip().upper()
    if k in _CV_CLASS_PCT_CV:
        pct = _CV_CLASS_PCT_CV[k]
        leak_cv = cv100 * pct / 100.0 if cv100 else None
        return {"class": k, "pct_cv": pct, "leak_cv": leak_cv,
                "desc": f"Class {k}: <= {pct}% of rated Cv (FCI 70-2 / IEC 60534-4)"}
    if k == "V":
        dp_bar = max(0.0, dp_pa) / 1e5
        leak = 0.18 * (seat_dia_mm or 0.0) * math.sqrt(dp_bar)
        return {"class": "V", "leak_ml_min": leak,
                "desc": "Class V: indicative liquid leakage (order-of-magnitude) "
                        "- confirm against FCI 70-2 / certified test data"}
    if k == "VI":
        seat_in = (seat_dia_mm or 0.0) / 25.4
        bubbles = _CV_CLASS_VI_TABLE[-1][1]
        for dia, b in _CV_CLASS_VI_TABLE:
            if seat_in <= dia:
                bubbles = b
                break
        return {"class": "VI", "bubbles_per_min": bubbles,
                "desc": "Class VI: max bubbles/min of air by port size "
                        "(FCI 70-2 Table 2 - indicative)"}
    return {"class": k, "desc": "Unrecognized leakage class - see IEC 60534-4"}


def _cv_sonic_velocity_gas(k, z, t_k, mw):
    if not t_k or not mw:
        return None
    return math.sqrt((k or 1.4) * (z or 1.0) * _CV_R_KMOL * t_k / mw)


# ── IEC 60534-8-3 / -8-4  external dB(A) noise prediction ──────────────────
# Engineering implementation of the standards' method chain.  Aerodynamic
# (-8-3): mechanical stream power -> acoustic power via a regime-dependent
# acoustical efficiency (η) -> internal sound pressure at the pipe wall ->
# external SPL at 1 m through the pipe-wall transmission loss (TL).
# Hydrodynamic (-8-4): turbulent baseline SPL from valve ΔP and style, with a
# cavitation increment once the service σ falls below the incipient value.
# Coefficient tables are simplified relative to the full standards; results
# are screening-grade dB(A) for adequacy against a project limit (e.g. 85
# dB(A)), not a substitute for a vendor's certified acoustic prediction.
_CV_SPEED_SOUND_PIPE = 5000.0    # m/s, longitudinal wave speed in steel wall
_CV_RHO_STEEL        = 7800.0    # kg/m³

def _cv_pipe_wall_thk_m(pipe_id_m):
    """Approximate carbon-steel (Sch-40) wall thickness from ID, m."""
    if not pipe_id_m or pipe_id_m <= 0:
        return 0.005
    d_in = pipe_id_m / 0.0254
    # Sch-40 wall grows ~ with NPS; bounded to a sane 3–15 mm band.
    return max(0.003, min(0.015, 0.0033 + 0.0015 * d_in))


def _cv_pipe_tl_db(pipe_id_m, fp_hz, c2):
    """Pipe-wall transmission loss TL [dB, negative] — simplified IEC 60534-8-3
    Annex form: coincidence-limited mass-law loss, referenced to the ring
    frequency fr of the pipe.  Screening-grade."""
    di = max(0.01, pipe_id_m)
    tp = _cv_pipe_wall_thk_m(pipe_id_m)
    fr = _CV_SPEED_SOUND_PIPE / (math.pi * di)                 # ring frequency
    fo = 0.25 * fr                                             # first coincidence
    # mass-law surface density term, normalised
    gy = (_CV_RHO_STEEL * tp) * fp_hz / ((c2 or 340.0) * 1.2 * 101325.0 / 1e5)
    tl = -(10.0 + 10.0 * math.log10(max(1e-6, gy))
           - 10.0 * math.log10(1.0 + (fo / max(1.0, fp_hz)) ** 1.5))
    return max(-75.0, min(-25.0, tl))                          # bounded band


def _cv_noise_aero(w_kgs, p1_pa, p2_pa, t1_k, mw, z, k, xt, fl, fd,
                    pipe_id_m, rho2, c2):
    """IEC 60534-8-3-style external A-weighted SPL at 1 m for a gas/vapor
    valve.  Returns (Lpe_dBA, detail_dict).

    Chain: mechanical stream power Wm=½·ṁ·Uvc² → acoustic power Wa=η·Wm →
    sound-power level Lw → internal pipe-wall SPL Lpi (IEC 8-3 relation, pipe
    ID in mm) → external SPL at 1 m through the wall transmission loss TL."""
    if (not w_kgs or w_kgs <= 0 or not p1_pa or p1_pa <= 0 or not pipe_id_m
            or pipe_id_m <= 0 or not rho2 or rho2 <= 0):
        return None, {}
    k = max(1.001, k or 1.4)
    x = max(1e-6, (p1_pa - p2_pa) / p1_pa)
    fk = k / 1.4
    x_choked = max(1e-6, fk * (xt or 0.7))
    c1 = c2 or _cv_sonic_velocity_gas(k, z, t1_k, mw) or 340.0
    if x < x_choked:                                   # subsonic
        uvc = min(c1, c1 * math.sqrt(max(0.0, x / x_choked)))
        regime = "subsonic"
    else:                                              # choked
        uvc = c1
        regime = "choked"
    wm = 0.5 * w_kgs * uvc ** 2                         # mechanical power, W
    mach_vc = uvc / c1 if c1 else 0.0
    # acoustical efficiency η (IEC 8-3 regime form): ~1e-4·M³ subsonic,
    # saturating near ~1e-3 when choked; Fd nudges the jet efficiency.
    eta = 1.0e-4 * (mach_vc ** 3)
    if regime == "choked":
        eta = max(eta, 3.0e-4 * (fd or 0.5) / 0.5)
    eta = min(eta, 3.0e-3)
    wa = max(1e-30, eta * wm)                           # acoustic power, W
    lw = 10.0 * math.log10(wa / 1e-12)                 # sound-power level, dB
    di_mm = pipe_id_m * 1000.0
    # internal pipe SPL — IEC 60534-8-3 relation (di in mm, Wa in W):
    #   Lpi = 10·log10( 3.2e9 · Wa · ρ2 · c2 / di² )   [dB re 2e-5 Pa]
    lpi = 10.0 * math.log10(
        max(1e-30, 3.2e9 * wa * rho2 * (c2 or c1) / (di_mm ** 2)))
    fp = max(1.0, 0.2 * uvc / max(1e-3, pipe_id_m))    # peak frequency, Hz
    tl = _cv_pipe_tl_db(pipe_id_m, fp, c2 or c1)
    lpe = lpi + tl - 1.0                               # external at 1 m, dB(A)
    lpe = max(20.0, lpe)
    return lpe, {"regime": regime, "uvc": uvc, "mach_vc": mach_vc,
                 "eta": eta, "wa": wa, "lw": lw, "lpi": lpi, "tl": tl, "fp": fp}


def _cv_noise_hydro(w_kgs, p1_pa, p2_pa, pv_pa, fl, rho_l, pipe_id_m, sigma):
    """IEC 60534-8-4-style external A-weighted SPL at 1 m for liquid / flashing
    service, with a cavitation increment once σ drops below the incipient
    value.  Returns (Lpe_dBA, detail_dict).  Screening-grade."""
    if (not w_kgs or w_kgs <= 0 or not p1_pa or not pipe_id_m
            or pipe_id_m <= 0 or not rho_l or rho_l <= 0):
        return None, {}
    dp = max(1.0, p1_pa - p2_pa)
    area = math.pi / 4.0 * pipe_id_m ** 2
    u = (w_kgs / rho_l) / area                          # downstream velocity
    # turbulent (non-cavitating) internal SPL baseline: grows with ΔP & U.
    lpi = 85.0 + 10.0 * math.log10(dp / 1e5) + 18.0 * math.log10(max(0.3, u))
    sigma_i = 1.0 / max(1e-6, (fl or 0.9) ** 2)         # incipient index proxy
    cav = ""
    if sigma is not None and sigma < sigma_i:
        inc = min(25.0, 18.0 * math.log10(max(1.0, sigma_i / max(1e-6, sigma))))
        lpi += inc
        cav = f"cavitating (σ {sigma:.2f} < σi {sigma_i:.2f}); +{inc:.0f} dB"
    fp = max(1.0, u / max(1e-3, pipe_id_m))
    tl = _cv_pipe_tl_db(pipe_id_m, fp, 1400.0)          # c ~ water sound speed
    lpe = max(20.0, lpi + tl)
    return lpe, {"u": u, "sigma_i": sigma_i, "cav": cav, "tl": tl, "lpi": lpi}


# ── Valve inputs, per-case sizing, rated-Cv selection, two-mode evaluation ──
@dataclass
class CvInputs:
    """Normalised per-valve inputs for datasheet sizing (SI internally)."""
    tag: str | None = None
    service: str | None = None
    line_no: str | None = None
    fluid: str | None = None
    temp_f: float | None = None
    nps: float | None = None
    body_key: str = "GLOBE"
    char: str = "Equal Percentage"
    rangeability: float = 50.0
    leak_class: str = "IV"
    action: str = "F"                      # F | P | T | L
    design_dp_pa: float | None = None      # fixed design ΔP (P/L; T floor base)
    exch_floor_pa: float | None = None     # AJ min-ΔP floor (P & T)
    cv100: float | None = None             # None -> selection mode
    design_open_pct: float = 80.0          # target max-flow % travel
    noise_limit_dba: float = 85.0
    min_mult: float = 0.35
    max_mult: float = 1.20
    bore_m: float | None = None
    line_in_m: float | None = None
    line_out_m: float | None = None
    fl: float = 0.90
    xt: float = 0.72
    fd: float = 0.46
    # normal-flow hydraulic state (from the march)
    p_src_pa: float | None = None
    p_dest_pa: float | None = None
    p1_norm_pa: float | None = None
    p2_norm_pa: float | None = None
    w_norm_kgs: float | None = None
    # fluid properties at the valve
    phase: str | None = None
    quality: float | None = None
    rho_l: float | None = None
    rho_v: float | None = None
    mu_l_cp: float | None = None
    mu_v_cp: float | None = None
    mw: float | None = None
    z: float | None = None
    k: float | None = None
    pv_pa: float | None = None
    pc_pa: float | None = None
    sg: float | None = None


def _cv_case_pressures(inp: CvInputs, mult: float):
    """(P1, P2, ΔP) [Pa] for a flow multiplier.

    Fixed-ΔP control (P/L, and T with its floor) holds ΔP constant across the
    turndown; flow control (F) floats — line losses grow ~flow², so the valve
    inlet falls and the valve ΔP shrinks as flow rises (the classic
    min-ΔP-at-max-flow datasheet trend)."""
    p1n, p2n = inp.p1_norm_pa, inp.p2_norm_pa
    if not p1n or not p2n:
        return None, None, None
    act = (inp.action or "F").upper()[:1]
    if act in ("P", "L", "T"):
        dp = p1n - p2n
        return p1n, p1n - dp, dp                        # constant ΔP, constant P1
    # flow control: scale the up/down line losses by mult²
    m2 = mult * mult
    if inp.p_src_pa and inp.p_dest_pa:
        l_up = max(0.0, inp.p_src_pa - p1n)
        l_dn = max(0.0, p2n - inp.p_dest_pa)
        p1 = inp.p_src_pa - l_up * m2
        p2 = inp.p_dest_pa + l_dn * m2
        if p1 - p2 < 1.0:
            p1, p2 = p1n, p2n                           # degenerate; hold normal
        return p1, p2, p1 - p2
    # no boundary info: hold P1, keep ΔP at normal (documented fallback)
    return p1n, p2n, p1n - p2n


def _cv_regime(inp: CvInputs):
    ph = (inp.phase or "").strip().lower()
    if "two" in ph or "mixed" in ph:
        return "two-phase", (inp.quality if inp.quality is not None else 0.5)
    if "liq" in ph:
        if inp.quality is not None and inp.quality > 0.01:
            return "two-phase", inp.quality
        return "liquid", 0.0
    if "vap" in ph or "gas" in ph:
        return "gas", 1.0
    if inp.rho_l and not inp.rho_v:
        return "liquid", 0.0
    return "gas", (inp.quality if inp.quality is not None else 1.0)


def _cv_size_case(inp: CvInputs, mult: float, cv100):
    """Full sizing + noise for one flow case.  Returns a result dict."""
    r = {"mult": mult}
    p1, p2, dp = _cv_case_pressures(inp, mult)
    if not p1 or not p2 or p2 >= p1:
        r["valid"] = False
        return r
    r.update(valid=True, p1=p1, p2=p2, dp=dp)
    w = (inp.w_norm_kgs or 0.0) * mult
    r["w"] = w
    regime, x_q = _cv_regime(inp)
    r["regime"] = regime
    ff = _cv_ff_factor(inp.pv_pa, inp.pc_pa)
    sigma = _cv_sigma_index(p1, p2, inp.pv_pa)
    r["ff"], r["sigma"] = ff, sigma

    if regime == "gas":
        x, x_ch, x_eff, y, choked = _cv_gas_xy(p1, p2, inp.k, inp.xt)
        rho1 = inp.rho_v
        if not rho1:
            t1_k = None
            rho1 = None
        cv_req = _cv_required_gas(w, p1, rho1 or 0.0, x_eff, y)
        r.update(x=x, x_choked=x_ch, y=y, choked=choked, cv_req=cv_req,
                 rho_eff=rho1)
    else:
        dp_choked = _cv_liquid_dp_choked(inp.fl, p1, inp.pv_pa, ff)
        choked = dp_choked > 0 and dp >= dp_choked
        dp_eff = min(dp, dp_choked) if dp_choked > 0 else dp
        r.update(dp_choked=dp_choked, choked=choked, dp_eff=dp_eff)
        if regime == "liquid":
            cv_req = _cv_required_liquid(w, inp.rho_l or 999.0, dp_eff)
            r.update(cv_req=cv_req, rho_eff=inp.rho_l)
        else:
            cv_req, rho_mix = _cv_required_two_phase(
                w, x_q, inp.rho_l, inp.rho_v, dp_eff)
            r.update(cv_req=cv_req, rho_eff=rho_mix)

    if cv100:
        r["travel"] = _cv_travel_pct(r.get("cv_req"), cv100, inp.char,
                                     inp.rangeability)
        r["pct_cv"] = 100.0 * (r.get("cv_req") or 0.0) / cv100

    # velocity + noise (downstream)
    line_m = inp.line_out_m or inp.line_in_m
    if line_m:
        area = math.pi / 4.0 * line_m ** 2
        if regime == "gas":
            rho2 = inp.rho_v
        elif regime == "liquid":
            rho2 = inp.rho_l
        else:
            if inp.rho_v and inp.rho_l and x_q is not None:
                rho2 = 1.0 / (x_q / inp.rho_v + (1.0 - x_q) / inp.rho_l)
            else:
                rho2 = r.get("rho_eff")
        if rho2 and rho2 > 0:
            r["v2"] = (w / rho2) / area
        # gas → aerodynamic (8-3); liquid & flashing/two-phase → hydrodynamic
        # (8-4), which carries the cavitation/flashing increment.
        if regime == "gas":
            c2 = _cv_sonic_velocity_gas(inp.k, inp.z, None, inp.mw)
            lpe, det = _cv_noise_aero(w, p1, p2, None, inp.mw, inp.z, inp.k,
                                      inp.xt, inp.fl, inp.fd, line_m,
                                      rho2, c2)
        else:
            lpe, det = _cv_noise_hydro(w, p1, p2, inp.pv_pa, inp.fl,
                                       inp.rho_l or rho2, line_m, sigma)
        r["noise_dba"], r["noise_detail"] = lpe, det
    return r


def _cv_select_rated(inp: CvInputs, cases):
    """Selection mode: pick the smallest ladder Rated Cv (≤ body full-open max)
    that puts MAX-flow travel at/under the design opening and keeps MIN-flow
    travel controllable.  Returns (cv100, note)."""
    body = _CV_BODY[inp.body_key]
    cv_max_body = body["cd"] * (inp.nps or 2.0) ** 2
    cv_req_max = max((c.get("cv_req") or 0.0) for c in cases.values())
    if cv_req_max <= 0:
        return None, "no positive required Cv - cannot size"
    target = inp.design_open_pct or 80.0
    ladder = [c for c in _CV_RATED_LADDER if c <= cv_max_body * 1.001]
    if not ladder:
        ladder = [round(cv_max_body, 1)]
    for rung in ladder:
        tmax = _cv_travel_pct(cv_req_max, rung, inp.char, inp.rangeability)
        cv_req_min = min((c.get("cv_req") or 0.0) for c in cases.values()
                         if c.get("cv_req"))
        tmin = _cv_travel_pct(cv_req_min, rung, inp.char, inp.rangeability)
        if tmax is not None and tmax <= target and (tmin is None or tmin >= 8.0):
            return rung, (f"selected {rung:g} from generic {body['label']} "
                          f"ladder (max-flow travel {tmax:.0f}% <= target "
                          f"{target:.0f}%; body full-open max ~{cv_max_body:.0f})")
    # nothing satisfied the window: take the largest rung <= body max
    rung = ladder[-1]
    return rung, (f"selected {rung:g} (largest generic {body['label']} rung "
                  f"<= body max ~{cv_max_body:.0f}); travel window not fully met "
                  "- verify size/trim with vendor")


def _cv_evaluate_valve(inp: CvInputs) -> dict:
    """Run Min/Norm/Max, select or check the Rated Cv, apply the dB(A) limit,
    and assemble verdicts + recommendations.  Returns the datasheet payload."""
    mults = {"MIN": inp.min_mult, "NOR": 1.0, "MAX": inp.max_mult}
    mode = "adequacy" if inp.cv100 else "selection"
    cv100 = inp.cv100
    sel_note = None
    if mode == "selection":
        prelim = {code: _cv_size_case(inp, m, None) for code, m in mults.items()}
        cv100, sel_note = _cv_select_rated(inp, prelim)
    cases = {code: _cv_size_case(inp, m, cv100) for code, m in mults.items()}

    # verdicts
    cv_req_max = max((c.get("cv_req") or 0.0) for c in cases.values())
    cap_ok = bool(cv100) and cv100 >= cv_req_max
    travels = [c.get("travel") for c in cases.values() if c.get("travel") is not None]
    travel_ok = bool(travels) and all(5.0 <= t <= 95.0 for t in travels)
    noises = [c.get("noise_dba") for c in cases.values() if c.get("noise_dba")]
    noise_max = max(noises) if noises else None
    noise_ok = (noise_max is None) or (noise_max <= inp.noise_limit_dba)
    choked_any = any(c.get("choked") for c in cases.values())

    recs = []
    if not noise_ok:
        recs.append(
            f"Predicted noise {noise_max:.0f} dB(A) exceeds the "
            f"{inp.noise_limit_dba:.0f} dB(A) limit at max flow. Rated-Cv "
            "choice does not change service ΔP/noise - mitigate with low-noise "
            "/ multistage trim, a larger body (lower outlet velocity), or split "
            "the ΔP across two valves in series.")
    if choked_any:
        recs.append("Choked / cavitating flow present - confirm trim and "
                    "material against the manufacturer's σ curves (ISA RP75.23) "
                    "and consider anti-cavitation trim.")
    if travels and (min(travels) < 10.0):
        recs.append("Low % travel near closed at min flow - controllability "
                    "may suffer; consider reduced trim.")
    if travels and (max(travels) > 90.0):
        recs.append("High % travel near max flow - little rangeability margin; "
                    "verify Rated Cv / body size.")

    return {
        "inp": inp, "mode": mode, "cv100": cv100, "sel_note": sel_note,
        "cases": cases, "verdicts": {
            "capacity": cap_ok, "travel": travel_ok, "noise": noise_ok,
            "noise_max": noise_max, "cv_req_max": cv_req_max,
        },
        "recommendations": recs,
    }


# ── Datasheet sheet (Min/Norm/Max vendor-style layout) ─────────────────────
_CVDS_HDR   = PatternFill("solid", fgColor="1F4E79")   # dark blue band
_CVDS_SEC   = PatternFill("solid", fgColor="D6E4F0")   # light blue section
_CVDS_LBL   = PatternFill("solid", fgColor="F2F2F2")   # label tint
_CVDS_PASS  = PatternFill("solid", fgColor="C6EFCE")
_CVDS_FAIL  = PatternFill("solid", fgColor="FFC7CE")
_CVDS_NORM  = PatternFill("solid", fgColor="FFF2CC")   # highlight Normal col
_cvds_thin  = Side(style="thin", color="BFBFBF")
_CVDS_BORD  = Border(left=_cvds_thin, right=_cvds_thin,
                     top=_cvds_thin, bottom=_cvds_thin)


def _cvc(ws, r, c, v=None, *, bold=False, fill=None, align="left",
         color="262626", size=10, border=True, wrap=False):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(bold=bold, color=color, size=size)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if fill:
        cell.fill = fill
    if border:
        cell.border = _CVDS_BORD
    return cell


def _cv_disp(res, key, conv, nd=3, default="—"):
    v = res.get(key)
    if v is None or not res.get("valid", True):
        return default
    try:
        return round(conv(v), nd)
    except Exception:
        return default


def _build_cv_datasheet_sheet(wb, payload: dict, run_label: str = "Main"):
    """One control-valve datasheet sheet mirroring the vendor layout:
    header block -> Min/Norm/Max service conditions -> flowing/sizing/noise
    results -> verdicts + recommendations."""
    inp: CvInputs = payload["inp"]
    cases = payload["cases"]
    v = payload["verdicts"]
    body = _CV_BODY[inp.body_key]
    tag = inp.tag or "CV"
    title = re.sub(r"[^A-Za-z0-9_-]", "_", f"CV_{tag}")[:31]
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 30, "B": 12, "C": 15, "D": 15, "E": 15, "F": 26}.items():
        ws.column_dimensions[col].width = w

    P = lambda pa: pa / _CV_PSI_TO_PA
    W = lambda kg: kg * _CV_KGS_TO_LBH
    LB = lambda kgm3: kgm3 / 16.018463

    # ── Title ──
    ws.merge_cells("A1:F1")
    _cvc(ws, 1, 1, f"CONTROL VALVE DATASHEET  —  {tag}", bold=True,
         fill=_CVDS_HDR, color="FFFFFF", size=13, align="center")
    ws.row_dimensions[1].height = 22

    # ── Header block ──
    act_lbl = {"F": "Flow", "P": "Pressure", "T": "Temperature",
               "L": "Level"}.get((inp.action or "F").upper()[:1], "Flow")
    mode_lbl = ("Sizing / selection (new valve)" if payload["mode"] == "selection"
                else "Adequacy check (existing valve)")
    hdr = [
        ("Valve Tag", tag, "Service", inp.service or "—"),
        ("Line No", inp.line_no or "—", "Fluid", inp.fluid or "—"),
        ("Body style", body["label"], "Characteristic", inp.char),
        ("Valve size (NPS)", inp.nps or "—", "Control action", act_lbl),
        ("Rated Cv (Cv100)", round(inp.cv100 or payload["cv100"] or 0, 2),
         "Leakage class", inp.leak_class),
        ("Seat / port bore (in)", round(inp.bore_m / 0.0254, 3) if inp.bore_m else "—",
         "Rangeability R", inp.rangeability),
        ("Mode", mode_lbl, "Noise limit dB(A)", inp.noise_limit_dba),
    ]
    r = 2
    for l1, v1, l2, v2 in hdr:
        _cvc(ws, r, 1, l1, bold=True, fill=_CVDS_LBL)
        _cvc(ws, r, 2, v1)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        _cvc(ws, r, 4, l2, bold=True, fill=_CVDS_LBL)
        _cvc(ws, r, 5, v2)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
        r += 1
    if payload.get("sel_note"):
        _cvc(ws, r, 1, "Selection basis", bold=True, fill=_CVDS_LBL)
        _cvc(ws, r, 2, payload["sel_note"], wrap=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 28
        r += 1

    # ── Service-conditions / results grid ──
    def sec(label):
        nonlocal r
        _cvc(ws, r, 1, label, bold=True, fill=_CVDS_SEC)
        for c in range(2, 7):
            _cvc(ws, r, c, "", fill=_CVDS_SEC)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        r += 1

    def row(label, unit, mn, no, mx, note=""):
        nonlocal r
        _cvc(ws, r, 1, label, fill=_CVDS_LBL)
        _cvc(ws, r, 2, unit, align="center")
        _cvc(ws, r, 3, mn, align="center")
        _cvc(ws, r, 4, no, align="center", fill=_CVDS_NORM, bold=True)
        _cvc(ws, r, 5, mx, align="center")
        _cvc(ws, r, 6, note, wrap=True)
        r += 1

    cMIN, cNOR, cMAX = cases["MIN"], cases["NOR"], cases["MAX"]
    # column header
    _cvc(ws, r, 1, "SERVICE CONDITIONS", bold=True, fill=_CVDS_HDR, color="FFFFFF")
    _cvc(ws, r, 2, "Units", bold=True, fill=_CVDS_HDR, color="FFFFFF", align="center")
    _cvc(ws, r, 3, "Minimum", bold=True, fill=_CVDS_HDR, color="FFFFFF", align="center")
    _cvc(ws, r, 4, "Normal", bold=True, fill=_CVDS_HDR, color="FFFFFF", align="center")
    _cvc(ws, r, 5, "Maximum", bold=True, fill=_CVDS_HDR, color="FFFFFF", align="center")
    _cvc(ws, r, 6, "Notes", bold=True, fill=_CVDS_HDR, color="FFFFFF", align="center")
    r += 1

    row("Flow Rate", "lb/h",
        _cv_disp(cMIN, "w", W, 1), _cv_disp(cNOR, "w", W, 1),
        _cv_disp(cMAX, "w", W, 1),
        f"turndown {inp.min_mult:g}× / {inp.max_mult:g}× of normal")
    row("Inlet Pressure P1", "psia",
        _cv_disp(cMIN, "p1", P, 2), _cv_disp(cNOR, "p1", P, 2),
        _cv_disp(cMAX, "p1", P, 2))
    row("Outlet Pressure P2", "psia",
        _cv_disp(cMIN, "p2", P, 2), _cv_disp(cNOR, "p2", P, 2),
        _cv_disp(cMAX, "p2", P, 2))
    row("Pressure Drop ΔP", "psi",
        _cv_disp(cMIN, "dp", P, 2), _cv_disp(cNOR, "dp", P, 2),
        _cv_disp(cMAX, "dp", P, 2))
    tf = round(inp.temp_f, 1) if inp.temp_f is not None else "—"
    row("Temperature", "°F", tf, tf, tf)
    pv = round(inp.pv_pa / _CV_PSI_TO_PA, 3) if inp.pv_pa else "—"
    pc = round(inp.pc_pa / _CV_PSI_TO_PA, 2) if inp.pc_pa else "—"
    row("Vapor Pressure Pv", "psia", pv, pv, pv)
    row("Critical Pressure Pc", "psia", pc, pc, pc)
    muv = round(inp.mu_l_cp or inp.mu_v_cp or 0, 4) if (inp.mu_l_cp or inp.mu_v_cp) else "—"
    row("Viscosity", "cP", muv, muv, muv)
    sg = round(inp.sg, 4) if inp.sg else "—"
    row("Liquid Gf / SG", "—", sg, sg, sg)

    sec("FLOWING CONDITIONS / SIZING")
    row("Flow regime", "",
        cMIN.get("regime", "—"), cNOR.get("regime", "—"), cMAX.get("regime", "—"))
    row("Choked / cavitating?", "",
        "Y" if cMIN.get("choked") else "N", "Y" if cNOR.get("choked") else "N",
        "Y" if cMAX.get("choked") else "N")
    row("Required Cv", "",
        _cv_disp(cMIN, "cv_req", lambda x: x, 4),
        _cv_disp(cNOR, "cv_req", lambda x: x, 4),
        _cv_disp(cMAX, "cv_req", lambda x: x, 4),
        "IEC 60534-2-1 (US Cv basis)")
    row("Oversized Req. Cv (×1.25)", "",
        _cv_disp(cMIN, "cv_req", lambda x: x * 1.25, 4),
        _cv_disp(cNOR, "cv_req", lambda x: x * 1.25, 4),
        _cv_disp(cMAX, "cv_req", lambda x: x * 1.25, 4))
    row("% of Rated Cv", "%",
        _cv_disp(cMIN, "pct_cv", lambda x: x, 2),
        _cv_disp(cNOR, "pct_cv", lambda x: x, 2),
        _cv_disp(cMAX, "pct_cv", lambda x: x, 2))
    row("% Travel", "%",
        _cv_disp(cMIN, "travel", lambda x: x, 1),
        _cv_disp(cNOR, "travel", lambda x: x, 1),
        _cv_disp(cMAX, "travel", lambda x: x, 1),
        f"trim: {inp.char}")
    row("FL / xT", "",
        round(inp.fl, 3), round(inp.fl, 3), round(inp.fl, 3),
        f"body typical (xT={inp.xt:g})")
    # sizing var: liquid ΔP_choked or gas Y
    if cNOR.get("regime") == "gas":
        row("Expansion factor Y", "",
            _cv_disp(cMIN, "y", lambda x: x, 3), _cv_disp(cNOR, "y", lambda x: x, 3),
            _cv_disp(cMAX, "y", lambda x: x, 3))
    else:
        row("ΔP choked = FL²(P1−FF·Pv)", "psia",
            _cv_disp(cMIN, "dp_choked", P, 2), _cv_disp(cNOR, "dp_choked", P, 2),
            _cv_disp(cMAX, "dp_choked", P, 2))
    row("Cavitation index σ", "",
        _cv_disp(cMIN, "sigma", lambda x: x, 2), _cv_disp(cNOR, "sigma", lambda x: x, 2),
        _cv_disp(cMAX, "sigma", lambda x: x, 2), "ISA RP75.23")
    row("Valve/outlet velocity", "m/s",
        _cv_disp(cMIN, "v2", lambda x: x, 2), _cv_disp(cNOR, "v2", lambda x: x, 2),
        _cv_disp(cMAX, "v2", lambda x: x, 2))
    row("Sound Level", "dB(A)",
        _cv_disp(cMIN, "noise_dba", lambda x: x, 0),
        _cv_disp(cNOR, "noise_dba", lambda x: x, 0),
        _cv_disp(cMAX, "noise_dba", lambda x: x, 0),
        "IEC 60534-8-3/-8-4 (screening)")

    # ── Verdicts ──
    sec("ADEQUACY VERDICTS")
    def verdict(label, ok, detail=""):
        nonlocal r
        _cvc(ws, r, 1, label, bold=True, fill=_CVDS_LBL)
        _cvc(ws, r, 2, "PASS" if ok else "FAIL", align="center", bold=True,
             fill=_CVDS_PASS if ok else _CVDS_FAIL)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        _cvc(ws, r, 4, detail, wrap=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
        r += 1
    verdict("Capacity (Cv100 ≥ Cv req)", v["capacity"],
            f"Rated {inp.cv100 or payload['cv100']:.2f} vs max req {v['cv_req_max']:.3f}")
    verdict("Travel window (5–95%)", v["travel"])
    nmax = v["noise_max"]
    verdict(f"Noise ≤ {inp.noise_limit_dba:.0f} dB(A)", v["noise"],
            f"max predicted {nmax:.0f} dB(A)" if nmax else "not evaluated")
    # beta ratio
    b_in = _cv_beta_ratio(inp.bore_m, inp.line_in_m)
    b_out = _cv_beta_ratio(inp.bore_m, inp.line_out_m)
    _cvc(ws, r, 1, "β ratio (in / out)", bold=True, fill=_CVDS_LBL)
    _cvc(ws, r, 2, f"{b_in:.3f}" if b_in else "—", align="center")
    _cvc(ws, r, 3, f"{b_out:.3f}" if b_out else "—", align="center")
    _cvc(ws, r, 4, "bore ÷ line ID", wrap=True)
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
    r += 1
    # seat leakage
    leak = _cv_seat_leakage(inp.leak_class, inp.cv100 or payload["cv100"],
                            (cNOR.get("dp") or 0.0),
                            (inp.bore_m / 0.0254 * 25.4) if inp.bore_m else None)
    _cvc(ws, r, 1, f"Seat leakage (Class {leak.get('class', '—')})",
         bold=True, fill=_CVDS_LBL)
    _cvc(ws, r, 2, leak.get("desc", ""), wrap=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.row_dimensions[r].height = 26
    r += 1

    # ── Recommendations ──
    if payload["recommendations"]:
        sec("RECOMMENDATIONS")
        for rec in payload["recommendations"]:
            _cvc(ws, r, 1, "•", align="center", fill=_CVDS_LBL)
            _cvc(ws, r, 2, rec, wrap=True)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            ws.row_dimensions[r].height = max(26, 13 * (1 + len(rec) // 60))
            r += 1

    # ── Method footer ──
    sec("METHOD")
    for line in [
        "Sizing: IEC 60534-2-1 / ISA 75.01.01 (liquid FF/FL choked; gas x/xT/Y; "
        "two-phase homogeneous).  US Cv basis (gpm·√SG/√psi liquid; lb/h,psia,"
        "lb/ft³ gas).",
        "Cavitation: ISA RP75.23 σ.  Seat leakage: FCI 70-2 / IEC 60534-4.",
        "Noise: IEC 60534-8-3 (aerodynamic) / -8-4 (hydrodynamic) — screening-"
        "grade dB(A); confirm severe/critical service with vendor acoustic data.",
        "Min/Max are turndown multiples of the marched Normal operating point; "
        "FL/xT/Fd are body-style typicals — confirm against the selected valve.",
    ]:
        _cvc(ws, r, 1, line, wrap=True, size=9, color="595959")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.row_dimensions[r].height = max(24, 12 * (1 + len(line) // 90))
        r += 1
    return title


def _cv_bubble_point_psia(feed, temp_f, p_hi_psia):
    """Bubble-point pressure [psia] at temp_f: the P where the first vapour
    appears (β≈0).  Bisection between a low bound and P1.  None on failure."""
    if feed is None or not p_hi_psia or p_hi_psia <= 0:
        return None
    def beta_at(pp):
        try:
            fr = flash(feed, pp, temp_f=temp_f)
            return (fr.beta if fr and fr.beta is not None else 0.0)
        except Exception:
            return None
    b_hi = beta_at(p_hi_psia)
    if b_hi is None:
        return None
    if b_hi > 1e-3:
        return None                    # already flashing at P1 (no subcooling)
    lo, hi = max(1e-3, p_hi_psia * 1e-4), p_hi_psia
    for _ in range(60):
        mid = math.sqrt(lo * hi)
        b = beta_at(mid)
        if b is None:
            return None
        if b > 1e-3:
            lo = mid
        else:
            hi = mid
        if hi / lo < 1.0001:
            break
    return 0.5 * (lo + hi)


def _cv_inputs_from_station(station, stream_name=None) -> "CvInputs | None":
    """Build a CvInputs from a control-valve Station's captured cv_raw."""
    raw = getattr(station, "cv_raw", None)
    if not raw:
        return None
    row = raw.get("row", {})
    PA = _CV_PSI_TO_PA
    to_pa = lambda psia: (psia * PA) if psia is not None else None
    lbft3 = 16.018463
    w_lbhr = (raw.get("vap_mass") or 0.0) + (raw.get("liq_mass") or 0.0)
    rho_l = (raw["liq_density"] * lbft3) if raw.get("liq_density") else None
    rho_v = (raw["vap_density"] * lbft3) if raw.get("vap_density") else None
    p1 = station.p_in_psia
    pv = _cv_bubble_point_psia(raw.get("feed"), raw.get("temp_f"), p1)
    nps = raw.get("bore_in")
    body_key = _cv_body_key(row.get("Valve Body Style"))
    body = _CV_BODY[body_key]
    cv_over = row.get("Rated Cv")
    return CvInputs(
        tag=row.get("Comp ID") or station.comp_id,
        service=row.get("Notes"),
        line_no=row.get("Line No"),
        fluid=stream_name,
        temp_f=raw.get("temp_f"),
        nps=nps,
        body_key=body_key,
        char=_cv_trim_name(row.get("Valve Characteristic")),
        rangeability=50.0,
        leak_class=(row.get("Seat Leakage Class") or "IV"),
        action=raw.get("ctype") or "F",
        design_dp_pa=to_pa(raw.get("fixed_dp") or None),
        exch_floor_pa=to_pa(raw.get("exch_dp")),
        cv100=(float(cv_over) if cv_over else None),
        design_open_pct=(row.get("Design Opening %") or 80.0),
        noise_limit_dba=(row.get("Noise Limit dBA") or 85.0),
        min_mult=(row.get("Min Flow Mult") or 0.35),
        max_mult=(row.get("Max Flow Mult") or 1.20),
        bore_m=(nps * 0.0254) if nps else None,
        line_in_m=(raw["line_bore_in"] * 0.0254) if raw.get("line_bore_in") else None,
        line_out_m=(raw["line_bore_in"] * 0.0254) if raw.get("line_bore_in") else None,
        fl=body["fl"], xt=body["xt"], fd=body["fd"],
        p_src_pa=to_pa(raw.get("p_src")),
        p_dest_pa=to_pa(raw.get("dest_p")),
        p1_norm_pa=to_pa(station.p_in_psia),
        p2_norm_pa=to_pa(station.p_out_psia),
        w_norm_kgs=(w_lbhr / _CV_KGS_TO_LBH) if w_lbhr else None,
        phase=raw.get("phase"),
        quality=raw.get("quality"),
        rho_l=rho_l, rho_v=rho_v,
        mu_l_cp=raw.get("liq_visc"), mu_v_cp=raw.get("vap_visc"),
        mw=(raw.get("vap_mw") or raw.get("mol_weight")),
        z=raw.get("vap_z"), k=raw.get("gamma"),
        pv_pa=to_pa(pv), pc_pa=to_pa(raw.get("pc_psia")),
        sg=((rho_l / 999.0) if rho_l else None),
    )


def _build_cv_datasheets(wb, stations, stream_name=None, run_label="Main"):
    """Build one datasheet sheet per control valve in `stations`.  Returns the
    list of created sheet titles (empty if the block has no control valves)."""
    titles = []
    for st in stations:
        if not getattr(st, "cv_raw", None):
            continue
        try:
            inp = _cv_inputs_from_station(st, stream_name=stream_name)
            if inp is None:
                continue
            payload = _cv_evaluate_valve(inp)
            titles.append(_build_cv_datasheet_sheet(wb, payload, run_label))
        except Exception as exc:
            print(f"  (CV datasheet for {getattr(st, 'comp_id', '?')} "
                  f"skipped: {exc})")
    return titles


if __name__ == "__main__":
    main()
