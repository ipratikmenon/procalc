#!/usr/bin/env python3
"""HyCalign Hydraulics — unified engine (single file to run).

This file merges the former  asme_data.py + hydraulics_engine.py +
hydraulics_engine_flash_noiso.py + flash_vle.py + flow_pattern_data.py +
flow_pattern_maps.py  into ONE engine.  Everything the engine needs is in this file: the units layer (UNITS sheet), the project piping
classes (PMS_CLASSES — rename a class / change its schedule or wall thickness there), the
stream map (actual stream number → simulation stream number) and the 98-component PR databank.
The only companion is  hmbbuilder.py  (PRO/II dump → HMB workbook).

Manual (no-HMB) streams may carry a mass/mole composition via the
Stream_Composition input sheet; streams entering at different Seq are mixed.
Pressure_Profile and FIV_EI_T2.2 report Mach (vapor) and ρv² (vapor / liquid).

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

import sys as _sys
A = _sys.modules[__name__]   # asme_data merged into this file

# ════════════════════════════════════════════════════════════════════════
#  UNITS LAYER  (embedded copy of the shared Procalc units layer, common/units.py)
#  Every quantity is computed internally in FPS; the UNITS sheet of the input workbook
#  decides how values are entered and presented.
# ════════════════════════════════════════════════════════════════════════
# ── conversion constants (exact / NIST) ────────────────────────────────────
LB_PER_KG = 2.2046226218487757
FT_PER_M = 3.2808398950131233
IN_PER_MM = 1.0 / 25.4
PSI_PER_KPA = 0.14503773773020922
PSIA_ATM = 14.695948775513449
KPA_ATM = 101.325
BTU_PER_KJ = 0.9478171203133172

# quantity → {unit: (to_internal(v), from_internal(v))}
_LIN = lambda f: (lambda v: v * f, lambda v: v / f)  # noqa: E731

_CONVERTERS: dict[str, dict[str, tuple]] = {
    # absolute pressure → psia
    "P": {
        "psia": _LIN(1.0),
        "psig": (lambda v: v + PSIA_ATM, lambda v: v - PSIA_ATM),
        "kPa":  _LIN(PSI_PER_KPA),                      # kPa(a)
        "kPag": (lambda v: (v + KPA_ATM) * PSI_PER_KPA,
                 lambda v: v / PSI_PER_KPA - KPA_ATM),
        "MPa":  _LIN(PSI_PER_KPA * 1000.0),
        "bara": _LIN(PSI_PER_KPA * 100.0),
        "barg": (lambda v: v * 100.0 * PSI_PER_KPA + PSIA_ATM,
                 lambda v: (v - PSIA_ATM) / (100.0 * PSI_PER_KPA)),
        "kg/cm2a": _LIN(14.223343307120154),
        "kg/cm2g": (lambda v: v * 14.223343307120154 + PSIA_ATM,
                    lambda v: (v - PSIA_ATM) / 14.223343307120154),
    },
    # differential pressure → psi
    "dP": {
        "psi":  _LIN(1.0),
        "kPa":  _LIN(PSI_PER_KPA),
        "bar":  _LIN(PSI_PER_KPA * 100.0),
        "mbar": _LIN(PSI_PER_KPA / 10.0),
        "kg/cm2": _LIN(14.223343307120154),
    },
    # temperature → °F
    "T": {
        "degF": _LIN(1.0), "°F": _LIN(1.0),
        "degC": (lambda v: v * 1.8 + 32.0, lambda v: (v - 32.0) / 1.8),
        "°C":   (lambda v: v * 1.8 + 32.0, lambda v: (v - 32.0) / 1.8),
        "K":    (lambda v: (v - 273.15) * 1.8 + 32.0,
                 lambda v: (v - 32.0) / 1.8 + 273.15),
    },
    # temperature difference → °F
    "dT": {"degF": _LIN(1.0), "degC": _LIN(1.8), "K": _LIN(1.8)},
    # mass flow → lb/hr
    "mflow": {
        "lb/hr": _LIN(1.0), "lb/h": _LIN(1.0),
        "kg/hr": _LIN(LB_PER_KG), "kg/h": _LIN(LB_PER_KG),
        "kg/s":  _LIN(LB_PER_KG * 3600.0),
        "t/h":   _LIN(LB_PER_KG * 1000.0),
    },
    # molar flow → lb-mol/hr
    "molflow": {
        "lb-mol/hr": _LIN(1.0),
        "kgmol/hr":  _LIN(LB_PER_KG), "kmol/h": _LIN(LB_PER_KG),
    },
    # density → lb/ft³
    "rho": {
        "lb/ft3": _LIN(1.0), "lb/ft³": _LIN(1.0),
        "kg/m3":  _LIN(LB_PER_KG / FT_PER_M**3),
        "kg/m³":  _LIN(LB_PER_KG / FT_PER_M**3),
    },
    # length → ft
    "L": {"ft": _LIN(1.0), "m": _LIN(FT_PER_M), "mm": _LIN(FT_PER_M / 1000.0)},
    # small length / diameter → in   (NPS stays in inches by convention)
    "Lin": {"in": _LIN(1.0), "mm": _LIN(IN_PER_MM)},
    # area → ft²
    "A": {"ft2": _LIN(1.0), "ft²": _LIN(1.0),
          "m2": _LIN(FT_PER_M**2), "m²": _LIN(FT_PER_M**2)},
    # small area (orifice) → in²
    "Ain": {"in2": _LIN(1.0), "in²": _LIN(1.0),
            "cm2": _LIN((IN_PER_MM * 10.0)**2), "cm²": _LIN((IN_PER_MM * 10.0)**2),
            "mm2": _LIN(IN_PER_MM**2), "mm²": _LIN(IN_PER_MM**2)},
    # volume → ft³
    "V": {"ft3": _LIN(1.0), "ft³": _LIN(1.0),
          "m3": _LIN(FT_PER_M**3), "m³": _LIN(FT_PER_M**3),
          "L": _LIN(FT_PER_M**3 / 1000.0)},
    # heat duty → BTU/hr
    "Q": {"BTU/hr": _LIN(1.0),
          "kW": _LIN(BTU_PER_KJ * 3600.0),
          "W":  _LIN(BTU_PER_KJ * 3.6),
          "MW": _LIN(BTU_PER_KJ * 3.6e6),
          "MMBTU/hr": _LIN(1.0e6), "kcal/h": _LIN(3.968320719312103)},
    # specific enthalpy / latent heat → BTU/lb
    "h": {"BTU/lb": _LIN(1.0), "kJ/kg": _LIN(BTU_PER_KJ / LB_PER_KG),
          "kcal/kg": _LIN(1.7988310432962014)},
    # velocity → ft/s
    "v": {"ft/s": _LIN(1.0), "m/s": _LIN(FT_PER_M)},
    # volumetric flow → ft³/hr
    "qvol": {"ft3/hr": _LIN(1.0), "m3/h": _LIN(FT_PER_M**3),
             "gpm": _LIN(8.020833333333334), "m3/hr": _LIN(FT_PER_M**3)},
    # dimensionless / identical in both systems
    "visc": {"cP": _LIN(1.0)},
    "MW":   {"g/mol": _LIN(1.0)},
    "-":    {"-": _LIN(1.0), "—": _LIN(1.0)},
    "pct":  {"%": _LIN(1.0)},
    "st":   {"dyn/cm": _LIN(1.0), "mN/m": _LIN(1.0)},   # numerically equal
}

# per-system default unit per quantity
SYSTEM_DEFAULTS: dict[str, dict[str, str]] = {
    "FPS": {
        "P": "psia", "dP": "psi", "T": "degF", "dT": "degF",
        "mflow": "lb/hr", "molflow": "lb-mol/hr", "rho": "lb/ft3",
        "L": "ft", "Lin": "in", "A": "ft2", "Ain": "in2", "V": "ft3",
        "Q": "BTU/hr", "h": "BTU/lb", "v": "ft/s", "qvol": "ft3/hr",
        "visc": "cP", "MW": "g/mol", "-": "-", "pct": "%", "st": "dyn/cm",
    },
    "SI": {
        "P": "kPa", "dP": "kPa", "T": "degC", "dT": "degC",
        "mflow": "kg/hr", "molflow": "kgmol/hr", "rho": "kg/m3",
        "L": "m", "Lin": "mm", "A": "m2", "Ain": "cm2", "V": "m3",
        "Q": "kW", "h": "kJ/kg", "v": "m/s", "qvol": "m3/h",
        "visc": "cP", "MW": "g/mol", "-": "-", "pct": "%", "st": "dyn/cm",
    },
}

QUANTITY_NAMES = {
    "P": "Pressure (absolute)", "dP": "Pressure drop", "T": "Temperature",
    "dT": "Temperature difference", "mflow": "Mass flow", "molflow": "Molar flow",
    "rho": "Density", "L": "Length / elevation", "Lin": "Bore / small length",
    "A": "Surface area", "Ain": "Orifice area", "V": "Volume",
    "Q": "Heat duty", "h": "Specific enthalpy / latent heat", "v": "Velocity",
    "qvol": "Volumetric flow", "visc": "Viscosity", "MW": "Molecular weight",
}

UNITS_SHEET = "UNITS"


class UnitSystem:
    """Converts between user-facing units and internal FPS canonical units."""

    def __init__(self, system: str = "FPS",
                 overrides: dict[str, str] | None = None):
        system = (system or "FPS").strip().upper()
        if system not in SYSTEM_DEFAULTS:
            system = "FPS"
        self.system = system
        self.sel: dict[str, str] = dict(SYSTEM_DEFAULTS[system])
        for qty, unit in (overrides or {}).items():
            if qty in _CONVERTERS and unit in _CONVERTERS[qty]:
                self.sel[qty] = unit

    # ── conversion ──────────────────────────────────────────────────────
    def to_internal(self, qty: str, value):
        if value is None or qty not in _CONVERTERS:
            return value
        try:
            v = float(value)
        except (TypeError, ValueError):
            return value
        return _CONVERTERS[qty][self.sel[qty]][0](v)

    def from_internal(self, qty: str, value):
        if value is None or qty not in _CONVERTERS:
            return value
        try:
            v = float(value)
        except (TypeError, ValueError):
            return value
        return _CONVERTERS[qty][self.sel[qty]][1](v)

    def disp(self, qty: str, value, nd: int | None = 4):
        """from_internal + optional rounding (None → no rounding)."""
        out = self.from_internal(qty, value)
        if nd is not None and isinstance(out, float):
            out = round(out, nd)
        return out

    # ── labels ──────────────────────────────────────────────────────────
    def label(self, qty: str) -> str:
        u = self.sel.get(qty, "")
        return {"degF": "°F", "degC": "°C"}.get(u, u)

    def hdr(self, base: str, qty: str) -> str:
        """'Start P', 'P'  →  'Start P (psia)' / 'Start P (kPa)'."""
        lbl = self.label(qty)
        return f"{base} ({lbl})" if lbl and lbl != "-" else base

    @property
    def is_fps(self) -> bool:
        return self.system == "FPS"

    # ── workbook I/O ────────────────────────────────────────────────────
    @classmethod
    def from_workbook(cls, wb) -> "UnitSystem":
        """Read the UNITS sheet (graceful: absent sheet → FPS)."""
        if UNITS_SHEET not in wb.sheetnames:
            return cls("FPS")
        ws = wb[UNITS_SHEET]
        system, overrides = "FPS", {}
        for row in ws.iter_rows(min_row=1, max_row=60, max_col=4,
                                values_only=True):
            if not row or row[0] is None:
                continue
            key = str(row[0]).strip()
            if key.lower() in ("unit system", "system"):
                system = str(row[1] or "FPS").strip().upper()
            elif key in _CONVERTERS and row[2]:
                unit = str(row[2]).strip()
                # normalise pretty labels back to converter keys
                unit = {"°F": "degF", "°C": "degC"}.get(unit, unit)
                overrides[key] = unit
        return cls(system, overrides)


def add_units_sheet(wb, system: str = "FPS", position: int = 0):
    """Insert a styled UNITS sheet into ``wb`` (openpyxl Workbook)."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation

    NAVY, WHITE, LGRAY, AMBER = "1F4973", "FFFFFF", "F2F2F2", "FFF2CC"
    if UNITS_SHEET in wb.sheetnames:
        del wb[UNITS_SHEET]
    ws = wb.create_sheet(UNITS_SHEET, position)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "C00000"

    def cell(r, c, v, bg=None, fg="000000", bold=False, sz=10):
        x = ws.cell(row=r, column=c, value=v)
        x.font = Font(name="Calibri", size=sz, bold=bold, color=fg)
        if bg:
            x.fill = PatternFill("solid", fgColor=bg)
        x.alignment = Alignment(horizontal="left", vertical="center")
        return x

    ws.merge_cells("A1:D1")
    cell(1, 1, "UNITS  —  choose FPS or SI; per-quantity overrides optional",
         bg=NAVY, fg=WHITE, bold=True, sz=11)
    ws.row_dimensions[1].height = 22

    cell(3, 1, "Unit System", bg=LGRAY, bold=True)
    c = cell(3, 2, (system or "FPS").upper(), bg=AMBER, bold=True)
    dv = DataValidation(type="list", formula1='"FPS,SI"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B3")
    cell(3, 3, "← FPS or SI.  All inputs read & all results written in these units.")

    cell(5, 1, "Qty code", bg=NAVY, fg=WHITE, bold=True)
    cell(5, 2, "Quantity", bg=NAVY, fg=WHITE, bold=True)
    cell(5, 3, "Unit override (optional)", bg=NAVY, fg=WHITE, bold=True)
    cell(5, 4, "Options", bg=NAVY, fg=WHITE, bold=True)
    r = 6
    sysd = SYSTEM_DEFAULTS.get((system or "FPS").upper(), SYSTEM_DEFAULTS["FPS"])
    for qty, name in QUANTITY_NAMES.items():
        cell(r, 1, qty, bg=LGRAY)
        cell(r, 2, name)
        cell(r, 3, None, bg=AMBER)   # blank = system default
        opts = " | ".join(_CONVERTERS[qty].keys())
        cell(r, 4, f"default: {sysd[qty]}   |   {opts}", fg="808080", sz=9)
        r += 1
    cell(r + 1, 1, "Leave overrides blank to use the system defaults. "
         "Overrides apply per quantity (e.g. P = barg while system = SI).",
         fg="808080", sz=9)
    for col, w in (("A", 10), ("B", 30), ("C", 22), ("D", 70)):
        ws.column_dimensions[col].width = w
    return ws

UN = _sys.modules[__name__]          # units layer lives in this file (UN.UnitSystem, UN.add_units_sheet …)

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
# ── PROJECT PIPING CLASSES (PMS) ────────────────────────────────────────────
# One line per class:   "<class name>": ("<material of construction>", (rule, rule, …)),
#   rule = (NPS from, NPS to, schedule-or-wall [, minimum schedule for 'CAL'])
#   schedule-or-wall:  a B36.10 / B36.19 designation ("40S", "80", "160", "STD", "XS", "XXS" …),
#                      or a WALL THICKNESS in inches (a number, e.g. 0.375) which overrides the
#                      schedule table, or "CAL" (calculate — falls back to the minimum schedule).
# Moving to a new project normally means two edits only:
#   · rename the class (e.g. "G1A-3" → "A1S6") — the name must match 'Piping Spec' on Pipeline_Input;
#   · change the schedule / wall thickness of its rules.
# 'Piping Spec' may carry a one-off override:  "G1A-3/100"  (schedule 100)  or  "G1A-3/0.5"  (wall 0.5 in).
# Source: GEMS catalogue extract (253 classes) with the G1S-1 / G1S-2 curated overrides.
PMS_CLASSES: dict[str, tuple] = {
    'G10Z-K': ('Type 304 Tubing (c,f)', ()),
    'G10Z-L': ('Type 316 Tubing (c,f)', ()),
    'G10Z-P': ('Allowance:', ()),
    'G10Z-Q': ('Allowance:', ()),
    'G12U-1': ('Nodular (Ductile) Iron,', (('4', '24', '350'),)),
    'G12U-2': ('High Density Polyethylene', (('1/2', '2', '80'), ('1/2', '2', '80'), ('2', '24', 'CL 200'))),
    'G15A-1': ('Carbon Steel', (('3/4', '3', '160'), ('4', '24', 'CAL'))),
    'G15A-11': ('Carbon Steel', (('3/4', '2', 'XXS'), ('3', '24', 'CAL'))),
    'G15A-15': ('Carbon Steel', (('1', '24', 'CAL'),)),
    'G15A-17': ('Carbon Steel (21)', (('1', '24', 'CAL'),)),
    'G15A-2': ('Carbon Steel', (('3/4', '3', '160'), ('4', '24', 'CAL'))),
    'G15A-4': ('Carbon Steel', (('3/4', '2', 'XXS'), ('3', '24', '-'))),
    'G15A-5': ('Carbon Steel', (('3/4', '2', 'XXS'), ('3', '24', 'CAL'))),
    'G15A-7': ('Carbon Steel (21)', (('3/4', '2', 'XXS'), ('3', '24', 'CAL'))),
    'G15B-2': ('0.125" (0.10" min.) (26)', (('3/4', '2', '160'), ('2', '2', 'XXS'), ('3', '12', '-'), ('3', '12', 'CAL'))),
    'G15F-1': ('1 1/4% Cr-1/2% Mo', (('3/4', '3', '160'), ('4', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G15F-2': ('1 1/4% Cr-1/2% Mo', (('3/4', '3', '160'), ('4', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G15G-1': ('2 1/4% Cr-1% Mo', (('3/4', '4', '160'), ('6', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G15G-2': ('2 1/4% Cr-1% Mo', (('1/2', '4', '160'), ('6', '24', 'CAL'), ('6', '24', 'CAL'), ('14', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G15H-1': ('5% Cr-1/2% Mo  (21)', (('3/4', '3', '160'), ('4', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G15H-2': ('5% Cr-1/2% Mo  (21)', (('3/4', '3', '160'), ('4', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G15J-1': ('9% Cr-1% Mo (21)', (('3/4', '3', '160'), ('4', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G15J-2': ('9% Cr-1% Mo  (21)', (('3/4', '3', '160'), ('4', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G15K-1': ('Type 304', ()),
    'G15K-10': ('Type 304 (76)', (('3/4', '4', '80S'), ('6', '8', 'CAL'), ('6', '8', 'CAL'), ('6', '8', 'CAL'))),
    'G15K-2': ('Type 304', (('3/4', '1-1/2', '80S'), ('2', '3', '160'), ('4', '4', '120'), ('6', '24', 'CAL'), ('6', '24', 'CAL'), ('6', '24', 'CAL'), ('6', '24', 'CAL'))),
    'G15K-3': ('Type 304L', (('3/4', '1', '80S'), ('1-1/2', '3', '160'), ('4', '8', '120'), ('10', '24', 'CAL'), ('10', '24', 'CAL'), ('10', '24', 'CAL'))),
    'G15L-4': ('Type 316L', (('3/4', '1', '160'), ('1-1/2', '1-1/2', 'XXH'), ('3/4', '1', '80S'), ('1-1/2', '3', '160'), ('4', '12', 'CAL'), ('4', '12', 'CAL'), ('4', '12', 'CAL'), ('4', '12', 'CAL'))),
    'G15N-1': ('Type 321', (('3/4', '1', '80S'), ('1-1/2', '3', '160'), ('4', '24', 'CAL'), ('4', '24', 'CAL'), ('4', '24', 'CAL'), ('4', '24', 'CAL'))),
    'G15N-2': ('Type 321', (('3/4', '1', '80S'), ('1-1/2', '3', '160'), ('4', '24', 'CAL'), ('4', '24', 'CAL'), ('4', '24', 'CAL'), ('4', '24', 'CAL'))),
    'G15N-6': ('Type 347', (('3/4', '1', '80S'), ('1-1/2', '3', '160'), ('4', '6', '120'), ('8', '8', '120'), ('8', '8', '120'), ('10', '10', 'CAL'), ('10', '10', 'CAL'), ('10', '10', 'CAL'))),
    'G15R-2': ('Alloy 825 (UNS N08825)', (('3/4', '2', '160'), ('3', '3', '80S'), ('4', '8', 'CAL'), ('10', '12', 'CAL'))),
    'G1A-1': ('Carbon Steel', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '24', 'Std'), ('26', '48', 'CAL'), ('26', '48', 'CAL'))),
    'G1A-10': ('Carbon Steel', ()),
    'G1A-11': ('Carbon Steel', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '24', 'Std'))),
    'G1A-13': ('Carbon Steel', (('3/4', '1-1/2', 'XXS'), ('2', '3', '160'), ('4', '8', '80'), ('10', '24', 'CAL'), ('26', '48', 'CAL'), ('26', '48', 'CAL'))),
    'G1A-15': ('Carbon Steel', (('3/4', '1-1/2', 'XXS'), ('2', '3', '160'), ('4', '8', '80'), ('10', '24', 'CAL'))),
    'G1A-17': ('Carbon Steel (21)', (('3/4', '1-1/2', 'XXS'), ('2', '3', '160'), ('4', '8', '80'), ('10', '24', 'CAL'))),
    'G1A-2': ('Carbon Steel', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '24', 'Std'))),
    'G1A-21': ('Carbon Steel', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '24', 'Std'), ('26', '48', 'CAL'), ('26', '48', 'CAL'))),
    'G1A-22': ('Carbon Steel', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '24', 'Std'))),
    'G1A-23': ('Carbon Steel', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '24', 'Std'), ('26', '48', 'CAL'), ('26', '48', 'CAL'))),
    'G1A-25': ('Carbon Steel', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '24', 'Std'))),
    'G1A-26': ('Carbon Steel   (21)', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '24', 'Std'))),
    'G1A-27': ('Carbon Steel  (21)', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '24', 'Std'))),
    'G1A-28': ('Carbon Steel', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '24', 'Std'))),
    'G1A-3': ('Carbon Steel', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '24', 'Std'), ('26', '48', 'CAL'), ('26', '48', 'CAL'))),
    'G1A-35': ('Carbon Steel', (('1', '1-1/2', 'XXS'), ('2', '3', '160'), ('4', '6', '80'), ('10', '24', 'CAL'))),
    'G1A-4': ('Carbon Steel', ()),
    'G1A-40': ('Carbon Steel (g)', (('1', '2', '160'), ('3', '4', '80'), ('6', '16', 'Std'))),
    'G1A-5': ('Carbon Steel', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '24', 'Std'))),
    'G1A-50': ('Carbon Steel - Lokring\uf8ea (#1)', ()),
    'G1A-6': ('Carbon Steel   (21)', (('3/4', '1', '160'),)),
    'G1A-7': ('Carbon Steel  (21)', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '24', 'Std'))),
    'G1A-8': ('Carbon Steel', (('3/4', '1', '160'), ('1-1/2', '2', '80'))),
    'G1A-9': ('Carbon Steel', (('3/4', '1-1/2', 'XXS'), ('2', '3', '160'), ('4', '6', '80'), ('8', '8', '60'), ('10', '18', '40'))),
    'G1B-1': ('0.125" (0.10" min.)', (('3/4', '2', '80S'), ('2', '2', '80'), ('3', '24', 'Std'))),
    'G1B-21': ('0.125" (0.10" min.)', (('3/4', '1-1/2', '80S'), ('2', '2', '80'), ('2', '2', '80S'), ('3', '24', 'Std'))),
    'G1B-3': ('0.125" (0.10" min.)', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '24', 'Std'))),
    'G1C-1': ('Aluminum (a)', ()),
    'G1D-1': ('Carbon Steel – PTFE Lined', (('1', '6', '40'), ('8', '10', '30'))),
    'G1D-2': ('Carbon Steel – PVDF Lined', (('1', '6', '40'), ('8', '8', '30'))),
    'G1D-3': ('Carbon Steel – PTFE Lined', (('1', '4', '40'), ('2', '4', '40'), ('installation,', 'installation,', 'inspection,'), ('joint', 'joint', 'make-up)'), ('the Owner’s', 'the Owner’s', 'Engineer'), ('Integrity', 'Integrity', 'Flange (HIF™)'), ('of and', 'of and', 'are provided'), ('torque', 'torque', 'and general'), ('type torque', 'type torque', 'wrenches'))),
    'G1F-1': ('1 1/4%Cr-1/2% Mo', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '10', '40'), ('12', '24', 'Std'), ('12', '24', 'Std'))),
    'G1F-21': ('1 1/4%Cr-1/2% Mo', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '10', '40'), ('12', '24', 'Std'), ('12', '24', 'Std'))),
    'G1G-1': ('2 1/4% Cr-1% Mo', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '10', '40'), ('12', '24', 'Std'), ('12', '24', 'Std'))),
    'G1G-21': ('2 1/4% Cr-1% Mo', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '10', '40'), ('12', '24', 'Std'), ('12', '24', 'Std'))),
    'G1H-1': ('5% Cr-1/2% Mo    (21)', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '10', '40'), ('12', '24', 'Std'), ('12', '24', 'Std'))),
    'G1H-21': ('5% Cr-1/2% Mo   (21)', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '10', '40'), ('12', '24', 'Std'), ('12', '24', 'Std'))),
    'G1J-1': ('9% Cr-1% Mo   (21)', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '10', '40'), ('12', '24', 'Std'), ('12', '24', 'Std'))),
    'G1J-21': ('9% Cr-1% Mo   (21)', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '10', '40'), ('12', '24', 'Std'), ('12', '24', 'Std'))),
    'G1K-1': ('Type 304', (('3/4', '1-1/2', '80S'), ('2', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G1K-10': ('Type 304 (76)', (('3/4', '1-1/2', '80S'), ('2', '4', '40S'), ('6', '8', '40S'), ('6', '8', 'CAL'))),
    'G1K-14': ('Type 304', (('3/4', '1-1/2', '80S'), ('2', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G1K-3': ('Type 304L', (('3/4', '1-1/2', '80S'), ('2', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G1K-4': ('Type 304', (('3/4', '1-1/2', '80S'), ('2', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G1K-40': ('Type 304', (('1', '1-1/2', '80S'), ('2', '16', '40S'))),
    'G1K-5': ('Type 304', (('3/4', '1-1/2', '80S'), ('2', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G1K-6': ('Type 304L', (('1/2', '6', '40S'), ('1/2', '6', '40S'))),
    'G1L-1': ('Type 316', (('3/4', '1-1/2', '80S'), ('2', '2', '40S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G1L-10': ('Type 316 (76)', ()),
    'G1L-3': ('Type 316L', (('3/4', '1-1/2', '80S'), ('2', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G1L-4': ('Type 316', (('3/4', '1-1/2', '80S'), ('2', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G1L-50': ('Type 316 - Lokring\uf8ea (#1)', ()),
    'G1M-1': ('Type 317L', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('3', '6', '40S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G1M-2': ('0.063" (0.05" min.)', (('3/4', '1-1/2', '80S'), ('3/4', '1-1/2', '40S'), ('2', '6', '40S'), ('2', '6', '40S'), ('8', '24', '-'), ('8', '24', '-'))),
    'G1M-3': ('0.063" (0.05" min.)', (('3/4', '1-1/2', '80S'), ('3/4', '1-1/2', '40S'), ('2', '6', '40S'), ('2', '6', '40S'), ('8', '24', '-'), ('8', '24', '-'))),
    'G1N-1': ('Type 321', (('3/4', '2', '80S'), ('3', '6', '40S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G1N-10': ('AL-6XN (UNS N08367)', (('3/4', '2', '40S'), ('3', '10', '10S'), ('3', '10', '10S'))),
    'G1P-1': ('Alloy 400 (UNS N04400)', (('3/4', '2', '80S'), ('3', '6', '40S'), ('8', '12', 'CAL'), ('8', '12', 'CAL'))),
    'G1P-2': ('Alloy 400 (UNS N04400)', (('3/4', '2', '80S'), ('3', '6', '40S'), ('8', '12', 'CAL'), ('8', '12', 'CAL'))),
    'G1P-40': ('Alloy 400 (UNS N04400)', (('1', '2', '80S'), ('3', '16', '40S'))),
    'G1P-5': ('Nickel 200 (UNS N02200)', (('3/4', '1-1/2', '80S'), ('2', '2', '40S'), ('2', '2', '40S'), ('3', '12', 'CAL'), ('3', '12', 'CAL'))),
    'G1Q-10': ('Alloy C-276 (UNS N10276)', (('3/4', '1-1/2', '80S'), ('3/4', '2', '40S'), ('3', '3', '40S'), ('3', '3', '40S'), ('4', '12', 'CAL'))),
    'G1Q-11': ('Alloy C-276 (UNS N10276)', (('3/4', '1-1/2', '80S'), ('3/4', '2', '40S'), ('3', '3', '40S'), ('3', '3', '40S'), ('4', '12', 'CAL'), ('4', '8', 'CAL'))),
    'G1Q-40': ('Alloy 600 (UNS N06600)', (('1', '2', '80S'), ('3', '16', '40S'))),
    'G1R-2': ('Alloy 825 (UNS N08825)', (('3/4', '1-1/2', '80S'), ('2', '8', '40S'), ('10', '18', 'CAL'))),
    'G1S-1': ('Alloy 20 (UNS N08020)', (('3/4', '1-1/2', '80S'), ('2', '6', '40S'), ('2', '6', '40S'), ('8', '18', 'CAL'))),
    'G1S-2': ('Alloy 20 (UNS N08020)', (('3/4', '1-1/2', '80S'), ('2', '6', '40S'), ('2', '6', '40S'), ('8', '18', 'CAL'))),
    'G1T-1': ('Titanium (UNS R50400) (c)', (('2', '16', 'CAL'), ('2', '16', 'CAL'))),
    'G1U-1': ('Carbon Steel', (('1/2', '2', '80'), ('1/2', '2', '80'), ('3', '12', 'Std'))),
    'G1U-10': ('Carbon Steel', (('1/2', '2', '40'), ('1/2', '2', '40'))),
    'G1U-14': ('Carbon Steel', (('1/2', '1-1/2', 'XXS'), ('2', '3', '160'), ('4', '6', '80'), ('8', '16', '40'), ('18', '24', 'CAL'), ('26', '48', 'CAL'), ('26', '48', 'CAL'))),
    'G1U-15': ('Carbon Steel', (('1/2', '1-1/2', 'XXS'), ('2', '3', '160'), ('4', '6', '80'), ('8', '24', 'CAL'))),
    'G1U-2': ('Carbon Steel', (('1/2', '2', '80'), ('1/2', '2', '80'), ('3', '3', '80'), ('3', '3', '80'), ('3', '12', 'Std'))),
    'G1U-3': ('Carbon Steel', (('1/2', '2', '80'), ('1/2', '2', '80'), ('2', '2', '80'), ('3', '24', 'Std'))),
    'G1U-4': ('Carbon Steel', (('1/2', '2', '80'), ('3', '24', 'Std'), ('26', '48', 'CAL'), ('26', '48', 'CAL'))),
    'G1U-5': ('Carbon Steel', (('1/2', '2', '80'), ('1/2', '2', '80'), ('3', '24', 'Std'), ('3', '24', 'Std'), ('3', '24', 'Std'), ('3', '24', 'Std'))),
    'G1U-7': ('Carbon Steel', (('1/2', '2', '80'), ('3', '24', 'Std'))),
    'G1U-8': ('Carbon Steel', (('1/2', '2', '80'), ('1/2', '2', '80'), ('3', '4', '40'), ('3', '4', '40'), ('6', '24', 'Std'))),
    'G1U-9': ('Carbon Steel', (('3/4', '2', '80'), ('3', '3', '40'), ('GP 03-06-01', 'GP 03-06-01', 'or engineering'), ('temperatures', 'temperatures', 'less'))),
    'G25G-2': ('2 1/4% Cr-1% Mo', (('1/2', '3', 'XXS'), ('4', '12', 'CAL'), ('4', '12', 'CAL'))),
    'G3A-1': ('Carbon Steel', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'), ('26', '48', 'CAL'), ('26', '48', 'CAL'))),
    'G3A-10': ('Carbon Steel', ()),
    'G3A-15': ('Carbon Steel', (('3/4', '1-1/2', 'XXS'), ('2', '3', '160'), ('4', '4', '120'), ('6', '8', '80'), ('10', '24', 'CAL'))),
    'G3A-17': ('Carbon Steel  (21)', (('3/4', '1-1/2', 'XXS'), ('2', '3', '160'), ('4', '4', '120'), ('6', '8', '80'), ('10', '24', 'CAL'))),
    'G3A-2': ('Carbon Steel', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'))),
    'G3A-21': ('Carbon Steel', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'), ('26', '48', 'CAL'), ('26', '48', 'CAL'))),
    'G3A-22': ('Carbon Steel', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'))),
    'G3A-23': ('Carbon Steel', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'), ('26', '48', 'CAL'), ('26', '48', 'CAL'))),
    'G3A-25': ('Carbon Steel', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'))),
    'G3A-26': ('Carbon Steel   (21)', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'))),
    'G3A-27': ('Carbon Steel  (21)', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'))),
    'G3A-28': ('Carbon Steel', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'))),
    'G3A-3': ('Carbon Steel', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'), ('26', '48', 'CAL'), ('26', '48', 'CAL'))),
    'G3A-4': ('Carbon Steel', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '16', '40'), ('18', '24', '-'), ('26', '48', '-'))),
    'G3A-40': ('Carbon Steel', (('1', '2', '160'), ('3', '4', '80'), ('6', '10', 'Std'), ('12', '16', 'CAL'))),
    'G3A-45': ('Carbon Steel', (('1', '1', '160'), ('1-1/2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'), ('16', '24', 'CAL'))),
    'G3A-46': ('Carbon Steel', (('1', '1', '160'), ('1-1/2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'), ('16', '24', 'CAL'))),
    'G3A-5': ('Carbon Steel', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'))),
    'G3A-50': ('Carbon Steel - Lokring\uf8ea (#1)', ()),
    'G3A-6': ('Carbon Steel   (21)', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'))),
    'G3A-7': ('Carbon Steel  (21)', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'))),
    'G3A-8': ('Carbon Steel', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'))),
    'G3B-1': ('0.125" (0.10" min.)', (('3/4', '2', '80S'), ('2', '2', '80'), ('3', '16', '40'), ('18', '24', '-'))),
    'G3B-2': ('0.125" (0.10" min.)', (('3/4', '3/4', '160'), ('1', '2', '80'), ('3', '14', '40'), ('16', '24', '-'))),
    'G3B-21': ('0.125" (0.10" min.)', (('3/4', '1-1/2', '80S'), ('2', '2', '80'), ('3', '16', '40'), ('18', '24', '-'))),
    'G3B-22': ('0.125" (0.10" min.)', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '14', '40'), ('16', '24', '-'))),
    'G3B-3': ('0.125" (0.10" min.)', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '16', '40'), ('18', '24', '-'))),
    'G3D-1': ('Carbon Steel – PTFE Lined', (('1', '6', '40'), ('8', '8', '30'))),
    'G3F-1': ('1 1/4% Cr-1/2% Mo', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'), ('18', '24', 'CAL'))),
    'G3F-21': ('1 1/4% Cr-1/2% Mo', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'), ('18', '24', 'CAL'))),
    'G3G-1': ('2 1/4% Cr-1% Mo', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'), ('18', '24', 'CAL'))),
    'G3G-21': ('2 1/4% Cr-1% Mo', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '16', '40'), ('18', '24', 'CAL'), ('18', '24', 'CAL'))),
    'G3H-1': ('5% Cr-1/2% Mo    (21)', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '10', '40'), ('12', '12', 'XS'), ('14', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G3H-2': ('5% Cr-1/2% Mo (21)', (('3/4', '3/4', '160'), ('1', '2', '80'), ('3', '10', '40'), ('12', '24', 'CAL'), ('12', '24', 'CAL'))),
    'G3H-21': ('5% Cr-1/2% Mo (21)', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '10', '40'), ('12', '12', 'XS'), ('14', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G3H-22': ('5% Cr-1/2% Mo  (21)', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '10', '40'), ('12', '24', 'CAL'), ('12', '24', 'CAL'))),
    'G3J-1': ('9% Cr-1% Mo   (21)', (('3/4', '1', '160'), ('1-1/2', '2', '80'), ('3', '10', '40'), ('12', '12', 'XS'), ('14', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G3J-2': ('9% Cr-1% Mo   (21)', (('3/4', '3/4', '160'), ('1', '2', '80'), ('3', '10', '40'), ('12', '24', 'CAL'), ('12', '24', 'CAL'))),
    'G3J-21': ('9% Cr-1% Mo   (21)', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '10', '40'), ('12', '12', 'XS'), ('14', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G3J-22': ('9%Cr-1% Mo   (21)', (('3/4', '1-1/2', '160'), ('2', '2', '80'), ('3', '10', '40'), ('12', '24', 'CAL'), ('12', '24', 'CAL'))),
    'G3K-1': ('Type 304', (('3/4', '1-1/2', '80S'), ('2', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G3K-10': ('Type 304 (76)', (('3/4', '1-1/2', '80S'), ('2', '4', '40S'), ('6', '8', '40S'))),
    'G3K-3': ('Type 304L', (('3/4', '1-1/2', '80S'), ('2', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G3K-4': ('Type 304', (('3/4', '2', '80S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G3K-40': ('Type 304', (('1', '1-1/2', '80S'), ('2', '16', '40S'))),
    'G3K-5': ('Type 304', (('3/4', '2', '80S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G3K-6': ('Type 304', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('3', '10', '40S'), ('12', '12', 'Std'), ('12', '12', 'Std'), ('14', '24', 'CAL'), ('14', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G3K-7': ('Type 304H', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('3', '10', '40S'), ('12', '16', 'CAL'), ('12', '16', 'CAL'), ('12', '16', 'CAL'))),
    'G3K-8': ('Type 304H', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('3', '10', '40S'), ('12', '16', 'CAL'), ('12', '16', 'CAL'), ('12', '16', 'CAL'))),
    'G3L-1': ('Type 316', (('3/4', '1-1/2', '80S'), ('2', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G3L-3': ('Type 316L', (('3/4', '1-1/2', '80S'), ('2', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G3L-50': ('Type 316 - Lokring\uf8ea (#1)', ()),
    'G3L-7': ('Type 316H', (('3/4', '1-1/2', '80S'), ('3/4', '6', '40S'), ('8', '16', 'CAL'), ('8', '16', 'CAL'), ('8', '16', 'CAL'))),
    'G3L-8': ('Type 316H', (('3/4', '1-1/2', '80S'), ('3/4', '6', '40S'), ('8', '16', 'CAL'), ('8', '16', 'CAL'), ('8', '16', 'CAL'))),
    'G3M-1': ('Type 317L', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('3', '6', '40S'), ('3', '6', '40S'), ('8', '24', '-'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G3M-2': ('0.063" (0.05" min.)', (('3/4', '1-1/2', '80S'), ('3/4', '1-1/2', '40S'), ('2', '6', '40S'), ('2', '6', '40S'), ('8', '24', '-'), ('8', '24', '-'))),
    'G3M-3': ('0.063" (0.05" min.)', (('3/4', '1-1/2', '80S'), ('3/4', '1-1/2', '40S'), ('2', '6', '40S'), ('2', '6', '40S'), ('8', '24', '-'), ('8', '24', '-'))),
    'G3N-1': ('Type 321', (('3/4', '2', '80S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G3N-6': ('Type 347', (('3/4', '1-1/2', '80S'), ('3/4', '2', '40S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G3N-7': ('Type 347', (('3/4', '1-1/2', '80S'), ('3/4', '2', '40S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G3P-1': ('Alloy 400 (UNS N04400)', (('3/4', '2', '80S'), ('3', '6', '40S'), ('8', '12', 'CAL'), ('8', '12', 'CAL'))),
    'G3P-2': ('Alloy 400 (UNS N04400)', (('3/4', '2', '80S'), ('3', '6', '40S'), ('8', '12', 'CAL'), ('8', '12', 'CAL'))),
    'G3P-40': ('Alloy 400 (UNS N04400)', (('1', '2', '80S'), ('3', '16', '40S'))),
    'G3P-45': ('Alloy 400 (UNS N04400)', (('1', '2', '80S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G3P-46': ('Alloy 400 (UNS N04400)', (('1', '2', '80S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G3P-47': ('Alloy 400 (UNS N04400) (21)', (('1', '2', '80S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', '-'))),
    'G3Q-10': ('Alloy C-276 (UNS N10276)', (('3/4', '1-1/2', '80S'), ('3/4', '2', '40S'), ('3', '3', '40S'), ('3', '3', '40S'), ('4', '12', 'CAL'), ('4', '8', 'CAL'))),
    'G3Q-11': ('Alloy C-276 (UNS N10276)', (('3/4', '1-1/2', '80S'), ('3/4', '2', '40S'), ('3', '3', '40S'), ('3', '3', '40S'), ('4', '12', 'CAL'), ('4', '8', 'CAL'))),
    'G3Q-40': ('Alloy 600 (UNS N06600)', (('1', '2', '80S'), ('3', '16', '40S'))),
    'G3R-2': ('Alloy 825 (UNS N08825)', (('3/4', '1-1/2', '80S'), ('2', '8', '40S'), ('10', '18', 'CAL'))),
    'G6A-1': ('Carbon Steel', (('3/4', '2', '160'), ('3', '8', '80'), ('10', '24', 'CAL'))),
    'G6A-15': ('Carbon Steel', (('3/4', '2', 'XXS'), ('3', '4', '160'), ('6', '6', '120'), ('8', '24', 'CAL'))),
    'G6A-17': ('Carbon Steel (21)', (('3/4', '2', 'XXS'), ('3', '4', '160'), ('6', '6', '120'), ('8', '24', 'CAL'))),
    'G6A-21': ('Carbon Steel', (('3/4', '2', '160'), ('3', '24', '80'))),
    'G6A-23': ('Carbon Steel', (('3/4', '2', '160'), ('3', '24', '80'))),
    'G6A-25': ('Carbon Steel', (('3/4', '1-1/2', 'XXS'), ('2', '2', '160'), ('3', '24', '80'))),
    'G6A-27': ('Carbon Steel  (21)', (('3/4', '1-1/2', 'XXS'), ('2', '2', '160'), ('3', '24', '80'))),
    'G6A-3': ('Carbon Steel', (('3/4', '2', '160'), ('3', '8', '80'), ('10', '24', 'CAL'))),
    'G6A-4': ('Carbon Steel', (('3/4', '2', '160'), ('3', '8', '80'), ('10', '24', '-'))),
    'G6A-5': ('Carbon Steel', (('3/4', '2', '160'), ('3', '8', '80'), ('10', '24', 'CAL'))),
    'G6A-50': ('Carbon Steel - Lokring\uf8ea (#1)', ()),
    'G6A-7': ('Carbon Steel  (21)', (('3/4', '2', '160'), ('3', '8', '80'), ('10', '24', 'CAL'))),
    'G6B-1': ('0.125" (0.10" min.)', (('3/4', '2', '80S'), ('2', '2', '160'), ('3', '8', '80'), ('10', '24', '-'))),
    'G6B-21': ('0.125" (0.10" min.)', (('3/4', '1-1/2', '80S'), ('2', '8', '80'), ('10', '24', '-'))),
    'G6B-3': ('0.125" (0.10" min.)', (('3/4', '2', '160'), ('3', '8', '80'), ('10', '24', '-'))),
    'G6F-1': ('1 1/4% Cr-1/2% Mo', (('3/4', '2', '160'), ('3', '6', '80'), ('8', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G6F-2': ('1 1/4% Cr-1/2% Mo', (('3/4', '2', '160'), ('3', '6', '80'), ('8', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G6F-21': ('1 1/4% Cr-1/2% Mo', (('3/4', '1-1/2', '160'), ('2', '6', '80'), ('8', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G6F-22': ('1 1/4% Cr-1/2% Mo', (('3/4', '1-1/2', '160'), ('2', '6', '80'), ('8', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G6G-1': ('2 1/4% Cr-1% Mo', (('3/4', '1', '160'), ('1-1/2', '6', '80'), ('8', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G6G-21': ('2 1/4% Cr-1% Mo', (('3/4', '1-1/2', '160'), ('2', '6', '80'), ('8', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G6H-1': ('5% Cr-1/2% Mo    (21)', (('3/4', '2', '160'), ('3', '8', '80'), ('10', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G6H-2': ('5% Cr-1/2% Mo    (21)', (('3/4', '2', '160'), ('3', '8', '80'), ('10', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G6H-21': ('5% Cr-1/2% Mo    (21)', (('3/4', '1-1/2', '160'), ('2', '8', '80'), ('10', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G6H-22': ('5% Cr-1/2% Mo    (21)', (('3/4', '1-1/2', '160'), ('2', '8', '80'), ('10', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G6J-1': ('9% Cr-1% Mo   (21)', (('3/4', '2', '160'), ('3', '8', '80'), ('10', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G6J-2': ('9% Cr-1% Mo  (21)', (('3/4', '2', '160'), ('3', '8', '80'), ('10', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G6J-21': ('9% Cr-1% Mo   (21)', (('3/4', '2', '160'), ('3', '8', '80'), ('10', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G6J-22': ('9% Cr-1% Mo   (21)', (('3/4', '2', '160'), ('3', '8', '80'), ('10', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G6K-1': ('Type 304', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('3', '4', '40S'), ('6', '24', 'CAL'), ('6', '24', 'CAL'), ('6', '24', 'CAL'), ('6', '24', 'CAL'))),
    'G6K-10': ('Type 304 (76)', (('3/4', '1-1/2', '80S'), ('2', '4', '40S'), ('6', '8', 'CAL'), ('6', '8', 'CAL'))),
    'G6K-3': ('Type 304L', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G6K-4': ('Type 304', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('3', '4', '40S'), ('6', '24', 'CAL'), ('6', '24', 'CAL'), ('6', '24', 'CAL'))),
    'G6K-5': ('Type 304', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G6L-1': ('Type 316', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('3', '4', '40S'), ('6', '24', 'CAL'), ('6', '24', 'CAL'), ('6', '24', 'CAL'), ('6', '24', 'CAL'))),
    'G6L-3': ('Type 316L', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G6L-50': ('Type 316 - Lokring\uf8ea (#1)', ()),
    'G6M-3': ('0.063" (0.05" min.)', (('3/4', '1-1/2', '80S'), ('3/4', '1-1/2', '40S'), ('2', '4', '40S'), ('2', '4', '40S'), ('6', '24', '-'), ('6', '24', '-'))),
    'G6N-1': ('Type 321', (('3/4', '2', '80S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G6N-6': ('Type 347', (('3/4', '1-1/2', '80S'), ('3/4', '2', '40S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G6N-7': ('Type 347', (('3/4', '1-1/2', '80S'), ('3/4', '2', '40S'), ('3', '6', '40S'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'), ('8', '24', 'CAL'))),
    'G6P-1': ('Alloy 400 (UNS N04400)', (('3/4', '2', '80S'), ('3', '6', '40S'), ('8', '12', 'CAL'), ('8', '12', 'CAL'))),
    'G6P-40': ('Alloy 400 (UNS N04400)', (('1', '2', '80S'), ('3', '6', '40S'), ('8', '16', 'CAL'))),
    'G6R-2': ('Alloy 825 (UNS N08825)', (('3/4', '1-1/2', '160'), ('3/4', '2', '40S'), ('3', '4', '40S'), ('6', '8', 'CAL'), ('10', '18', 'CAL'))),
    'G9A-1': ('Carbon Steel', (('3/4', '1-1/2', '160'), ('3/4', '2', '80'), ('3', '4', '80'), ('6', '24', 'CAL'))),
    'G9A-11': ('Carbon Steel', (('3/4', '1-1/2', 'XXS'), ('3/4', '2', '160'), ('3', '3', '160'), ('4', '6', '120'), ('8', '24', 'CAL'))),
    'G9A-15': ('Carbon Steel', (('1', '4', 'XXS'), ('6', '6', '160'), ('8', '24', 'CAL'))),
    'G9A-3': ('Carbon Steel', (('3/4', '1-1/2', '160'), ('3/4', '2', '80'), ('2', '4', '80'), ('6', '24', 'CAL'))),
    'G9A-4': ('Carbon Steel', (('3/4', '1-1/2', 'XXS'), ('3/4', '2', '160'), ('3', '3', '160'), ('4', '6', '120'), ('8', '24', '-'))),
    'G9A-5': ('Carbon Steel', (('3/4', '1-1/2', 'XXS'), ('3/4', '2', '160'), ('3', '3', '160'), ('4', '6', '120'), ('8', '24', 'CAL'))),
    'G9A-7': ('Carbon Steel  (21)', (('3/4', '1-1/2', 'XXS'), ('3/4', '2', '160'), ('3', '3', '160'), ('4', '6', '120'), ('8', '24', 'CAL'))),
    'G9B-1': ('0.063" (0.05" min.)', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('2', '4', '80'), ('6', '24', '-'))),
    'G9F-1': ('1 1/4% Cr-1/2% Mo', (('3/4', '1-1/2', '160'), ('3/4', '2', '80'), ('3', '3', '80'), ('4', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G9F-2': ('1 1/4% Cr-1/2% Mo', (('3/4', '1-1/2', '160'), ('3/4', '2', '80'), ('3', '4', '80'), ('6', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G9G-2': ('2 1/4% Cr-1% Mo (73)', (('3/4', '1-1/2', '160'), ('3/4', '2', '80'), ('3', '4', '80'), ('6', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G9G-3': ('2 1/4% Cr-1% Mo', (('1/2', '4', '80'), ('6', '24', 'CAL'), ('6', '24', 'CAL'), ('14', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G9H-1': ('5% Cr-1/2% Mo  (21)', (('3/4', '1-1/2', '160'), ('3/4', '2', '80'), ('3', '3', '80'), ('4', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G9H-2': ('5% Cr-1/2% Mo    (21)', (('3/4', '1-1/2', '160'), ('3/4', '2', '80'), ('3', '3', '80'), ('4', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G9J-1': ('9% Cr-1% Mo   (21)', (('3/4', '1-1/2', '160'), ('3/4', '2', '80'), ('3', '3', '80'), ('4', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G9J-2': ('9% Cr-1% Mo (21)', (('3/4', '1-1/2', '160'), ('3/4', '2', '80'), ('3', '3', '80'), ('4', '24', 'CAL'), ('14', '24', 'CAL'))),
    'G9K-1': ('Type 304', (('3/4', '1-1/2', '160'),)),
    'G9K-2': ('Type 304', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('3', '4', '80S'), ('6', '24', 'CAL'), ('6', '24', 'CAL'), ('6', '24', 'CAL'), ('6', '24', 'CAL'))),
    'G9K-3': ('Type 304L', ()),
    'G9K-4': ('Type 304', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('3', '3', '80S'), ('4', '24', 'CAL'), ('4', '24', 'CAL'), ('4', '24', 'CAL'))),
    'G9K-40': ('Type 304', (('3/4', '2', '80S'), ('3', '4', '80S'), ('6', '12', 'CAL'))),
    'G9L-2': ('Type 316', (('1/2', '2', '80S'), ('3', '3', '40S'), ('4', '4', '80S'), ('6', '24', 'CAL'), ('6', '24', 'CAL'), ('6', '24', 'CAL'), ('6', '24', 'CAL'))),
    'G9M-3': ('0.063" (0.05" min.)', (('3/4', '1-1/2', '160'), ('3/4', '1-1/2', '40S'), ('2', '2', '80S'), ('3', '3', '40S'), ('4', '24', '-'), ('4', '24', '-'))),
    'G9N-1': ('Type 321', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('3', '3', '40S'), ('4', '24', 'CAL'), ('4', '24', 'CAL'), ('4', '24', 'CAL'), ('4', '24', 'CAL'))),
    'G9N-2': ('Type 321', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('3', '3', '40S'), ('4', '24', 'CAL'), ('4', '24', 'CAL'), ('4', '24', 'CAL'), ('4', '24', 'CAL'))),
    'G9N-6': ('Type 347', (('3/4', '1-1/2', '160'), ('3/4', '2', '80S'), ('3', '3', '40S'), ('4', '24', 'CAL'), ('4', '24', 'CAL'), ('4', '24', 'CAL'), ('4', '24', 'CAL'))),
    'G9R-2': ('Alloy 825 (UNS N08825)', (('3/4', '1-1/2', '160'), ('3/4', '1', '80S'), ('1-1/2', '4', '80S'), ('6', '8', 'CAL'), ('10', '12', 'CAL'))),
}


class _PmsRule:
    __slots__ = ("nps_low", "nps_high", "schedule", "min_schedule")

    def __init__(self, nps_low, nps_high, schedule, min_schedule="STD"):
        self.nps_low, self.nps_high = nps_low, nps_high
        self.schedule, self.min_schedule = schedule, min_schedule


class _PmsClass:
    __slots__ = ("name", "moc", "pipe_rules")

    def __init__(self, name, moc, rules):
        self.name, self.moc = name, moc
        self.pipe_rules = tuple(_PmsRule(*r) for r in rules)


_CLASS_BY_NAME = {k.upper(): _PmsClass(k, v[0], v[1]) for k, v in PMS_CLASSES.items()}

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


# ASME B36.10M 8" schedules missing from the base table (wall, in) — verify before design use
PIPE_WALL_IN["8"].update({"20": 0.250, "30": 0.277, "60": 0.406, "100": 0.594,
                          "120": 0.719, "140": 0.812})


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

    # Optional schedule override:  "<class>/<schedule>"  e.g.  G9A-3/100  or  G9A-3/SCH80
    spec_raw = (spec or "").strip()
    sched_ov = None
    if "/" in spec_raw:
        spec_raw, _ov = spec_raw.split("/", 1)
        _o = _ov.strip().upper().replace("SCH", "").strip()
        sched_ov = (_o if re.fullmatch(r"\d*\.\d+", _o) else _o.replace(".", "")) or None
        spec = spec_raw.strip()
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
    if sched_ov:
        sched = sched_ov
        basis += f" Sch {sched_ov} (override)"
    if sched is None:
        sched = "STD"
        basis += " (spec n/a -> STD)"
    wall_given = None
    try:                                   # a number = wall thickness in inches
        w_ = float(str(sched).replace("IN", "").strip())
        if 0.01 < w_ < 3.0:
            wall_given = w_
    except ValueError:
        pass
    if wall_given is not None:
        if od - 2 * wall_given <= 0:
            return IDResult(None, od, wall_given, f"{wall_given:g}\"", nps, basis + " (wall > OD/2)")
        return IDResult(round(od - 2 * wall_given, 4), od, wall_given, f"t={wall_given:g}\"", nps,
                        basis + f" wall {wall_given:g}\" (thickness)")
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
    import hmbbuilder as HMBPROII          # optional — PRO/II case-sheet exports
except ImportError:                        # pragma: no cover
    HMBPROII = None

# ════════════════════════════════════════════════════════════════════════
#  STREAM MAP  — actual (plant) stream number  →  Simulation Stream Number in the HMB
# ════════════════════════════════════════════════════════════════════════
# No-ISO input : type the ACTUAL stream number in 'Stream Lookup'; the 'Stream_Map' sheet of the
#                input workbook (Stream Number | Simulation Stream Number) translates it to the
#                stream name used in the HMB.  A stream not listed is looked up by its own name.
# ISO / PCF    : stream_map.xlsx (next to this file) keys the line (Fluid Code + Sequence No +
#                Material Class) to the same two stream numbers.
# Both files are write-once templates — never overwritten.
import types as _types

_SM_FILENAME = "stream_map.xlsx"
_SM_SHEET = "STREAM_MAP"
_SM_HEADERS = ["Fluid Code", "Sequence No", "Material Class", "Stream Number", "Simulation Stream Number"]
STREAM_MAP_INPUT_SHEET = "Stream_Map"


@dataclass(frozen=True)
class _SmRow:
    fluid_code: str
    seq_no: str
    material_class: str
    stream_number: str          # plant / actual name
    sim_stream_number: str      # id present in the HMB (may be "")

    @property
    def lookup_key(self) -> str:
        """Preferred key into the HMB."""
        return self.sim_stream_number or self.stream_number


class StreamMapError(RuntimeError):
    """Raised when a line cannot be resolved (map-only, no interactive fallback)."""


def _sm_norm(v) -> str:
    return re.sub(r"\s+", " ", str(v).strip()).upper() if v is not None else ""


def _sm_norm_seq(v) -> str:
    s = "" if v is None else str(v).strip()
    if not s:
        return ""
    m = re.match(r"^0*(\d+)(?:\.0+)?$", s)
    return m.group(1) if m else s.upper()


def _sm_parse_line_no(line_no: str) -> tuple[str, str]:
    """'ER-162' -> ('ER','162').  Tolerates 'ER 162', 'ER162', 'ER-162-H12'."""
    s = str(line_no).strip()
    m = re.match(r"^\s*([A-Za-z]+)[\s\-_]*0*(\d+)", s)
    if m:
        return m.group(1).upper(), m.group(2)
    parts = re.split(r"[\s\-_]+", s, maxsplit=1)
    return (parts[0].upper(), parts[1] if len(parts) > 1 else "")


def _sm_default_path(base_dir=None) -> Path:
    return (Path(base_dir) if base_dir else Path(__file__).resolve().parent) / _SM_FILENAME


def _sm_ensure_template(path=None) -> Path:
    """Create the write-once stream_map.xlsx template if absent."""
    p = Path(path) if path else _sm_default_path()
    if p.exists():
        return p
    wb = Workbook()
    ws = wb.active
    ws.title = _SM_SHEET
    ws.sheet_view.showGridLines = False
    for ci, (h, w) in enumerate(zip(_SM_HEADERS, [14, 13, 16, 22, 24]), start=1):
        ws.column_dimensions[get_column_letter(ci)].width = w
        c = ws.cell(row=1, column=ci, value=h)
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4973")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28
    for ci, v in enumerate(["ER", "162", "G1A-5", "T-801 OH", "T801-OH"], start=1):
        c = ws.cell(row=2, column=ci, value=v)
        c.font = Font(name="Calibri", size=10, italic=True, color="A6A6A6")
        c.fill = PatternFill("solid", fgColor="FFF2CC")
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=4, column=1, value="↑ example row — edit or delete, then add your own.  Leave 'Simulation "
            "Stream Number' blank if the plant stream is not in the HMB.").font = Font(
        name="Calibri", size=9, italic=True, color="A6A6A6")
    ws.freeze_panes = "A2"
    wb.save(p)
    return p


def _sm_load_map(path=None) -> list:
    p = Path(path) if path else _sm_default_path()
    if not p.exists():
        return []
    wb = load_workbook(p, read_only=True, data_only=True)
    try:
        ws = wb[_SM_SHEET] if _SM_SHEET in wb.sheetnames else wb.worksheets[0]
        rows = []
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r is None:
                continue
            fc, seq, mc, sn, sim = (list(r) + [None] * 5)[:5]
            if not (fc and seq and mc):
                continue
            rows.append(_SmRow(str(fc).strip(), str(seq).strip(), str(mc).strip(),
                               str(sn).strip() if sn else "", str(sim).strip() if sim else ""))
    finally:
        wb.close()
    return rows


def _sm_resolve(fluid_code, seq_no, material_class, path=None):
    fc, sq, mc = _sm_norm(fluid_code), _sm_norm_seq(seq_no), _sm_norm(material_class)
    for row in _sm_load_map(path):
        if (_sm_norm(row.fluid_code) == fc and _sm_norm_seq(row.seq_no) == sq
                and _sm_norm(row.material_class) == mc):
            return row
    return None


def _sm_resolve_from_line(line_no, material_class, path=None):
    fc, seq = _sm_parse_line_no(line_no)
    return _sm_resolve(fc, seq, material_class, path)


def _sm_require(line_no, material_class, path=None):
    """Resolve or raise with an actionable message."""
    p = Path(path) if path else _sm_default_path()
    if not p.exists():
        _sm_ensure_template(p)
        raise StreamMapError(f"stream_map not found — created template at {p}. "
                             f"Fill in the row for {line_no} / {material_class} and re-run.")
    row = _sm_resolve_from_line(line_no, material_class, p)
    if row is None:
        fc, seq = _sm_parse_line_no(line_no)
        raise StreamMapError(f"No stream_map entry for Fluid Code='{fc}', Sequence No='{seq}', "
                             f"Material Class='{material_class}' in {p}. Add a row and re-run.")
    if not row.lookup_key:
        raise StreamMapError(f"stream_map row for {line_no}/{material_class} has neither a "
                             f"Stream Number nor a Simulation Stream Number.")
    return row


SMAP = _types.SimpleNamespace(
    MapRow=_SmRow, StreamMapError=StreamMapError, ensure_template=_sm_ensure_template,
    load_map=_sm_load_map, resolve=_sm_resolve, resolve_from_line=_sm_resolve_from_line,
    require=_sm_require, default_path=_sm_default_path, parse_line_no=_sm_parse_line_no)


def _read_stream_map_sheet(input_path) -> dict:
    """'Stream_Map' sheet of the input workbook → {NORM(actual stream number): simulation stream}."""
    out: dict = {}
    try:
        wb = load_workbook(input_path, read_only=True, data_only=True)
    except Exception:
        return out
    try:
        if STREAM_MAP_INPUT_SHEET not in wb.sheetnames:
            return out
        for r in wb[STREAM_MAP_INPUT_SHEET].iter_rows(min_row=3, values_only=True):
            if not r or r[0] is None or not str(r[0]).strip():
                continue
            sim = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ""
            if sim:
                out[_sm_norm(r[0])] = sim
    finally:
        wb.close()
    return out


def _stream_alias(input_path, map_path=None) -> dict:
    """{NORM(actual stream number): simulation stream} from the input workbook's Stream_Map sheet,
    then (lower priority) the legacy stream_map.xlsx rows."""
    out: dict = {}
    try:
        for row in _sm_load_map(map_path):
            if row.stream_number and row.sim_stream_number:
                out[_sm_norm(row.stream_number)] = row.sim_stream_number
    except Exception:
        pass
    out.update(_read_stream_map_sheet(input_path))
    return out


_ALIAS_NOTED: set = set()


def _alias_stream(name, alias):
    """Simulation stream name for an actual stream number (unchanged when not mapped)."""
    if not name or not alias:
        return name
    sim = alias.get(_sm_norm(name))
    if sim and sim != name:
        if name not in _ALIAS_NOTED:
            _ALIAS_NOTED.add(name)
            print(f"  stream map: '{name}'  ->  simulation stream '{sim}'")
        return sim
    return name


def _build_stream_map_sheet(wb) -> None:
    ws = wb.create_sheet(STREAM_MAP_INPUT_SHEET)
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:C1")
    _c(ws, 1, 1,
       "STREAM MAP — actual (plant) stream number → Simulation Stream Number in the HMB.  "
       "Type the ACTUAL stream number in 'Stream Lookup' on Pipeline_Input; it is translated here.  "
       "A stream not listed is looked up by its own name.",
       bg=NAVY, fg=WHITE, sz=10, bold=True, wrap=True)
    ws.row_dimensions[1].height = 42
    _hdr(ws, ["Stream Number (actual)", "Simulation Stream Number", "Description"], row=2)
    ex = [("T-801 OH", "T801-OH", "example — T-801 overhead"), ("D-703 CV OUT", "D703-CV-OUT", "example")]
    for i, rowv in enumerate(ex):
        for c, v in enumerate(rowv, 1):
            _c(ws, 3 + i, c, v, bg=AMBER, sz=9, italic=True, fg="808080")
    for r in range(5, 60):
        for c in range(1, 4):
            _c(ws, r, c, None, bg=AMBER, sz=9)
    for c, w in enumerate([26, 28, 40], 1):
        cw(ws, c, w)
    ws.freeze_panes = "A3"

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
    # Per-phase momentum flux & vapour Mach (SI; None when not computable)
    mach_vap: float | None = None       # v_mix / c_vapour
    rhov2_vap: float | None = None      # ρ_g · v_sg²   (kg/m·s²)
    rhov2_liq: float | None = None      # ρ_l · v_sl²   (kg/m·s²)
    props: dict | None = None           # stream properties at this station
    instr: dict | None = None           # control-valve / inline-instrument capture (kind, tag, downstream flash…)


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
    ph = _station_phase_state(station, sp, gas_density_fn)
    rho_gas_si = ph["rho_g"] * LBFT3_TO_KGM3
    vm = ph["mg"] * LBHR_TO_KGS
    lm = ph["ml"] * LBHR_TO_KGS
    q_gas = vm / rho_gas_si if rho_gas_si > 0 else 0.0
    rho_l_si = ph["rho_l"] * LBFT3_TO_KGM3
    q_liq = lm / rho_l_si if rho_l_si > 0 else 0.0
    gvf = q_gas / (q_gas + q_liq) if (q_gas + q_liq) > 0 else 1.0
    mu_gas = (ph["mu_g"] or 0.0) * CP_TO_PAS
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
        "Mach (vapor)", "ρv² Vapor (kg/m·s²)", "ρv² Liquid (kg/m·s²)",
    ]
    n_base = len(base_headers)               # 19
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
            round(s.mach_vap, 4)  if s.mach_vap  is not None else None,
            round(s.rhov2_vap, 1) if s.rhov2_vap is not None else None,
            round(s.rhov2_liq, 1) if s.rhov2_liq is not None else None,
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

    base_widths = [5, 9, 16, 6, 9, 9, 8, 12, 11, 14, 8, 8, 11, 12, 11, 12,
                   11, 16, 16]
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
        "  ρv² Vapor / Liquid = ρ_phase · (superficial velocity of that phase)²;  "
        "Mach (vapor) = mixture velocity / √(k·Z·R·T/MW), k = Vap Cp/Cv.",
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

    ph = _station_phase_state(station, sp, gas_density_fn)
    rho_g = ph["rho_g"] * LBFT3_TO_KGM3
    rho_l = ph["rho_l"] * LBFT3_TO_KGM3
    mu_g = (ph["mu_g"] or 0.01) * CP_TO_PAS
    mu_l = (ph["mu_l"] or 0.5) * CP_TO_PAS
    sigma = (ph["sigma"] or 20.0) * DYNCM_TO_NM

    mg = ph["mg"] * LBHR_TO_KGS      # kg/s
    ml = ph["ml"] * LBHR_TO_KGS
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
    _apply_light_theme(wb)
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


_INSTR_ALIASES = {
    "PT": "PT", "PI": "PT", "PRESSURE TRANSMITTER": "PT",
    "TT": "TT", "TE": "TT", "TW": "TT", "TEMPERATURE TRANSMITTER": "TT", "THERMOWELL": "TT",
    "FT": "FT", "FT-I": "FT", "FT INLINE": "FT", "FT (INLINE)": "FT", "INLINE FT": "FT",
    "FLOW TRANSMITTER": "FT", "MAG": "FT", "CORIOLIS": "FT", "VORTEX": "FT",
    "FO": "FO", "FT-O": "FO", "FT ORIFICE": "FO", "FT (ORIFICE)": "FO", "ORIFICE": "FO",
    "FLOW ORIFICE": "FO", "FE": "FO",
    "RO": "RO", "RESTRICTION ORIFICE": "RO",
}


def _instr_kind(fitting: str | None, instr_type: str | None) -> str | None:
    """Classify a row: 'CV' control valve, 'PT'/'TT'/'FT' (inline), 'FO' (orifice FT), 'RO', or None."""
    if (fitting or "").lower().strip() == "control valve":
        return "CV"
    it = (instr_type or "").upper().strip()
    if it in _INSTR_ALIASES:
        return _INSTR_ALIASES[it]
    if (fitting or "").lower().strip() == "orifice":
        return "FO"
    return None


def _is_orifice(fitting: str | None, instr_type: str | None) -> bool:
    return _instr_kind(fitting, instr_type) in ("FO", "RO")


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
    "Liq MW",                 # liquid molecular weight (g/mol) — manual streams
    "Liq Surf Tension (N/m)", # liquid surface tension, N/m (flow-pattern maps)
    "Vap Cp (BTU/lb-F)",      # vapour Cp — used for the enthalpy balance at tee joins
    "Liq Cp (BTU/lb-F)",      # liquid Cp — used for the enthalpy balance at tee joins
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
    "Set P (psia)",           # boundary pressure for Source / Destination rows; on a Control Valve row = desired OUTLET pressure (ΔP back-calculated if Fixed dP is blank)
    "Control Valve Type",     # P | T | F | L  (Control Valve rows only)
    "Exch Max Allow dP (psi)",  # T-control floor: max allowable dP across exchanger
    "Notes",                  # free notes
]
_NCOL = len(INPUT_HEADERS)   # 37

# Manual stream-property columns (styled yellow in the template; used when no stream)
_MANUAL_PROP_COLS = [
    "Temp (degF)", "Vapor Mass Flow", "Vap MW", "Vap Visc (cP)", "Vap Z",
    "Vap Cp/Cv", "Vap Density (lb/ft3)", "Liq Mass Flow (lb/h)",
    "Liq Density (lb/ft3)", "Liq Visc (cP)", "Liq MW", "Liq Surf Tension (N/m)",
    "Vap Cp (BTU/lb-F)", "Liq Cp (BTU/lb-F)",
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
    "Liq MW": ("Liq MW", None),
    "Liq Surf Tension (N/m)": ("Liq Surf Tension", None),
    "Vap Cp (BTU/lb-F)": ("Vap Cp", None),
    "Liq Cp (BTU/lb-F)": ("Liq Cp", None),
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
def _build_composition_input_sheet(wb, usys) -> None:
    """Stream_Composition input sheet: one column per phase of each inlet stream
    (see the block comment above ``_read_stream_composition``)."""
    ws = wb.create_sheet(_COMP_SHEET)
    ws.sheet_view.showGridLines = False
    n_ex = 4                                   # example columns
    ncol = 1 + max(n_ex, 8)                    # leave room for more streams
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    _c(ws, 1, 1,
       "STREAM COMPOSITION  —  use when NOT referencing an HMB stream.  One COLUMN per phase "
       "of each inlet stream; Seq says where it enters the line.  Streams entering at "
       "different Seq (or in extra columns) are MIXED.  Component names = HMB 'COMPONENTS' sheet.",
       bg=NAVY, fg=WHITE, sz=10, bold=True, wrap=True)
    ws.row_dimensions[1].height = 40
    labels = ["Circuit (optional)", "Line No", "Seq", "Comp ID (optional)",
              "Phase", "Basis", f"Phase Flow ({usys.label('mflow')}, mass)"]
    for i, lab in enumerate(labels):
        _c(ws, 2 + i, 1, lab, bg=NAVY, fg=WHITE, sz=9, bold=True)
    _c(ws, 9, 1, "Component", bg=GRNHDR, fg=WHITE, sz=9, bold=True)
    for ci in range(2, ncol + 1):
        _c(ws, 9, ci, "amount ↓", bg=GRNHDR, fg=WHITE, sz=9, bold=True, ha="center")
        for r in range(2, 9):
            _c(ws, r, ci, None, bg=AMBER, sz=9)
    cw(ws, 1, 30)
    for ci in range(2, ncol + 1):
        cw(ws, ci, 14)
    ws.freeze_panes = "B10"

    comps = ["H2", "METHANE", "ETHANE", "PROPANE", "NITROGEN", "H2O", "H2S", "CO2"]
    for i in range(60):                        # component-name rows (free to edit)
        _c(ws, 10 + i, 1, comps[i] if i < len(comps) else None, bg=WHITE, sz=9)
        for ci in range(2, ncol + 1):
            _c(ws, 10 + i, ci, None, bg=AMBER, sz=9)

    # Examples: Circuit C2 / Line H-285.  Gas A enters at Seq 1 (flow taken from
    # the Pipeline_Input row, 12000 lb/hr); gas B (own flow) is mixed in at Seq 2.
    flow_a = round(usys.from_internal("mflow", 12000.0), 3)
    flow_b = round(usys.from_internal("mflow", 3000.0), 3)
    ex = [
        ("C2", "H-285", 1, None, "Vapor", "Mass", None,
         {"H2": 0.10, "METHANE": 0.60, "ETHANE": 0.30}),
        ("C2", "H-285", 2, None, "Vapor", "Mole", flow_b,
         {"NITROGEN": 0.80, "CO2": 0.20}),
    ]
    for k, (cir, ln, sq, cid, ph, bs, fl, comp) in enumerate(ex):
        ci = 2 + k
        for r, v in zip(range(2, 9), (cir, ln, sq, cid, ph, bs, fl)):
            _c(ws, r, ci, v, bg=AMBER, sz=9, italic=True, fg="808080")
        for r in range(10, 70):
            nm = ws.cell(r, 1).value
            if nm in comp:
                _c(ws, r, ci, comp[nm], bg=AMBER, sz=9, italic=True, fg="808080")

    # Component-name suggestions (pure HMB components + built-in library).  The
    # list is advisory only — any HMB COMP_CONSTANTS name may be typed.
    ws_cl = wb.create_sheet("_CompList")
    ws_cl.sheet_state = "hidden"
    suggest = list(dict.fromkeys(
        ["H2O", "H2S", "NH3", "H2", "NITROGEN", "METHANE", "ETHANE", "PROPANE",
         "IC4", "NC4", "IC5", "NC5", "CYCLOPENTANE", "ISOHEXANE", "CO2", "HCL",
         "SO2", "SO3", "CL2", "O2", "CO"] + [r[0] for r in _PURE_LIB_ROWS]))
    for i, nm in enumerate(suggest, 1):
        ws_cl.cell(row=i, column=1, value=nm)
    dv_cn = DataValidation(type="list", formula1=f"_CompList!$A$1:$A${len(suggest)}",
                           allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(dv_cn)
    dv_cn.add("A10:A200")

    dv_ph = DataValidation(type="list", formula1='"Vapor,Liquid"', allow_blank=True)
    dv_bs = DataValidation(type="list", formula1='"Mass,Mole"', allow_blank=True)
    ws.add_data_validation(dv_ph)
    ws.add_data_validation(dv_bs)
    last = get_column_letter(ncol + 40)
    dv_ph.add(f"B6:{last}6")
    dv_bs.add(f"B7:{last}7")


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
        "Liq MW": 9, "Liq Surf Tension (N/m)": 14, "Vap Cp (BTU/lb-F)": 12,
        "Liq Cp (BTU/lb-F)": 12,
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

    # ── Drop-down: Instr Type (PT / TT / FT / FO / RO) ────────────────
    it_col = _hcol_letter("Instr Type")
    dv_it = DataValidation(
        type="list", formula1='"PT,TT,FT,FO,RO"', allow_blank=True,
        showErrorMessage=False)
    ws.add_data_validation(dv_it)
    dv_it.add(f"{it_col}3:{it_col}2000")

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

    # ── Stream_Composition sheet (manual mass/mole composition, no HMB) ──
    _build_composition_input_sheet(wb, usys)
    _build_pseudo_input_sheet(wb, usys)
    _build_stream_map_sheet(wb)
    _build_flow_cases_sheet(wb)
    _build_instrument_data_sheet(wb)

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
        ("MANUAL COMPOSITION  (sheet 'Stream_Composition' — use INSTEAD of an HMB stream)", True),
        ("  Each COLUMN is one phase of one inlet stream.  Header rows: Circuit (optional), Line No,", False),
        ("  Seq (where the stream enters), Comp ID (optional, if Seq repeats on a line), Phase", False),
        ("  (Vapor/Liquid), Basis (Mass/Mole — fractions or flows, normalised), Phase Flow (mass).", False),
        ("  Component rows use the EXACT names of the HMB 'COMPONENTS' sheet (MW/Tc/Pc/ω are read", False),
        ("  from the HMB 'COMP_CONSTANTS' sheet, so set HMB File on the row or pass the HMB on the CLI).", False),
        ("  PSEUDO-COMPONENTS: define any extra component on the 'Pseudo_Components' sheet from MW, NBP", False),
        ("  and SG (Tc/Pc via Exxon-style correlation, ω via Edmister); then use its Name here.", False),
        ("  Names missing from the HMB fall back to a built-in PURE library (benzene, toluene, xylenes,", False),
        ("  styrene, cumene, trimethylbenzenes, butylbenzene, naphthalene, methylnaphthalene, biphenyl,", False),
        ("  NC6-NC16, cyclohexanes, olefins, He/Ar, methanol/ethanol/phenol — see _PURE_LIB_ROWS).", False),
        ("  Phase Flow blank → the Vapor / Liq Mass Flow cell of that Pipeline_Input row is used.", False),
        ("  Still type Temp, Vap Z, Vap Visc, Vap Cp/Cv, Liq Density, Liq Visc on that row (yellow);", False),
        ("  MW and the vapour/liquid split come from the composition.", False),
        ("  TEMPERATURE: composition streams are marched ADIABATICALLY (isenthalpic, PR enthalpy):", False),
        ("  the temperature falls/rises across every fitting with its pressure drop (Joule-Thomson),", False),
        ("  carried to all downstream stations.  Use  --isothermal-comp  on the command line to disable.", False),
        ("  MIXING: the first composition on a circuit starts the flow.  Every later column (a later Seq,", False),
        ("  or another column at the same Seq) is MIXED in: component flows add, the blend is re-flashed", False),
        ("  and the mixed MW / properties carry downstream.  Provide BOTH a vapour and a liquid column", False),
        ("  for true K = y/x; with one phase only, Wilson K from the HMB constants is used.", False),
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
        ("FLOW CASES  (Flow_Cases sheet)", True),
        ("  Normal 100 %, Maximum 110 %, Minimum 60 % by default — edit or add cases.  Every inlet flow is scaled and the whole", False),
        ("  hydraulics is repeated per case: one output workbook per case; 'CV Datasheet' shows all cases side by side.", False),
        ("", False),
        ("STREAM MAP  (Stream Lookup)", True),
        ("  Type the ACTUAL stream number in Stream Lookup; the 'Stream_Map' sheet translates it to the", False),
        ("  Simulation Stream Number used in the HMB.  A stream that is not listed is looked up by its own name.", False),
        ("", False),
        ("INSTRUMENTS  (Instr Type = PT / TT / FT / FO / RO)", True),
        ("  PT / TT: pressure / temperature transmitter taps — no pressure drop; listed on 'Inline Instruments'.", False),
        ("  TT also gets a thermowell wake-frequency check (enter thermowell data in 'Instrument_Data').", False),
        ("  FT = inline flow meter (rated dP in Instr dP).  FO = orifice flow element.  RO = restriction orifice (rated dP in Instr dP).", False),
        ("  Use the Instr Tag column as the tag; 'Instrument_Data' rows are matched on that tag.", False),
        ("  Control valves get their own 'CV Datasheet' sheet; every line gets a 'Line Data' row.", False),
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
        ("", False),
        ("CONTROL VALVE ΔP FROM SET PRESSURE", True),
        ("  Fitting = 'Control Valve': leave Fixed dP BLANK and type the desired OUTLET pressure in", False),
        ("  'Set P (psia)' on the same row — the engine back-calculates ΔP = P_upstream − Set P.", False),
        ("  (A typed Fixed dP wins; Set P ≥ upstream pressure = valve fully open, ΔP = 0.)", False),
        ("  Composition streams cool/heat across the valve (isenthalpic) and carry that downstream.", False),
        ("", False),
        ("TEE JOINS (temperature)", True),
        ("  When two composition streams (or a Branch) join, the mixed temperature comes from an", False),
        ("  ENTHALPY BALANCE (PR enthalpy) — not a mass-weighted average.", False),
        ("", False),
        ("AUTO-LINKED SPLITS & JOINS (Main ↔ Branch)", True),
        ("  JOIN (Branch → Main): on a Main row, set Fitting = 'Tee Join Branch Flow N' and (optional)", False),
        ("  type the joining Branch's Line No in 'Upstream Line No' of that row to name it explicitly.", False),
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
        next((s for s in wb.worksheets if s.title not in ("Notes", "_FittingList", _COMP_SHEET,
                                                "Pseudo_Components", "_CompList")),
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
            "_Up_defaulted":    bool(not ul_raw and upstream_line is not None),
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
            "Liq MW":               num(g(rv, "Liq MW")),
            "Liq Surf Tension (N/m)": num(g(rv, "Liq Surf Tension (N/m)")),
            "Vap Cp (BTU/lb-F)":    num(g(rv, "Vap Cp (BTU/lb-F)")),
            "Liq Cp (BTU/lb-F)":    num(g(rv, "Liq Cp (BTU/lb-F)")),
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
    _attach_stream_composition(rows, _read_stream_composition(wb, usys))
    _USER_COMPONENTS.clear()
    _USER_COMPONENTS.update(_read_pseudo_components(wb, usys))
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
    rho_g = _gas_rho_sp(sp, fr.vap_mw, sp.temp_f, fr.p_psia)
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


def _gas_rho_sp(sp, vap_mw, T_F, P_psia, z=None) -> float:
    """Gas density kg/m3.  A manual-property stream with a typed ``Vap Density`` uses it
    (at the row's Start P) scaled with local pressure; otherwise MW·P/(Z·R·T)."""
    if (sp is not None and getattr(sp, "stream", None) == "(manual)"
            and sp.vap_density and sp.pres_psia and P_psia):
        return sp.vap_density * LBFT3_TO_KGM3 * (P_psia / sp.pres_psia)
    return _gas_rho_kgm3(vap_mw, z if z is not None else (sp.vap_z if sp else None),
                         T_F, P_psia)


def _phase_momentum(g: HE.Geom, sp: HE.StreamProps, fr, v_mix_ms: float) -> dict:
    """Vapour Mach number and per-phase ρv² at one station (SI).

    ρ_g·v_sg² and ρ_l·v_sl² use each phase's SUPERFICIAL velocity (phase
    volumetric flow / pipe area), so for a single-phase stream the relevant
    one equals the mixture ρv².  Mach = homogeneous mixture velocity divided
    by the vapour sound speed  c = √(k·Z·R·T / MW)  (k = Cp/Cv).
    Any term whose inputs are missing is returned as None.
    """
    out = dict(mach_vap=None, rhov2_vap=None, rhov2_liq=None)
    if fr is None or g.area_m2 <= 0:
        return out
    vm = fr.vap_mass * LBHR_TO_KGS
    lm = fr.liq_mass * LBHR_TO_KGS
    rho_g = _gas_rho_sp(sp, fr.vap_mw, sp.temp_f, fr.p_psia)
    rho_l = (sp.liq_density or 0.0) * LBFT3_TO_KGM3
    if vm > 0 and rho_g > 0:
        v_sg = vm / rho_g / g.area_m2
        out["rhov2_vap"] = rho_g * v_sg * v_sg
        k = sp.vap_cp_cv
        if k and k > 0 and fr.vap_mw and sp.temp_f is not None:
            t_k = (sp.temp_f + F_TO_K_OFFSET) / 1.8
            c = math.sqrt(k * (sp.vap_z or 1.0) * R_UNIV * t_k
                          / (fr.vap_mw / 1000.0))
            out["mach_vap"] = v_mix_ms / c if c > 0 else None
    elif vm <= 0:
        out["rhov2_vap"] = 0.0
    if lm > 0 and rho_l > 0:
        v_sl = lm / rho_l / g.area_m2
        out["rhov2_liq"] = rho_l * v_sl * v_sl
    elif lm <= 0:
        out["rhov2_liq"] = 0.0
    return out


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
        mol_weight  = vap_mw if vap_mw else num(row.get("Liq MW")),
        vap_mw      = vap_mw,
        liq_mw      = num(row.get("Liq MW")),
        vap_density = vap_den,
        liq_density = liq_den,
        vap_visc    = num(row.get("Vap Visc (cP)")),
        liq_visc    = num(row.get("Liq Visc (cP)")),
        vap_z       = num(row.get("Vap Z")),
        vap_cp_cv   = num(row.get("Vap Cp/Cv")),
        vap_cp      = num(row.get("Vap Cp (BTU/lb-F)")),
        liq_cp      = num(row.get("Liq Cp (BTU/lb-F)")),
        liq_surf_tens = ((num(row.get("Liq Surf Tension (N/m)")) * 1000.0)   # N/m → dyn/cm
                         if num(row.get("Liq Surf Tension (N/m)")) is not None else None),
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
    if len(names) < 1 or not p_ref:
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
# ════════════════════════════════════════════════════════════════════════
#  Manual stream COMPOSITION  (no HMB stream) — Stream_Composition sheet
# ════════════════════════════════════════════════════════════════════════
# Each column of the sheet is ONE phase of ONE inlet stream:
#
#     Circuit | Line No | Seq | Comp ID | Phase | Basis | Phase Flow | <component amounts…>
#
# * Seq (+ Line No) picks the Pipeline_Input row where the stream ENTERS.
#   Comp ID (optional) disambiguates when a Seq number repeats on a line.
# * Phase = Vapor | Liquid ; Basis = Mass | Mole (fractions OR flows — normalised).
# * Phase Flow = MASS flow of that phase (blank → the row's Vapor Mass Flow /
#   Liq Mass Flow cell is used).
# * The FIRST composition stream on a circuit starts the flow.  Every later
#   composition column (new Seq, or the same Seq in another column) is MIXED
#   into the running stream: component flows add, the blend is re-flashed.
# Component names must match the HMB COMPONENTS / COMP_CONSTANTS sheet — that
# sheet also supplies MW / Tc / Pc / ω.
_COMP_SHEET = "Stream_Composition"
_COMP_HDR_KEYS = (            # (label prefix in col A, key) — longest first
    ("phase flow", "flow"), ("comp id", "comp_id"), ("line no", "line"),
    ("circuit", "circuit"), ("phase", "phase"), ("basis", "basis"),
    ("seq", "seq"),
)


def _read_stream_composition(wb, usys) -> list[dict]:
    """Parse the Stream_Composition sheet → list of column dicts
    {circuit, line, seq, comp_id, phase('vap'|'liq'), basis('mass'|'mole'),
     flow (lb/hr or None), amounts {name: value}, col (letter)}."""
    if _COMP_SHEET not in wb.sheetnames:
        return []
    ws = wb[_COMP_SHEET]
    grid = [list(r) for r in ws.iter_rows(values_only=True)]
    if not grid:
        return []
    ncols = max(len(r) for r in grid)
    hdr: dict[str, int] = {}          # key → grid row index
    comp_start = None
    for ri, r in enumerate(grid):
        a = txt(r[0]) if r else None
        if not a:
            continue
        al = a.lower()
        if al == "component":
            comp_start = ri + 1
            break
        for pref, key in _COMP_HDR_KEYS:
            if al.startswith(pref):
                hdr.setdefault(key, ri)
                break
    if comp_start is None or "line" not in hdr or "phase" not in hdr:
        return []

    def cell(key, ci):
        ri = hdr.get(key)
        if ri is None or ci >= len(grid[ri]):
            return None
        return grid[ri][ci]

    cols: list[dict] = []
    for ci in range(1, ncols):
        line = txt(cell("line", ci))
        ph = (txt(cell("phase", ci)) or "").lower()
        if not line or not ph:
            continue
        amounts: dict[str, float] = {}
        for r in grid[comp_start:]:
            nm = txt(r[0]) if r else None
            v = num(r[ci]) if (nm and ci < len(r)) else None
            if nm and v is not None and v > 0:
                amounts[nm] = amounts.get(nm, 0.0) + v
        if not amounts:
            continue
        seq_v = num(cell("seq", ci))
        flow_v = num(cell("flow", ci))
        basis = (txt(cell("basis", ci)) or "mass").lower()
        cols.append({
            "circuit": txt(cell("circuit", ci)),
            "line":    line,
            "seq":     int(seq_v) if seq_v is not None else None,
            "comp_id": txt(cell("comp_id", ci)),
            "phase":   "liq" if ph.startswith("l") else "vap",
            "basis":   "mole" if basis.startswith("mol") else "mass",
            "flow":    (usys.to_internal("mflow", flow_v)
                        if flow_v is not None else None),
            "amounts": amounts,
            "col":     get_column_letter(ci + 1),
        })
    return cols


def _attach_stream_composition(rows: list[dict], cols: list[dict]) -> None:
    """Attach each composition column to the FIRST matching input row (same
    Line No / Seq [/ Circuit] [/ Comp ID]) and make that row's phase-flow cells
    consistent with the composition flows."""
    for col in cols:
        hit = next((r for r in rows
                    if r["Line No"] == col["line"]
                    and (col["seq"] is None or r["Seq"] == col["seq"])
                    and (col["circuit"] is None or r["Circuit"] == col["circuit"])
                    and (col["comp_id"] is None
                         or (r.get("Comp ID") or "") == col["comp_id"])),
                   None)
        if hit is None:
            print(f"  WARNING: {_COMP_SHEET} column {col['col']} (Line "
                  f"{col['line']}, Seq {col['seq']}) matches no Pipeline_Input "
                  f"row — ignored.")
            continue
        hit.setdefault("_comp", []).append(col)

    for r in rows:
        cl = r.get("_comp")
        if not cl:
            continue
        for ph, cell_key in (("vap", "Vapor Mass Flow"),
                             ("liq", "Liq Mass Flow (lb/h)")):
            pcols = [c for c in cl if c["phase"] == ph]
            if not pcols:
                continue
            blank = [c for c in pcols if c["flow"] is None]
            if blank:
                if len(pcols) == 1 and r.get(cell_key):
                    blank[0]["flow"] = r[cell_key]
                else:
                    print(f"  WARNING: {_COMP_SHEET} column {blank[0]['col']} has "
                          f"no Phase Flow (and no single {cell_key} on the "
                          f"row) — column ignored.")
                    for c in blank:
                        c["flow"] = 0.0
            r[cell_key] = sum(c["flow"] or 0.0 for c in pcols)
        # a phase with no composition column is NOT part of the composition
        # stream: zero its flow so the row cell can't leak into the mix
        for ph, cell_key in (("vap", "Vapor Mass Flow"),
                             ("liq", "Liq Mass Flow (lb/h)")):
            if not any(c["phase"] == ph for c in cl):
                r[cell_key] = None


def _wilson_k(tc_r: float, pc: float, omega: float, t_r: float, p: float) -> float:
    return (pc / p) * math.exp(WILSON_C * (1.0 + omega) * (1.0 - tc_r / t_r))


def _comp_payload(row: dict, p_ref: float, const_map: dict) -> dict:
    """Build the inlet stream described by ``row['_comp']`` at pressure ``p_ref``.

    Returns {"feed": FlashFeed, "sp": StreamProps, "moles": {name: lb-mol/hr}}.
    Raises ValueError (with an actionable message) on unknown components.

    Equilibrium ratios: true K = y/x when both a vapour AND a liquid
    composition are supplied; otherwise Wilson K from the HMB COMP_CONSTANTS
    (anchored to the stated vapour fraction if Wilson disagrees with it).
    """
    if not const_map:
        raise ValueError(
            "composition input needs component data: the HMB workbook must have "
            "a COMP_CONSTANTS sheet (MW, Tc, Pc, ω).  Set 'HMB File' on the row "
            "(or pass the HMB on the command line).")
    vmol: dict[str, float] = {}
    lmol: dict[str, float] = {}
    mw: dict[str, float] = {}
    for col in row["_comp"]:
        flow = col.get("flow") or 0.0
        if flow <= 0:
            continue
        amt: dict[str, float] = {}
        for nm, a in col["amounts"].items():
            c = const_map.get(nm.strip().upper())
            if c is None or not c.get("mw"):
                raise ValueError(
                    f"component '{nm}' ({_COMP_SHEET} col {col['col']}) is not in "
                    f"the HMB COMP_CONSTANTS sheet — use the exact HMB component "
                    f"name.")
            cn = c.get("name") or nm.strip().upper()
            mw[cn] = c["mw"]
            amt[cn] = amt.get(cn, 0.0) + a
        tot = sum(amt.values())
        if col["basis"] == "mass":
            mol = {n: flow * a / tot / mw[n] for n, a in amt.items()}
        else:
            avg_mw = sum(a / tot * mw[n] for n, a in amt.items())
            ntot = flow / avg_mw
            mol = {n: ntot * a / tot for n, a in amt.items()}
        tgt = vmol if col["phase"] == "vap" else lmol
        for n, m in mol.items():
            tgt[n] = tgt.get(n, 0.0) + m

    names = list(dict.fromkeys(list(vmol) + list(lmol)))
    if not names:
        raise ValueError(f"no usable composition on Line {row.get('Line No')} "
                         f"Seq {row.get('Seq')} (check Phase Flow and amounts).")
    nv, nl = sum(vmol.values()), sum(lmol.values())
    n_i = {n: vmol.get(n, 0.0) + lmol.get(n, 0.0) for n in names}
    N = nv + nl
    vap_mass = sum(vmol.get(n, 0.0) * mw[n] for n in names)
    liq_mass = sum(lmol.get(n, 0.0) * mw[n] for n in names)
    vmw = (vap_mass / nv) if nv > 0 else None
    lmw = (liq_mass / nl) if nl > 0 else None

    base = _streamprops_from_manual(row)
    if base is None:
        base = HE.StreamProps(stream="(manual)", source_sheet="manual-input",
                              temp_f=num(row.get("Temp (degF)")))
    sp = replace(
        base, stream="(manual composition)", source_sheet="manual-input",
        phase=("TWO-PHASE" if (vap_mass > 0 and liq_mass > 0)
               else "VAPOR" if vap_mass > 0 else "LIQUID"),
        pres_psia=p_ref, total_mass=vap_mass + liq_mass,
        vap_mass=vap_mass, liq_mass=liq_mass, vap_mw=vmw, liq_mw=lmw,
        mol_weight=vmw or lmw)
    # ── equilibrium ratios ───────────────────────────────────────────────
    z = [n_i[n] / N for n in names]
    x_d = {n: lmol[n] / nl for n in lmol} if nl > 0 else {}
    y_d = {n: vmol[n] / nv for n in vmol} if nv > 0 else {}
    both = sum(1 for n in names if n in vmol and n in lmol)
    tc_r, omega, pc_l = [], [], []
    for n in names:
        c = const_map.get(n.upper(), {})
        tc_r.append(c["tc_f"] + F_TO_R if c.get("tc_f") is not None else None)
        omega.append(c.get("omega"))
        pc_l.append(c.get("pc_psia"))

    if both >= 2:
        feed = _build_feed(
            {n: n_i[n] / N for n in names}, x_d,
            {n: n_i[n] * mw[n] for n in names}, n_i, sp,
            y_d=y_d, const_map={k.upper(): v for k, v in const_map.items()})
    else:
        feed = None
    if feed is None:
        t_f = sp.temp_f
        have = (t_f is not None
                and all(t is not None and o is not None and p is not None
                        for t, o, p in zip(tc_r, omega, pc_l)))
        beta0 = min(max(nv / N, 1e-4), 1.0 - 1e-4)
        if have:
            t_r = t_f + F_TO_R
            k_raw = [_wilson_k(tc_r[i], pc_l[i], omega[i], t_r, p_ref)
                     for i in range(len(names))]
            if abs(rachford_rice(z, k_raw) - beta0) < 0.01:
                k_ref, basis = [min(1e10, max(1e-10, k)) for k in k_raw], "wilson"
            else:
                lam = _solve_lambda(z, k_raw, beta0)
                k_ref = [min(1e10, max(1e-10, lam * k)) for k in k_raw]
                basis = "wilson-lambda"
        else:
            # no T or no criticals: neutral volatility ordering by MW
            k_raw = [1.0 / max(mw[n], 1.0) for n in names]
            lam = _solve_lambda(z, k_raw, beta0)
            k_ref = [min(1e10, max(1e-10, lam * k)) for k in k_raw]
            basis = "zx-lambda"
        total_mw = sum(zi * mw[n] for zi, n in zip(z, names))
        feed = FlashFeed(
            names=names, z=z, k_ref=k_ref, mw=[mw[n] for n in names],
            total_mass_lbhr=vap_mass + liq_mass, p_ref_psia=p_ref,
            temp_f=sp.temp_f, total_mw=total_mw, total_moles_lbmolhr=N,
            tc_r=tc_r, omega=omega, pc_psia=pc_l,
            vap_cp=sp.vap_cp, liq_cp=sp.liq_cp, k_basis=basis)
        feed.beta_ref = rachford_rice(z, k_ref)
    est_used: list[str] = []
    if sp.temp_f is not None and feed is not None:
        try:
            fr0 = flash(feed, p_ref)
            est = _est_props(const_map, feed.names, fr0.y, fr0.x, fr0.vap_moles,
                             fr0.liq_moles, sp.temp_f, p_ref)
        except Exception:
            est = {}
        fill = {}
        for k, a in (("z_vap", "vap_z"), ("visc_vap", "vap_visc"),
                     ("cpcv_vap", "vap_cp_cv"), ("rho_liq", "liq_density"),
                     ("visc_liq", "liq_visc")):
            if getattr(sp, a, None) is None and est.get(k) is not None:
                fill[a] = est[k]
                est_used.append(a)
        if fill:
            sp = replace(sp, **fill)
    if sp.vap_density is None and vmw and sp.temp_f is not None:
        sp = replace(sp, vap_density=_gas_rho_kgm3(vmw, sp.vap_z, sp.temp_f, p_ref)
                     / LBFT3_TO_KGM3)

    return {"feed": feed, "sp": sp, "moles": n_i, "estimated": est_used}


# ════════════════════════════════════════════════════════════════════════
#  Property ESTIMATION from Tb / SG / MW / Tc / Pc / ω  (no property package)
# ════════════════════════════════════════════════════════════════════════
# Generalised correlations, used ONLY where neither a yellow input cell nor an
# HMB value exists.  Typical accuracy: Z ±2-5 %, liquid ρ ±3-5 %, gas μ ±5-10 %,
# liquid μ ±20-30 %, Cp ±5-10 %, k ±10-20 %, σ ±10 %.  Basis for enthalpy:
# ideal gas at 77 °F = 0 (liquid = vapour − latent heat).
#   Z (vapour)     Peng-Robinson, Kay-type quadratic mixing, kij = 0
#   ρ (liquid)     Rackett (Spencer-Danner Zra), ideal volume additivity
#   μ (vapour)     Lucas dilute-gas + Jossi-Stiel-Thodos density term; Herning-Zipperer mixing
#   μ (liquid)     Letsou-Stiel ; Kendall-Monroe mixing
#   Cp° (vapour)   Smith-Van Ness-Abbott polynomials, else n-paraffin trend by MW
#                  corrected with Watson K (from Tb, SG) ; Cp/Cv = Cp/(Cp−R) ideal
#   k  (vapour)    modified Eucken ;  k (liquid) Sato-Riedel
#   Cp (liquid)    Rowlinson-Bondi
#   σ              Brock-Bird (mole-fraction mixing)
#   h              ∫Cp° dT ; latent heat Riedel (Tb) + Watson scaling
_RJ = 8.314462618                       # J/mol/K
_CP_ABCD = {        # Cp°/R = A + B·T + C·T² + D/T²   (T in K)  [SVNA Table C.1]
    "METHANE": (1.702, 9.081e-3, -2.164e-6, 0), "ETHANE": (1.131, 19.225e-3, -5.561e-6, 0),
    "PROPANE": (1.213, 28.785e-3, -8.824e-6, 0), "NC4": (1.935, 36.915e-3, -11.402e-6, 0),
    "IC4": (1.677, 37.853e-3, -11.945e-6, 0), "NC5": (2.464, 45.351e-3, -14.111e-6, 0),
    "NC6": (3.025, 53.722e-3, -16.791e-6, 0), "NC7": (3.570, 62.127e-3, -20.001e-6, 0),
    "BENZENE": (-0.206, 39.064e-3, -13.301e-6, 0), "TOLUENE": (0.290, 47.052e-3, -15.716e-6, 0),
    "ETHYLBENZENE": (1.124, 55.380e-3, -18.476e-6, 0), "STYRENE": (2.050, 50.192e-3, -16.662e-6, 0),
    "ETHYLENE": (1.424, 14.394e-3, -4.392e-6, 0), "PROPYLENE": (1.637, 22.706e-3, -6.915e-6, 0),
    "1-BUTENE": (1.967, 31.630e-3, -9.873e-6, 0), "BUTADIENE": (2.734, 26.786e-3, -8.882e-6, 0),
    "ACETYLENE": (6.132, 1.952e-3, 0, -1.299e5), "CYCLOHEXANE": (-3.876, 63.249e-3, -20.928e-6, 0),
    "METHANOL": (2.211, 12.216e-3, -3.450e-6, 0), "ETHANOL": (3.518, 20.001e-3, -6.002e-6, 0),
    "H2": (3.249, 0.422e-3, 0, 0.083e5), "NITROGEN": (3.280, 0.593e-3, 0, 0.040e5),
    "O2": (3.639, 0.506e-3, 0, -0.227e5), "CO": (3.376, 0.557e-3, 0, -0.031e5),
    "CO2": (5.457, 1.045e-3, 0, -1.157e5), "H2O": (3.470, 1.450e-3, 0, 0.121e5),
    "NH3": (3.578, 3.020e-3, 0, -0.186e5), "H2S": (3.931, 1.490e-3, 0, -0.232e5),
    "CL2": (4.442, 0.089e-3, 0, -0.344e5), "SO2": (5.699, 0.801e-3, 0, -1.015e5),
    "HCL": (3.156, 0.623e-3, 0, 0.151e5), "HELIUM": (2.5, 0, 0, 0), "ARGON": (2.5, 0, 0, 0),
}


def _est_comp_data(names, cmap):
    """Per-component SI data list, or None if any component lacks constants."""
    out = []
    for n in names:
        c = (cmap or {}).get(str(n).upper()) or {}
        if (any(c.get(k) is None for k in ("tc_f", "pc_psia", "omega"))
                and c.get("mw") and c.get("nbp_f") is not None
                and c.get("sld") and 20.0 <= c["sld"] <= 125.0):
            c = {**c, **{k: v for k, v in _pseudo_constants(
                c["mw"], c["nbp_f"], c["sld"] / 62.428).items()
                if c.get(k) is None}}
        if any(c.get(k) is None for k in ("mw", "tc_f", "pc_psia", "omega")):
            return None
        nb = c.get("nbp_f")
        sld = c.get("sld")
        out.append(dict(
            name=str(n).upper(), mw=c["mw"], tc=(c["tc_f"] + 459.67) / 1.8,
            pc=c["pc_psia"] / 14.5038, w=c["omega"],
            tb=((nb + 459.67) / 1.8) if nb is not None else None,
            sg=(sld / 62.428) if (sld and 20.0 <= sld <= 125.0) else None))
    return out


def _cp_ig(ci, T):
    """Ideal-gas Cp, J/mol/K."""
    ab = _CP_ABCD.get(ci["name"])
    if ab:
        A, B, C, D = ab
        return _RJ * (A + B * T + C * T * T + D / (T * T))
    n = ci["mw"] / 14.027
    if n < 2.5:
        return _RJ * (3.5 + 0.5e-3 * T)
    cp = _RJ * ((0.5575 * n - 0.35) + (8.467 * n + 2.5) * 1e-3 * T
                - 2.8e-6 * n * T * T)
    if ci["tb"] and ci["sg"]:
        K = (ci["tb"] * 1.8) ** (1.0 / 3.0) / ci["sg"]
        cp *= min(1.0, max(0.55, 1.0 - 0.12 * (12.5 - K)))
    return cp


def _dh_ig(ci, T):
    """∫Cp° dT from 298.15 K, J/mol."""
    T0 = 298.15
    ab = _CP_ABCD.get(ci["name"])
    if ab:
        A, B, C, D = ab
        return _RJ * (A * (T - T0) + B / 2 * (T * T - T0 * T0)
                      + C / 3 * (T ** 3 - T0 ** 3) - D * (1 / T - 1 / T0))
    return 0.5 * (_cp_ig(ci, T) + _cp_ig(ci, T0)) * (T - T0)     # trapezoid


def _dhvap(ci, T):
    """Latent heat J/mol: Riedel at Tb, Watson-scaled to T."""
    if not ci["tb"] or T >= ci["tc"]:
        return 0.0
    tbr, tr = ci["tb"] / ci["tc"], T / ci["tc"]
    hvb = (1.093 * _RJ * ci["tb"] * (math.log(ci["pc"]) - 1.013)
           / max(0.930 - tbr, 0.05))
    return hvb * ((1 - tr) / (1 - tbr)) ** 0.38


def _pr_z(cd, y, T, P_bar):
    """Vapour compressibility, Peng-Robinson, kij = 0."""
    R = 83.14462618
    a, b = [], []
    for c in cd:
        k = 0.37464 + 1.54226 * c["w"] - 0.26992 * c["w"] ** 2
        al = (1 + k * (1 - math.sqrt(T / c["tc"]))) ** 2
        a.append(0.45724 * R * R * c["tc"] ** 2 / c["pc"] * al)
        b.append(0.07780 * R * c["tc"] / c["pc"])
    am = sum(y[i] * y[j] * math.sqrt(a[i] * a[j])
             for i in range(len(cd)) for j in range(len(cd)))
    bm = sum(y[i] * b[i] for i in range(len(cd)))
    A = am * P_bar / (R * T) ** 2
    B = bm * P_bar / (R * T)
    Z = 1.0
    for _ in range(100):
        f = Z ** 3 - (1 - B) * Z ** 2 + (A - 3 * B * B - 2 * B) * Z - (A * B - B * B - B ** 3)
        df = 3 * Z * Z - 2 * (1 - B) * Z + (A - 3 * B * B - 2 * B)
        if df == 0:
            break
        zn = Z - f / df
        if abs(zn - Z) < 1e-10:
            Z = zn
            break
        Z = zn
    return Z if Z > B else 1.0


def _lucas_dilute(c, T):
    """Dilute-gas viscosity, µP (nonpolar Lucas)."""
    tr = T / c["tc"]
    xi = 0.176 * (c["tc"] / (c["mw"] ** 3 * c["pc"] ** 4)) ** (1.0 / 6.0)
    return (0.807 * tr ** 0.618 - 0.357 * math.exp(-0.449 * tr)
            + 0.340 * math.exp(-4.058 * tr) + 0.018) / xi


def _est_props(cmap, names, y, x, nv, nl, T_F, P_psia, only=None):
    """Estimate stream properties for a vapour (y, nv) / liquid (x, nl) split.

    Returns FPS values: z_vap, visc_vap, cpcv_vap, cp_vap (BTU/lb-F),
    k_vap (BTU/hr-ft-F), rho_liq (lb/ft3), visc_liq (cP), cp_liq, k_liq,
    sigma (dyn/cm), h_vap / h_liq (BTU/lb).  {} if constants are incomplete.
    """
    if T_F is None or not P_psia or not names:
        return {}
    cd = _est_comp_data(names, cmap)
    if cd is None:
        return {}
    T = (T_F + 459.67) / 1.8
    P = P_psia / 14.5038
    n = len(cd)
    out: dict = {}
    want = (lambda k: only is None or k in only)
    J_G_TO_BTU_LB = 0.238846
    # ── vapour ────────────────────────────────────────────────────────────
    sy = sum(y) if y else 0.0
    if nv > 0 and sy > 0:
        yv = [v / sy for v in y]
        Mv = sum(yv[i] * cd[i]["mw"] for i in range(n))
        Z = _pr_z(cd, yv, T, P)
        out["z_vap"] = Z
        cpm = sum(yv[i] * _cp_ig(cd[i], T) for i in range(n))      # J/mol/K
        out["cp_vap"] = cpm / Mv * J_G_TO_BTU_LB
        out["cpcv_vap"] = cpm / (cpm - _RJ)
        eta_i = [_lucas_dilute(c, T) for c in cd]                   # µP
        num = sum(yv[i] * eta_i[i] * math.sqrt(cd[i]["mw"]) for i in range(n))
        den = sum(yv[i] * math.sqrt(cd[i]["mw"]) for i in range(n))
        eta0 = num / den                                            # µP
        # density term (Jossi-Stiel-Thodos) only where ρr > 0.1
        tcm = sum(yv[i] * cd[i]["tc"] for i in range(n))
        pcm = sum(yv[i] * cd[i]["pc"] for i in range(n))
        vcm = sum(yv[i] * (0.291 - 0.080 * cd[i]["w"]) * 83.14462618
                  * cd[i]["tc"] / cd[i]["pc"] for i in range(n))    # cm3/mol
        rho_r = (P / (Z * 83.14462618 * T)) * vcm
        eta = eta0
        if rho_r > 0.1:
            xi = tcm ** (1 / 6) / (Mv ** 0.5 * (pcm / 1.01325) ** (2 / 3))
            rhs = (1.023 + 0.23364 * rho_r + 0.58533 * rho_r ** 2
                   - 0.40758 * rho_r ** 3 + 0.093324 * rho_r ** 4)
            eta = eta0 + (rhs ** 4 - 1.0) / xi
        out["visc_vap"] = eta * 1e-4                                # µP → cP
        cv = cpm - _RJ
        out["k_vap"] = (eta * 1e-7 / (Mv * 1e-3)) * (1.32 * cv + 1.77 * _RJ) * 0.577789
        out["h_vap"] = (sum(yv[i] * _dh_ig(cd[i], T) for i in range(n)) / Mv
                        * 0.429923)                                 # J/g → BTU/lb
    # PR-package override for the vapour when every component is covered
    # (Z, Cp, Cp/Cv, h, µ, k calibrated to the HMB; formation enthalpy basis)
    pr_ok = _pr_ready(names)
    if pr_ok and nv > 0 and sy > 0:
        try:
            out.update(_pr_vapor_props(names, y, T_F, P_psia))
        except Exception:
            pass
    # ── liquid ────────────────────────────────────────────────────────────
    sx = sum(x) if x else 0.0
    if nl > 0 and sx > 0:
        xl = [v / sx for v in x]
        Ml = sum(xl[i] * cd[i]["mw"] for i in range(n))
        trs = [min(T / c["tc"], 0.97) for c in cd]
        vs = [83.14462618 * c["tc"] / c["pc"]
              * (0.29056 - 0.08775 * c["w"]) ** (1 + (1 - trs[i]) ** (2 / 7))
              for i, c in enumerate(cd)]                            # cm3/mol
        rho = Ml / sum(xl[i] * vs[i] for i in range(n))             # g/cm3
        out["rho_liq"] = rho * 62.428
        # Letsou-Stiel viscosity, Kendall-Monroe cube-root mixing
        cube = 0.0
        for i, c in enumerate(cd):
            tr = trs[i]
            xi = 2173.424 * c["tc"] ** (1 / 6) * c["mw"] ** -0.5 * (c["pc"] * 1e5) ** (-2 / 3)
            e0 = 0.015174 - 0.02135 * tr + 0.0075 * tr * tr
            e1 = 0.042552 - 0.07674 * tr + 0.0340 * tr * tr
            cube += xl[i] * ((e0 + c["w"] * e1) / xi) ** (1 / 3)
        out["visc_liq"] = cube ** 3
        # Rowlinson-Bondi liquid Cp
        cpl = 0.0
        for i, c in enumerate(cd):
            tr = trs[i]
            d = (1.586 + 0.49 / (1 - tr)
                 + c["w"] * (4.2775 + 6.3 * (1 - tr) ** (1 / 3) / tr + 0.4355 / (1 - tr)))
            cpl += xl[i] * (_cp_ig(c, T) + _RJ * d)
        out["cp_liq"] = cpl / Ml * J_G_TO_BTU_LB
        # Sato-Riedel thermal conductivity, weight-fraction mixing
        wl = [xl[i] * cd[i]["mw"] / Ml for i in range(n)]
        kl = 0.0
        for i, c in enumerate(cd):
            tbr = (c["tb"] / c["tc"]) if c["tb"] else 0.65
            kl += wl[i] * (1.11 / c["mw"] ** 0.5
                           * (3 + 20 * (1 - trs[i]) ** (2 / 3))
                           / (3 + 20 * (1 - tbr) ** (2 / 3)))
        out["k_liq"] = kl * 0.577789
        # Brock-Bird surface tension, mole-fraction mixing
        sg = 0.0
        for i, c in enumerate(cd):
            if not c["tb"] or T >= c["tc"]:
                continue
            tbr = c["tb"] / c["tc"]
            q = 0.1196 * (1 + tbr * math.log(c["pc"] / 1.01325) / (1 - tbr)) - 0.279
            sg += xl[i] * (c["pc"] ** (2 / 3) * c["tc"] ** (1 / 3) * q
                           * (1 - T / c["tc"]) ** (11 / 9))
        out["sigma"] = sg
        out["h_liq"] = (sum(xl[i] * (_dh_ig(cd[i], T) - _dhvap(cd[i], T))
                            for i in range(n)) / Ml * 0.429923)
        if pr_ok:                      # PR liquid root: enthalpy on the HMB basis
            try:
                pl = _PRMix(names, xl).props((T_F + 459.67) / 1.8, P_psia * 0.0689476,
                                             liquid=True)
                out["h_liq"], out["cp_liq"] = pl["h"], pl["cp"]
            except Exception:
                pass
        # Generalised correlations are poor for water: dedicated correlations
        # when the liquid is (almost) pure water.
        iw = [i for i, c in enumerate(cd) if c["name"] in ("H2O", "WATER")]
        if iw and sum(xl[i] for i in iw) >= 0.99 and T < 640.0:
            tr = 1.0 - T / 647.096
            out["rho_liq"] = 62.4 * (1.0 - 0.0002 * max(T - 288.15, 0.0)
                                     - 2.5e-6 * max(T - 288.15, 0.0) ** 2)
            out["visc_liq"] = 2.414e-2 * 10 ** (247.8 / (T - 140.0))   # cP (Vogel)
            out["cp_liq"] = 1.0 + 0.0004 * max(T - 298.15, 0.0)         # BTU/lb-F
            out["k_liq"] = 0.3466 + 0.00038 * (T - 273.15) - 1.0e-6 * (T - 273.15) ** 2 * 3.0
            out["sigma"] = 235.8 * tr ** 1.256 * (1 - 0.625 * tr)      # dyn/cm
    return out


# ════════════════════════════════════════════════════════════════════════
#  Peng-Robinson (+Peneloux) component data  — PR package of the HMB simulator
#  98 components (source: Some_Properties_rev1_17Nov26.xlsx — Therm, Ideal
#  Enthalpy, Binary Parameters, PR Peneloux).  Regenerate, do not hand-edit.
# ════════════════════════════════════════════════════════════════════════
# name: (MW, Tc K, Pc bar, ω, Vc cm3/mol, H offset kJ/kmol, Hf kJ/kmol,
#        ideal-gas H(T) polynomial a..f [kJ/kg, T in K], Peneloux c cm3/mol)
# Ideal enthalpy H_i(T) [kJ/kmol] = MW·poly(T) + H offset  (includes heat of formation).
# Peneloux c = PR Peneloux (ft3/lbmol) × 62.428 where the source workbook lists it;
# components the workbook does not list carry 0.0 (no volume shift) unless noted.
_PR_DATA = {
    'H2O': (18.015100479126, 647.2990, 221.20000, 0.34400001168251, 57.1000, -251712.490605563, -241814,
        (-5.7296, 1.9145, -0.00039574, 8.76206e-07, -4.95055e-10, 1.03846e-13), 0.164),
    'H2S': (34.0808982849121, 373.6000, 90.07790, 0.0810000002384186, 97.9999, -30140.4609401633, -20179,
        (-1.435, 0.9985, -0.0001843, 5.57076e-07, -3.17706e-10, 6.36625e-14), -3.1259),
    'NH3': (17.0300006866455, 405.5490, 112.76900, 0.25499901175499, 80.3999, -55757.6900649454, -45710,
        (-1.8514, 1.9937, -0.00053266, 2.0615e-06, -1.33856e-09, 3.05326e-13), -0.8656),
    'H2': (2.01600003242493, 33.4400, 13.15500, -0.120090000331402, 51.5000, -8287.94689311488, 0,
        (-49.68312, 13.83761, 0.0002999806, 3.458931e-07, -9.712927e-11, 7.731201e-15), -4.4166),
    'NITROGEN': (28.0130004882813, 126.1940, 33.94370, 0.0399998016655445, 90.0000, -8529.70379943045, 0,
        (2.888634, 0.9827466, 9.714241e-05, -4.157947e-10, -3.655484e-12, 4.050133e-16), -3.4576),
    'METHANE': (16.0429000854492, 190.6990, 46.40680, 0.0114984000101686, 98.9999, -84928.6245558341, -74900,
        (-12.98, 2.36459, -0.00213247, 5.6618e-06, -3.72476e-09, 8.60896e-13), -4.2086),
    'ETHANE': (30.0699005126953, 305.4280, 48.83850, 0.0986000001430511, 148.0000, -96704.1624207535, -84738,
        (-1.7675, 1.1429, -0.0003236, 4.2431e-06, -3.39316e-09, 8.82096e-13), -4.5933),
    'PROPANE': (44.0970001220703, 369.8980, 42.56660, 0.152400001883507, 200.0000, -119360.05427959, -103890,
        (39.4889, 0.395, 0.00211409, 3.96486e-07, -6.67176e-10, 1.67936e-13), -4.9033),
    'IC4': (58.1240005493164, 408.0960, 36.47620, 0.184790000319481, 263.0000, -152466.512358435, -134590,
        (30.903, 0.1533, 0.00263479, 7.27226e-08, -7.27896e-10, 2.36736e-13), -5.2769),
    'NC4': (58.1240005493164, 425.1990, 37.96620, 0.20100000500679, 254.9900, -145576.710057205, -126190,
        (67.721, 0.00854058, 0.00327699, -1.10968e-06, 1.76646e-10, -6.39926e-15), -4.5838),
    'IC5': (72.1510009765625, 460.3980, 33.33590, 0.222240000963211, 307.9900, -176693.306123832, -154590,
        (64.25, -0.131798, 0.003541, -1.3332e-06, 2.51446e-10, -1.29576e-14), -4.3789),
    'NC5': (72.1510009765625, 469.6000, 33.75120, 0.253890007734299, 310.9900, -169942.680608612, -146490,
        (63.198, -0.0117017, 0.0033164, -1.1705e-06, 1.99636e-10, -8.66485e-15), -3.4042),
    'CYCLOPENTANE': (70.1350021362305, 511.6000, 45.08950, 0.19200000166893, 259.9900, -82868.4898206678, -77288,
        (1.56197e-08, -0.764518, 0.00386825, -1.44055e-06, 2.31157e-10, -8.26099e-23), -4.8091),
    'CYCLOHEXANE': (84.1600036621094, 553.2000, 40.53000, 0.213300004601479, 307.9900, -131896.191402247, -123190,
        (4.55567e-09, -0.64805, 0.00363175, -9.99283e-07, 3.92403e-11, -2.5816e-23), -5.1503),
    'CO2': (44.0097007751465, 304.1000, 73.70000, 0.238940000534058, 93.9000, -403629.911229437, -393790,
        (1.25255e-09, 0.618139, 0.000484485, -1.49353e-07, 2.2905e-11, -1.37045e-15), -1.3758),
    'HCL': (36.4608001708984, 324.7000, 83.00000, 0.133000001311302, 80.8999, -101289.335448284, -92359,
        (0.0, 0.841536, -9.88125e-05, 1.13982e-07, -2.67445e-11, 0.0), -2.4609),
    'SO2': (64.0630035400391, 430.7990, 78.70000, 0.256000012159348, 122.1700, -306776.127886296, -297100,
        (-2.11957e-09, 0.37228, 0.000522843, -2.58127e-07, 5.18235e-11, 1.2589e-23), -1.442),
    'SO3': (80.0578002929688, 491.0000, 82.10000, 0.481000006198883, 127.2500, -406440.964168187, -396000,
        (-3.69848e-08, 0.206948, 0.000904133, -4.66872e-07, 9.09358e-11, 1.15978e-22), 1.8544),
    'CL2': (70.9057006835938, 416.8980, 77.40000, 0.0900000035762787, 123.8000, -9228.0586888887, 0,
        (0.0, 0.38005, 0.00023878, -1.82e-07, 5.45787e-11, 0.0), -4.4992),
    'O2': (32, 154.7700, 50.80020, 0.0189999006688595, 73.2000, -8659.50097304527, 0,
        (13.4497, 0.813132, 0.00016558, 6.82e-09, -2.331e-11, 3.7644e-15), -3.0085),
    'CO': (28.0109004974365, 132.9490, 34.98750, 0.0930000022053719, 89.2999, -119260.420539854, -110590,
        (-2.269, 1.0739, -0.00017265, 3.02226e-07, -1.37526e-10, 2.00356e-14), -3.4112),
    'ARGON': (39.9480018615723, 150.7070, 48.63600, -0.00400000018998981, 74.9000, -6196.0367739833, 0,
        (0.0, 0.520298, -4.01929e-07, 4.31096e-10, 0.0, 0.0), 0.3675),
    'NC6': (86.1779022216797, 507.8980, 30.31620, 0.300700008869171, 368.0000, -195010.486723487, -167290,
        (74.513, -0.096697, 0.00347649, -1.3212e-06, 2.52365e-10, -1.34666e-14), -1.4144),
    'NC7': (100.205001831055, 540.1580, 27.36780, 0.349790006875992, 425.9800, -219753.308034262, -187890,
        (71.41, -0.0968949, 0.003473, -1.3302e-06, 2.55766e-10, -1.37726e-14), 1.3938),
    'NC8': (114.232002258301, 568.5980, 24.96620, 0.401800006628036, 486.0000, -249006.099292054, -208590,
        (126.507, -0.2701, 0.00399829, -1.973e-06, 6.22796e-10, -9.38135e-14), 5.292),
    'CYCLOPROPANE': (42.0806007385254, 397.8110, 54.95080, 0.129960000514984, 170.6900, 49263.8549037915, 53338,
        (1.47133e-08, -0.83742, 0.00453056, -2.28211e-06, 5.36762e-10, -7.86756e-23), -4.6699),
    'CYCLOBUTANE': (56.1077003479004, 459.9270, 49.84890, 0.180950000882149, 210.1300, 22256.9678252389, 26669,
        (8.32829e-09, -0.895568, 0.0044771, -2.11377e-06, 4.66509e-10, -4.19138e-23), -4.2381),
    'MCYCLOHEXANE': (98.1890029907227, 572.0980, 34.75370, 0.232960000634193, 368.0000, -176812.857380919, -154890,
        (107.59, -0.704597, 0.0041033, -1.5285e-06, 1.83396e-10, 2.66336e-14), -4.9021),
    'MCYCLOPENTAN': (84.1619033813477, 532.7000, 37.89550, 0.238940000534058, 319.0000, -126781.658499949, -106790,
        (127.24, -0.6841, 0.0040055, -1.6815e-06, 3.57826e-10, -2.62676e-14), -4.4001),
    'ECYCLOPENTAN': (98.1890029907227, 569.5200, 33.97620, 0.282590001821518, 374.4600, -145597.255194235, -127190,
        (64.915, -0.677, 0.0042233, -2.1238e-06, 6.99946e-10, -1.12975e-13), -3.0119),
    'ETHYLENE': (28.0538005828857, 282.3590, 50.31790, 0.0850000008940697, 128.9400, 41693.0302911846, 52329.80078125,
        (1.12008e-09, 1.137, -0.00024462, 2.92095e-06, -2.10761e-09, 4.85356e-13), -4.3409),
    'PROPENE': (42.0806007385254, 365.0000, 46.20410, 0.148000001907349, 180.9500, 9881.42710286316, 20429,
        (1.92663e-08, 0.0881636, 0.0027863, -9.18863e-07, 1.30992e-10, -1.04978e-22), -4.7278),
    'PROPADIENE': (40.0648002624512, 393.1480, 54.71540, 0.159600004553795, 162.0000, 179541.868575046, 192290,
        (13.85, 0.537, 0.001655, -1.82495e-08, -3.32996e-10, 9.66386e-14), -4.8669),
    '1-BUTENE': (56.1077003479004, 419.6000, 40.22600, 0.187000006437302, 239.9500, -13261.9392185581, -125.970001220703,
        (8.30081e-09, -0.0533615, 0.0031475, -1.18222e-06, 1.98854e-10, -4.27421e-23), -4.5729),
    'CIS2-BUTENE': (56.1077003479004, 435.5780, 42.05770, 0.202950000762939, 233.9400, -19345.3761160482, -6989.7900390625,
        (5.41921e-09, 0.00783491, 0.00263154, -6.04784e-07, -2.74467e-12, -2.79141e-23), -4.0496),
    'TR2-BUTENE': (56.1077003479004, 428.6270, 41.02350, 0.218180000782013, 238.1500, -27399.4742719071, -11178.900390625,
        (2.68367e-08, 0.3265, 0.00228487, -4.16634e-07, -4.00519e-11, -1.48189e-22), -3.6909),
    'I-BUTENE': (56.1077003479004, 417.8980, 40.02330, 0.189980000257492, 238.9400, -33211.1038833511, -16909,
        (2.0158e-08, 0.28605, 0.00249875, -6.48153e-07, 4.05376e-11, -1.09334e-22), -4.4741),
    '12-BUTADIENE': (54.0918006896973, 443.6990, 44.98830, 0.254990011453629, 218.9700, 145238.05651725, 162290,
        (23.1848, 0.4646, 0.0017529, 4.79806e-09, -3.31946e-10, 8.91906e-14), -2.3204),
    '13-BUTADIENE': (54.0918006896973, 425.0000, 43.26560, 0.194940000772476, 221.0000, 94358.7906863635, 110190,
        (47.206, 0.059928, 0.002823, -8.74555e-07, -3.91706e-11, 6.01406e-14), -4.4107),
    'CYCLOBUTENE': (54.0900001525879, 446.2990, 52.70000, 0.188997000455856, 195.5000, 122612.431881767, 129500,
        (-4.93855e-08, -0.51185, 0.00360519, -1.62771e-06, 3.29771e-10, 1.15662e-21), 0.0),  # no Peneloux in source: 0
    '2E1C4=': (84.1600036621094, 510.3980, 32.70000, 0.266999006271362, 353.5000, -71022.3663142847, -51550,
        (-3.02911e-08, -0.157817, 0.00357645, -1.59105e-06, 3.38132e-10, 7.04671e-22), -2.518),
    '1-PENTENE': (70.1350021362305, 464.6990, 35.28700, 0.232960000634193, 300.0000, -38175.5555322329, -20929,
        (-1.44437e-08, -0.0019105, 0.00308619, -1.1012e-06, 1.66852e-10, 8.01642e-23), -4.0973),
    '3M1C5=': (84.1600036621094, 495.2990, 32.70000, 0.261999011039734, 346.5000, -68173.7365673392, -45010,
        (-4.83573e-08, 0.0389258, 0.00337785, -1.47346e-06, 3.05248e-10, 1.12611e-21), -3.069),
    'CYCLOPENTENE': (68.1184005737305, 507.0000, 48.01990, 0.195859000086784, 245.0000, 19476.8700199973, 33100,
        (0.0, 0.166132, 0.00138023, 1.47209e-06, -1.54839e-09, 4.18547e-13), -4.4676),
    'CIS2-PENTENE': (70.1350021362305, 475.9280, 36.54120, 0.240600004792213, 302.0600, -44027.3576706616, -28089,
        (0.0, 0.0054891, 0.0026993, -4.80186e-07, -2.31546e-10, 9.43486e-14), -3.4518),
    'TR2-PENTENE': (70.1350021362305, 475.3710, 36.54120, 0.237250000238419, 302.0600, -51187.9557325557, -31779,
        (0.0, 0.26349, 0.00230989, -1.73906e-07, -3.56725e-10, 1.15236e-13), -3.4518),
    '12PENTADIENE': (68.1190032958984, 503.1480, 40.73370, 0.172940000891685, 275.9500, 119855.048479211, 145690,
        (134.0, 0.03955, 0.00282779, -6.75446e-07, 0.0, 0.0), 0.0),  # no Peneloux in source: 0
    '1-CI3-C5==': (68.1190032958984, 498.7040, 40.12620, 0.18690000474453, 276.3600, 52322.32026512, 78220,
        (133.92, 0.043098, 0.002827, -6.74946e-07, 0.0, 0.0), -7.7417),
    '1-TR3-C5==': (68.1190032958984, 496.4830, 39.92050, 0.174970000982285, 276.3600, 51500.765916387, 77868,
        (123.55, 0.102289, 0.002821, -6.69335e-07, 0.0, 0.0), -4.5881),
    '23PENTADIENE': (68.120002746582, 505.7990, 41.39990, 0.203997001051903, 275.4980, 120831.696736459, 138490,
        (4.67411e-10, 0.183469, 0.00254535, -8.56916e-07, 1.22582e-10, -2.65387e-24), -4.4676),
    '14PENTADIENE': (68.1190032958984, 477.5920, 37.89340, 0.104000002145767, 276.3600, 78029.2192284248, 105490,
        (135.0, 0.1192, 0.00281429, -6.63376e-07, 0.0, 0.0), -10.0397),
    '2M1C5=': (84.1600036621094, 506.5000, 32.70000, 0.266999006271362, 353.5000, -73933.0854622737, -52260,
        (-5.55527e-08, 0.0220057, 0.00315862, -1.18318e-06, 1.94804e-10, 1.31803e-21), -4.4803),
    '1-HEXENE': (84.1619033813477, 501.9980, 31.71370, 0.284990012645721, 349.9900, -62329.0128787959, -41698,
        (1.51543e-08, -0.0207457, 0.00315402, -1.14975e-06, 1.79827e-10, -8.0402e-23), -3.6136),
    'CIS2-HEXENE': (84.1619033813477, 517.5930, 32.82580, 0.256000012159348, 350.9100, -83605.1364926019, -52378,
        (97.7618, 0.134, 0.0028255, -6.73335e-07, 0.0, 0.0), -3.3774),
    'TR2-HEXENE': (84.1619033813477, 516.4810, 32.68790, 0.241990000009537, 350.9100, -85575.0947942842, -53928,
        (92.3118, 0.169898, 0.0028215, -6.69846e-07, 0.0, 0.0), -3.3774),
    'CIS3-HEXENE': (84.1619033813477, 517.0390, 32.82580, 0.22495000064373, 350.9100, -79248.7396715975, -47648,
        (95.2489, 0.157998, 0.0028224, -6.70597e-07, 0.0, 0.0), -3.3774),
    'TR3-HEXENE': (84.1619033813477, 517.5930, 32.82580, 0.226980000734329, 350.9100, -86172.773766851, -54468,
        (92.8288, 0.1705, 0.0028213, -6.69596e-07, 0.0, 0.0), -3.3774),
    '15-HEXADIENE': (82.1455001831055, 520.9270, 34.45290, 0.159970000386238, 333.2700, 53091.1853986921, 83739,
        (113.9, 0.086335, 0.0028275, -6.75196e-07, 0.0, 0.0), 0.0),  # no Peneloux in source: 0
    '1-HEPTENE': (98.1890029907227, 537.2000, 28.37100, 0.358000010251999, 439.9900, -86388.9131301472, -62338,
        (1.94118e-08, -0.033639, 0.00320657, -1.19225e-06, 1.93678e-10, -1.07199e-22), 0.2161),
    '1-OCTENE': (112.208000183105, 566.5980, 26.24310, 0.386000007390976, 463.9900, -117349.255482361, -82977,
        (63.518, -0.053958, 0.0032715, -1.2781e-06, 2.51605e-10, -1.39996e-14), 3.4707),
    'CIS2-OCTENE': (112.208000183105, 581.4810, 27.76500, 0.341500014066696, 464.3700, -129639.936826879, -91230.0078125,
        (69.864, 0.128097, 0.0028398, -6.86156e-07, 0.0, 0.0), 0.1445),
    'TR2-OCTENE': (112.208000183105, 580.3720, 27.66120, 0.349990010261536, 464.3700, -133069.267237399, -94578,
        (66.273, 0.142898, 0.0028383, -6.84786e-07, 0.0, 0.0), -2.6053),
    'TR3-OCTENE': (112.208000183105, 578.1480, 27.46120, 0.335990011692047, 464.3700, -129863.266922178, -91230.0078125,
        (63.273, 0.157597, 0.0028365, -6.83166e-07, 0.0, 0.0), -3.2014),
    'CIS4-OCTENE': (112.208000183105, 577.5930, 27.35820, 0.330700010061264, 464.3700, -129792.376475307, -91230.0078125,
        (69.91, 0.132798, 0.0028385, -6.85156e-07, 0.0, 0.0), 0.7861),
    'TR4-OCTENE': (112.208000183105, 577.5930, 27.35820, 0.327300012111664, 464.3700, -129918.02493175, -91230.0078125,
        (63.22, 0.1595, 0.00283609, -6.82786e-07, 0.0, 0.0), 0.7861),
    '26-OCTADIENE': (110.194000244141, 617.0000, 28.67510, 0.12347000092268, 474.6100, -11736.8462463131, 25989.982421875,
        (93.8218, 0.0463809, 0.0028469, -6.92495e-07, 0.0, 0.0), 0.0),  # no Peneloux in source: 0
    '1-NONENE': (126.236000061035, 592.0390, 23.44120, 0.430000007152557, 583.1000, -145904.018331166, -103590,
        (54.148, 0.15659, 0.0028415, -6.87766e-07, 0.0, 0.0), 7.3378),
    'BENZENE': (78.1100006103516, 562.0980, 49.24390, 0.215000003576279, 259.9900, 68751.1831371061, 82977,
        (84.4658, -0.513298, 0.00324869, -1.5437e-06, 3.65006e-10, -2.48204e-14), -4.3339),
    'TOLUENE': (92.1408004760742, 591.7990, 41.00040, 0.259600013494492, 316.0000, 32018.3026345654, 50029,
        (74.162, -0.4231, 0.0031845, -1.4397e-06, 3.26596e-10, -2.12746e-14), -3.0284),
    'O-XYLENE': (106.166000366211, 630.3720, 37.32810, 0.302300006151199, 369.1200, -4319.77915409395, 19000,
        (32.515, -0.062588, 0.0025189, -6.56626e-07, -1.12345e-10, 7.49536e-14), -0.9242),
    'M-XYLENE': (106.166000366211, 617.0490, 35.41120, 0.32600000500679, 375.7300, 1100.60994710727, 17249,
        (1.66491e-08, -0.274749, 0.00296562, -1.17645e-06, 1.99636e-10, -9.16113e-23), -0.0911),
    'P-XYLENE': (106.166000366211, 616.2600, 35.10790, 0.325890004634857, 379.0600, -4457.73766655836, 17950,
        (43.03, -0.125888, 0.00248729, -5.34936e-07, -2.00494e-10, 9.55146e-14), -0.3377),
    'E-BENZENE': (106.166000366211, 617.0980, 36.07120, 0.300980001688004, 374.0000, 7506.12541100584, 29809,
        (70.552, -0.392, 0.0033088, -1.5277e-06, 3.56046e-10, -2.38736e-14), -0.0911),
    'N-PBENZENE': (120.194000244141, 638.3770, 31.99840, 0.344390004873276, 439.9900, -33787.2589473717, 7820,
        (212.449, -0.3466, 0.0028828, -7.24606e-07, 0.0, 0.0), 1.1057),
    'STYRENE': (104.152000427246, 636.0000, 38.40000, 0.297096014022827, 351.9990, 119196.247312532, 147400,
        (0.0, 0.749642, -0.000151457, 2.96449e-06, -2.44436e-09, 6.35729e-13), -1.6178),
    'M-MERCAPTAN': (48.1068000793457, 469.9970, 72.34600, 0.155000001192093, 144.9600, -32216.7621926298, -22989,
        (0.0, 0.275799, 0.0015137, -1.065e-06, 4.12186e-10, 0.0), 0.0),  # no Peneloux in source: 0
    'E-MERCAPTAN': (62.1338005065918, 498.9970, 54.91810, 0.189980000257492, 207.0000, -59080.0588536164, -46138,
        (0.0, 0.24009, 0.00189169, -1.30946e-06, 4.12186e-10, 0.0), 0.0),  # no Peneloux in source: 0
    'T-B-MERCAPTA': (90.1890029907227, 520.0000, 40.59960, 0.190696001052856, 307.0000, -132460.470303713, -109290,
        (4.98484e-09, 0.321917, 0.00189912, -2.29705e-07, -2.53572e-10, 8.61886e-14), 0.0),  # no Peneloux in source: 0
    'NPMERCAPTAN': (76.1500015258789, 536.0000, 45.19080, 0.279000014066696, 263.0000, -88235.5944880054, -67860,
        (0.0, 0.0712536, 0.00297867, -6.96382e-07, 0.0, 0.0), -3.5653),
    'NBMERCAPTAN': (90.1890029907227, 567.0000, 39.70000, 0.278400003910065, 307.0000, -108991.367079625, -88070,
        (-3.39726e-07, 0.197386, 0.00211784, -5.90501e-07, 6.45643e-11, 4.10087e-21), 0.0),  # no Peneloux in source: 0
    'ACETYLENE': (26.0380001068115, 308.3200, 61.39100, 0.187270000576973, 112.9600, 215950.043025049, 226900,
        (5.42055e-10, 1.03002, 0.00145516, -6.40982e-07, 1.35567e-10, -3.32651e-24), -2.2633),
    'M-ACETYLENE': (40.0648002624512, 402.3900, 56.27600, 0.216100007295609, 164.0000, 173802.975937188, 185500,
        (-9.82541e-09, 0.367149, 0.00232621, -9.76746e-07, 2.0117e-10, 5.6159e-23), -2.5597),
    'VNYLACETLEN': (52.076000213623, 451.9980, 48.60000, 0.118170000612736, 217.9600, 292011.804189425, 304800,
        (2.10868e-08, 0.129748, 0.00272774, -1.4498e-06, 3.58174e-10, -1.11127e-22), -6.6261),
    'DMACETYLENE': (54.0914993286133, 473.2000, 48.70000, 0.238542005419731, 221.0000, 127827.704234732, 145700,
        (0.0, 0.810409, 0.00069292, 1.34396e-06, -1.15192e-09, 2.85182e-13), -2.7782),
    'EACETYLENE': (54.0918006896973, 443.2000, 49.50000, 0.246900007128716, 222.0000, 150658.921160214, 165300,
        (1.18819e-08, 0.232009, 0.00253642, -9.52081e-07, 1.59447e-10, -6.87113e-23), -2.3204),
    'BIACETYLENE': (50.060001373291, 476.9970, 58.60000, 0.100000001490116, 183.5000, 457326.985696091, 472780,
        (8.83578e-10, 0.486754, 0.0022498, -1.50546e-06, 4.39354e-10, -2.36519e-23), 0.0),  # no Peneloux in source: 0
    '3-M-1-BUTYNE': (68.120002746582, 476.0000, 42.10000, 0.175999000668526, 271.4990, 119076.706225336, 136400,
        (-1.98448e-08, 0.0682618, 0.00299295, -1.28782e-06, 2.57091e-10, 4.69972e-22), 0.0),  # no Peneloux in source: 0
    '1-PENTYNE': (68.120002746582, 493.4970, 40.50000, 0.164000004529953, 277.9990, 125016.277406163, 144400,
        (9.60907e-09, 0.265259, 0.0025771, -9.36102e-07, 1.50394e-10, -2.05015e-22), 0.0),  # no Peneloux in source: 0
    '2-PENTYNE': (68.120002746582, 522.0000, 42.30000, 0.186000004410744, 277.5000, 111704.634323598, 128870,
        (-5.56367e-08, 0.179035, 0.00245492, -7.67906e-07, 9.28725e-11, 1.31607e-21), 0.0),  # no Peneloux in source: 0
    '1-HEXYNE': (82.1500015258789, 527.0000, 36.70000, 0.248998001217842, 333.4990, 101705.556128153, 123640,
        (5.05029e-08, 0.143651, 0.00283152, -1.096e-06, 1.92312e-10, -1.18354e-21), -3.7632),
    '1-HEPTYNE': (96.1697006225586, 559.7000, 32.00000, 0.293000012636185, 389.4990, 77612.1329423758, 103000,
        (-3.84125e-08, 0.110281, 0.00292034, -1.13431e-06, 1.99222e-10, 9.05374e-22), -3.1291),
    '1-OCTYNE': (110.195999145508, 586.9000, 29.60000, 0.337998002767563, 445.4990, 53551.2098431329, 82420,
        (-2.20938e-08, 0.087812, 0.00297794, -1.1505e-06, 1.98517e-10, 4.97007e-22), -0.275),
    'HELIUM': (4.00298023223877, 5.1900, 2.26970, -0.39003199338913, 57.3000, -6237.68185510886, 0,
        (9.1317, 5.1958, 0.0, 0.0, 0.0, 0.0), 0.0),  # no Peneloux in source: 0
    'NO': (30.0060005187988, 180.0000, 64.70000, 0.587970018386841, 57.7000, 81637.3266229526, 90427,
        (0.0, 0.97845, -1.56371e-05, 1.08345e-07, -3.49075e-11, 0.0), 0.0),  # no Peneloux in source: 0
    'NO2': (46.0060005187988, 430.9970, 100.00000, 0.833980023860931, 80.7000, 24677.3235619132, 33869,
        (0.0, 0.52668, 0.000525904, -1.50872e-07, 1.5932e-12, 0.0), 0.0),  # no Peneloux in source: 0
    'N2O': (44.0130004882813, 309.6000, 72.40000, 0.165000006556511, 97.4000, 72384.1895605416, 81598,
        (0.0, 0.49153, 0.000827688, -4.37852e-07, 1.04009e-10, 0.0), 0.0),  # no Peneloux in source: 0
    'COS': (60.0699005126953, 378.7990, 61.79810, 0.104070000350475, 142.0000, -151449.027753762, -142040,
        (0.0, 0.32, 0.000861838, -6.60416e-07, 2.83086e-10, -5.03625e-14), 0.0),  # no Peneloux in source: 0
    'CS2': (76.130500793457, 552.0000, 79.03390, 0.104070000350475, 172.0000, 105472.432885416, 115999,
        (0.0, 0.293499, 0.000731809, -6.18336e-07, 2.81036e-10, -5.19276e-14), -4.8697),
}
# alternative / simulator names → _PR_DATA key (normalised: upper-case, A-Z 0-9 = only)
_PR_ALIAS = {
    '12BD': '12-BUTADIENE', '12BUTADIENE': '12-BUTADIENE', '12PENTADIENE': '12PENTADIENE',
    '13BD': '13-BUTADIENE', '13BUTADIENE': '13-BUTADIENE', '14PD': '14PENTADIENE',
    '14PENTADIENE': '14PENTADIENE', '15HEXADIENE': '15-HEXADIENE', '1BUTENE': '1-BUTENE',
    '1BUTYNE': 'EACETYLENE', '1CI3C5==': '1-CI3-C5==', '1HEPTENE': '1-HEPTENE',
    '1HEPTYNE': '1-HEPTYNE', '1HEXENE': '1-HEXENE', '1HEXYNE': '1-HEXYNE', '1M1E': 'CIS2-OCTENE',
    '1NONENE': '1-NONENE', '1OCTENE': '1-OCTENE', '1OCTYNE': '1-OCTYNE', '1PENTENE': '1-PENTENE',
    '1PENTYNE': '1-PENTYNE', '1TR3C5==': '1-TR3-C5==', '23PD': '1-TR3-C5==',
    '23PENTADIENE': '23PENTADIENE', '26OCTADIENE': '26-OCTADIENE', '2BUTYNE': 'DMACETYLENE',
    '2E1C4=': '2E1C4=', '2M1C5=': '2M1C5=', '2M1PNTEN': '2M1C5=', '2PENTYNE': '2-PENTYNE',
    '3M1BUTYNE': '3-M-1-BUTYNE', '3M1C5=': '3M1C5=', '3M1PNTEN': '3M1C5=', '3MT2PNTE': 'TR3-HEXENE',
    '4MCPEN': '1-HEXYNE', '4MT2PNTE': '1-HEXENE', 'ACETYLENE': 'ACETYLENE', 'ACETYLN': 'ACETYLENE',
    'AMMONIA': 'NH3', 'AR': 'ARGON', 'ARGON': 'ARGON', 'BENZENE': 'BENZENE',
    'BIACETYLENE': 'BIACETYLENE', 'BUTADIENE': '13-BUTADIENE', 'BUTANE': 'NC4',
    'C2BUTENE': 'CIS2-BUTENE', 'C2H6': 'ETHANE', 'C2PENTEN': 'TR2-PENTENE', 'C3H8': 'PROPANE',
    'C3HEXENE': '2E1C4=', 'CARBONDIOXIDE': 'CO2', 'CARBONDISULFIDE': 'CS2', 'CARBONMONOXIDE': 'CO',
    'CARBONYLSULFIDE': 'COS', 'CH': 'CYCLOHEXANE', 'CH4': 'METHANE', 'CIS2BUTENE': 'CIS2-BUTENE',
    'CIS2HEXENE': 'CIS2-HEXENE', 'CIS2OCTENE': 'CIS2-OCTENE', 'CIS2PENTENE': 'CIS2-PENTENE',
    'CIS3HEXENE': 'CIS3-HEXENE', 'CIS4OCTENE': 'CIS4-OCTENE', 'CL2': 'CL2', 'CO': 'CO',
    'CO2': 'CO2', 'COS': 'COS', 'CP': 'CYCLOPENTANE', 'CS2': 'CS2', 'CT4C': 'TR3-OCTENE',
    'CTCP': 'TR2-OCTENE', 'CYBU': 'CYCLOBUTANE', 'CYCLOBUTANE': 'CYCLOBUTANE',
    'CYCLOBUTENE': 'CYCLOBUTENE', 'CYCLOHEXANE': 'CYCLOHEXANE', 'CYCLOPENTANE': 'CYCLOPENTANE',
    'CYCLOPENTENE': 'CYCLOPENTENE', 'CYCLOPROPANE': 'CYCLOPROPANE', 'CYPENTEN': '23PENTADIENE',
    'CYPR': 'CYCLOPROPANE', 'DMACETYLENE': 'DMACETYLENE', 'EACETYLENE': 'EACETYLENE',
    'EBENZENE': 'E-BENZENE', 'ECP': 'ECYCLOPENTAN', 'ECYCLOPENTAN': 'ECYCLOPENTAN',
    'EMERCAPTAN': 'E-MERCAPTAN', 'ETHANE': 'ETHANE', 'ETHANETHIOL': 'E-MERCAPTAN',
    'ETHENE': 'ETHYLENE', 'ETHYLBENZENE': 'E-BENZENE', 'ETHYLENE': 'ETHYLENE',
    'ETHYLMERCAPTAN': 'E-MERCAPTAN', 'ETHYNE': 'ACETYLENE', 'H2': 'H2', 'H2O': 'H2O', 'H2S': 'H2S',
    'HCL': 'HCL', 'HE': 'HELIUM', 'HELIUM': 'HELIUM', 'HEPTANE': 'NC7', 'HEXANE': 'NC6',
    'HYDROGEN': 'H2', 'HYDROGENSULFIDE': 'H2S', 'IBUTANE': 'IC4', 'IBUTENE': 'I-BUTENE',
    'IPENTANE': 'IC5', 'ISOBUTENE': 'I-BUTENE', 'ISOBUTYLENE': 'I-BUTENE',
    'MACETYLENE': 'M-ACETYLENE', 'MCH': 'MCYCLOHEXANE', 'MCP': 'MCYCLOPENTAN',
    'MCYCLOHEXANE': 'MCYCLOHEXANE', 'MCYCLOPENTAN': 'MCYCLOPENTAN', 'METHANE': 'METHANE',
    'METHANETHIOL': 'M-MERCAPTAN', 'METHYLBENZENE': 'TOLUENE', 'METHYLCYCLOHEXANE': 'MCYCLOHEXANE',
    'METHYLCYCLOPENTANE': 'MCYCLOPENTAN', 'METHYLMERCAPTAN': 'M-MERCAPTAN',
    'MMERCAPTAN': 'M-MERCAPTAN', 'MXYLENE': 'E-BENZENE', 'N2': 'NITROGEN', 'N2O': 'N2O',
    'NBMERCAPTAN': 'NBMERCAPTAN', 'NBUTANE': 'NC4', 'NHEPTANE': 'NC7', 'NHEXANE': 'NC6',
    'NITROGEN': 'NITROGEN', 'NITROUSOXIDE': 'N2O', 'NO': 'NO', 'NO2': 'NO2', 'NOCTANE': 'NC8',
    'NPBENZENE': 'N-PBENZENE', 'NPENTANE': 'NC5', 'NPMERCAPTAN': 'NPMERCAPTAN',
    'NPROPYLBENZENE': 'N-PBENZENE', 'O2': 'O2', 'OCTANE': 'NC8', 'OXYGEN': 'O2',
    'OXYLENE': 'O-XYLENE', 'PENTANE': 'NC5', 'PPLN': '1-CI3-C5==', 'PR1THIOL': 'NPMERCAPTAN',
    'PRBENZEN': 'N-PBENZENE', 'PROPADIENE': 'PROPADIENE', 'PROPANE': 'PROPANE',
    'PROPDIEN': 'PROPADIENE', 'PROPENE': 'PROPENE', 'PROPYLENE': 'PROPENE',
    'PROPYNE': 'M-ACETYLENE', 'PXYLENE': 'P-XYLENE', 'SO2': 'SO2', 'SO3': 'SO3',
    'STYRENE': 'STYRENE', 'T2BUTENE': 'TR2-BUTENE', 'T2OCTENE': 'TR4-OCTENE',
    'TBMERCAPTA': 'T-B-MERCAPTA', 'TOLUENE': 'TOLUENE', 'TR2BUTENE': 'TR2-BUTENE',
    'TR2HEXENE': 'TR2-HEXENE', 'TR2OCTENE': 'TR2-OCTENE', 'TR2PENTENE': 'TR2-PENTENE',
    'TR3HEXENE': 'TR3-HEXENE', 'TR3OCTENE': 'TR3-OCTENE', 'TR4OCTENE': 'TR4-OCTENE',
    'VAC': 'VNYLACETLEN', 'VINYLBENZENE': 'STYRENE', 'VNYLACETLEN': 'VNYLACETLEN', 'WATER': 'H2O',
}
_PR_EXTRA = {   # key: (original databank name, normal boiling point K, standard liquid density kg/m3)
    'H2O': ('H2O', 373.148, 997.986),
    'H2S': ('H2S', 213.498, 788.408),
    'NH3': ('Ammonia', 239.699, 616.070),
    'H2': ('Hydrogen', 20.555, 69.859),
    'NITROGEN': ('Nitrogen', 77.350, 806.374),
    'METHANE': ('Methane', 111.625, 299.394),
    'ETHANE': ('Ethane', 184.550, 355.683),
    'PROPANE': ('Propane', 231.048, 506.678),
    'IC4': ('i-Butane', 261.420, 561.966),
    'NC4': ('n-Butane', 272.648, 583.223),
    'IC5': ('i-Pentane', 301.028, 623.442),
    'NC5': ('n-Pentane', 309.209, 629.729),
    'CYCLOPENTANE': ('Cyclopentane', 322.398, 748.874),
    'CYCLOHEXANE': ('Cyclohexane', 353.880, 781.822),
    'CO2': ('CO2', 194.598, 825.335),
    'HCL': ('HCl', 188.100, 870.330),
    'SO2': ('SO2', 263.200, 1395.560),
    'SO3': ('SO3', 318.000, 1880.180),
    'CL2': ('Cl2', 239.199, 1561.319),
    'O2': ('Oxygen', 90.200, 1137.680),
    'CO': ('CO', 81.700, 799.388),
    'ARGON': ('Argon', 87.299, 1370.180),
    'NC6': ('n-Hexane', 341.880, 662.664),
    'NC7': ('n-Heptane', 371.579, 686.815),
    'NC8': ('n-Octane', 398.820, 705.377),
    'CYCLOPROPANE': ('Cyclopropane', 240.371, 420.540),
    'CYCLOBUTANE': ('Cyclobutane', 285.661, 698.393),
    'MCYCLOHEXANE': ('Mcyclohexane', 374.079, 772.427),
    'MCYCLOPENTAN': ('Mcyclopentan', 344.959, 752.065),
    'ECYCLOPENTAN': ('Ecyclopentan', 376.617, 769.565),
    'ETHYLENE': ('Ethylene', 169.399, 383.226),
    'PROPENE': ('Propene', 225.399, 520.955),
    'PROPADIENE': ('Propadiene', 238.699, 655.593),
    '1-BUTENE': ('1-Butene', 266.898, 593.789),
    'CIS2-BUTENE': ('cis2-Butene', 276.867, 625.953),
    'TR2-BUTENE': ('tr2-Butene', 274.027, 608.684),
    'I-BUTENE': ('i-Butene', 266.299, 592.793),
    '12-BUTADIENE': ('12-Butadiene', 283.000, 650.700),
    '13-BUTADIENE': ('13-Butadiene', 268.700, 619.744),
    'CYCLOBUTENE': ('Cyclobutene', 275.799, 710.348),
    '2E1C4=': ('2E1C4=', 337.799, 694.064),
    '1-PENTENE': ('1-Pentene', 303.096, 638.723),
    '3M1C5=': ('3M1C5=', 327.299, 672.351),
    'CYCLOPENTENE': ('Cyclopentene', 317.380, 776.516),
    'CIS2-PENTENE': ('cis2-Pentene', 310.083, 659.575),
    'TR2-PENTENE': ('tr2-Pentene', 309.493, 652.263),
    '12PENTADIENE': ('12Pentadiene', 317.993, 696.327),
    '1-CI3-C5==': ('1-ci3-C5==', 317.204, 694.955),
    '1-TR3-C5==': ('1-tr3-C5==', 315.170, 679.900),
    '23PENTADIENE': ('23Pentadiene', 321.398, 699.802),
    '14PENTADIENE': ('14Pentadiene', 299.109, 664.541),
    '2M1C5=': ('2M1C5=', 335.299, 684.583),
    '1-HEXENE': ('1-Hexene', 336.600, 671.653),
    'CIS2-HEXENE': ('cis2-hexene', 342.032, 690.539),
    'TR2-HEXENE': ('tr2-Hexene', 341.027, 681.210),
    'CIS3-HEXENE': ('cis3-Hexene', 339.588, 683.189),
    'TR3-HEXENE': ('tr3-Hexene', 340.226, 680.554),
    '15-HEXADIENE': ('15-Hexadiene', 332.600, 695.640),
    '1-HEPTENE': ('1-Heptene', 366.799, 695.606),
    '1-OCTENE': ('1-Octene', 394.398, 713.567),
    'CIS2-OCTENE': ('cis2-Octene', 398.799, 727.164),
    'TR2-OCTENE': ('tr2-Octene', 398.148, 723.067),
    'TR3-OCTENE': ('tr3-Octene', 396.427, 718.286),
    'CIS4-OCTENE': ('cis4-Octene', 395.700, 724.179),
    'TR4-OCTENE': ('tr4-Octene', 395.410, 717.192),
    '26-OCTADIENE': ('26-Octadiene', 397.648, 746.775),
    '1-NONENE': ('1-Nonene', 420.039, 732.064),
    'BENZENE': ('Benzene', 353.239, 882.190),
    'TOLUENE': ('Toluene', 383.799, 870.044),
    'O-XYLENE': ('o-Xylene', 417.575, 883.145),
    'M-XYLENE': ('m-Xylene', 412.265, 866.882),
    'P-XYLENE': ('p-Xylene', 411.510, 864.229),
    'E-BENZENE': ('E-Benzene', 409.350, 870.033),
    'N-PBENZENE': ('n-PBenzene', 432.389, 864.757),
    'STYRENE': ('Styrene', 418.308, 908.830),
    'M-MERCAPTAN': ('M-Mercaptan', 279.096, 863.192),
    'E-MERCAPTAN': ('E-Mercaptan', 308.200, 836.507),
    'T-B-MERCAPTA': ('t-B-Mercapta', 337.368, 804.820),
    'NPMERCAPTAN': ('nPMercaptan', 339.200, 845.931),
    'NBMERCAPTAN': ('nBMercaptan', 371.609, 845.913),
    'ACETYLENE': ('Acetylene', 189.350, 417.053),
    'M-ACETYLENE': ('M-Acetylene', 249.940, 620.536),
    'VNYLACETLEN': ('VnylAcetlen', 278.250, 688.432),
    'DMACETYLENE': ('DMAcetylene', 300.130, 695.843),
    'EACETYLENE': ('EAcetylene', 281.220, 659.288),
    'BIACETYLENE': ('Biacetylene', 283.500, 717.260),
    '3-M-1-BUTYNE': ('3-M-1-Butyne', 299.500, 598.737),
    '1-PENTYNE': ('1-Pentyne', 313.299, 694.890),
    '2-PENTYNE': ('2-Pentyne', 329.200, 715.265),
    '1-HEXYNE': ('1-Hexyne', 344.500, 719.960),
    '1-HEPTYNE': ('1-Heptyne', 372.898, 736.846),
    '1-OCTYNE': ('1-Octyne', 399.398, 749.840),
    'HELIUM': ('Helium', 4.210, 124.064),
    'NO': ('NO', 121.400, 71.323),
    'NO2': ('NO2', 294.299, 1469.760),
    'N2O': ('N2O', 184.699, 816.880),
    'COS': ('COS', 223.000, 1015.400),
    'CS2': ('CS2', 319.359, 1269.310),
}
_PR_KIJ = {   # binary interaction parameters (symmetric, sorted pair; zeros omitted)
    ('1-BUTENE', '1-CI3-C5=='): 0.000277161598205566, ('1-BUTENE', '1-HEPTENE'): 0.00508427619934082,
    ('1-BUTENE', '1-HEPTYNE'): 0.00325047969818115, ('1-BUTENE', '1-HEXENE'): 0.0019756555557251,
    ('1-BUTENE', '1-HEXYNE'): 0.0015033483505249, ('1-BUTENE', '1-NONENE'): 0.0108512043952942,
    ('1-BUTENE', '1-OCTENE'): 0.00600934028625488, ('1-BUTENE', '1-OCTYNE'): 0.00529414415359497,
    ('1-BUTENE', '1-PENTENE'): 0.000692427158355713, ('1-BUTENE', '1-PENTYNE'): 0.000300824642181396,
    ('1-BUTENE', '1-TR3-C5=='): 0.000277161598205566, ('1-BUTENE', '12-BUTADIENE'): 0.000116288661956787,
    ('1-BUTENE', '12PENTADIENE'): 0.000271320343017578, ('1-BUTENE', '13-BUTADIENE'): 0.00219990010373294,
    ('1-BUTENE', '14PENTADIENE'): 0.000277161598205566, ('1-BUTENE', '15-HEXADIENE'): 0.00149720907211304,
    ('1-BUTENE', '2-PENTYNE'): 0.000293493270874023, ('1-BUTENE', '23PENTADIENE'): 0.000265002250671387,
    ('1-BUTENE', '26-OCTADIENE'): 0.00642663240432739, ('1-BUTENE', '2E1C4='): 0.00208139419555664,
    ('1-BUTENE', '2M1C5='): 0.00208139419555664, ('1-BUTENE', '3-M-1-BUTYNE'): 0.000211954116821289,
    ('1-BUTENE', '3M1C5='): 0.00187236070632935, ('1-BUTENE', 'ACETYLENE'): 0.00783199071884155,
    ('1-BUTENE', 'BENZENE'): 8.93473625183105e-05, ('1-BUTENE', 'BIACETYLENE'): 0.000998258590698242,
    ('1-BUTENE', 'CIS2-BUTENE'): 8.94069671630859e-06, ('1-BUTENE', 'CIS2-HEXENE'): 0.0020032525062561,
    ('1-BUTENE', 'CIS2-OCTENE'): 0.00602418184280396, ('1-BUTENE', 'CIS2-PENTENE'): 0.00073552131652832,
    ('1-BUTENE', 'CIS3-HEXENE'): 0.0020032525062561, ('1-BUTENE', 'CIS4-OCTENE'): 0.00602418184280396,
    ('1-BUTENE', 'CO2'): 0.0568980015814304, ('1-BUTENE', 'CS2'): 0.0335654430091381,
    ('1-BUTENE', 'CYCLOBUTANE'): 0.000244498252868652, ('1-BUTENE', 'CYCLOBUTENE'): 0.000582635402679443,
    ('1-BUTENE', 'CYCLOHEXANE'): 0.000864982604980469, ('1-BUTENE', 'CYCLOPENTANE'): 8.93473625183105e-05,
    ('1-BUTENE', 'CYCLOPENTENE'): 6.02006912231445e-06, ('1-BUTENE', 'CYCLOPROPANE'): 0.00160884857177734,
    ('1-BUTENE', 'DMACETYLENE'): 9.3996524810791e-05, ('1-BUTENE', 'E-BENZENE'): 0.00272959470748901,
    ('1-BUTENE', 'E-MERCAPTAN'): 0.0596264488995075, ('1-BUTENE', 'EACETYLENE'): 8.392333984375e-05,
    ('1-BUTENE', 'ECYCLOPENTAN'): 0.00274473428726196, ('1-BUTENE', 'ETHANE'): 0.00323432683944702,
    ('1-BUTENE', 'ETHYLENE'): 0.00533372163772583, ('1-BUTENE', 'H2'): 0.292100012302399,
    ('1-BUTENE', 'H2O'): 0.5, ('1-BUTENE', 'H2S'): 0.0450000017881393,
    ('1-BUTENE', 'I-BUTENE'): 2.98023223876953e-07, ('1-BUTENE', 'IC4'): 0.000116825103759766,
    ('1-BUTENE', 'IC5'): 0.000864982604980469, ('1-BUTENE', 'M-ACETYLENE'): 0.0399990007281303,
    ('1-BUTENE', 'M-MERCAPTAN'): 0.0596264488995075, ('1-BUTENE', 'M-XYLENE'): 0.00278657674789429,
    ('1-BUTENE', 'MCYCLOHEXANE'): 0.00253474712371826, ('1-BUTENE', 'MCYCLOPENTAN'): 0.00112515687942505,
    ('1-BUTENE', 'METHANE'): 0.0107879042625427, ('1-BUTENE', 'N-PBENZENE'): 0.00508427619934082,
    ('1-BUTENE', 'NC4'): 0.00100000004749745, ('1-BUTENE', 'NC5'): 0.000933349132537842,
    ('1-BUTENE', 'NC6'): 0.00253474712371826, ('1-BUTENE', 'NC7'): 0.0045580267906189,
    ('1-BUTENE', 'NC8'): 0.00687873363494873, ('1-BUTENE', 'NITROGEN'): 0.100000001490116,
    ('1-BUTENE', 'O-XYLENE'): 0.0025707483291626, ('1-BUTENE', 'P-XYLENE'): 0.00289702415466309,
    ('1-BUTENE', 'PROPADIENE'): 0.0149990003556013, ('1-BUTENE', 'PROPANE'): 0.000460445880889893,
    ('1-BUTENE', 'PROPENE'): -0.00100000004749745, ('1-BUTENE', 'STYRENE'): 0.00203597545623779,
    ('1-BUTENE', 'TOLUENE'): 0.00105184316635132, ('1-BUTENE', 'TR2-BUTENE'): 7.74860382080078e-07,
    ('1-BUTENE', 'TR2-HEXENE'): 0.0020032525062561, ('1-BUTENE', 'TR2-OCTENE'): 0.00602418184280396,
    ('1-BUTENE', 'TR2-PENTENE'): 0.00073552131652832, ('1-BUTENE', 'TR3-HEXENE'): 0.0020032525062561,
    ('1-BUTENE', 'TR3-OCTENE'): 0.00602418184280396, ('1-BUTENE', 'TR4-OCTENE'): 0.00602418184280396,
    ('1-BUTENE', 'VNYLACETLEN'): 0.000128328800201416,
    ('1-CI3-C5==', '1-HEPTENE'): 0.00299620628356934, ('1-CI3-C5==', '1-HEPTYNE'): 0.00163328647613525,
    ('1-CI3-C5==', '1-HEXENE'): 0.000774383544921875, ('1-CI3-C5==', '1-HEXYNE'): 0.000490367412567139,
    ('1-CI3-C5==', '1-NONENE'): 0.00769335031509399, ('1-CI3-C5==', '1-OCTENE'): 0.00371742248535156,
    ('1-CI3-C5==', '1-OCTYNE'): 0.0031583309173584, ('1-CI3-C5==', '1-PENTENE'): 9.35196876525879e-05,
    ('1-CI3-C5==', '1-PENTYNE'): 4.76837158203125e-07, ('1-CI3-C5==', '12-BUTADIENE'): 0.000751972198486328,
    ('1-CI3-C5==', '13-BUTADIENE'): 0.000693619251251221,
    ('1-CI3-C5==', '15-HEXADIENE'): 0.000486671924591064, ('1-CI3-C5==', '2-PENTYNE'): 1.78813934326172e-07,
    ('1-CI3-C5==', '23PENTADIENE'): 1.78813934326172e-07, ('1-CI3-C5==', '26-OCTADIENE'): 0.00404816865921021,
    ('1-CI3-C5==', '2E1C4='): 0.000841140747070313, ('1-CI3-C5==', '2M1C5='): 0.000841140747070313,
    ('1-CI3-C5==', '3-M-1-BUTYNE'): 4.29153442382813e-06, ('1-CI3-C5==', '3M1C5='): 0.000710129737854004,
    ('1-CI3-C5==', 'ACETYLENE'): 0.0110151171684265, ('1-CI3-C5==', 'BENZENE'): 5.17368316650391e-05,
    ('1-CI3-C5==', 'BIACETYLENE'): 0.00232440233230591, ('1-CI3-C5==', 'CIS2-BUTENE'): 0.000385582447052002,
    ('1-CI3-C5==', 'CIS2-HEXENE'): 0.000791668891906738, ('1-CI3-C5==', 'CIS2-OCTENE'): 0.00372916460037231,
    ('1-CI3-C5==', 'CIS2-PENTENE'): 0.00010979175567627, ('1-CI3-C5==', 'CIS3-HEXENE'): 0.000791668891906738,
    ('1-CI3-C5==', 'CIS4-OCTENE'): 0.00372916460037231, ('1-CI3-C5==', 'CO2'): 0.101000003516674,
    ('1-CI3-C5==', 'CS2'): 0.0396388955414295, ('1-CI3-C5==', 'CYCLOBUTANE'): 0.0010417103767395,
    ('1-CI3-C5==', 'CYCLOBUTENE'): 0.00166171789169312, ('1-CI3-C5==', 'CYCLOHEXANE'): 0.000163018703460693,
    ('1-CI3-C5==', 'CYCLOPENTANE'): 5.17368316650391e-05, ('1-CI3-C5==', 'CYCLOPENTENE'): 0.00020146369934082,
    ('1-CI3-C5==', 'CYCLOPROPANE'): 0.00321614742279053, ('1-CI3-C5==', 'DMACETYLENE'): 0.000693619251251221,
    ('1-CI3-C5==', 'E-BENZENE'): 0.00126999616622925, ('1-CI3-C5==', 'E-MERCAPTAN'): 0.0560881048440933,
    ('1-CI3-C5==', 'EACETYLENE'): 0.000665903091430664, ('1-CI3-C5==', 'ECYCLOPENTAN'): 0.00128030776977539,
    ('1-CI3-C5==', 'ETHANE'): 0.00539225339889526, ('1-CI3-C5==', 'ETHYLENE'): 0.0080181360244751,
    ('1-CI3-C5==', 'H2'): 0.292100012302399, ('1-CI3-C5==', 'H2O'): 0.5,
    ('1-CI3-C5==', 'H2S'): 0.0450000017881393, ('1-CI3-C5==', 'I-BUTENE'): 0.000293910503387451,
    ('1-CI3-C5==', 'IC4'): 3.4034252166748e-05, ('1-CI3-C5==', 'IC5'): 0.000163018703460693,
    ('1-CI3-C5==', 'M-ACETYLENE'): 0.00377029180526733, ('1-CI3-C5==', 'M-MERCAPTAN'): 0.0560881048440933,
    ('1-CI3-C5==', 'M-XYLENE'): 0.00130897760391235, ('1-CI3-C5==', 'MCYCLOHEXANE'): 0.00113797187805176,
    ('1-CI3-C5==', 'MCYCLOPENTAN'): 0.000285923480987549, ('1-CI3-C5==', 'METHANE'): 0.0144607424736023,
    ('1-CI3-C5==', 'N-PBENZENE'): 0.00299620628356934, ('1-CI3-C5==', 'NC4'): 8.99434089660645e-05,
    ('1-CI3-C5==', 'NC5'): 0.000193595886230469, ('1-CI3-C5==', 'NC6'): 0.00113797187805176,
    ('1-CI3-C5==', 'NC7'): 0.00259464979171753, ('1-CI3-C5==', 'NC8'): 0.00440961122512817,
    ('1-CI3-C5==', 'NITROGEN'): 0.100000001490116, ('1-CI3-C5==', 'O-XYLENE'): 0.00116223096847534,
    ('1-CI3-C5==', 'P-XYLENE'): 0.00138527154922485, ('1-CI3-C5==', 'PROPADIENE'): 0.00394898653030396,
    ('1-CI3-C5==', 'PROPANE'): 0.00145077705383301, ('1-CI3-C5==', 'PROPENE'): 0.00248563289642334,
    ('1-CI3-C5==', 'STYRENE'): 0.000812351703643799, ('1-CI3-C5==', 'TOLUENE'): 0.000249505043029785,
    ('1-CI3-C5==', 'TR2-BUTENE'): 0.00030750036239624, ('1-CI3-C5==', 'TR2-HEXENE'): 0.000791668891906738,
    ('1-CI3-C5==', 'TR2-OCTENE'): 0.00372916460037231, ('1-CI3-C5==', 'TR2-PENTENE'): 0.00010979175567627,
    ('1-CI3-C5==', 'TR3-HEXENE'): 0.000791668891906738, ('1-CI3-C5==', 'TR3-OCTENE'): 0.00372916460037231,
    ('1-CI3-C5==', 'TR4-OCTENE'): 0.00372916460037231, ('1-CI3-C5==', 'VNYLACETLEN'): 0.000782191753387451,
    ('1-HEPTENE', '1-HEPTYNE'): 0.000206291675567627, ('1-HEPTENE', '1-HEXENE'): 0.000726997852325439,
    ('1-HEPTENE', '1-HEXYNE'): 0.00106561183929443, ('1-HEPTENE', '1-NONENE'): 0.00110042095184326,
    ('1-HEPTENE', '1-OCTENE'): 3.91602516174316e-05, ('1-HEPTENE', '1-OCTYNE'): 2.08616256713867e-06,
    ('1-HEPTENE', '1-PENTENE'): 0.00203359127044678, ('1-HEPTENE', '1-PENTYNE'): 0.00292074680328369,
    ('1-HEPTENE', '1-TR3-C5=='): 0.00299620628356934, ('1-HEPTENE', '12-BUTADIENE'): 0.00672531127929688,
    ('1-HEPTENE', '12PENTADIENE'): 0.00301539897918701, ('1-HEPTENE', '13-BUTADIENE'): 0.00654953718185425,
    ('1-HEPTENE', '14PENTADIENE'): 0.00299620628356934, ('1-HEPTENE', '15-HEXADIENE'): 0.00107091665267944,
    ('1-HEPTENE', '2-PENTYNE'): 0.00294357538223267, ('1-HEPTENE', '23PENTADIENE'): 0.0030364990234375,
    ('1-HEPTENE', '26-OCTADIENE'): 7.96914100646973e-05, ('1-HEPTENE', '2E1C4='): 0.000664889812469482,
    ('1-HEPTENE', '2M1C5='): 0.000664889812469482, ('1-HEPTENE', '3-M-1-BUTYNE'): 0.00322866439819336,
    ('1-HEPTENE', '3M1C5='): 0.000791966915130615, ('1-HEPTENE', 'ACETYLENE'): 0.0251400470733643,
    ('1-HEPTENE', 'BENZENE'): 0.00383204221725464, ('1-HEPTENE', 'BIACETYLENE'): 0.0105292201042175,
    ('1-HEPTENE', 'CIS2-BUTENE'): 0.00551652908325195, ('1-HEPTENE', 'CIS2-HEXENE'): 0.00071030855178833,
    ('1-HEPTENE', 'CIS2-OCTENE'): 4.04119491577148e-05, ('1-HEPTENE', 'CIS2-PENTENE'): 0.00196164846420288,
    ('1-HEPTENE', 'CIS3-HEXENE'): 0.00071030855178833, ('1-HEPTENE', 'CIS4-OCTENE'): 4.04119491577148e-05,
    ('1-HEPTENE', 'CO2'): 0.101000003516674, ('1-HEPTENE', 'CS2'): 0.0435004197061062,
    ('1-HEPTENE', 'CYCLOBUTANE'): 0.00753790140151978, ('1-HEPTENE', 'CYCLOBUTENE'): 0.00907027721405029,
    ('1-HEPTENE', 'CYCLOHEXANE'): 0.00176441669464111, ('1-HEPTENE', 'CYCLOPENTANE'): 0.00383204221725464,
    ('1-HEPTENE', 'CYCLOPENTENE'): 0.0047423243522644, ('1-HEPTENE', 'CYCLOPROPANE'): 0.0123251676559448,
    ('1-HEPTENE', 'DMACETYLENE'): 0.00654953718185425, ('1-HEPTENE', 'E-BENZENE'): 0.000366628170013428,
    ('1-HEPTENE', 'E-MERCAPTAN'): 0.0557244718074799, ('1-HEPTENE', 'EACETYLENE'): 0.00646436214447021,
    ('1-HEPTENE', 'ECYCLOPENTAN'): 0.000361084938049316, ('1-HEPTENE', 'ETHANE'): 0.0162639021873474,
    ('1-HEPTENE', 'ETHYLENE'): 0.0205651521682739, ('1-HEPTENE', 'H2'): 0.292100012302399,
    ('1-HEPTENE', 'H2O'): 0.5, ('1-HEPTENE', 'H2S'): 0.0450000017881393,
    ('1-HEPTENE', 'I-BUTENE'): 0.00515490770339966, ('1-HEPTENE', 'IC4'): 0.00366675853729248,
    ('1-HEPTENE', 'IC5'): 0.00176441669464111, ('1-HEPTENE', 'M-ACETYLENE'): 0.0133762359619141,
    ('1-HEPTENE', 'M-MERCAPTAN'): 0.0557244718074799, ('1-HEPTENE', 'M-XYLENE'): 0.000346064567565918,
    ('1-HEPTENE', 'MCYCLOHEXANE'): 0.000443220138549805, ('1-HEPTENE', 'MCYCLOPENTAN'): 0.00143438577651978,
    ('1-HEPTENE', 'METHANE'): 0.0301260948181152, ('1-HEPTENE', 'NC4'): 0.00411909818649292,
    ('1-HEPTENE', 'NC5'): 0.00166988372802734, ('1-HEPTENE', 'NC6'): 0.000443220138549805,
    ('1-HEPTENE', 'NC7'): 1.44839286804199e-05, ('1-HEPTENE', 'NC8'): 0.000137388706207275,
    ('1-HEPTENE', 'NITROGEN'): 0.100000001490116, ('1-HEPTENE', 'O-XYLENE'): 0.000428199768066406,
    ('1-HEPTENE', 'P-XYLENE'): 0.000308454036712646, ('1-HEPTENE', 'PROPADIENE'): 0.0137070417404175,
    ('1-HEPTENE', 'PROPANE'): 0.00857210159301758, ('1-HEPTENE', 'PROPENE'): 0.0108658671379089,
    ('1-HEPTENE', 'STYRENE'): 0.000691056251525879, ('1-HEPTENE', 'TOLUENE'): 0.00151985883712769,
    ('1-HEPTENE', 'TR2-BUTENE'): 0.0052107572555542, ('1-HEPTENE', 'TR2-HEXENE'): 0.00071030855178833,
    ('1-HEPTENE', 'TR2-OCTENE'): 4.04119491577148e-05, ('1-HEPTENE', 'TR2-PENTENE'): 0.00196164846420288,
    ('1-HEPTENE', 'TR3-HEXENE'): 0.00071030855178833, ('1-HEPTENE', 'TR3-OCTENE'): 4.04119491577148e-05,
    ('1-HEPTENE', 'TR4-OCTENE'): 4.04119491577148e-05, ('1-HEPTENE', 'VNYLACETLEN'): 0.00681412220001221,
    ('1-HEPTYNE', '1-HEXENE'): 0.000158846378326416, ('1-HEPTYNE', '1-HEXYNE'): 0.000334501266479492,
    ('1-HEPTYNE', '1-NONENE'): 0.00225698947906494, ('1-HEPTYNE', '1-OCTENE'): 0.000425279140472412,
    ('1-HEPTYNE', '1-OCTYNE'): 0.000250518321990967, ('1-HEPTYNE', '1-PENTENE'): 0.000946044921875,
    ('1-HEPTYNE', '1-PENTYNE'): 0.00157755613327026, ('1-HEPTYNE', '1-TR3-C5=='): 0.00163328647613525,
    ('1-HEPTYNE', '12-BUTADIENE'): 0.00458931922912598, ('1-HEPTYNE', '12PENTADIENE'): 0.00164735317230225,
    ('1-HEPTYNE', '13-BUTADIENE'): 0.00444388389587402, ('1-HEPTYNE', '14PENTADIENE'): 0.00163328647613525,
    ('1-HEPTYNE', '15-HEXADIENE'): 0.000337541103363037, ('1-HEPTYNE', '2-PENTYNE'): 0.0015944242477417,
    ('1-HEPTYNE', '23PENTADIENE'): 0.00166314840316772, ('1-HEPTYNE', '26-OCTADIENE'): 0.000542223453521729,
    ('1-HEPTYNE', '2E1C4='): 0.000130593776702881, ('1-HEPTYNE', '2M1C5='): 0.000130593776702881,
    ('1-HEPTYNE', '3-M-1-BUTYNE'): 0.00180637836456299, ('1-HEPTYNE', '3M1C5='): 0.000190019607543945,
    ('1-HEPTYNE', 'ACETYLENE'): 0.02090984582901, ('1-HEPTYNE', 'BENZENE'): 0.00226503610610962,
    ('1-HEPTYNE', 'BIACETYLENE'): 0.00781643390655518, ('1-HEPTYNE', 'CIS2-BUTENE'): 0.00359880924224854,
    ('1-HEPTYNE', 'CIS2-HEXENE'): 0.000151216983795166, ('1-HEPTYNE', 'CIS2-OCTENE'): 0.000429153442382813,
    ('1-HEPTYNE', 'CIS2-PENTENE'): 0.000897109508514404, ('1-HEPTYNE', 'CIS3-HEXENE'): 0.000151216983795166,
    ('1-HEPTYNE', 'CIS4-OCTENE'): 0.000429153442382813, ('1-HEPTYNE', 'CO2'): 0.101000003516674,
    ('1-HEPTYNE', 'CS2'): 0.0446295291185379, ('1-HEPTYNE', 'CYCLOBUTANE'): 0.00526642799377441,
    ('1-HEPTYNE', 'CYCLOBUTENE'): 0.00656306743621826, ('1-HEPTYNE', 'CYCLOHEXANE'): 0.000765204429626465,
    ('1-HEPTYNE', 'CYCLOPENTANE'): 0.00226503610610962, ('1-HEPTYNE', 'CYCLOPENTENE'): 0.00297766923904419,
    ('1-HEPTYNE', 'CYCLOPROPANE'): 0.00937956571578979, ('1-HEPTYNE', 'DMACETYLENE'): 0.00444388389587402,
    ('1-HEPTYNE', 'E-BENZENE'): 2.28285789489746e-05, ('1-HEPTYNE', 'E-MERCAPTAN'): 0.0518617741763592,
    ('1-HEPTYNE', 'EACETYLENE'): 0.00437355041503906, ('1-HEPTYNE', 'ECYCLOPENTAN'): 2.1517276763916e-05,
    ('1-HEPTYNE', 'ETHANE'): 0.0128653049468994, ('1-HEPTYNE', 'ETHYLENE'): 0.0167376399040222,
    ('1-HEPTYNE', 'H2'): 0.292100012302399, ('1-HEPTYNE', 'H2O'): 0.5,
    ('1-HEPTYNE', 'H2S'): 0.0450000017881393, ('1-HEPTYNE', 'I-BUTENE'): 0.00330722332000732,
    ('1-HEPTYNE', 'IC4'): 0.00213801860809326, ('1-HEPTYNE', 'IC5'): 0.000765204429626465,
    ('1-HEPTYNE', 'M-ACETYLENE'): 0.0103026032447815, ('1-HEPTYNE', 'M-MERCAPTAN'): 0.0518617741763592,
    ('1-HEPTYNE', 'M-XYLENE'): 1.8000602722168e-05, ('1-HEPTYNE', 'MCYCLOHEXANE'): 4.48226928710938e-05,
    ('1-HEPTYNE', 'MCYCLOPENTAN'): 0.000553488731384277, ('1-HEPTYNE', 'METHANE'): 0.0255040526390076,
    ('1-HEPTYNE', 'N-PBENZENE'): 0.000206291675567627, ('1-HEPTYNE', 'NC4'): 0.0024874210357666,
    ('1-HEPTYNE', 'NC5'): 0.000703334808349609, ('1-HEPTYNE', 'NC6'): 4.48226928710938e-05,
    ('1-HEPTYNE', 'NC7'): 0.000111281871795654, ('1-HEPTYNE', 'NC8'): 0.000680088996887207,
    ('1-HEPTYNE', 'NITROGEN'): 0.100000001490116, ('1-HEPTYNE', 'O-XYLENE'): 4.01139259338379e-05,
    ('1-HEPTYNE', 'P-XYLENE'): 1.02519989013672e-05, ('1-HEPTYNE', 'PROPADIENE'): 0.0105944871902466,
    ('1-HEPTYNE', 'PROPANE'): 0.00613903999328613, ('1-HEPTYNE', 'PROPENE'): 0.00810784101486206,
    ('1-HEPTYNE', 'STYRENE'): 0.000142335891723633, ('1-HEPTYNE', 'TOLUENE'): 0.000607073307037354,
    ('1-HEPTYNE', 'TR2-BUTENE'): 0.00335204601287842, ('1-HEPTYNE', 'TR2-HEXENE'): 0.000151216983795166,
    ('1-HEPTYNE', 'TR2-OCTENE'): 0.000429153442382813, ('1-HEPTYNE', 'TR2-PENTENE'): 0.000897109508514404,
    ('1-HEPTYNE', 'TR3-HEXENE'): 0.000151216983795166, ('1-HEPTYNE', 'TR3-OCTENE'): 0.000429153442382813,
    ('1-HEPTYNE', 'TR4-OCTENE'): 0.000429153442382813, ('1-HEPTYNE', 'VNYLACETLEN'): 0.00466293096542358,
    ('1-HEXENE', '1-HEXYNE'): 3.24249267578125e-05, ('1-HEXENE', '1-NONENE'): 0.00360804796218872,
    ('1-HEXENE', '1-OCTENE'): 0.00110316276550293, ('1-HEXENE', '1-OCTYNE'): 0.000808119773864746,
    ('1-HEXENE', '1-PENTENE'): 0.000329852104187012, ('1-HEXENE', '1-PENTYNE'): 0.000736117362976074,
    ('1-HEXENE', '1-TR3-C5=='): 0.000774383544921875, ('1-HEXENE', '12-BUTADIENE'): 0.00304681062698364,
    ('1-HEXENE', '12PENTADIENE'): 0.000784158706665039, ('1-HEXENE', '13-BUTADIENE'): 0.00292831659317017,
    ('1-HEXENE', '14PENTADIENE'): 0.000774383544921875, ('1-HEXENE', '15-HEXADIENE'): 3.3259391784668e-05,
    ('1-HEXENE', '2-PENTYNE'): 0.000747621059417725, ('1-HEXENE', '23PENTADIENE'): 0.00079500675201416,
    ('1-HEXENE', '26-OCTADIENE'): 0.00128710269927979, ('1-HEXENE', '2E1C4='): 1.37090682983398e-06,
    ('1-HEXENE', '2M1C5='): 1.37090682983398e-06, ('1-HEXENE', '3-M-1-BUTYNE'): 0.00089496374130249,
    ('1-HEXENE', '3M1C5='): 1.37090682983398e-06, ('1-HEXENE', 'ACETYLENE'): 0.0175028443336487,
    ('1-HEXENE', 'BENZENE'): 0.00122594833374023, ('1-HEXENE', 'BIACETYLENE'): 0.00576269626617432,
    ('1-HEXENE', 'CIS2-BUTENE'): 0.00224965810775757, ('1-HEXENE', 'CIS2-HEXENE'): 5.96046447753906e-08,
    ('1-HEXENE', 'CIS2-OCTENE'): 0.00110965967178345, ('1-HEXENE', 'CIS2-PENTENE'): 0.000301122665405273,
    ('1-HEXENE', 'CIS3-HEXENE'): 5.96046447753906e-08, ('1-HEXENE', 'CIS4-OCTENE'): 0.00110965967178345,
    ('1-HEXENE', 'CO2'): 0.101000003516674, ('1-HEXENE', 'CS2'): 0.040765468031168,
    ('1-HEXENE', 'CYCLOBUTANE'): 0.00360417366027832, ('1-HEXENE', 'CYCLOBUTENE'): 0.00469160079956055,
    ('1-HEXENE', 'CYCLOHEXANE'): 0.000226914882659912, ('1-HEXENE', 'CYCLOPENTANE'): 0.00122594833374023,
    ('1-HEXENE', 'CYCLOPENTENE'): 0.00176405906677246, ('1-HEXENE', 'CYCLOPROPANE'): 0.00711864233016968,
    ('1-HEXENE', 'DMACETYLENE'): 0.00292831659317017, ('1-HEXENE', 'E-BENZENE'): 6.11543655395508e-05,
    ('1-HEXENE', 'E-MERCAPTAN'): 0.0575355552136898, ('1-HEXENE', 'EACETYLENE'): 0.00287133455276489,
    ('1-HEXENE', 'ECYCLOPENTAN'): 6.3478946685791e-05, ('1-HEXENE', 'ETHANE'): 0.0177970007061958,
    ('1-HEXENE', 'ETHYLENE'): 0.0136908292770386, ('1-HEXENE', 'H2'): 0.292100012302399,
    ('1-HEXENE', 'H2O'): 0.5, ('1-HEXENE', 'H2S'): 0.0450000017881393,
    ('1-HEXENE', 'I-BUTENE'): 0.00202006101608276, ('1-HEXENE', 'IC4'): 0.00113296508789063,
    ('1-HEXENE', 'IC5'): 0.000226914882659912, ('1-HEXENE', 'M-ACETYLENE'): 0.00792813301086426,
    ('1-HEXENE', 'M-MERCAPTAN'): 0.0575355552136898, ('1-HEXENE', 'M-XYLENE'): 6.99758529663086e-05,
    ('1-HEXENE', 'MCYCLOHEXANE'): 3.49879264831543e-05, ('1-HEXENE', 'MCYCLOPENTAN'): 0.000119388103485107,
    ('1-HEXENE', 'METHANE'): 0.0217461585998535, ('1-HEXENE', 'N-PBENZENE'): 0.000726997852325439,
    ('1-HEXENE', 'NC4'): 0.00139129161834717, ('1-HEXENE', 'NC5'): 0.000193774700164795,
    ('1-HEXENE', 'NC6'): 3.49879264831543e-05, ('1-HEXENE', 'NC7'): 0.000536024570465088,
    ('1-HEXENE', 'NC8'): 0.00149506330490112, ('1-HEXENE', 'NITROGEN'): 0.100000001490116,
    ('1-HEXENE', 'O-XYLENE'): 3.93986701965332e-05, ('1-HEXENE', 'P-XYLENE'): 8.83936882019043e-05,
    ('1-HEXENE', 'PROPADIENE'): 0.00818508863449097, ('1-HEXENE', 'PROPANE'): 0.00433343648910522,
    ('1-HEXENE', 'PROPENE'): 0.00601398944854736, ('1-HEXENE', 'STYRENE'): 5.36441802978516e-07,
    ('1-HEXENE', 'TOLUENE'): 0.00014495849609375, ('1-HEXENE', 'TR2-BUTENE'): 0.00205522775650024,
    ('1-HEXENE', 'TR2-HEXENE'): 5.96046447753906e-08, ('1-HEXENE', 'TR2-OCTENE'): 0.00110965967178345,
    ('1-HEXENE', 'TR2-PENTENE'): 0.000301122665405273, ('1-HEXENE', 'TR3-HEXENE'): 5.96046447753906e-08,
    ('1-HEXENE', 'TR3-OCTENE'): 0.00110965967178345, ('1-HEXENE', 'TR4-OCTENE'): 0.00110965967178345,
    ('1-HEXENE', 'VNYLACETLEN'): 0.00310701131820679,
    ('1-HEXYNE', '1-NONENE'): 0.00432002544403076, ('1-HEXYNE', '1-OCTENE'): 0.00151264667510986,
    ('1-HEXYNE', '1-OCTYNE'): 0.00116336345672607, ('1-HEXYNE', '1-PENTENE'): 0.000155627727508545,
    ('1-HEXYNE', '1-PENTYNE'): 0.000459969043731689, ('1-HEXYNE', '1-TR3-C5=='): 0.000490367412567139,
    ('1-HEXYNE', '12-BUTADIENE'): 0.00245320796966553, ('1-HEXYNE', '12PENTADIENE'): 0.000498116016387939,
    ('1-HEXYNE', '13-BUTADIENE'): 0.00234705209732056, ('1-HEXYNE', '14PENTADIENE'): 0.000490367412567139,
    ('1-HEXYNE', '2-PENTYNE'): 0.0004691481590271, ('1-HEXYNE', '23PENTADIENE'): 0.000506758689880371,
    ('1-HEXYNE', '26-OCTADIENE'): 0.00172674655914307, ('1-HEXYNE', '2E1C4='): 4.70876693725586e-05,
    ('1-HEXYNE', '2M1C5='): 4.70876693725586e-05, ('1-HEXYNE', '3-M-1-BUTYNE'): 0.000587224960327148,
    ('1-HEXYNE', '3M1C5='): 2.02655792236328e-05, ('1-HEXYNE', 'ACETYLENE'): 0.0160602927207947,
    ('1-HEXYNE', 'BENZENE'): 0.000860393047332764, ('1-HEXYNE', 'BIACETYLENE'): 0.0049368143081665,
    ('1-HEXYNE', 'CIS2-BUTENE'): 0.00174355506896973, ('1-HEXYNE', 'CIS2-HEXENE'): 3.59416007995605e-05,
    ('1-HEXYNE', 'CIS2-OCTENE'): 0.00152015686035156, ('1-HEXYNE', 'CIS2-PENTENE'): 0.000136137008666992,
    ('1-HEXYNE', 'CIS3-HEXENE'): 3.59416007995605e-05, ('1-HEXYNE', 'CIS4-OCTENE'): 0.00152015686035156,
    ('1-HEXYNE', 'CO2'): 0.101000003516674, ('1-HEXYNE', 'CS2'): 0.0423111580312252,
    ('1-HEXYNE', 'CYCLOBUTANE'): 0.0029560923576355, ('1-HEXYNE', 'CYCLOBUTENE'): 0.0039486289024353,
    ('1-HEXYNE', 'CYCLOHEXANE'): 8.79764556884766e-05, ('1-HEXYNE', 'CYCLOPENTANE'): 0.000860393047332764,
    ('1-HEXYNE', 'CYCLOPENTENE'): 0.00131940841674805, ('1-HEXYNE', 'CYCLOPROPANE'): 0.00619858503341675,
    ('1-HEXYNE', 'DMACETYLENE'): 0.00234705209732056, ('1-HEXYNE', 'E-BENZENE'): 0.000182449817657471,
    ('1-HEXYNE', 'E-MERCAPTAN'): 0.0540839247405529, ('1-HEXYNE', 'EACETYLENE'): 0.0022958517074585,
    ('1-HEXYNE', 'ECYCLOPENTAN'): 0.000186383724212646, ('1-HEXYNE', 'ETHANE'): 0.0090976357460022,
    ('1-HEXYNE', 'ETHYLENE'): 0.0124127864837646, ('1-HEXYNE', 'H2'): 0.292100012302399,
    ('1-HEXYNE', 'H2O'): 0.5, ('1-HEXYNE', 'H2S'): 0.0450000017881393,
    ('1-HEXYNE', 'I-BUTENE'): 0.00154209136962891, ('1-HEXYNE', 'IC4'): 0.000782787799835205,
    ('1-HEXYNE', 'IC5'): 8.79764556884766e-05, ('1-HEXYNE', 'M-ACETYLENE'): 0.00695639848709106,
    ('1-HEXYNE', 'M-MERCAPTAN'): 0.0540839247405529, ('1-HEXYNE', 'M-XYLENE'): 0.000197410583496094,
    ('1-HEXYNE', 'MCYCLOHEXANE'): 0.000134587287902832, ('1-HEXYNE', 'MCYCLOPENTAN'): 2.74777412414551e-05,
    ('1-HEXYNE', 'METHANE'): 0.0201429724693298, ('1-HEXYNE', 'N-PBENZENE'): 0.00106561183929443,
    ('1-HEXYNE', 'NC4'): 0.000999867916107178, ('1-HEXYNE', 'NC5'): 6.78300857543945e-05,
    ('1-HEXYNE', 'NC6'): 0.000134587287902832, ('1-HEXYNE', 'NC7'): 0.000831484794616699,
    ('1-HEXYNE', 'NC8'): 0.00196623802185059, ('1-HEXYNE', 'NITROGEN'): 0.100000001490116,
    ('1-HEXYNE', 'O-XYLENE'): 0.000143051147460938, ('1-HEXYNE', 'P-XYLENE'): 0.000227689743041992,
    ('1-HEXYNE', 'PROPADIENE'): 0.00719743967056274, ('1-HEXYNE', 'PROPANE'): 0.00362026691436768,
    ('1-HEXYNE', 'PROPENE'): 0.00516974925994873, ('1-HEXYNE', 'STYRENE'): 4.05311584472656e-05,
    ('1-HEXYNE', 'TOLUENE'): 4.03523445129395e-05, ('1-HEXYNE', 'TR2-BUTENE'): 0.00157284736633301,
    ('1-HEXYNE', 'TR2-HEXENE'): 3.59416007995605e-05, ('1-HEXYNE', 'TR2-OCTENE'): 0.00152015686035156,
    ('1-HEXYNE', 'TR2-PENTENE'): 0.000136137008666992, ('1-HEXYNE', 'TR3-HEXENE'): 3.59416007995605e-05,
    ('1-HEXYNE', 'TR3-OCTENE'): 0.00152015686035156, ('1-HEXYNE', 'TR4-OCTENE'): 0.00152015686035156,
    ('1-HEXYNE', 'VNYLACETLEN'): 0.00250726938247681,
    ('1-NONENE', '1-OCTENE'): 0.000724732875823975, ('1-NONENE', '1-OCTYNE'): 0.00100535154342651,
    ('1-NONENE', '1-PENTENE'): 0.00610297918319702, ('1-NONENE', '1-PENTYNE'): 0.0075727105140686,
    ('1-NONENE', '1-TR3-C5=='): 0.00769335031509399, ('1-NONENE', '12-BUTADIENE'): 0.0131767988204956,
    ('1-NONENE', '12PENTADIENE'): 0.00772368907928467, ('1-NONENE', '13-BUTADIENE'): 0.0129324197769165,
    ('1-NONENE', '14PENTADIENE'): 0.00769335031509399, ('1-NONENE', '15-HEXADIENE'): 0.00433063507080078,
    ('1-NONENE', '2-PENTYNE'): 0.00760930776596069, ('1-NONENE', '23PENTADIENE'): 0.00775742530822754,
    ('1-NONENE', '26-OCTADIENE'): 0.000588357448577881, ('1-NONENE', '2E1C4='): 0.00346875190734863,
    ('1-NONENE', '2M1C5='): 0.00346875190734863, ('1-NONENE', '3-M-1-BUTYNE'): 0.00806087255477905,
    ('1-NONENE', '3M1C5='): 0.00375068187713623, ('1-NONENE', 'ACETYLENE'): 0.0362837910652161,
    ('1-NONENE', 'BENZENE'): 0.00899326801300049, ('1-NONENE', 'BIACETYLENE'): 0.018281877040863,
    ('1-NONENE', 'CIS2-BUTENE'): 0.0114740133285522, ('1-NONENE', 'CIS2-HEXENE'): 0.00357115268707275,
    ('1-NONENE', 'CIS2-OCTENE'): 0.00071948766708374, ('1-NONENE', 'CIS2-PENTENE'): 0.00597858428955078,
    ('1-NONENE', 'CIS3-HEXENE'): 0.00357115268707275, ('1-NONENE', 'CIS4-OCTENE'): 0.00071948766708374,
    ('1-NONENE', 'CO2'): 0.101000003516674, ('1-NONENE', 'CS2'): 0.0469782538712025,
    ('1-NONENE', 'CYCLOBUTANE'): 0.0142955183982849, ('1-NONENE', 'CYCLOBUTENE'): 0.0163600444793701,
    ('1-NONENE', 'CYCLOHEXANE'): 0.00563204288482666, ('1-NONENE', 'CYCLOPENTANE'): 0.00899326801300049,
    ('1-NONENE', 'CYCLOPENTENE'): 0.0103523731231689, ('1-NONENE', 'CYCLOPROPANE'): 0.0206016302108765,
    ('1-NONENE', 'DMACETYLENE'): 0.0129324197769165, ('1-NONENE', 'E-BENZENE'): 0.00273299217224121,
    ('1-NONENE', 'E-MERCAPTAN'): 0.0502057857811451, ('1-NONENE', 'EACETYLENE'): 0.0128135681152344,
    ('1-NONENE', 'ECYCLOPENTAN'): 0.00271791219711304, ('1-NONENE', 'ETHANE'): 0.0255553722381592,
    ('1-NONENE', 'ETHYLENE'): 0.0308139324188232, ('1-NONENE', 'H2'): 0.292100012302399,
    ('1-NONENE', 'H2O'): 0.5, ('1-NONENE', 'H2S'): 0.0450000017881393,
    ('1-NONENE', 'I-BUTENE'): 0.0109535455703735, ('1-NONENE', 'IC4'): 0.00874060392379761,
    ('1-NONENE', 'IC5'): 0.00563204288482666, ('1-NONENE', 'M-ACETYLENE'): 0.0219394564628601,
    ('1-NONENE', 'M-MERCAPTAN'): 0.0502057857811451, ('1-NONENE', 'M-XYLENE'): 0.00267666578292847,
    ('1-NONENE', 'MCYCLOHEXANE'): 0.00293523073196411, ('1-NONENE', 'MCYCLOPENTAN'): 0.0050317645072937,
    ('1-NONENE', 'METHANE'): 0.042137086391449, ('1-NONENE', 'N-PBENZENE'): 0.00110042095184326,
    ('1-NONENE', 'NC4'): 0.00942736864089966, ('1-NONENE', 'NC5'): 0.00546300411224365,
    ('1-NONENE', 'NC6'): 0.00293523073196411, ('1-NONENE', 'NC7'): 0.00136750936508179,
    ('1-NONENE', 'NC8'): 0.000460684299468994, ('1-NONENE', 'NITROGEN'): 0.100000001490116,
    ('1-NONENE', 'O-XYLENE'): 0.00289666652679443, ('1-NONENE', 'P-XYLENE'): 0.0025705099105835,
    ('1-NONENE', 'PROPADIENE'): 0.0223580598831177, ('1-NONENE', 'PROPANE'): 0.0156945586204529,
    ('1-NONENE', 'PROPENE'): 0.0187202095985413, ('1-NONENE', 'STYRENE'): 0.00352787971496582,
    ('1-NONENE', 'TOLUENE'): 0.00518995523452759, ('1-NONENE', 'TR2-BUTENE'): 0.0110343098640442,
    ('1-NONENE', 'TR2-HEXENE'): 0.00357115268707275, ('1-NONENE', 'TR2-OCTENE'): 0.00071948766708374,
    ('1-NONENE', 'TR2-PENTENE'): 0.00597858428955078, ('1-NONENE', 'TR3-HEXENE'): 0.00357115268707275,
    ('1-NONENE', 'TR3-OCTENE'): 0.00071948766708374, ('1-NONENE', 'TR4-OCTENE'): 0.00071948766708374,
    ('1-NONENE', 'VNYLACETLEN'): 0.0133001804351807,
    ('1-OCTENE', '1-OCTYNE'): 2.29477882385254e-05, ('1-OCTENE', '1-PENTENE'): 0.00263535976409912,
    ('1-OCTENE', '1-PENTYNE'): 0.00363337993621826, ('1-OCTENE', '1-TR3-C5=='): 0.00371742248535156,
    ('1-OCTENE', '12-BUTADIENE'): 0.00778108835220337, ('1-OCTENE', '12PENTADIENE'): 0.00373870134353638,
    ('1-OCTENE', '13-BUTADIENE'): 0.00759220123291016, ('1-OCTENE', '14PENTADIENE'): 0.00371742248535156,
    ('1-OCTENE', '15-HEXADIENE'): 0.00151890516281128, ('1-OCTENE', '2-PENTYNE'): 0.00365877151489258,
    ('1-OCTENE', '23PENTADIENE'): 0.00376230478286743, ('1-OCTENE', '26-OCTADIENE'): 7.15255737304688e-06,
    ('1-OCTENE', '2E1C4='): 0.00102651119232178, ('1-OCTENE', '2M1C5='): 0.00102651119232178,
    ('1-OCTENE', '3-M-1-BUTYNE'): 0.00397557020187378, ('1-OCTENE', '3M1C5='): 0.0011829137802124,
    ('1-OCTENE', 'ACETYLENE'): 0.0270971059799194, ('1-OCTENE', 'BENZENE'): 0.00464165210723877,
    ('1-OCTENE', 'BIACETYLENE'): 0.0118339061737061, ('1-OCTENE', 'CIS2-BUTENE'): 0.00647807121276855,
    ('1-OCTENE', 'CIS2-HEXENE'): 0.0010826587677002, ('1-OCTENE', 'CIS2-PENTENE'): 0.00255358219146729,
    ('1-OCTENE', 'CIS3-HEXENE'): 0.0010826587677002, ('1-OCTENE', 'CO2'): 0.101000003516674,
    ('1-OCTENE', 'CS2'): 0.0455515347421169, ('1-OCTENE', 'CYCLOBUTANE'): 0.0086522102355957,
    ('1-OCTENE', 'CYCLOBUTENE'): 0.010286271572113, ('1-OCTENE', 'CYCLOHEXANE'): 0.00232785940170288,
    ('1-OCTENE', 'CYCLOPENTANE'): 0.00464165210723877, ('1-OCTENE', 'CYCLOPENTENE'): 0.00563746690750122,
    ('1-OCTENE', 'CYCLOPROPANE'): 0.0137303471565247, ('1-OCTENE', 'DMACETYLENE'): 0.00759220123291016,
    ('1-OCTENE', 'E-BENZENE'): 0.000645339488983154, ('1-OCTENE', 'E-MERCAPTAN'): 0.0534539632499218,
    ('1-OCTENE', 'EACETYLENE'): 0.00750064849853516, ('1-OCTENE', 'ECYCLOPENTAN'): 0.000638008117675781,
    ('1-OCTENE', 'ETHANE'): 0.0178639888763428, ('1-OCTENE', 'ETHYLENE'): 0.0223495364189148,
    ('1-OCTENE', 'H2'): 0.292100012302399, ('1-OCTENE', 'H2O'): 0.5, ('1-OCTENE', 'H2S'): 0.0450000017881393,
    ('1-OCTENE', 'I-BUTENE'): 0.00608611106872559, ('1-OCTENE', 'IC4'): 0.00445961952209473,
    ('1-OCTENE', 'IC5'): 0.00232785940170288, ('1-OCTENE', 'M-ACETYLENE'): 0.0148363709449768,
    ('1-OCTENE', 'M-MERCAPTAN'): 0.0534539632499218, ('1-OCTENE', 'M-XYLENE'): 0.00061798095703125,
    ('1-OCTENE', 'MCYCLOHEXANE'): 0.000745773315429688, ('1-OCTENE', 'MCYCLOPENTAN'): 0.00194656848907471,
    ('1-OCTENE', 'METHANE'): 0.0322508215904236, ('1-OCTENE', 'N-PBENZENE'): 3.91602516174316e-05,
    ('1-OCTENE', 'NC4'): 0.00495684146881104, ('1-OCTENE', 'NC5'): 0.00221925973892212,
    ('1-OCTENE', 'NC6'): 0.000745773315429688, ('1-OCTENE', 'NC7'): 0.000101447105407715,
    ('1-OCTENE', 'NC8'): 2.98023223876953e-05, ('1-OCTENE', 'NITROGEN'): 0.100000001490116,
    ('1-OCTENE', 'O-XYLENE'): 0.000726282596588135, ('1-OCTENE', 'P-XYLENE'): 0.000567436218261719,
    ('1-OCTENE', 'PROPADIENE'): 0.0151839852333069, ('1-OCTENE', 'PROPANE'): 0.00975602865219116,
    ('1-OCTENE', 'PROPENE'): 0.0121899843215942, ('1-OCTENE', 'STYRENE'): 0.00105887651443481,
    ('1-OCTENE', 'TOLUENE'): 0.00204569101333618, ('1-OCTENE', 'TR2-BUTENE'): 0.00614672899246216,
    ('1-OCTENE', 'TR2-HEXENE'): 0.0010826587677002, ('1-OCTENE', 'TR2-PENTENE'): 0.00255358219146729,
    ('1-OCTENE', 'TR3-HEXENE'): 0.0010826587677002, ('1-OCTENE', 'VNYLACETLEN'): 0.00787651538848877,
    ('1-OCTYNE', '1-PENTENE'): 0.00216764211654663, ('1-OCTYNE', '1-PENTYNE'): 0.00308078527450562,
    ('1-OCTYNE', '1-TR3-C5=='): 0.0031583309173584, ('1-OCTYNE', '12-BUTADIENE'): 0.00696587562561035,
    ('1-OCTYNE', '12PENTADIENE'): 0.0031779408454895, ('1-OCTYNE', '13-BUTADIENE'): 0.00678694248199463,
    ('1-OCTYNE', '14PENTADIENE'): 0.0031583309173584, ('1-OCTYNE', '15-HEXADIENE'): 0.00116890668869019,
    ('1-OCTYNE', '2-PENTYNE'): 0.00310420989990234, ('1-OCTYNE', '23PENTADIENE'): 0.00319963693618774,
    ('1-OCTYNE', '26-OCTADIENE'): 5.56707382202148e-05, ('1-OCTYNE', '2E1C4='): 0.000742673873901367,
    ('1-OCTYNE', '2M1C5='): 0.000742673873901367, ('1-OCTYNE', '3-M-1-BUTYNE'): 0.00339674949645996,
    ('1-OCTYNE', '3M1C5='): 0.000876486301422119, ('1-OCTYNE', 'ACETYLENE'): 0.0255923867225647,
    ('1-OCTYNE', 'BENZENE'): 0.00401479005813599, ('1-OCTYNE', 'BIACETYLENE'): 0.0108282566070557,
    ('1-OCTYNE', 'CIS2-BUTENE'): 0.00573498010635376, ('1-OCTYNE', 'CIS2-HEXENE'): 0.000790596008300781,
    ('1-OCTYNE', 'CIS2-OCTENE'): 2.38418579101563e-05, ('1-OCTYNE', 'CIS2-PENTENE'): 0.00209337472915649,
    ('1-OCTYNE', 'CIS3-HEXENE'): 0.000790596008300781, ('1-OCTYNE', 'CIS4-OCTENE'): 2.38418579101563e-05,
    ('1-OCTYNE', 'CO2'): 0.101000003516674, ('1-OCTYNE', 'CS2'): 0.0463193133473396,
    ('1-OCTYNE', 'CYCLOBUTANE'): 0.00779211521148682, ('1-OCTYNE', 'CYCLOBUTENE'): 0.00934839248657227,
    ('1-OCTYNE', 'CYCLOHEXANE'): 0.00188946723937988, ('1-OCTYNE', 'CYCLOPENTANE'): 0.00401479005813599,
    ('1-OCTYNE', 'CYCLOPENTENE'): 0.00494527816772461, ('1-OCTYNE', 'CYCLOPROPANE'): 0.0126477479934692,
    ('1-OCTYNE', 'DMACETYLENE'): 0.00678694248199463, ('1-OCTYNE', 'E-BENZENE'): 0.000424981117248535,
    ('1-OCTYNE', 'E-MERCAPTAN'): 0.0492285564541817, ('1-OCTYNE', 'EACETYLENE'): 0.00670033693313599,
    ('1-OCTYNE', 'ECYCLOPENTAN'): 0.000418961048126221, ('1-OCTYNE', 'ETHANE'): 0.0166323781013489,
    ('1-OCTYNE', 'ETHYLENE'): 0.0209768414497375, ('1-OCTYNE', 'H2'): 0.292100012302399,
    ('1-OCTYNE', 'H2O'): 0.5, ('1-OCTYNE', 'H2S'): 0.0450000017881393,
    ('1-OCTYNE', 'I-BUTENE'): 0.00536626577377319, ('1-OCTYNE', 'IC4'): 0.00384557247161865,
    ('1-OCTYNE', 'IC5'): 0.00188946723937988, ('1-OCTYNE', 'M-ACETYLENE'): 0.0137117505073547,
    ('1-OCTYNE', 'M-MERCAPTAN'): 0.0492285564541817, ('1-OCTYNE', 'M-XYLENE'): 0.000402748584747314,
    ('1-OCTYNE', 'MCYCLOHEXANE'): 0.000507056713104248, ('1-OCTYNE', 'MCYCLOPENTAN'): 0.00154739618301392,
    ('1-OCTYNE', 'METHANE'): 0.0306179523468018, ('1-OCTYNE', 'N-PBENZENE'): 2.08616256713867e-06,
    ('1-OCTYNE', 'NC4'): 0.00430846214294434, ('1-OCTYNE', 'NC5'): 0.00179165601730347,
    ('1-OCTYNE', 'NC6'): 0.000507056713104248, ('1-OCTYNE', 'NC7'): 2.78949737548828e-05,
    ('1-OCTYNE', 'NC8'): 0.000105142593383789, ('1-OCTYNE', 'NITROGEN'): 0.100000001490116,
    ('1-OCTYNE', 'O-XYLENE'): 0.000491082668304443, ('1-OCTYNE', 'P-XYLENE'): 0.000362157821655273,
    ('1-OCTYNE', 'PROPADIENE'): 0.0140465497970581, ('1-OCTYNE', 'PROPANE'): 0.00884270668029785,
    ('1-OCTYNE', 'PROPENE'): 0.0111694931983948, ('1-OCTYNE', 'STYRENE'): 0.000770151615142822,
    ('1-OCTYNE', 'TOLUENE'): 0.0016360878944397, ('1-OCTYNE', 'TR2-BUTENE'): 0.00542318820953369,
    ('1-OCTYNE', 'TR2-HEXENE'): 0.000790596008300781, ('1-OCTYNE', 'TR2-OCTENE'): 2.38418579101563e-05,
    ('1-OCTYNE', 'TR2-PENTENE'): 0.00209337472915649, ('1-OCTYNE', 'TR3-HEXENE'): 0.000790596008300781,
    ('1-OCTYNE', 'TR3-OCTENE'): 2.38418579101563e-05, ('1-OCTYNE', 'TR4-OCTENE'): 2.38418579101563e-05,
    ('1-OCTYNE', 'VNYLACETLEN'): 0.00705635547637939,
    ('1-PENTENE', '1-PENTYNE'): 8.05258750915527e-05, ('1-PENTENE', '1-TR3-C5=='): 9.35196876525879e-05,
    ('1-PENTENE', '12-BUTADIENE'): 0.00137519836425781, ('1-PENTENE', '12PENTADIENE'): 9.69767570495605e-05,
    ('1-PENTENE', '13-BUTADIENE'): 0.00129592418670654, ('1-PENTENE', '14PENTADIENE'): 9.35196876525879e-05,
    ('1-PENTENE', '15-HEXADIENE'): 0.000153601169586182, ('1-PENTENE', '2-PENTYNE'): 8.44001770019531e-05,
    ('1-PENTENE', '23PENTADIENE'): 0.000100791454315186, ('1-PENTENE', '26-OCTADIENE'): 0.00291526317596436,
    ('1-PENTENE', '2E1C4='): 0.000373959541320801, ('1-PENTENE', '2M1C5='): 0.000373959541320801,
    ('1-PENTENE', '3-M-1-BUTYNE'): 0.000138342380523682, ('1-PENTENE', '3M1C5='): 0.000288307666778564,
    ('1-PENTENE', 'ACETYLENE'): 0.0131058096885681, ('1-PENTENE', 'BENZENE'): 0.000284492969512939,
    ('1-PENTENE', 'BIACETYLENE'): 0.00334668159484863, ('1-PENTENE', 'CIS2-BUTENE'): 0.000858664512634277,
    ('1-PENTENE', 'CIS2-HEXENE'): 0.000341117382049561, ('1-PENTENE', 'CIS2-OCTENE'): 0.00264531373977661,
    ('1-PENTENE', 'CIS2-PENTENE'): 7.15255737304688e-07, ('1-PENTENE', 'CIS3-HEXENE'): 0.000341117382049561,
    ('1-PENTENE', 'CIS4-OCTENE'): 0.00264531373977661, ('1-PENTENE', 'CO2'): 0.101000003516674,
    ('1-PENTENE', 'CS2'): 0.0373697280883789, ('1-PENTENE', 'CYCLOBUTANE'): 0.00175821781158447,
    ('1-PENTENE', 'CYCLOBUTENE'): 0.00254142284393311, ('1-PENTENE', 'CYCLOHEXANE'): 9.59634780883789e-06,
    ('1-PENTENE', 'CYCLOPENTANE'): 0.000284492969512939, ('1-PENTENE', 'CYCLOPENTENE'): 0.000569343566894531,
    ('1-PENTENE', 'CYCLOPROPANE'): 0.00440067052841187, ('1-PENTENE', 'DMACETYLENE'): 0.00129592418670654,
    ('1-PENTENE', 'E-BENZENE'): 0.000674724578857422, ('1-PENTENE', 'E-MERCAPTAN'): 0.0588542185723782,
    ('1-PENTENE', 'EACETYLENE'): 0.00125789642333984, ('1-PENTENE', 'ECYCLOPENTAN'): 0.000682294368743896,
    ('1-PENTENE', 'ETHANE'): 0.00689405202865601, ('1-PENTENE', 'ETHYLENE'): 0.00982266664505005,
    ('1-PENTENE', 'H2'): 0.292100012302399, ('1-PENTENE', 'H2O'): 0.5,
    ('1-PENTENE', 'H2S'): 0.0450000017881393, ('1-PENTENE', 'I-BUTENE'): 0.000718891620635986,
    ('1-PENTENE', 'IC4'): 0.000240564346313477, ('1-PENTENE', 'IC5'): 9.59634780883789e-06,
    ('1-PENTENE', 'M-ACETYLENE'): 0.00504416227340698, ('1-PENTENE', 'M-MERCAPTAN'): 0.0588542185723782,
    ('1-PENTENE', 'M-XYLENE'): 0.000703334808349609, ('1-PENTENE', 'MCYCLOHEXANE'): 0.000579357147216797,
    ('1-PENTENE', 'MCYCLOPENTAN'): 5.2332878112793e-05, ('1-PENTENE', 'METHANE'): 0.0168318152427673,
    ('1-PENTENE', 'N-PBENZENE'): 0.00203359127044678, ('1-PENTENE', 'NC4'): 0.000366866588592529,
    ('1-PENTENE', 'NC5'): 1.79409980773926e-05, ('1-PENTENE', 'NC6'): 0.000579357147216797,
    ('1-PENTENE', 'NC7'): 0.0017048716545105, ('1-PENTENE', 'NC8'): 0.00322377681732178,
    ('1-PENTENE', 'NITROGEN'): 0.100000001490116, ('1-PENTENE', 'O-XYLENE'): 0.000596821308135986,
    ('1-PENTENE', 'P-XYLENE'): 0.000759482383728027, ('1-PENTENE', 'PROPADIENE'): 0.00525033473968506,
    ('1-PENTENE', 'PROPANE'): 0.00227898359298706, ('1-PENTENE', 'PROPENE'): 0.00353944301605225,
    ('1-PENTENE', 'STYRENE'): 0.0003548264503479, ('1-PENTENE', 'TOLUENE'): 3.74317169189453e-05,
    ('1-PENTENE', 'TR2-BUTENE'): 0.000739991664886475, ('1-PENTENE', 'TR2-HEXENE'): 0.000341117382049561,
    ('1-PENTENE', 'TR2-OCTENE'): 0.00264531373977661, ('1-PENTENE', 'TR2-PENTENE'): 7.15255737304688e-07,
    ('1-PENTENE', 'TR3-HEXENE'): 0.000341117382049561, ('1-PENTENE', 'TR3-OCTENE'): 0.00264531373977661,
    ('1-PENTENE', 'TR4-OCTENE'): 0.00264531373977661, ('1-PENTENE', 'VNYLACETLEN'): 0.00141578912734985,
    ('1-PENTYNE', '1-TR3-C5=='): 4.76837158203125e-07, ('1-PENTYNE', '12-BUTADIENE'): 0.000790715217590332,
    ('1-PENTYNE', '12PENTADIENE'): 8.34465026855469e-07, ('1-PENTYNE', '13-BUTADIENE'): 0.000730752944946289,
    ('1-PENTYNE', '14PENTADIENE'): 4.76837158203125e-07, ('1-PENTYNE', '15-HEXADIENE'): 0.000456571578979492,
    ('1-PENTYNE', '23PENTADIENE'): 1.07288360595703e-06, ('1-PENTYNE', '26-OCTADIENE'): 0.00396043062210083,
    ('1-PENTYNE', '2E1C4='): 0.000801205635070801, ('1-PENTYNE', '2M1C5='): 0.000801205635070801,
    ('1-PENTYNE', '3-M-1-BUTYNE'): 7.80820846557617e-06, ('1-PENTYNE', '3M1C5='): 0.000673472881317139,
    ('1-PENTYNE', 'ACETYLENE'): 0.0111598372459412, ('1-PENTYNE', 'BENZENE'): 6.22272491455078e-05,
    ('1-PENTYNE', 'BIACETYLENE'): 0.00239187479019165, ('1-PENTYNE', 'CIS2-BUTENE'): 0.000413417816162109,
    ('1-PENTYNE', 'CIS2-HEXENE'): 0.000752925872802734, ('1-PENTYNE', 'CIS2-OCTENE'): 0.00364494323730469,
    ('1-PENTYNE', 'CIS2-PENTENE'): 9.5665454864502e-05, ('1-PENTYNE', 'CIS3-HEXENE'): 0.000752925872802734,
    ('1-PENTYNE', 'CIS4-OCTENE'): 0.00364494323730469, ('1-PENTYNE', 'CO2'): 0.101000003516674,
    ('1-PENTYNE', 'CS2'): 0.0393104813992977, ('1-PENTYNE', 'CYCLOBUTANE'): 0.00108706951141357,
    ('1-PENTYNE', 'CYCLOBUTENE'): 0.00171893835067749, ('1-PENTYNE', 'CYCLOHEXANE'): 0.000145792961120605,
    ('1-PENTYNE', 'CYCLOPENTANE'): 6.22272491455078e-05, ('1-PENTYNE', 'CYCLOPENTENE'): 0.000221729278564453,
    ('1-PENTYNE', 'CYCLOPROPANE'): 0.00329530239105225, ('1-PENTYNE', 'DMACETYLENE'): 0.000730752944946289,
    ('1-PENTYNE', 'E-BENZENE'): 0.00122088193893433, ('1-PENTYNE', 'E-MERCAPTAN'): 0.0560731887817383,
    ('1-PENTYNE', 'EACETYLENE'): 0.000702381134033203, ('1-PENTYNE', 'ECYCLOPENTAN'): 0.00123101472854614,
    ('1-PENTYNE', 'ETHANE'): 0.00549429655075073, ('1-PENTYNE', 'ETHYLENE'): 0.00814205408096313,
    ('1-PENTYNE', 'H2'): 0.292100012302399, ('1-PENTYNE', 'H2O'): 0.5,
    ('1-PENTYNE', 'H2S'): 0.0450000017881393, ('1-PENTYNE', 'I-BUTENE'): 0.000318229198455811,
    ('1-PENTYNE', 'IC4'): 4.27961349487305e-05, ('1-PENTYNE', 'IC5'): 0.000145792961120605,
    ('1-PENTYNE', 'M-ACETYLENE'): 0.00385594367980957, ('1-PENTYNE', 'M-MERCAPTAN'): 0.0560731887817383,
    ('1-PENTYNE', 'M-XYLENE'): 0.00125908851623535, ('1-PENTYNE', 'MCYCLOHEXANE'): 0.00109148025512695,
    ('1-PENTYNE', 'MCYCLOPENTAN'): 0.000262856483459473, ('1-PENTYNE', 'METHANE'): 0.0146257877349854,
    ('1-PENTYNE', 'N-PBENZENE'): 0.00292074680328369, ('1-PENTYNE', 'NC4'): 0.000103592872619629,
    ('1-PENTYNE', 'NC5'): 0.000174641609191895, ('1-PENTYNE', 'NC6'): 0.00109148025512695,
    ('1-PENTYNE', 'NC7'): 0.00252437591552734, ('1-PENTYNE', 'NC8'): 0.00431811809539795,
    ('1-PENTYNE', 'NITROGEN'): 0.100000001490116, ('1-PENTYNE', 'O-XYLENE'): 0.00111526250839233,
    ('1-PENTYNE', 'P-XYLENE'): 0.00133389234542847, ('1-PENTYNE', 'PROPADIENE'): 0.00403672456741333,
    ('1-PENTYNE', 'PROPANE'): 0.00150418281555176, ('1-PENTYNE', 'PROPENE'): 0.00255537033081055,
    ('1-PENTYNE', 'STYRENE'): 0.000773131847381592, ('1-PENTYNE', 'TOLUENE'): 0.000227928161621094,
    ('1-PENTYNE', 'TR2-BUTENE'): 0.000332355499267578, ('1-PENTYNE', 'TR2-HEXENE'): 0.000752925872802734,
    ('1-PENTYNE', 'TR2-OCTENE'): 0.00364494323730469, ('1-PENTYNE', 'TR2-PENTENE'): 9.5665454864502e-05,
    ('1-PENTYNE', 'TR3-HEXENE'): 0.000752925872802734, ('1-PENTYNE', 'TR3-OCTENE'): 0.00364494323730469,
    ('1-PENTYNE', 'TR4-OCTENE'): 0.00364494323730469, ('1-PENTYNE', 'VNYLACETLEN'): 0.000821590423583984,
    ('1-TR3-C5==', '12-BUTADIENE'): 0.000751972198486328,
    ('1-TR3-C5==', '13-BUTADIENE'): 0.000693619251251221,
    ('1-TR3-C5==', '15-HEXADIENE'): 0.000486671924591064, ('1-TR3-C5==', '2-PENTYNE'): 1.78813934326172e-07,
    ('1-TR3-C5==', '23PENTADIENE'): 1.78813934326172e-07, ('1-TR3-C5==', '26-OCTADIENE'): 0.00404816865921021,
    ('1-TR3-C5==', '2E1C4='): 0.000841140747070313, ('1-TR3-C5==', '2M1C5='): 0.000841140747070313,
    ('1-TR3-C5==', '3-M-1-BUTYNE'): 4.29153442382813e-06, ('1-TR3-C5==', '3M1C5='): 0.000710129737854004,
    ('1-TR3-C5==', 'ACETYLENE'): 0.0110151171684265, ('1-TR3-C5==', 'BENZENE'): 5.17368316650391e-05,
    ('1-TR3-C5==', 'BIACETYLENE'): 0.00232440233230591, ('1-TR3-C5==', 'CIS2-BUTENE'): 0.000385582447052002,
    ('1-TR3-C5==', 'CIS2-HEXENE'): 0.000791668891906738, ('1-TR3-C5==', 'CIS2-OCTENE'): 0.00372916460037231,
    ('1-TR3-C5==', 'CIS2-PENTENE'): 0.00010979175567627, ('1-TR3-C5==', 'CIS3-HEXENE'): 0.000791668891906738,
    ('1-TR3-C5==', 'CIS4-OCTENE'): 0.00372916460037231, ('1-TR3-C5==', 'CO2'): 0.101000003516674,
    ('1-TR3-C5==', 'CS2'): 0.0391802303493023, ('1-TR3-C5==', 'CYCLOBUTANE'): 0.0010417103767395,
    ('1-TR3-C5==', 'CYCLOBUTENE'): 0.00166171789169312, ('1-TR3-C5==', 'CYCLOHEXANE'): 0.000163018703460693,
    ('1-TR3-C5==', 'CYCLOPENTANE'): 5.17368316650391e-05, ('1-TR3-C5==', 'CYCLOPENTENE'): 0.00020146369934082,
    ('1-TR3-C5==', 'CYCLOPROPANE'): 0.00321614742279053, ('1-TR3-C5==', 'DMACETYLENE'): 0.000693619251251221,
    ('1-TR3-C5==', 'E-BENZENE'): 0.00126999616622925, ('1-TR3-C5==', 'E-MERCAPTAN'): 0.0570147968828678,
    ('1-TR3-C5==', 'EACETYLENE'): 0.000665903091430664, ('1-TR3-C5==', 'ECYCLOPENTAN'): 0.00128030776977539,
    ('1-TR3-C5==', 'ETHANE'): 0.00539225339889526, ('1-TR3-C5==', 'ETHYLENE'): 0.0080181360244751,
    ('1-TR3-C5==', 'H2'): 0.292100012302399, ('1-TR3-C5==', 'H2O'): 0.5,
    ('1-TR3-C5==', 'H2S'): 0.0450000017881393, ('1-TR3-C5==', 'I-BUTENE'): 0.000293910503387451,
    ('1-TR3-C5==', 'IC4'): 3.4034252166748e-05, ('1-TR3-C5==', 'IC5'): 0.000163018703460693,
    ('1-TR3-C5==', 'M-ACETYLENE'): 0.00377029180526733, ('1-TR3-C5==', 'M-MERCAPTAN'): 0.0570147968828678,
    ('1-TR3-C5==', 'M-XYLENE'): 0.00130897760391235, ('1-TR3-C5==', 'MCYCLOHEXANE'): 0.00113797187805176,
    ('1-TR3-C5==', 'MCYCLOPENTAN'): 0.000285923480987549, ('1-TR3-C5==', 'METHANE'): 0.0144607424736023,
    ('1-TR3-C5==', 'N-PBENZENE'): 0.00299620628356934, ('1-TR3-C5==', 'NC4'): 8.99434089660645e-05,
    ('1-TR3-C5==', 'NC5'): 0.000193595886230469, ('1-TR3-C5==', 'NC6'): 0.00113797187805176,
    ('1-TR3-C5==', 'NC7'): 0.00259464979171753, ('1-TR3-C5==', 'NC8'): 0.00440961122512817,
    ('1-TR3-C5==', 'NITROGEN'): 0.100000001490116, ('1-TR3-C5==', 'O-XYLENE'): 0.00116223096847534,
    ('1-TR3-C5==', 'P-XYLENE'): 0.00138527154922485, ('1-TR3-C5==', 'PROPADIENE'): 0.00394898653030396,
    ('1-TR3-C5==', 'PROPANE'): 0.00145077705383301, ('1-TR3-C5==', 'PROPENE'): 0.00248563289642334,
    ('1-TR3-C5==', 'STYRENE'): 0.000812351703643799, ('1-TR3-C5==', 'TOLUENE'): 0.000249505043029785,
    ('1-TR3-C5==', 'TR2-BUTENE'): 0.00030750036239624, ('1-TR3-C5==', 'TR2-HEXENE'): 0.000791668891906738,
    ('1-TR3-C5==', 'TR2-OCTENE'): 0.00372916460037231, ('1-TR3-C5==', 'TR2-PENTENE'): 0.00010979175567627,
    ('1-TR3-C5==', 'TR3-HEXENE'): 0.000791668891906738, ('1-TR3-C5==', 'TR3-OCTENE'): 0.00372916460037231,
    ('1-TR3-C5==', 'TR4-OCTENE'): 0.00372916460037231, ('1-TR3-C5==', 'VNYLACETLEN'): 0.000782191753387451,
    ('12-BUTADIENE', '12PENTADIENE'): 0.000742495059967041,
    ('12-BUTADIENE', '13-BUTADIENE'): 1.19209289550781e-06,
    ('12-BUTADIENE', '14PENTADIENE'): 0.000751972198486328,
    ('12-BUTADIENE', '15-HEXADIENE'): 0.00244522094726563,
    ('12-BUTADIENE', '2-PENTYNE'): 0.000778913497924805,
    ('12-BUTADIENE', '23PENTADIENE'): 0.000731945037841797,
    ('12-BUTADIENE', '26-OCTADIENE'): 0.00825381278991699, ('12-BUTADIENE', '2E1C4='): 0.00317764282226563,
    ('12-BUTADIENE', '2M1C5='): 0.00317764282226563, ('12-BUTADIENE', '3-M-1-BUTYNE'): 0.000641763210296631,
    ('12-BUTADIENE', '3M1C5='): 0.00291842222213745, ('12-BUTADIENE', 'ACETYLENE'): 0.00605422258377075,
    ('12-BUTADIENE', 'BENZENE'): 0.000409424304962158, ('12-BUTADIENE', 'BIACETYLENE'): 0.000433564186096191,
    ('12-BUTADIENE', 'CIS2-BUTENE'): 6.0737133026123e-05,
    ('12-BUTADIENE', 'CIS2-HEXENE'): 0.00308096408843994, ('12-BUTADIENE', 'CIS2-OCTENE'): 0.0077979564666748,
    ('12-BUTADIENE', 'CIS2-PENTENE'): 0.00143551826477051,
    ('12-BUTADIENE', 'CIS3-HEXENE'): 0.00308096408843994, ('12-BUTADIENE', 'CIS4-OCTENE'): 0.0077979564666748,
    ('12-BUTADIENE', 'CO2'): 0.101000003516674, ('12-BUTADIENE', 'CS2'): 0.03591638058424,
    ('12-BUTADIENE', 'CYCLOBUTANE'): 2.34842300415039e-05,
    ('12-BUTADIENE', 'CYCLOBUTENE'): 0.00017845630645752,
    ('12-BUTADIENE', 'CYCLOHEXANE'): 0.00161409378051758,
    ('12-BUTADIENE', 'CYCLOPENTANE'): 0.000409424304962158,
    ('12-BUTADIENE', 'CYCLOPENTENE'): 0.000175178050994873,
    ('12-BUTADIENE', 'CYCLOPROPANE'): 0.000861048698425293,
    ('12-BUTADIENE', 'DMACETYLENE'): 1.19209289550781e-06, ('12-BUTADIENE', 'E-BENZENE'): 0.00396692752838135,
    ('12-BUTADIENE', 'E-MERCAPTAN'): 0.0580847598612309, ('12-BUTADIENE', 'EACETYLENE'): 2.68220901489258e-06,
    ('12-BUTADIENE', 'ECYCLOPENTAN'): 0.00398516654968262, ('12-BUTADIENE', 'ETHANE'): 0.00212746858596802,
    ('12-BUTADIENE', 'ETHYLENE'): 0.00388270616531372, ('12-BUTADIENE', 'H2'): 0.292100012302399,
    ('12-BUTADIENE', 'H2O'): 0.5, ('12-BUTADIENE', 'H2S'): 0.0450000017881393,
    ('12-BUTADIENE', 'I-BUTENE'): 0.000105857849121094, ('12-BUTADIENE', 'IC4'): 0.000466048717498779,
    ('12-BUTADIENE', 'IC5'): 0.00161409378051758, ('12-BUTADIENE', 'M-ACETYLENE'): 0.0011594295501709,
    ('12-BUTADIENE', 'M-MERCAPTAN'): 0.0580847598612309, ('12-BUTADIENE', 'M-XYLENE'): 0.0040353536605835,
    ('12-BUTADIENE', 'MCYCLOHEXANE'): 0.00373154878616333,
    ('12-BUTADIENE', 'MCYCLOPENTAN'): 0.00196301937103271, ('12-BUTADIENE', 'METHANE'): 0.00868862867355347,
    ('12-BUTADIENE', 'N-PBENZENE'): 0.00672531127929688, ('12-BUTADIENE', 'NC4'): 0.000322103500366211,
    ('12-BUTADIENE', 'NC5'): 0.00170695781707764, ('12-BUTADIENE', 'NC6'): 0.00373154878616333,
    ('12-BUTADIENE', 'NC7'): 0.00611913204193115, ('12-BUTADIENE', 'NC8'): 0.00876390933990479,
    ('12-BUTADIENE', 'NITROGEN'): 0.100000001490116, ('12-BUTADIENE', 'O-XYLENE'): 0.00377529859542847,
    ('12-BUTADIENE', 'P-XYLENE'): 0.00416791439056396, ('12-BUTADIENE', 'PROPADIENE'): 0.00125980377197266,
    ('12-BUTADIENE', 'PROPANE'): 0.000114023685455322, ('12-BUTADIENE', 'PROPENE'): 0.000504970550537109,
    ('12-BUTADIENE', 'STYRENE'): 0.00312149524688721, ('12-BUTADIENE', 'TOLUENE'): 0.00186574459075928,
    ('12-BUTADIENE', 'TR2-BUTENE'): 9.78708267211914e-05, ('12-BUTADIENE', 'TR2-HEXENE'): 0.00308096408843994,
    ('12-BUTADIENE', 'TR2-OCTENE'): 0.0077979564666748, ('12-BUTADIENE', 'TR2-PENTENE'): 0.00143551826477051,
    ('12-BUTADIENE', 'TR3-HEXENE'): 0.00308096408843994, ('12-BUTADIENE', 'TR3-OCTENE'): 0.0077979564666748,
    ('12-BUTADIENE', 'TR4-OCTENE'): 0.0077979564666748, ('12-BUTADIENE', 'VNYLACETLEN'): 2.98023223876953e-07,
    ('12PENTADIENE', '13-BUTADIENE'): 0.000684440135955811,
    ('12PENTADIENE', '15-HEXADIENE'): 0.000494539737701416,
    ('12PENTADIENE', '2-PENTYNE'): 4.76837158203125e-07,
    ('12PENTADIENE', '26-OCTADIENE'): 0.00407034158706665, ('12PENTADIENE', '2E1C4='): 0.000851273536682129,
    ('12PENTADIENE', '2M1C5='): 0.000851273536682129, ('12PENTADIENE', '3-M-1-BUTYNE'): 3.69548797607422e-06,
    ('12PENTADIENE', '3M1C5='): 0.000719428062438965, ('12PENTADIENE', 'ACETYLENE'): 0.0109788179397583,
    ('12PENTADIENE', 'BENZENE'): 4.9293041229248e-05, ('12PENTADIENE', 'BIACETYLENE'): 0.00230765342712402,
    ('12PENTADIENE', 'CIS2-BUTENE'): 0.000378727912902832,
    ('12PENTADIENE', 'CIS2-HEXENE'): 0.000801563262939453,
    ('12PENTADIENE', 'CIS2-OCTENE'): 0.0037505030632019,
    ('12PENTADIENE', 'CIS2-PENTENE'): 0.000113546848297119,
    ('12PENTADIENE', 'CIS3-HEXENE'): 0.000801563262939453,
    ('12PENTADIENE', 'CIS4-OCTENE'): 0.0037505030632019, ('12PENTADIENE', 'CO2'): 0.101000003516674,
    ('12PENTADIENE', 'CS2'): 0.0397308394312859, ('12PENTADIENE', 'CYCLOBUTANE'): 0.00103038549423218,
    ('12PENTADIENE', 'CYCLOBUTENE'): 0.00164765119552612,
    ('12PENTADIENE', 'CYCLOHEXANE'): 0.000167548656463623,
    ('12PENTADIENE', 'CYCLOPENTANE'): 4.9293041229248e-05,
    ('12PENTADIENE', 'CYCLOPENTENE'): 0.000196516513824463,
    ('12PENTADIENE', 'CYCLOPROPANE'): 0.0031963586807251,
    ('12PENTADIENE', 'DMACETYLENE'): 0.000684440135955811, ('12PENTADIENE', 'E-BENZENE'): 0.00128257274627686,
    ('12PENTADIENE', 'E-MERCAPTAN'): 0.0559986978769302, ('12PENTADIENE', 'EACETYLENE'): 0.000656962394714355,
    ('12PENTADIENE', 'ECYCLOPENTAN'): 0.001292884349823, ('12PENTADIENE', 'ETHANE'): 0.00536662340164185,
    ('12PENTADIENE', 'ETHYLENE'): 0.00798708200454712, ('12PENTADIENE', 'H2'): 0.292100012302399,
    ('12PENTADIENE', 'H2O'): 0.5, ('12PENTADIENE', 'H2S'): 0.0450000017881393,
    ('12PENTADIENE', 'I-BUTENE'): 0.000287890434265137, ('12PENTADIENE', 'IC4'): 3.21269035339355e-05,
    ('12PENTADIENE', 'IC5'): 0.000167548656463623, ('12PENTADIENE', 'M-ACETYLENE'): 0.00374889373779297,
    ('12PENTADIENE', 'M-MERCAPTAN'): 0.0559986978769302, ('12PENTADIENE', 'M-XYLENE'): 0.00132161378860474,
    ('12PENTADIENE', 'MCYCLOHEXANE'): 0.00114977359771729,
    ('12PENTADIENE', 'MCYCLOPENTAN'): 0.000291883945465088, ('12PENTADIENE', 'METHANE'): 0.0144195556640625,
    ('12PENTADIENE', 'N-PBENZENE'): 0.00301539897918701, ('12PENTADIENE', 'NC4'): 8.6665153503418e-05,
    ('12PENTADIENE', 'NC5'): 0.0001983642578125, ('12PENTADIENE', 'NC6'): 0.00114977359771729,
    ('12PENTADIENE', 'NC7'): 0.00261247158050537, ('12PENTADIENE', 'NC8'): 0.00443273782730103,
    ('12PENTADIENE', 'NITROGEN'): 0.100000001490116, ('12PENTADIENE', 'O-XYLENE'): 0.00117427110671997,
    ('12PENTADIENE', 'P-XYLENE'): 0.00139826536178589, ('12PENTADIENE', 'PROPADIENE'): 0.00392723083496094,
    ('12PENTADIENE', 'PROPANE'): 0.00143742561340332, ('12PENTADIENE', 'PROPENE'): 0.00246834754943848,
    ('12PENTADIENE', 'STYRENE'): 0.000822365283966064, ('12PENTADIENE', 'TOLUENE'): 0.000255048274993896,
    ('12PENTADIENE', 'TR2-BUTENE'): 0.000301361083984375,
    ('12PENTADIENE', 'TR2-HEXENE'): 0.000801563262939453, ('12PENTADIENE', 'TR2-OCTENE'): 0.0037505030632019,
    ('12PENTADIENE', 'TR2-PENTENE'): 0.000113546848297119,
    ('12PENTADIENE', 'TR3-HEXENE'): 0.000801563262939453, ('12PENTADIENE', 'TR3-OCTENE'): 0.0037505030632019,
    ('12PENTADIENE', 'TR4-OCTENE'): 0.0037505030632019, ('12PENTADIENE', 'VNYLACETLEN'): 0.000772416591644287,
    ('13-BUTADIENE', '14PENTADIENE'): 0.000693619251251221,
    ('13-BUTADIENE', '15-HEXADIENE'): 0.00233918428421021,
    ('13-BUTADIENE', '2-PENTYNE'): 0.000719428062438965,
    ('13-BUTADIENE', '23PENTADIENE'): 0.000674426555633545,
    ('13-BUTADIENE', '26-OCTADIENE'): 0.00805944204330444, ('13-BUTADIENE', '2E1C4='): 0.00305658578872681,
    ('13-BUTADIENE', '2M1C5='): 0.00305658578872681, ('13-BUTADIENE', '3-M-1-BUTYNE'): 0.000587999820709229,
    ('13-BUTADIENE', '3M1C5='): 0.00280243158340454, ('13-BUTADIENE', 'ACETYLENE'): 0.00622332096099854,
    ('13-BUTADIENE', 'BENZENE'): 0.000366508960723877, ('13-BUTADIENE', 'BIACETYLENE'): 0.000480055809020996,
    ('13-BUTADIENE', 'CIS2-BUTENE'): 4.50015068054199e-05,
    ('13-BUTADIENE', 'CIS2-HEXENE'): 0.00296193361282349,
    ('13-BUTADIENE', 'CIS2-OCTENE'): 0.00760889053344727,
    ('13-BUTADIENE', 'CIS2-PENTENE'): 0.00135445594787598,
    ('13-BUTADIENE', 'CIS3-HEXENE'): 0.00296193361282349,
    ('13-BUTADIENE', 'CIS4-OCTENE'): 0.00760889053344727, ('13-BUTADIENE', 'CO2'): 0.101000003516674,
    ('13-BUTADIENE', 'CS2'): 0.0341714471578598, ('13-BUTADIENE', 'CYCLOBUTANE'): 3.52859497070313e-05,
    ('13-BUTADIENE', 'CYCLOBUTENE'): 0.000208735466003418,
    ('13-BUTADIENE', 'CYCLOHEXANE'): 0.00152814388275146,
    ('13-BUTADIENE', 'CYCLOPENTANE'): 0.000366508960723877,
    ('13-BUTADIENE', 'CYCLOPENTENE'): 0.000147640705108643,
    ('13-BUTADIENE', 'CYCLOPROPANE'): 0.000926077365875244,
    ('13-BUTADIENE', 'E-BENZENE'): 0.00383174419403076, ('13-BUTADIENE', 'E-MERCAPTAN'): 0.0589505396783352,
    ('13-BUTADIENE', 'EACETYLENE'): 2.98023223876953e-07,
    ('13-BUTADIENE', 'ECYCLOPENTAN'): 0.00384968519210815, ('13-BUTADIENE', 'ETHANE'): 0.00222861766815186,
    ('13-BUTADIENE', 'ETHYLENE'): 0.00401872396469116, ('13-BUTADIENE', 'H2'): 0.292100012302399,
    ('13-BUTADIENE', 'H2O'): 0.5, ('13-BUTADIENE', 'H2S'): 0.0450000017881393,
    ('13-BUTADIENE', 'I-BUTENE'): 8.45789909362793e-05, ('13-BUTADIENE', 'IC4'): 0.000420331954956055,
    ('13-BUTADIENE', 'IC5'): 0.00152814388275146, ('13-BUTADIENE', 'M-ACETYLENE'): 0.00123465061187744,
    ('13-BUTADIENE', 'M-MERCAPTAN'): 0.0589505396783352, ('13-BUTADIENE', 'M-XYLENE'): 0.00389909744262695,
    ('13-BUTADIENE', 'MCYCLOHEXANE'): 0.00360053777694702,
    ('13-BUTADIENE', 'MCYCLOPENTAN'): 0.00186800956726074, ('13-BUTADIENE', 'METHANE'): 0.00889021158218384,
    ('13-BUTADIENE', 'N-PBENZENE'): 0.00654953718185425, ('13-BUTADIENE', 'NC4'): 0.0141000002622604,
    ('13-BUTADIENE', 'NC5'): 0.00161850452423096, ('13-BUTADIENE', 'NC6'): 0.00360053777694702,
    ('13-BUTADIENE', 'NC7'): 0.0059514045715332, ('13-BUTADIENE', 'NC8'): 0.00856369733810425,
    ('13-BUTADIENE', 'NITROGEN'): 0.100000001490116, ('13-BUTADIENE', 'O-XYLENE'): 0.0036434531211853,
    ('13-BUTADIENE', 'P-XYLENE'): 0.00402939319610596, ('13-BUTADIENE', 'PROPADIENE'): 0.00133812427520752,
    ('13-BUTADIENE', 'PROPANE'): 0.000138401985168457, ('13-BUTADIENE', 'PROPENE'): 0.000554859638214111,
    ('13-BUTADIENE', 'STYRENE'): 0.00300157070159912, ('13-BUTADIENE', 'TOLUENE'): 0.00177323818206787,
    ('13-BUTADIENE', 'TR2-BUTENE'): 7.76052474975586e-05, ('13-BUTADIENE', 'TR2-HEXENE'): 0.00296193361282349,
    ('13-BUTADIENE', 'TR2-OCTENE'): 0.00760889053344727, ('13-BUTADIENE', 'TR2-PENTENE'): 0.00135445594787598,
    ('13-BUTADIENE', 'TR3-HEXENE'): 0.00296193361282349, ('13-BUTADIENE', 'TR3-OCTENE'): 0.00760889053344727,
    ('13-BUTADIENE', 'TR4-OCTENE'): 0.00760889053344727,
    ('13-BUTADIENE', 'VNYLACETLEN'): 2.68220901489258e-06,
    ('14PENTADIENE', '15-HEXADIENE'): 0.000486671924591064,
    ('14PENTADIENE', '2-PENTYNE'): 1.78813934326172e-07,
    ('14PENTADIENE', '23PENTADIENE'): 1.78813934326172e-07,
    ('14PENTADIENE', '26-OCTADIENE'): 0.00404816865921021, ('14PENTADIENE', '2E1C4='): 0.000841140747070313,
    ('14PENTADIENE', '2M1C5='): 0.000841140747070313, ('14PENTADIENE', '3-M-1-BUTYNE'): 4.29153442382813e-06,
    ('14PENTADIENE', '3M1C5='): 0.000710129737854004, ('14PENTADIENE', 'ACETYLENE'): 0.0110151171684265,
    ('14PENTADIENE', 'BENZENE'): 5.17368316650391e-05, ('14PENTADIENE', 'BIACETYLENE'): 0.00232440233230591,
    ('14PENTADIENE', 'CIS2-BUTENE'): 0.000385582447052002,
    ('14PENTADIENE', 'CIS2-HEXENE'): 0.000791668891906738,
    ('14PENTADIENE', 'CIS2-OCTENE'): 0.00372916460037231,
    ('14PENTADIENE', 'CIS2-PENTENE'): 0.00010979175567627,
    ('14PENTADIENE', 'CIS3-HEXENE'): 0.000791668891906738,
    ('14PENTADIENE', 'CIS4-OCTENE'): 0.00372916460037231, ('14PENTADIENE', 'CO2'): 0.101000003516674,
    ('14PENTADIENE', 'CS2'): 0.0375353135168552, ('14PENTADIENE', 'CYCLOBUTANE'): 0.0010417103767395,
    ('14PENTADIENE', 'CYCLOBUTENE'): 0.00166171789169312,
    ('14PENTADIENE', 'CYCLOHEXANE'): 0.000163018703460693,
    ('14PENTADIENE', 'CYCLOPENTANE'): 5.17368316650391e-05,
    ('14PENTADIENE', 'CYCLOPENTENE'): 0.00020146369934082,
    ('14PENTADIENE', 'CYCLOPROPANE'): 0.00321614742279053,
    ('14PENTADIENE', 'DMACETYLENE'): 0.000693619251251221, ('14PENTADIENE', 'E-BENZENE'): 0.00126999616622925,
    ('14PENTADIENE', 'E-MERCAPTAN'): 0.0576850026845932, ('14PENTADIENE', 'EACETYLENE'): 0.000665903091430664,
    ('14PENTADIENE', 'ECYCLOPENTAN'): 0.00128030776977539, ('14PENTADIENE', 'ETHANE'): 0.00539225339889526,
    ('14PENTADIENE', 'ETHYLENE'): 0.0080181360244751, ('14PENTADIENE', 'H2'): 0.292100012302399,
    ('14PENTADIENE', 'H2O'): 0.5, ('14PENTADIENE', 'H2S'): 0.0450000017881393,
    ('14PENTADIENE', 'I-BUTENE'): 0.000293910503387451, ('14PENTADIENE', 'IC4'): 3.4034252166748e-05,
    ('14PENTADIENE', 'IC5'): 0.000163018703460693, ('14PENTADIENE', 'M-ACETYLENE'): 0.00377029180526733,
    ('14PENTADIENE', 'M-MERCAPTAN'): 0.0576850026845932, ('14PENTADIENE', 'M-XYLENE'): 0.00130897760391235,
    ('14PENTADIENE', 'MCYCLOHEXANE'): 0.00113797187805176,
    ('14PENTADIENE', 'MCYCLOPENTAN'): 0.000285923480987549, ('14PENTADIENE', 'METHANE'): 0.0144607424736023,
    ('14PENTADIENE', 'N-PBENZENE'): 0.00299620628356934, ('14PENTADIENE', 'NC4'): 8.99434089660645e-05,
    ('14PENTADIENE', 'NC5'): 0.000193595886230469, ('14PENTADIENE', 'NC6'): 0.00113797187805176,
    ('14PENTADIENE', 'NC7'): 0.00259464979171753, ('14PENTADIENE', 'NC8'): 0.00440961122512817,
    ('14PENTADIENE', 'NITROGEN'): 0.100000001490116, ('14PENTADIENE', 'O-XYLENE'): 0.00116223096847534,
    ('14PENTADIENE', 'P-XYLENE'): 0.00138527154922485, ('14PENTADIENE', 'PROPADIENE'): 0.00394898653030396,
    ('14PENTADIENE', 'PROPANE'): 0.00145077705383301, ('14PENTADIENE', 'PROPENE'): 0.00248563289642334,
    ('14PENTADIENE', 'STYRENE'): 0.000812351703643799, ('14PENTADIENE', 'TOLUENE'): 0.000249505043029785,
    ('14PENTADIENE', 'TR2-BUTENE'): 0.00030750036239624, ('14PENTADIENE', 'TR2-HEXENE'): 0.000791668891906738,
    ('14PENTADIENE', 'TR2-OCTENE'): 0.00372916460037231, ('14PENTADIENE', 'TR2-PENTENE'): 0.00010979175567627,
    ('14PENTADIENE', 'TR3-HEXENE'): 0.000791668891906738, ('14PENTADIENE', 'TR3-OCTENE'): 0.00372916460037231,
    ('14PENTADIENE', 'TR4-OCTENE'): 0.00372916460037231,
    ('14PENTADIENE', 'VNYLACETLEN'): 0.000782191753387451,
    ('15-HEXADIENE', '2-PENTYNE'): 0.000465571880340576,
    ('15-HEXADIENE', '23PENTADIENE'): 0.000503063201904297,
    ('15-HEXADIENE', '26-OCTADIENE'): 0.00173342227935791, ('15-HEXADIENE', '2E1C4='): 4.82797622680664e-05,
    ('15-HEXADIENE', '2M1C5='): 4.82797622680664e-05, ('15-HEXADIENE', '3-M-1-BUTYNE'): 0.000583291053771973,
    ('15-HEXADIENE', '3M1C5='): 2.11000442504883e-05, ('15-HEXADIENE', 'ACETYLENE'): 0.0160402655601501,
    ('15-HEXADIENE', 'BENZENE'): 0.000855803489685059, ('15-HEXADIENE', 'BIACETYLENE'): 0.00492548942565918,
    ('15-HEXADIENE', 'CIS2-BUTENE'): 0.00173693895339966,
    ('15-HEXADIENE', 'CIS2-HEXENE'): 3.68952751159668e-05,
    ('15-HEXADIENE', 'CIS2-OCTENE'): 0.00152647495269775,
    ('15-HEXADIENE', 'CIS2-PENTENE'): 0.00013422966003418,
    ('15-HEXADIENE', 'CIS3-HEXENE'): 3.68952751159668e-05,
    ('15-HEXADIENE', 'CIS4-OCTENE'): 0.00152647495269775, ('15-HEXADIENE', 'CO2'): 0.101000003516674,
    ('15-HEXADIENE', 'CS2'): 0.0409149825572968, ('15-HEXADIENE', 'CYCLOBUTANE'): 0.00294727087020874,
    ('15-HEXADIENE', 'CYCLOBUTENE'): 0.00393849611282349,
    ('15-HEXADIENE', 'CYCLOHEXANE'): 8.64267349243164e-05,
    ('15-HEXADIENE', 'CYCLOPENTANE'): 0.000855803489685059,
    ('15-HEXADIENE', 'CYCLOPENTENE'): 0.00131350755691528,
    ('15-HEXADIENE', 'CYCLOPROPANE'): 0.00618594884872437,
    ('15-HEXADIENE', 'DMACETYLENE'): 0.00233918428421021, ('15-HEXADIENE', 'E-BENZENE'): 0.00018465518951416,
    ('15-HEXADIENE', 'E-MERCAPTAN'): 0.0560605749487877, ('15-HEXADIENE', 'EACETYLENE'): 0.00228816270828247,
    ('15-HEXADIENE', 'ECYCLOPENTAN'): 0.000188589096069336, ('15-HEXADIENE', 'ETHANE'): 0.00908243656158447,
    ('15-HEXADIENE', 'ETHYLENE'): 0.012394905090332, ('15-HEXADIENE', 'H2'): 0.292100012302399,
    ('15-HEXADIENE', 'H2O'): 0.5, ('15-HEXADIENE', 'H2S'): 0.0450000017881393,
    ('15-HEXADIENE', 'I-BUTENE'): 0.00153583288192749, ('15-HEXADIENE', 'IC4'): 0.000778257846832275,
    ('15-HEXADIENE', 'IC5'): 8.64267349243164e-05, ('15-HEXADIENE', 'M-ACETYLENE'): 0.0069429874420166,
    ('15-HEXADIENE', 'M-MERCAPTAN'): 0.0560605749487877, ('15-HEXADIENE', 'M-XYLENE'): 0.000199735164642334,
    ('15-HEXADIENE', 'MCYCLOHEXANE'): 0.000136435031890869,
    ('15-HEXADIENE', 'MCYCLOPENTAN'): 2.65836715698242e-05, ('15-HEXADIENE', 'METHANE'): 0.0201205611228943,
    ('15-HEXADIENE', 'N-PBENZENE'): 0.00107091665267944, ('15-HEXADIENE', 'NC4'): 0.000994741916656494,
    ('15-HEXADIENE', 'NC5'): 6.64591789245605e-05, ('15-HEXADIENE', 'NC6'): 0.000136435031890869,
    ('15-HEXADIENE', 'NC7'): 0.000836074352264404, ('15-HEXADIENE', 'NC8'): 0.00197345018386841,
    ('15-HEXADIENE', 'NITROGEN'): 0.100000001490116, ('15-HEXADIENE', 'O-XYLENE'): 0.000145018100738525,
    ('15-HEXADIENE', 'P-XYLENE'): 0.000230073928833008, ('15-HEXADIENE', 'PROPADIENE'): 0.00718384981155396,
    ('15-HEXADIENE', 'PROPANE'): 0.00361055135726929, ('15-HEXADIENE', 'PROPENE'): 0.0051581859588623,
    ('15-HEXADIENE', 'STYRENE'): 4.15444374084473e-05, ('15-HEXADIENE', 'TOLUENE'): 3.92794609069824e-05,
    ('15-HEXADIENE', 'TR2-BUTENE'): 0.00156641006469727, ('15-HEXADIENE', 'TR2-HEXENE'): 3.68952751159668e-05,
    ('15-HEXADIENE', 'TR2-OCTENE'): 0.00152647495269775, ('15-HEXADIENE', 'TR2-PENTENE'): 0.00013422966003418,
    ('15-HEXADIENE', 'TR3-HEXENE'): 3.68952751159668e-05, ('15-HEXADIENE', 'TR3-OCTENE'): 0.00152647495269775,
    ('15-HEXADIENE', 'TR4-OCTENE'): 0.00152647495269775, ('15-HEXADIENE', 'VNYLACETLEN'): 0.00249916315078735,
    ('2-PENTYNE', '23PENTADIENE'): 7.15255737304688e-07, ('2-PENTYNE', '26-OCTADIENE'): 0.00398695468902588,
    ('2-PENTYNE', '2E1C4='): 0.000813305377960205, ('2-PENTYNE', '2M1C5='): 0.000813305377960205,
    ('2-PENTYNE', '3-M-1-BUTYNE'): 6.61611557006836e-06, ('2-PENTYNE', '3M1C5='): 0.000684559345245361,
    ('2-PENTYNE', 'ACETYLENE'): 0.0111157894134521, ('2-PENTYNE', 'BENZENE'): 5.90085983276367e-05,
    ('2-PENTYNE', 'BIACETYLENE'): 0.00237125158309937, ('2-PENTYNE', 'CIS2-BUTENE'): 0.000404894351959229,
    ('2-PENTYNE', 'CIS2-HEXENE'): 0.000764548778533936, ('2-PENTYNE', 'CIS2-OCTENE'): 0.00367045402526855,
    ('2-PENTYNE', 'CIS2-PENTENE'): 9.98377799987793e-05, ('2-PENTYNE', 'CIS3-HEXENE'): 0.000764548778533936,
    ('2-PENTYNE', 'CIS4-OCTENE'): 0.00367045402526855, ('2-PENTYNE', 'CO2'): 0.101000003516674,
    ('2-PENTYNE', 'CS2'): 0.0410122312605381, ('2-PENTYNE', 'CYCLOBUTANE'): 0.00107306241989136,
    ('2-PENTYNE', 'CYCLOBUTENE'): 0.00170153379440308, ('2-PENTYNE', 'CYCLOHEXANE'): 0.000150859355926514,
    ('2-PENTYNE', 'CYCLOPENTANE'): 5.90085983276367e-05, ('2-PENTYNE', 'CYCLOPENTENE'): 0.000215470790863037,
    ('2-PENTYNE', 'CYCLOPROPANE'): 0.00327116250991821, ('2-PENTYNE', 'DMACETYLENE'): 0.000719428062438965,
    ('2-PENTYNE', 'E-BENZENE'): 0.0012357234954834, ('2-PENTYNE', 'E-MERCAPTAN'): 0.0545828454196453,
    ('2-PENTYNE', 'EACETYLENE'): 0.000691235065460205, ('2-PENTYNE', 'ECYCLOPENTAN'): 0.00124591588973999,
    ('2-PENTYNE', 'ETHANE'): 0.00546324253082275, ('2-PENTYNE', 'ETHYLENE'): 0.00810432434082031,
    ('2-PENTYNE', 'H2'): 0.292100012302399, ('2-PENTYNE', 'H2O'): 0.5,
    ('2-PENTYNE', 'H2S'): 0.0450000017881393, ('2-PENTYNE', 'I-BUTENE'): 0.000310838222503662,
    ('2-PENTYNE', 'IC4'): 3.99947166442871e-05, ('2-PENTYNE', 'IC5'): 0.000150859355926514,
    ('2-PENTYNE', 'M-ACETYLENE'): 0.00382977724075317, ('2-PENTYNE', 'M-MERCAPTAN'): 0.0545828454196453,
    ('2-PENTYNE', 'M-XYLENE'): 0.0012742280960083, ('2-PENTYNE', 'MCYCLOHEXANE'): 0.00110548734664917,
    ('2-PENTYNE', 'MCYCLOPENTAN'): 0.000269770622253418, ('2-PENTYNE', 'METHANE'): 0.0145755410194397,
    ('2-PENTYNE', 'N-PBENZENE'): 0.00294357538223267, ('2-PENTYNE', 'NC4'): 9.94205474853516e-05,
    ('2-PENTYNE', 'NC5'): 0.000180304050445557, ('2-PENTYNE', 'NC6'): 0.00110548734664917,
    ('2-PENTYNE', 'NC7'): 0.00254565477371216, ('2-PENTYNE', 'NC8'): 0.00434577465057373,
    ('2-PENTYNE', 'NITROGEN'): 0.100000001490116, ('2-PENTYNE', 'O-XYLENE'): 0.00112944841384888,
    ('2-PENTYNE', 'P-XYLENE'): 0.00134932994842529, ('2-PENTYNE', 'PROPADIENE'): 0.00400996208190918,
    ('2-PENTYNE', 'PROPANE'): 0.00148779153823853, ('2-PENTYNE', 'PROPENE'): 0.00253409147262573,
    ('2-PENTYNE', 'STYRENE'): 0.000784933567047119, ('2-PENTYNE', 'TOLUENE'): 0.000234305858612061,
    ('2-PENTYNE', 'TR2-BUTENE'): 0.000324606895446777, ('2-PENTYNE', 'TR2-HEXENE'): 0.000764548778533936,
    ('2-PENTYNE', 'TR2-OCTENE'): 0.00367045402526855, ('2-PENTYNE', 'TR2-PENTENE'): 9.98377799987793e-05,
    ('2-PENTYNE', 'TR3-HEXENE'): 0.000764548778533936, ('2-PENTYNE', 'TR3-OCTENE'): 0.00367045402526855,
    ('2-PENTYNE', 'TR4-OCTENE'): 0.00367045402526855, ('2-PENTYNE', 'VNYLACETLEN'): 0.00080949068069458,
    ('23PENTADIENE', '26-OCTADIENE'): 0.00409489870071411, ('23PENTADIENE', '2E1C4='): 0.000862598419189453,
    ('23PENTADIENE', '2M1C5='): 0.000862598419189453, ('23PENTADIENE', '3-M-1-BUTYNE'): 2.92062759399414e-06,
    ('23PENTADIENE', '3M1C5='): 0.000729858875274658, ('23PENTADIENE', 'ACETYLENE'): 0.0109390616416931,
    ('23PENTADIENE', 'BENZENE'): 4.66108322143555e-05, ('23PENTADIENE', 'BIACETYLENE'): 0.00228911638259888,
    ('23PENTADIENE', 'CIS2-BUTENE'): 0.000371336936950684,
    ('23PENTADIENE', 'CIS2-HEXENE'): 0.000812530517578125,
    ('23PENTADIENE', 'CIS2-OCTENE'): 0.00377404689788818,
    ('23PENTADIENE', 'CIS2-PENTENE'): 0.000117719173431396,
    ('23PENTADIENE', 'CIS3-HEXENE'): 0.000812530517578125,
    ('23PENTADIENE', 'CIS4-OCTENE'): 0.00377404689788818, ('23PENTADIENE', 'CO2'): 0.101000003516674,
    ('23PENTADIENE', 'CS2'): 0.0400799140334129, ('23PENTADIENE', 'CYCLOBUTANE'): 0.00101804733276367,
    ('23PENTADIENE', 'CYCLOBUTENE'): 0.00163191556930542,
    ('23PENTADIENE', 'CYCLOHEXANE'): 0.00017249584197998,
    ('23PENTADIENE', 'CYCLOPENTANE'): 4.66108322143555e-05,
    ('23PENTADIENE', 'CYCLOPENTENE'): 0.000191092491149902,
    ('23PENTADIENE', 'CYCLOPROPANE'): 0.00317478179931641,
    ('23PENTADIENE', 'DMACETYLENE'): 0.000674426555633545, ('23PENTADIENE', 'E-BENZENE'): 0.00129634141921997,
    ('23PENTADIENE', 'E-MERCAPTAN'): 0.0557683110237122, ('23PENTADIENE', 'EACETYLENE'): 0.000647127628326416,
    ('23PENTADIENE', 'ECYCLOPENTAN'): 0.00130677223205566, ('23PENTADIENE', 'ETHANE'): 0.00533866882324219,
    ('23PENTADIENE', 'ETHYLENE'): 0.00795292854309082, ('23PENTADIENE', 'H2'): 0.292100012302399,
    ('23PENTADIENE', 'H2O'): 0.5, ('23PENTADIENE', 'H2S'): 0.0450000017881393,
    ('23PENTADIENE', 'I-BUTENE'): 0.000281393527984619, ('23PENTADIENE', 'IC4'): 2.99215316772461e-05,
    ('23PENTADIENE', 'IC5'): 0.00017249584197998, ('23PENTADIENE', 'M-ACETYLENE'): 0.00372534990310669,
    ('23PENTADIENE', 'M-MERCAPTAN'): 0.0557683110237122, ('23PENTADIENE', 'M-XYLENE'): 0.00133579969406128,
    ('23PENTADIENE', 'MCYCLOHEXANE'): 0.00116294622421265,
    ('23PENTADIENE', 'MCYCLOPENTAN'): 0.000298500061035156, ('23PENTADIENE', 'METHANE'): 0.0143739581108093,
    ('23PENTADIENE', 'N-PBENZENE'): 0.0030364990234375, ('23PENTADIENE', 'NC4'): 8.31484794616699e-05,
    ('23PENTADIENE', 'NC5'): 0.000203907489776611, ('23PENTADIENE', 'NC6'): 0.00116294622421265,
    ('23PENTADIENE', 'NC7'): 0.00263214111328125, ('23PENTADIENE', 'NC8'): 0.00445836782455444,
    ('23PENTADIENE', 'NITROGEN'): 0.100000001490116, ('23PENTADIENE', 'O-XYLENE'): 0.00118744373321533,
    ('23PENTADIENE', 'P-XYLENE'): 0.00141274929046631, ('23PENTADIENE', 'PROPADIENE'): 0.00390321016311646,
    ('23PENTADIENE', 'PROPANE'): 0.00142288208007813, ('23PENTADIENE', 'PROPENE'): 0.0024491548538208,
    ('23PENTADIENE', 'STYRENE'): 0.000833451747894287, ('23PENTADIENE', 'TOLUENE'): 0.000261247158050537,
    ('23PENTADIENE', 'TR2-BUTENE'): 0.000294685363769531,
    ('23PENTADIENE', 'TR2-HEXENE'): 0.000812530517578125, ('23PENTADIENE', 'TR2-OCTENE'): 0.00377404689788818,
    ('23PENTADIENE', 'TR2-PENTENE'): 0.000117719173431396,
    ('23PENTADIENE', 'TR3-HEXENE'): 0.000812530517578125, ('23PENTADIENE', 'TR3-OCTENE'): 0.00377404689788818,
    ('23PENTADIENE', 'TR4-OCTENE'): 0.00377404689788818,
    ('23PENTADIENE', 'VNYLACETLEN'): 0.000761806964874268,
    ('26-OCTADIENE', '2E1C4='): 0.00120425224304199, ('26-OCTADIENE', '2M1C5='): 0.00120425224304199,
    ('26-OCTADIENE', '3-M-1-BUTYNE'): 0.00431716442108154, ('26-OCTADIENE', '3M1C5='): 0.00137317180633545,
    ('26-OCTADIENE', 'ACETYLENE'): 0.027951717376709, ('26-OCTADIENE', 'BENZENE'): 0.00500982999801636,
    ('26-OCTADIENE', 'BIACETYLENE'): 0.0124122500419617, ('26-OCTADIENE', 'CIS2-BUTENE'): 0.00691074132919312,
    ('26-OCTADIENE', 'CIS2-HEXENE'): 0.00126516819000244,
    ('26-OCTADIENE', 'CIS2-OCTENE'): 6.61611557006836e-06,
    ('26-OCTADIENE', 'CIS2-PENTENE'): 0.00282925367355347,
    ('26-OCTADIENE', 'CIS3-HEXENE'): 0.00126516819000244,
    ('26-OCTADIENE', 'CIS4-OCTENE'): 6.61611557006836e-06, ('26-OCTADIENE', 'CO2'): 0.101000003516674,
    ('26-OCTADIENE', 'CS2'): 0.0461949445307255, ('26-OCTADIENE', 'CYCLOBUTANE'): 0.00914978981018066,
    ('26-OCTADIENE', 'CYCLOBUTENE'): 0.0108270049095154, ('26-OCTADIENE', 'CYCLOHEXANE'): 0.00259149074554443,
    ('26-OCTADIENE', 'CYCLOPENTANE'): 0.00500982999801636,
    ('26-OCTADIENE', 'CYCLOPENTENE'): 0.0060420036315918,
    ('26-OCTADIENE', 'CYCLOPROPANE'): 0.0143513083457947,
    ('26-OCTADIENE', 'DMACETYLENE'): 0.00805944204330444, ('26-OCTADIENE', 'E-BENZENE'): 0.000787794589996338,
    ('26-OCTADIENE', 'E-MERCAPTAN'): 0.0497129671275616, ('26-OCTADIENE', 'EACETYLENE'): 0.007965087890625,
    ('26-OCTADIENE', 'ECYCLOPENTAN'): 0.000779688358306885, ('26-OCTADIENE', 'ETHANE'): 0.0185675621032715,
    ('26-OCTADIENE', 'ETHYLENE'): 0.0231311917304993, ('26-OCTADIENE', 'H2'): 0.292100012302399,
    ('26-OCTADIENE', 'H2O'): 0.5, ('26-OCTADIENE', 'H2S'): 0.0450000017881393,
    ('26-OCTADIENE', 'I-BUTENE'): 0.00650596618652344, ('26-OCTADIENE', 'IC4'): 0.00482082366943359,
    ('26-OCTADIENE', 'IC5'): 0.00259149074554443, ('26-OCTADIENE', 'M-ACETYLENE'): 0.0154805779457092,
    ('26-OCTADIENE', 'M-MERCAPTAN'): 0.0497129671275616, ('26-OCTADIENE', 'M-XYLENE'): 0.000757575035095215,
    ('26-OCTADIENE', 'MCYCLOHEXANE'): 0.000898301601409912,
    ('26-OCTADIENE', 'MCYCLOPENTAN'): 0.00218832492828369, ('26-OCTADIENE', 'METHANE'): 0.0331763029098511,
    ('26-OCTADIENE', 'N-PBENZENE'): 7.96914100646973e-05, ('26-OCTADIENE', 'NC4'): 0.0053369402885437,
    ('26-OCTADIENE', 'NC5'): 0.00247687101364136, ('26-OCTADIENE', 'NC6'): 0.000898301601409912,
    ('26-OCTADIENE', 'NC7'): 0.000162243843078613, ('26-OCTADIENE', 'NC8'): 7.74860382080078e-06,
    ('26-OCTADIENE', 'NITROGEN'): 0.100000001490116, ('26-OCTADIENE', 'O-XYLENE'): 0.000876903533935547,
    ('26-OCTADIENE', 'P-XYLENE'): 0.000701487064361572, ('26-OCTADIENE', 'PROPADIENE'): 0.0158354640007019,
    ('26-OCTADIENE', 'PROPANE'): 0.010283350944519, ('26-OCTADIENE', 'PROPENE'): 0.0127766728401184,
    ('26-OCTADIENE', 'STYRENE'): 0.00123924016952515, ('26-OCTADIENE', 'TOLUENE'): 0.00229352712631226,
    ('26-OCTADIENE', 'TR2-BUTENE'): 0.00656861066818237, ('26-OCTADIENE', 'TR2-HEXENE'): 0.00126516819000244,
    ('26-OCTADIENE', 'TR2-OCTENE'): 6.61611557006836e-06,
    ('26-OCTADIENE', 'TR2-PENTENE'): 0.00282925367355347, ('26-OCTADIENE', 'TR3-HEXENE'): 0.00126516819000244,
    ('26-OCTADIENE', 'TR3-OCTENE'): 6.61611557006836e-06,
    ('26-OCTADIENE', 'TR4-OCTENE'): 6.61611557006836e-06,
    ('26-OCTADIENE', 'VNYLACETLEN'): 0.00835216045379639,
    ('2E1C4=', '3-M-1-BUTYNE'): 0.000966727733612061, ('2E1C4=', '3M1C5='): 5.60283660888672e-06,
    ('2E1C4=', 'ACETYLENE'): 0.0178086161613464, ('2E1C4=', 'BENZENE'): 0.00130957365036011,
    ('2E1C4=', 'BIACETYLENE'): 0.00594127178192139, ('2E1C4=', 'CIS2-BUTENE'): 0.00236231088638306,
    ('2E1C4=', 'CIS2-HEXENE'): 7.74860382080078e-07, ('2E1C4=', 'CIS2-OCTENE'): 0.00103271007537842,
    ('2E1C4=', 'CIS2-PENTENE'): 0.00034332275390625, ('2E1C4=', 'CIS3-HEXENE'): 7.74860382080078e-07,
    ('2E1C4=', 'CIS4-OCTENE'): 0.00103271007537842, ('2E1C4=', 'CO2'): 0.101000003516674,
    ('2E1C4=', 'CS2'): 0.0412995591759682, ('2E1C4=', 'CYCLOBUTANE'): 0.00374609231948853,
    ('2E1C4=', 'CYCLOBUTENE'): 0.00485318899154663, ('2E1C4=', 'CYCLOHEXANE'): 0.000263810157775879,
    ('2E1C4=', 'CYCLOPENTANE'): 0.00130957365036011, ('2E1C4=', 'CYCLOPENTENE'): 0.00186395645141602,
    ('2E1C4=', 'CYCLOPROPANE'): 0.00731676816940308, ('2E1C4=', 'DMACETYLENE'): 0.00305658578872681,
    ('2E1C4=', 'E-BENZENE'): 4.41074371337891e-05, ('2E1C4=', 'E-MERCAPTAN'): 0.0561534017324448,
    ('2E1C4=', 'EACETYLENE'): 0.00299829244613647, ('2E1C4=', 'ECYCLOPENTAN'): 4.6074390411377e-05,
    ('2E1C4=', 'ETHANE'): 0.0104371309280396, ('2E1C4=', 'ETHYLENE'): 0.0139625668525696,
    ('2E1C4=', 'H2'): 0.292100012302399, ('2E1C4=', 'H2O'): 0.5, ('2E1C4=', 'H2S'): 0.0450000017881393,
    ('2E1C4=', 'I-BUTENE'): 0.00212681293487549, ('2E1C4=', 'IC4'): 0.0012134313583374,
    ('2E1C4=', 'IC5'): 0.000263810157775879, ('2E1C4=', 'M-ACETYLENE'): 0.0081368088722229,
    ('2E1C4=', 'M-MERCAPTAN'): 0.0561534017324448, ('2E1C4=', 'M-XYLENE'): 5.16772270202637e-05,
    ('2E1C4=', 'MCYCLOHEXANE'): 2.24113464355469e-05, ('2E1C4=', 'MCYCLOPENTAN'): 0.000146448612213135,
    ('2E1C4=', 'METHANE'): 0.0220847725868225, ('2E1C4=', 'N-PBENZENE'): 0.000664889812469482,
    ('2E1C4=', 'NC4'): 0.00148016214370728, ('2E1C4=', 'NC5'): 0.000227987766265869,
    ('2E1C4=', 'NC6'): 2.24113464355469e-05, ('2E1C4=', 'NC7'): 0.000482857227325439,
    ('2E1C4=', 'NC8'): 0.00140571594238281, ('2E1C4=', 'NITROGEN'): 0.100000001490116,
    ('2E1C4=', 'O-XYLENE'): 2.59280204772949e-05, ('2E1C4=', 'P-XYLENE'): 6.76512718200684e-05,
    ('2E1C4=', 'PROPADIENE'): 0.00839704275131226, ('2E1C4=', 'PROPANE'): 0.00448876619338989,
    ('2E1C4=', 'PROPENE'): 0.00619620084762573, ('2E1C4=', 'STYRENE'): 2.38418579101563e-07,
    ('2E1C4=', 'TOLUENE'): 0.000174641609191895, ('2E1C4=', 'TR2-BUTENE'): 0.00216299295425415,
    ('2E1C4=', 'TR2-HEXENE'): 7.74860382080078e-07, ('2E1C4=', 'TR2-OCTENE'): 0.00103271007537842,
    ('2E1C4=', 'TR2-PENTENE'): 0.00034332275390625, ('2E1C4=', 'TR3-HEXENE'): 7.74860382080078e-07,
    ('2E1C4=', 'TR3-OCTENE'): 0.00103271007537842, ('2E1C4=', 'TR4-OCTENE'): 0.00103271007537842,
    ('2E1C4=', 'VNYLACETLEN'): 0.00323909521102905,
    ('2M1C5=', '3-M-1-BUTYNE'): 0.000966727733612061, ('2M1C5=', '3M1C5='): 5.60283660888672e-06,
    ('2M1C5=', 'ACETYLENE'): 0.0178086161613464, ('2M1C5=', 'BENZENE'): 0.00130957365036011,
    ('2M1C5=', 'BIACETYLENE'): 0.00594127178192139, ('2M1C5=', 'CIS2-BUTENE'): 0.00236231088638306,
    ('2M1C5=', 'CIS2-HEXENE'): 7.74860382080078e-07, ('2M1C5=', 'CIS2-OCTENE'): 0.00103271007537842,
    ('2M1C5=', 'CIS2-PENTENE'): 0.00034332275390625, ('2M1C5=', 'CIS3-HEXENE'): 7.74860382080078e-07,
    ('2M1C5=', 'CIS4-OCTENE'): 0.00103271007537842, ('2M1C5=', 'CO2'): 0.101000003516674,
    ('2M1C5=', 'CS2'): 0.040916882455349, ('2M1C5=', 'CYCLOBUTANE'): 0.00374609231948853,
    ('2M1C5=', 'CYCLOBUTENE'): 0.00485318899154663, ('2M1C5=', 'CYCLOHEXANE'): 0.000263810157775879,
    ('2M1C5=', 'CYCLOPENTANE'): 0.00130957365036011, ('2M1C5=', 'CYCLOPENTENE'): 0.00186395645141602,
    ('2M1C5=', 'CYCLOPROPANE'): 0.00731676816940308, ('2M1C5=', 'DMACETYLENE'): 0.00305658578872681,
    ('2M1C5=', 'E-BENZENE'): 4.41074371337891e-05, ('2M1C5=', 'E-MERCAPTAN'): 0.0567792877554893,
    ('2M1C5=', 'EACETYLENE'): 0.00299829244613647, ('2M1C5=', 'ECYCLOPENTAN'): 4.6074390411377e-05,
    ('2M1C5=', 'ETHANE'): 0.0104371309280396, ('2M1C5=', 'ETHYLENE'): 0.0139625668525696,
    ('2M1C5=', 'H2'): 0.292100012302399, ('2M1C5=', 'H2O'): 0.5, ('2M1C5=', 'H2S'): 0.0450000017881393,
    ('2M1C5=', 'I-BUTENE'): 0.00212681293487549, ('2M1C5=', 'IC4'): 0.0012134313583374,
    ('2M1C5=', 'IC5'): 0.000263810157775879, ('2M1C5=', 'M-ACETYLENE'): 0.0081368088722229,
    ('2M1C5=', 'M-MERCAPTAN'): 0.0567792877554893, ('2M1C5=', 'M-XYLENE'): 5.16772270202637e-05,
    ('2M1C5=', 'MCYCLOHEXANE'): 2.24113464355469e-05, ('2M1C5=', 'MCYCLOPENTAN'): 0.000146448612213135,
    ('2M1C5=', 'METHANE'): 0.0220847725868225, ('2M1C5=', 'N-PBENZENE'): 0.000664889812469482,
    ('2M1C5=', 'NC4'): 0.00148016214370728, ('2M1C5=', 'NC5'): 0.000227987766265869,
    ('2M1C5=', 'NC6'): 2.24113464355469e-05, ('2M1C5=', 'NC7'): 0.000482857227325439,
    ('2M1C5=', 'NC8'): 0.00140571594238281, ('2M1C5=', 'NITROGEN'): 0.100000001490116,
    ('2M1C5=', 'O-XYLENE'): 2.59280204772949e-05, ('2M1C5=', 'P-XYLENE'): 6.76512718200684e-05,
    ('2M1C5=', 'PROPADIENE'): 0.00839704275131226, ('2M1C5=', 'PROPANE'): 0.00448876619338989,
    ('2M1C5=', 'PROPENE'): 0.00619620084762573, ('2M1C5=', 'STYRENE'): 2.38418579101563e-07,
    ('2M1C5=', 'TOLUENE'): 0.000174641609191895, ('2M1C5=', 'TR2-BUTENE'): 0.00216299295425415,
    ('2M1C5=', 'TR2-HEXENE'): 7.74860382080078e-07, ('2M1C5=', 'TR2-OCTENE'): 0.00103271007537842,
    ('2M1C5=', 'TR2-PENTENE'): 0.00034332275390625, ('2M1C5=', 'TR3-HEXENE'): 7.74860382080078e-07,
    ('2M1C5=', 'TR3-OCTENE'): 0.00103271007537842, ('2M1C5=', 'TR4-OCTENE'): 0.00103271007537842,
    ('2M1C5=', 'VNYLACETLEN'): 0.00323909521102905,
    ('3-M-1-BUTYNE', '3M1C5='): 0.000825881958007813, ('3-M-1-BUTYNE', 'ACETYLENE'): 0.0105862617492676,
    ('3-M-1-BUTYNE', 'BENZENE'): 2.60472297668457e-05, ('3-M-1-BUTYNE', 'BIACETYLENE'): 0.00212764739990234,
    ('3-M-1-BUTYNE', 'CIS2-BUTENE'): 0.000307917594909668,
    ('3-M-1-BUTYNE', 'CIS2-HEXENE'): 0.000913560390472412,
    ('3-M-1-BUTYNE', 'CIS2-OCTENE'): 0.00398772954940796,
    ('3-M-1-BUTYNE', 'CIS2-PENTENE'): 0.000158071517944336,
    ('3-M-1-BUTYNE', 'CIS3-HEXENE'): 0.000913560390472412,
    ('3-M-1-BUTYNE', 'CIS4-OCTENE'): 0.00398772954940796, ('3-M-1-BUTYNE', 'CO2'): 0.101000003516674,
    ('3-M-1-BUTYNE', 'CS2'): 0.0362778306007385, ('3-M-1-BUTYNE', 'CYCLOBUTANE'): 0.000911235809326172,
    ('3-M-1-BUTYNE', 'CYCLOBUTENE'): 0.00149595737457275,
    ('3-M-1-BUTYNE', 'CYCLOHEXANE'): 0.000220775604248047,
    ('3-M-1-BUTYNE', 'CYCLOPENTANE'): 2.60472297668457e-05,
    ('3-M-1-BUTYNE', 'CYCLOPENTENE'): 0.000146448612213135,
    ('3-M-1-BUTYNE', 'CYCLOPROPANE'): 0.00298410654067993,
    ('3-M-1-BUTYNE', 'DMACETYLENE'): 0.000587999820709229, ('3-M-1-BUTYNE', 'E-BENZENE'): 0.001423180103302,
    ('3-M-1-BUTYNE', 'E-MERCAPTAN'): 0.0600791312754154, ('3-M-1-BUTYNE', 'EACETYLENE'): 0.000562429428100586,
    ('3-M-1-BUTYNE', 'ECYCLOPENTAN'): 0.0014340877532959, ('3-M-1-BUTYNE', 'ETHANE'): 0.00509136915206909,
    ('3-M-1-BUTYNE', 'ETHYLENE'): 0.00765162706375122, ('3-M-1-BUTYNE', 'H2'): 0.292100012302399,
    ('3-M-1-BUTYNE', 'H2O'): 0.5, ('3-M-1-BUTYNE', 'H2S'): 0.0450000017881393,
    ('3-M-1-BUTYNE', 'I-BUTENE'): 0.000226676464080811, ('3-M-1-BUTYNE', 'IC4'): 1.40666961669922e-05,
    ('3-M-1-BUTYNE', 'IC5'): 0.000220775604248047, ('3-M-1-BUTYNE', 'M-ACETYLENE'): 0.00351893901824951,
    ('3-M-1-BUTYNE', 'M-MERCAPTAN'): 0.0600791312754154, ('3-M-1-BUTYNE', 'M-XYLENE'): 0.00146442651748657,
    ('3-M-1-BUTYNE', 'MCYCLOHEXANE'): 0.00128322839736938,
    ('3-M-1-BUTYNE', 'MCYCLOPENTAN'): 0.000361025333404541, ('3-M-1-BUTYNE', 'METHANE'): 0.0139708518981934,
    ('3-M-1-BUTYNE', 'N-PBENZENE'): 0.00322866439819336, ('3-M-1-BUTYNE', 'NC4'): 5.47170639038086e-05,
    ('3-M-1-BUTYNE', 'NC5'): 0.000256061553955078, ('3-M-1-BUTYNE', 'NC6'): 0.00128322839736938,
    ('3-M-1-BUTYNE', 'NC7'): 0.00281131267547607, ('3-M-1-BUTYNE', 'NC8'): 0.00469017028808594,
    ('3-M-1-BUTYNE', 'NITROGEN'): 0.100000001490116, ('3-M-1-BUTYNE', 'O-XYLENE'): 0.00130897760391235,
    ('3-M-1-BUTYNE', 'P-XYLENE'): 0.0015450119972229, ('3-M-1-BUTYNE', 'PROPADIENE'): 0.00369173288345337,
    ('3-M-1-BUTYNE', 'PROPANE'): 0.00129604339599609, ('3-M-1-BUTYNE', 'PROPENE'): 0.00228208303451538,
    ('3-M-1-BUTYNE', 'STYRENE'): 0.000935792922973633, ('3-M-1-BUTYNE', 'TOLUENE'): 0.000319838523864746,
    ('3-M-1-BUTYNE', 'TR2-BUTENE'): 0.000238418579101563,
    ('3-M-1-BUTYNE', 'TR2-HEXENE'): 0.000913560390472412, ('3-M-1-BUTYNE', 'TR2-OCTENE'): 0.00398772954940796,
    ('3-M-1-BUTYNE', 'TR2-PENTENE'): 0.000158071517944336,
    ('3-M-1-BUTYNE', 'TR3-HEXENE'): 0.000913560390472412, ('3-M-1-BUTYNE', 'TR3-OCTENE'): 0.00398772954940796,
    ('3-M-1-BUTYNE', 'TR4-OCTENE'): 0.00398772954940796,
    ('3-M-1-BUTYNE', 'VNYLACETLEN'): 0.000669658184051514,
    ('3M1C5=', 'ACETYLENE'): 0.0171984434127808, ('3M1C5=', 'BENZENE'): 0.0011448860168457,
    ('3M1C5=', 'BIACETYLENE'): 0.00558596849441528, ('3M1C5=', 'CIS2-BUTENE'): 0.0021393895149231,
    ('3M1C5=', 'CIS2-HEXENE'): 2.26497650146484e-06, ('3M1C5=', 'CIS2-OCTENE'): 0.00118952989578247,
    ('3M1C5=', 'CIS2-PENTENE'): 0.000261604785919189, ('3M1C5=', 'CIS3-HEXENE'): 2.26497650146484e-06,
    ('3M1C5=', 'CIS4-OCTENE'): 0.00118952989578247, ('3M1C5=', 'CO2'): 0.101000003516674,
    ('3M1C5=', 'CS2'): 0.0400316268205643, ('3M1C5=', 'CYCLOBUTANE'): 0.00346434116363525,
    ('3M1C5=', 'CYCLOBUTENE'): 0.00453221797943115, ('3M1C5=', 'CYCLOHEXANE'): 0.000192761421203613,
    ('3M1C5=', 'CYCLOPENTANE'): 0.0011448860168457, ('3M1C5=', 'CYCLOPENTENE'): 0.0016663670539856,
    ('3M1C5=', 'CYCLOPROPANE'): 0.00692248344421387, ('3M1C5=', 'DMACETYLENE'): 0.00280243158340454,
    ('3M1C5=', 'E-BENZENE'): 8.10027122497559e-05, ('3M1C5=', 'E-MERCAPTAN'): 0.0574909187853336,
    ('3M1C5=', 'EACETYLENE'): 0.00274652242660522, ('3M1C5=', 'ECYCLOPENTAN'): 8.3625316619873e-05,
    ('3M1C5=', 'ETHANE'): 0.00996702909469604, ('3M1C5=', 'ETHYLENE'): 0.0134204626083374,
    ('3M1C5=', 'H2'): 0.292100012302399, ('3M1C5=', 'H2O'): 0.5, ('3M1C5=', 'H2S'): 0.0450000017881393,
    ('3M1C5=', 'I-BUTENE'): 0.00191551446914673, ('3M1C5=', 'IC4'): 0.00105500221252441,
    ('3M1C5=', 'IC5'): 0.000192761421203613, ('3M1C5=', 'M-ACETYLENE'): 0.0077211856842041,
    ('3M1C5=', 'M-MERCAPTAN'): 0.0574909187853336, ('3M1C5=', 'M-XYLENE'): 9.10758972167969e-05,
    ('3M1C5=', 'MCYCLOHEXANE'): 5.03063201904297e-05, ('3M1C5=', 'MCYCLOPENTAN'): 9.49501991271973e-05,
    ('3M1C5=', 'METHANE'): 0.0214084386825562, ('3M1C5=', 'N-PBENZENE'): 0.000791966915130615,
    ('3M1C5=', 'NC4'): 0.00130468606948853, ('3M1C5=', 'NC5'): 0.000162303447723389,
    ('3M1C5=', 'NC6'): 5.03063201904297e-05, ('3M1C5=', 'NC7'): 0.000592052936553955,
    ('3M1C5=', 'NC8'): 0.00158768892288208, ('3M1C5=', 'NITROGEN'): 0.100000001490116,
    ('3M1C5=', 'O-XYLENE'): 5.54919242858887e-05, ('3M1C5=', 'P-XYLENE'): 0.000111997127532959,
    ('3M1C5=', 'PROPADIENE'): 0.00797498226165771, ('3M1C5=', 'PROPANE'): 0.0041801929473877,
    ('3M1C5=', 'PROPENE'): 0.00583344697952271, ('3M1C5=', 'STYRENE'): 3.45706939697266e-06,
    ('3M1C5=', 'TOLUENE'): 0.000117957592010498, ('3M1C5=', 'TR2-BUTENE'): 0.00194978713989258,
    ('3M1C5=', 'TR2-HEXENE'): 2.26497650146484e-06, ('3M1C5=', 'TR2-OCTENE'): 0.00118952989578247,
    ('3M1C5=', 'TR2-PENTENE'): 0.000261604785919189, ('3M1C5=', 'TR3-HEXENE'): 2.26497650146484e-06,
    ('3M1C5=', 'TR3-OCTENE'): 0.00118952989578247, ('3M1C5=', 'TR4-OCTENE'): 0.00118952989578247,
    ('3M1C5=', 'VNYLACETLEN'): 0.00297737121582031,
    ('ACETYLENE', 'BENZENE'): 0.00957441329956055, ('ACETYLENE', 'BIACETYLENE'): 0.00326049327850342,
    ('ACETYLENE', 'CIS2-BUTENE'): 0.00731664896011353, ('ACETYLENE', 'CIS2-HEXENE'): 0.0175831317901611,
    ('ACETYLENE', 'CIS2-OCTENE'): 0.027127742767334, ('ACETYLENE', 'CIS2-PENTENE'): 0.0132880806922913,
    ('ACETYLENE', 'CIS3-HEXENE'): 0.0175831317901611, ('ACETYLENE', 'CIS4-OCTENE'): 0.027127742767334,
    ('ACETYLENE', 'CO2'): 0.101000003516674, ('ACETYLENE', 'CS2'): 0.0258935857564211,
    ('ACETYLENE', 'CYCLOBUTANE'): 0.00532710552215576, ('ACETYLENE', 'CYCLOBUTENE'): 0.00416439771652222,
    ('ACETYLENE', 'CYCLOHEXANE'): 0.0138124227523804, ('ACETYLENE', 'CYCLOPENTANE'): 0.00957441329956055,
    ('ACETYLENE', 'CYCLOPENTENE'): 0.00826787948608398, ('ACETYLENE', 'CYCLOPROPANE'): 0.00236219167709351,
    ('ACETYLENE', 'DMACETYLENE'): 0.00622332096099854, ('ACETYLENE', 'E-BENZENE'): 0.0195825099945068,
    ('ACETYLENE', 'E-MERCAPTAN'): 0.059960912913084, ('ACETYLENE', 'EACETYLENE'): 0.00630682706832886,
    ('ACETYLENE', 'ECYCLOPENTAN'): 0.0196220278739929, ('ACETYLENE', 'ETHANE'): 0.00101304054260254,
    ('ACETYLENE', 'ETHYLENE'): 0.065200001001358, ('ACETYLENE', 'H2'): 0.292100012302399,
    ('ACETYLENE', 'H2O'): 0.5, ('ACETYLENE', 'H2S'): 0.0450000017881393,
    ('ACETYLENE', 'I-BUTENE'): 0.00774508714675903, ('ACETYLENE', 'IC4'): 0.00983846187591553,
    ('ACETYLENE', 'IC5'): 0.0138124227523804, ('ACETYLENE', 'M-ACETYLENE'): 0.00192743539810181,
    ('ACETYLENE', 'M-MERCAPTAN'): 0.059960912913084, ('ACETYLENE', 'M-XYLENE'): 0.0197312831878662,
    ('ACETYLENE', 'MCYCLOHEXANE'): 0.0190653800964355, ('ACETYLENE', 'MCYCLOPENTAN'): 0.014784574508667,
    ('ACETYLENE', 'METHANE'): 0.000241637229919434, ('ACETYLENE', 'N-PBENZENE'): 0.0251400470733643,
    ('ACETYLENE', 'NC4'): 0.00913691520690918, ('ACETYLENE', 'NC5'): 0.0140775442123413,
    ('ACETYLENE', 'NC6'): 0.0190653800964355, ('ACETYLENE', 'NC7'): 0.0239813327789307,
    ('ACETYLENE', 'NC8'): 0.0288604497909546, ('ACETYLENE', 'NITROGEN'): 0.100000001490116,
    ('ACETYLENE', 'O-XYLENE'): 0.0191620588302612, ('ACETYLENE', 'P-XYLENE'): 0.0200172662734985,
    ('ACETYLENE', 'PROPADIENE'): 0.00180298089981079, ('ACETYLENE', 'PROPANE'): 0.00451582670211792,
    ('ACETYLENE', 'PROPENE'): 0.065200001001358, ('ACETYLENE', 'STYRENE'): 0.017677903175354,
    ('ACETYLENE', 'TOLUENE'): 0.0145199298858643, ('ACETYLENE', 'TR2-BUTENE'): 0.00767719745635986,
    ('ACETYLENE', 'TR2-HEXENE'): 0.0175831317901611, ('ACETYLENE', 'TR2-OCTENE'): 0.027127742767334,
    ('ACETYLENE', 'TR2-PENTENE'): 0.0132880806922913, ('ACETYLENE', 'TR3-HEXENE'): 0.0175831317901611,
    ('ACETYLENE', 'TR3-OCTENE'): 0.027127742767334, ('ACETYLENE', 'TR4-OCTENE'): 0.027127742767334,
    ('ACETYLENE', 'VNYLACETLEN'): 0.00597041845321655,
    ('ARGON', 'H2O'): 0.5, ('ARGON', 'METHANE'): 0.0230000000447035, ('ARGON', 'NH3'): -0.180000007152557,
    ('ARGON', 'O2'): 0.0104000000283122,
    ('BENZENE', 'BIACETYLENE'): 0.00168383121490479, ('BENZENE', 'CIS2-BUTENE'): 0.000154793262481689,
    ('BENZENE', 'CIS2-HEXENE'): 0.00124776363372803, ('BENZENE', 'CIS2-OCTENE'): 0.00465470552444458,
    ('BENZENE', 'CIS2-PENTENE'): 0.000312328338623047, ('BENZENE', 'CIS3-HEXENE'): 0.00124776363372803,
    ('BENZENE', 'CIS4-OCTENE'): 0.00465470552444458, ('BENZENE', 'CO'): 0.0984990000724792,
    ('BENZENE', 'CO2'): 0.0806000009179115, ('BENZENE', 'CS2'): 0.0454601384699345,
    ('BENZENE', 'CYCLOBUTANE'): 0.000629305839538574, ('BENZENE', 'CYCLOBUTENE'): 0.00112777948379517,
    ('BENZENE', 'CYCLOHEXANE'): 0.000398576259613037, ('BENZENE', 'CYCLOPENTENE'): 4.90546226501465e-05,
    ('BENZENE', 'CYCLOPROPANE'): 0.00245422124862671, ('BENZENE', 'DMACETYLENE'): 0.000366508960723877,
    ('BENZENE', 'E-BENZENE'): 0.00183349847793579, ('BENZENE', 'E-MERCAPTAN'): 0.0110213570296764,
    ('BENZENE', 'EACETYLENE'): 0.000346481800079346, ('BENZENE', 'ECYCLOPENTAN'): 0.00184589624404907,
    ('BENZENE', 'ETHANE'): 0.0199970006942749, ('BENZENE', 'ETHYLENE'): 0.0099999001249671,
    ('BENZENE', 'H2'): 0.285100013017654, ('BENZENE', 'H2O'): 0.5, ('BENZENE', 'H2S'): 0.00899980030953884,
    ('BENZENE', 'I-BUTENE'): 9.89437103271484e-05, ('BENZENE', 'IC4'): 1.84774398803711e-06,
    ('BENZENE', 'IC5'): 0.000398576259613037, ('BENZENE', 'M-ACETYLENE'): 0.00294160842895508,
    ('BENZENE', 'M-MERCAPTAN'): 0.0110213570296764, ('BENZENE', 'M-XYLENE'): 0.0018802285194397,
    ('BENZENE', 'MCYCLOHEXANE'): 0.00167429447174072, ('BENZENE', 'MCYCLOPENTAN'): 0.000580847263336182,
    ('BENZENE', 'METHANE'): 0.0399990007281303, ('BENZENE', 'N-PBENZENE'): 0.00383204221725464,
    ('BENZENE', 'NC4'): 5.24520874023438e-06, ('BENZENE', 'NC5'): 0.0160000007599592,
    ('BENZENE', 'NC6'): 0.00700000021606684, ('BENZENE', 'NC7'): -0.0020000000949949,
    ('BENZENE', 'NC8'): 0.00300000002607703, ('BENZENE', 'NITROGEN'): 0.159700006246567,
    ('BENZENE', 'O-XYLENE'): 0.00170356035232544, ('BENZENE', 'P-XYLENE'): 0.00197130441665649,
    ('BENZENE', 'PROPADIENE'): 0.00309997797012329, ('BENZENE', 'PROPANE'): 0.0199970006942749,
    ('BENZENE', 'PROPENE'): 0.00182151794433594, ('BENZENE', 'SO2'): 0.00150000001303852,
    ('BENZENE', 'STYRENE'): 0.00127369165420532, ('BENZENE', 'TOLUENE'): 0.000528454780578613,
    ('BENZENE', 'TR2-BUTENE'): 0.000106871128082275, ('BENZENE', 'TR2-HEXENE'): 0.00124776363372803,
    ('BENZENE', 'TR2-OCTENE'): 0.00465470552444458, ('BENZENE', 'TR2-PENTENE'): 0.000312328338623047,
    ('BENZENE', 'TR3-HEXENE'): 0.00124776363372803, ('BENZENE', 'TR3-OCTENE'): 0.00465470552444458,
    ('BENZENE', 'TR4-OCTENE'): 0.00465470552444458, ('BENZENE', 'VNYLACETLEN'): 0.000431716442108154,
    ('BIACETYLENE', 'CIS2-BUTENE'): 0.000818550586700439, ('BIACETYLENE', 'CIS2-HEXENE'): 0.00580942630767822,
    ('BIACETYLENE', 'CIS2-OCTENE'): 0.0118545889854431, ('BIACETYLENE', 'CIS2-PENTENE'): 0.00344032049179077,
    ('BIACETYLENE', 'CIS3-HEXENE'): 0.00580942630767822, ('BIACETYLENE', 'CIS4-OCTENE'): 0.0118545889854431,
    ('BIACETYLENE', 'CO2'): 0.101000003516674, ('BIACETYLENE', 'CS2'): 0.0371479839086533,
    ('BIACETYLENE', 'CYCLOBUTANE'): 0.000254988670349121,
    ('BIACETYLENE', 'CYCLOBUTENE'): 5.57303428649902e-05, ('BIACETYLENE', 'CYCLOHEXANE'): 0.00371301174163818,
    ('BIACETYLENE', 'CYCLOPENTANE'): 0.00168383121490479,
    ('BIACETYLENE', 'CYCLOPENTENE'): 0.00115925073623657,
    ('BIACETYLENE', 'CYCLOPROPANE'): 7.27176666259766e-05,
    ('BIACETYLENE', 'DMACETYLENE'): 0.000480055809020996, ('BIACETYLENE', 'E-BENZENE'): 0.0070006251335144,
    ('BIACETYLENE', 'E-MERCAPTAN'): 0.0541355572640896, ('BIACETYLENE', 'EACETYLENE'): 0.000503659248352051,
    ('BIACETYLENE', 'ECYCLOPENTAN'): 0.00702470541000366, ('BIACETYLENE', 'ETHANE'): 0.00064164400100708,
    ('BIACETYLENE', 'ETHYLENE'): 0.00172698497772217, ('BIACETYLENE', 'H2'): 0.292100012302399,
    ('BIACETYLENE', 'H2O'): 0.5, ('BIACETYLENE', 'H2S'): 0.0450000017881393,
    ('BIACETYLENE', 'I-BUTENE'): 0.000967204570770264, ('BIACETYLENE', 'IC4'): 0.0017966628074646,
    ('BIACETYLENE', 'IC5'): 0.00371301174163818, ('BIACETYLENE', 'M-ACETYLENE'): 0.000175237655639648,
    ('BIACETYLENE', 'M-MERCAPTAN'): 0.0541355572640896, ('BIACETYLENE', 'M-XYLENE'): 0.00709110498428345,
    ('BIACETYLENE', 'MCYCLOHEXANE'): 0.00668793916702271,
    ('BIACETYLENE', 'MCYCLOPENTAN'): 0.00423204898834229, ('BIACETYLENE', 'METHANE'): 0.00526577234268188,
    ('BIACETYLENE', 'N-PBENZENE'): 0.0105292201042175, ('BIACETYLENE', 'NC4'): 0.00150161981582642,
    ('BIACETYLENE', 'NC5'): 0.00385290384292603, ('BIACETYLENE', 'NC6'): 0.00668793916702271,
    ('BIACETYLENE', 'NC7'): 0.00977063179016113, ('BIACETYLENE', 'NC8'): 0.0130327939987183,
    ('BIACETYLENE', 'NITROGEN'): 0.100000001490116, ('BIACETYLENE', 'O-XYLENE'): 0.00674617290496826,
    ('BIACETYLENE', 'P-XYLENE'): 0.00726574659347534, ('BIACETYLENE', 'PROPADIENE'): 0.000215649604797363,
    ('BIACETYLENE', 'PROPANE'): 0.000102996826171875, ('BIACETYLENE', 'PROPENE'): 2.74181365966797e-06,
    ('BIACETYLENE', 'STYRENE'): 0.00586479902267456, ('BIACETYLENE', 'TOLUENE'): 0.00408905744552612,
    ('BIACETYLENE', 'TR2-BUTENE'): 0.00094306468963623, ('BIACETYLENE', 'TR2-HEXENE'): 0.00580942630767822,
    ('BIACETYLENE', 'TR2-OCTENE'): 0.0118545889854431, ('BIACETYLENE', 'TR2-PENTENE'): 0.00344032049179077,
    ('BIACETYLENE', 'TR3-HEXENE'): 0.00580942630767822, ('BIACETYLENE', 'TR3-OCTENE'): 0.0118545889854431,
    ('BIACETYLENE', 'TR4-OCTENE'): 0.0118545889854431, ('BIACETYLENE', 'VNYLACETLEN'): 0.00041121244430542,
    ('CIS2-BUTENE', 'CIS2-HEXENE'): 0.00227910280227661, ('CIS2-BUTENE', 'CIS2-OCTENE'): 0.00649338960647583,
    ('CIS2-BUTENE', 'CIS2-PENTENE'): 0.000906407833099365,
    ('CIS2-BUTENE', 'CIS3-HEXENE'): 0.00227910280227661, ('CIS2-BUTENE', 'CIS4-OCTENE'): 0.00649338960647583,
    ('CIS2-BUTENE', 'CO2'): 0.101000003516674, ('CIS2-BUTENE', 'CS2'): 0.0349528342485428,
    ('CIS2-BUTENE', 'CYCLOBUTANE'): 0.000159978866577148,
    ('CIS2-BUTENE', 'CYCLOBUTENE'): 0.000447452068328857, ('CIS2-BUTENE', 'CYCLOHEXANE'): 0.0010494589805603,
    ('CIS2-BUTENE', 'CYCLOPENTANE'): 0.000154793262481689,
    ('CIS2-BUTENE', 'CYCLOPENTENE'): 2.95639038085938e-05,
    ('CIS2-BUTENE', 'CYCLOPROPANE'): 0.00137841701507568,
    ('CIS2-BUTENE', 'DMACETYLENE'): 4.50015068054199e-05, ('CIS2-BUTENE', 'E-BENZENE'): 0.00304967164993286,
    ('CIS2-BUTENE', 'E-MERCAPTAN'): 0.0589056797325611, ('CIS2-BUTENE', 'EACETYLENE'): 3.814697265625e-05,
    ('CIS2-BUTENE', 'ECYCLOPENTAN'): 0.00306564569473267, ('CIS2-BUTENE', 'ETHANE'): 0.00290441513061523,
    ('CIS2-BUTENE', 'ETHYLENE'): 0.00490874052047729, ('CIS2-BUTENE', 'H2'): 0.292100012302399,
    ('CIS2-BUTENE', 'H2O'): 0.5, ('CIS2-BUTENE', 'H2S'): 0.0450000017881393,
    ('CIS2-BUTENE', 'I-BUTENE'): 6.25848770141602e-06, ('CIS2-BUTENE', 'IC4'): 0.000190377235412598,
    ('CIS2-BUTENE', 'IC5'): 0.0010494589805603, ('CIS2-BUTENE', 'M-ACETYLENE'): 0.00174969434738159,
    ('CIS2-BUTENE', 'M-MERCAPTAN'): 0.0589056797325611, ('CIS2-BUTENE', 'M-XYLENE'): 0.00310987234115601,
    ('CIS2-BUTENE', 'MCYCLOHEXANE'): 0.00284361839294434,
    ('CIS2-BUTENE', 'MCYCLOPENTAN'): 0.00133430957794189, ('CIS2-BUTENE', 'METHANE'): 0.0101836919784546,
    ('CIS2-BUTENE', 'N-PBENZENE'): 0.00551652908325195, ('CIS2-BUTENE', 'NC4'): 0.00010305643081665,
    ('CIS2-BUTENE', 'NC5'): 0.00112468004226685, ('CIS2-BUTENE', 'NC6'): 0.00284361839294434,
    ('CIS2-BUTENE', 'NC7'): 0.00496810674667358, ('CIS2-BUTENE', 'NC8'): 0.00737893581390381,
    ('CIS2-BUTENE', 'NITROGEN'): 0.100000001490116, ('CIS2-BUTENE', 'O-XYLENE'): 0.00288176536560059,
    ('CIS2-BUTENE', 'P-XYLENE'): 0.00322645902633667, ('CIS2-BUTENE', 'PROPADIENE'): 0.00187253952026367,
    ('CIS2-BUTENE', 'PROPANE'): 0.000341176986694336, ('CIS2-BUTENE', 'PROPENE'): 0.00091552734375,
    ('CIS2-BUTENE', 'STYRENE'): 0.00231385231018066, ('CIS2-BUTENE', 'TOLUENE'): 0.0012543797492981,
    ('CIS2-BUTENE', 'TR2-BUTENE'): 4.41074371337891e-06, ('CIS2-BUTENE', 'TR2-HEXENE'): 0.00227910280227661,
    ('CIS2-BUTENE', 'TR2-OCTENE'): 0.00649338960647583, ('CIS2-BUTENE', 'TR2-PENTENE'): 0.000906407833099365,
    ('CIS2-BUTENE', 'TR3-HEXENE'): 0.00227910280227661, ('CIS2-BUTENE', 'TR3-OCTENE'): 0.00649338960647583,
    ('CIS2-BUTENE', 'TR4-OCTENE'): 0.00649338960647583, ('CIS2-BUTENE', 'VNYLACETLEN'): 6.96182250976563e-05,
    ('CIS2-HEXENE', 'CIS2-OCTENE'): 0.00108909606933594,
    ('CIS2-HEXENE', 'CIS2-PENTENE'): 0.000312089920043945,
    ('CIS2-HEXENE', 'CIS4-OCTENE'): 0.00108909606933594, ('CIS2-HEXENE', 'CO2'): 0.101000003516674,
    ('CIS2-HEXENE', 'CS2'): 0.0415643863379955, ('CIS2-HEXENE', 'CYCLOBUTANE'): 0.00364124774932861,
    ('CIS2-HEXENE', 'CYCLOBUTENE'): 0.0047338604927063, ('CIS2-HEXENE', 'CYCLOHEXANE'): 0.00109929998870939,
    ('CIS2-HEXENE', 'CYCLOPENTANE'): 0.00124776363372803,
    ('CIS2-HEXENE', 'CYCLOPENTENE'): 0.00178998708724976,
    ('CIS2-HEXENE', 'CYCLOPROPANE'): 0.00717061758041382, ('CIS2-HEXENE', 'DMACETYLENE'): 0.00296193361282349,
    ('CIS2-HEXENE', 'E-BENZENE'): 5.64455986022949e-05, ('CIS2-HEXENE', 'E-MERCAPTAN'): 0.0563712045550346,
    ('CIS2-HEXENE', 'EACETYLENE'): 0.00290447473526001, ('CIS2-HEXENE', 'ECYCLOPENTAN'): 5.8591365814209e-05,
    ('CIS2-HEXENE', 'ETHANE'): 0.0102631449699402, ('CIS2-HEXENE', 'ETHYLENE'): 0.0137620568275452,
    ('CIS2-HEXENE', 'H2'): 0.292100012302399, ('CIS2-HEXENE', 'H2O'): 0.5,
    ('CIS2-HEXENE', 'H2S'): 0.0450000017881393, ('CIS2-HEXENE', 'I-BUTENE'): 0.00204789638519287,
    ('CIS2-HEXENE', 'IC4'): 0.00115388631820679, ('CIS2-HEXENE', 'IC5'): 0.000236332416534424,
    ('CIS2-HEXENE', 'M-ACETYLENE'): 0.00798279047012329, ('CIS2-HEXENE', 'M-MERCAPTAN'): 0.0563712045550346,
    ('CIS2-HEXENE', 'M-XYLENE'): 6.49094581604004e-05, ('CIS2-HEXENE', 'MCYCLOHEXANE'): 3.14116477966309e-05,
    ('CIS2-HEXENE', 'MCYCLOPENTAN'): 0.000126183032989502, ('CIS2-HEXENE', 'METHANE'): 0.0218349695205688,
    ('CIS2-HEXENE', 'N-PBENZENE'): 0.00071030855178833, ('CIS2-HEXENE', 'NC4'): 0.00141441822052002,
    ('CIS2-HEXENE', 'NC5'): 0.000202596187591553, ('CIS2-HEXENE', 'NC6'): 3.14116477966309e-05,
    ('CIS2-HEXENE', 'NC7'): 0.000521719455718994, ('CIS2-HEXENE', 'NC8'): 0.00147128105163574,
    ('CIS2-HEXENE', 'NITROGEN'): 0.100000001490116, ('CIS2-HEXENE', 'O-XYLENE'): 3.55243682861328e-05,
    ('CIS2-HEXENE', 'P-XYLENE'): 8.27312469482422e-05, ('CIS2-HEXENE', 'PROPADIENE'): 0.00824052095413208,
    ('CIS2-HEXENE', 'PROPANE'): 0.00437402725219727, ('CIS2-HEXENE', 'PROPENE'): 0.0060616135597229,
    ('CIS2-HEXENE', 'STYRENE'): 1.78813934326172e-07, ('CIS2-HEXENE', 'TOLUENE'): 0.000152468681335449,
    ('CIS2-HEXENE', 'TR2-BUTENE'): 0.00208330154418945, ('CIS2-HEXENE', 'TR2-OCTENE'): 0.00108909606933594,
    ('CIS2-HEXENE', 'TR2-PENTENE'): 0.000312089920043945, ('CIS2-HEXENE', 'TR3-OCTENE'): 0.00108909606933594,
    ('CIS2-HEXENE', 'TR4-OCTENE'): 0.00108909606933594, ('CIS2-HEXENE', 'VNYLACETLEN'): 0.00314152240753174,
    ('CIS2-OCTENE', 'CIS2-PENTENE'): 0.0025632381439209, ('CIS2-OCTENE', 'CIS3-HEXENE'): 0.00108909606933594,
    ('CIS2-OCTENE', 'CO2'): 0.101000003516674, ('CIS2-OCTENE', 'CS2'): 0.045986469835043,
    ('CIS2-OCTENE', 'CYCLOBUTANE'): 0.00866991281509399, ('CIS2-OCTENE', 'CYCLOBUTENE'): 0.0103057026863098,
    ('CIS2-OCTENE', 'CYCLOHEXANE'): 0.00233715772628784, ('CIS2-OCTENE', 'CYCLOPENTANE'): 0.00465470552444458,
    ('CIS2-OCTENE', 'CYCLOPENTENE'): 0.00565195083618164, ('CIS2-OCTENE', 'CYCLOPROPANE'): 0.0137525796890259,
    ('CIS2-OCTENE', 'DMACETYLENE'): 0.00760889053344727, ('CIS2-OCTENE', 'E-BENZENE'): 0.000650167465209961,
    ('CIS2-OCTENE', 'E-MERCAPTAN'): 0.0519249103963375, ('CIS2-OCTENE', 'EACETYLENE'): 0.00751721858978271,
    ('CIS2-OCTENE', 'ECYCLOPENTAN'): 0.000642776489257813, ('CIS2-OCTENE', 'ETHANE'): 0.0178892016410828,
    ('CIS2-OCTENE', 'ETHYLENE'): 0.022377610206604, ('CIS2-OCTENE', 'H2'): 0.292100012302399,
    ('CIS2-OCTENE', 'H2O'): 0.5, ('CIS2-OCTENE', 'H2S'): 0.0450000017881393,
    ('CIS2-OCTENE', 'I-BUTENE'): 0.00610101222991943, ('CIS2-OCTENE', 'IC4'): 0.00447249412536621,
    ('CIS2-OCTENE', 'IC5'): 0.00233715772628784, ('CIS2-OCTENE', 'M-ACETYLENE'): 0.0148594379425049,
    ('CIS2-OCTENE', 'M-MERCAPTAN'): 0.0519249103963375, ('CIS2-OCTENE', 'M-XYLENE'): 0.000622749328613281,
    ('CIS2-OCTENE', 'MCYCLOHEXANE'): 0.000750899314880371,
    ('CIS2-OCTENE', 'MCYCLOPENTAN'): 0.00195497274398804, ('CIS2-OCTENE', 'METHANE'): 0.0322840809822083,
    ('CIS2-OCTENE', 'N-PBENZENE'): 4.04119491577148e-05, ('CIS2-OCTENE', 'NC4'): 0.00497031211853027,
    ('CIS2-OCTENE', 'NC5'): 0.00222837924957275, ('CIS2-OCTENE', 'NC6'): 0.000750899314880371,
    ('CIS2-OCTENE', 'NC7'): 0.000103354454040527, ('CIS2-OCTENE', 'NC8'): 2.87890434265137e-05,
    ('CIS2-OCTENE', 'NITROGEN'): 0.100000001490116, ('CIS2-OCTENE', 'O-XYLENE'): 0.000731468200683594,
    ('CIS2-OCTENE', 'P-XYLENE'): 0.000571966171264648, ('CIS2-OCTENE', 'PROPADIENE'): 0.0152073502540588,
    ('CIS2-OCTENE', 'PROPANE'): 0.00977498292922974, ('CIS2-OCTENE', 'PROPENE'): 0.0122110247612,
    ('CIS2-OCTENE', 'STYRENE'): 0.00106513500213623, ('CIS2-OCTENE', 'TOLUENE'): 0.00205451250076294,
    ('CIS2-OCTENE', 'TR2-BUTENE'): 0.00616168975830078, ('CIS2-OCTENE', 'TR2-HEXENE'): 0.00108909606933594,
    ('CIS2-OCTENE', 'TR2-PENTENE'): 0.0025632381439209, ('CIS2-OCTENE', 'TR3-HEXENE'): 0.00108909606933594,
    ('CIS2-OCTENE', 'VNYLACETLEN'): 0.00789350271224976,
    ('CIS2-PENTENE', 'CIS3-HEXENE'): 0.000312089920043945,
    ('CIS2-PENTENE', 'CIS4-OCTENE'): 0.0025632381439209, ('CIS2-PENTENE', 'CO2'): 0.101000003516674,
    ('CIS2-PENTENE', 'CS2'): 0.0383603684604168, ('CIS2-PENTENE', 'CYCLOBUTANE'): 0.00182634592056274,
    ('CIS2-PENTENE', 'CYCLOBUTENE'): 0.00262314081192017,
    ('CIS2-PENTENE', 'CYCLOHEXANE'): 5.30481338500977e-06,
    ('CIS2-PENTENE', 'CYCLOPENTANE'): 0.000312328338623047,
    ('CIS2-PENTENE', 'CYCLOPENTENE'): 0.000608503818511963,
    ('CIS2-PENTENE', 'CYCLOPROPANE'): 0.00450778007507324,
    ('CIS2-PENTENE', 'DMACETYLENE'): 0.00135445594787598, ('CIS2-PENTENE', 'E-BENZENE'): 0.000633478164672852,
    ('CIS2-PENTENE', 'E-MERCAPTAN'): 0.058039303869009, ('CIS2-PENTENE', 'EACETYLENE'): 0.0013156533241272,
    ('CIS2-PENTENE', 'ECYCLOPENTAN'): 0.000640869140625, ('CIS2-PENTENE', 'ETHANE'): 0.00702744722366333,
    ('CIS2-PENTENE', 'ETHYLENE'): 0.00998121500015259, ('CIS2-PENTENE', 'H2'): 0.292100012302399,
    ('CIS2-PENTENE', 'H2O'): 0.5, ('CIS2-PENTENE', 'H2S'): 0.0450000017881393,
    ('CIS2-PENTENE', 'I-BUTENE'): 0.000762701034545898, ('CIS2-PENTENE', 'IC4'): 0.000266194343566895,
    ('CIS2-PENTENE', 'IC5'): 5.30481338500977e-06, ('CIS2-PENTENE', 'M-ACETYLENE'): 0.00515866279602051,
    ('CIS2-PENTENE', 'M-MERCAPTAN'): 0.058039303869009, ('CIS2-PENTENE', 'M-XYLENE'): 0.000661194324493408,
    ('CIS2-PENTENE', 'MCYCLOHEXANE'): 0.000541329383850098,
    ('CIS2-PENTENE', 'MCYCLOPENTAN'): 4.13656234741211e-05, ('CIS2-PENTENE', 'METHANE'): 0.0170372128486633,
    ('CIS2-PENTENE', 'N-PBENZENE'): 0.00196164846420288, ('CIS2-PENTENE', 'NC4'): 0.000398457050323486,
    ('CIS2-PENTENE', 'NC5'): 1.18613243103027e-05, ('CIS2-PENTENE', 'NC6'): 0.000541329383850098,
    ('CIS2-PENTENE', 'NC7'): 0.00163906812667847, ('CIS2-PENTENE', 'NC8'): 0.00313323736190796,
    ('CIS2-PENTENE', 'NITROGEN'): 0.100000001490116, ('CIS2-PENTENE', 'O-XYLENE'): 0.000558078289031982,
    ('CIS2-PENTENE', 'P-XYLENE'): 0.000715672969818115, ('CIS2-PENTENE', 'PROPADIENE'): 0.00536710023880005,
    ('CIS2-PENTENE', 'PROPANE'): 0.00235652923583984, ('CIS2-PENTENE', 'PROPENE'): 0.00363552570343018,
    ('CIS2-PENTENE', 'STYRENE'): 0.00032508373260498, ('CIS2-PENTENE', 'TOLUENE'): 2.83122062683105e-05,
    ('CIS2-PENTENE', 'TR2-BUTENE'): 0.000784456729888916,
    ('CIS2-PENTENE', 'TR2-HEXENE'): 0.000312089920043945, ('CIS2-PENTENE', 'TR2-OCTENE'): 0.0025632381439209,
    ('CIS2-PENTENE', 'TR3-HEXENE'): 0.000312089920043945, ('CIS2-PENTENE', 'TR3-OCTENE'): 0.0025632381439209,
    ('CIS2-PENTENE', 'TR4-OCTENE'): 0.0025632381439209, ('CIS2-PENTENE', 'VNYLACETLEN'): 0.00147712230682373,
    ('CIS3-HEXENE', 'CIS4-OCTENE'): 0.00108909606933594, ('CIS3-HEXENE', 'CO2'): 0.101000003516674,
    ('CIS3-HEXENE', 'CS2'): 0.041229635477066, ('CIS3-HEXENE', 'CYCLOBUTANE'): 0.00364124774932861,
    ('CIS3-HEXENE', 'CYCLOBUTENE'): 0.0047338604927063, ('CIS3-HEXENE', 'CYCLOHEXANE'): 0.000236332416534424,
    ('CIS3-HEXENE', 'CYCLOPENTANE'): 0.00124776363372803,
    ('CIS3-HEXENE', 'CYCLOPENTENE'): 0.00178998708724976,
    ('CIS3-HEXENE', 'CYCLOPROPANE'): 0.00717061758041382, ('CIS3-HEXENE', 'DMACETYLENE'): 0.00296193361282349,
    ('CIS3-HEXENE', 'E-BENZENE'): 5.64455986022949e-05, ('CIS3-HEXENE', 'E-MERCAPTAN'): 0.0568540394306183,
    ('CIS3-HEXENE', 'EACETYLENE'): 0.00290447473526001, ('CIS3-HEXENE', 'ECYCLOPENTAN'): 5.8591365814209e-05,
    ('CIS3-HEXENE', 'ETHANE'): 0.0102631449699402, ('CIS3-HEXENE', 'ETHYLENE'): 0.0137620568275452,
    ('CIS3-HEXENE', 'H2'): 0.292100012302399, ('CIS3-HEXENE', 'H2O'): 0.5,
    ('CIS3-HEXENE', 'H2S'): 0.0450000017881393, ('CIS3-HEXENE', 'I-BUTENE'): 0.00204789638519287,
    ('CIS3-HEXENE', 'IC4'): 0.00115388631820679, ('CIS3-HEXENE', 'IC5'): 0.000236332416534424,
    ('CIS3-HEXENE', 'M-ACETYLENE'): 0.00798279047012329, ('CIS3-HEXENE', 'M-MERCAPTAN'): 0.0568540394306183,
    ('CIS3-HEXENE', 'M-XYLENE'): 6.49094581604004e-05, ('CIS3-HEXENE', 'MCYCLOHEXANE'): 3.14116477966309e-05,
    ('CIS3-HEXENE', 'MCYCLOPENTAN'): 0.000126183032989502, ('CIS3-HEXENE', 'METHANE'): 0.0218349695205688,
    ('CIS3-HEXENE', 'N-PBENZENE'): 0.00071030855178833, ('CIS3-HEXENE', 'NC4'): 0.00141441822052002,
    ('CIS3-HEXENE', 'NC5'): 0.000202596187591553, ('CIS3-HEXENE', 'NC6'): 3.14116477966309e-05,
    ('CIS3-HEXENE', 'NC7'): 0.000521719455718994, ('CIS3-HEXENE', 'NC8'): 0.00147128105163574,
    ('CIS3-HEXENE', 'NITROGEN'): 0.100000001490116, ('CIS3-HEXENE', 'O-XYLENE'): 3.55243682861328e-05,
    ('CIS3-HEXENE', 'P-XYLENE'): 8.27312469482422e-05, ('CIS3-HEXENE', 'PROPADIENE'): 0.00824052095413208,
    ('CIS3-HEXENE', 'PROPANE'): 0.00437402725219727, ('CIS3-HEXENE', 'PROPENE'): 0.0060616135597229,
    ('CIS3-HEXENE', 'STYRENE'): 1.78813934326172e-07, ('CIS3-HEXENE', 'TOLUENE'): 0.000152468681335449,
    ('CIS3-HEXENE', 'TR2-BUTENE'): 0.00208330154418945, ('CIS3-HEXENE', 'TR2-OCTENE'): 0.00108909606933594,
    ('CIS3-HEXENE', 'TR2-PENTENE'): 0.000312089920043945, ('CIS3-HEXENE', 'TR3-OCTENE'): 0.00108909606933594,
    ('CIS3-HEXENE', 'TR4-OCTENE'): 0.00108909606933594, ('CIS3-HEXENE', 'VNYLACETLEN'): 0.00314152240753174,
    ('CIS4-OCTENE', 'CO2'): 0.101000003516674, ('CIS4-OCTENE', 'CS2'): 0.0457814186811447,
    ('CIS4-OCTENE', 'CYCLOBUTANE'): 0.00866991281509399, ('CIS4-OCTENE', 'CYCLOBUTENE'): 0.0103057026863098,
    ('CIS4-OCTENE', 'CYCLOHEXANE'): 0.00233715772628784, ('CIS4-OCTENE', 'CYCLOPENTANE'): 0.00465470552444458,
    ('CIS4-OCTENE', 'CYCLOPENTENE'): 0.00565195083618164, ('CIS4-OCTENE', 'CYCLOPROPANE'): 0.0137525796890259,
    ('CIS4-OCTENE', 'DMACETYLENE'): 0.00760889053344727, ('CIS4-OCTENE', 'E-BENZENE'): 0.000650167465209961,
    ('CIS4-OCTENE', 'E-MERCAPTAN'): 0.0523723736405373, ('CIS4-OCTENE', 'EACETYLENE'): 0.00751721858978271,
    ('CIS4-OCTENE', 'ECYCLOPENTAN'): 0.000642776489257813, ('CIS4-OCTENE', 'ETHANE'): 0.0178892016410828,
    ('CIS4-OCTENE', 'ETHYLENE'): 0.022377610206604, ('CIS4-OCTENE', 'H2'): 0.292100012302399,
    ('CIS4-OCTENE', 'H2O'): 0.5, ('CIS4-OCTENE', 'H2S'): 0.0450000017881393,
    ('CIS4-OCTENE', 'I-BUTENE'): 0.00610101222991943, ('CIS4-OCTENE', 'IC4'): 0.00447249412536621,
    ('CIS4-OCTENE', 'IC5'): 0.00233715772628784, ('CIS4-OCTENE', 'M-ACETYLENE'): 0.0148594379425049,
    ('CIS4-OCTENE', 'M-MERCAPTAN'): 0.0523723736405373, ('CIS4-OCTENE', 'M-XYLENE'): 0.000622749328613281,
    ('CIS4-OCTENE', 'MCYCLOHEXANE'): 0.000750899314880371,
    ('CIS4-OCTENE', 'MCYCLOPENTAN'): 0.00195497274398804, ('CIS4-OCTENE', 'METHANE'): 0.0322840809822083,
    ('CIS4-OCTENE', 'N-PBENZENE'): 4.04119491577148e-05, ('CIS4-OCTENE', 'NC4'): 0.00497031211853027,
    ('CIS4-OCTENE', 'NC5'): 0.00222837924957275, ('CIS4-OCTENE', 'NC6'): 0.000750899314880371,
    ('CIS4-OCTENE', 'NC7'): 0.000103354454040527, ('CIS4-OCTENE', 'NC8'): 2.87890434265137e-05,
    ('CIS4-OCTENE', 'NITROGEN'): 0.100000001490116, ('CIS4-OCTENE', 'O-XYLENE'): 0.000731468200683594,
    ('CIS4-OCTENE', 'P-XYLENE'): 0.000571966171264648, ('CIS4-OCTENE', 'PROPADIENE'): 0.0152073502540588,
    ('CIS4-OCTENE', 'PROPANE'): 0.00977498292922974, ('CIS4-OCTENE', 'PROPENE'): 0.0122110247612,
    ('CIS4-OCTENE', 'STYRENE'): 0.00106513500213623, ('CIS4-OCTENE', 'TOLUENE'): 0.00205451250076294,
    ('CIS4-OCTENE', 'TR2-BUTENE'): 0.00616168975830078, ('CIS4-OCTENE', 'TR2-HEXENE'): 0.00108909606933594,
    ('CIS4-OCTENE', 'TR2-PENTENE'): 0.0025632381439209, ('CIS4-OCTENE', 'TR3-HEXENE'): 0.00108909606933594,
    ('CIS4-OCTENE', 'VNYLACETLEN'): 0.00789350271224976,
    ('CO', 'CO2'): -0.0313990004360676, ('CO', 'ETHANE'): 0.0253970008343458,
    ('CO', 'H2'): 0.0252960007637739, ('CO', 'H2O'): -0.38960000872612, ('CO', 'H2S'): 0.0874990001320839,
    ('CO', 'METHANE'): 0.0209970008581877, ('CO', 'NC8'): 0.14640000462532,
    ('CO', 'NITROGEN'): 0.0115000000223517, ('CO', 'PROPANE'): 0.0612980015575886,
    ('CO2', 'CYCLOBUTANE'): 0.101000003516674, ('CO2', 'CYCLOBUTENE'): 0.101000003516674,
    ('CO2', 'CYCLOHEXANE'): 0.09009800106287, ('CO2', 'CYCLOPENTANE'): 0.101000003516674,
    ('CO2', 'CYCLOPENTENE'): 0.101000003516674, ('CO2', 'CYCLOPROPANE'): 0.101000003516674,
    ('CO2', 'DMACETYLENE'): 0.101000003516674, ('CO2', 'E-BENZENE'): 0.101000003516674,
    ('CO2', 'EACETYLENE'): 0.101000003516674, ('CO2', 'ECYCLOPENTAN'): 0.101000003516674,
    ('CO2', 'ETHANE'): 0.129800006747246, ('CO2', 'ETHYLENE'): 0.066200003027916,
    ('CO2', 'H2'): 0.120200000703335, ('CO2', 'H2O'): 0.0445000007748604, ('CO2', 'H2S'): 0.100000001490116,
    ('CO2', 'HELIUM'): 0.796700000762939, ('CO2', 'I-BUTENE'): 0.101000003516674,
    ('CO2', 'IC4'): 0.129800006747246, ('CO2', 'IC5'): 0.125, ('CO2', 'M-ACETYLENE'): 0.101000003516674,
    ('CO2', 'M-XYLENE'): 0.087898001074791, ('CO2', 'MCYCLOHEXANE'): 0.101800002157688,
    ('CO2', 'MCYCLOPENTAN'): 0.101000003516674, ('CO2', 'METHANE'): 0.100000001490116,
    ('CO2', 'N-PBENZENE'): 0.100000001490116, ('CO2', 'NC4'): 0.129800006747246, ('CO2', 'NC5'): 0.125,
    ('CO2', 'NC6'): 0.125, ('CO2', 'NC7'): 0.1199000030756, ('CO2', 'NC8'): 0.115000002086163,
    ('CO2', 'NITROGEN'): -0.0199970006942749, ('CO2', 'O-XYLENE'): 0.0900000035762787,
    ('CO2', 'O2'): 0.0974970012903214, ('CO2', 'P-XYLENE'): 0.0900000035762787,
    ('CO2', 'PROPADIENE'): 0.101000003516674, ('CO2', 'PROPANE'): 0.135000005364418,
    ('CO2', 'PROPENE'): 0.0833000019192696, ('CO2', 'STYRENE'): 0.101000003516674,
    ('CO2', 'TOLUENE'): 0.0935980007052422, ('CO2', 'TR2-BUTENE'): 0.101000003516674,
    ('CO2', 'TR2-HEXENE'): 0.101000003516674, ('CO2', 'TR2-OCTENE'): 0.101000003516674,
    ('CO2', 'TR2-PENTENE'): 0.101000003516674, ('CO2', 'TR3-HEXENE'): 0.101000003516674,
    ('CO2', 'TR3-OCTENE'): 0.101000003516674, ('CO2', 'TR4-OCTENE'): 0.101000003516674,
    ('CO2', 'VNYLACETLEN'): 0.101000003516674,
    ('COS', 'ETHANE'): 0.0249589998275042, ('COS', 'METHANE'): 0.0288980007171631,
    ('COS', 'NC4'): 0.0157570000737906, ('COS', 'PROPANE'): 0.0105010000988841,
    ('CS2', 'CYCLOBUTANE'): 0.0370052196085453, ('CS2', 'CYCLOBUTENE'): 0.0363510772585869,
    ('CS2', 'CYCLOHEXANE'): 0.0440719537436962, ('CS2', 'CYCLOPENTANE'): 0.0410678796470165,
    ('CS2', 'CYCLOPENTENE'): 0.0411330536007881, ('CS2', 'CYCLOPROPANE'): 0.0287012159824371,
    ('CS2', 'DMACETYLENE'): 0.0382092446088791, ('CS2', 'E-BENZENE'): 0.0476453192532063,
    ('CS2', 'EACETYLENE'): 0.03592224791646, ('CS2', 'ECYCLOPENTAN'): 0.0453663803637028,
    ('CS2', 'ETHANE'): 0.0281559992581606, ('CS2', 'ETHYLENE'): 0.0628990009427071,
    ('CS2', 'I-BUTENE'): 0.0335002057254314, ('CS2', 'IC4'): 0.0325740128755569,
    ('CS2', 'IC5'): 0.0368969962000847, ('CS2', 'M-ACETYLENE'): 0.0326670631766319,
    ('CS2', 'M-XYLENE'): 0.0476763434708118, ('CS2', 'MCYCLOHEXANE'): 0.0452557094395161,
    ('CS2', 'MCYCLOPENTAN'): 0.042920708656311, ('CS2', 'METHANE'): 0.0222671125084162,
    ('CS2', 'N-PBENZENE'): 0.0477120839059353, ('CS2', 'NC4'): -0.00461000017821789,
    ('CS2', 'NC5'): 0.0376949980854988, ('CS2', 'NC6'): 0.0410015769302845,
    ('CS2', 'NC7'): 0.0436688475310802, ('CS2', 'NC8'): 0.0456682927906513,
    ('CS2', 'O-XYLENE'): 0.0477494038641453, ('CS2', 'P-XYLENE'): 0.0476571470499039,
    ('CS2', 'PROPADIENE'): 0.0323005989193916, ('CS2', 'PROPANE'): -0.00146000005770475,
    ('CS2', 'PROPENE'): -0.0170770008116961, ('CS2', 'STYRENE'): 0.0477396622300148,
    ('CS2', 'TOLUENE'): 0.0469186007976532, ('CS2', 'TR2-BUTENE'): 0.0344058759510517,
    ('CS2', 'TR2-HEXENE'): 0.0413039401173592, ('CS2', 'TR2-OCTENE'): 0.0458946190774441,
    ('CS2', 'TR2-PENTENE'): 0.0381676107645035, ('CS2', 'TR3-HEXENE'): 0.0412281602621078,
    ('CS2', 'TR3-OCTENE'): 0.045733630657196, ('CS2', 'TR4-OCTENE'): 0.0456622317433357,
    ('CS2', 'VNYLACETLEN'): 0.0361861623823643,
    ('CYCLOBUTANE', 'CYCLOBUTENE'): 7.22408294677734e-05, ('CYCLOBUTANE', 'CYCLOHEXANE'): 0.00202697515487671,
    ('CYCLOBUTANE', 'CYCLOPENTANE'): 0.000629305839538574,
    ('CYCLOBUTANE', 'CYCLOPENTENE'): 0.00032728910446167,
    ('CYCLOBUTANE', 'CYCLOPROPANE'): 0.000599861145019531,
    ('CYCLOBUTANE', 'DMACETYLENE'): 3.52859497070313e-05, ('CYCLOBUTANE', 'E-BENZENE'): 0.00459873676300049,
    ('CYCLOBUTANE', 'E-MERCAPTAN'): 0.0555635988712311, ('CYCLOBUTANE', 'EACETYLENE'): 4.19020652770996e-05,
    ('CYCLOBUTANE', 'ECYCLOPENTAN'): 0.00461828708648682, ('CYCLOBUTANE', 'ETHANE'): 0.00170397758483887,
    ('CYCLOBUTANE', 'ETHYLENE'): 0.00330358743667603, ('CYCLOBUTANE', 'H2'): 0.292100012302399,
    ('CYCLOBUTANE', 'H2O'): 0.5, ('CYCLOBUTANE', 'H2S'): 0.0450000017881393,
    ('CYCLOBUTANE', 'I-BUTENE'): 0.000229299068450928, ('CYCLOBUTANE', 'IC4'): 0.000699102878570557,
    ('CYCLOBUTANE', 'IC5'): 0.00202697515487671, ('CYCLOBUTANE', 'M-ACETYLENE'): 0.000852584838867188,
    ('CYCLOBUTANE', 'M-MERCAPTAN'): 0.0555635988712311, ('CYCLOBUTANE', 'M-XYLENE'): 0.00467240810394287,
    ('CYCLOBUTANE', 'MCYCLOHEXANE'): 0.00434529781341553,
    ('CYCLOBUTANE', 'MCYCLOPENTAN'): 0.00241565704345703, ('CYCLOBUTANE', 'METHANE'): 0.00781559944152832,
    ('CYCLOBUTANE', 'N-PBENZENE'): 0.00753790140151978, ('CYCLOBUTANE', 'NC4'): 0.000519752502441406,
    ('CYCLOBUTANE', 'NC5'): 0.00213080644607544, ('CYCLOBUTANE', 'NC6'): 0.00434529781341553,
    ('CYCLOBUTANE', 'NC7'): 0.00689595937728882, ('CYCLOBUTANE', 'NC8'): 0.00968587398529053,
    ('CYCLOBUTANE', 'NITROGEN'): 0.100000001490116, ('CYCLOBUTANE', 'O-XYLENE'): 0.00439244508743286,
    ('CYCLOBUTANE', 'P-XYLENE'): 0.00481480360031128, ('CYCLOBUTANE', 'PROPADIENE'): 0.000939130783081055,
    ('CYCLOBUTANE', 'PROPANE'): 3.39746475219727e-05, ('CYCLOBUTANE', 'PROPENE'): 0.000310301780700684,
    ('CYCLOBUTANE', 'STYRENE'): 0.00368523597717285, ('CYCLOBUTANE', 'TOLUENE'): 0.00230777263641357,
    ('CYCLOBUTANE', 'TR2-BUTENE'): 0.000217616558074951, ('CYCLOBUTANE', 'TR2-HEXENE'): 0.00364124774932861,
    ('CYCLOBUTANE', 'TR2-OCTENE'): 0.00866991281509399, ('CYCLOBUTANE', 'TR2-PENTENE'): 0.00182634592056274,
    ('CYCLOBUTANE', 'TR3-HEXENE'): 0.00364124774932861, ('CYCLOBUTANE', 'TR3-OCTENE'): 0.00866991281509399,
    ('CYCLOBUTANE', 'TR4-OCTENE'): 0.00866991281509399, ('CYCLOBUTANE', 'VNYLACETLEN'): 1.86562538146973e-05,
    ('CYCLOBUTENE', 'CYCLOHEXANE'): 0.00286227464675903, ('CYCLOBUTENE', 'CYCLOPENTANE'): 0.00112777948379517,
    ('CYCLOBUTENE', 'CYCLOPENTENE'): 0.000707089900970459,
    ('CYCLOBUTENE', 'CYCLOPROPANE'): 0.000255703926086426,
    ('CYCLOBUTENE', 'DMACETYLENE'): 0.000208735466003418, ('CYCLOBUTENE', 'E-BENZENE'): 0.00581622123718262,
    ('CYCLOBUTENE', 'E-MERCAPTAN'): 0.0545637346804142, ('CYCLOBUTENE', 'EACETYLENE'): 0.00022435188293457,
    ('CYCLOBUTENE', 'ECYCLOPENTAN'): 0.00583815574645996, ('CYCLOBUTENE', 'ETHANE'): 0.0010751485824585,
    ('CYCLOBUTENE', 'ETHYLENE'): 0.00240123271942139, ('CYCLOBUTENE', 'H2'): 0.292100012302399,
    ('CYCLOBUTENE', 'H2O'): 0.5, ('CYCLOBUTENE', 'H2S'): 0.0450000017881393,
    ('CYCLOBUTENE', 'I-BUTENE'): 0.000558972358703613, ('CYCLOBUTENE', 'IC4'): 0.00122052431106567,
    ('CYCLOBUTENE', 'IC5'): 0.00286227464675903, ('CYCLOBUTENE', 'M-ACETYLENE'): 0.000428557395935059,
    ('CYCLOBUTENE', 'M-MERCAPTAN'): 0.0545637346804142, ('CYCLOBUTENE', 'M-XYLENE'): 0.00589883327484131,
    ('CYCLOBUTENE', 'MCYCLOHEXANE'): 0.00553107261657715,
    ('CYCLOBUTENE', 'MCYCLOPENTAN'): 0.00332057476043701, ('CYCLOBUTENE', 'METHANE'): 0.00639635324478149,
    ('CYCLOBUTENE', 'N-PBENZENE'): 0.00907027721405029, ('CYCLOBUTENE', 'NC4'): 0.000979483127593994,
    ('CYCLOBUTENE', 'NC5'): 0.00298529863357544, ('CYCLOBUTENE', 'NC6'): 0.00553107261657715,
    ('CYCLOBUTENE', 'NC7'): 0.00836598873138428, ('CYCLOBUTENE', 'NC8'): 0.0114082098007202,
    ('CYCLOBUTENE', 'NITROGEN'): 0.100000001490116, ('CYCLOBUTENE', 'O-XYLENE'): 0.00558418035507202,
    ('CYCLOBUTENE', 'P-XYLENE'): 0.00605833530426025, ('CYCLOBUTENE', 'PROPADIENE'): 0.000490486621856689,
    ('CYCLOBUTENE', 'PROPANE'): 7.15255737304688e-06, ('CYCLOBUTENE', 'PROPENE'): 8.30888748168945e-05,
    ('CYCLOBUTENE', 'STYRENE'): 0.00478404760360718, ('CYCLOBUTENE', 'TOLUENE'): 0.00319391489028931,
    ('CYCLOBUTENE', 'TR2-BUTENE'): 0.000540614128112793, ('CYCLOBUTENE', 'TR2-HEXENE'): 0.0047338604927063,
    ('CYCLOBUTENE', 'TR2-OCTENE'): 0.0103057026863098, ('CYCLOBUTENE', 'TR2-PENTENE'): 0.00262314081192017,
    ('CYCLOBUTENE', 'TR3-HEXENE'): 0.0047338604927063, ('CYCLOBUTENE', 'TR3-OCTENE'): 0.0103057026863098,
    ('CYCLOBUTENE', 'TR4-OCTENE'): 0.0103057026863098, ('CYCLOBUTENE', 'VNYLACETLEN'): 0.000164151191711426,
    ('CYCLOHEXANE', 'CYCLOPENTANE'): 0.000398576259613037,
    ('CYCLOHEXANE', 'CYCLOPENTENE'): 0.000726699829101563,
    ('CYCLOHEXANE', 'CYCLOPROPANE'): 0.00481879711151123, ('CYCLOHEXANE', 'DMACETYLENE'): 0.00152814388275146,
    ('CYCLOHEXANE', 'E-BENZENE'): 0.000523507595062256, ('CYCLOHEXANE', 'E-MERCAPTAN'): 0.0461881197988987,
    ('CYCLOHEXANE', 'EACETYLENE'): 0.00148683786392212, ('CYCLOHEXANE', 'ECYCLOPENTAN'): 0.0005301833152771,
    ('CYCLOHEXANE', 'ETHANE'): 0.0261000003665686, ('CYCLOHEXANE', 'ETHYLENE'): 0.0310980007052422,
    ('CYCLOHEXANE', 'H2'): 0.285100013017654, ('CYCLOHEXANE', 'H2O'): 0.5,
    ('CYCLOHEXANE', 'H2S'): 0.0450000017881393, ('CYCLOHEXANE', 'I-BUTENE'): 0.000894367694854736,
    ('CYCLOHEXANE', 'IC4'): 0.000346183776855469, ('CYCLOHEXANE', 'M-ACETYLENE'): 0.00549077987670898,
    ('CYCLOHEXANE', 'M-MERCAPTAN'): 0.0461881197988987, ('CYCLOHEXANE', 'M-XYLENE'): 0.000548720359802246,
    ('CYCLOHEXANE', 'MCYCLOHEXANE'): 0.000439941883087158,
    ('CYCLOHEXANE', 'MCYCLOPENTAN'): 1.71065330505371e-05, ('CYCLOHEXANE', 'METHANE'): 0.0392000004649162,
    ('CYCLOHEXANE', 'N-PBENZENE'): 0.00176441669464111, ('CYCLOHEXANE', 'NC4'): 0.00049513578414917,
    ('CYCLOHEXANE', 'NC5'): 0.00130000000353903, ('CYCLOHEXANE', 'NC6'): -0.00300000002607703,
    ('CYCLOHEXANE', 'NC7'): 0.00145912170410156, ('CYCLOHEXANE', 'NC8'): 0.00288277864456177,
    ('CYCLOHEXANE', 'NITROGEN'): 0.100000001490116, ('CYCLOHEXANE', 'O-XYLENE'): 0.000455141067504883,
    ('CYCLOHEXANE', 'P-XYLENE'): 0.000598490238189697, ('CYCLOHEXANE', 'PROPADIENE'): 0.00570565462112427,
    ('CYCLOHEXANE', 'PROPANE'): 0.0142980003729463, ('CYCLOHEXANE', 'PROPENE'): 0.00391578674316406,
    ('CYCLOHEXANE', 'STYRENE'): 0.000247776508331299, ('CYCLOHEXANE', 'TOLUENE'): 9.11951065063477e-06,
    ('CYCLOHEXANE', 'TR2-BUTENE'): 0.00091785192489624, ('CYCLOHEXANE', 'TR2-HEXENE'): 0.000236332416534424,
    ('CYCLOHEXANE', 'TR2-OCTENE'): 0.00233715772628784, ('CYCLOHEXANE', 'TR2-PENTENE'): 5.30481338500977e-06,
    ('CYCLOHEXANE', 'TR3-HEXENE'): 0.000236332416534424, ('CYCLOHEXANE', 'TR3-OCTENE'): 0.00233715772628784,
    ('CYCLOHEXANE', 'TR4-OCTENE'): 0.00233715772628784, ('CYCLOHEXANE', 'VNYLACETLEN'): 0.00165802240371704,
    ('CYCLOPENTANE', 'CYCLOPENTENE'): 4.90546226501465e-05,
    ('CYCLOPENTANE', 'CYCLOPROPANE'): 0.00245422124862671,
    ('CYCLOPENTANE', 'DMACETYLENE'): 0.000366508960723877, ('CYCLOPENTANE', 'E-BENZENE'): 0.00183349847793579,
    ('CYCLOPENTANE', 'E-MERCAPTAN'): 0.0512871593236923, ('CYCLOPENTANE', 'EACETYLENE'): 0.000346481800079346,
    ('CYCLOPENTANE', 'ECYCLOPENTAN'): 0.00184589624404907, ('CYCLOPENTANE', 'ETHANE'): 0.00439292192459106,
    ('CYCLOPENTANE', 'ETHYLENE'): 0.00679212808609009, ('CYCLOPENTANE', 'H2'): 0.292100012302399,
    ('CYCLOPENTANE', 'H2O'): 0.5, ('CYCLOPENTANE', 'H2S'): 0.0450000017881393,
    ('CYCLOPENTANE', 'I-BUTENE'): 9.89437103271484e-05, ('CYCLOPENTANE', 'IC4'): 1.84774398803711e-06,
    ('CYCLOPENTANE', 'IC5'): 0.000398576259613037, ('CYCLOPENTANE', 'M-ACETYLENE'): 0.00294160842895508,
    ('CYCLOPENTANE', 'M-MERCAPTAN'): 0.0512871593236923, ('CYCLOPENTANE', 'M-XYLENE'): 0.0018802285194397,
    ('CYCLOPENTANE', 'MCYCLOHEXANE'): 0.00167429447174072,
    ('CYCLOPENTANE', 'MCYCLOPENTAN'): 0.000580847263336182, ('CYCLOPENTANE', 'METHANE'): 0.0128094553947449,
    ('CYCLOPENTANE', 'N-PBENZENE'): 0.00383204221725464, ('CYCLOPENTANE', 'NC4'): 5.24520874023438e-06,
    ('CYCLOPENTANE', 'NC5'): 0.00044548511505127, ('CYCLOPENTANE', 'NC6'): 0.00167429447174072,
    ('CYCLOPENTANE', 'NC7'): 0.00337636470794678, ('CYCLOPENTANE', 'NC8'): 0.00541073083877563,
    ('CYCLOPENTANE', 'NITROGEN'): 0.100000001490116, ('CYCLOPENTANE', 'O-XYLENE'): 0.00170356035232544,
    ('CYCLOPENTANE', 'P-XYLENE'): 0.00197130441665649, ('CYCLOPENTANE', 'PROPADIENE'): 0.00309997797012329,
    ('CYCLOPENTANE', 'PROPANE'): 0.000954985618591309, ('CYCLOPENTANE', 'PROPENE'): 0.00182151794433594,
    ('CYCLOPENTANE', 'STYRENE'): 0.00127369165420532, ('CYCLOPENTANE', 'TOLUENE'): 0.000528454780578613,
    ('CYCLOPENTANE', 'TR2-BUTENE'): 0.000106871128082275, ('CYCLOPENTANE', 'TR2-HEXENE'): 0.00124776363372803,
    ('CYCLOPENTANE', 'TR2-OCTENE'): 0.00465470552444458,
    ('CYCLOPENTANE', 'TR2-PENTENE'): 0.000312328338623047,
    ('CYCLOPENTANE', 'TR3-HEXENE'): 0.00124776363372803, ('CYCLOPENTANE', 'TR3-OCTENE'): 0.00465470552444458,
    ('CYCLOPENTANE', 'TR4-OCTENE'): 0.00465470552444458,
    ('CYCLOPENTANE', 'VNYLACETLEN'): 0.000431716442108154,
    ('CYCLOPENTENE', 'CYCLOPROPANE'): 0.00181138515472412,
    ('CYCLOPENTENE', 'DMACETYLENE'): 0.000147640705108643, ('CYCLOPENTENE', 'E-BENZENE'): 0.00247997045516968,
    ('CYCLOPENTENE', 'E-MERCAPTAN'): 0.0475529879331589, ('CYCLOPENTENE', 'EACETYLENE'): 0.000134944915771484,
    ('CYCLOPENTENE', 'ECYCLOPENTAN'): 0.00249439477920532, ('CYCLOPENTENE', 'ETHANE'): 0.00351828336715698,
    ('CYCLOPENTENE', 'ETHYLENE'): 0.00569581985473633, ('CYCLOPENTENE', 'H2'): 0.292100012302399,
    ('CYCLOPENTENE', 'H2O'): 0.5, ('CYCLOPENTENE', 'H2S'): 0.0450000017881393,
    ('CYCLOPENTENE', 'I-BUTENE'): 8.70227813720703e-06, ('CYCLOPENTENE', 'IC4'): 6.98566436767578e-05,
    ('CYCLOPENTENE', 'IC5'): 0.000726699829101563, ('CYCLOPENTENE', 'M-ACETYLENE'): 0.00223356485366821,
    ('CYCLOPENTENE', 'M-MERCAPTAN'): 0.0475529879331589, ('CYCLOPENTENE', 'M-XYLENE'): 0.00253421068191528,
    ('CYCLOPENTENE', 'MCYCLOHEXANE'): 0.00229424238204956,
    ('CYCLOPENTENE', 'MCYCLOPENTAN'): 0.000966787338256836, ('CYCLOPENTENE', 'METHANE'): 0.0112966299057007,
    ('CYCLOPENTENE', 'N-PBENZENE'): 0.0047423243522644, ('CYCLOPENTENE', 'NC4'): 2.22325325012207e-05,
    ('CYCLOPENTENE', 'NC5'): 0.000789463520050049, ('CYCLOPENTENE', 'NC6'): 0.00229424238204956,
    ('CYCLOPENTENE', 'NC7'): 0.00423449277877808, ('CYCLOPENTENE', 'NC8'): 0.00648081302642822,
    ('CYCLOPENTENE', 'NITROGEN'): 0.100000001490116, ('CYCLOPENTENE', 'O-XYLENE'): 0.00232869386672974,
    ('CYCLOPENTENE', 'P-XYLENE'): 0.00263965129852295, ('CYCLOPENTENE', 'PROPADIENE'): 0.00237190723419189,
    ('CYCLOPENTENE', 'PROPANE'): 0.000571727752685547, ('CYCLOPENTENE', 'PROPENE'): 0.00127410888671875,
    ('CYCLOPENTENE', 'STYRENE'): 0.00182110071182251, ('CYCLOPENTENE', 'TOLUENE'): 0.000898778438568115,
    ('CYCLOPENTENE', 'TR2-BUTENE'): 1.10864639282227e-05, ('CYCLOPENTENE', 'TR2-HEXENE'): 0.00178998708724976,
    ('CYCLOPENTENE', 'TR2-OCTENE'): 0.00565195083618164,
    ('CYCLOPENTENE', 'TR2-PENTENE'): 0.000608503818511963,
    ('CYCLOPENTENE', 'TR3-HEXENE'): 0.00178998708724976, ('CYCLOPENTENE', 'TR3-OCTENE'): 0.00565195083618164,
    ('CYCLOPENTENE', 'TR4-OCTENE'): 0.00565195083618164,
    ('CYCLOPENTENE', 'VNYLACETLEN'): 0.000189840793609619,
    ('CYCLOPROPANE', 'DMACETYLENE'): 0.000926077365875244, ('CYCLOPROPANE', 'E-BENZENE'): 0.00848525762557983,
    ('CYCLOPROPANE', 'E-MERCAPTAN'): 0.0609394013881683, ('CYCLOPROPANE', 'EACETYLENE'): 0.000958621501922607,
    ('CYCLOPROPANE', 'ECYCLOPENTAN'): 0.00851172208786011, ('CYCLOPROPANE', 'ETHANE'): 0.000282466411590576,
    ('CYCLOPROPANE', 'ETHYLENE'): 0.00109183788299561, ('CYCLOPROPANE', 'H2'): 0.292100012302399,
    ('CYCLOPROPANE', 'H2O'): 0.5, ('CYCLOPROPANE', 'H2S'): 0.0450000017881393,
    ('CYCLOPROPANE', 'I-BUTENE'): 0.00156939029693604, ('CYCLOPROPANE', 'IC4'): 0.00259000062942505,
    ('CYCLOPROPANE', 'IC5'): 0.00481879711151123, ('CYCLOPROPANE', 'M-ACETYLENE'): 2.21729278564453e-05,
    ('CYCLOPROPANE', 'M-MERCAPTAN'): 0.0609394013881683, ('CYCLOPROPANE', 'M-XYLENE'): 0.00858473777770996,
    ('CYCLOPROPANE', 'MCYCLOHEXANE'): 0.00814139842987061,
    ('CYCLOPROPANE', 'MCYCLOPENTAN'): 0.00540679693222046, ('CYCLOPROPANE', 'METHANE'): 0.00410711765289307,
    ('CYCLOPROPANE', 'N-PBENZENE'): 0.0123251676559448, ('CYCLOPROPANE', 'NC4'): 0.00223338603973389,
    ('CYCLOPROPANE', 'NC5'): 0.00497782230377197, ('CYCLOPROPANE', 'NC6'): 0.00814139842987061,
    ('CYCLOPROPANE', 'NC7'): 0.0115051865577698, ('CYCLOPROPANE', 'NC8'): 0.0150160789489746,
    ('CYCLOPROPANE', 'NITROGEN'): 0.100000001490116, ('CYCLOPROPANE', 'O-XYLENE'): 0.00820541381835938,
    ('CYCLOPROPANE', 'P-XYLENE'): 0.00877642631530762, ('CYCLOPROPANE', 'PROPADIENE'): 3.7848949432373e-05,
    ('CYCLOPROPANE', 'PROPANE'): 0.000348687171936035, ('CYCLOPROPANE', 'PROPENE'): 4.73260879516602e-05,
    ('CYCLOPROPANE', 'STYRENE'): 0.00723201036453247, ('CYCLOPROPANE', 'TOLUENE'): 0.0052453875541687,
    ('CYCLOPROPANE', 'TR2-BUTENE'): 0.00153863430023193, ('CYCLOPROPANE', 'TR2-HEXENE'): 0.00717061758041382,
    ('CYCLOPROPANE', 'TR2-OCTENE'): 0.0137525796890259, ('CYCLOPROPANE', 'TR2-PENTENE'): 0.00450778007507324,
    ('CYCLOPROPANE', 'TR3-HEXENE'): 0.00717061758041382, ('CYCLOPROPANE', 'TR3-OCTENE'): 0.0137525796890259,
    ('CYCLOPROPANE', 'TR4-OCTENE'): 0.0137525796890259, ('CYCLOPROPANE', 'VNYLACETLEN'): 0.000829458236694336,
    ('DMACETYLENE', 'E-BENZENE'): 0.00383174419403076, ('DMACETYLENE', 'E-MERCAPTAN'): 0.0559069104492664,
    ('DMACETYLENE', 'EACETYLENE'): 2.98023223876953e-07, ('DMACETYLENE', 'ECYCLOPENTAN'): 0.00384968519210815,
    ('DMACETYLENE', 'ETHANE'): 0.00222861766815186, ('DMACETYLENE', 'ETHYLENE'): 0.00401872396469116,
    ('DMACETYLENE', 'H2'): 0.292100012302399, ('DMACETYLENE', 'H2O'): 0.5,
    ('DMACETYLENE', 'H2S'): 0.0450000017881393, ('DMACETYLENE', 'I-BUTENE'): 8.45789909362793e-05,
    ('DMACETYLENE', 'IC4'): 0.000420331954956055, ('DMACETYLENE', 'IC5'): 0.00152814388275146,
    ('DMACETYLENE', 'M-ACETYLENE'): 0.00123465061187744, ('DMACETYLENE', 'M-MERCAPTAN'): 0.0559069104492664,
    ('DMACETYLENE', 'M-XYLENE'): 0.00389909744262695, ('DMACETYLENE', 'MCYCLOHEXANE'): 0.00360053777694702,
    ('DMACETYLENE', 'MCYCLOPENTAN'): 0.00186800956726074, ('DMACETYLENE', 'METHANE'): 0.00889021158218384,
    ('DMACETYLENE', 'N-PBENZENE'): 0.00654953718185425, ('DMACETYLENE', 'NC4'): 0.000284194946289063,
    ('DMACETYLENE', 'NC5'): 0.00161850452423096, ('DMACETYLENE', 'NC6'): 0.00360053777694702,
    ('DMACETYLENE', 'NC7'): 0.0059514045715332, ('DMACETYLENE', 'NC8'): 0.00856369733810425,
    ('DMACETYLENE', 'NITROGEN'): 0.100000001490116, ('DMACETYLENE', 'O-XYLENE'): 0.0036434531211853,
    ('DMACETYLENE', 'P-XYLENE'): 0.00402939319610596, ('DMACETYLENE', 'PROPADIENE'): 0.00133812427520752,
    ('DMACETYLENE', 'PROPANE'): 0.000138401985168457, ('DMACETYLENE', 'PROPENE'): 0.000554859638214111,
    ('DMACETYLENE', 'STYRENE'): 0.00300157070159912, ('DMACETYLENE', 'TOLUENE'): 0.00177323818206787,
    ('DMACETYLENE', 'TR2-BUTENE'): 7.76052474975586e-05, ('DMACETYLENE', 'TR2-HEXENE'): 0.00296193361282349,
    ('DMACETYLENE', 'TR2-OCTENE'): 0.00760889053344727, ('DMACETYLENE', 'TR2-PENTENE'): 0.00135445594787598,
    ('DMACETYLENE', 'TR3-HEXENE'): 0.00296193361282349, ('DMACETYLENE', 'TR3-OCTENE'): 0.00760889053344727,
    ('DMACETYLENE', 'TR4-OCTENE'): 0.00760889053344727, ('DMACETYLENE', 'VNYLACETLEN'): 2.68220901489258e-06,
    ('E-BENZENE', 'E-MERCAPTAN'): 0.0145937940105796, ('E-BENZENE', 'EACETYLENE'): 0.00376653671264648,
    ('E-BENZENE', 'ECYCLOPENTAN'): 5.96046447753906e-08, ('E-BENZENE', 'ETHANE'): 0.0118186473846436,
    ('E-BENZENE', 'ETHYLENE'): 0.0155462622642517, ('E-BENZENE', 'H2'): 0.292100012302399,
    ('E-BENZENE', 'H2O'): 0.5, ('E-BENZENE', 'H2S'): 0.0450000017881393,
    ('E-BENZENE', 'I-BUTENE'): 0.00278162956237793, ('E-BENZENE', 'IC4'): 0.00171941518783569,
    ('E-BENZENE', 'IC5'): 0.000523507595062256, ('E-BENZENE', 'M-ACETYLENE'): 0.0093654990196228,
    ('E-BENZENE', 'M-MERCAPTAN'): 0.0145937940105796, ('E-BENZENE', 'M-XYLENE'): 2.38418579101563e-07,
    ('E-BENZENE', 'MCYCLOHEXANE'): 3.63588333129883e-06, ('E-BENZENE', 'MCYCLOPENTAN'): 0.000351250171661377,
    ('E-BENZENE', 'METHANE'): 0.024044394493103, ('E-BENZENE', 'N-PBENZENE'): 0.000366628170013428,
    ('E-BENZENE', 'NC4'): 0.00203418731689453, ('E-BENZENE', 'NC5'): 0.000472605228424072,
    ('E-BENZENE', 'NC6'): 3.63588333129883e-06, ('E-BENZENE', 'NC7'): 0.000235199928283691,
    ('E-BENZENE', 'NC8'): -0.00179909996222705, ('E-BENZENE', 'NITROGEN'): 0.100000001490116,
    ('E-BENZENE', 'O-XYLENE'): 2.32458114624023e-06, ('E-BENZENE', 'P-XYLENE'): 2.50339508056641e-06,
    ('E-BENZENE', 'PROPADIENE'): 0.00964409112930298, ('E-BENZENE', 'PROPANE'): 0.0054171085357666,
    ('E-BENZENE', 'PROPENE'): 0.00727689266204834, ('E-BENZENE', 'STYRENE'): 5.0961971282959e-05,
    ('E-BENZENE', 'TOLUENE'): 0.000394284725189209, ('E-BENZENE', 'TR2-BUTENE'): 0.00282281637191772,
    ('E-BENZENE', 'TR2-HEXENE'): 5.64455986022949e-05, ('E-BENZENE', 'TR2-OCTENE'): 0.000650167465209961,
    ('E-BENZENE', 'TR2-PENTENE'): 0.000633478164672852, ('E-BENZENE', 'TR3-HEXENE'): 5.64455986022949e-05,
    ('E-BENZENE', 'TR3-OCTENE'): 0.000650167465209961, ('E-BENZENE', 'TR4-OCTENE'): 0.000650167465209961,
    ('E-BENZENE', 'VNYLACETLEN'): 0.00403553247451782,
    ('E-MERCAPTAN', 'EACETYLENE'): 0.05767822265625, ('E-MERCAPTAN', 'ECYCLOPENTAN'): 0.047463420778513,
    ('E-MERCAPTAN', 'ETHANE'): 0.0607680007815361, ('E-MERCAPTAN', 'ETHYLENE'): 0.0597206763923168,
    ('E-MERCAPTAN', 'I-BUTENE'): 0.0596382468938828, ('E-MERCAPTAN', 'IC4'): 0.0601705983281136,
    ('E-MERCAPTAN', 'IC5'): 0.059373065829277, ('E-MERCAPTAN', 'M-ACETYLENE'): 0.0585366822779179,
    ('E-MERCAPTAN', 'M-XYLENE'): 0.0160421524196863, ('E-MERCAPTAN', 'MCYCLOHEXANE'): 0.0471121594309807,
    ('E-MERCAPTAN', 'MCYCLOPENTAN'): 0.0507104769349098, ('E-MERCAPTAN', 'METHANE'): 0.0591587759554386,
    ('E-MERCAPTAN', 'N-PBENZENE'): 0.0148921068757772, ('E-MERCAPTAN', 'NC4'): 0.0368640013039112,
    ('E-MERCAPTAN', 'NC5'): 0.0592578798532486, ('E-MERCAPTAN', 'NC6'): 0.0580051429569721,
    ('E-MERCAPTAN', 'NC7'): 0.0562547668814659, ('E-MERCAPTAN', 'NC8'): 0.0540067404508591,
    ('E-MERCAPTAN', 'O-XYLENE'): 0.00548804644495249, ('E-MERCAPTAN', 'P-XYLENE'): 0.017520809546113,
    ('E-MERCAPTAN', 'PROPADIENE'): 0.0569896809756756, ('E-MERCAPTAN', 'PROPANE'): 0.0564489997923374,
    ('E-MERCAPTAN', 'PROPENE'): 0.0599718652665615, ('E-MERCAPTAN', 'STYRENE'): -0.00999999977648258,
    ('E-MERCAPTAN', 'TOLUENE'): 0.0167241860181093, ('E-MERCAPTAN', 'TR2-BUTENE'): 0.0593789592385292,
    ('E-MERCAPTAN', 'TR2-HEXENE'): 0.056970052421093, ('E-MERCAPTAN', 'TR2-OCTENE'): 0.0523778945207596,
    ('E-MERCAPTAN', 'TR2-PENTENE'): 0.0583673492074013, ('E-MERCAPTAN', 'TR3-HEXENE'): 0.057012852281332,
    ('E-MERCAPTAN', 'TR3-OCTENE'): 0.0529257394373417, ('E-MERCAPTAN', 'TR4-OCTENE'): 0.0530732236802578,
    ('E-MERCAPTAN', 'VNYLACETLEN'): 0.0560894645750523,
    ('EACETYLENE', 'ECYCLOPENTAN'): 0.00378423929214478, ('EACETYLENE', 'ETHANE'): 0.00227898359298706,
    ('EACETYLENE', 'ETHYLENE'): 0.00408607721328735, ('EACETYLENE', 'H2'): 0.292100012302399,
    ('EACETYLENE', 'H2O'): 0.5, ('EACETYLENE', 'H2S'): 0.0450000017881393,
    ('EACETYLENE', 'I-BUTENE'): 7.50422477722168e-05, ('EACETYLENE', 'IC4'): 0.000398814678192139,
    ('EACETYLENE', 'IC5'): 0.00148683786392212, ('EACETYLENE', 'M-ACETYLENE'): 0.00127226114273071,
    ('EACETYLENE', 'M-MERCAPTAN'): 0.05767822265625, ('EACETYLENE', 'M-XYLENE'): 0.00383329391479492,
    ('EACETYLENE', 'MCYCLOHEXANE'): 0.00353729724884033, ('EACETYLENE', 'MCYCLOPENTAN'): 0.00182241201400757,
    ('EACETYLENE', 'METHANE'): 0.00898987054824829, ('EACETYLENE', 'N-PBENZENE'): 0.00646436214447021,
    ('EACETYLENE', 'NC4'): 0.000266492366790771, ('EACETYLENE', 'NC5'): 0.00157606601715088,
    ('EACETYLENE', 'NC6'): 0.00353729724884033, ('EACETYLENE', 'NC7'): 0.00587016344070435,
    ('EACETYLENE', 'NC8'): 0.00846654176712036, ('EACETYLENE', 'NITROGEN'): 0.100000001490116,
    ('EACETYLENE', 'O-XYLENE'): 0.00357979536056519, ('EACETYLENE', 'P-XYLENE'): 0.00396245718002319,
    ('EACETYLENE', 'PROPADIENE'): 0.00137728452682495, ('EACETYLENE', 'PROPANE'): 0.000151216983795166,
    ('EACETYLENE', 'PROPENE'): 0.000580310821533203, ('EACETYLENE', 'STYRENE'): 0.00294375419616699,
    ('EACETYLENE', 'TOLUENE'): 0.00172877311706543, ('EACETYLENE', 'TR2-BUTENE'): 6.84857368469238e-05,
    ('EACETYLENE', 'TR2-HEXENE'): 0.00290447473526001, ('EACETYLENE', 'TR2-OCTENE'): 0.00751721858978271,
    ('EACETYLENE', 'TR2-PENTENE'): 0.0013156533241272, ('EACETYLENE', 'TR3-HEXENE'): 0.00290447473526001,
    ('EACETYLENE', 'TR3-OCTENE'): 0.00751721858978271, ('EACETYLENE', 'TR4-OCTENE'): 0.00751721858978271,
    ('EACETYLENE', 'VNYLACETLEN'): 4.64916229248047e-06,
    ('ECYCLOPENTAN', 'ETHANE'): 0.0118497014045715, ('ECYCLOPENTAN', 'ETHYLENE'): 0.0155817270278931,
    ('ECYCLOPENTAN', 'H2'): 0.292100012302399, ('ECYCLOPENTAN', 'H2O'): 0.5,
    ('ECYCLOPENTAN', 'H2S'): 0.0450000017881393, ('ECYCLOPENTAN', 'I-BUTENE'): 0.00279688835144043,
    ('ECYCLOPENTAN', 'IC4'): 0.00173145532608032, ('ECYCLOPENTAN', 'IC5'): 0.0005301833152771,
    ('ECYCLOPENTAN', 'M-ACETYLENE'): 0.00939321517944336, ('ECYCLOPENTAN', 'M-MERCAPTAN'): 0.047463420778513,
    ('ECYCLOPENTAN', 'M-XYLENE'): 1.78813934326172e-07,
    ('ECYCLOPENTAN', 'MCYCLOHEXANE'): 4.23192977905273e-06,
    ('ECYCLOPENTAN', 'MCYCLOPENTAN'): 0.000356793403625488, ('ECYCLOPENTAN', 'METHANE'): 0.0240878462791443,
    ('ECYCLOPENTAN', 'N-PBENZENE'): 0.000361084938049316, ('ECYCLOPENTAN', 'NC4'): 0.00204730033874512,
    ('ECYCLOPENTAN', 'NC5'): 0.000478863716125488, ('ECYCLOPENTAN', 'NC6'): 4.23192977905273e-06,
    ('ECYCLOPENTAN', 'NC7'): 0.000230669975280762, ('ECYCLOPENTAN', 'NC8'): 0.000943362712860107,
    ('ECYCLOPENTAN', 'NITROGEN'): 0.100000001490116, ('ECYCLOPENTAN', 'O-XYLENE'): 2.80141830444336e-06,
    ('ECYCLOPENTAN', 'P-XYLENE'): 2.08616256713867e-06, ('ECYCLOPENTAN', 'PROPADIENE'): 0.00967222452163696,
    ('ECYCLOPENTAN', 'PROPANE'): 0.00543826818466187, ('ECYCLOPENTAN', 'PROPENE'): 0.00730133056640625,
    ('ECYCLOPENTAN', 'STYRENE'): 5.3107738494873e-05, ('ECYCLOPENTAN', 'TOLUENE'): 0.000400066375732422,
    ('ECYCLOPENTAN', 'TR2-BUTENE'): 0.00283819437026978, ('ECYCLOPENTAN', 'TR2-HEXENE'): 5.8591365814209e-05,
    ('ECYCLOPENTAN', 'TR2-OCTENE'): 0.000642776489257813, ('ECYCLOPENTAN', 'TR2-PENTENE'): 0.000640869140625,
    ('ECYCLOPENTAN', 'TR3-HEXENE'): 5.8591365814209e-05, ('ECYCLOPENTAN', 'TR3-OCTENE'): 0.000642776489257813,
    ('ECYCLOPENTAN', 'TR4-OCTENE'): 0.000642776489257813,
    ('ECYCLOPENTAN', 'VNYLACETLEN'): 0.00405395030975342,
    ('ETHANE', 'ETHYLENE'): 0.0122990002855659, ('ETHANE', 'H2'): 0.223100006580353, ('ETHANE', 'H2O'): 0.5,
    ('ETHANE', 'H2S'): 0.0839990004897118, ('ETHANE', 'HELIUM'): 1.12320005893707,
    ('ETHANE', 'I-BUTENE'): 0.00317829847335815, ('ETHANE', 'IC4'): 0.00457346439361572,
    ('ETHANE', 'IC5'): 0.00741332769393921, ('ETHANE', 'M-ACETYLENE'): 0.000146448612213135,
    ('ETHANE', 'M-MERCAPTAN'): 0.0693709999322891, ('ETHANE', 'M-XYLENE'): 0.0295000001788139,
    ('ETHANE', 'MCYCLOHEXANE'): 0.0121990004554391, ('ETHANE', 'MCYCLOPENTAN'): 0.00813591480255127,
    ('ETHANE', 'METHANE'): 0.00224137306213379, ('ETHANE', 'N-PBENZENE'): 0.0379990004003048,
    ('ETHANE', 'NC4'): 0.0040963888168335, ('ETHANE', 'NC5'): 0.00760942697525024,
    ('ETHANE', 'NC6'): 0.0114138126373291, ('ETHANE', 'NC7'): 0.0153242945671082,
    ('ETHANE', 'NC8'): 0.019318699836731, ('ETHANE', 'NITROGEN'): 0.0500000007450581,
    ('ETHANE', 'O-XYLENE'): 0.0329990014433861, ('ETHANE', 'P-XYLENE'): 0.0329990014433861,
    ('ETHANE', 'PROPADIENE'): 0.000113427639007568, ('ETHANE', 'PROPANE'): 0.00125795602798462,
    ('ETHANE', 'PROPENE'): -0.0019000000320375, ('ETHANE', 'STYRENE'): 0.0103362202644348,
    ('ETHANE', 'TOLUENE'): 0.034400001168251, ('ETHANE', 'TR2-BUTENE'): 0.00313454866409302,
    ('ETHANE', 'TR2-HEXENE'): 0.0102631449699402, ('ETHANE', 'TR2-OCTENE'): 0.0178892016410828,
    ('ETHANE', 'TR2-PENTENE'): 0.00702744722366333, ('ETHANE', 'TR3-HEXENE'): 0.0102631449699402,
    ('ETHANE', 'TR3-OCTENE'): 0.0178892016410828, ('ETHANE', 'TR4-OCTENE'): 0.0178892016410828,
    ('ETHANE', 'VNYLACETLEN'): 0.00207757949829102,
    ('ETHYLENE', 'H2'): 0.00740000000223517, ('ETHYLENE', 'H2O'): 0.5,
    ('ETHYLENE', 'H2S'): 0.0500000007450581, ('ETHYLENE', 'I-BUTENE'): 0.00526183843612671,
    ('ETHYLENE', 'IC4'): 0.00701570510864258, ('ETHYLENE', 'IC5'): 0.0104383230209351,
    ('ETHYLENE', 'M-ACETYLENE'): 0.000802993774414063, ('ETHYLENE', 'M-MERCAPTAN'): 0.0597206763923168,
    ('ETHYLENE', 'M-XYLENE'): 0.0156795978546143, ('ETHYLENE', 'MCYCLOHEXANE'): 0.0150836110115051,
    ('ETHYLENE', 'MCYCLOPENTAN'): 0.0112893581390381, ('ETHYLENE', 'METHANE'): 0.0215000007301569,
    ('ETHYLENE', 'N-PBENZENE'): 0.0205651521682739, ('ETHYLENE', 'NC4'): 0.0939980000257492,
    ('ETHYLENE', 'NC5'): 0.0106698870658875, ('ETHYLENE', 'NC6'): 0.0150836110115051,
    ('ETHYLENE', 'NC7'): 0.017000000923872, ('ETHYLENE', 'NC8'): 0.0239636898040771,
    ('ETHYLENE', 'NITROGEN'): 0.0722000002861023, ('ETHYLENE', 'O-XYLENE'): 0.0151700377464294,
    ('ETHYLENE', 'P-XYLENE'): 0.0159359574317932, ('ETHYLENE', 'PROPADIENE'): 0.000723183155059814,
    ('ETHYLENE', 'PROPANE'): 0.00267040729522705, ('ETHYLENE', 'PROPENE'): 0.00159275531768799,
    ('ETHYLENE', 'STYRENE'): 0.0138463973999023, ('ETHYLENE', 'TOLUENE'): 0.0099999001249671,
    ('ETHYLENE', 'TR2-BUTENE'): 0.00520581007003784, ('ETHYLENE', 'TR2-HEXENE'): 0.0137620568275452,
    ('ETHYLENE', 'TR2-OCTENE'): 0.022377610206604, ('ETHYLENE', 'TR2-PENTENE'): 0.00998121500015259,
    ('ETHYLENE', 'TR3-HEXENE'): 0.0137620568275452, ('ETHYLENE', 'TR3-OCTENE'): 0.022377610206604,
    ('ETHYLENE', 'TR4-OCTENE'): 0.022377610206604, ('ETHYLENE', 'VNYLACETLEN'): 0.0038154125213623,
    ('H2', 'H2O'): -0.299800008535385, ('H2', 'H2S'): 0.75, ('H2', 'HELIUM'): 0.405800014734268,
    ('H2', 'I-BUTENE'): 0.292100012302399, ('H2', 'IC4'): 0.203700006008148, ('H2', 'IC5'): 0.292100012302399,
    ('H2', 'M-ACETYLENE'): 0.292100012302399, ('H2', 'M-XYLENE'): 0.292100012302399,
    ('H2', 'MCYCLOHEXANE'): 0.292100012302399, ('H2', 'MCYCLOPENTAN'): 0.292100012302399,
    ('H2', 'METHANE'): 0.202000007033348, ('H2', 'N-PBENZENE'): 0.292100012302399,
    ('H2', 'NC4'): 0.194100007414818, ('H2', 'NC5'): 0.292100012302399, ('H2', 'NC6'): 0.292100012302399,
    ('H2', 'NC7'): -0.116700001060963, ('H2', 'NC8'): 0.292100012302399,
    ('H2', 'NITROGEN'): -0.0359990000724792, ('H2', 'O-XYLENE'): 0.292100012302399,
    ('H2', 'P-XYLENE'): 0.292100012302399, ('H2', 'PROPADIENE'): 0.292100012302399,
    ('H2', 'PROPANE'): 0.214200004935265, ('H2', 'PROPENE'): -0.10360000282526,
    ('H2', 'STYRENE'): 0.292100012302399, ('H2', 'TOLUENE'): 0.285100013017654,
    ('H2', 'TR2-BUTENE'): 0.292100012302399, ('H2', 'TR2-HEXENE'): 0.292100012302399,
    ('H2', 'TR2-OCTENE'): 0.292100012302399, ('H2', 'TR2-PENTENE'): 0.292100012302399,
    ('H2', 'TR3-HEXENE'): 0.292100012302399, ('H2', 'TR3-OCTENE'): 0.292100012302399,
    ('H2', 'TR4-OCTENE'): 0.292100012302399, ('H2', 'VNYLACETLEN'): 0.292100012302399,
    ('H2O', 'H2S'): 0.0819000005722046, ('H2O', 'I-BUTENE'): 0.5, ('H2O', 'IC4'): 0.5, ('H2O', 'IC5'): 0.5,
    ('H2O', 'M-ACETYLENE'): 0.5, ('H2O', 'M-XYLENE'): 0.5, ('H2O', 'MCYCLOHEXANE'): 0.5,
    ('H2O', 'MCYCLOPENTAN'): 0.5, ('H2O', 'METHANE'): 0.5, ('H2O', 'N-PBENZENE'): 0.5, ('H2O', 'NC4'): 0.5,
    ('H2O', 'NC5'): 0.479999989271164, ('H2O', 'NC6'): 0.5, ('H2O', 'NC7'): 0.5, ('H2O', 'NC8'): 0.5,
    ('H2O', 'NH3'): -0.25330001115799, ('H2O', 'NITROGEN'): -0.315600007772446, ('H2O', 'O-XYLENE'): 0.5,
    ('H2O', 'P-XYLENE'): 0.5, ('H2O', 'PROPADIENE'): 0.5, ('H2O', 'PROPANE'): 0.5, ('H2O', 'PROPENE'): 0.5,
    ('H2O', 'SO2'): -0.114090003073215, ('H2O', 'STYRENE'): 0.5, ('H2O', 'TOLUENE'): 0.5,
    ('H2O', 'TR2-BUTENE'): 0.5, ('H2O', 'TR2-HEXENE'): 0.5, ('H2O', 'TR2-OCTENE'): 0.5,
    ('H2O', 'TR2-PENTENE'): 0.5, ('H2O', 'TR3-HEXENE'): 0.5, ('H2O', 'TR3-OCTENE'): 0.5,
    ('H2O', 'TR4-OCTENE'): 0.5, ('H2O', 'VNYLACETLEN'): 0.5,
    ('H2S', 'I-BUTENE'): 0.0450000017881393, ('H2S', 'IC4'): 0.0500000007450581,
    ('H2S', 'IC5'): 0.0599990002810955, ('H2S', 'M-ACETYLENE'): 0.0450000017881393,
    ('H2S', 'M-XYLENE'): 0.0171000007539988, ('H2S', 'MCYCLOHEXANE'): 0.0689980015158653,
    ('H2S', 'MCYCLOPENTAN'): 0.0450000017881393, ('H2S', 'METHANE'): 0.0850000008940697,
    ('H2S', 'N-PBENZENE'): 0.0399990007281303, ('H2S', 'NC4'): 0.0599990002810955,
    ('H2S', 'NC5'): 0.0649980008602142, ('H2S', 'NC6'): 0.0599990002810955,
    ('H2S', 'NC7'): 0.0599990002810955, ('H2S', 'NC8'): 0.0549990013241768,
    ('H2S', 'NITROGEN'): 0.167600005865097, ('H2S', 'O-XYLENE'): -0.0230960007756948,
    ('H2S', 'P-XYLENE'): -0.0230960007756948, ('H2S', 'PROPADIENE'): 0.0450000017881393,
    ('H2S', 'PROPANE'): 0.0750000029802322, ('H2S', 'PROPENE'): 0.0450000017881393,
    ('H2S', 'STYRENE'): 0.0450000017881393, ('H2S', 'TOLUENE'): 0.00810000021010637,
    ('H2S', 'TR2-BUTENE'): 0.0450000017881393, ('H2S', 'TR2-HEXENE'): 0.0450000017881393,
    ('H2S', 'TR2-OCTENE'): 0.0450000017881393, ('H2S', 'TR2-PENTENE'): 0.0450000017881393,
    ('H2S', 'TR3-HEXENE'): 0.0450000017881393, ('H2S', 'TR3-OCTENE'): 0.0450000017881393,
    ('H2S', 'TR4-OCTENE'): 0.0450000017881393, ('H2S', 'VNYLACETLEN'): 0.0450000017881393,
    ('HELIUM', 'METHANE'): 0.764900028705597, ('HELIUM', 'NITROGEN'): 0.0684980005025864,
    ('HELIUM', 'PROPANE'): 1.06420004367828,
    ('I-BUTENE', 'IC4'): 0.00570000009611249, ('I-BUTENE', 'IC5'): 0.000894367694854736,
    ('I-BUTENE', 'M-ACETYLENE'): 0.00196385383605957, ('I-BUTENE', 'M-MERCAPTAN'): 0.0596382468938828,
    ('I-BUTENE', 'M-XYLENE'): 0.00283908843994141, ('I-BUTENE', 'MCYCLOHEXANE'): 0.00258487462997437,
    ('I-BUTENE', 'MCYCLOPENTAN'): 0.00115871429443359, ('I-BUTENE', 'METHANE'): 0.0106862187385559,
    ('I-BUTENE', 'N-PBENZENE'): 0.00515490770339966, ('I-BUTENE', 'NC4'): 5.87105751037598e-05,
    ('I-BUTENE', 'NC5'): 0.000963926315307617, ('I-BUTENE', 'NC6'): 0.00258487462997437,
    ('I-BUTENE', 'NC7'): 0.00462502241134644, ('I-BUTENE', 'NC8'): 0.00696069002151489,
    ('I-BUTENE', 'NITROGEN'): 0.100000001490116, ('I-BUTENE', 'O-XYLENE'): 0.00262123346328735,
    ('I-BUTENE', 'P-XYLENE'): 0.00295060873031616, ('I-BUTENE', 'PROPADIENE'): 0.00209391117095947,
    ('I-BUTENE', 'PROPANE'): 0.00043940544128418, ('I-BUTENE', 'PROPENE'): 0.00107234716415405,
    ('I-BUTENE', 'STYRENE'): 0.00208097696304321, ('I-BUTENE', 'TOLUENE'): 0.00108426809310913,
    ('I-BUTENE', 'TR2-BUTENE'): 1.78813934326172e-07, ('I-BUTENE', 'TR2-HEXENE'): 0.00204789638519287,
    ('I-BUTENE', 'TR2-OCTENE'): 0.00610101222991943, ('I-BUTENE', 'TR2-PENTENE'): 0.000762701034545898,
    ('I-BUTENE', 'TR3-HEXENE'): 0.00204789638519287, ('I-BUTENE', 'TR3-OCTENE'): 0.00610101222991943,
    ('I-BUTENE', 'TR4-OCTENE'): 0.00610101222991943, ('I-BUTENE', 'VNYLACETLEN'): 0.000117301940917969,
    ('IC4', 'IC5'): 0.000346183776855469, ('IC4', 'M-ACETYLENE'): 0.0750000029802322,
    ('IC4', 'M-MERCAPTAN'): 0.0399000011384487, ('IC4', 'M-XYLENE'): 0.00176465511322021,
    ('IC4', 'MCYCLOHEXANE'): 0.00156527757644653, ('IC4', 'MCYCLOPENTAN'): 0.000517368316650391,
    ('IC4', 'METHANE'): 0.0131134390830994, ('IC4', 'N-PBENZENE'): 0.00366675853729248,
    ('IC4', 'NC4'): 1.33514404296875e-05, ('IC4', 'NC5'): 0.000390052795410156,
    ('IC4', 'NC6'): 0.00156527757644653, ('IC4', 'NC7'): 0.00322127342224121,
    ('IC4', 'NC8'): 0.0052141547203064, ('IC4', 'NITROGEN'): 0.0949990004301071,
    ('IC4', 'O-XYLENE'): 0.00159376859664917, ('IC4', 'P-XYLENE'): 0.00185292959213257,
    ('IC4', 'PROPADIENE'): 0.0359990000724792, ('IC4', 'PROPANE'): 0.00104051828384399,
    ('IC4', 'PROPENE'): -0.0149999996647239, ('IC4', 'STYRENE'): 0.00117886066436768,
    ('IC4', 'TOLUENE'): 0.000467896461486816, ('IC4', 'TR2-BUTENE'): 0.000136733055114746,
    ('IC4', 'TR2-HEXENE'): 0.00115388631820679, ('IC4', 'TR2-OCTENE'): 0.00447249412536621,
    ('IC4', 'TR2-PENTENE'): 0.000266194343566895, ('IC4', 'TR3-HEXENE'): 0.00115388631820679,
    ('IC4', 'TR3-OCTENE'): 0.00447249412536621, ('IC4', 'TR4-OCTENE'): 0.00447249412536621,
    ('IC4', 'VNYLACETLEN'): 0.000489890575408936,
    ('IC5', 'M-ACETYLENE'): 0.00549077987670898, ('IC5', 'M-MERCAPTAN'): 0.059373065829277,
    ('IC5', 'M-XYLENE'): 0.000548720359802246, ('IC5', 'MCYCLOHEXANE'): 0.000439941883087158,
    ('IC5', 'MCYCLOPENTAN'): 1.71065330505371e-05, ('IC5', 'METHANE'): 0.017627477645874,
    ('IC5', 'N-PBENZENE'): 0.00176441669464111, ('IC5', 'NC4'): 0.00049513578414917,
    ('IC5', 'NC5'): 1.2516975402832e-06, ('IC5', 'NC6'): 0.000439941883087158,
    ('IC5', 'NC7'): 0.00145912170410156, ('IC5', 'NC8'): 0.00288277864456177,
    ('IC5', 'NITROGEN'): 0.0949990004301071, ('IC5', 'O-XYLENE'): 0.000455141067504883,
    ('IC5', 'P-XYLENE'): 0.000598490238189697, ('IC5', 'PROPADIENE'): 0.00570565462112427,
    ('IC5', 'PROPANE'): 0.00258338451385498, ('IC5', 'PROPENE'): 0.00391578674316406,
    ('IC5', 'STYRENE'): 0.000247776508331299, ('IC5', 'TOLUENE'): 9.11951065063477e-06,
    ('IC5', 'TR2-BUTENE'): 0.00091785192489624, ('IC5', 'TR2-HEXENE'): 0.000236332416534424,
    ('IC5', 'TR2-OCTENE'): 0.00233715772628784, ('IC5', 'TR2-PENTENE'): 5.30481338500977e-06,
    ('IC5', 'TR3-HEXENE'): 0.000236332416534424, ('IC5', 'TR3-OCTENE'): 0.00233715772628784,
    ('IC5', 'TR4-OCTENE'): 0.00233715772628784, ('IC5', 'VNYLACETLEN'): 0.00165802240371704,
    ('M-ACETYLENE', 'M-MERCAPTAN'): 0.0585366822779179, ('M-ACETYLENE', 'M-XYLENE'): 0.00946974754333496,
    ('M-ACETYLENE', 'MCYCLOHEXANE'): 0.00900429487228394,
    ('M-ACETYLENE', 'MCYCLOPENTAN'): 0.00611668825149536, ('M-ACETYLENE', 'METHANE'): 0.00352805852890015,
    ('M-ACETYLENE', 'N-PBENZENE'): 0.0133762359619141, ('M-ACETYLENE', 'NC4'): 0.00269949436187744,
    ('M-ACETYLENE', 'NC5'): 0.00566011667251587, ('M-ACETYLENE', 'NC6'): 0.00900429487228394,
    ('M-ACETYLENE', 'NC7'): 0.0125223398208618, ('M-ACETYLENE', 'NC8'): 0.0161698460578918,
    ('M-ACETYLENE', 'NITROGEN'): 0.100000001490116, ('M-ACETYLENE', 'O-XYLENE'): 0.0090717077255249,
    ('M-ACETYLENE', 'P-XYLENE'): 0.00967073440551758, ('M-ACETYLENE', 'PROPADIENE'): 2.08616256713867e-06,
    ('M-ACETYLENE', 'PROPANE'): 0.0750000029802322, ('M-ACETYLENE', 'PROPENE'): 0.0399990007281303,
    ('M-ACETYLENE', 'STYRENE'): 0.00804758071899414, ('M-ACETYLENE', 'TOLUENE'): 0.00594502687454224,
    ('M-ACETYLENE', 'TR2-BUTENE'): 0.00192958116531372, ('M-ACETYLENE', 'TR2-HEXENE'): 0.00798279047012329,
    ('M-ACETYLENE', 'TR2-OCTENE'): 0.0148594379425049, ('M-ACETYLENE', 'TR2-PENTENE'): 0.00515866279602051,
    ('M-ACETYLENE', 'TR3-HEXENE'): 0.00798279047012329, ('M-ACETYLENE', 'TR3-OCTENE'): 0.0148594379425049,
    ('M-ACETYLENE', 'TR4-OCTENE'): 0.0148594379425049, ('M-ACETYLENE', 'VNYLACETLEN'): 0.00112259387969971,
    ('M-MERCAPTAN', 'M-XYLENE'): 0.0160421524196863, ('M-MERCAPTAN', 'MCYCLOHEXANE'): 0.0471121594309807,
    ('M-MERCAPTAN', 'MCYCLOPENTAN'): 0.0507104769349098, ('M-MERCAPTAN', 'METHANE'): 0.0591587759554386,
    ('M-MERCAPTAN', 'N-PBENZENE'): 0.0148921068757772, ('M-MERCAPTAN', 'NC4'): 0.0719690024852753,
    ('M-MERCAPTAN', 'NC5'): 0.0592578798532486, ('M-MERCAPTAN', 'NC6'): 0.0580051429569721,
    ('M-MERCAPTAN', 'NC7'): 0.0562547668814659, ('M-MERCAPTAN', 'NC8'): 0.0540067404508591,
    ('M-MERCAPTAN', 'O-XYLENE'): 0.00548804644495249, ('M-MERCAPTAN', 'P-XYLENE'): 0.017520809546113,
    ('M-MERCAPTAN', 'PROPADIENE'): 0.0569896809756756, ('M-MERCAPTAN', 'PROPANE'): 0.0656580030918121,
    ('M-MERCAPTAN', 'PROPENE'): 0.0599718652665615, ('M-MERCAPTAN', 'STYRENE'): -0.00999999977648258,
    ('M-MERCAPTAN', 'TOLUENE'): 0.0167241860181093, ('M-MERCAPTAN', 'TR2-BUTENE'): 0.0593789592385292,
    ('M-MERCAPTAN', 'TR2-HEXENE'): 0.056970052421093, ('M-MERCAPTAN', 'TR2-OCTENE'): 0.0523778945207596,
    ('M-MERCAPTAN', 'TR2-PENTENE'): 0.0583673492074013, ('M-MERCAPTAN', 'TR3-HEXENE'): 0.057012852281332,
    ('M-MERCAPTAN', 'TR3-OCTENE'): 0.0529257394373417, ('M-MERCAPTAN', 'TR4-OCTENE'): 0.0530732236802578,
    ('M-MERCAPTAN', 'VNYLACETLEN'): 0.0560894645750523,
    ('M-XYLENE', 'MCYCLOHEXANE'): 6.07967376708984e-06, ('M-XYLENE', 'MCYCLOPENTAN'): 0.000371992588043213,
    ('M-XYLENE', 'METHANE'): 0.049100000411272, ('M-XYLENE', 'N-PBENZENE'): 0.000346064567565918,
    ('M-XYLENE', 'NC4'): 0.002083420753479, ('M-XYLENE', 'NC5'): 0.000496506690979004,
    ('M-XYLENE', 'NC6'): 6.07967376708984e-06, ('M-XYLENE', 'NC7'): 0.000218749046325684,
    ('M-XYLENE', 'NC8'): 0.000919103622436523, ('M-XYLENE', 'NITROGEN'): 0.216900005936623,
    ('M-XYLENE', 'O-XYLENE'): 4.41074371337891e-06, ('M-XYLENE', 'P-XYLENE'): 1.07288360595703e-06,
    ('M-XYLENE', 'PROPADIENE'): 0.0097498893737793, ('M-XYLENE', 'PROPANE'): 0.0299980007112026,
    ('M-XYLENE', 'PROPENE'): 0.00736904144287109, ('M-XYLENE', 'STYRENE'): 5.91278076171875e-05,
    ('M-XYLENE', 'TOLUENE'): 0.000416159629821777, ('M-XYLENE', 'TR2-BUTENE'): 0.00288069248199463,
    ('M-XYLENE', 'TR2-HEXENE'): 6.49094581604004e-05, ('M-XYLENE', 'TR2-OCTENE'): 0.000622749328613281,
    ('M-XYLENE', 'TR2-PENTENE'): 0.000661194324493408, ('M-XYLENE', 'TR3-HEXENE'): 6.49094581604004e-05,
    ('M-XYLENE', 'TR3-OCTENE'): 0.000622749328613281, ('M-XYLENE', 'TR4-OCTENE'): 0.000622749328613281,
    ('M-XYLENE', 'VNYLACETLEN'): 0.00410455465316772,
    ('MCYCLOHEXANE', 'MCYCLOPENTAN'): 0.000283479690551758, ('MCYCLOHEXANE', 'METHANE'): 0.00650000013411045,
    ('MCYCLOHEXANE', 'N-PBENZENE'): 0.000443220138549805, ('MCYCLOHEXANE', 'NC4'): 0.00186634063720703,
    ('MCYCLOHEXANE', 'NC5'): 0.000899980019312352, ('MCYCLOHEXANE', 'NC6'): -0.000699990021530539,
    ('MCYCLOHEXANE', 'NC7'): 0.000297248363494873, ('MCYCLOHEXANE', 'NC8'): 0.00107330083847046,
    ('MCYCLOHEXANE', 'NITROGEN'): 0.106200002133846, ('MCYCLOHEXANE', 'O-XYLENE'): 1.78813934326172e-07,
    ('MCYCLOHEXANE', 'P-XYLENE'): 1.21593475341797e-05, ('MCYCLOHEXANE', 'PROPADIENE'): 0.00927770137786865,
    ('MCYCLOHEXANE', 'PROPANE'): 0.0124000003561378, ('MCYCLOHEXANE', 'PROPENE'): 0.0069580078125,
    ('MCYCLOHEXANE', 'STYRENE'): 2.74777412414551e-05, ('MCYCLOHEXANE', 'TOLUENE'): 0.000322282314300537,
    ('MCYCLOHEXANE', 'TR2-BUTENE'): 0.00262457132339478, ('MCYCLOHEXANE', 'TR2-HEXENE'): 3.14116477966309e-05,
    ('MCYCLOHEXANE', 'TR2-OCTENE'): 0.000750899314880371,
    ('MCYCLOHEXANE', 'TR2-PENTENE'): 0.000541329383850098,
    ('MCYCLOHEXANE', 'TR3-HEXENE'): 3.14116477966309e-05,
    ('MCYCLOHEXANE', 'TR3-OCTENE'): 0.000750899314880371,
    ('MCYCLOHEXANE', 'TR4-OCTENE'): 0.000750899314880371,
    ('MCYCLOHEXANE', 'VNYLACETLEN'): 0.00379818677902222,
    ('MCYCLOPENTAN', 'METHANE'): 0.0187180638313293, ('MCYCLOPENTAN', 'N-PBENZENE'): 0.00143438577651978,
    ('MCYCLOPENTAN', 'NC4'): 0.000696241855621338, ('MCYCLOPENTAN', 'NC5'): 0.00109929998870939,
    ('MCYCLOPENTAN', 'NC6'): -0.00230000005103648, ('MCYCLOPENTAN', 'NC7'): 0.00116044282913208,
    ('MCYCLOPENTAN', 'NC8'): 0.00245684385299683, ('MCYCLOPENTAN', 'NITROGEN'): 0.100000001490116,
    ('MCYCLOPENTAN', 'O-XYLENE'): 0.000295698642730713, ('MCYCLOPENTAN', 'P-XYLENE'): 0.000413119792938232,
    ('MCYCLOPENTAN', 'PROPADIENE'): 0.00634312629699707, ('MCYCLOPENTAN', 'PROPANE'): 0.00301980972290039,
    ('MCYCLOPENTAN', 'PROPENE'): 0.00444817543029785, ('MCYCLOPENTAN', 'STYRENE'): 0.000134527683258057,
    ('MCYCLOPENTAN', 'TOLUENE'): 1.2516975402832e-06, ('MCYCLOPENTAN', 'TR2-BUTENE'): 0.00118541717529297,
    ('MCYCLOPENTAN', 'TR2-HEXENE'): 0.000126183032989502, ('MCYCLOPENTAN', 'TR2-OCTENE'): 0.00195497274398804,
    ('MCYCLOPENTAN', 'TR2-PENTENE'): 4.13656234741211e-05,
    ('MCYCLOPENTAN', 'TR3-HEXENE'): 0.000126183032989502, ('MCYCLOPENTAN', 'TR3-OCTENE'): 0.00195497274398804,
    ('MCYCLOPENTAN', 'TR4-OCTENE'): 0.00195497274398804, ('MCYCLOPENTAN', 'VNYLACETLEN'): 0.00201147794723511,
    ('METHANE', 'N-PBENZENE'): 0.0700000002980232, ('METHANE', 'N2O'): 0.0299980007112026,
    ('METHANE', 'NC4'): 0.0123046040534973, ('METHANE', 'NC5'): 0.0179253816604614,
    ('METHANE', 'NC6'): 0.0234740376472473, ('METHANE', 'NC7'): 0.0288642644882202,
    ('METHANE', 'NC8'): 0.0341590642929077, ('METHANE', 'NITROGEN'): 0.0359990000724792,
    ('METHANE', 'O-XYLENE'): 0.0500000007450581, ('METHANE', 'P-XYLENE'): 0.0500000007450581,
    ('METHANE', 'PROPADIENE'): 0.00335907936096191, ('METHANE', 'PROPANE'): 0.00682878494262695,
    ('METHANE', 'PROPENE'): 0.0329990014433861, ('METHANE', 'SO2'): 0.134200006723404,
    ('METHANE', 'STYRENE'): 0.0219399929046631, ('METHANE', 'TOLUENE'): 0.0649000033736229,
    ('METHANE', 'TR2-BUTENE'): 0.0106067657470703, ('METHANE', 'TR2-HEXENE'): 0.0218349695205688,
    ('METHANE', 'TR2-OCTENE'): 0.0322840809822083, ('METHANE', 'TR2-PENTENE'): 0.0170372128486633,
    ('METHANE', 'TR3-HEXENE'): 0.0218349695205688, ('METHANE', 'TR3-OCTENE'): 0.0322840809822083,
    ('METHANE', 'TR4-OCTENE'): 0.0322840809822083, ('METHANE', 'VNYLACETLEN'): 0.00858843326568604,
    ('N-PBENZENE', 'NC4'): 0.00411909818649292, ('N-PBENZENE', 'NC5'): 0.00166988372802734,
    ('N-PBENZENE', 'NC6'): 0.000443220138549805, ('N-PBENZENE', 'NC7'): 1.44839286804199e-05,
    ('N-PBENZENE', 'NC8'): 0.000137388706207275, ('N-PBENZENE', 'NITROGEN'): 0.319900006055832,
    ('N-PBENZENE', 'O-XYLENE'): 0.000428199768066406, ('N-PBENZENE', 'P-XYLENE'): 0.000308454036712646,
    ('N-PBENZENE', 'PROPADIENE'): 0.0137070417404175, ('N-PBENZENE', 'PROPANE'): 0.0260000005364418,
    ('N-PBENZENE', 'PROPENE'): 0.0108658671379089, ('N-PBENZENE', 'STYRENE'): 0.000691056251525879,
    ('N-PBENZENE', 'TOLUENE'): 0.00151985883712769, ('N-PBENZENE', 'TR2-BUTENE'): 0.0052107572555542,
    ('N-PBENZENE', 'TR2-HEXENE'): 0.00071030855178833, ('N-PBENZENE', 'TR2-OCTENE'): 4.04119491577148e-05,
    ('N-PBENZENE', 'TR2-PENTENE'): 0.00196164846420288, ('N-PBENZENE', 'TR3-HEXENE'): 0.00071030855178833,
    ('N-PBENZENE', 'TR3-OCTENE'): 4.04119491577148e-05, ('N-PBENZENE', 'TR4-OCTENE'): 4.04119491577148e-05,
    ('N-PBENZENE', 'VNYLACETLEN'): 0.00681412220001221,
    ('N2O', 'NITROGEN'): 0.0120000001043081, ('N2O', 'O2'): 0.0439990013837814,
    ('NC4', 'NC5'): 0.000547230243682861, ('NC4', 'NC6'): 0.00186634063720703,
    ('NC4', 'NC7'): 0.0036463737487793, ('NC4', 'NC8'): 0.00575017929077148,
    ('NC4', 'NITROGEN'): 0.0900000035762787, ('NC4', 'O-XYLENE'): 0.00189733505249023,
    ('NC4', 'P-XYLENE'): 0.00217920541763306, ('NC4', 'PROPADIENE'): 0.00285130739212036,
    ('NC4', 'PROPANE'): 0.000818967819213867, ('NC4', 'PROPENE'): 0.00163173675537109,
    ('NC4', 'STYRENE'): 0.00144195556640625, ('NC4', 'TOLUENE'): 0.000638782978057861,
    ('NC4', 'TR2-BUTENE'): 6.4849853515625e-05, ('NC4', 'TR2-HEXENE'): 0.00141441822052002,
    ('NC4', 'TR2-OCTENE'): 0.00497031211853027, ('NC4', 'TR2-PENTENE'): 0.000398457050323486,
    ('NC4', 'TR3-HEXENE'): 0.00141441822052002, ('NC4', 'TR3-OCTENE'): 0.00497031211853027,
    ('NC4', 'TR4-OCTENE'): 0.00497031211853027, ('NC4', 'VNYLACETLEN'): 0.000341951847076416,
    ('NC5', 'NC6'): 0.000393450260162354, ('NC5', 'NC7'): 0.001373291015625,
    ('NC5', 'NC8'): 0.0027618408203125, ('NC5', 'NITROGEN'): 0.100000001490116,
    ('NC5', 'O-XYLENE'): 0.000407755374908447, ('NC5', 'P-XYLENE'): 0.000543892383575439,
    ('NC5', 'PROPADIENE'): 0.0058782696723938, ('NC5', 'PROPANE'): 0.00270050764083862,
    ('NC5', 'PROPENE'): 0.0040593147277832, ('NC5', 'STYRENE'): 0.000213086605072021,
    ('NC5', 'TOLUENE'): 3.57627868652344e-06, ('NC5', 'TR2-BUTENE'): 0.000988304615020752,
    ('NC5', 'TR2-HEXENE'): 0.000202596187591553, ('NC5', 'TR2-OCTENE'): 0.00222837924957275,
    ('NC5', 'TR2-PENTENE'): 1.18613243103027e-05, ('NC5', 'TR3-HEXENE'): 0.000202596187591553,
    ('NC5', 'TR3-OCTENE'): 0.00222837924957275, ('NC5', 'TR4-OCTENE'): 0.00222837924957275,
    ('NC5', 'VNYLACETLEN'): 0.00175225734710693,
    ('NC6', 'NC7'): 0.000297248363494873, ('NC6', 'NC8'): 0.00107330083847046,
    ('NC6', 'NITROGEN'): 0.149000003933907, ('NC6', 'O-XYLENE'): 1.78813934326172e-07,
    ('NC6', 'P-XYLENE'): 1.21593475341797e-05, ('NC6', 'PROPADIENE'): 0.00927770137786865,
    ('NC6', 'PROPANE'): 0.0051419734954834, ('NC6', 'PROPENE'): 0.0069580078125,
    ('NC6', 'STYRENE'): 2.74777412414551e-05, ('NC6', 'TOLUENE'): 0.000322282314300537,
    ('NC6', 'TR2-BUTENE'): 0.00262457132339478, ('NC6', 'TR2-HEXENE'): 3.14116477966309e-05,
    ('NC6', 'TR2-OCTENE'): 0.000750899314880371, ('NC6', 'TR2-PENTENE'): 0.000541329383850098,
    ('NC6', 'TR3-HEXENE'): 3.14116477966309e-05, ('NC6', 'TR3-OCTENE'): 0.000750899314880371,
    ('NC6', 'TR4-OCTENE'): 0.000750899314880371, ('NC6', 'VNYLACETLEN'): 0.00379818677902222,
    ('NC7', 'NC8'): 0.000241279602050781, ('NC7', 'NITROGEN'): 0.143900007009506,
    ('NC7', 'O-XYLENE'): 0.000285029411315918, ('NC7', 'P-XYLENE'): 0.000189125537872314,
    ('NC7', 'PROPADIENE'): 0.0128429532051086, ('NC7', 'PROPANE'): 0.00788742303848267,
    ('NC7', 'PROPENE'): 0.0100952982902527, ('NC7', 'STYRENE'): 0.000505208969116211,
    ('NC7', 'TOLUENE'): 0.00600000005215406, ('NC7', 'TR2-BUTENE'): 0.00467795133590698,
    ('NC7', 'TR2-HEXENE'): 0.000521719455718994, ('NC7', 'TR2-OCTENE'): 0.000103354454040527,
    ('NC7', 'TR2-PENTENE'): 0.00163906812667847, ('NC7', 'TR3-HEXENE'): 0.000521719455718994,
    ('NC7', 'TR3-OCTENE'): 0.000103354454040527, ('NC7', 'TR4-OCTENE'): 0.000103354454040527,
    ('NC7', 'VNYLACETLEN'): 0.00620394945144653,
    ('NC8', 'NITROGEN'): 0.100000001490116, ('NC8', 'O-XYLENE'): 0.00105011463165283,
    ('NC8', 'P-XYLENE'): 0.00669990014284849, ('NC8', 'PROPADIENE'): 0.0165321230888367,
    ('NC8', 'PROPANE'): 0.0108502507209778, ('NC8', 'PROPENE'): 0.0134056806564331,
    ('NC8', 'STYRENE'): 0.00144350528717041, ('NC8', 'TOLUENE'): 0.0099999001249671,
    ('NC8', 'TR2-BUTENE'): 0.00702548027038574, ('NC8', 'TR2-HEXENE'): 0.00147128105163574,
    ('NC8', 'TR2-OCTENE'): 2.87890434265137e-05, ('NC8', 'TR2-PENTENE'): 0.00313323736190796,
    ('NC8', 'TR3-HEXENE'): 0.00147128105163574, ('NC8', 'TR3-OCTENE'): 2.87890434265137e-05,
    ('NC8', 'TR4-OCTENE'): 2.87890434265137e-05, ('NC8', 'VNYLACETLEN'): 0.00886505842208862,
    ('NH3', 'NITROGEN'): 0.222000002861023,
    ('NITROGEN', 'O-XYLENE'): 0.214000001549721, ('NITROGEN', 'O2'): -0.0120000001043081,
    ('NITROGEN', 'P-XYLENE'): 0.214000001549721, ('NITROGEN', 'PROPADIENE'): 0.100000001490116,
    ('NITROGEN', 'PROPANE'): 0.0799980014562607, ('NITROGEN', 'PROPENE'): 0.0672990009188652,
    ('NITROGEN', 'SO2'): 0.0799980014562607, ('NITROGEN', 'STYRENE'): 0.100000001490116,
    ('NITROGEN', 'TOLUENE'): 0.193200007081032, ('NITROGEN', 'TR2-BUTENE'): 0.100000001490116,
    ('NITROGEN', 'TR2-HEXENE'): 0.100000001490116, ('NITROGEN', 'TR2-OCTENE'): 0.100000001490116,
    ('NITROGEN', 'TR2-PENTENE'): 0.100000001490116, ('NITROGEN', 'TR3-HEXENE'): 0.100000001490116,
    ('NITROGEN', 'TR3-OCTENE'): 0.100000001490116, ('NITROGEN', 'TR4-OCTENE'): 0.100000001490116,
    ('NITROGEN', 'VNYLACETLEN'): 0.100000001490116,
    ('O-XYLENE', 'P-XYLENE'): 9.83476638793945e-06, ('O-XYLENE', 'PROPADIENE'): 0.00934600830078125,
    ('O-XYLENE', 'PROPANE'): 0.0299980007112026, ('O-XYLENE', 'PROPENE'): 0.00701737403869629,
    ('O-XYLENE', 'STYRENE'): 3.13520431518555e-05, ('O-XYLENE', 'TOLUENE'): 0.000335276126861572,
    ('O-XYLENE', 'TR2-BUTENE'): 0.00266128778457642, ('O-XYLENE', 'TR2-HEXENE'): 3.55243682861328e-05,
    ('O-XYLENE', 'TR2-OCTENE'): 0.000731468200683594, ('O-XYLENE', 'TR2-PENTENE'): 0.000558078289031982,
    ('O-XYLENE', 'TR3-HEXENE'): 3.55243682861328e-05, ('O-XYLENE', 'TR3-OCTENE'): 0.000731468200683594,
    ('O-XYLENE', 'TR4-OCTENE'): 0.000731468200683594, ('O-XYLENE', 'VNYLACETLEN'): 0.00384217500686646,
    ('P-XYLENE', 'PROPADIENE'): 0.00995385646820068, ('P-XYLENE', 'PROPANE'): 0.0299980007112026,
    ('P-XYLENE', 'PROPENE'): 0.00754702091217041, ('P-XYLENE', 'STYRENE'): 7.62343406677246e-05,
    ('P-XYLENE', 'TOLUENE'): 0.000459671020507813, ('P-XYLENE', 'TR2-BUTENE'): 0.00299298763275146,
    ('P-XYLENE', 'TR2-HEXENE'): 8.27312469482422e-05, ('P-XYLENE', 'TR2-OCTENE'): 0.000571966171264648,
    ('P-XYLENE', 'TR2-PENTENE'): 0.000715672969818115, ('P-XYLENE', 'TR3-HEXENE'): 8.27312469482422e-05,
    ('P-XYLENE', 'TR3-OCTENE'): 0.000571966171264648, ('P-XYLENE', 'TR4-OCTENE'): 0.000571966171264648,
    ('P-XYLENE', 'VNYLACETLEN'): 0.00423818826675415,
    ('PROPADIENE', 'PROPANE'): 0.0359990000724792, ('PROPADIENE', 'PROPENE'): 0.0149990003556013,
    ('PROPADIENE', 'STYRENE'): 0.00830632448196411, ('PROPADIENE', 'TOLUENE'): 0.00616836547851563,
    ('PROPADIENE', 'TR2-BUTENE'): 0.00205838680267334, ('PROPADIENE', 'TR2-HEXENE'): 0.00824052095413208,
    ('PROPADIENE', 'TR2-OCTENE'): 0.0152073502540588, ('PROPADIENE', 'TR2-PENTENE'): 0.00536710023880005,
    ('PROPADIENE', 'TR3-HEXENE'): 0.00824052095413208, ('PROPADIENE', 'TR3-OCTENE'): 0.0152073502540588,
    ('PROPADIENE', 'TR4-OCTENE'): 0.0152073502540588, ('PROPADIENE', 'VNYLACETLEN'): 0.00122147798538208,
    ('PROPANE', 'PROPENE'): 0.00789980031549931, ('PROPANE', 'STYRENE'): 0.00442224740982056,
    ('PROPANE', 'TOLUENE'): 0.0309980008751154, ('PROPANE', 'TR2-BUTENE'): 0.000423252582550049,
    ('PROPANE', 'TR2-HEXENE'): 0.00437402725219727, ('PROPANE', 'TR2-OCTENE'): 0.00977498292922974,
    ('PROPANE', 'TR2-PENTENE'): 0.00235652923583984, ('PROPANE', 'TR3-HEXENE'): 0.00437402725219727,
    ('PROPANE', 'TR3-OCTENE'): 0.00977498292922974, ('PROPANE', 'TR4-OCTENE'): 0.00977498292922974,
    ('PROPANE', 'VNYLACETLEN'): 0.000102698802947998,
    ('PROPENE', 'STYRENE'): 0.00611811876296997, ('PROPENE', 'TOLUENE'): 0.00430160760879517,
    ('PROPENE', 'TR2-BUTENE'): 0.00104695558547974, ('PROPENE', 'TR2-HEXENE'): 0.0060616135597229,
    ('PROPENE', 'TR2-OCTENE'): 0.0122110247612, ('PROPENE', 'TR2-PENTENE'): 0.00363552570343018,
    ('PROPENE', 'TR3-HEXENE'): 0.0060616135597229, ('PROPENE', 'TR3-OCTENE'): 0.0122110247612,
    ('PROPENE', 'TR4-OCTENE'): 0.0122110247612, ('PROPENE', 'VNYLACETLEN'): 0.000480771064758301,
    ('STYRENE', 'TOLUENE'): 0.000161647796630859, ('STYRENE', 'TR2-BUTENE'): 0.00211668014526367,
    ('STYRENE', 'TR2-HEXENE'): 1.78813934326172e-07, ('STYRENE', 'TR2-OCTENE'): 0.00106513500213623,
    ('STYRENE', 'TR2-PENTENE'): 0.00032508373260498, ('STYRENE', 'TR3-HEXENE'): 1.78813934326172e-07,
    ('STYRENE', 'TR3-OCTENE'): 0.00106513500213623, ('STYRENE', 'TR4-OCTENE'): 0.00106513500213623,
    ('STYRENE', 'VNYLACETLEN'): 0.00318241119384766,
    ('TOLUENE', 'TR2-BUTENE'): 0.00111013650894165, ('TOLUENE', 'TR2-HEXENE'): 0.000152468681335449,
    ('TOLUENE', 'TR2-OCTENE'): 0.00205451250076294, ('TOLUENE', 'TR2-PENTENE'): 2.83122062683105e-05,
    ('TOLUENE', 'TR3-HEXENE'): 0.000152468681335449, ('TOLUENE', 'TR3-OCTENE'): 0.00205451250076294,
    ('TOLUENE', 'TR4-OCTENE'): 0.00205451250076294, ('TOLUENE', 'VNYLACETLEN'): 0.00191301107406616,
    ('TR2-BUTENE', 'TR2-HEXENE'): 0.00208330154418945, ('TR2-BUTENE', 'TR2-OCTENE'): 0.00616168975830078,
    ('TR2-BUTENE', 'TR2-PENTENE'): 0.000784456729888916, ('TR2-BUTENE', 'TR3-HEXENE'): 0.00208330154418945,
    ('TR2-BUTENE', 'TR3-OCTENE'): 0.00616168975830078, ('TR2-BUTENE', 'TR4-OCTENE'): 0.00616168975830078,
    ('TR2-BUTENE', 'VNYLACETLEN'): 0.000109016895294189,
    ('TR2-HEXENE', 'TR2-OCTENE'): 0.00108909606933594, ('TR2-HEXENE', 'TR2-PENTENE'): 0.000312089920043945,
    ('TR2-HEXENE', 'TR3-OCTENE'): 0.00108909606933594, ('TR2-HEXENE', 'TR4-OCTENE'): 0.00108909606933594,
    ('TR2-HEXENE', 'VNYLACETLEN'): 0.00314152240753174,
    ('TR2-OCTENE', 'TR2-PENTENE'): 0.0025632381439209, ('TR2-OCTENE', 'TR3-HEXENE'): 0.00108909606933594,
    ('TR2-OCTENE', 'VNYLACETLEN'): 0.00789350271224976,
    ('TR2-PENTENE', 'TR3-HEXENE'): 0.000312089920043945, ('TR2-PENTENE', 'TR3-OCTENE'): 0.0025632381439209,
    ('TR2-PENTENE', 'TR4-OCTENE'): 0.0025632381439209, ('TR2-PENTENE', 'VNYLACETLEN'): 0.00147712230682373,
    ('TR3-HEXENE', 'TR3-OCTENE'): 0.00108909606933594, ('TR3-HEXENE', 'TR4-OCTENE'): 0.00108909606933594,
    ('TR3-HEXENE', 'VNYLACETLEN'): 0.00314152240753174,
    ('TR3-OCTENE', 'VNYLACETLEN'): 0.00789350271224976,
    ('TR4-OCTENE', 'VNYLACETLEN'): 0.00789350271224976,
}


def _pr_key(name) -> str | None:
    """Resolve any component name (HMB, simulator or common) to its _PR_DATA key."""
    u = str(name).strip().upper()
    if u in _PR_DATA:
        return u
    return _PR_ALIAS.get(re.sub(r"[^A-Z0-9=]", "", u))

_PR_RG = 83.14462618          # cm3·bar/mol/K
_PR_H2_CP = (3.004605, 1.6031e-03, -1.3188e-06)       # BTU/lb-F, T in K (HMB fit)
_PR_TUNE = {   # coefficients fitted to HMB H2-rich states (applied when H2 >= 50 mol%)
    "m_H2": 0.258927,
    "c_H2": 0.213118,
    "k_H2": 1,
    "c_HC": 0.592434,
    "h_res": 1,
    "e0": 550.947,
    "e1": -0.463104,
    "g_cv": 1.0011,
    "mu_s": 1.34035,
    "k_s": 3,
    "mu_hc": 0.968744,
    "mu_q": -0.0108797,
    "k_hc": 0.899878,
    "k_q": -0.091523,
}
_PR_TUNE_OFF = {"m_H2": None, "c_H2": 1.0, "k_H2": 1.0, "c_HC": 1.0, "h_res": 1.0, "e0": 0.0,
                "e1": 0.0, "g_cv": 1.0, "mu_s": 1.0, "k_s": 1.0, "mu_hc": 1.0, "mu_q": 0.0,
                "k_hc": 1.0, "k_q": 0.0}
_PR_TUNE_FIT = dict(_PR_TUNE)   # frozen copy of the fitted values   # HMB-fitted H2 coefficients
_PR_KEPS = 1.25554               # Mason-Saxena cross-term factor (calibrated to the HMB)
_PR_RJ = 8.314462618          # J/mol/K


def _pr_ready(names) -> bool:
    return bool(names) and all(_pr_key(n) is not None for n in names)


def _pr_poly(co, T):
    return sum(c * T ** i for i, c in enumerate(co))


def _pr_dpoly(co, T):
    return sum(i * c * T ** (i - 1) for i, c in enumerate(co) if i)


class _PRMix:
    """Peng-Robinson mixture (kij from the PR package) with Peneloux volume shift.

    Vapour/liquid Z, residual enthalpy / Cp (numerical dT), Cp-Cv, and the
    ideal-gas enthalpy basis of the simulator (formation basis)."""

    def __init__(s, names, y):
        s.n = [_pr_key(x) for x in names]
        s.y = list(y)
        s.N = len(s.n)
        s.d = [_PR_DATA[x] for x in s.n]
        # HMB-fitted coefficients apply to H2-rich gas only
        yh2 = sum(s.y[i] for i in range(s.N) if s.n[i] == "H2")
        s.tn = _PR_TUNE if yh2 >= 0.5 else _PR_TUNE_OFF
        s.M = sum(s.y[i] * s.d[i][0] for i in range(s.N))
        s.b = sum(s.y[i] * 0.07780 * _PR_RG * s.d[i][1] / s.d[i][2] for i in range(s.N))
        s.c = sum(s.y[i] * s.d[i][8] * (s.tn["c_H2"] if s.n[i] == "H2"
                                        else s.tn["c_HC"]) for i in range(s.N))
        s.k = [[(_PR_KIJ.get(tuple(sorted((s.n[i], s.n[j]))), 0.0)
                 * (s.tn["k_H2"] if "H2" in (s.n[i], s.n[j]) else 1.0)
                 if i != j else 0.0)
                for j in range(s.N)] for i in range(s.N)]

    def _alpha(s, d, T):
        w = d[3]
        m = 0.37464 + 1.54226 * w - 0.26992 * w * w
        if d is _PR_DATA["H2"] and s.tn["m_H2"] is not None:
            m = s.tn["m_H2"]
        return (1 + m * (1 - math.sqrt(T / d[1]))) ** 2

    def a(s, T):
        ai = [0.45724 * _PR_RG ** 2 * d[1] ** 2 / d[2] * s._alpha(d, T) for d in s.d]
        return sum(s.y[i] * s.y[j] * math.sqrt(ai[i] * ai[j]) * (1 - s.k[i][j])
                   for i in range(s.N) for j in range(s.N))

    def Z(s, T, P, liquid=False):
        a = s.a(T)
        A = a * P / (_PR_RG * T) ** 2
        B = s.b * P / (_PR_RG * T)
        co = (1.0, -(1 - B), A - 3 * B * B - 2 * B, -(A * B - B * B - B ** 3))
        z = (B * 1.0001) if liquid else 1.0
        for _ in range(400):
            f = ((z + co[1]) * z + co[2]) * z + co[3]
            df = (3 * z + 2 * co[1]) * z + co[2]
            if df == 0:
                break
            zn = z - f / df
            if liquid:
                zn = max(zn, B * 1.0001)
            if abs(zn - z) < 1e-13:
                z = zn
                break
            z = zn
        return z

    def hres(s, T, P, liquid=False):       # J/mol
        z = s.Z(T, P, liquid)
        h = 0.01
        ap = (s.a(T + h) - s.a(T - h)) / (2 * h)
        a = s.a(T)
        B = s.b * P / (_PR_RG * T)
        r2 = math.sqrt(2.0)
        return s.tn["h_res"] * 0.1 * (
            _PR_RG * T * (z - 1) + (T * ap - a) / (2 * r2 * s.b)
            * math.log((z + (1 + r2) * B) / (z + (1 - r2) * B)))

    def ideal(s, T):
        """(Cp_ig J/mol/K, H_ig J/mol) of the mixture.  H2 uses ideal-gas Cp/h
        equations fitted to the HMB pure-H2 states (reference: 77 F = 0)."""
        cp = h = 0.0
        for i in range(s.N):
            d = s.d[i]
            if s.n[i] == "H2":
                a0, b0, c0 = _PR_H2_CP
                T0 = (77 + 459.67) / 1.8
                cp_i = (a0 + b0 * T + c0 * T * T) / 0.238846 * d[0]          # J/mol/K
                h_i = 1.8 * (a0 * (T - T0) + b0 * (T * T - T0 * T0) / 2
                             + c0 * (T ** 3 - T0 ** 3) / 3) / 0.429923 * d[0]  # J/mol
            else:
                cp_i = d[0] * _pr_dpoly(d[7], T) + s.tn["e1"]
                h_i = (d[0] * _pr_poly(d[7], T) + d[5]
                       + s.tn["e0"] + s.tn["e1"] * (T - 298.15))
            cp += s.y[i] * cp_i
            h += s.y[i] * h_i
        return cp, h

    def h_only(s, T, P, liquid=False):
        """Specific enthalpy, BTU/lb (formation basis) — lighter than props()."""
        return (s.ideal(T)[1] + s.hres(T, P, liquid)) / s.M * 0.429923

    def props(s, T, P, liquid=False):
        """Single-phase properties at T[K], P[bar]; FPS outputs."""
        z = s.Z(T, P, liquid)
        v = z * _PR_RG * T / P
        a = s.a(T)
        h = 0.01
        ap = (s.a(T + h) - s.a(T - h)) / (2 * h)
        cpi, hi = s.ideal(T)
        dT = 0.05
        cp = cpi + (s.hres(T + dT, P, liquid) - s.hres(T - dT, P, liquid)) / (2 * dT)
        q = v * v + 2 * s.b * v - s.b * s.b
        dPdT = _PR_RG / (v - s.b) - ap / q
        dPdv = -_PR_RG * T / (v - s.b) ** 2 + a * (2 * v + 2 * s.b) / q ** 2
        cpmcv = -T * dPdT ** 2 / dPdv * 0.1
        vt = v - s.c
        return {"Zk": z, "Z": P * vt / (_PR_RG * T),
                "rho": s.M / vt * 62.428,                       # lb/ft3
                "cp": cp / s.M * 0.238846,                      # BTU/lb-F
                "gamma": cp / (cp - s.tn["g_cv"] * cpmcv),
                "h": (hi + s.hres(T, P, liquid)) / s.M * 0.429923}   # BTU/lb


def _pr_lucas_dilute(d, T):
    """Non-polar Lucas dilute-gas viscosity, µP (d = _PR_DATA tuple)."""
    tr = T / d[1]
    xi = 0.176 * (d[1] / (d[0] ** 3 * d[2] ** 4)) ** (1.0 / 6.0)
    return (0.807 * tr ** 0.618 - 0.357 * math.exp(-0.449 * tr)
            + 0.340 * math.exp(-4.058 * tr) + 0.018) / xi


def _pr_vapor_props(names, y, T_F, P_psia) -> dict:
    """Vapour properties of a PR-covered mixture, calibrated to the HMB:
    Z, ρ, Cp, Cp/Cv, h from PR+Peneloux; μ (Lucas + Wilke + Jossi-Stiel-Thodos);
    k (Mason-Saxena, ε = 1.043, + Stiel-Thodos density term).  Pure H2 uses
    equations fitted to the HMB.  Returns FPS: z_vap, rho_vap, cp_vap, cpcv_vap,
    h_vap, visc_vap (cP), k_vap (BTU/hr-ft-F)."""
    sy = sum(y)
    y = [v / sy for v in y]
    m = _PRMix(names, y)
    T = (T_F + 459.67) / 1.8
    P = P_psia * 0.0689476
    p = m.props(T, P)
    n, d = m.N, m.d
    # H2 transport fitted to the HMB (pure-H2 states): Sutherland mu, power-law k
    mu_i, k_i, cpi = [], [], []
    for i in range(n):
        cpi.append(d[i][0] * _pr_dpoly(d[i][7], T))
        if m.n[i] == "H2":
            mu_i.append(0.009346 * (T / 310.928) ** 1.5 * (310.928 + 90.91) / (T + 90.91) * 1e4)
            k_i.append(0.106675 * (T / 310.928) ** 0.705 * 1.7307)
        else:
            e = _pr_lucas_dilute(d[i], T) * m.tn["mu_hc"] * (T / 373.15) ** m.tn["mu_q"]
            mu_i.append(e)
            k_i.append(e * 1e-7 / (d[i][0] * 1e-3) * (1.32 * (cpi[i] - _PR_RJ) + 1.77 * _PR_RJ)
                       * m.tn["k_hc"] * (T / 373.15) ** m.tn["k_q"])

    def phi(i, j):
        return ((1 + math.sqrt(mu_i[i] / mu_i[j]) * (d[j][0] / d[i][0]) ** 0.25) ** 2
                / math.sqrt(8 * (1 + d[i][0] / d[j][0])))

    def aij(i, j):
        return (((_PR_KEPS if m.tn is _PR_TUNE else 1.0) if i != j else 1.0)
                * (1 + math.sqrt(k_i[i] / k_i[j]) * (d[i][0] / d[j][0]) ** 0.25) ** 2
                / math.sqrt(8 * (1 + d[i][0] / d[j][0])))
    mu0 = sum(y[i] * mu_i[i] / sum(y[j] * phi(i, j) for j in range(n)) for i in range(n))
    k0 = sum(y[i] * k_i[i] / sum(y[j] * aij(i, j) for j in range(n)) for i in range(n))
    tcm = sum(y[i] * d[i][1] for i in range(n))
    pcm = sum(y[i] * d[i][2] for i in range(n))
    vcm = sum(y[i] * d[i][4] for i in range(n))
    wm = sum(y[i] * d[i][3] for i in range(n))
    rr = P / (p["Zk"] * _PR_RG * T) * vcm
    mu = mu0
    if rr > 0.1:
        xi = tcm ** (1 / 6) / (m.M ** 0.5 * (pcm / 1.01325) ** (2 / 3))
        rhs = (1.023 + 0.23364 * rr + 0.58533 * rr ** 2 - 0.40758 * rr ** 3
               + 0.093324 * rr ** 4)
        mu = mu0 + m.tn["mu_s"] * (rhs ** 4 - 1) / xi
    k = k0
    if m.n == ["H2"]:        # pure H2: HMB-fitted k(T, P) already has the P effect
        k = 0.106675 * (T / 310.928) ** 0.705 * (1 + 4.99e-5 * P_psia) * 1.7307
    elif rr < 0.5:
        zc = 0.291 - 0.080 * wm
        gam = 210 * (tcm * m.M ** 3 / pcm ** 4) ** (1 / 6)
        k = k0 + m.tn["k_s"] * 1.22e-2 * (math.exp(0.535 * rr) - 1) / (gam * zc ** 5)
    return {"z_vap": p["Z"], "rho_vap": p["rho"], "cp_vap": p["cp"],
            "cpcv_vap": p["gamma"], "h_vap": p["h"],
            "visc_vap": mu * 1e-4, "k_vap": k * 0.577789}


_COMP_ISOTHERMAL = False      # True (CLI --isothermal-comp) → composition streams stay isothermal


def _pr_stream_h(names, fr, T_F, P_psia) -> float | None:
    """Total specific enthalpy (BTU/lb) of a flashed composition stream (PR)."""
    try:
        T, P = (T_F + 459.67) / 1.8, P_psia * 0.0689476
        vm, lm = fr.vap_mass, fr.liq_mass
        tot = vm + lm
        if tot <= 0:
            return None
        h = 0.0
        if vm > 0 and fr.y:
            sy = sum(fr.y) or 1.0
            h += vm * _PRMix(names, [v / sy for v in fr.y]).h_only(T, P)
        if lm > 0 and fr.x:
            sx = sum(fr.x) or 1.0
            h += lm * _PRMix(names, [v / sx for v in fr.x]).h_only(T, P, liquid=True)
        return h / tot
    except Exception:
        return None


def _iso_T_out(feed, T_in_F, P_in, P_out):
    """Outlet temperature (°F) of an adiabatic, no-work (isenthalpic) pressure
    change P_in → P_out of a composition stream, using the PR enthalpy and the
    temperature-corrected flash.  Returns T_in if it cannot be solved."""
    if feed is None or not _pr_ready(feed.names) or abs(P_in - P_out) < 1e-9:
        return T_in_F
    fr_in = FV.flash(feed, P_in, temp_f=T_in_F)
    h_in = _pr_stream_h(feed.names, fr_in, T_in_F, P_in)
    if h_in is None:
        return T_in_F

    def f(T):
        h = _pr_stream_h(feed.names, FV.flash(feed, P_out, temp_f=T), T, P_out)
        return None if h is None else h - h_in
    t0, t1 = T_in_F, T_in_F - 1.0
    f0, f1 = f(t0), f(t1)
    if f0 is None or f1 is None:
        return T_in_F
    for _ in range(40):
        if f1 == f0:
            break
        t2 = t1 - f1 * (t1 - t0) / (f1 - f0)
        t2 = min(max(t2, T_in_F - 250.0), T_in_F + 250.0)
        f2 = f(t2)
        if f2 is None:
            return T_in_F
        t0, f0, t1, f1 = t1, f1, t2, f2
        if abs(t1 - t0) < 1e-5:
            break
    return t1


def _enthalpy_mix_T(feed_a, T_a, mass_a, feed_b, T_b, mass_b, blended, p_psia, T_guess):
    """Mixed-stream temperature (°F) from an ENTHALPY BALANCE:
    mass_a·h_a(T_a) + mass_b·h_b(T_b) = (mass_a+mass_b)·h_mix(T_mix), all at p_psia,
    with the PR enthalpy and the temperature-corrected flash.  Falls back to T_guess."""
    try:
        if not (_pr_ready(feed_a.names) and _pr_ready(feed_b.names)
                and _pr_ready(blended.names)) or (mass_a + mass_b) <= 0:
            return T_guess
        h_a = _pr_stream_h(feed_a.names, FV.flash(feed_a, p_psia, temp_f=T_a), T_a, p_psia)
        h_b = _pr_stream_h(feed_b.names, FV.flash(feed_b, p_psia, temp_f=T_b), T_b, p_psia)
        if h_a is None or h_b is None:
            return T_guess
        h_t = (h_a * mass_a + h_b * mass_b) / (mass_a + mass_b)

        def f(T):
            h = _pr_stream_h(blended.names, FV.flash(blended, p_psia, temp_f=T), T, p_psia)
            return None if h is None else h - h_t
        t0, t1 = T_guess, T_guess - 1.0
        f0, f1 = f(t0), f(t1)
        if f0 is None or f1 is None:
            return T_guess
        for _ in range(40):
            if f1 == f0:
                break
            t2 = t1 - f1 * (t1 - t0) / (f1 - f0)
            t2 = min(max(t2, min(T_a, T_b) - 250.0), max(T_a, T_b) + 250.0)
            f2 = f(t2)
            if f2 is None:
                return T_guess
            t0, f0, t1, f1 = t1, f1, t2, f2
            if abs(t1 - t0) < 1e-5:
                break
        return t1
    except Exception:
        return T_guess


# ════════════════════════════════════════════════════════════════════════
#  Per-station stream properties (Component_Detail)
# ════════════════════════════════════════════════════════════════════════
# Standard-liquid specific gravity (60°F, volume-additive basis, GPSA-style) for
# light pure components whose HMB 'SLD' entries are not usable.  Priority:
# this table → library/HMB SLD (20–125 lb/ft3) → SG 0.70 assumed.  VERIFY.
_STD_LIQ_SG = {
    "H2O": 0.999, "H2S": 0.8014, "NH3": 0.6173, "H2": 0.0709, "NITROGEN": 0.8094,
    "METHANE": 0.3000, "ETHANE": 0.3564, "PROPANE": 0.5077, "IC4": 0.5631,
    "NC4": 0.5844, "IC5": 0.6247, "NC5": 0.6311, "CO2": 0.8172, "O2": 1.1410,
    "CO": 0.7931, "CL2": 1.4247, "SO2": 1.3970, "PROPYLENE": 0.5220,
    "1-BUTENE": 0.6004, "BUTADIENE": 0.6272,
}
_V_STD_FT3_LBMOL = 379.48          # ideal gas at 14.696 psia, 60°F


def _std_liq_density(name: str, c: dict | None, fallback: float | None = None) -> float:
    sg = _STD_LIQ_SG.get(str(name).upper())
    if sg:
        return sg * 62.428
    sld = (c or {}).get("sld")
    if sld and 20.0 <= sld <= 125.0:
        return sld
    return fallback or 0.70 * 62.428


def _station_props(fr, sp, feed, cmap, p_psia, est_inputs=None) -> dict | None:
    """Everything Component_Detail reports for one station (internal FPS).

    Computed from the station's flash (phase split, component flows) and the
    component constants (HMB COMP_CONSTANTS / built-in library).  Properties
    that need a property package (Cp, k, σ, enthalpy) are CARRIED from the HMB
    stream (``sp``) and are blank for manual-composition streams.
    """
    if fr is None:
        return None
    cmap = cmap or {}
    T = fr.temp_f if fr.temp_f is not None else (sp.temp_f if sp else None)
    vm, lm, nv, nl = fr.vap_mass, fr.liq_mass, fr.vap_moles, fr.liq_moles
    N, M = nv + nl, vm + lm
    d: dict = {"phase": _phase_of(fr.quality), "T": T, "P": p_psia,
               "n_tot": N, "m_tot": M, "n_vap": nv, "n_liq": nl,
               "m_vap": vm, "m_liq": lm,
               "mw": (M / N) if N > 0 else None,
               "beta": (nv / N) if N > 0 else None,
               "x_liq": (nl / N) if N > 0 else None,
               "w_liq": (lm / M) if M > 0 else None}
    names = list(getattr(feed, "names", None) or [])
    y, x = (fr.y or []), (fr.x or [])
    est: dict = {}
    if names and len(y) == len(names) and len(x) == len(names) and N > 0:
        try:
            est = _est_props(cmap, names, y, x, nv, nl, T, p_psia)
        except Exception:
            est = {}
    est_keys: list[str] = []

    def pick(sp_attr, est_key, label, active=True):
        """Input/HMB value if present, else the correlation estimate."""
        v = getattr(sp, sp_attr, None) if sp else None
        if v is None and active and est.get(est_key) is not None:
            est_keys.append(label)
            return est[est_key]
        return v
    z_vap = pick("vap_z", "z_vap", "Z", vm > 0)
    rho_g = (_gas_rho_sp(sp, fr.vap_mw, T, p_psia, z_vap) / LBFT3_TO_KGM3
             if (vm > 0 and T is not None) else None)
    rho_l = pick("liq_density", "rho_liq", "liq ρ", lm > 0)
    vg = (vm / rho_g) if rho_g else 0.0
    vl = (lm / rho_l) if rho_l else 0.0
    d["q_std_vap_ph"] = nv * _V_STD_FT3_LBMOL
    d["z_k"] = None
    if vm > 0 and T is not None and _pr_ready(names) and len(y) == len(names):
        try:
            sy_ = sum(y) or 1.0
            d["z_k"] = _PRMix(names, [v / sy_ for v in y]).Z((T + 459.67) / 1.8,
                                                           p_psia * 0.0689476)
        except Exception:
            d["z_k"] = None
    d.update(rho_vap=rho_g, rho_liq=rho_l, q_vap=vg or None, q_liq=vl or None,
             q_tot=(vg + vl) or None,
             rho_tot=(M / (vg + vl)) if (vg + vl) > 0 else None,
             z_tot=((p_psia * (vg + vl)) / (N * 10.7316 * (T + 459.67))
                    if ((vg + vl) > 0 and N > 0 and T is not None) else None),
             q_std_vap=N * _V_STD_FT3_LBMOL,
             z_vap=(z_vap if vm > 0 else None),
             visc_vap=pick("vap_visc", "visc_vap", "vap μ", vm > 0),
             visc_liq=pick("liq_visc", "visc_liq", "liq μ", lm > 0),
             cpcv_vap=pick("vap_cp_cv", "cpcv_vap", "Cp/Cv", vm > 0))
    # HMB-carried properties; correlation estimate where the HMB has none
    for k, a, lab, act in (
            ("cp_vap", "vap_cp", "vap Cp", vm > 0), ("cp_liq", "liq_cp", "liq Cp", lm > 0),
            ("k_vap", "vap_therm_cond", "vap k", vm > 0),
            ("k_liq", "liq_therm_cond", "liq k", lm > 0),
            ("sigma", "liq_surf_tens", "σ", lm > 0),
            ("h_vap", "vap_sp_enthalpy", "h vap", vm > 0),
            ("h_liq", "liq_sp_enthalpy", "h liq", lm > 0)):
        d[k] = pick(a, k, lab, act)
    _lab = {"vap_z": "Z", "vap_visc": "vap μ", "vap_cp_cv": "Cp/Cv",
            "liq_density": "liq ρ", "liq_visc": "liq μ"}
    est_keys += [_lab[k] for k in (est_inputs or []) if k in _lab]
    d["est_keys"] = est_keys
    # Manual stream with a liquid but no liquid MW: molar quantities would rest on a
    # fictitious MW, so leave them blank instead of showing misleading numbers.
    if feed is None and lm > 0 and not (getattr(sp, "liq_mw", None) if sp else None):
        for k_ in ("n_tot", "n_liq", "mw", "beta", "x_liq", "z_tot", "q_std_vap"):
            d[k_] = None
    if d["h_vap"] is not None or d["h_liq"] is not None:
        hv, hl = (d["h_vap"] or 0.0), (d["h_liq"] or 0.0)
        wv = (vm / M) if M > 0 else 0.0
        d["h_spec"] = wv * hv + (1 - wv) * hl
        d["h_tot"] = d["h_spec"] * M / 1e6                 # MMBTU/hr

    # ── composition-based properties ─────────────────────────────────────
    if names and len(y) == len(names) and len(x) == len(names) and N > 0:
        mwl = feed.mw
        vmol = [nv * yi for yi in y]
        lmol = [nl * xi for xi in x]
        tot = [a + b for a, b in zip(vmol, lmol)]
        zi = [t / N for t in tot]
        fb = None                    # HMB stream's own average std density
        if sp and sp.total_std_liq and sp.total_mass:
            fb = sp.total_mass / sp.total_std_liq
        rho_std = [_std_liq_density(n, cmap.get(str(n).upper()), fb)
                   for n in names]
        svol_v = sum(vmol[i] * mwl[i] / rho_std[i] for i in range(len(names)))
        svol_l = sum(lmol[i] * mwl[i] / rho_std[i] for i in range(len(names)))
        svol = svol_v + svol_l
        d.update(q_std_liq=svol, q_std_liq_liq=svol_l,
                 rho_std_liq=(lm / svol_l) if svol_l > 0 else None)
        if svol > 0:
            sg = (M / svol) / 62.428
            d["api"] = 141.5 / sg - 131.5
        d["x_water"] = sum(zi[i] for i, n in enumerate(names)
                           if str(n).upper() in ("H2O", "WATER"))

        def mavg(key, scale=None):
            # Kay mole-average; only trusted when the constants cover ≥98 % of
            # the moles (HMB pseudo-components may be missing from the sheet)
            num_, den_ = 0.0, 0.0
            for i, n in enumerate(names):
                v = (cmap.get(str(n).upper()) or {}).get(key)
                if v is None:
                    continue
                num_ += zi[i] * (scale(v) if scale else v)
                den_ += zi[i]
            return (num_ / den_) if den_ >= 0.98 else None
        d["tc"] = mavg("tc_f")
        d["pc"] = mavg("pc_psia")
        d["tc_kay"], d["pc_kay"] = d["tc"], d["pc"]       # pure Kay's rule (no HMB fallback)
        d["comp_y"] = {n_: y[i] for i, n_ in enumerate(names)} if nv > 0 else {}
        d["comp_x"] = {n_: x[i] for i, n_ in enumerate(names)} if nl > 0 else {}
        d["vc"] = mavg("vc")
        d["omega"] = mavg("omega")
        # fall back to the HMB stream's own (rigorous) values where incomplete
        if sp is not None:
            for k, a in (("tc", "tc_f"), ("pc", "pc_psia"), ("omega", "acentric")):
                if d.get(k) is None:
                    d[k] = getattr(sp, a, None)
        tb = mavg("nbp_f", lambda v: v + 459.67)
        if tb and svol > 0:
            d["watson"] = (tb ** (1.0 / 3.0)) / ((M / svol) / 62.428)
    return d


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
    "Liq MW":               "liq_mw",
    "Liq Surf Tension (N/m)": "liq_surf_tens",
    "Vap Cp (BTU/lb-F)":    "vap_cp",
    "Liq Cp (BTU/lb-F)":    "liq_cp",
}


def _row_yellow_overrides(row: dict) -> dict:
    """Non-blank yellow manual-input cells on this row, keyed by field name."""
    out = {}
    for col, field_name in _YELLOW_OVERRIDE_MAP.items():
        v = num(row.get(col))
        if v is not None:
            out[field_name] = v * 1000.0 if field_name == "liq_surf_tens" else v   # N/m → dyn/cm
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
        const_map: dict | None = None,
) -> tuple[list[HE.Station], list, list, dict]:
    """
    March pressure along ``block_rows`` with flash-coupled VLE.

    Manual composition (no HMB): a row carrying ``_comp`` (from the
    Stream_Composition sheet) STARTS the stream if none is active, otherwise
    its component flows are MIXED into the running stream and re-flashed.
    ``const_map`` (HMB COMP_CONSTANTS) supplies MW / Tc / Pc / ω for it.

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
    est_inputs: list[str] = []      # sp fields filled by correlation (composition streams)
    iso_T = False                   # composition stream marched adiabatically (isenthalpic)
    T_run = None                    # running stream temperature, °F (when iso_T)

    def _flash(fd, pp):
        if iso_T and T_run is not None:
            return FV.flash(fd, pp, temp_f=T_run)
        t_ov = override_state.get("temp_f")      # typed Temp: flash at that temperature
        if t_ov is not None and not use_isoH:
            return FV.flash(fd, pp, temp_f=t_ov)
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
            iso_T = False
        elif cur_feed is None and row_sp is not None:
            cur_sp = row_sp                       # manual rows refresh sp

        # ── Manual composition stream (Stream_Composition sheet) ─────────
        # No active stream → this composition STARTS the flow.  Otherwise it
        # is MIXED into the running stream below.  Its yellow cells describe
        # THIS inlet only, so they must not become persistent overrides.
        comp_pl, mix_pl = None, None
        scale_tag_mix = ""
        if row.get("_comp"):
            try:
                comp_pl = _comp_payload(row, max(p, p_floor), const_map)
            except ValueError as exc:
                raise SystemExit(f"  ERROR (Line {row.get('Line No')}, Seq "
                                 f"{row.get('Seq')}): {exc}")
            for _k in comp_pl.get("estimated") or []:
                if _k not in est_inputs:
                    est_inputs.append(_k)
            if comp_pl.get("estimated"):
                scale_tag_mix += (" | estimated from composition: "
                                  + ", ".join(comp_pl["estimated"]))
            if cur_feed is None:
                cur_feed, cur_sp = comp_pl["feed"], comp_pl["sp"]
                iso_T = (not _COMP_ISOTHERMAL and _pr_ready(cur_feed.names)
                         and cur_sp.temp_f is not None)
                T_run = cur_sp.temp_f if iso_T else None
                flow_scale = vap_cf = liq_cf = 1.0
                mass_scale, mass_scale_set = 1.0, False
                override_state = {}
            else:
                mix_pl = comp_pl
        else:
            override_state.update(_row_yellow_overrides(row))
        a_feed, a_sp = cur_feed, (cur_sp or default_sp)

        if mix_pl is not None:
            p_mix = max(p, p_floor)
            fr0 = _flash(a_feed, p_mix)
            main_moles = _feed_comp_flows(
                a_feed, fr0, flow_scale * vap_cf * mass_scale,
                flow_scale * liq_cf * mass_scale)
            blended = _blend_feeds(a_feed, main_moles,
                                   mix_pl["feed"], mix_pl["moles"])
            if blended is not None:
                mwm = {n: a_feed.mw[i] for i, n in enumerate(a_feed.names)}
                mwm.update({n: mix_pl["feed"].mw[i]
                            for i, n in enumerate(mix_pl["feed"].names)})
                main_mass = sum(m * mwm.get(n, 50.0) for n, m in main_moles.items())
                add_mass = sum(m * mwm.get(n, 50.0)
                               for n, m in mix_pl["moles"].items())
                new_sp = _blend_sp(a_sp, mix_pl["sp"], main_mass, add_mass)
                if new_sp is not None and (main_mass + add_mass) > 0:
                    wa = main_mass / (main_mass + add_mass)
                    ssp = mix_pl["sp"]

                    def _mx(av, bv):
                        return bv if av is None else av if bv is None else wa * av + (1 - wa) * bv
                    new_sp = replace(
                        new_sp,
                        temp_f=_mx(a_sp.temp_f if a_sp else None, ssp.temp_f),
                        vap_cp_cv=_mx(a_sp.vap_cp_cv if a_sp else None, ssp.vap_cp_cv))
                if (iso_T and T_run is not None and new_sp is not None
                        and mix_pl["feed"] is not None):
                    # ENTHALPY BALANCE for the mixed temperature
                    T_b = mix_pl["sp"].temp_f if mix_pl["sp"].temp_f is not None else T_run
                    T_mix = _enthalpy_mix_T(a_feed, T_run, main_mass, mix_pl["feed"],
                                            T_b, add_mass, blended, p_mix, new_sp.temp_f)
                    scale_tag_mix += (f" | T mix {T_run:.2f} + {T_b:.2f} → {T_mix:.2f} °F "
                                      f"(enthalpy balance)")
                    new_sp = replace(new_sp, temp_f=T_mix)
                cur_feed, cur_sp = blended, new_sp
                a_feed, a_sp = cur_feed, (cur_sp or default_sp)
                iso_T = (not _COMP_ISOTHERMAL and _pr_ready(blended.names)
                         and new_sp is not None and new_sp.temp_f is not None)
                T_run = new_sp.temp_f if iso_T else None
                flow_scale = vap_cf = liq_cf = 1.0
                mass_scale = 1.0
                override_state.pop("vap_mass", None)
                override_state.pop("liq_mass", None)
                scale_tag_mix = (f" | composition mixed in (+{add_mass:.1f} lb/hr)"
                                 f" → re-flash blended feed")

        if iso_T and T_run is not None and a_sp is not None:
            a_sp = replace(a_sp, temp_f=T_run)         # carry the running temperature
            # refresh ESTIMATED vapour Z / µ / Cp-Cv at this station's local T, P
            # (values you typed on the input sheet are never touched)
            if est_inputs and a_feed is not None and (a_sp.vap_mass or 0) > 0:
                try:
                    _fl = FV.flash(a_feed, max(p, p_floor), temp_f=T_run)
                    if _fl.vap_mass > 0 and _fl.y:
                        _vp = _pr_vapor_props(a_feed.names, _fl.y, T_run, max(p, p_floor))
                        _upd = {}
                        for _a, _k in (("vap_z", "z_vap"), ("vap_visc", "visc_vap"),
                                       ("vap_cp_cv", "cpcv_vap")):
                            if _a in est_inputs:
                                _upd[_a] = _vp[_k]
                        if _upd:
                            a_sp = replace(a_sp, **_upd)
                except Exception:
                    pass
            cur_sp = a_sp
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
        mom = dict(mach_vap=None, rhov2_vap=None, rhov2_liq=None)

        # ── Tee split / merge: update running flow scale ─────────────────
        # negative = SPLIT: a delta fraction of the CURRENT flow leaves the line.
        # positive = MERGE: the running flow becomes that fraction of the
        # original Main flow (an ABSOLUTE target), e.g. 0.125 → 0.25 → 0.5 → 1.0
        # builds a header up in stages as branches join it.
        frac = num(row.get("Flow Fraction from Main"))
        split_n = _is_tee_split_branch(fitting)
        pending_split_scale = None    # filled below if this row splits off a branch
        scale_tag += scale_tag_mix
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
                                              flow_scale * vap_cf * mass_scale,
                                              flow_scale * liq_cf * mass_scale)
                blended = _blend_feeds(a_feed, main_moles, b_feed, b_comp)
                if blended is not None:
                    mwmap = {n: a_feed.mw[i] for i, n in enumerate(a_feed.names)}
                    main_mass = sum(m * mwmap.get(n, 50.0)
                                    for n, m in main_moles.items())
                    t_note = ""
                    new_sp_j = _blend_sp(a_sp, inj.get("sp"), main_mass, b_mass)
                    if iso_T and T_run is not None and new_sp_j is not None:
                        T_b = inj.get("T_out")
                        if T_b is None:
                            T_b = (inj["sp"].temp_f if inj.get("sp") is not None
                                   and inj["sp"].temp_f is not None else T_run)
                        T_mix = _enthalpy_mix_T(a_feed, T_run, main_mass, b_feed, T_b,
                                                b_mass, blended, p_eval, T_run)
                        t_note = (f" | T mix {T_run:.2f} + {T_b:.2f} → {T_mix:.2f} °F "
                                  f"(enthalpy balance)")
                        new_sp_j = replace(new_sp_j, temp_f=T_mix)
                        iso_T = _pr_ready(blended.names)
                        T_run = T_mix if iso_T else None
                    cur_feed = blended
                    cur_sp   = new_sp_j
                    a_feed, a_sp = cur_feed, cur_sp
                    flow_scale = vap_cf = liq_cf = 1.0   # blend = new full flow
                    mass_scale = 1.0                     # main's flow override is now in the blend
                    override_state.pop("vap_mass", None)
                    override_state.pop("liq_mass", None)
                    scale_tag += (f" | Tee join {inj.get('label','Branch')} "
                                  f"→ re-flash blended feed (+{b_mass:.1f} lb/hr)" + t_note)
            elif (a_feed is None and b_feed is None and a_sp is not None
                  and inj.get("sp") is not None and b_mass > 0):
                # ── Manual (frozen-split) streams: add the flows, merge the phase
                # properties (vapour props from whichever run carries vapour, etc.)
                bsp = inj["sp"]
                mv = ((override_state.get("vap_mass", a_sp.vap_mass) or 0.0)
                      * flow_scale * vap_cf)
                ml = ((override_state.get("liq_mass", a_sp.liq_mass) or 0.0)
                      * flow_scale * liq_cf)
                bv, bl = inj.get("vap_mass", 0.0), inj.get("liq_mass", 0.0)
                mwa = a_sp.vap_mw or a_sp.mol_weight
                mwb = bsp.vap_mw or bsp.mol_weight
                mol_v = ((mv / mwa) if (mv > 0 and mwa) else 0.0) + \
                        ((bv / mwb) if (bv > 0 and mwb) else 0.0)
                vmw = ((mv + bv) / mol_v) if mol_v > 0 else (mwa or mwb)
                nsp = _blend_sp(a_sp, bsp, mv + ml, bv + bl)
                if nsp is not None:
                    def _wavg(pairs):
                        """weighted mean over (weight, value) pairs with a value."""
                        pr_ = [(w, v) for w, v in pairs if v is not None and w > 0]
                        tw_ = sum(w for w, _ in pr_)
                        return (sum(w * v for w, v in pr_) / tw_) if tw_ > 0 else None
                    cpv = _wavg([(mv, a_sp.vap_cp), (bv, bsp.vap_cp)])
                    cpl = _wavg([(ml, a_sp.liq_cp), (bl, bsp.liq_cp)])
                    sig = _wavg([(ml, a_sp.liq_surf_tens), (bl, bsp.liq_surf_tens)])
                    # liquid MW: mass-weighted harmonic (mole-weighted)
                    lmw_a = a_sp.liq_mw or (a_sp.mol_weight if ml > 0 and mv <= 0 else None)
                    mol_l = ((ml / lmw_a) if (ml > 0 and lmw_a) else 0.0) + \
                            ((bl / bsp.liq_mw) if (bl > 0 and bsp.liq_mw) else 0.0)
                    lmw = ((ml + bl) / mol_l) if mol_l > 0 else None
                    # ENTHALPY (sensible-heat) balance for the mixed temperature:
                    #   Σ m_i·Cp_i·(T_mix − T_i) = 0   (no phase change / heat of mixing)
                    Ta_ = override_state.get("temp_f", a_sp.temp_f)
                    Tb_ = bsp.temp_f
                    ca_ = (mv * (a_sp.vap_cp or 0.0) + ml * (a_sp.liq_cp or 0.0))
                    cb_ = (bv * (bsp.vap_cp or 0.0) + bl * (bsp.liq_cp or 0.0))
                    ok_cp = (Ta_ is not None and Tb_ is not None
                             and (mv <= 0 or a_sp.vap_cp) and (ml <= 0 or a_sp.liq_cp)
                             and (bv <= 0 or bsp.vap_cp) and (bl <= 0 or bsp.liq_cp)
                             and (ca_ + cb_) > 0)
                    if ok_cp:
                        T_mix_m = (ca_ * Ta_ + cb_ * Tb_) / (ca_ + cb_)
                        t_note_m = (f" | T mix {Ta_:.2f} + {Tb_:.2f} → {T_mix_m:.2f} °F "
                                    f"(enthalpy balance, Cp)")
                    else:
                        T_mix_m = (((mv + ml) * Ta_ + (bv + bl) * Tb_) / (mv + ml + bv + bl)
                                   if (Ta_ is not None and Tb_ is not None) else nsp.temp_f)
                        t_note_m = (f" | T mix {T_mix_m:.2f} °F (mass-weighted — enter Vap Cp / "
                                    f"Liq Cp for an enthalpy balance)"
                                    if (Ta_ is not None and Tb_ is not None
                                        and abs(Ta_ - Tb_) > 0.05) else "")
                    nsp = replace(
                        nsp, vap_mass=mv + bv, liq_mass=ml + bl,
                        total_mass=mv + bv + ml + bl, vap_mw=vmw,
                        mol_weight=vmw if vmw else lmw, liq_mw=lmw,
                        vap_cp=cpv, liq_cp=cpl, liq_surf_tens=sig, temp_f=T_mix_m,
                        vap_cp_cv=(a_sp.vap_cp_cv if (mv > 0 and a_sp.vap_cp_cv)
                                   else bsp.vap_cp_cv),
                        phase=("TWO-PHASE" if (mv + bv) > 0 and (ml + bl) > 0
                               else "VAPOR" if (mv + bv) > 0 else "LIQUID"))
                    cur_sp, a_sp = nsp, nsp
                    flow_scale = vap_cf = liq_cf = 1.0
                    for _k in set(_YELLOW_OVERRIDE_MAP.values()):   # merged values now in sp
                        override_state.pop(_k, None)
                    scale_tag += (f" | Tee join {inj.get('label','Branch')} (manual streams): "
                                  f"+{bv:.1f} vapour, +{bl:.1f} liquid lb/hr" + t_note_m)
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
            # A flow typed on THIS row (not just carried from an earlier row)
            # re-targets the running flow: mass_scale is re-solved so the flow
            # leaving this row equals the typed value (e.g. 100% → 50% → 25%
            # down successive tee rows).
            row_flow = {k: v for k, v in _row_yellow_overrides(row).items()
                        if k in ("vap_mass", "liq_mass")}
            if row_flow and mass_scale_set:
                eff_v = fr_full.vap_mass * flow_scale * vap_cf * mass_scale
                eff_l = fr_full.liq_mass * flow_scale * liq_cf * mass_scale
                if eff_v + eff_l > 0:
                    tgt = (row_flow.get("vap_mass", eff_v)
                           + row_flow.get("liq_mass", eff_l))
                    mass_scale *= tgt / (eff_v + eff_l)
            elif not mass_scale_set and ("vap_mass" in override_state
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

        elif _instr_kind(fitting, instr_type) == "FT" and instr_dp > 0:
            # inline flow transmitter (mag / Coriolis / vortex …) with a rated ΔP
            dp_t  = instr_dp
            dp_f  = dp_k = dp_z = 0.0
            res   = dict(rho=0.0, v=0.0, Re=0.0, f=0.0, dp_f=0.0, dp_k=0.0, dp_z=0.0)
            note  = (f"Inline FT {instr_tag or ''}:  −{instr_dp:.3f} psi"
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
            if not fixed_dp and set_p is not None:
                # ΔP back-calculated from the desired valve OUTLET pressure (Set P)
                base_dp = max(0.0, p - set_p)
                src_tag = (f"ΔP back-calculated from Set P = {set_p:.3f} psia (outlet)"
                           if p > set_p else
                           f"Set P {set_p:.3f} ≥ upstream {p:.3f} psia → valve fully open, ΔP = 0")
            elif dest_p is not None:
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
                if fr is not None:
                    mom = _phase_momentum(g_obj, sp, fr, res["v"])

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

        # ── Adiabatic (isenthalpic) temperature change across this fitting ──
        t_out = None
        if (iso_T and T_run is not None and a_feed is not None and fr_full is not None
                and not _is_source(fitting) and not _is_destination(fitting)):
            t_out = _iso_T_out(a_feed, T_run, max(p, p_floor), max(p_out, p_floor))
            if abs(t_out - T_run) >= 0.005:
                note += f" | T {T_run:.2f} → {t_out:.2f} °F (isenthalpic)"

        # ── Control valve / inline instrument: capture the DOWNSTREAM state ──
        instr_info = None
        _kind = _instr_kind(fitting, instr_type)
        if _kind is not None:
            instr_info = dict(
                kind=_kind,
                tag=(instr_tag or txt(row.get("Comp ID"))
                     or f"{row.get('Line No')}-{int(num(row.get('Seq')) or 0)}"),
                tag_given=bool(instr_tag),
                line_no=row.get("Line No"), run_type=row.get("Run Type"),
                circuit=row.get("Circuit"),
                control_type=(_cv_control_type(row) if _kind == "CV" else None),
                bore_in=num(row.get("Bore (in)")),
                line_bore_in=(line_bore_before or last_bore),
                set_p=set_p, fixed_dp=fixed_dp, instr_dp=instr_dp,
                dp_psi=dp_t, dn=None, pv_psia=None, note=note)
            if a_feed is not None and fr_full is not None and fr is not None:
                try:
                    pe_dn = max(p_out, p_floor)
                    if iso_T and t_out is not None:
                        fd_dn = FV.flash(a_feed, pe_dn, temp_f=t_out)
                    else:
                        fd_dn = _flash(a_feed, pe_dn)
                    instr_info["dn"] = _station_props(
                        _scale_flash_result(fd_dn, vap_eff, liq_eff), sp, a_feed,
                        const_map, pe_dn, est_inputs)
                except Exception:
                    instr_info["dn"] = None
                # bubble-point pressure at the inlet temperature (liquid-bearing, for Pv)
                if fr.liq_mass > 0 and _kind in ("CV", "FO", "RO", "FT"):
                    try:
                        T_b = (T_run if (iso_T and T_run is not None) else fr.temp_f)
                        if fr.vap_mass > 0:
                            instr_info["pv_psia"] = max(p, p_floor)       # already saturated
                        elif FV.flash(a_feed, p_floor, temp_f=T_b).vap_mass > 0:
                            lo_, hi_ = p_floor, max(p, p_floor)
                            for _ in range(24):
                                mid_ = 0.5 * (lo_ + hi_)
                                if FV.flash(a_feed, mid_, temp_f=T_b).vap_mass > 0:
                                    lo_ = mid_
                                else:
                                    hi_ = mid_
                            instr_info["pv_psia"] = 0.5 * (lo_ + hi_)
                    except Exception:
                        pass

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
            mach_vap     = mom["mach_vap"],
            rhov2_vap    = mom["rhov2_vap"],
            rhov2_liq    = mom["rhov2_liq"],
            props        = (_station_props(fr, sp, a_feed, const_map, p_eval,
                                           est_inputs)
                            if fr is not None else None),
            instr        = instr_info,
        ))
        flashes.append(fr)
        comps.append(comp_vec)
        if stations[-1].props is not None:
            stations[-1].props["T_out"] = t_out if t_out is not None else (
                fr.temp_f if fr is not None else None)
        if t_out is not None:
            T_run = t_out
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
        default_feed=None,
        default_vap_cf: float = 1.0,
        default_liq_cf: float = 1.0,
        splits_out: dict | None = None,
        const_map: dict | None = None,
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
                and _cv_control_type(row) == "F"
                and num(row.get("Set P (psia)")) is None
                and not num(row.get("Fixed dP (psi)"))):
            di = next((j for j in range(ri + 1, len(block_rows))
                       if _is_destination(block_rows[j].get("Fitting Name"))
                       and num(block_rows[j].get("Set P (psia)")) is not None),
                      None)
            if di is not None:
                targets[ri] = (di, num(block_rows[di].get("Set P (psia)")))

    common = dict(p_floor=p_floor, flash_mode=flash_mode,
                  default_sp=default_sp, injections=injections,
                  default_feed=default_feed, default_vap_cf=default_vap_cf,
                  default_liq_cf=default_liq_cf, splits_out=splits_out,
                  const_map=const_map)
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
        "GVF (local, vol)", "Mach (vapor)",
        "ρv² Vapor (kg/m·s²)", "ρv² Liquid (kg/m·s²)", "Note",
    ]


_PP_NCOL = 33


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
       f"{datetime.now():%d-%b-%Y %H:%M}" + _case_tag(),
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

        vap_rho = ((s.props or {}).get("rho_vap")
                   or (_gas_rho_sp(sp, fr.vap_mw, sp.temp_f, s.p_in_psia)
                       / LBFT3_TO_KGM3)) if fr else None
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
            round(fr.vap_mw, 2)   if (fr and fr.vap_mass > 0) else None,
            round(fr.liq_mw, 2)   if (fr and fr.liq_mass > 0) else None,
            u.disp("rho", vap_rho, 6)     if vap_rho else None,
            u.disp("rho", rho_l_lbft3, 4) if rho_l_lbft3 else None,
            round(gvf, 4)         if gvf is not None else None,
            round(s.mach_vap, 4)  if s.mach_vap  is not None else None,
            round(s.rhov2_vap, 1) if s.rhov2_vap is not None else None,
            round(s.rhov2_liq, 1) if s.rhov2_liq is not None else None,
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
              + [12, 12, 18, 18, 10, 10, 16, 16, 14, 12, 18, 18, 36])
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


def _build_component_detail_noiso(wb, stations, phase, station_lines=None,
                                  line_colors=None):
    """Component_Detail: hydraulic results + full stream property set per fitting.

    Stream properties come from each station's flash: HMB composition when a
    Stream Lookup is used, else the Stream_Composition sheet.  Columns marked †
    (Cp, k, σ, enthalpy) are carried from the HMB stream and blank for manual
    composition.
    """
    u = _u()
    ws = wb.create_sheet("Component_Detail")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = PURPLE

    def D(code, v, nd=4):
        return None if v is None else u.disp(code, v, nd)

    def R(v, nd=4):
        return None if v is None else round(v, nd)

    base = [("Seq", None, 5), ("Comp ID", None, 9), ("Fitting", None, 24),
            (u.hdr("ID", "Lin"), None, 9), (u.hdr("Length", "L"), None, 11),
            (u.hdr("Velocity", "v"), None, 13), ("Reynolds", None, 11),
            ("Friction f", None, 10), (u.hdr("Mix Density", "rho"), None, 14),
            (u.hdr("dP Total", "dP"), None, 13), (u.hdr("P Out", "P"), None, 12),
            ("Mach (vapor)", None, 11),
            ("ρv² Vapor (kg/m·s²)", None, 14), ("ρv² Liquid (kg/m·s²)", None, 14)]
    # (header, getter(props) → value, width, group)
    tot = [
        ("Phase", lambda p: p["phase"], 11),
        (u.hdr("Temperature (in)", "T"), lambda p: D("T", p["T"], 2), 12),
        (u.hdr("Temperature (out)", "T"), lambda p: D("T", p.get("T_out"), 2), 12),
        (u.hdr("Pressure (in)", "P"), lambda p: D("P", p["P"], 3), 13),
        (u.hdr("Total Molar Rate", "molflow"), lambda p: D("molflow", p["n_tot"], 3), 15),
        (u.hdr("Total Mass Rate", "mflow"), lambda p: D("mflow", p["m_tot"], 1), 14),
        (u.hdr("Std Liq Vol Rate", "qvol"), lambda p: D("qvol", p.get("q_std_liq"), 2), 14),
        (u.hdr("Std Vap Vol Rate", "qvol"), lambda p: D("qvol", p["q_std_vap"], 1), 15),
        (u.hdr("Actual Vol Rate", "qvol"), lambda p: D("qvol", p["q_tot"], 2), 14),
        ("Molecular Weight", lambda p: R(p["mw"], 3), 11),
        ("Enthalpy † (MM BTU/HR)", lambda p: R(p.get("h_tot"), 4), 13),
        ("Specific Enthalpy † (BTU/LB)", lambda p: R(p.get("h_spec"), 2), 14),
        ("Liquid Mole Fraction", lambda p: R(p["x_liq"], 5), 11),
        ("Vapor Mole Fraction", lambda p: R(p["beta"], 5), 11),
        ("Liquid Weight Fraction", lambda p: R(p["w_liq"], 5), 11),
        ("Water Mole Fraction", lambda p: R(p.get("x_water"), 5), 11),
        (u.hdr("Actual Density (total)", "rho"), lambda p: D("rho", p["rho_tot"], 5), 14),
        (u.hdr("True Critical Temp", "T"), lambda p: D("T", p.get("tc"), 2), 12),
        (u.hdr("True Critical Pressure", "P"), lambda p: D("P", p.get("pc"), 2), 13),
        ("True Critical Volume (FT3/LB-MOL)", lambda p: R(p.get("vc"), 4), 14),
        ("Acentric Factor", lambda p: R(p.get("omega"), 4), 10),
        ("Watson K (UOP)", lambda p: R(p.get("watson"), 3), 10),
        ("Std API Gravity (°API)", lambda p: R(p.get("api"), 2), 11),
        ("Z Factor (total)", lambda p: R(p["z_tot"], 4), 10),
        ("Z Factor (from K)", lambda p: R(p.get("z_k") if p.get("z_k") is not None
                                          else p["z_vap"], 4), 10),
        (u.hdr("Critical Temp (Kay's Rule)", "T"), lambda p: D("T", p.get("tc_kay"), 2), 13),
        (u.hdr("Critical Press (Kay's Rule)", "P"), lambda p: D("P", p.get("pc_kay"), 2), 13),
        ("Molecular Weight (for EOS)", lambda p: R(p["mw"], 3), 12),
        ("Estimated by correlation (‡)",
         lambda p: ", ".join(dict.fromkeys(p.get("est_keys") or [])) or "—", 24),
    ]
    vap = [
        (u.hdr("Vapor Molar Rate", "molflow"), lambda p: D("molflow", p["n_vap"], 3), 15),
        (u.hdr("Vapor Mass Rate", "mflow"), lambda p: D("mflow", p["m_vap"], 1), 14),
        (u.hdr("Vapor Actual Density", "rho"), lambda p: D("rho", p["rho_vap"], 5), 14),
        ("Vapor Viscosity (cP)", lambda p: R(p["visc_vap"], 5), 12),
        ("Vapor Cp † (BTU/LB-F)", lambda p: R(p.get("cp_vap"), 4), 12),
        ("Vapor Cp/Cv", lambda p: R(p["cpcv_vap"], 4), 10),
        (u.hdr("Vapor Std Vol Rate", "qvol"), lambda p: D("qvol", p.get("q_std_vap_ph"), 1), 15),
        ("Vapor Z (from density)", lambda p: R(p["z_vap"], 4), 11),
        ("Vapor Z (from K)", lambda p: R(p.get("z_k") if p.get("z_k") is not None
                                         else p["z_vap"], 4), 10),
        ("Vapor Thermal Cond † (BTU/HR-FT-F)", lambda p: R(p.get("k_vap"), 5), 15),
    ]
    liq = [
        (u.hdr("Liquid Molar Rate", "molflow"), lambda p: D("molflow", p["n_liq"], 3), 15),
        (u.hdr("Liquid Mass Rate", "mflow"), lambda p: D("mflow", p["m_liq"], 1), 14),
        (u.hdr("Liquid Actual Density", "rho"), lambda p: D("rho", p["rho_liq"], 4), 14),
        (u.hdr("Liquid Std Density", "rho"), lambda p: D("rho", p.get("rho_std_liq"), 4), 14),
        ("Liquid Viscosity (cP)", lambda p: R(p["visc_liq"], 5), 12),
        ("Liquid Cp † (BTU/LB-F)", lambda p: R(p.get("cp_liq"), 4), 12),
        ("Liquid Surface Tension † (DYNE/CM)", lambda p: R(p.get("sigma"), 4), 14),
        ("Liquid Thermal Cond † (BTU/HR-FT-F)", lambda p: R(p.get("k_liq"), 5), 15),
        (u.hdr("Liquid Std LV Rate", "qvol"), lambda p: D("qvol", p.get("q_std_liq_liq"), 3), 14),
    ]
    groups = [(tot, "375623"), (vap, STEEL), (liq, PURPLE)]
    ncol = len(base) + sum(len(g) for g, _ in groups)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    _c(ws, 1, 1,
       f"COMPONENT DETAIL   |   Phase model: {phase}   |   per-fitting stream properties "
       "from the local flash (HMB composition, or Stream_Composition when no HMB stream).  "
       "Viscosity, Z, Cp/Cv, Cp, k, σ, ρ and enthalpy: your input / the HMB value when "
       "present (†), otherwise ESTIMATED from Tb, SG, MW, Tc, Pc, ω by generalised "
       "correlations (listed in column ‡; enthalpy basis = ideal gas @ 77°F).  Critical "
       "properties = mole-averaged (Kay); Std liquid basis 60°F, volume-additive.",
       bg=NAVY, fg=WHITE, sz=10, bold=True, wrap=True)
    ws.row_dimensions[1].height = 42

    _hdr(ws, [h for h, _, _ in base] + [h for g, _ in groups for h, _, _ in g], row=2)
    ci = len(base) + 1
    for g, colr in groups:
        for _ in g:
            ws.cell(2, ci).fill = _fill(colr)
            ci += 1
    ws.row_dimensions[2].height = 54
    for c in range(1, ncol + 1):
        ws.cell(2, c).alignment = Alignment(wrap_text=True, horizontal="center",
                                            vertical="center")
    ws.freeze_panes = "D3"
    multi = bool(station_lines and line_colors and len(set(station_lines)) > 1)
    r = 3
    prev = None
    for i, st in enumerate(stations):
        cur = station_lines[i] if (station_lines and i < len(station_lines)) else None
        if multi and cur and cur != prev:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
            _c(ws, r, 1, f"  ▶  LINE: {cur}", bg=line_colors.get(cur, LGRAY),
               fg=NAVY, sz=9, bold=True)
            r += 1
            prev = cur
        bg = (line_colors.get(cur, LGRAY) if (multi and cur)
              else (LGRAY if i % 2 else WHITE))
        vals = [st.seq, st.comp_id, st.fitting, u.disp("Lin", st.id_in),
                u.disp("L", st.length_ft), u.disp("v", st.v_fts), st.Re, st.f,
                u.disp("rho", st.rho_lbft3, 5), u.disp("dP", st.dp_total_psi, 5),
                u.disp("P", st.p_out_psia),
                R(st.mach_vap), R(st.rhov2_vap, 1), R(st.rhov2_liq, 1)]
        pr = st.props
        for g, _ in groups:
            for _h, fn, _w in g:
                try:
                    vals.append(fn(pr) if pr else None)
                except Exception:
                    vals.append(None)
        for c, v in enumerate(vals, 1):
            _c(ws, r, c, "" if v is None else v, bg=bg, sz=9,
               ha="right" if isinstance(v, (int, float)) else "left")
        r += 1
    widths = [w for _, _, w in base] + [w for g, _ in groups for _, _, w in g]
    for c, w in enumerate(widths, 1):
        cw(ws, c, w)


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
                round(fr.vap_mw, 2) if fr.vap_mass > 0 else None,
                round(fr.liq_mw, 2) if fr.liq_mass > 0 else None] if fr else [
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

    # ── Extra blocks: ppm (total) and vapour / liquid phase compositions ──
    def _block(title, getter, nd, tot_label=None, total_fn=None):
        nonlocal r
        r += 2
        ws.merge_cells(f"A{r}:{last_col}{r}")
        _c(ws, r, 1, title, bg=GRNHDR, fg=WHITE, sz=9, bold=True)
        r += 1
        _station_header(r)
        r += 1
        for nm in names:
            _c(ws, r, 1, nm, bg=LGRAY, sz=8, ha="left")
            for j in range(len(stations)):
                v = getter(j, nm)
                _c(ws, r, 2 + j, round(v, nd) if v else "", sz=8, ha="right")
            r += 1
        if tot_label:
            _c(ws, r, 1, tot_label, bg=GREEN, sz=8, bold=True)
            for j in range(len(stations)):
                v = total_fn(j)
                _c(ws, r, 2 + j, round(v, nd) if v else "", bg=GREEN, sz=8,
                   bold=True, ha="right")

    def _tot_mole_ppm(j, nm):
        return (1e6 * comps[j].get(nm, 0.0) / mole_tot[j]) if mole_tot[j] else None

    def _tot_mass_ppm(j, nm):
        return (1e6 * mass_cols[j].get(nm, 0.0) / col_tot[j]) if col_tot[j] else None
    _block("MOLE FRACTION  (ppm, total)", _tot_mole_ppm, 2)
    _block("MASS FRACTION  (ppm, total)", _tot_mass_ppm, 2)

    def _pp(j):
        return (stations[j].props or {}) if j < len(stations) else {}

    def _phase_mole(which):
        def g(j, nm):
            return 100.0 * (_pp(j).get(which) or {}).get(nm, 0.0) or None
        return g

    def _phase_mass(which):
        def g(j, nm):
            dct = _pp(j).get(which) or {}
            tot = sum(v * mw_map.get(k, 0.0) for k, v in dct.items())
            return (100.0 * dct.get(nm, 0.0) * mw_map.get(nm, 0.0) / tot) if tot else None
        return g
    _block("VAPOR PHASE — MOLE FRACTION  (mol %)", _phase_mole("comp_y"), 4,
           "TOTAL (mol %)", lambda j: 100.0 if _pp(j).get("comp_y") else None)
    _block("VAPOR PHASE — MASS FRACTION  (mass %)", _phase_mass("comp_y"), 4,
           "TOTAL (mass %)", lambda j: 100.0 if _pp(j).get("comp_y") else None)
    _block("LIQUID PHASE — MOLE FRACTION  (mol %)", _phase_mole("comp_x"), 4,
           "TOTAL (mol %)", lambda j: 100.0 if _pp(j).get("comp_x") else None)
    _block("LIQUID PHASE — MASS FRACTION  (mass %)", _phase_mass("comp_x"), 4,
           "TOTAL (mass %)", lambda j: 100.0 if _pp(j).get("comp_x") else None)

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
# ════════════════════════════════════════════════════════════════════════
#  FLOW CASES  (Normal / Max / Min …) — every inlet flow × a multiplier, whole run repeated
# ════════════════════════════════════════════════════════════════════════
FLOW_CASES_SHEET = "Flow_Cases"
_RUN_CTX: dict = {"name": "Normal", "mult": 1.0, "n": 1}     # flow case being written (for titles)


def _case_tag() -> str:
    """' | Flow case: Max (110 %)' for multi-case runs, '' for a single normal case."""
    if _RUN_CTX.get("n", 1) <= 1:
        return ""
    return f"   |   Flow case: {_RUN_CTX['name']} ({_RUN_CTX['mult'] * 100:g} % of normal)"


def _read_flow_cases(input_path) -> list:
    """Flow_Cases sheet → [(name, factor)] with Normal (100 %) always first.  Blank rows are
    skipped; a missing sheet gives the single Normal case."""
    cases = []
    try:
        wb = load_workbook(input_path, read_only=True, data_only=True)
    except Exception:
        return [("Normal", 1.0)]
    try:
        if FLOW_CASES_SHEET in wb.sheetnames:
            for r in wb[FLOW_CASES_SHEET].iter_rows(min_row=3, values_only=True):
                if not r or r[0] is None or not str(r[0]).strip():
                    continue
                pct = num(r[1]) if len(r) > 1 else None
                if pct is None or pct <= 0:
                    continue
                cases.append((str(r[0]).strip(), pct / 100.0))
    finally:
        wb.close()
    if not cases:
        return [("Normal", 1.0)]
    if not any(abs(m_ - 1.0) < 1e-9 for _, m_ in cases):
        cases.insert(0, ("Normal", 1.0))
    seen, out = set(), []
    for n_, m_ in cases:
        if n_.upper() in seen:
            continue
        seen.add(n_.upper())
        out.append((n_, m_))
    out.sort(key=lambda t: 0 if abs(t[1] - 1.0) < 1e-9 else 1)      # Normal first, then as listed
    return out


def _scale_rows(rows: list, m: float) -> list:
    """Copy of the input rows with every typed / composition flow × m."""
    out = []
    for r in rows:
        r2 = dict(r)
        for k in ("Vapor Mass Flow", "Liq Mass Flow (lb/h)"):
            if r2.get(k) is not None:
                r2[k] = r2[k] * m
        if r2.get("_comp"):
            r2["_comp"] = [dict(c, flow=(c["flow"] * m if c.get("flow") is not None else None))
                           for c in r2["_comp"]]
        out.append(r2)
    return out


def _scale_sp(sp, m: float):
    if sp is None or m == 1.0:
        return sp
    f = {a: getattr(sp, a) * m
         for a in ("total_mass", "vap_mass", "liq_mass", "total_std_liq", "total_std_vap",
                   "vap_act_vol", "liq_act_rate") if getattr(sp, a, None) is not None}
    return replace(sp, **f)


def _scale_feed(feed, m: float):
    if feed is None or m == 1.0:
        return feed
    return replace(feed, total_mass_lbhr=feed.total_mass_lbhr * m,
                   total_moles_lbmolhr=feed.total_moles_lbmolhr * m)


def _build_flow_cases_sheet(wb) -> None:
    ws = wb.create_sheet(FLOW_CASES_SHEET)
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:C1")
    _c(ws, 1, 1,
       "FLOW CASES — every inlet flow (HMB stream, manual, Stream_Composition) is multiplied by the case "
       "percentage and the WHOLE hydraulics is repeated: one workbook per circuit per case.  The CV Datasheet "
       "lists Normal / Max / Min side by side.  Normal (100 %) is always run.  Blank rows are skipped.",
       bg=NAVY, fg=WHITE, sz=10, bold=True, wrap=True)
    ws.row_dimensions[1].height = 54
    _hdr(ws, ["Flow case", "Flow multiplier (% of normal)", "Note"], row=2)
    rows = [("Normal", 100, "design / normal flow"), ("Maximum", 110, "e.g. 1.1 × normal"),
            ("Minimum", 60, "turndown, e.g. 0.6 × normal")]
    for i, rowv in enumerate(rows):
        for c, v in enumerate(rowv, 1):
            _c(ws, 3 + i, c, v, bg=AMBER, sz=10, bold=(c == 1))
    for r in range(6, 14):
        for c in range(1, 4):
            _c(ws, r, c, None, bg=AMBER, sz=10)
    for c, w in enumerate([22, 30, 44], 1):
        cw(ws, c, w)
    ws.freeze_panes = "A3"


def _make_stream_resolver(default_hmb: str, default_case: str, map_path, alias: dict | None = None,
                          mult: float = 1.0):
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

        sim = _alias_stream(stream, alias)
        sp = HE.load_stream_props(hmb, sim, case)
        if sp is None and sim != stream:              # not in the HMB under the mapped name → try as typed
            sim = stream
            sp = HE.load_stream_props(hmb, sim, case)
        feed = None
        if sp is not None:
            try:
                is_case = HE.is_proii_export(hmb)
                feed = FV.read_feed(hmb, sim, case, sp, is_casesheet=is_case)
            except Exception as exc:
                print(f"  (feed load failed for {stream}@{case}: {exc})")
        cache[key] = (_scale_sp(sp, mult), _scale_feed(feed, mult))
        return cache[key]

    return resolver


def _circuit_case(cid, c, all_rows, input_path, hmb_path, case, map_path, flash_mode,
                  stream_name, out_path, alias, mult, fcase, n_cases):
    """March ONE circuit at ONE flow case (all inlet flows × ``mult``).

    Returns ``(items, write)`` — the control-valve / instrument capture and a writer that builds
    the workbook once every flow case of the circuit has been computed — or None to skip."""
    sk  = _alias_stream(stream_name or c["stream_lookup"], alias)
    p0  = c["start_p"]
    lns = c["line_order"]

    # ── Per-row stream resolver (HMB File / Case / manual) ────────────
    def_hmb  = c["hmb_file"] or hmb_path
    def_case = c["case"] or case
    resolver = _make_stream_resolver(def_hmb, def_case, map_path, alias, mult)
    # HMB COMP_CONSTANTS (MW/Tc/Pc/ω) — only needed for Stream_Composition
    try:
        cmap = FV._read_comp_constants(def_hmb)
    except Exception:
        cmap = None

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
            return None
        is_case = HE.is_proii_export(def_hmb)
        feed = FV.read_feed(def_hmb, sk, def_case, sp, is_casesheet=is_case)
        sp, feed = _scale_sp(sp, mult), _scale_feed(feed, mult)
        if feed is None:
            print("  WARNING: no composition for default stream — flash uses "
                  "frozen split where applicable.")
        else:
            print(f"  Flash feed: {feed.n} components | β₀={feed.beta_ref:.4f} "
                  f"| P_ref={feed.p_ref_psia:.2f} psia | mode: {flash_mode}")
    else:
        # No HMB stream at circuit level → manual mode.  Use the first main
        # row carrying manual data as the representative StreamProps.
        comp_row = next((r for r in c["main_rows"] if r.get("_comp")), None)
        if comp_row is not None:
            try:
                pl0 = _comp_payload(comp_row, p0 or 100.0, cmap)
            except ValueError as exc:
                raise SystemExit(f"  ERROR (circuit {cid}, Line "
                                 f"{comp_row.get('Line No')}, Seq "
                                 f"{comp_row.get('Seq')}): {exc}")
            sp, feed = pl0["sp"], pl0["feed"]
            sk = "(manual composition)"
            print(f"  Manual composition mode — {feed.n} components from "
                  f"{_COMP_SHEET} | K basis: {feed.k_basis} | "
                  f"β₀={feed.beta_ref:.4f}")
        else:
            sp = next((m for m in (_streamprops_from_manual(r)
                                   for r in c["main_rows"]) if m), None)
            if sp is None:
                print(f"  WARNING: circuit '{cid}' has no Stream Lookup and no "
                      f"manual properties — skipped.")
                return None
            sk = "(manual)"
            print("  Manual stream mode — properties from yellow input cells.")

    if n_cases > 1:
        print(f"\n  ── Flow case {fcase}: {mult * 100:g} % of normal flow ──")
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
                          default_sp=sp, splits_out=splits_out,
                          const_map=cmap)

    # ── March branch blocks FIRST (independent sub-runs) ─────────────
    # A branch whose outlet matches a Main 'Tee Join Branch Flow N' row on
    # the same line is INJECTED into the Main at that row (its outlet flow +
    # composition merge in); others stay standalone Branch sheets.
    # A branch whose INLET matches a Main 'Tee Split Branch Flow N' row
    # (via "Upstream Line No" / "Upstream Seq") is SEEDED from that row's
    # captured split-off flow/composition instead of marching standalone.
    branch_runs: list = []            # (label, sts, fls, cmps, b_lns)
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
            if split_cap is None and not b0.get("_Up_defaulted"):
                print(f"  WARNING: branch (line {b0.get('Line No')}) specifies "
                      f"Upstream Line No '{up_line}' but no matching "
                      f"'Tee Split Branch Flow N' output was found on Main.")

        solve_kwargs = dict(flash_mode=flash_mode, default_sp=sp,
                            const_map=cmap)
        if split_cap is not None:
            solve_kwargs.update(default_feed=split_cap["feed"],
                                 default_vap_cf=split_cap["vap_cf"],
                                 default_liq_cf=split_cap["liq_cf"])
            if split_cap.get("sp") is not None:
                solve_kwargs["default_sp"] = split_cap["sp"]
        sts, fls, cmps, _bmw = build_profile_solved(block, resolver, bp, **solve_kwargs)
        b_lns = [r["Line No"] for r in block]
        label = f"Branch_{bi+1}" if len(branch_blocks) > 1 else "Branch"
        branch_runs.append((label, sts, fls, cmps, b_lns))
        if split_cap is not None:
            print(f"  Branch '{label}' (line {b0.get('Line No')}) splits off "
                  f"Main line {up_line} at Tee Split Branch Flow {split_cap['n']}.")

        bline  = block[0].get("Line No") if block else None
        line_branch_count[bline] = line_branch_count.get(bline, 0) + 1
        join_n = line_branch_count[bline]
        # A Main 'Tee Join Branch Flow N' row may name the joining Branch
        # explicitly through its "Upstream Line No" cell; otherwise the Nth
        # join row sharing the Branch's Line No is used.
        join_idx = next(
            (mi for mi, mr in enumerate(main_rows)
             if _is_tee_join(mr.get("Fitting Name")) is not None
             and mr.get("Upstream Line No") == bline
             and mr.get("Run Type") == "Main"),
            None)
        if join_idx is None:
            join_idx = next(
                (mi for mi, mr in enumerate(main_rows)
                 if mr.get("Line No") == bline
                 and _is_tee_join(mr.get("Fitting Name")) == join_n),
                None)
        if join_idx is not None and fls and fls[-1] is not None:
            out_fr = fls[-1]
            b_sp, b_feed = resolver(block[0]) if block else (None, None)
            if b_feed is None and block and block[0].get("_comp"):
                _bpl = _comp_payload(block[0], bp, cmap)
                b_sp, b_feed = _bpl["sp"], _bpl["feed"]
            injections[join_idx] = {
                "vap_mass": out_fr.vap_mass,
                "liq_mass": out_fr.liq_mass,
                "vmol":     out_fr.vap_moles,
                "lmol":     out_fr.liq_moles,
                "comp":     cmps[-1] if cmps else {},
                "feed":     b_feed,
                "sp":       b_sp,
                "T_out":    ((sts[-1].props or {}).get("T_out") if sts else None),
                "label":    label,
            }
            print(f"  Branch '{label}' (line {bline}) joins Main at "
                  f"Tee Join Branch Flow {join_n}.")

    # ── March the Main with branch injections ────────────────────────
    main_stations, main_flashes, main_comps, main_mw = build_profile_solved(
        main_rows, resolver, p_start, flash_mode=flash_mode,
        default_sp=sp, injections=injections, const_map=cmap)

    # Manual circuit with joins: the screening sheets (FIV / AIV / regime / maps)
    # should describe the MIXED stream at the end of the line, not the first inlet.
    if (injections and main_stations and feed is None
            and (main_stations[-1].props or None)):
        pe = main_stations[-1].props
        try:
            sp = replace(
                sp, vap_mass=pe["m_vap"], liq_mass=pe["m_liq"],
                total_mass=pe["m_vap"] + pe["m_liq"],
                vap_mw=((pe["m_vap"] / pe["n_vap"]) if pe["n_vap"] else sp.vap_mw),
                mol_weight=((pe["m_vap"] / pe["n_vap"]) if pe["n_vap"] else sp.mol_weight),
                vap_visc=pe.get("visc_vap") or sp.vap_visc,
                liq_density=pe.get("rho_liq") or sp.liq_density,
                liq_visc=pe.get("visc_liq") or sp.liq_visc,
                vap_z=pe.get("z_vap") or sp.vap_z,
                vap_cp_cv=pe.get("cpcv_vap") or sp.vap_cp_cv)
            phase = HE.classify_phase(sp)

            def gdf(p_, _sp=sp):
                return (HE.vapor_density_kgm3(_sp, p_) / LBFT3_TO_KGM3
                        if _sp.mol_weight else (_sp.vap_density or 0.0))
        except Exception:
            pass


    # station_lines: map each station back to its source Line No
    station_lines = [r["Line No"] for r in main_rows if r["Run Type"] == "Main"]
    runs_all = ([("Main", main_stations, station_lines)]
                + [(lb_, sts_, bl_) for (lb_, sts_, _f, _c2, bl_) in branch_runs])
    items = _collect_instruments(runs_all)


    def _write(case_items):
        _RUN_CTX.update(name=fcase, mult=mult, n=n_cases)

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
            _build_component_detail_noiso(wb, main_stations, phase,
                                          station_lines=station_lines,
                                          line_colors=lc)
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

        # ── Line Data / CV Datasheet / Inline Instruments ─────────────────
        try:
            _build_line_data_sheet(wb, runs_all, sp, gdf, cid, sk)
            _idata = _read_instrument_data(input_path)
            _ins = [x for x in items if x[1].instr["kind"] != "CV"]
            _order = list(case_items)
            _cvs = sorted(
                ((lb_, st_, ln_, fc_, case_items[fc_]["mult"])
                 for fc_ in _order for (lb_, st_, ln_) in case_items[fc_]["items"]
                 if st_.instr["kind"] == "CV"),
                key=lambda t: (t[0], t[1].seq, t[1].instr["tag"], _order.index(t[3])))
            if _cvs:
                _build_cv_datasheet(wb, _cvs, _idata, cid, sk)
            if _ins:
                _build_instruments_sheet(wb, _ins, _idata, cid, sk)
        except Exception as exc:                      # never lose the run over a datasheet
            import traceback
            traceback.print_exc()
            print(f"  (Line Data / instrument sheets skipped: {exc})")

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
                  "Line Data", "Inline Instruments", "CV Datasheet",
                  "FIV_EI_T2.2", "AIV", "Two_Phase_Regime"]
                 + fp_map_sheets
                 + ["Input_Pipeline", "README"])
        for name in reversed(order):
            if name in [s.title for s in wb.worksheets]:
                wb.move_sheet(name, offset=-len(wb.worksheets))
        if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
            del wb["Sheet"]
        # keep the line / instrument sheets next to Stream_Props
        _anchor = wb.sheetnames.index("Stream_Props") if "Stream_Props" in wb.sheetnames else None
        if _anchor is not None:
            _mv = [wb[n_] for n_ in ("Line Data", "Inline Instruments", "CV Datasheet")
                   if n_ in wb.sheetnames]
            for w_ in _mv:
                wb._sheets.remove(w_)
            _anchor = wb.sheetnames.index("Stream_Props") + 1
            wb._sheets[_anchor:_anchor] = _mv

        safe_f = re.sub(r"[^A-Za-z0-9_-]", "_", str(fcase))
        if out_path is None:
            safe_c = re.sub(r"[^A-Za-z0-9_-]", "_", str(cid))
            safe_s = re.sub(r"[^A-Za-z0-9_-]", "_", str(sk))
            op = (f"hydraulics_flash_noiso_{safe_c}_{safe_s}"
                  + (f"_{safe_f}" if n_cases > 1 else "") + ".xlsx")
        elif n_cases > 1:
            _b, _e = os.path.splitext(out_path)
            op = f"{_b}_{safe_f}{_e or '.xlsx'}"
        else:
            op = out_path

        _apply_light_theme(wb)
        wb.save(op)
        _res = (op, main_stations, main_flashes, sp)

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
        return _res


    return items, _write


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

    Flow cases (Flow_Cases sheet: Normal 100 %, Max 110 %, Min 60 % …): every inlet flow is
    multiplied by the case factor and the WHOLE hydraulics is repeated — one workbook per
    circuit per case; the CV Datasheet in each shows all cases side by side.

    Returns list of (out_xlsx, main_stations, main_flashes, sp) — one per circuit per flow case.
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
    alias = _stream_alias(input_path, map_path)
    flow_cases = _read_flow_cases(input_path)
    if len(flow_cases) > 1:
        print("  Flow cases: " + ", ".join(f"{n_} {m_ * 100:g} %" for n_, m_ in flow_cases))
    target_ids = [circuit_id] if circuit_id else list(circuits.keys())
    results = []

    for cid in target_ids:
        if cid not in circuits:
            print(f"  WARNING: circuit '{cid}' not found — skipped.")
            continue
        built = []
        for fname, fmult in flow_cases:
            circ_c = circuits if fmult == 1.0 else _group_circuits(_scale_rows(all_rows, fmult))
            r_ = _circuit_case(cid, circ_c[cid], all_rows, input_path, hmb_path, case, map_path,
                               flash_mode, stream_name, out_path, alias, fmult, fname, len(flow_cases))
            if r_ is None:
                built = []
                break
            built.append((fname, fmult, r_))
        if not built:
            continue
        case_items = {fn_: dict(mult=fm_, items=r_[0]) for fn_, fm_, r_ in built}
        for fn_, fm_, r_ in built:
            results.append(r_[1](case_items))

    return results



# ════════════════════════════════════════════════════════════════════════
#  Control-valve datasheet · Inline Instruments · Line Data
# ════════════════════════════════════════════════════════════════════════
#  Instr Type drop-down:  PT · TT · FT (inline) · FO (orifice FT) · RO     Fitting = "Control Valve" → CV
#  Optional per-tag data (thermowell, orifice bore, valve rating) lives in the
#  'Instrument_Data' input sheet; anything not entered falls back to a flagged default.
_INSTR_DATA_SHEET = "Instrument_Data"
_INSTR_DATA_HEADERS = [
    "Tag", "Description", "Orifice / throat d (in)", "TW insertion U (in)",
    "TW root OD (in)", "TW tip OD (in)", "TW bore (in)", "TW material",
    "CV rated Cv", "CV FL", "CV xT", "CV characteristic", "CV rangeability",
    "CV curve points", "Notes"]

# thermowell materials: E at 70 °F (psi), dE/dT (psi/°F), density (lb/in³)
_TW_MATERIALS = {
    "SS316":  (28.3e6, -5660.0, 0.290), "SS304": (28.3e6, -5660.0, 0.289),
    "SS316L": (28.3e6, -5660.0, 0.290), "CS":    (29.5e6, -4500.0, 0.284),
    "C276":   (29.8e6, -5000.0, 0.321), "MONEL": (26.0e6, -4000.0, 0.319),
    "DUPLEX": (29.0e6, -5200.0, 0.280), "INC625": (30.1e6, -5300.0, 0.305),
    "1.25CR": (30.0e6, -4800.0, 0.284), "2.25CR": (30.0e6, -4800.0, 0.284),
    "5CR":    (30.0e6, -4800.0, 0.280), "9CR":    (30.0e6, -4800.0, 0.280),
    "ALLOY825": (28.3e6, -5000.0, 0.294), "ALLOY600": (31.0e6, -5000.0, 0.304),
    "ALLOY20": (28.0e6, -5000.0, 0.292),
}

# Thermowell defaults PER PIPING CLASS.  Key = piping class (the part of 'Piping Spec' before
# the '/', e.g. 'G9A-3').  Any of: tw_root, tw_tip, tw_bore, tw_u (all inches), tw_mat (a key of
# _TW_MATERIALS).  Edit this table for the project; a tag typed in 'Instrument_Data' still wins.
# When a class is not listed, the thermowell MATERIAL follows the class's MOC (see _TW_MOC_RULES)
# and the stem geometry falls back to the generic screening default.
THERMOWELL_BY_CLASS: dict[str, dict] = {
    # "G9A-3": {"tw_root": 1.0, "tw_tip": 0.75, "tw_bore": 0.26, "tw_u": 6.0, "tw_mat": "SS316"},
}
_TW_MOC_RULES = [          # (text found in the class MOC, upper-case) → thermowell material
    ("C-276", "C276"), ("ALLOY 400", "MONEL"), ("ALLOY 825", "ALLOY825"),
    ("ALLOY 600", "ALLOY600"), ("ALLOY 20", "ALLOY20"), ("DUPLEX", "DUPLEX"), ("2205", "DUPLEX"),
    ("316", "SS316"), ("304", "SS304"), ("321", "SS304"), ("347", "SS304"), ("317", "SS316"),
    ("9% CR", "9CR"), ("5% CR", "5CR"), ("2 1/4% CR", "2.25CR"), ("1 1/4% CR", "1.25CR"),
    ("1 1/4%CR", "1.25CR"), ("CARBON STEEL", "CS"),
]


def _tw_class_data(spec) -> dict:
    """Thermowell defaults for a piping spec: THERMOWELL_BY_CLASS entry, else MOC-derived material."""
    cls = (spec or "").split("/")[0].strip().upper()
    d = dict(THERMOWELL_BY_CLASS.get(cls, {}))
    if d:
        d["_src"] = f"THERMOWELL_BY_CLASS[{cls}]"
    if "tw_mat" not in d:
        pc = _CLASS_BY_NAME.get(cls)
        moc = (getattr(pc, "moc", "") or "").upper() if pc else ""
        for txt_, mat in _TW_MOC_RULES:
            if txt_ in moc:
                d["tw_mat"] = mat
                d["_mat_src"] = f"piping class {cls} MOC '{pc.moc}'"
                break
    return d


def _read_instrument_data(path) -> dict:
    """Instrument_Data sheet → {TAG_UPPER: {field: value}} (empty when absent)."""
    out: dict = {}
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return out
    try:
        if _INSTR_DATA_SHEET not in wb.sheetnames:
            return out
        rows = list(wb[_INSTR_DATA_SHEET].iter_rows(values_only=True))
    finally:
        wb.close()
    if len(rows) < 3:
        return out
    key = {"tag": "tag", "description": "desc", "orifice / throat d": "d_or",
           "tw insertion u": "tw_u", "tw root od": "tw_root", "tw tip od": "tw_tip",
           "tw bore": "tw_bore", "tw material": "tw_mat", "cv rated cv": "cv_rated",
           "cv fl": "fl", "cv xt": "xt", "cv characteristic": "cv_char",
           "cv rangeability": "cv_range", "cv curve points": "cv_curve", "notes": "notes"}
    hdr = {}
    for i, h in enumerate(rows[1]):
        if h:
            k = re.sub(r"\s*[\(\[].*$", "", str(h)).strip().lower()
            if k in key:
                hdr[key[k]] = i
    if "tag" not in hdr:
        return out
    for r in rows[2:]:
        tg = txt(r[hdr["tag"]]) if hdr["tag"] < len(r) else None
        if not tg:
            continue
        d = {}
        for f, i in hdr.items():
            if f == "tag" or i >= len(r):
                continue
            d[f] = (txt(r[i]) if f in ("desc", "tw_mat", "notes", "cv_char", "cv_curve") else num(r[i]))
        out[tg.upper()] = d
    return out


def _build_instrument_data_sheet(wb) -> None:
    ws = wb.create_sheet(_INSTR_DATA_SHEET)
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:O1")
    _c(ws, 1, 1,
       "INSTRUMENT DATA (optional) — one row per Instr Tag (or Comp ID) used on Pipeline_Input.  "
       "TT: thermowell geometry for the wake-frequency check.  FO / RO: orifice bore d.  "
       "Control Valve: rated Cv (100 % travel), FL, xT and the Cv CURVE — either a characteristic "
       "(Equal %, Linear, Quick opening + rangeability) or measured points 'travel %:Cv' separated by ; "
       "e.g.  10:4.2; 25:11; 50:30; 75:58; 100:85.  Points win over the characteristic.  Blank cells use a flagged default.",
       bg=NAVY, fg=WHITE, sz=10, bold=True, wrap=True)
    ws.row_dimensions[1].height = 40
    _hdr(ws, _INSTR_DATA_HEADERS, row=2)
    ex = [("TT-101", "Reactor effluent TT", None, 6.0, 1.0, 0.75, 0.26, "SS316", None, None, None, None, None, None, "example"),
          ("RO-201", "Restriction orifice", 1.25, None, None, None, None, None, None, None, None, None, None, None, "example"),
          ("FV-301", "Feed control valve", None, None, None, None, None, None, 85, 0.9, 0.7, "Equal %", 50, None, "example"),
          ("FV-302", "Reflux control valve", None, None, None, None, None, None, 120, 0.85, 0.72, None, None,
           "10:6; 20:11; 30:17; 40:25; 50:35; 60:50; 70:68; 80:90; 90:112; 100:120", "example — measured curve")]
    for i, rowv in enumerate(ex):
        for c, v in enumerate(rowv, 1):
            _c(ws, 3 + i, c, v, bg=AMBER, sz=9, italic=True, fg="808080")
    for r in range(6, 40):
        for c in range(1, 16):
            _c(ws, r, c, None, bg=AMBER, sz=9)
    for c, w in enumerate([14, 24, 16, 14, 12, 12, 11, 12, 11, 8, 8, 16, 12, 46, 24], 1):
        cw(ws, c, w)
    dvc = DataValidation(type="list", formula1='"Equal %,Linear,Quick opening"', allow_blank=True)
    ws.add_data_validation(dvc)
    dvc.add("L3:L39")
    dv = DataValidation(type="list", formula1='"' + ",".join(_TW_MATERIALS) + '"',
                        allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("H3:H39")
    ws.freeze_panes = "A3"


# ── small helpers ────────────────────────────────────────────────────────
_GPM_PER_FT3H = 0.124675          # 1 ft3/h → US gpm
_LBF_PER_N = 1.0 / 4.448222


def _pr_comp_pct(pr: dict) -> dict:
    """Mole % of key species — total, vapour, liquid — from a station property set."""
    if not pr or not pr.get("comp_y") and not pr.get("comp_x"):
        return {}
    nv, nl = pr.get("n_vap") or 0.0, pr.get("n_liq") or 0.0
    N = nv + nl
    cy, cx = pr.get("comp_y") or {}, pr.get("comp_x") or {}
    out: dict = {}
    for key in ("H2", "H2S", "NH3", "CO2", "H2O"):
        yv = sum(v for n, v in cy.items() if _pr_key(n) == key)
        xv = sum(v for n, v in cx.items() if _pr_key(n) == key)
        out[key] = dict(vap=yv * 100.0 if nv > 0 else None,
                        liq=xv * 100.0 if nl > 0 else None,
                        tot=((nv * yv + nl * xv) / N * 100.0) if N > 0 else None)
    return out


def _ph(pr: dict, phase: str, key: str):
    """Property of one phase, blank when that phase is absent from the stream."""
    m = (pr or {}).get("m_vap" if phase == "v" else "m_liq") or 0.0
    return (pr or {}).get(key) if m > 0 else None


def _flow_state(st) -> dict:
    """Inlet flows, velocities, ρv² of a station (SI where noted)."""
    pr = st.props or {}
    d_in = st.id_in or 0.0
    area_ft2 = math.pi / 4.0 * (d_in / 12.0) ** 2
    qv, ql = pr.get("q_vap") or 0.0, pr.get("q_liq") or 0.0       # ft3/h
    mv, ml = pr.get("m_vap") or 0.0, pr.get("m_liq") or 0.0
    q = qv + ql
    out = dict(area_ft2=area_ft2, qv=qv, ql=ql, q=q, mv=mv, ml=ml, mt=mv + ml)
    if area_ft2 > 0:
        out.update(vsg=qv / 3600.0 / area_ft2, vsl=ql / 3600.0 / area_ft2,
                   vm=q / 3600.0 / area_ft2)        # ft/s
    else:
        out.update(vsg=None, vsl=None, vm=None)
    rg, rl = pr.get("rho_vap"), pr.get("rho_liq")
    out["rhov2_m"] = None
    if out["vm"] and out["q"] > 0 and (mv + ml) > 0:
        rho_h = (mv + ml) / q                                   # lb/ft3 homogeneous
        out["rho_h"] = rho_h
        out["rhov2_m"] = rho_h * LBFT3_TO_KGM3 * (out["vm"] * FT_TO_M) ** 2   # kg/m·s²
    else:
        out["rho_h"] = None
    out["rhov2_g"] = (rg * LBFT3_TO_KGM3 * (out["vsg"] * FT_TO_M) ** 2
                      if (rg and out["vsg"]) else None)
    out["rhov2_l"] = (rl * LBFT3_TO_KGM3 * (out["vsl"] * FT_TO_M) ** 2
                      if (rl and out["vsl"]) else None)
    return out


def _sound_speed_ms(pr: dict) -> float | None:
    """Vapour speed of sound (m/s) from γ, Z, T, MW of the station's vapour."""
    try:
        nv, mv = pr.get("n_vap") or 0.0, pr.get("m_vap") or 0.0
        if nv <= 0 or mv <= 0:
            return None
        mw = mv / nv
        k = pr.get("cpcv_vap") or 1.3
        z = pr.get("z_vap") or 1.0
        T = ((pr.get("T") if pr.get("T") is not None else 60.0) + 459.67) / 1.8
        return math.sqrt(k * z * 8.314462618 * T / (mw * 1e-3))
    except Exception:
        return None


def _flashing_info(pr_in: dict, pr_dn: dict | None) -> dict:
    """Did extra vapour form across the device?  (needs a downstream flash)."""
    if not pr_dn or not pr_in:
        return dict(known=False)
    mt_i = (pr_in.get("m_vap") or 0.0) + (pr_in.get("m_liq") or 0.0)
    mt_d = (pr_dn.get("m_vap") or 0.0) + (pr_dn.get("m_liq") or 0.0)
    wv_i = (pr_in.get("m_vap") or 0.0) / mt_i if mt_i > 0 else 0.0
    wv_d = (pr_dn.get("m_vap") or 0.0) / mt_d if mt_d > 0 else 0.0
    gen = (pr_dn.get("m_vap") or 0.0) - (pr_in.get("m_vap") or 0.0)
    flash = gen > 1e-3 * max(mt_i, 1.0)
    if flash and wv_i <= 1e-9:
        label = "YES — liquid flashes to vapour"
    elif flash:
        label = "YES — additional vapour forms"
    elif gen < -1e-3 * max(mt_i, 1.0):
        label = "No — vapour condenses"
    else:
        label = "No"
    return dict(known=True, wv_in=wv_i, wv_dn=wv_d, gen=gen, flash=flash, label=label)


def _crit_ratio(k: float) -> float:
    return (2.0 / (k + 1.0)) ** (k / (k - 1.0)) if k and k > 1.0 else 0.5283


# ── control-valve sizing (ISA / IEC 60534-2-1 screening) ──────────────────
def _cv_size(st, info: dict, data: dict) -> dict:
    pr = st.props or {}
    p1, p2 = st.p_in_psia, st.p_out_psia
    dP = max(p1 - p2, 0.0)
    FL = data.get("fl") or 0.9
    xT = data.get("xt") or 0.7
    mv, ml = pr.get("m_vap") or 0.0, pr.get("m_liq") or 0.0
    rg, rl = pr.get("rho_vap"), pr.get("rho_liq")
    k = pr.get("cpcv_vap") or 1.3
    out = dict(FL=FL, xT=xT, FL_default=not data.get("fl"), xT_default=not data.get("xt"),
               dP=dP, flags=[], cv_l=0.0, cv_g=0.0, choked_l=False, choked_g=False)
    if mv > 0 and rg and p1 > 0:
        Fk = k / 1.4
        x = dP / p1
        xmax = Fk * xT
        xe = min(x, xmax)
        out["choked_g"] = x >= xmax
        Y = 1.0 - xe / (3.0 * Fk * xT)
        out.update(x=x, xmax=xmax, Y=Y, Fk=Fk)
        if xe > 0:
            out["cv_g"] = mv / (63.3 * Y * math.sqrt(xe * p1 * rg))
    if ml > 0 and rl:
        q_gpm = (ml / rl) * _GPM_PER_FT3H
        sg = rl / 62.428
        dPe = dP
        if mv <= 0:                                   # liquid only: choked-flow limit
            pv = info.get("pv_psia")
            pc = pr.get("pc") or 3208.0
            if pv is not None:
                FF = 0.96 - 0.28 * math.sqrt(max(pv, 0.0) / pc) if pc > 0 else 0.96
                dPmax = FL * FL * (p1 - FF * pv)
                out.update(Pv=pv, FF=FF, dP_max_liq=dPmax)
                if dP > dPmax > 0:
                    dPe = dPmax
                    out["choked_l"] = True
        if dPe > 0:
            out["cv_l"] = q_gpm * math.sqrt(sg / dPe)
        out["sg"] = sg
        out["q_gpm"] = q_gpm
    out["cv"] = out["cv_l"] + out["cv_g"]
    out["kv"] = out["cv"] * 0.865
    if mv > 0 and ml > 0:
        out["flags"].append("two-phase inlet: Cv = vapour part + liquid part (additive screening) — vendor to confirm")
    rated = data.get("cv_rated")
    out["rated"] = rated
    out.update(_cv_travel(out["cv"], data))
    out["pct_open"] = out.get("travel")
    return out


def _parse_cv_curve(txt_):
    """'10:4.2; 25:11; 100:85' → sorted [(travel %, Cv)] (needs ≥ 2 numeric points)."""
    pts = []
    for tok in re.split(r"[;\n]+|,(?=\s*\d+\s*[:=])", str(txt_ or "")):
        m = re.match(r"\s*([-+]?\d+(?:\.\d+)?)\s*[:=]\s*([-+]?\d+(?:\.\d+)?)\s*$", tok)
        if m:
            pts.append((float(m.group(1)), float(m.group(2))))
    pts = sorted(set(pts))
    return pts if len(pts) >= 2 else []


def _cv_travel(cv_req: float, data: dict) -> dict:
    """Valve travel (% open) that passes the required Cv.

    Measured points (``CV curve points``) win; otherwise the inherent characteristic scaled to the
    rated Cv at 100 % travel — Equal % (Cv = Cvr·R^(x−1)), Linear (Cv = Cvr·x), Quick opening
    (Cv = Cvr·√x).  Travel > 100 % means the valve is undersized."""
    rated = data.get("cv_rated")
    out = dict(travel=None, travel_basis=None, travel_status=None)
    pts = _parse_cv_curve(data.get("cv_curve"))
    if pts:
        if rated is None:
            rated = max(c for _, c in pts)
            out["rated"] = rated
        xs, cs = [p[0] for p in pts], [p[1] for p in pts]
        if cv_req <= cs[0]:
            x = xs[0] * cv_req / cs[0] if cs[0] > 0 else xs[0]
        elif cv_req >= cs[-1]:
            x = xs[-1] + (cv_req - cs[-1]) / max(cs[-1] - cs[-2], 1e-9) * (xs[-1] - xs[-2])
        else:
            x = xs[-1]
            for (x0, c0), (x1, c1) in zip(pts, pts[1:]):
                if c0 <= cv_req <= c1:
                    x = x0 + (cv_req - c0) / (c1 - c0) * (x1 - x0) if c1 > c0 else x0
                    break
        out["travel_basis"] = f"measured curve ({len(pts)} points)"
    elif rated:
        char = (data.get("cv_char") or "Equal %").strip().lower()
        r_ = data.get("cv_range") or 50.0
        ratio = cv_req / rated
        if ratio <= 0:
            x = 0.0
        elif char.startswith("lin"):
            x = ratio * 100.0
        elif char.startswith("quick"):
            x = ratio * ratio * 100.0
        else:
            char = "equal %"
            x = (1.0 + math.log(ratio) / math.log(r_)) * 100.0 if r_ > 1 else ratio * 100.0
        out["travel_basis"] = (f"{char.title()}" + (f", R = {r_:g}" if char == "equal %" else "")
                               + ("" if data.get("cv_char") else " (default characteristic)"))
    else:
        return out
    out["travel"] = x
    out["travel_status"] = ("UNDERSIZED — needs > 100 % travel" if x > 100.0 else
                            "Too open (> 90 %)" if x > 90.0 else
                            "Too closed (< 10 %)" if x < 10.0 else
                            "OK (10–90 %)")
    return out


# ── thermowell wake-frequency (simplified cantilever) ─────────────────────
def _tw_wake(st, info: dict, data: dict) -> dict:
    pr = st.props or {}
    fs = _flow_state(st)
    id_in = st.id_in or 0.0
    out = dict(flags=[])
    cls_d = _tw_class_data(st.spec)
    user = {k: v for k, v in (data or {}).items() if v is not None}
    data = {**{k: v for k, v in cls_d.items() if not k.startswith("_")}, **user}
    if cls_d.get("_src"):
        out["flags"].append(f"class defaults from {cls_d['_src']}")
    if "tw_mat" not in user and cls_d.get("_mat_src") and "tw_mat" in cls_d and not cls_d.get("_src"):
        out["flags"].append(f"material from {cls_d['_mat_src']}")
    mat = (data.get("tw_mat") or "SS316").upper()
    if mat not in _TW_MATERIALS:
        mat = "SS316"
        out["flags"].append("unknown material → SS316")
    E0, dE, rho_m = _TW_MATERIALS[mat]
    Tf = pr.get("T") if pr.get("T") is not None else 70.0
    E = max(E0 + dE * (Tf - 70.0), 1.0e7)
    root = data.get("tw_root") or 1.0
    tip = data.get("tw_tip") or 0.75
    bore = data.get("tw_bore") or 0.26
    U = data.get("tw_u")
    defaults = []
    if not data.get("tw_root"): defaults.append("root OD")
    if not data.get("tw_tip"): defaults.append("tip OD")
    if not data.get("tw_bore"): defaults.append("bore")
    if U is None:
        U = min(12.0, max(3.0, 0.5 * id_in))
        defaults.append("U")
    if not data.get("tw_mat"): defaults.append("material")
    d_avg = 0.5 * (root + tip)
    # fluid for the wake: phase with the larger volume share
    rg, rl = pr.get("rho_vap"), pr.get("rho_liq")
    if (fs["qv"] or 0) >= (fs["ql"] or 0) and rg:
        rho_f, mu = rg, pr.get("visc_vap"); phase = "vapour"
    elif rl:
        rho_f, mu = rl, pr.get("visc_liq"); phase = "liquid"
    else:
        rho_f, mu, phase = rg or 0.0, pr.get("visc_vap"), "vapour"
    if fs["qv"] > 0 and fs["ql"] > 0:
        out["flags"].append("two-phase at TW: wake check uses the larger-volume phase")
    V = fs["vm"]                                           # ft/s
    out.update(mat=mat, E=E, root=root, tip=tip, bore=bore, U=U, defaults=defaults,
               phase=phase, rho_f=rho_f, mu=mu, V=V)
    if not V or not rho_f or not mu:
        out["flags"].append("no velocity / fluid properties at this station")
        return out
    d_m = d_avg * IN_TO_M
    Re = rho_f * LBFT3_TO_KGM3 * (V * FT_TO_M) * d_m / (mu * 1e-3)
    if Re <= 2.0e5:
        S = 0.20
    elif Re >= 3.5e6:
        S = 0.27
    else:
        S = 0.20 + 0.07 * math.log10(Re / 2.0e5) / math.log10(3.5e6 / 2.0e5)
    fw = S * (V * FT_TO_M) / (tip * IN_TO_M)               # Hz (tip diameter)
    I = math.pi / 64.0 * (d_avg ** 4 - bore ** 4)          # in^4
    A_m = math.pi / 4.0 * (d_avg ** 2 - bore ** 2)         # in²
    # metal + displaced-fluid added mass (ρ_f lb/ft³ → lb/in³ = /1728), lb per inch of length
    mass_pl = rho_m * A_m + (rho_f / 1728.0) * math.pi / 4.0 * d_avg ** 2
    mass_pl_slug = mass_pl / 386.088                       # lb·s²/in²  (per inch length)
    wn = (1.875 ** 2) * math.sqrt(E * I / (mass_pl_slug * U ** 4))   # rad/s
    fn = wn / (2.0 * math.pi)
    ratio = fw / fn if fn > 0 else None
    v_lim = 0.8 * fn * (tip * IN_TO_M) / S / FT_TO_M       # ft/s for fs/fn = 0.8
    if ratio is None:
        verdict = "n/a"
    elif ratio < 0.4:
        verdict = "PASS (fs/fn < 0.4)"
    elif ratio < 0.8:
        verdict = "CHECK — in-line drag resonance band (0.4–0.8): ASME PTC 19.3 TW stress check"
    else:
        verdict = "FAIL (fs/fn ≥ 0.8) — shorten U / thicken stem / lower velocity"
    out.update(Re=Re, S=S, fw=fw, fn=fn, ratio=ratio, v_lim=v_lim, verdict=verdict,
               d_avg=d_avg, I=I)
    if defaults:
        out["flags"].append("DEFAULT used for: " + ", ".join(defaults)
                            + " — enter real data in Instrument_Data")
    return out


# ── sheet helpers ────────────────────────────────────────────────────────
def _fmt_or_dash(v, nd=3):
    if v is None:
        return "—"
    if isinstance(v, float):
        return round(v, nd)
    return v


def _datasheet(ws, title, cols, sections, label_w=34, col_w=17):
    """Vertical datasheet: sections = [(name, [(label, [v per col])...])…]."""
    n = len(cols)
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n + 2)
    _c(ws, 1, 1, title, bg=NAVY, fg=WHITE, sz=11, bold=True)
    ws.row_dimensions[1].height = 22
    r = 2
    _c(ws, r, 1, "Item", bg=NAVY, fg=WHITE, bold=True)
    _c(ws, r, 2, "Unit", bg=NAVY, fg=WHITE, bold=True, ha="center")
    for j, cname in enumerate(cols):
        _c(ws, r, 3 + j, cname, bg=NAVY, fg=WHITE, bold=True, ha="center", wrap=True)
    ws.row_dimensions[r].height = 30
    r += 1
    for sec, rows in sections:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n + 2)
        _c(ws, r, 1, sec, bg=STEEL, fg=WHITE, bold=True)
        r += 1
        for lab, unit, vals in rows:
            _c(ws, r, 1, lab, bg=LGRAY)
            _c(ws, r, 2, unit, bg=LGRAY, ha="center", fg="7F7F7F")
            for j, v in enumerate(vals):
                bg = None
                if isinstance(v, str) and (v.startswith("YES") or v.startswith("FAIL") or v.startswith("UNDERSIZED") or v.startswith("Too")
                                           or v.startswith("CHOKED")):
                    bg = RED
                _c(ws, r, 3 + j, _fmt_or_dash(v), ha="center", bg=bg,
                   wrap=isinstance(v, str) and len(v) > 24)
            r += 1
    cw(ws, 1, label_w)
    cw(ws, 2, 12)
    for j in range(n):
        cw(ws, 3 + j, col_w)
    ws.freeze_panes = "C3"
    return r


def _collect_instruments(runs):
    """[(run_label, station, line_no)] for every CV / instrument station."""
    out = []
    for label, sts, lns in runs:
        for i, s in enumerate(sts):
            if getattr(s, "instr", None):
                out.append((label, s, (lns[i] if lns and i < len(lns) else s.instr.get("line_no"))))
    return out


def _build_cv_datasheet(wb, items, idata, circuit_id, stream_name) -> None:
    """'CV Datasheet' — one column per control valve."""
    u = _u()
    ws = wb.create_sheet("CV Datasheet")
    ws.sheet_properties.tabColor = ORANGE
    cols = []
    rows: dict[str, list] = {}

    def add(key, v):
        rows.setdefault(key, []).append(v)
    for label, st, ln, fc, fm in items:
        info = st.instr
        pr = st.props or {}
        dn = info.get("dn") or {}
        data = idata.get((info["tag"] or "").upper(), {})
        sz = _cv_size(st, info, data)
        fs = _flow_state(st)
        fl = _flashing_info(pr, dn)
        cm = _pr_comp_pct(pr)
        cols.append(f"{info['tag']}\n{fc} ({fm * 100:g} %)")
        add("fcase", fc); add("fmult", fm * 100.0)
        mt = (pr.get("m_vap") or 0.0) + (pr.get("m_liq") or 0.0)
        # general
        add("tag", info["tag"]); add("desc", data.get("desc"))
        add("line", ln); add("run", label); add("seq", st.seq)
        add("stream", stream_name); add("ctype", {"F": "F — flow", "P": "P — pressure",
                                                   "T": "T — temperature", "L": "L — level"}.get(info["control_type"], info["control_type"]))
        add("bore", u.disp("Lin", info.get("bore_in"), 3) if info.get("bore_in") else None)
        add("lbore", u.disp("Lin", info.get("line_bore_in"), 3) if info.get("line_bore_in") else None)
        add("beta", (info["bore_in"] / info["line_bore_in"]) if (info.get("bore_in") and info.get("line_bore_in")) else None)
        add("basis", (info.get("note") or "")[:220])
        # operating
        add("p1", u.disp("P", st.p_in_psia)); add("p2", u.disp("P", st.p_out_psia))
        add("dp", u.disp("dP", sz["dP"], 3))
        add("dpr", (sz["dP"] / st.p_in_psia * 100.0) if st.p_in_psia else None)
        add("t1", u.disp("T", pr.get("T"), 2)); add("t2", u.disp("T", dn.get("T_out") or dn.get("T"), 2) if dn else None)
        add("ph1", pr.get("phase")); add("ph2", dn.get("phase") if dn else None)
        add("bv1", pr.get("beta")); add("wv1", fl.get("wv_in") if fl["known"] else ((pr.get("m_vap") or 0) / mt if mt else None))
        add("bv2", dn.get("beta") if dn else None); add("wv2", fl.get("wv_dn"))
        add("flash", fl["label"] if fl["known"] else "n/a (no composition / frozen split)")
        add("gen", u.disp("mflow", fl["gen"], 1) if fl["known"] else None)
        # flows
        add("mt", u.disp("mflow", mt, 1)); add("mv", u.disp("mflow", pr.get("m_vap"), 1))
        add("ml", u.disp("mflow", pr.get("m_liq"), 1))
        add("qv", u.disp("qvol", pr.get("q_vap"), 2) if pr.get("q_vap") else None)
        add("ql", u.disp("qvol", pr.get("q_liq"), 2) if pr.get("q_liq") else None)
        add("qgpm", sz.get("q_gpm"))
        add("qstd", (pr.get("q_std_vap_ph") or 0) * 24.0 / 1e6 if pr.get("q_std_vap_ph") else None)
        # inlet fluid
        add("rg", u.disp("rho", pr.get("rho_vap"), 4) if pr.get("m_vap") else None)
        add("rl", u.disp("rho", pr.get("rho_liq"), 3) if pr.get("m_liq") else None)
        add("mwv", (pr["m_vap"] / pr["n_vap"]) if (pr.get("n_vap") and pr.get("m_vap")) else None)
        add("mwl", (pr["m_liq"] / pr["n_liq"]) if (pr.get("n_liq") and pr.get("m_liq") and (pr.get("m_liq") / pr["n_liq"]) > 1.01) else None)
        add("mug", _ph(pr, "v", "visc_vap")); add("mul", _ph(pr, "l", "visc_liq"))
        add("z", _ph(pr, "v", "z_vap")); add("k", _ph(pr, "v", "cpcv_vap"))
        add("cpv", _ph(pr, "v", "cp_vap")); add("cpl", _ph(pr, "l", "cp_liq"))
        add("sig", _ph(pr, "l", "sigma"))
        add("pv", u.disp("P", info.get("pv_psia")) if info.get("pv_psia") is not None else None)
        add("pc", u.disp("P", pr.get("pc")) if pr.get("pc") else None)
        add("h2", cm.get("H2", {}).get("tot")); add("h2v", cm.get("H2", {}).get("vap"))
        add("h2s", cm.get("H2S", {}).get("tot")); add("h2sv", cm.get("H2S", {}).get("vap"))
        add("nh3", cm.get("NH3", {}).get("tot")); add("co2", cm.get("CO2", {}).get("tot"))
        add("h2o", cm.get("H2O", {}).get("tot"))
        h2v = (cm.get("H2", {}).get("vap"))
        add("pph2", u.disp("P", st.p_in_psia * h2v / 100.0, 2) if h2v is not None else None)
        h2sv = (cm.get("H2S", {}).get("vap"))
        add("pph2s", u.disp("P", st.p_in_psia * h2sv / 100.0, 3) if h2sv is not None else None)
        # sizing
        add("fl", sz["FL"]); add("xt", sz["xT"])
        add("cv", sz["cv"]); add("kv", sz["kv"]); add("cvg", sz["cv_g"] or None); add("cvl", sz["cv_l"] or None)
        add("chk", ("CHOKED (gas, x ≥ Fk·xT)" if sz["choked_g"] else "") +
                   ("CHOKED / flashing-cavitating (liquid ΔP > ΔPmax)" if sz["choked_l"] else "") or "No")
        add("dpmax", u.disp("dP", sz.get("dP_max_liq"), 3) if sz.get("dP_max_liq") is not None else None)
        add("rated", sz.get("rated")); add("open", sz.get("pct_open"))
        add("tbasis", sz.get("travel_basis")); add("tstat", sz.get("travel_status"))
        add("rec", sz["cv"] / 0.7 if sz["cv"] else None)
        # velocities / momentum
        vin = fs["vm"]
        add("vin", u.disp("v", vin, 2) if vin else None)
        dnst = None
        if dn and fs["area_ft2"] > 0 and (dn.get("q_vap") or dn.get("q_liq")):
            vdn = ((dn.get("q_vap") or 0) + (dn.get("q_liq") or 0)) / 3600.0 / fs["area_ft2"]
            add("vout", u.disp("v", vdn, 2))
            c_dn = _sound_speed_ms(dn)
            vg_dn = (dn.get("q_vap") or 0) / 3600.0 / fs["area_ft2"] * FT_TO_M
            add("mach2", (vg_dn / c_dn) if (c_dn and vg_dn) else None)
            rhov2_dn = (((dn.get("m_vap") or 0) + (dn.get("m_liq") or 0)) /
                        ((dn.get("q_vap") or 0) + (dn.get("q_liq") or 0))) * LBFT3_TO_KGM3 * (vdn * FT_TO_M) ** 2
            add("rv2o", rhov2_dn)
        else:
            add("vout", None); add("mach2", None); add("rv2o", None)
        c_in = _sound_speed_ms(pr)
        add("mach1", (fs["vsg"] * FT_TO_M / c_in) if (c_in and fs["vsg"]) else None)
        add("rv2i", fs["rhov2_m"])
        add("flags", "; ".join(sz["flags"]) or "—")
    L = lambda k: rows[k]
    pu = u.label("P"); dpu = u.label("dP"); tu = u.label("T"); mu_ = u.label("mflow")
    ru = u.label("rho"); vu = u.label("v"); qu = u.label("qvol"); lu = u.label("Lin")
    sections = [
        ("1  GENERAL", [
            ("Tag", "", L("tag")), ("Flow case", "", L("fcase")),
            ("Flow multiplier", "% of normal", L("fmult")), ("Service / description", "", L("desc")),
            ("Line No", "", L("line")), ("Run", "", L("run")), ("Seq", "", L("seq")),
            ("Stream", "", L("stream")), ("Control type", "", L("ctype")),
            ("Valve port bore", lu, L("bore")), ("Upstream line bore (NPS)", lu, L("lbore")),
            ("β = valve bore / line bore", "-", L("beta")),
            ("ΔP basis", "", L("basis"))]),
        ("2  OPERATING CONDITIONS", [
            ("Inlet pressure P1", pu, L("p1")), ("Outlet pressure P2", pu, L("p2")),
            ("Pressure drop ΔP", dpu, L("dp")), ("ΔP / P1", "%", L("dpr")),
            ("Inlet temperature", tu, L("t1")), ("Outlet temperature (flash-consistent)", tu, L("t2")),
            ("Inlet phase", "", L("ph1")), ("Outlet phase", "", L("ph2")),
            ("Inlet vapour fraction (mole β)", "-", L("bv1")), ("Inlet vapour fraction (mass)", "-", L("wv1")),
            ("Outlet vapour fraction (mole β)", "-", L("bv2")), ("Outlet vapour fraction (mass)", "-", L("wv2")),
            ("Flashing across the valve?", "", L("flash")),
            ("Vapour generated", mu_, L("gen"))]),
        ("3  FLOWS", [
            ("Total mass flow", mu_, L("mt")), ("Vapour mass flow", mu_, L("mv")),
            ("Liquid mass flow", mu_, L("ml")),
            ("Vapour actual volumetric flow", qu, L("qv")), ("Liquid actual volumetric flow", qu, L("ql")),
            ("Liquid flow (US gpm)", "gpm", L("qgpm")), ("Vapour standard flow", "MMSCFD", L("qstd"))]),
        ("4  FLUID PROPERTIES AT INLET", [
            ("Vapour density", ru, L("rg")), ("Liquid density", ru, L("rl")),
            ("Vapour MW", "g/mol", L("mwv")), ("Liquid MW", "g/mol", L("mwl")),
            ("Vapour viscosity", "cP", L("mug")), ("Liquid viscosity", "cP", L("mul")),
            ("Vapour Z", "-", L("z")), ("Vapour Cp/Cv (γ)", "-", L("k")),
            ("Vapour Cp", "BTU/lb-°F", L("cpv")), ("Liquid Cp", "BTU/lb-°F", L("cpl")),
            ("Liquid surface tension", "dyn/cm", L("sig")),
            ("Liquid vapour pressure Pv (at inlet T)", pu, L("pv")),
            ("Critical pressure Pc (mixture)", pu, L("pc"))]),
        ("5  COMPOSITION (HMB / Stream_Composition streams)", [
            ("H2 — total", "mol %", L("h2")), ("H2 — vapour phase", "mol %", L("h2v")),
            ("H2 partial pressure (vapour)", pu, L("pph2")),
            ("H2S — total", "mol %", L("h2s")), ("H2S — vapour phase", "mol %", L("h2sv")),
            ("H2S partial pressure (vapour)", pu, L("pph2s")),
            ("NH3 — total", "mol %", L("nh3")), ("CO2 — total", "mol %", L("co2")),
            ("H2O — total", "mol %", L("h2o"))]),
        ("6  SIZING (ISA / IEC 60534-2-1 screening)", [
            ("FL (liquid pressure-recovery factor)", "-", L("fl")),
            ("xT (gas pressure-drop ratio factor)", "-", L("xt")),
            ("Required Cv (full flow)", "-", L("cv")), ("Required Kv", "-", L("kv")),
            ("  Cv — vapour portion", "-", L("cvg")), ("  Cv — liquid portion", "-", L("cvl")),
            ("Choked flow?", "", L("chk")), ("ΔP max (liquid choked limit)", dpu, L("dpmax")),
            ("Suggested rated Cv (required ÷ 0.7)", "-", L("rec")),
            ("Rated Cv at 100 % travel (Instrument_Data)", "-", L("rated")),
            ("Cv curve basis", "", L("tbasis")),
            ("Valve travel at this flow", "% open", L("open")),
            ("Travel check", "", L("tstat"))]),
        ("7  VELOCITY / MOMENTUM (at valve port bore)", [
            ("Inlet mixture velocity", vu, L("vin")), ("Outlet mixture velocity", vu, L("vout")),
            ("Inlet vapour Mach", "-", L("mach1")), ("Outlet vapour Mach", "-", L("mach2")),
            ("Inlet ρv² (homogeneous)", "kg/m·s²", L("rv2i")),
            ("Outlet ρv² (homogeneous)", "kg/m·s²", L("rv2o"))]),
        ("8  NOT CALCULATED — VENDOR TO PROVIDE", [
            ("Noise (IEC 60534-8-3 / -8-4)", "dBA", ["vendor"] * len(cols)),
            ("Cavitation index σ, Fd, travel, trim, leakage class", "", ["vendor"] * len(cols)),
            ("Calculation notes", "", L("flags"))]),
    ]
    title = (f"CONTROL VALVE DATASHEET   |   Circuit {circuit_id}   |   Stream {stream_name}   |   "
             f"units: {u.system}   |   {datetime.now():%d-%b-%Y %H:%M}")
    r = _datasheet(ws, title, cols, sections)
    notes = ["Cv follows ISA / IEC 60534-2-1 with the standard-state Y, xT, FL, FF factors; liquid choked-flow limit uses the bubble-point pressure of the stream at the inlet temperature.",
             "Outlet state = flash of the same stream at P2 (isothermal, or adiabatic for composition streams) — 'flashing' means vapour mass increases across the valve.",
             "FL = 0.9 and xT = 0.7 are generic defaults unless a rating is typed in Instrument_Data; the final sizing, noise and cavitation checks belong to the valve vendor."]
    r += 1
    for n_ in notes:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(cols) + 2)
        _c(ws, r, 1, n_, fg=DGRAY, sz=8, wrap=True)
        ws.row_dimensions[r].height = 24
        r += 1


_INSTR_COLS = [
    # (header, unit-qty/label, group)
    ("Tag", "", "ID"), ("Type", "", "ID"), ("Line No", "", "ID"), ("Run", "", "ID"),
    ("Seq", "", "ID"), ("Comp ID", "", "ID"), ("NPS", "", "ID"), ("ID", "Lin", "ID"),
    ("P inlet", "P", "INLET CONDITIONS"), ("T inlet", "T", "INLET CONDITIONS"),
    ("Phase", "", "INLET CONDITIONS"), ("Vapour frac (mole β)", "", "INLET CONDITIONS"),
    ("Vapour frac (mass)", "", "INLET CONDITIONS"),
    ("Mass flow total", "mflow", "FLOWS"), ("Mass flow vapour", "mflow", "FLOWS"),
    ("Mass flow liquid", "mflow", "FLOWS"), ("Vol flow vapour", "qvol", "FLOWS"),
    ("Vol flow liquid", "qvol", "FLOWS"), ("Velocity (mixture)", "v", "FLOWS"),
    ("ρv² (homog.)", "kg/m·s²", "FLOWS"), ("Vapour Mach", "", "FLOWS"),
    ("Vapour density", "rho", "VAPOUR PROPERTIES"), ("Vapour MW", "g/mol", "VAPOUR PROPERTIES"),
    ("Vapour visc", "cP", "VAPOUR PROPERTIES"), ("Vapour Z", "", "VAPOUR PROPERTIES"),
    ("Cp/Cv", "", "VAPOUR PROPERTIES"),
    ("Liquid density", "rho", "LIQUID PROPERTIES"), ("Liquid MW", "g/mol", "LIQUID PROPERTIES"),
    ("Liquid visc", "cP", "LIQUID PROPERTIES"), ("Surface tension", "dyn/cm", "LIQUID PROPERTIES"),
    ("H2 total", "mol %", "COMPOSITION"), ("H2 vapour", "mol %", "COMPOSITION"),
    ("H2 partial P", "P", "COMPOSITION"), ("H2S total", "mol %", "COMPOSITION"),
    ("H2S vapour", "mol %", "COMPOSITION"), ("H2S partial P", "P", "COMPOSITION"),
    ("NH3 total", "mol %", "COMPOSITION"), ("CO2 total", "mol %", "COMPOSITION"),
    ("H2O total", "mol %", "COMPOSITION"),
    ("ΔP (instrument)", "dP", "PRESSURE DROP & DOWNSTREAM"), ("ΔP / P1", "%", "PRESSURE DROP & DOWNSTREAM"),
    ("P outlet", "P", "PRESSURE DROP & DOWNSTREAM"), ("T outlet", "T", "PRESSURE DROP & DOWNSTREAM"),
    ("Outlet phase", "", "PRESSURE DROP & DOWNSTREAM"),
    ("Outlet vapour frac (mass)", "", "PRESSURE DROP & DOWNSTREAM"),
    ("Flashing downstream?", "", "PRESSURE DROP & DOWNSTREAM"),
    ("Vapour generated", "mflow", "PRESSURE DROP & DOWNSTREAM"),
    ("Orifice d", "Lin", "ORIFICE / RO CHECKS"), ("β = d/D", "", "ORIFICE / RO CHECKS"),
    ("Critical P2/P1 (gas)", "", "ORIFICE / RO CHECKS"), ("P2/P1", "", "ORIFICE / RO CHECKS"),
    ("Choked?", "", "ORIFICE / RO CHECKS"),
    ("Notes", "", "NOTES"),
]


def _build_instruments_sheet(wb, items, idata, circuit_id, stream_name) -> None:
    """'Inline Instruments' — one row per PT / TT / FT / FO / RO, plus the thermowell wake block."""
    u = _u()
    ws = wb.create_sheet("Inline Instruments")
    ws.sheet_properties.tabColor = PURPLE
    ws.sheet_view.showGridLines = False
    ncol = len(_INSTR_COLS)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    _c(ws, 1, 1, f"INLINE INSTRUMENTS (PT · TT · FT inline / orifice · RO)   |   Circuit {circuit_id}   |   "
                 f"Stream {stream_name}   |   units: {u.system}   |   {datetime.now():%d-%b-%Y %H:%M}" + _case_tag(),
       bg=NAVY, fg=WHITE, sz=11, bold=True)
    ws.row_dimensions[1].height = 22
    # group header row 2, header row 3
    groups: list = []
    for ci, (h, q, g) in enumerate(_INSTR_COLS, 1):
        if groups and groups[-1][0] == g:
            groups[-1][2] = ci
        else:
            groups.append([g, ci, ci])
    for g, c1, c2 in groups:
        if c2 > c1:
            ws.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
        _c(ws, 2, c1, g, bg=STEEL, fg=WHITE, bold=True, ha="center")
    qlab = {"P": u.label("P"), "T": u.label("T"), "mflow": u.label("mflow"), "qvol": u.label("qvol"),
            "v": u.label("v"), "rho": u.label("rho"), "Lin": u.label("Lin"), "dP": u.label("dP")}
    for ci, (h, q, g) in enumerate(_INSTR_COLS, 1):
        unit = qlab.get(q, q)
        _c(ws, 3, ci, f"{h} ({unit})" if unit else h, bg=NAVY, fg=WHITE, sz=9,
           bold=True, ha="center", wrap=True)
    ws.row_dimensions[3].height = 40
    r = 4
    tw_items = []
    for label, st, ln in [x for x in items if x[1].instr["kind"] != "CV"]:
        info = st.instr; pr = st.props or {}
        dn = info.get("dn") or {}
        data = idata.get((info["tag"] or "").upper(), {})
        fs = _flow_state(st)
        fl = _flashing_info(pr, dn)
        cm = _pr_comp_pct(pr)
        mt = fs["mt"]
        P1 = st.p_in_psia
        dP = info.get("dp_psi") or 0.0
        k = pr.get("cpcv_vap") or 1.3
        d_or = data.get("d_or")
        D_in = st.id_in or info.get("line_bore_in")
        crit = _crit_ratio(k) if (pr.get("m_vap") or 0) > 0 else None
        ratio21 = ((P1 - dP) / P1) if P1 else None
        choked = None
        if info["kind"] in ("FO", "RO") and crit is not None and ratio21 is not None:
            choked = ("CHOKED (P2/P1 < critical)" if ratio21 < crit else "No")
        mwl = (pr["m_liq"] / pr["n_liq"]) if (pr.get("n_liq") and pr.get("m_liq")) else None
        vals = [info["tag"], info["kind"] + {"FT": " (inline)", "FO": " (orifice FT)"}.get(info["kind"], ""),
                ln, label, st.seq, st.comp_id, st.nps, u.disp("Lin", st.id_in, 3) if st.id_in else None,
                u.disp("P", P1), u.disp("T", pr.get("T"), 2), pr.get("phase"),
                pr.get("beta"), ((pr.get("m_vap") or 0) / mt) if mt else None,
                u.disp("mflow", mt, 1), u.disp("mflow", pr.get("m_vap"), 1), u.disp("mflow", pr.get("m_liq"), 1),
                u.disp("qvol", pr.get("q_vap"), 2) if pr.get("q_vap") else None,
                u.disp("qvol", pr.get("q_liq"), 2) if pr.get("q_liq") else None,
                u.disp("v", fs["vm"], 2) if fs["vm"] else None, fs["rhov2_m"],
                ((fs["vsg"] * FT_TO_M / _sound_speed_ms(pr)) if (fs["vsg"] and _sound_speed_ms(pr)) else None),
                (u.disp("rho", pr.get("rho_vap"), 4) if pr.get("m_vap") else None),
                ((pr["m_vap"] / pr["n_vap"]) if (pr.get("n_vap") and pr.get("m_vap")) else None),
                _ph(pr, "v", "visc_vap"), _ph(pr, "v", "z_vap"), _ph(pr, "v", "cpcv_vap"),
                u.disp("rho", pr.get("rho_liq"), 3) if pr.get("m_liq") else None,
                (mwl if (mwl and mwl > 1.01) else None),
                _ph(pr, "l", "visc_liq"), _ph(pr, "l", "sigma"),
                cm.get("H2", {}).get("tot"), cm.get("H2", {}).get("vap"),
                (u.disp("P", P1 * cm["H2"]["vap"] / 100.0, 2) if cm.get("H2", {}).get("vap") is not None else None),
                cm.get("H2S", {}).get("tot"), cm.get("H2S", {}).get("vap"),
                (u.disp("P", P1 * cm["H2S"]["vap"] / 100.0, 3) if cm.get("H2S", {}).get("vap") is not None else None),
                cm.get("NH3", {}).get("tot"), cm.get("CO2", {}).get("tot"), cm.get("H2O", {}).get("tot"),
                u.disp("dP", dP, 3), (dP / P1 * 100.0) if P1 else None, u.disp("P", st.p_out_psia),
                u.disp("T", dn.get("T_out") or dn.get("T"), 2) if dn else None,
                dn.get("phase") if dn else None, fl.get("wv_dn") if fl["known"] else None,
                fl["label"] if fl["known"] else "n/a (no composition / frozen split)",
                u.disp("mflow", fl["gen"], 1) if fl["known"] else None,
                u.disp("Lin", d_or, 3) if d_or else None,
                (d_or / D_in) if (d_or and D_in) else None, crit, ratio21, choked,
                ("; ".join(x for x in [data.get("desc"), info.get("note")] if x) or None)]
        vals[-1] = (vals[-1] or "")[:160]
        for ci, v in enumerate(vals, 1):
            bg = RED if (isinstance(v, str) and (v.startswith("YES") or v.startswith("CHOKED"))) else None
            _c(ws, r, ci, _fmt_or_dash(v, 4), ha="center", bg=bg, wrap=(ci == ncol))
        if info["kind"] == "TT":
            tw_items.append((r, label, st, ln, info, data))
        r += 1
    for ci in range(1, ncol + 1):
        cw(ws, ci, 13)
    cw(ws, 1, 12); cw(ws, 2, 12); cw(ws, ncol, 50)
    ws.freeze_panes = "C4"

    # ── thermowell wake-frequency block ──────────────────────────────────
    if tw_items:
        r += 2
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
        _c(ws, r, 1, "THERMOWELL WAKE-FREQUENCY CHECK (TT)  —  simplified Euler-Bernoulli cantilever, "
                     "average stem diameter, fluid added mass  |  final check: ASME PTC 19.3 TW-2016 by the vendor",
           bg=NAVY, fg=WHITE, bold=True, sz=10)
        r += 1
        th = ["Tag", "Line No", "Fluid used", "Velocity (m/s)", "Fluid ρ (kg/m³)", "Fluid µ (cP)",
              "Material", "E at T (psi)", "Root OD (in)", "Tip OD (in)", "Bore (in)", "Insertion U (in)",
              "Reynolds", "Strouhal S", "Wake fs (Hz)", "Natural fn (Hz)", "fs / fn",
              "Max V for fs/fn = 0.8 (m/s)", "Verdict", "Notes"]
        for ci, h in enumerate(th, 1):
            _c(ws, r, ci, h, bg=STEEL, fg=WHITE, bold=True, ha="center", wrap=True)
        ws.row_dimensions[r].height = 36
        r += 1
        for rr, label, st, ln, info, data in tw_items:
            w = _tw_wake(st, info, data)
            vals = [info["tag"], ln, w.get("phase"),
                    (w["V"] * FT_TO_M) if w.get("V") else None,
                    (w["rho_f"] * LBFT3_TO_KGM3) if w.get("rho_f") else None, w.get("mu"),
                    w.get("mat"), w.get("E"), w.get("root"), w.get("tip"), w.get("bore"), w.get("U"),
                    w.get("Re"), w.get("S"), w.get("fw"), w.get("fn"), w.get("ratio"),
                    (w["v_lim"] * FT_TO_M) if w.get("v_lim") else None, w.get("verdict"),
                    "; ".join(w["flags"]) or None]
            for ci, v in enumerate(vals, 1):
                if isinstance(v, float):
                    v = round(v, 4 if abs(v) < 1000 else 0)
                bg = RED if (isinstance(v, str) and v.startswith("FAIL")) else (
                    AMBER if (isinstance(v, str) and v.startswith("CHECK")) else None)
                _c(ws, r, ci, v if v is not None else "—", ha="center", bg=bg,
                   wrap=(ci in (19, 20)))
            ws.row_dimensions[r].height = 36
            r += 1
    # ── notes ────────────────────────────────────────────────────────────
    r += 1
    for n_ in [
        "Instr Type drop-down on Pipeline_Input: PT, TT, FT (inline meter, rated ΔP in 'Instr dP'), FO (orifice flow element), RO (restriction orifice).  Control valves appear on 'CV Datasheet'.",
        "Inlet properties are the flashed station properties; 'Flashing downstream' re-flashes the same stream at the outlet pressure (isothermal, or adiabatic for composition streams) and reports vapour created.",
        "H2 / H2S / NH3 / CO2 / H2O come from the stream composition (HMB or Stream_Composition); blank for manual streams.  Partial pressure = vapour mole fraction × P.",
        "RO / FO: 'Choked?' compares P2/P1 with the critical ratio (2/(γ+1))^(γ/(γ−1)).  Orifice d is read from Instrument_Data."]:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
        _c(ws, r, 1, n_, fg=DGRAY, sz=8, wrap=True)
        ws.row_dimensions[r].height = 24
        r += 1


# ── Line Data ─────────────────────────────────────────────────────────────
def _build_line_data_sheet(wb, runs, sp, gdf, circuit_id, stream_name) -> None:
    """'Line Data' — per Line No: conditions at the highest-pressure straight pipe,
    vapour / liquid properties, all IDs, flows, ρv², two-phase slug data."""
    u = _u()
    ws = wb.create_sheet("Line Data")
    ws.sheet_properties.tabColor = GRNHDR
    ws.sheet_view.showGridLines = False
    cols = [
        ("Line No", "", "LINE"), ("Run", "", "LINE"), ("Seq of max-P pipe", "", "LINE"),
        ("Comp ID", "", "LINE"), ("Basis", "", "LINE"),
        ("Max operating P", "P", "OPERATING POINT (highest pressure in line)"),
        ("T", "T", "OPERATING POINT (highest pressure in line)"),
        ("Phase", "", "OPERATING POINT (highest pressure in line)"),
        ("Vapour frac (mole β)", "", "OPERATING POINT (highest pressure in line)"),
        ("Vapour frac (mass)", "", "OPERATING POINT (highest pressure in line)"),
        ("GVF (vol)", "", "OPERATING POINT (highest pressure in line)"),
        ("Line NPS", "", "LINE SIZE"), ("Sched", "", "LINE SIZE"), ("ID at max-P pipe", "Lin", "LINE SIZE"),
        ("All IDs in this line (NPS / sch = ID)", "", "LINE SIZE"), ("No. of sizes", "", "LINE SIZE"),
        ("Mass flow", "mflow", "VAPOUR"), ("Vol flow (actual)", "qvol", "VAPOUR"),
        ("Density", "rho", "VAPOUR"), ("Viscosity", "cP", "VAPOUR"), ("MW", "g/mol", "VAPOUR"),
        ("Z", "", "VAPOUR"), ("Cp/Cv", "", "VAPOUR"), ("Superficial vel", "v", "VAPOUR"),
        ("ρv²", "kg/m·s²", "VAPOUR"), ("Mach", "", "VAPOUR"),
        ("Mass flow", "mflow", "LIQUID"), ("Vol flow (actual)", "qvol", "LIQUID"),
        ("Density", "rho", "LIQUID"), ("Viscosity", "cP", "LIQUID"), ("MW", "g/mol", "LIQUID"),
        ("Surface tension", "dyn/cm", "LIQUID"), ("Superficial vel", "v", "LIQUID"),
        ("ρv²", "kg/m·s²", "LIQUID"),
        ("Total mass flow", "mflow", "MIXTURE"), ("Homog. density", "rho", "MIXTURE"),
        ("Mixture vel", "v", "MIXTURE"), ("ρv² (homog.)", "kg/m·s²", "MIXTURE"),
        ("Regime (Taitel-Dukler)", "", "TWO-PHASE / SLUG"), ("Regime (Shell/BB)", "", "TWO-PHASE / SLUG"),
        ("Slug freq", "/min", "TWO-PHASE / SLUG"), ("Slug holdup", "", "TWO-PHASE / SLUG"),
        ("Slug force @ 90° bend", "N", "TWO-PHASE / SLUG"), ("Slug force @ 90° bend", "lbf", "TWO-PHASE / SLUG"),
        ("Straight pipe length", "L", "LINE TOTALS"), ("P in (line)", "P", "LINE TOTALS"),
        ("P out (line)", "P", "LINE TOTALS"), ("ΔP (line)", "dP", "LINE TOTALS"),
    ]
    ncol = len(cols)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    _c(ws, 1, 1, f"LINE DATA — straight pipe at the highest operating pressure per Line No   |   Circuit {circuit_id}   |   "
                 f"Stream {stream_name}   |   units: {u.system}   |   {datetime.now():%d-%b-%Y %H:%M}" + _case_tag(),
       bg=NAVY, fg=WHITE, sz=11, bold=True)
    ws.row_dimensions[1].height = 22
    groups: list = []
    for ci, (h, q, g) in enumerate(cols, 1):
        if groups and groups[-1][0] == g:
            groups[-1][2] = ci
        else:
            groups.append([g, ci, ci])
    gcol = {"LINE": NAVY, "VAPOUR": STEEL, "LIQUID": "1F7070", "MIXTURE": PURPLE,
            "TWO-PHASE / SLUG": ORANGE, "LINE TOTALS": GRNHDR}
    for g, c1, c2 in groups:
        if c2 > c1:
            ws.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
        _c(ws, 2, c1, g, bg=gcol.get(g, STEEL), fg=WHITE, bold=True, ha="center")
    qlab = {q_: u.label(q_) for q_ in ("P", "T", "mflow", "qvol", "v", "rho", "Lin", "dP", "L")}
    for ci, (h, q, g) in enumerate(cols, 1):
        unit = qlab.get(q, q)
        _c(ws, 3, ci, f"{h} ({unit})" if unit else h, bg=gcol.get(g, STEEL), fg=WHITE,
           bold=True, ha="center", wrap=True)
    ws.row_dimensions[3].height = 44
    r = 4
    sizes_rows = []
    for label, sts, lns in runs:
        order = list(dict.fromkeys(lns))
        for ln in order:
            idx = [i for i, l_ in enumerate(lns) if l_ == ln]
            lst = [sts[i] for i in idx]
            pipes = [s for s in lst if (s.fitting or "").lower().strip() == "straight pipeline"
                     and s.props and s.id_in]
            basis = "straight pipe"
            if not pipes:
                pipes = [s for s in lst if s.props and s.id_in]
                basis = "no 'Straight Pipeline' row — highest-P fitting used"
            if not pipes:
                continue
            st = max(pipes, key=lambda s: s.p_in_psia)
            pr = st.props
            fs = _flow_state(st)
            ph = _station_phase_state(st, sp, gdf)
            t = None
            try:
                t = station_tp_inputs(st, sp, gdf)
            except Exception:
                t = None
            mwv = (pr["m_vap"] / pr["n_vap"]) if (pr.get("n_vap") and pr.get("m_vap")) else None
            mwl = (pr["m_liq"] / pr["n_liq"]) if (pr.get("n_liq") and pr.get("m_liq")) else None
            if mwl is not None and mwl <= 1.01:
                mwl = None                      # fictitious liquid MW (manual stream)
            c_s = _sound_speed_ms(pr)
            mt = fs["mt"]
            # all IDs under the line
            sz_map: dict = {}
            for s in lst:
                if not s.id_in or (getattr(s, "instr", None) or {}).get("kind") == "CV":
                    continue                      # a valve port bore is not a line size
                kk = (s.nps, s.schedule, s.id_in)
                e = sz_map.setdefault(kk, dict(L=0.0, vmax=0.0, n=0, od=s.od_in, wall=s.wall_in))
                if (s.fitting or "").lower().strip() == "straight pipeline":
                    e["L"] += s.length_ft or 0.0
                e["n"] += 1
                e["vmax"] = max(e["vmax"], s.v_fts or 0.0)
            id_txt = "; ".join(f"{k[0]}\" {k[1] or ''} = {u.disp('Lin', k[2], 3)} {u.label('Lin')}"
                               for k in sz_map)
            for k, e in sz_map.items():
                sizes_rows.append((ln, label, k, e))
            # two-phase
            reg_td = reg_sh = freq = hl = force = None
            if t is not None:
                spl = _is_single_phase(t)
                if spl:
                    reg_td = reg_sh = spl
                else:
                    try:
                        reg_td = taitel_dukler(t)["regime"]
                        reg_sh = shell_beggs_brill(t)
                        freq = slug_freq_gregory_scott(t.vsl, t.vm, t.d) * 60.0
                        hl = slug_holdup_gregory(t.vm)
                        rho_s = t.rho_l * hl + t.rho_g * (1.0 - hl)
                        force = slug_force_bend(rho_s, t.vm, t.d)
                    except Exception:
                        pass
            len_pipe = sum(s.length_ft or 0.0 for s in lst
                           if (s.fitting or "").lower().strip() == "straight pipeline")
            vals = [ln, label, st.seq, st.comp_id, basis,
                    u.disp("P", st.p_in_psia), u.disp("T", pr.get("T"), 2), pr.get("phase"),
                    pr.get("beta"), ((pr.get("m_vap") or 0) / mt) if mt else None,
                    ((fs["qv"] / fs["q"]) if fs["q"] > 0 else None),
                    st.nps, st.schedule, u.disp("Lin", st.id_in, 3), id_txt, len(sz_map),
                    u.disp("mflow", pr.get("m_vap"), 1) if pr.get("m_vap") else None,
                    u.disp("qvol", pr.get("q_vap"), 2) if pr.get("q_vap") else None,
                    u.disp("rho", pr.get("rho_vap"), 4) if pr.get("rho_vap") else None,
                    _ph(pr, "v", "visc_vap"), mwv, _ph(pr, "v", "z_vap"), _ph(pr, "v", "cpcv_vap"),
                    u.disp("v", fs["vsg"], 2) if fs["vsg"] else None, fs["rhov2_g"],
                    ((fs["vsg"] * FT_TO_M / c_s) if (c_s and fs["vsg"]) else None),
                    u.disp("mflow", pr.get("m_liq"), 1) if pr.get("m_liq") else None,
                    u.disp("qvol", pr.get("q_liq"), 2) if pr.get("q_liq") else None,
                    u.disp("rho", pr.get("rho_liq"), 3) if pr.get("rho_liq") else None,
                    _ph(pr, "l", "visc_liq"), mwl, _ph(pr, "l", "sigma"),
                    u.disp("v", fs["vsl"], 3) if fs["vsl"] else None, fs["rhov2_l"],
                    u.disp("mflow", mt, 1), u.disp("rho", fs["rho_h"], 4) if fs["rho_h"] else None,
                    u.disp("v", fs["vm"], 2) if fs["vm"] else None, fs["rhov2_m"],
                    reg_td, reg_sh, freq, hl, force, (force * _LBF_PER_N) if force else None,
                    u.disp("L", len_pipe, 1), u.disp("P", lst[0].p_in_psia), u.disp("P", lst[-1].p_out_psia),
                    u.disp("dP", lst[0].p_in_psia - lst[-1].p_out_psia, 3)]
            for ci, v in enumerate(vals, 1):
                if isinstance(v, float):
                    v = round(v, 4 if abs(v) < 100 else 1)
                _c(ws, r, ci, v if v is not None else "—", ha="center",
                   wrap=(ci in (5, 15)), bg=(LGRAY if (r % 2) else None))
            ws.row_dimensions[r].height = 30
            r += 1
    for ci in range(1, ncol + 1):
        cw(ws, ci, 12)
    cw(ws, 5, 22); cw(ws, 15, 40)
    ws.freeze_panes = "B4"

    # ── all line sizes ───────────────────────────────────────────────────
    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    _c(ws, r, 1, "ALL SIZES UNDER EACH LINE NUMBER", bg=NAVY, fg=WHITE, bold=True)
    r += 1
    sh = ["Line No", "Run", "NPS", "Schedule", f"ID ({u.label('Lin')})", f"OD ({u.label('Lin')})",
          f"Wall ({u.label('Lin')})", f"Straight length ({u.label('L')})", f"Max velocity ({u.label('v')})"]
    for ci, h in enumerate(sh, 1):
        _c(ws, r, ci, h, bg=STEEL, fg=WHITE, bold=True, ha="center", wrap=True)
    r += 1
    for ln, label, k, e in sizes_rows:
        vals = [ln, label, k[0], k[1], u.disp("Lin", k[2], 3),
                u.disp("Lin", e["od"], 3) if e["od"] else None,
                u.disp("Lin", e["wall"], 3) if e["wall"] else None,
                u.disp("L", e["L"], 1), u.disp("v", e["vmax"], 2)]
        for ci, v in enumerate(vals, 1):
            _c(ws, r, ci, v if v is not None else "—", ha="center")
        r += 1
    r += 1
    for n_ in [
        "Operating point = the 'Straight Pipeline' row with the highest inlet pressure within each Line No (usually the first pipe).  Properties are the flashed station properties there.",
        "ρv² in kg/m·s² (Pa): vapour = ρg·vsg², liquid = ρl·vsl², mixture = ρ_homogeneous·v_mixture².  Slug force = √2·ρ_slug·vm²·A at a 90° bend (Gregory holdup), screening only.",
        "Two-phase regime / slug values are blank for single-phase lines; manual streams show '—' for MW of liquid when no Liq MW was typed."]:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
        _c(ws, r, 1, n_, fg=DGRAY, sz=8, wrap=True)
        ws.row_dimensions[r].height = 24
        r += 1



# ════════════════════════════════════════════════════════════════════════
#  OUTPUT THEME  — Technip Energies light theme + TEN logo
# ════════════════════════════════════════════════════════════════════════
# Every output workbook is re-coloured just before it is saved: dark banner / header fills become the
# light Technip Energies tones with dark-blue text, and the TEN logo is placed on every sheet.
_TEN_NAVY = "001742"
_LIGHT_FILL = {            # dark fill used by the writers → light theme fill
    "1F4973": "AED3FF", "2E75B6": "CDE3FF", "375623": "A0EEC0", "7030A0": "DCDCE1",
    "C55A11": "FFB4AA", "1F7070": "CDF5E1", "404040": "DCDCE1", "A0A0A0": "DCDCE1",
    "C00000": "FFB4AA", "808080": "DCDCE1", "595959": "DCDCE1", "505050": "DCDCE1",
}
_TEN_LOGO_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAPAAAABdCAIAAAAUtXz1AABBV0lEQVR42u29d7gkV3UtvtY+p6rTDZOTZpRmpFHOSEICoQACEyWC"
    "hEw02MaGR/LD+Bn7PdvgwANsbGOCBQaMyDkKkJCQkFBCAQWUNZJmNJrRxDtzU3dXnb3fH6eqb/cNM3dE+H0/uPXVJ121qiv1Pvus"
    "vfba+9DM8Ovf1KAGMyPpZZoDcoMZSJjBCzjdGYICgBBCkJjb5rapG3+tBm0GNQBwXUa8fVw37gqbdoetI7p+hw23MdLmrjENpgJV"
    "aD3FQIWNCpcOuv0G/YK623+eX9EvqUwYd7R+mTPrue03YNDRjrsNbmhcb1yf3fpY++aNYf3WsGmXbR6zkAvaBnEQBwYwwHIgwAIM"
    "gMITXhrVdFVfZfmC5PAlctwynnaQP3KJ71wr6Jxlz22/NoNWQzAkpSvdNhJ+9GD+vXvb1z/UfmLIRloGM/gEDuJMnCcNJCCgAgoo"
    "TAVmYuaBPGiA5rBcoA7q6LG4bocsDc88JD3/6PqRy7wXAsgDyJ55YG6bM+hfzisDQeEIEqOZ/fyx7Is3t75xd3vjTgMMDnDOiXMI"
    "hgAGVSWdmigBOBBduJigAQoNQjjNhRB6pcsDAgzBoHmlZqcdIK88rvHMQ2r7z/MAsgARuDlvPWfQv+QWwzUnGM/0279oXnJj83v3"
    "5mg7VB0rpAVTAwEVWKBkpJkqSRhAkqKEgqADCBjooIQFB6MFg4JUigkpJAgJmivaOYIetKhy4fG1i46vHrMsjTczFzXOGfSTxxiR"
    "eQDwmZubn/jp2DUPt2FwNU/HEJwpYSANzI0GMwgAwgxQMSOUBEg1BWiIICQFEpoBwWgGA1F6cYOCZoSJEMI8A9rZygV6/rGNt5w2"
    "sGahL8x6OrZkbpsz6Bm3LCBxAHDtuuzd3x/78X15DnM10FlQMzMaANKc0axjkSyhrgVAEU2TZqYAFAYRWAI4QIlAKqAKgHEkUNSJ"
    "iTPmYoHmHEQsyxUZlg/iT59af9vT+vorMhcvzhn0PvAYQeEdNg3rv/xo7IPXhJAbq7mjBtBMACMCoYDRTAkTD3qwCw2YRT9dgnDr"
    "QGhACuwBJYpjLH7PnKhz5miSiyoDjUaI90LLQ47x1vEr+ffPm//ctTUAucLPBYtzBr1nmBFd7aX3ZO/6xvjt63MZcJBMQ5ugMQEq"
    "AIEczIg2kQMEvSE1CJCXkV+Xt+aEPYMAA2AwAA4mAAv0wMhCg0YxFo7fxAAwB0ABPXW0JdZ829MG/vo58+fXOGfTcwa9p/jPCYLZ"
    "ey4dfd8PMM4kqbYDm4oMKmACCCgwAwKgjGQcAOcMDjBaTpgVZioGgiBpphNGLQoARpoQPjpw0Ixdx5jBWIwEdp6HZua9QTUfzk5d"
    "nfz7S+Y9ZVUl8uJz6GPOoHu2dkDqsHk4f+OXdn7j5uBqfUySPGRwI2AGVGEpYGAosURhRQbCASRMxXKWAMKiQSNGg1ZgDMIiQW2k"
    "UUwAGNQIFQVtwjCLe1dAYNEJC4KRFE9hyEZbgw3+x/kLX3lCTRWYg9RzBt3Z4sR99xPh1Z/efsuGVtLvcq1YqIACjoNNIjHzAGAl"
    "LJ5gllnQFAANJZyIVs3iyMKgQbKEy5jkVa3bmnv+x8QTFZclzCzxlrWVbfv758/7i7PrphCZs+k5gy6Rxg0Pt17xqfa6ocw1miEo"
    "rAJUoAKOU5qAB1IAE+dkr01Pg5c7h1kXhv4VWFw5IuCoprmOJm8/u+8D59XNQMzZ9G/tNqtYKQtwgmsfar744yPrdqrvYwikgkoq"
    "IrkmFrpMd8/2YtOQJr8miZRpUJCS1uyDl4294Uu7HCGRBJ/bfjcNOpLNP30ku+DikU2j4huWtzMGderEIgUBgxhdNxogSRJ7D8MM"
    "1F5k8itynmZmFhM2lmcQpP3uEz/R//HV3WoW9Xpz2++cQashcfjFpuzVH29uGqv5CvKQO80SE1hicIAAAjhFxSCYlae13v3Xj6tA"
    "Ncuc9/PqH/5x+Mvv7gYtD3M2/Ttm0JHq2rw7XPTJ4XXb6RrI85BYYK5mDBAlLXJxkgAV0E92yZzWJXfv3QezB4T8ku6585dQKbnB"
    "AM1bfrDy/svCh69tJr7I289tvxMGHbMnQe1Pvjh85wa6QYR201kQNTqfwZkjGMAsZjRiUvrJxm9PHmZwFhDFJIE4aEuZKTJXTf/y"
    "m+Pfv2/cCfI5m/4dMehcQeJvLh391i2WDkjIW0QAGMxlTOAcGCAZ2Ia2YRktcFq3GmFxj9lN4A0y6pZsMnSeNZK2GS7KntFCQGkZ"
    "0NaQm+hY0D/6/NC6HbnnnJ/+HTDooEgdvnd3+19/GFzdZxpAgAK4gARIgJidzmiZWDuxpkebCLMD0L+JzaYafpRKWab5eFrVjdv5"
    "pi+PjOemOgemf6sNWg1CbB21P/tqc0yFiZkqCIMovZEQg1gUxrmARI1WaDpnHRTqZLTAX5FNRcFTl4XSwGCMWUujGcmQt9XV/A/u"
    "zD98TdM75DpnCb/VHprEuy8dvX9DcLUsz1oQg4IqiEk+5GAGA82JJYZKYJKL0140TM6onig4vV+tt7YZqTgxAAKLKF/MnBoNwkry"
    "T5c2b9+Ye1fUKMxtv20GHQxCXPVA+6M/HnEDUFMgQ1SBwgRKBCAvXKw6aDVYNUgFkuwttrMS1fI3iT0AWLTmQlItZgkl0Vx9kuzY"
    "of/4naEo85jDHb9tBh3FFM3M/s93x4LBmJkZxEMJoyHK4SL94WBiYC5QYdTpo8hmFAGfQY3RKxpjCAhlHBkFaddLRLMnXgRLHRKn"
    "I0K6vzzVN8erGePtxmw3LO4s8Q3zdpYMyJdvbV96V9M9qfRhB+DMcu+5wX05flpkGBRBseckUXGqvXmaPVwuat/3HDpPe5WZTvvk"
    "nneWm58UC3qHz9/cvuH+zNVVAwFfGr0VajYTaFJ+YsY2QJhO+N/4LIzDpcdcYw3hFMXRBC3R9UTKWG0F64w6s0lM3x4gezkCCgmU"
    "ESx01dHCo3LVNIiB6d9+v/nsI6qx+mYfScMnyzbiSXKVsUVEIbFib1wyXYXObO5wppuJL5ws6o6t6weezXuY6bR7vqVfUhTpu0e8"
    "E+wYCx+9spmZ88wtXjkaVo9yyHpsFxaNzyyGjpOoM5v0GJxeMWedY2all+LMxhxHXo8/tAmdKVBOGAZAc0jd3fpg9pVbxl9xSm1f"
    "SwGambVtdoIYgxDVhEKoYSzb+zOqoebZPczi7UXz2jik67ZrO8eiPq6cJwsbjF0cugvPgmI8K952PeFMVjLWLqToNT/R18pKKnXr"
    "iN7zhK4a5EGLXIcz6H0JyNQIpA6pZ+fmx7M4a6OaTJy2naMVjFNQpwFeUE0o5VM8uaYUPQbtBT+8u3n9w5lvpEFLYb7N5Jw4aZLp"
    "IS3Y40G7aWGb9pzsrsDidOCbk//gzFi8K1E4MfZsOriCQrb9b1eOvuzEinNitnevFn/U3S17xSdHf/pIXq9xzzGlE7RbduhCd8nr"
    "GwfNl0d26LM/PDLSUvGcaSpPBNuH7X3n1970tEq04/jPraP2lZtbX7g1+8UWHWsaFEjRX+GxS9wLj/YvPDY9cIF0DOK6R/KLPjli"
    "QgEu/x99hy1x3eYYn3TnmJ35H8Nbdlkz2Ode03ju4Umwoh3FlhF993ebX70j2z6mVc/zj0r+6fzafoPFK4plFUNNe/WnRm94NCdx"
    "wYnph15aj+Kf9Tv1mR8aGctUFJ94VeM5hydRTP+Pl41/6IpWfYAhoHuEmSF1WDUgzz7U/8HT0v0G5MnZ9IRBC9HM9NPXjQEOjpa5"
    "sq4vWoPONFuw47M5QxTInjEw/ZzVox/FFETMKV7Z9uilp8WJnamm85MSgBkldXeuy795e+uCE2vxvc9mC4ptw7Zzl+3OukiSrpKa"
    "4gUCThCaNj+1yA9miseGtNk0JN2QqmdLHdpDNty0zkTsBdc/kr/x82M/XxdQBTo5oRZbuV1xf3bFndl7L2u99ZzKq09Ol/QLgPEM"
    "G3cYHAC0ZkgS5IbHhmzHLkNmo80J8xrL7OWfHvvxrW0MiPMcU1xydeuhIf3+n/b3pRO+JTc8vsu27jAIto5MOKtc8diQttqGgNHW"
    "xOV2jtmOHbYjMgtBe2o1BI9uCdf+IvvM9e2LX1M/c41/EjYt3S7nlg3tH9+v0vB5HkOnKXB9b2CAxRYLTziRImSsTVHSON3W9d2y"
    "thtKs27dh5Qpx3LXHhg5EQgajFSis1uZlo9/l6cvLqJ0ZLPtv3xTlqvtQ3RoGM8MTQstQ9ybBoIpmEBSIDc0tXPAaMu0HPipBxO4"
    "BAiGZvndrr3dMjStHQdAgAiueCB/wYdGfr4hJAsIBTIsbHBZn6QeGFMY3CC3NPV/fXbsXd9tRlgihE/gEvhkRlRKoJKACZAgTvcx"
    "Sfz129s/vj3zCwVAGDbNkS6V6+7MP3VDq9Dflk7GueKRUz/h1kikvrh0t1F6AVIkHkkFSxZw0SAXD3LRIJcsYKVGKCoL+MDW8Puf"
    "HL1/qwr3mU7tCQo/d2Mra4mvqbaM0ZqKglaZbla3mVzy1GO5lxhuctBQgl6beRBxBrQ9BZl0QDMwMT107kUBQwBZsavvyW/fGE5c"
    "5XOF3yPqiMbRV+UHXlLbOmzOM8Zqice/X9W65t7MNRjG7cy1/lUnV/oqNMCCDVS4vJ8ARMqHyuxdv1c9aX8/nsP3vh0RtNt24oEe"
    "QMXjge36J5eMbR9HtZ/NYXvR8embzqisXiKOGBq3Wx4NX76tfcV9WZ4zGeBLjkkm3kXHHe2NCekpugd+8lCQBJphsMLXnpF+6458"
    "/Y7gGvzQ1c3XnZr2VRiBRyLoS2mxCM72dNquUAlZ005bnVzymnr379TK8d83tP7lilZlgJu26sXXtj5wfm1fydSi/lSIHWP2jV8Y"
    "qk7zELOBapGK6X7WWeiby3CQE1GuUTq0nE2bFCxYDrMJcFIMGCugQoEx9h79lsjaeugOAloOKy2IdUZez6ABScJtQ3bZXdmJq/ws"
    "310ieOZhyaQPL70ru0YRMhy+n/vOn/T1pZw2GihuVXHO2uTsQ/yegQ2Jj1/TenBTqC1ga8Te/aLq/35OrXPAAfNx7Ar32qemNzyS"
    "v+MrY6uXuecflUTA3f2is4AsFHmGjsEJ0Q7T0G0Ato2ZEmjbMfu7951XP3ZF+3WfHEkG+dAm/cbt2atPTuOpEsf+6r5z+IZagoMX"
    "THaU731h7fFhu+S6tlRw1f35tjFbVOdsopoeyBG9+uX3tDdv9/RULfppSKl3Zgk8J3ZOVAnGXQgREhCyM8cTSqqwY0cWuztPtxuh"
    "Qot/ELHpFySekJhKGYnAOROxntdpHZTeQ18XjHaH/VbQShLRygpd8vKf5yMt+Fn/QLE5ZdwzRTC0ckCANo5f4fpSxg/jPokwJsA9"
    "6hODIhicYNOwfeGWttQwPornHZf87+fUOgy0GlQRFDCcdqD/0dv6P3pB3aYb90v6mThUPVJX7BWPxGH5AN10wYh0TZh5wAuOTvZf"
    "7kILcPj0z9oae69Yhy7a5y0Y8u6nMDRzqOHM1V7MlNg5bq1sn089gXou/UVAbkIqaWZKEROoqZRcrmJKgaB1iUY7oaGVUVxsE4ZY"
    "8+oFADMtRELTcCY2FbnEkWWUXt0HENRCMyY2gVQcETRqj2KrDhTDsHjltByglREbQVowAFSakYBqYILr12HdFj1mlZhOtAzZK/bo"
    "3JV0CBoi11jR2HOM9hJVKnbXplBPkYUeIigY6gmP38/FN/nAFn1sp7mUktmbn161SS23OTEAan6ynDYOcDP8209aS/tEbbLwsZVj"
    "pGUi0JkRYDvYogYvOiH9v99t+gHesi7/2fpwygE+C09e+CuYzJBWPAjcuyVEQ6sn5ve98aZXgyd2jYWfPZLDAxrtAgaoTfxAjQpS"
    "LzZBl7OTLCFifxcFGGsMASEBKiEgCMvNRtt0gsGaTDuiWYRiBWg2EFIU8Aow0rJ2YCS8Vanj2t8vh6zwC2vYNKIPPBFao0GqUnpA"
    "K3rnFcwGnGCgj7GxghlGmgi5IZKhMS6EQekTNHfy+vvzY1alv3zSintjYcxgnn/9rfGqo+pEQCJAyLD/Qv7oz/rn1wXAg1ty5qYJ"
    "l8yTE1Y5cnrm20kBW6dO0AZ84IfNGe8moXcz5gJZzuHPPir58FWttmJ43L52e3bKAd5+iUqMYBjLbCLdpiDxzTuyT1zbSurMdtvx"
    "q/zCBnV2nqXHoIW4eX14bEsuaVK4twL/OkBgecNw8YV9z16bjGfWM6o4lVyblHNBUFQ8f745P+9TQysG5fLXzat6xpzW3hMlRFA0"
    "Ur7rhyMfvWY8rTPPUAXecHb9jafVD57vnCAofv549r6rRr9887ivSOhl9pwgjNkzj0m/8Oq+scxIBMVbPjfyzRsz33B54ETG1ijB"
    "aO6Ku/CGc35TZeHEcNOGg01+FW1LKtL5OCvbnPRX2FchZs4yzog191Dvxj1HjMVEZIaTVrmTD/JX3pfR44q7s61nVxc3mD0JObnC"
    "N3jDhnD4PwxPIhA3D6l4IIev4A9OSz2RG/y+eWgFBLc9juEmknoWFKDAW+ksBIBQF/VxYSNqL5/MD7esT4TqRVb0u2Tf55HBmoAG"
    "Y8XsoxcMvOaEasfJOcGJK5MvvXLeIYvcP1w6KlWxSb9VQF/C+XUOGoUwoL8iCCia+dqEJDAoLeHNj2K8bbX0N2LRZocvcwtrzHsL"
    "hUNuK+ZJ4gpwtqhfom/ZMWY7x3TZgMBmTFbbdE0aSJywyg1WmHfjRoMQzYBbNoc9m6UIgqI/5XOOSa66N/NV3rkxXPtIfv5RyZMp"
    "jzCAaGW2fvvkL1MsH4Wjvf/C+llrEt33Hm4+GuhdGzPQaEoV+sQswLVMIeYNpmA7hxqaGbozsYUb7030a2+gkCsqHs2scP3tAC+I"
    "uagYJBn3xPsFReqRB4XT9ri985y+15xQbedwApGiTixe6z3P6b9jo37ntpZrSMiLgC8Cw5jOCAojckOeR+GglK0ki/43gTmS5uNb"
    "5IHH5ZgDvWrBy/66vDMhGf/hvNqZh/g89FZUAp6YVy0A3NplklREYTuG9OqH8guPTzPtyf50ZEBOJk91MSAX4GOvbBy+VLLQi/uJ"
    "kTZOfN/uTUO2ZzgcqcaXHJ3882WytWnawvfvzF50ZOKmJIz3BUf3flXQSHj6av/GsysvPCpR2zewURh04jDc1Ic2ZpAo4ilZ407W"
    "AXDCRgVC1NO9kGVRENIDzxwADFQLpUGkR6RL8rIXaswBQF9KtLHfQveqE2pRmhPPVlyLRcrpb5/TuOzuPM8jgp7gnuPzRIVD0fjR"
    "Oq3DyogBYnAUaw3jng045sBfu5o08pT9CedXmefTDB4t3fD+8+QpK+X6dTk9/+VHrecflTQStsNE/8piOQ7ioe2qikMWy1QmruLR"
    "SKh+cuo7y20PGKbbl2cBBy+UM9cmX7qx5Wr89h3tv3ludcWgcG9piClgH6Flx6/y//Li2qTk6LIBWTrARsJfSsuxZThs2Q2Ii2m8"
    "0KPkI4Cm2hdvbt+3Scdzc1KQGwJkua1Z4l5wZMVKDl+IzSP6/Xtbu5rmxECoMXHcsCtvBpRTaOxFh+vWt29a36qk0skG9Q5YxpRv"
    "1fPa9RmIAxe6lQPSCYmGW/az9dmx+/n5NSERDEcvdxcen37m2vFkwGXtUjVSalK6MdxEGnxCriQwT4Fl+MUGAr+JuiwDx4MB8H5P"
    "5N28Ki86Ob3+/rzSwE0P56/69OgHX1o/YH7PD/7okF5xb/5/vj2+eL7c9uf9U40hVpppL8NJFGzMrIcgXndK8qWftZjiie36xVuz"
    "Pz49HW5Zz3Q/mwdXLGrwzDV+pkd+0nOjB7Bpl23cRYqZggHwWtgBxEAj22IXX91E6KVzHTCqZ59UfcGRFbMCRTy8I7z8kqGbHm7D"
    "RShXhtze6FScQweQEF/7xfi/XDqEeR6hHN4dLYCy8KAElK4KeDQq6EvFgCyg4vHft4y/+dO7X/vM+qcuGMgVMCSOL39K5Ss/a7Xb"
    "kA5HZpiGkYqNp7XLppWgYwDMHtvBvTutXwF+hjn87XfHf3h31g6TiQsnaDftZaek565N8oCLTkq/cFP7uvvy6gJ+49bsjseGX3ZS"
    "esxK5wTbhu22R/OrHsgffFyljo079L2Xj/+vZ9UmPYGUc6N0ZUv3qSF8zNScfWhy4oH+1oeD9Ml7vt/80T3ZDeuD1Ghjtqh/H8qQ"
    "MkWmk9uyFSD2l0B6HsDuFkaDd6maasEiKyCRGCbgYOrqJsKuOm0kTpueg43iGYLBCz51y/hND7fnL5AoZ44Aw2BC7GoZe7VlgzUm"
    "8/3SAd/KJyJwKwoF2FHiRyJxqG2RAoTBaAA37lZJ8e3b2z9/WnbsiiQPUMMzDk5OXl25+hdt34DmAKUsVOl2ViysuWimHpVjAoPB"
    "IHb/5oIu3Kcc1US6R2Dck4+x0mSZ4NbHwq3rwrTzMoZs1VJ37tokVyys8T9f0XjxR0ce2BT8PD60U9/7vWYR/0euJoEbYBg3Cobb"
    "vVye7EV+7SKE613Bo8B11oMMY5T2jrMqF9076msczuwHd2WuVtjAuWt9N9MSTztJLCwsbinqYIlf8Wo4AmDbMJCDIgqxGPz35NsA"
    "IATkuYWAPFgeEIKGYLmiKzw2ACMtldSCWh4sD5prsQeNLtcmzSxZ0HbQzKytvXvQdrC2Wq6a6+SqgIhcUgdLsXNMv3Bbm2Uuo57w"
    "BcckBC0vMphQmSa3q53KGbIjadIMzJDkQ+NZFp7kix5rIoyajtrozLSvI6jIRywbNQTAc6Y9+irvEAxHLZPvvbnxnGMTGTXkRCfp"
    "7gyeUFQynHWYv/xt/e/+vVooZX3tUctGrT1qMxUCKzA0ZmHEwqi1u0bWcBNh1MKoDY+X4MuK0PClx6evO6eS7VTNDClDy3S3/unZ"
    "lWetTawcAEGxa6y4dPdpx9vF+xlu/lownQeweTfQgjQsVy0iqaIChb1ZvUn+ir1ayWJQRo65aPHcyW3Qiv4bXT5KzVyI5mswtQgT"
    "rOA91RQK7+CkAAA97UsRpSYQsR/8ovn2Z9SW9kns7vWSo9MP/mB84y5jkWnANJLVzq1ZF6xDgBFizWBjTRtsUG0fcEc88vQ1bjRL"
    "CZ62epqgN/KGB8yXf35J7et35D4pOw9P5zV3jtnhSx3KADooDlnkvv+mvu/elX3rF9m9W8KuEYTcag0u65fjl8vvHZGcdnBsZ1zA"
    "reX9fMFJKYRCLKhN76arDucfmwyNYjzYAQvYSX+edYijpTAcv9IVjEqphfGCj15YP3aF++pt2bZRXdTgRSemb3hapRBJGwD0V/jS"
    "E9JmblSs6jrtcavcs09OneC4lb+WVRU8gGYgQNNAimkoCqiilEe6+PqeIsAOXzBJEhSLmzqtY6JZQ2jCiTLw6HguOqZ+zPKk4Rls"
    "+kk5D/jQDcPXPtJcUPOI3U4xeUD11d1dm7KfrMsuOKZCQg0HLpSnHZp+6dq21Gk6XXBnEvWlpTiJpatifKR2y3Y3bbDBmejePSQ1"
    "3n5W5e1nVfaQ6YiT7B+eVvnD0yqzZQU4kQgk8PyjkucflQDYOW5ZsIGKVJOe1LqUyPjY/dy339A3U+Yl/j1Q5SWvaEwSXQH4i3Oq"
    "f3HONCM2Wm0ieMszKm95RmXLiC7pk86lO4Nm+QC/8rrGJAgO4I9Oq/zRrB/8SRp0nissA8QoRiVBE4vLRBTdNtAlAY1PZdPphSAw"
    "iDrGo2MCmyWcUifWnRA+YnFyxOJkz/f37XvHrjJLRMEg7CxCUaIlmvPQXL57R+u8I9NEGLntV5yUfun6ViEbmcpywGAdKoddKee4"
    "jC1MQ9An2WZhQoG0R14y6ARAtz2a8lQTzLWgnOfXijAjaDH0JoVTsdDQuuxpxpuZcrlgk2sKJ1l2lE0v6ZN45NTzd8ir7vX1Zvl+"
    "fimDRoj1Ax55pJGtq0xucsK0bJNrnTlosgMgyFgkURwZl5aY2jo/V2vPAFUj25GpNfPgqBTXqRXoyT/Q1DSp4dJ72+t36ppFLvq/"
    "Zx7u1yx3D25UV6HOFJX1likXL9kcjHHwzMY5Fz/PtLkhKxtJ2TRkQkzad5uRTUeDdNxe99clViZYYTFx2iuKCK2HOWYZm8bK7T0k"
    "zGXKZOJK2RmnHa6l/Ewxi2VOp32XtqdGW9M2pe+5OqdfNMeXvklpOURgNOl2o1OrBku5Gqfx0PGnizUr5fIpRpo4iPVWgQNeuDcY"
    "xXrK6CG8TFSgsTMb0MyslmL7UH7pPe23PL1mBgVqCS88IfmHh5q+VnSS3Pv7LiWkcSGYWWZ0ZapCmzOA697rha6kbtDpf7+Z+msL"
    "sXdWkZNxwh6Oj2Nj6mLpk0xZtUjQTmLfaNO7/2nZN3lSrTnjHSau57vT6hk9gEQMmpNpmWwo9M82wQiwuzS1YNRoXXbdmVwMYhJj"
    "H0GU6Yuj0ETAMjEbLee6x1rXbRjvSyR05rZy+RUrRV73bMsaFaeEuCluIHqGCIsSfP7W5htPr3kp6uBfdHz6ge+38jAt8cbuykIT"
    "0gAGiAGGYC6xNDHsmVEygNg2as28ZDhnsCs19lcwWGXn1UYx7ZYRfWzIBqtcvajAoJNsenfLhppwsFrCBfWJauodo9YMe7lorpxX"
    "xUC1UMePZ7ZjDJzuKwLMr7GaMDr4qTCgow8Rh1bAozvCeAYA9RT7DUg9ZSf33n3/QbFtVHMliSV99KX8ZLhlu5p7f2Opw6JG4cLi"
    "mcVhx5ht3KXBUPFYWGfE7uhaZ7Aw6ME0wEKxhJoa1YHBGGAecGU0FrGGMeZbpBDJT0WAHUBCBtAKqxOjakdzq4ADfvjg6Lt/vHPx"
    "gG+HkntGb1mVWdWzvyKgOjEnPVOXEHDmBJmhXuNtG7LrH8mefnASx98RK/wZhyaX35FFbDz5pRU5hvI6kTSniYRA1Krsr+6lc3as"
    "+n7TZ0dveyT4GrvLPXvVNtBxO+Jg/9U/bDgWKdJ2jvdcNv61m7LdTUsdjtjfve9FtSOWuY56JJ7/g5e3PntNSzyPOcR95bWNmA3e"
    "Omq///HRDVuDq3ImhsR7tHfr7z+t+rcvrGYBqceXbmv/0zeaLJV6k2LrRp1Hr3BvfFrlKfu7STnneCcEbl6ff/pn7RsfzrfvsiwU"
    "tOlgv5y4yl1wfHLOYYmUYzV+ZcMufcGHR9pto+Ijr2qcvdbnAYnDv13Z+syVLTdAnUFLLQ6haQctdBe/vrH/IDs15O+/ovnje/Lh"
    "cQuAFzQqPGipPP0g/+LjkkMWux4PvWq+Y4JcQSemxkB4mORQT3hjDqp0la6UwBCx6HUK9gFFaSbOUgFpuamQDpNH/0BFFvTL4gbj"
    "C4Kx0D9bh3qiGUJQIZzopLW7CYiYOKUxEY6O4pKbmk8/OIlBQSPBeSf6K+7MjHDOJiNoxmHZmXsY5Td0GRzrqTQq2CtnFxTrduoD"
    "WwMasdXwDET/uEk/rVTXGvDWr4197IdNNOg8g+Hhm9sPPB6+/9b+g+aXTRQMIB7brQ9uC0i4cAm7A4+HdoRHtwXUOJMmXzx0SB8f"
    "1s5QHhrH/fEr0+ZNhbesy790Y/u9L6m95YxKZ66IfzRz+9/fbn7k6tZYZnBdQbYAQ3bbo/lnr2ufutr9/Qtrpx/srUvyet/WkOWG"
    "HDvGJ6rBt4zoA1tD0ap2pjfWNCPyYAAThzs3hQs+NnLvJkUtrn1ZeKK7n9BLb8/+6bLWRSckbz27ethSKYK/gbp478zKxbedGUPX"
    "0ic2EUWVSFxmQHIR75LwjjTsbOrOVhAyFYpQeufUVtBAZZFfj/8s4xiCpXd3Ds7ROzrR7jiOAkf1VC8KWLWKy+5rbdilcXYzw+8d"
    "ney3hDaqbZ3MmsVpBxIggWJ0wTllpAXJRYMi3HsZFoFKQkmZVAjHiWrF7t0myPo4dd72WPjPa1ucJz5lGDcEVBbI/Rv0vZe1Jl0z"
    "9ZCUUmHVd4WtQLVz0Q5L0nvRom1M17kSD6nQp/QJ+6oc6Ow19lUJNVfhOOytnx/7+h1ZrLWOUcSull3wX6MfuLTZ8oADcizuk0MW"
    "uzWL3II6kQMGreKq27Kf3J93Zv84f9Yq9CmTChOZfCdphfR7emOmxagZzfDOr4/fuzFU5hE5FtS4ap7Mi/gtM3MYNfvY95p/f1kz"
    "QiYPoK/m+mvJDpBqoqBTFQW1lIJCrNMvtADKIpDIyU02aJL0jqNZSMTOOqgK2i2PNzND6ov5tPONZf1OJAxnFotxrbNyJqCmpSqR"
    "Zlp14p1N8tACFa/ixJsq2ahiw67wtZ+Pv+0ZjfhmD1ro/uzZ1b/5zOgJ+0sxqUiswAsRNVFNJDZdEMIo5r1mGQ9cKhEX+Vl0nFHA"
    "WnbRSekLj0ricmGTRnjIsaiPkdbwxI2P5MghKfKAV55euevxcPuGXObxcz9rvfOZlTWLpAP7zRC7/WivZCvEizbtlaekzzsyycNk"
    "8C3EeIajyqRMJ1GWt2ztEvePL6oN1qhdV7l9Y/jHS8d35wwJ/u7S8Wcf5msJIwvxV99ufuemdnWRNIftsGXyJ2dWTz7AL+6jGTYP"
    "650bw6V3Zd+7ub1mtf+jZ1QitdLBhapQBXvFTzFqz5p2xAr/V8+uTIXSJELA/AqX1gngns3h8vtzNyCtJl5zWuUPTk0Haxwas0d3"
    "6nUP5j+4O1u/MRx2mH/3c6txPvEAls1zy+alO7YFSWIyOIABAJzCgiOLNL+LYVj0o6QU2ZVep6XeWR6woMYPPmfhmQdUDbji4fG/"
    "vHLbrpb6MkqN7uzCI/qPWJSOZuplonaqrKAopl411hw/f9euS27fnbiezgniLHHmRc2BakJzXr92x9hrTq4NViWe481nV591RHLA"
    "AukE3eO5bR1VpOYSGE1gLkaWAMVSb2NtW7V4X+JwwgKeeqB7+fHpbA7f3TYTaGYrB92HL6zdviGc8c8jicPYsH38+tb/fUHNbBad"
    "7wgEnHaAf/lx6Wzyl50092DKcw9L+nq/9KxDvaP92Zeavo5Ht+hNG8JZqz2Iax/M//OqZnW+NEfsnMP8p1/dWDlv4gc/ZJE8/SD/"
    "6pPTb56QLqxhYa3QfO49o01YwIp+XrTHNxZP89hO1bZJjbUE//SC2vKBiQd6zVPSddv1ize2Tj/UH7xQYkTrzbC0X5YuSu7eHFwi"
    "SgbkE0FjR0rMnkCnbPdinOIbnNh4rq87sf/5h9RjIcZ5axu3PzH+rzcNebHu3gKNRE7drzYbI7h105iJTeKGhEi8pomQVDVVW9jP"
    "mza2f3R/62XH1mIwIYYjl7vuWHj7sP18k7o+xlViHDt5NY2LwCUpjl65z70UM+0ahpOgdm/egaWjqicYb+OUg9wzDnU/uS9nhd+8"
    "PfurZ1X7KwyzW1Asn5lcjG9+KnGWKXaOWcUxaOHXM0UiePqaZKDRHGlbRm4dtnjbn7+5nedQw6J+/terGivnSSuHlwnXboa+lK88"
    "Oe38IrMvYDHd47xnxf33V2lCETRb9sefH/2b59XWLpH+SiEtPnihvOu5tW5+xueKasI1y+XHdyjhQLEJk4sZQQJm1K7GRlbQFxJT"
    "G0VfVxMYTMScswMHnQGZKoCayOGLU+eUNM/iB45wUXXPzsUi+9hSS8VIiwRfeQZNRBNnUAtEMCYOtVTfd/XuM1anS/tcpnAstCWd"
    "go5//lFrx4jWBqmq4i0Sq64gz4MGLBzgEcv3zaBJPLJDb9yQZzmc67p7Q0IetlRir5mpdFg7t1TkZSelV9+Tpx6Pbtbv3J294oQ0"
    "z+FlovHhjBfdqTesz7vLSCPJUBEesUxqyTQXdUTVI3ETdukVInhwqw6PgxUkhvl1AGhl+OkjgVXoiL3i3MoB8yUPqHiMZzaadZVx"
    "WSFDNUNfBbXphAxTfS8Fu9p2/fp8mp9cccgit6ivUGYetdztv4Drd2ilj9/9efaD+/Jz1/izDvenHuTXLnaL+wignU8UUhXFg8ft"
    "Z+LVkJQpbZlozIai2Q8EpIIQgYhSILRaAiFSBxOSqCV0oqnH1RvGXn3cQM0LgHawHz0yUk+5o5k9uis7dEHamVPdXrrIFb03Gymd"
    "sJaw4opEeoyNEg8vCg+nDNCgWNjH+7a23vLNnf9x3vzFfa4rtQAA77+y+ZFrmsl8KFS8JoQXE1LiOJQwMs5Dl9lBC6V45FmR/pAq"
    "L76mfcl17e5sHAnL0ZfwK29snH6g38PQPWdtsv9it3E4hIDv3plddEIaTa2/SkyNvCMgUbDKD1/d/sQ17e5UJQnNsaTCb72176jl"
    "buKiZdovADvGLQL6iHedw883hL/73pg4aM7lS+UpqzyAHWP6yE6FEGJPPdhbqdb43M3ZJ69osotj6QhvlvXx/RfWD10se31jrsJb"
    "N4Zn//vI1FFq4/Zvr2z8walpvNzSfr73RbXXfHKsNax+QCzHpXdml96RVes4dpl/6iHu/GPTM1b7zvRYdD44aSXm1XVnTicGcZ0o"
    "g1F9IQQpVAgcGXU9oKUJNg3nP3poHIQaGinXDWWp40CFVz069meXPfHyo/pU8aV7d125fmRRXYba2buve+JVR8zrS2KcCcyoZSiE"
    "zwpUHB4ZavdXbEczu2b9KMGgWk/cI0OtmlcnIKkqTiy2fVk+n1c+PPrSz2V/+JT+Mw6q9VfZbNudT+QX3zD+jduzxgBJCMw7iJin"
    "uchSiaUujLTl5AOr3nFqmmPPoLCVodWapPwFcmTB9uCxIiQ4dJGcudZ/5rrgK7zm3vzuzeHIpW7vSpJ4UbWp2Dq1AgL1Qh9Ilfds"
    "C8/7yHAiE87bzJ7YbbuaBk+M258/szKvxsi7tULBow7WJpoor9sWrr87Q19XIw8DEsABY3bGkck7zqrM5o2FgOG2TcvZNcvF9oTI"
    "FRedlFYT/u13x+/YGKBAlYlHUNy4Ib9xXfbfV7dfekr6vhfX5lWpBh+pomNXpgctkx0Pt1FJigksdjCiGiVSy51miXSFNLS/gls3"
    "ty766lahiRNHTZwNVMWo/RV+Z93uyx7dHZMvgxUx03oFt28du/2qsbTwfkbp9C/qdEmPWvtCo2xWwPSl/XxgqPm2yzYUXTuMAOY1"
    "AA1BxVxu6s0QzNRsySAe2jX+jh8056W+L5Xdbds6Ys02Fs3zRa9zmHPmBJ5wMAeCKhIaNXn6mn0qmy/g4ECd86vsVlELoQF9njXf"
    "GaDTC/SEOO84/+Wb28Fh4w698v78qGWz6H9qGKxxoMqg3f1dYTkWVaTqp+0Ti2aOhzZpjwE5wDGpsC+3t7+49rpTKrnBE4NVLqpz"
    "025DwOYh7bAiBy+So45MkipCpz2bcf1OHWtb4Gwbt5oiTbB4QCbNPyS0ZoM1dmv08oDzj02edbi/9M7sa3dmt6wPT+zQkZbBoTog"
    "uwM+fnlz+5h9/rV1L/RxEFQSOX1NcsvDGRIgGEGQDoEMcAaJq0OQErt7GaWgOypCOnXCWOAQP49Z7sFECBXCORrgDCAaKZ3FjExH"
    "XEoKYuexYvKkFu1DSzY12rs3M9dZFIMaRdRm3lSDqJoB3iwu5b2wj6bSzmxniwZb3A/SZZYb4EhQHTXK6IUQ0DmMtfOVi9Lj9k/3"
    "zZwFmtkbz6m87YzK7ha6OddIBy4fKADMTKhDDeeuTdYsl7s2KQTfvK39ulPTvnSG1tFFRS+0aW99XuWPTqu0Qk/GKlaMLx8sLxp6"
    "UHvquGIRPRnLmXe3bXvLjAhtfOz1jQuOSzst3wfrXDnoNu3MQFx2X/7aUyPcw+tOrbz2lEp3nL1tTJ//HyO3PqZVj0WN2fSbQmjb"
    "KQf7z76mMQlPxRhgST878sC4qgSAvpQXnJhecGK6q2k3PRKueii78p78hocDK0gXyDdva19+evr8w5OJmpnnHpn++zUtlZwqAgFM"
    "qKSqGKTsQCdWdJ0Ti/8p0W1HoUXMnhARXhOQgqg2V4YglEJU1KnmIrSTjBUUTaRhKP/VyezQzFQ7aASmhDeYaSAcQ0FnQ9XMLIAw"
    "em81M9PoTbRahn+gCtXTHCEGwtLEjQY9+cBkeb8z28dGM4bFfVzaL/PrEy7KuirhZwIwkfTJFY2U5x2T3vXIuG/w2gfzOx4Ppx24"
    "l5JTGhb3c+U8mXbVAdXei0YOvmWH7ecuflV9fpVRWLl73C785MgjO8xyu/6h8KKjC+o96pDOWu1+dl/mBvit27MbHslPPdBPam5r"
    "htTj9sf1vs1Kz3l1PmV/mQ3egKKe8MD5Mkkw00mea5kxdYINQ/qpG9uvPzXdb1CygP4Kn3WYf9ZhfufZ9rFrW39/aTMTaIa7Htfn"
    "Hw7p6MefvqZ66HJannkP0hhTaGIiStHC7xIiFIfog4UQZ+IgjEFVNHEVUZEYO5qIOR8/NBFzNHHx4HJ35sUSZ76ze0u8+cQSj86H"
    "idPUI02Qeku9VRJLU02dpl7TVNNUK6lVU6umWq2gWkE91VpFa1WtVfK+Whio5/2NMNAIA43Q36cDDR3ss4GG9de1rx766pqmeb2u"
    "z1pbTxz3eYU3QzPrkdta19+yt0LUyMq/5uS00U8AWRsfv7ZtwFjLuiWPk5c1IKJICOxeaqm8qEx/0YrDIQvl0MVyxBJZu0iessr9"
    "5blVa1ta579e3vzxA7kThHIsXnhymtQpamM5Xv7J0Z+uy2OZoC/3isctj+Vv//LYKGCZnXSQX7PIZfms3lmmRbWnTXlp8Y1F4uAD"
    "VzbP+deRv/nK2HkXj17zUO66nmt+jX/ytHRhH+MVQ6yf73Cl9VQuOrryd+ubrqqZBREQAjEvwWhCV4AKxvITc6XbjslqIRxjr1FQ"
    "zNGcM0rsI2qOiA0XBSCLtaiKVGPRxTTm1SOkJkutmwFRgzOlFTXNYBJTpWKWd4pqCzWfmSEmqWJcG7OgkaWhmDpnrjyrgwyHcMiA"
    "nLSi0iWJm601I+Unftpev01lioBYiJDZqsX+L86tTJs46MhB1yySFx+TXnJtKxngZ25ubx62WzbkrEEyNNIugVWUKygs4cU/zTZs"
    "03xKXzUnaI7bWYf7i55S6a35RFA0syKzHVmO151SueSm7Jr7MyZ417fGzljdX3GkQA0nrHBvO7Py/m+PN5bIo7vs3A+NvOSE9PeO"
    "SJYOIHHcvNuuuj+75OZsuKVe6BO+61nVbp3qtG2pSSDA1fmzx8LrPjs2UEGYUkWbZ7Z8QN70zOrCGn90b/7AQ1nfSnfzhnDOvw+f"
    "uTZ53uHJ2mXSqHD7iH3iutbGXVqrsJXw6P3chEHHM770hPo/X91qqvo05gkkUlqxQDD6XdKir41eNlJ4sZRXxIQqTkRMnDmxaN/x"
    "4JLFNkf0GHRZAMuiYzNQWLx0i/2ll78yUyIuoxhLGzHRWhpFXQtLlityMlYAJIpTgfoYEcQVsgzNVvvpBy5Y0VcJZjK7+lgrBFJI"
    "Knh4p/7nVa3ppTZjtubg8OfPrHREDs4BgqRLEBsppzc/o/LVm9pZbpLwB7e3UWVSYTassYG0WakMBsQhqeDBbeFDj+XTq9yHLIAv"
    "f0rZAlbgPILBu1imCRqiZsMJ/vEF1XM/mIcablsf/uMn7XeeU4mdMYLiXc+pPrQ1fP36thuUlsMl17UuubaFFCCRGQhfIw02Zv/6"
    "qsZTD3KdTkMEElc0Kuh+nSFAHJIEzcw+c11rmkjZAU07eLG76LR0YdX91yvrrzf88LY26lTPy+/OLr8jgwc80TJ4pDWOb9VzT0jP"
    "WO210wgv5njWLvUvOib93C2tSg250ZlFiE0hS/wQlc3ORfAAoRZI2lHYkT7TFQeL0CiIllTWVqjQYHBRR0GjxdotY1noJUTZyLSz"
    "WGHJXUcLnNBvRYOc6HPOsvgkyvEARoCECJAI50CYp4o4MzVjUF1RwTkr52MW0vlCb2Dor7Bao+6yYvmSaSuKBHA9GHcsQ9htcNhd"
    "nWD0om2duL/7u/Or7/zCGBKgRuTItui5Jyd//NTUSkOppRxIoEPWymMAOM39ioM6604TjrUtHzEoRkctt8mFM089yL/mjPRj32mi"
    "zvd8Y+xpB7unHugjip1X5X+9qrF6sfvoT1ojuw0pkJYTWErklo/Y6qXu3S+q/v6JRQv0OPAyxdCwITPkKKq+DQAGq6LD2ozte2W6"
    "ugYHCJwrSgH265cv/2Hj4mv9J69t37M5wAAXaT/AE7m1h+zZxycfeUV9XpXBelfBShxfe0r163eP5fDOwStcXDJOlC4644iqY0QY"
    "hFYs/i1gJDpiAZmUjc0dZKL+ykgINHr0uBQmY0VtuRJbDARZGDRKig1lRWN050WfiZLRix+WSjWCFr9uKKNSJ0aaE3HOROidAsGJ"
    "UYIpCA7nelT/4KGDfRo5kFlENTEMf8/zqivrHM5tJnGeEHkbBy6dEBoevdydf2oKwfIBqfsJjBJN4X+eXV3RLx+6urVupy7rl1ec"
    "UHnDGdX+KjsMwIIa3/+S+sU/acY2DdMKJ7xgdMxO2N+zTA8dvtSdf3KqhtWLpeYnD1pH/M+zquMjNqI23sKdm8NpB3kr+w/Oq/J9"
    "59X+4KnpV27JfrIuf3B7GMtgQH/Co5bLs9cm5x2f7jfAYF3tsYGBKl7+lLSdGxSr5k1oaV5/erp1WDcNq7jpb14EeYaVA9KXShwY"
    "AxW+45zqq0+pXL8uv/r+/K4nwuYR3dWyRXVZu0BecEzygmOSSAo5ghNL+BXlGvqSS7Z+50HMrzoL5mAiOUTpVYRCo4OniphzSjHn"
    "zRW5QzrCkc6BoiJagpBeg6Y6htJVsyBJUFSGx+qBuHxxgUMKnKJC0MhSPRINnnHB2SgtkY7czyRyi4zDjF5CvEnvnHPwEigaFSkR"
    "izQzfdXKI9b0z9NZ441fx2ZdSo/d41pLGbPfT6Lfza/kNroNo5M3Hc9sNDMa6qnUkglp+CQt5K9w01K20K1UaQer+omqvE5dQs+y"
    "brlZ4uTNT+u//JHtmaWJ8zR6CXBmwuikXYTLkYlzcKSnipAuxojqBFK688h+MMISFpfvBKqFQSPE1F3RGaEwUBNqx6DL0lvtBL/l"
    "csvaiR9jfZcU3r1kXRxJiNOEJg7ehVToXKAzieWwJiPt5nF9S9b0zzPsszWrzQqidLe4ta6VjGQ6xx91NoN1mTAU7ttFrbuUsPei"
    "M9WVWdnqbuqp49QfrUoEtYS1hBNaUCuaIe3phN13EusGZjOoOCGtnmhgH5f2EvhSUGbouQHfO1VRDc9cXX/pkSOfvSNfWK/Qip6L"
    "sYrQF7bCKCN2LIoFRYyEozkxTpB3pX0XZh2rwUvqGiQholKs5FB43xKKWGnN8f8WDY6ke5EX6SzkVtAsLBqAqDi4AutTnNGpozki"
    "cfAO4oI4E0eK0zx44+kLVs0OPE9X8jmr1Pg09sSZWTyUZZcz1pnOXjI6sxFP+gr3uvAXe8cG99RCljOvl+y4b2+s8wg9bU5t+hWU"
    "OWlltDi7PbS9fcYnd4y200aSU5s+EYPRw0e861ScipgvGQwnEAcnwQnozEVMIrFuyoRwQsIocLHlWJexOkRYEvu+FCsPkdpB0lK0"
    "qpng+yJccVJwdAX3QhAUMRfRDiPsofOg00itJA7ew0lOZz7xVGnmemR68ImN1d0dP+a2//9ufuo4UMPqhelfntF4x6VDC1Jm3oxB"
    "RIr6KEJEhOYL+oIxhxIpPEoQkTgFlERHmXABy68XF/IsCnOjS5YCNBtpjhaVfQSFWpg+KURc5EdoQgGVMEfSUWACc6Ii6l2MAlWE"
    "zoEejpYQ3iHxRgdxdMI8CyuSgSPrK+fs4LfWoCPRpIbXn9j3kwd3X/Fwa/6Ay8wclVKkuJ0U2T4Wqku4+HmRTVShUZSiQnOkK1dt"
    "i1EaxaTTfjzyITHbQnNlbxrp5GgK36xCcfHMkeajiQjFImh2LMgWBxOnnhqJcCd0LqpN4AVprBLwRs+Qh5p3xySrq6zMueffmo3T"
    "LcZaxIz3bmm97IubdqjWUkLNCcWRVC+IRgNGlsMKlkM0Er3RRzoxoboy51IA30hdF+W0LHIxoBQpGJQCjyCIfjry4OaKYzRqmcq6"
    "xnI5w4jUCU8TCV4KMZ0jnEecJbwgEYhTc7ETOw6W1Wv9YXPW/Ntv0AAytUT4pbuG3/LDTQM1cRSSzgGi3tFRpeiVoRBEbTGJDnR2"
    "RR4xknoWIYcjCnFIx6EWK23SF1659NzUaOKR4PAFIDGhlTxGTLDDCaUAJFbSLOoFIsELndA5E1pCuPI+TdDW1iK39AR/isDNGcFv"
    "OeToMB7BcOFR/fdvH/nobTsXNtIcAlKcizQCC1Udo4YJcYqnCemEE7nxoqVGSbpFLR/LzEvUMFlwBSxWFsyJgeooJXERz6wisQS9"
    "1Po5OjMRugi7BSJBJIiYF3jAOXiBQL2jL5gZKEK/NA53RwncnHv+XTHoSJ+p4Z2nL318pPntR8aX9NVypSO8E5g6gSOdgGIQmJRm"
    "Jx0/XRB2BYuHwisL4CPF0dHfMdq9FYQdY0uxYplEEaMg0hSOcEUEaSIiYs7giytGoXaAhBiSRv/toUL14ilGKkS9yuFyXIMDc9b8"
    "OwQ5uin6ofH8j3/4+M2bxpY3pG0mzgklERXm9IHOKM5BvAsiwTlKYUwmDFJ4zTJui12ZBCX2jUA5kiGFQUeUgsKLUyJfURh0wQOW"
    "YyPK+ugIF8k+CSLqi+GEhPCEcwRyERUXEPQQOXmpO3DOmn8rt711/wSCYV7N/9s5y5+61O1ujtZcnjrzVGdaKKEBB3hQULRU5ERf"
    "XRZZD5kAGOz6I6asWYAQ6z6s6DHg0ElP+oIN7OwUb+IKRbUXS8gKXUVcSpeQqTDW5MRMk4hS84PlmDlr/t310AXpYebIjbtb77r6"
    "0du2tBY3qmoWvIqzuGxBoUnyCmcxSvMRZjgVmi946IgcLMo8WJAV8GKkehSut/DKogIF4UjGDA4KjM6OUxfS5Y7woCOTQrxljkZf"
    "OHVXUISktIn2wTx+uVs7Z82/6wYNIKp2toxmf331hp9uay/uTxSZUBOiYhYZDziFoxNhSXH4goouk+Gx851E/I3IvhWBI7SUp7Lk"
    "p0PR+VE0KYoJlDAvEaJQxChBHBLSA97K9TydMqIdmhPzToyZKA+W45a5uYzgnEH32vRwO/zDzU98a8OuRVWpiYrlnhAJnjAH5+iE"
    "oPlo0IxFsio0OPOMaT/zYoKYYtTIQkRv6qTMHbJkQgAR9TEBTnM0H3M3FkXY5sQ8mQJixSTAaNCikdejZBVN18jJC2W/OWueM+hJ"
    "2CNqZO3ie7Z86r6tVWJBioB2JMwSD1+Sx96ZiBbccOGS1Umk5NSJIpaOCT1NoOLgi7SfUYociivKy6IwtdCjeimSjrGYwNE8LBYN"
    "OSic0hmIWBmZozkfC9by5D4umLPmOYOe1k8XMd/VG4c+cuf6Lc32vJo3gVISqBd1ImXNbGGCjiZOHQqZnmNeZGEIEkn04iypPYEX"
    "oahDcEW5jNEhIcQiktGJTCEhNA84wkHpFFQ4pE5yawO20h28mkcnc8ntOYPeE5dXilA3jox/8p6NV2/eXU3Yl9AQSDgnkYCTmDJ0"
    "MW8XBR6F5YkzUBzNicaUiiskpkYiEREGh9hNzxxLVhtlFSO0oEQMAvgo6ysaO8EIZd7H/gPl0BXuQABz1jxn0LOBH+bIXO3Hj+/4"
    "6oMbt46MVaveJz52L3dizsFJEEQ/HXvhIWKPEiVHNGJCekbqwwg4B0cVaJTblgLUKMkwJ4hJxwJho9SpUkkYczG/wu1/sDusKjWb"
    "ZiWnuW3OoPcYJgLY2Wx/5+FNV27euVuzeiJVL7EswEnOLoN2nRYIRbkhRNSXTXhL2i6q5zQeGbOGvpQfuVgFUwaORZc6GkWVbQHm"
    "cdEaf9hitwyAQgUy9xvPGfS+wA/ASrN+eHjs+xs2/3zn0IjmVS9VLxQF1ElRMFI0pmEQRjAdmZAo8wilxsMcWWibUKpVC2FqEFIY"
    "hR8UoQMgwdgWw3yZd0By0H7ugNKIbc4xzxn0kzdrNXOFWY/csm37z7Zve7w9niRJNRFPTVB0F4dTkyBCT3qUgk8YGSDqXIz84GOp"
    "Iks5XtRpOIOYQD3FkQpTy2rOL/WLV7qVi93ShImVyyjO/bRzBv3LbmowFGa9o9W6e9fOm3ZsW9ccbpqmQFVYJSVBcLmjedIbHct4"
    "EQqxxFmsOXOlq04KkZ1B1GAUA5QwTzfAxop0yX7J8kEZdIz9Z5VzGGPOoH+1Z1Szop8ooGZPNMd+sXvng6O7tzRHR0I7WNs5SxwS"
    "SkLnBMJAYYdjjtUrRjUXVSJKC7Eiy5M1lw5IY6EbXJEuWZQsSOjLKWLOMc9tvx6Dnoqt47Y7az02PrppfPeO1ujO9ti4tccsBCpo"
    "gTkkCM2LxeUM6UjPBNInST/Tea46zzfm+b75vn+e6xfKnB3Pbb9Rg55k2R2fHbdWyMdCezhvZ6YtzUet2dKcUbhHOLLq0opLK3Q1"
    "JnWpVCXpMmJgzo7ntv+vDHrCss2sbCCyr012rKfPFucMeW6baft/ucNWFMf+c+cAAAAASUVORK5CYII="
)


def _lum(h6: str) -> float:
    r, g, b = int(h6[0:2], 16), int(h6[2:4], 16), int(h6[4:6], 16)
    return 0.299 * r + 0.587 * g + 0.114 * b


def _apply_light_theme(wb) -> None:
    """Dark fills → light Technip Energies tones, white text → dark blue; TEN logo on every sheet."""
    import base64
    from io import BytesIO
    from openpyxl.drawing.image import Image as XLImage
    cache: dict = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                f = c.fill
                if f is None or f.fill_type != "solid":
                    continue
                rgb = f.fgColor.rgb if isinstance(f.fgColor.rgb, str) else None
                if not rgb or len(rgb) < 6:
                    continue
                h6 = rgb[-6:].upper()
                if _lum(h6) >= 150:
                    continue
                new = _LIGHT_FILL.get(h6, "AED3FF")
                c.fill = PatternFill("solid", fgColor=new)
                ft = c.font
                key = (ft.name, ft.sz, ft.b, ft.i)
                nf = cache.get(key)
                if nf is None:
                    nf = cache[key] = Font(name=ft.name, size=ft.sz, bold=ft.b, italic=ft.i, color=_TEN_NAVY)
                c.font = nf
        # logo + room for it in the title row
        if ws.sheet_state == "visible" and ws["A1"].value:
            ws.row_dimensions[1].height = max(ws.row_dimensions[1].height or 15, 36)
            a = ws["A1"].alignment
            ws["A1"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=a.wrap_text, indent=9)
            img = XLImage(BytesIO(base64.b64decode("".join(_TEN_LOGO_B64))))
            img.width, img.height = 84, 33
            ws.add_image(img, "A1")


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
        f"Flow case:    {_RUN_CTX['name']}  ({_RUN_CTX['mult'] * 100:g} % of normal flow)",
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
        "  Mach (vapor) = mixture velocity / vapour sound speed √(k·Z·R·T/MW), k = Vap Cp/Cv.",
        "  ρv² Vapor / Liquid = ρ_phase·(superficial velocity of that phase)², kg/m·s²",
        "    (shown on Pressure_Profile and FIV_EI_T2.2; blank where Cp/Cv or densities are missing).",
        "",
        "LINE DATA / INSTRUMENTS / CONTROL VALVES",
        "  Line Data          — per Line No: conditions at the highest-pressure straight pipe, vapour & liquid properties,",
        "                       all IDs under the line, flows, rho-v2, slug frequency / holdup / bend force.",
        "  Inline Instruments — rows with Instr Type PT / TT / FT / FO / RO: inlet properties, H2 / H2S content, pressure drop,",
        "                       downstream flash, orifice choke check; thermowell wake-frequency block for every TT.",
        "  CV Datasheet       — one column per 'Control Valve': conditions, flashing, Cv (ISA/IEC 60534-2-1 screening),",
        "                       velocities, composition.  Optional tag data: 'Instrument_Data' input sheet.",
        "  PR databank        — 98 components (Tc, Pc, omega, ideal-gas H, Peneloux, 3302 kij) embedded in this file.",
        "",
        "MANUAL COMPOSITION (no HMB stream)",
        "  Stream_Composition input sheet: vapour / liquid mass or mole composition per Seq.",
        "  Later streams are mixed into the running flow and re-flashed (see input Notes sheet).",
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
    global _COMP_ISOTHERMAL
    _COMP_ISOTHERMAL = any(f.lower() == "--isothermal-comp" for f in flags)

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
        # An HMB is only needed for Stream Lookup rows / composition constants.
        needs_hmb = any(r.get("Stream Lookup") or r.get("_comp")
                        for r in read_pipeline_input(input_path))
        if needs_hmb:
            raise SystemExit(f"HMB not found: {hmb_path}")
        print("  (no HMB file — manual-property input, HMB not needed)")
        hmb_path = hmb_path or ""

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

# ── Built-in PURE-component library (fallback for names absent from the HMB) ─
# Used ONLY when a component is not in the HMB COMP_CONSTANTS sheet (HMB values
# always win).  Columns: name, MW (g/mol), NBP (K), Tc (K), Pc (bar), ω, SG 60°F
# (None for gases).  Literature values (Reid-Prausnitz-Poling / DIPPR style) —
# VERIFY against your project property basis before design use; add an HMB
# COMP_CONSTANTS row to override any entry.
_PURE_LIB_ROWS = [
    # aromatics (light → heavy)
    ("BENZENE",             78.114, 353.24, 562.05, 48.95, 0.2103, 0.8846),
    ("TOLUENE",             92.141, 383.78, 591.75, 41.08, 0.2640, 0.8719),
    ("ETHYLBENZENE",       106.167, 409.35, 617.15, 36.09, 0.3030, 0.8718),
    ("O-XYLENE",           106.167, 417.58, 630.33, 37.34, 0.3100, 0.8848),
    ("M-XYLENE",           106.167, 412.27, 617.05, 35.41, 0.3260, 0.8684),
    ("P-XYLENE",           106.167, 411.51, 616.23, 35.11, 0.3220, 0.8657),
    ("STYRENE",            104.152, 418.31, 636.00, 38.40, 0.2970, 0.9110),
    ("CUMENE",             120.194, 425.56, 631.00, 32.09, 0.3260, 0.8670),
    ("N-PROPYLBENZENE",    120.194, 432.39, 638.35, 32.00, 0.3440, 0.8620),
    ("124TMB",             120.194, 442.53, 649.17, 32.32, 0.3770, 0.8796),
    ("135TMB",             120.194, 437.89, 637.36, 31.27, 0.3996, 0.8652),
    ("N-BUTYLBENZENE",     134.221, 456.46, 660.50, 28.90, 0.3935, 0.8601),
    ("NAPHTHALENE",        128.174, 491.14, 748.40, 40.50, 0.3020, 1.0250),
    ("1-METHYLNAPHTHALENE", 142.201, 517.84, 772.00, 36.00, 0.3400, 1.0202),
    ("BIPHENYL",           154.211, 528.15, 789.00, 38.50, 0.3720, 1.0400),
    # paraffins / naphthenes
    ("NC6",                 86.178, 341.88, 507.60, 30.25, 0.3013, 0.6640),
    ("NC7",                100.205, 371.58, 540.20, 27.40, 0.3495, 0.6882),
    ("NC8",                114.232, 398.83, 568.70, 24.90, 0.3996, 0.7070),
    ("NC9",                128.259, 423.97, 594.60, 22.90, 0.4435, 0.7219),
    ("NC10",               142.286, 447.30, 617.70, 21.10, 0.4923, 0.7342),
    ("NC12",               170.340, 489.47, 658.00, 18.20, 0.5764, 0.7487),
    ("NC16",               226.446, 560.01, 723.00, 14.00, 0.7174, 0.7733),
    ("CYCLOHEXANE",         84.161, 353.87, 553.80, 40.80, 0.2096, 0.7835),
    ("METHYLCYCLOPENTANE",  84.161, 345.00, 532.80, 37.85, 0.2305, 0.7536),
    ("METHYLCYCLOHEXANE",   98.188, 374.08, 572.20, 34.71, 0.2350, 0.7740),
    # olefins / light
    ("ETHYLENE",            28.054, 169.42, 282.34, 50.41, 0.0866, None),
    ("PROPYLENE",           42.081, 225.43, 364.90, 46.00, 0.1424, None),
    ("1-BUTENE",            56.108, 266.90, 419.50, 40.20, 0.1919, None),
    ("ISOBUTYLENE",         56.108, 266.25, 417.90, 40.00, 0.1940, None),
    ("BUTADIENE",           54.092, 268.74, 425.00, 43.20, 0.1932, None),
    ("ACETYLENE",           26.038, 189.15, 308.30, 61.40, 0.1870, None),
    ("HELIUM",               4.003,   4.22,   5.19,  2.27, -0.390, None),
    ("ARGON",               39.948,  87.30, 150.86, 48.98, -0.002, None),
    # oxygenates
    ("METHANOL",            32.042, 337.70, 512.60, 80.97, 0.5649, 0.7960),
    ("ETHANOL",             46.069, 351.40, 513.90, 61.48, 0.6452, 0.7940),
    ("PHENOL",              94.113, 455.00, 694.30, 61.30, 0.4438, 1.0600),
]
_PURE_ALIASES = {
    "N-HEXANE": "NC6", "HEXANE": "NC6", "N-HEPTANE": "NC7", "HEPTANE": "NC7",
    "N-OCTANE": "NC8", "OCTANE": "NC8", "N-NONANE": "NC9", "N-DECANE": "NC10",
    "N-DODECANE": "NC12", "N-HEXADECANE": "NC16", "ISOPROPYLBENZENE": "CUMENE",
    "1,2,4-TRIMETHYLBENZENE": "124TMB", "1,3,5-TRIMETHYLBENZENE": "135TMB",
    "MESITYLENE": "135TMB", "PSEUDOCUMENE": "124TMB", "XYLENE": "M-XYLENE",
    "ETHENE": "ETHYLENE", "PROPENE": "PROPYLENE", "1,3-BUTADIENE": "BUTADIENE",
    "ISOBUTENE": "ISOBUTYLENE", "ETHYNE": "ACETYLENE", "AR": "ARGON",
    "HE": "HELIUM", "N2": "NITROGEN", "CH4": "METHANE", "C2H6": "ETHANE",
    "C3H8": "PROPANE", "NAPHTHALENE (C10H8)": "NAPHTHALENE",
}


def _pure_library() -> dict:
    """{NAME_UPPER: {name, mw, tc_f, pc_psia, omega, nbp_f, sld}} incl. aliases."""
    out: dict = {}
    for nm, mw, tb, tc, pc, om, sg in _PURE_LIB_ROWS:
        out[nm.upper()] = {
            "name": nm, "mw": mw,
            "tc_f": tc * 1.8 - 459.67, "pc_psia": pc * 14.5038, "omega": om,
            "nbp_f": tb * 1.8 - 459.67,
            "sld": (sg * 62.428) if sg else None,
            # Vc from Zc ≈ 0.291 − 0.080·ω (Reid) ; ft3/lbmol
            "vc": (0.291 - 0.080 * om) * 10.7316 * (tc * 1.8) / (pc * 14.5038),
        }
    for al, nm in _PURE_ALIASES.items():
        if nm.upper() in out:
            out.setdefault(al.upper(), out[nm.upper()])
    return out


# ── Pseudo-components from MW / Tb / SG  (Exxon-style Tc, Pc; Edmister ω) ───
# Tb in °R, SG at 60/60 °F, MW g/mol ;  Tc in °R, Pc in psia.
def _pseudo_constants(mw: float, nbp_f: float, sg: float) -> dict:
    Tb, SG, MW = nbp_f + 459.67, sg, mw
    tc_r = math.exp(
        5.00788 + 3.45052e-3 * Tb - 5.35056e-08 * Tb ** 2 - 1.07402e-09 * Tb ** 3
        + 1.64910 * SG - 0.55505 * SG ** 2 - 1.51362e-3 * Tb * SG
        + 6.72608e-07 * Tb ** 2 * SG - 1.06298e-2 * MW - 1.85057e-06 * MW ** 2
        + 1.14428e-05 * Tb * MW - 2.26285e-3 * SG * MW
        - 1.59585e-09 * Tb ** 2 * MW + 1.55423e-3 * SG ** 2 * MW)
    pc = math.exp(
        5.70163 - 7.32144e-3 * Tb + 2.23949e-05 * Tb ** 2 - 1.41165e-08 * Tb ** 3
        + 9.73622 * SG - 3.30416 * SG ** 2 - 8.70625e-3 * Tb * SG
        + 3.77458e-06 * Tb ** 2 * SG - 5.51438e-2 * MW - 1.74986e-05 * MW ** 2
        + 4.56742e-05 * Tb * MW - 1.19893e-2 * SG * MW
        + 4.17127e-09 * Tb ** 2 * MW + 8.25884e-3 * SG ** 2 * MW)
    omega = (3.0 / 7.0) * math.log10(pc / 14.696) / (tc_r / Tb - 1.0) - 1.0
    return {"mw": mw, "nbp_f": nbp_f, "sld": sg * 62.428,
            "tc_f": tc_r - 459.67, "pc_psia": pc, "omega": omega,
            "vc": (0.291 - 0.080 * omega) * 10.7316 * tc_r / pc}


_USER_COMPONENTS: dict = {}     # NAME_UPPER → constants, from Pseudo_Components sheet


def _read_pseudo_components(wb, usys) -> dict:
    """Pseudo_Components sheet: Name | MW | NBP | SG | [Tc | Pc | ω overrides]."""
    out: dict = {}
    if "Pseudo_Components" not in wb.sheetnames:
        return out
    rows = list(wb["Pseudo_Components"].iter_rows(values_only=True))
    if len(rows) < 3:
        return out
    hdr = {}
    for i, h in enumerate(rows[1]):
        if h:
            hdr.setdefault(re.sub(r"\s*[\(\[].*$", "", str(h)).strip().lower(), i)
    need = ("name", "mw", "nbp", "sg")
    if any(k not in hdr for k in need):
        return out
    for r in rows[2:]:
        def g(k):
            i = hdr.get(k)
            return r[i] if (i is not None and i < len(r)) else None
        nm = txt(g("name"))
        mw, nb, sg = num(g("mw")), num(g("nbp")), num(g("sg"))
        if not nm or mw is None or nb is None or sg is None:
            continue
        c = _pseudo_constants(mw, usys.to_internal("T", nb), sg)
        c["name"] = nm
        tcv, pcv, wv = num(g("tc")), num(g("pc")), num(g("ω")) if "ω" in hdr else num(g("omega"))
        if tcv is not None:
            c["tc_f"] = usys.to_internal("T", tcv)
        if pcv is not None:
            c["pc_psia"] = usys.to_internal("P", pcv)
        if wv is not None:
            c["omega"] = wv
        out[nm.upper()] = c
    return out


def _build_pseudo_input_sheet(wb, usys) -> None:
    ws = wb.create_sheet("Pseudo_Components")
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:G1")
    _c(ws, 1, 1,
       "PSEUDO-COMPONENTS — define a component the HMB does not have from MW, NBP and SG.  "
       "Tc, Pc (Exxon-style correlation) and ω (Edmister) are estimated; type Tc / Pc / ω "
       "only to override.  Use the Name in Stream_Composition.",
       bg=NAVY, fg=WHITE, sz=10, bold=True, wrap=True)
    ws.row_dimensions[1].height = 40
    _hdr(ws, ["Name", "MW (g/mol)", usys.hdr("NBP", "T"), "SG (60/60°F)",
              usys.hdr("Tc", "T") + " [opt]", usys.hdr("Pc", "P") + " [opt]",
              "ω [opt]"], row=2)
    ex = [("HVY_AROM_C11", 143.0, 4.9e2, 0.97), ("HVY_AROM_C14", 184.0, 6.2e2, 1.00)]
    for i, (nm, mw, nb, sg) in enumerate(ex):
        r = 3 + i
        _c(ws, r, 1, nm, bg=AMBER, sz=9, italic=True, fg="808080")
        _c(ws, r, 2, mw, bg=AMBER, sz=9, italic=True, fg="808080")
        _c(ws, r, 3, round(usys.from_internal("T", nb), 2), bg=AMBER, sz=9,
           italic=True, fg="808080")
        _c(ws, r, 4, sg, bg=AMBER, sz=9, italic=True, fg="808080")
    for r in range(5, 40):
        for c in range(1, 8):
            _c(ws, r, c, None, bg=AMBER, sz=9)
    for c, w in enumerate([22, 12, 14, 14, 16, 16, 10], 1):
        cw(ws, c, w)
    ws.freeze_panes = "A3"


def _read_comp_constants_hmb(hmb_path) -> dict:
    """Read the COMP_CONSTANTS sheet embedded in the HMB (if present) →
    {NAME_UPPER: {'tc_f','pc_psia','omega','mw'}}.  Cached per file.

    If the sheet is absent the built-in pure library + embedded PR databank still
    resolve the common components (build the HMB with hmbbuilder.py to get the rest).
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
                    "name":    name,
                    "nbp_f":   _f(row[5]),    # F
                    "sld":     _f(row[6]),    # G  (lb/ft3)
                    "vc":      _f(row[10]) if len(row) > 10 else None,  # K ft3/lbmol
                    "mw":      _f(row[4]),    # E
                    "tc_f":    _f(row[8]),    # I
                    "pc_psia": _f(row[9]),    # J
                    "omega":   _f(row[12]),   # M
                }
        wb.close()
    except Exception:
        pass

    for k, v in _pure_library().items():      # fallback: HMB entries win
        out.setdefault(k, v)
    for k, v in _pr_library().items():        # embedded 98-component PR databank (last resort)
        out.setdefault(k, v)
    _CONST_CACHE[key] = out
    return out


def _pr_library() -> dict:
    """{NAME_UPPER: constants} for every embedded PR-databank component, under its key,
    its databank name and every alias — used when a name is in neither the HMB nor the
    36-species pure library.  Tc/Pc/ω/MW/Vc/NBP/standard liquid density from the databank."""
    out: dict = {}
    for k, d in _PR_DATA.items():
        nm0, tb_k, sld_kg = _PR_EXTRA[k]
        rec = {"name": k, "mw": d[0], "tc_f": d[1] * 1.8 - 459.67,
               "pc_psia": d[2] * 14.5038, "omega": d[3],
               "vc": d[4] * 0.0160185,                    # cm3/mol → ft3/lbmol
               "nbp_f": tb_k * 1.8 - 459.67, "sld": sld_kg * 0.0624280}
        for nm in (k, nm0.upper()):
            out.setdefault(nm, rec)
    for a, k in _PR_ALIAS.items():
        out.setdefault(a, out[k])
    return out


def _read_comp_constants(hmb_path) -> dict:
    """HMB COMP_CONSTANTS + built-in pure library, plus any components the user
    defined on the Pseudo_Components sheet (those win)."""
    base = _read_comp_constants_hmb(hmb_path)
    if not _USER_COMPONENTS:
        return base
    merged = dict(base)
    merged.update(_USER_COMPONENTS)
    return merged


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

    # A single-component stream only gets a feed when it is pure VAPOUR (so it
    # can join/mix); a pure liquid would flash unphysically from the proxy K.
    if len(z) < 1 or (len(z) == 1 and ((sp.liq_mass or 0.0) > 1e-9
                                       or not (sp.vap_mass or 0.0) > 0)):
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
    if len(z) == 1:       # pure vapour: Rachford-Rice is a step; keep K >> 1 at any P
        k_ref = [1e10]

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


def _station_phase_state(station, sp, gas_density_fn) -> dict:
    """Phase flows / densities / viscosities / surface tension AT THIS STATION.

    Uses the station's own flash & property set (``station.props``) so that
    stations upstream of a tee join, flashed or cooled streams, etc. are each
    evaluated with their own flows — not with one circuit-level stream.  Falls back
    to the circuit stream ``sp`` where a station has no property set.
    """
    pr = getattr(station, "props", None) or {}
    rho_g = pr.get("rho_vap")
    if rho_g is None:
        rho_g = gas_density_fn(station.p_in_psia) if gas_density_fn else None
    rho_l = pr.get("rho_liq")
    if rho_l is None and sp is not None:
        rho_l = sp.liq_density
    return dict(
        mg=(pr["m_vap"] if pr.get("m_vap") is not None else (sp.vap_mass or 0.0)),
        ml=(pr["m_liq"] if pr.get("m_liq") is not None else (sp.liq_mass or 0.0)),
        rho_g=rho_g or 0.0, rho_l=rho_l or 0.0,
        mu_g=(pr.get("visc_vap") if pr.get("visc_vap") is not None
              else (sp.vap_visc if sp else None)),
        mu_l=(pr.get("visc_liq") if pr.get("visc_liq") is not None
              else (sp.liq_visc if sp else None)),
        sigma=(pr.get("sigma") if pr.get("sigma") is not None
               else (sp.liq_surf_tens if sp else None)))


def params_from_station(sp, station, gas_density_fn):
    """FPParams for one pipe station, using LOCAL (pressure-dependent) gas ρ."""
    ph = _station_phase_state(station, sp, gas_density_fn)
    rg = ph["rho_g"] * LBFT3_TO_KGM3
    rl = ph["rho_l"] * LBFT3_TO_KGM3
    mg = ph["mg"] * LBHR_TO_KGH
    ml = ph["ml"] * LBHR_TO_KGH
    sigma = (ph["sigma"] or 0.0) * DYNCM_TO_NM
    d_mm = (station.id_in or 0.0) * IN_TO_MM
    return compute_params(mg, rg, ph["mu_g"] or 0.0, ml, rl,
                          ph["mu_l"] or 0.0, sigma, d_mm)


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
           "Vsl", "Liquid apparent velocity (m/s)", 0.01, 10,
           "Vsg", "Gas apparent velocity (m/s)", 0.01, 10, "HEDH_DN_V"),
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


def _logsafe(mp, pt):
    """Floor non-positive coordinates on log axes (they cannot be plotted)."""
    x, y = pt
    if x is not None and mp.logx and x <= 0:
        x = mp.xmin
    if y is not None and mp.logy and y <= 0:
        y = mp.ymin
    return (x, y)


def _add_scatter(ws, mp, data_col, op_points, ctrl_point, regime_text=None):
    """Build one ScatterChart for map mp; data laid out from column data_col."""
    chart = ScatterChart()
    chart.title = mp.title + (f"   —   REGIME: {regime_text}" if regime_text else "")
    chart.height = 13
    chart.width = 22
    chart.legend = None          # regime is in the title/banner; a legend only squeezes the plot
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
        xc, yc, r0, r1 = _write_series_cols(
            ws, col, [_logsafe(mp, q) for q in poly], f"bnd{i}")
        xref = Reference(ws, min_col=xc, min_row=r0, max_row=r1)
        yref = Reference(ws, min_col=yc, min_row=r0, max_row=r1)
        s = Series(yref, xref, title=("regime boundary" if i == 0 else None))
        s.marker = Marker(symbol="none")
        s.graphicalProperties.line.solidFill = "808080"
        s.graphicalProperties.line.width = 14000
        chart.series.append(s)
        col += 3

    # operating path (blue line + small markers)
    # operating path: a point that cannot be plotted on a log axis (e.g. gas Froude = 0
    # at a liquid-only station) is DROPPED, not floored, so it cannot draw a false line
    op_ok = [q for q in op_points
             if q[0] is not None and q[1] is not None
             and (q[0] > 0 or not mp.logx) and (q[1] > 0 or not mp.logy)]
    xc, yc, r0, r1 = _write_series_cols(ws, col, op_ok, "path")
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
        xc, yc, r0, r1 = _write_series_cols(
            ws, col, [_logsafe(mp, ctrl_point)], "ctrl")
        xref = Reference(ws, min_col=xc, min_row=r0, max_row=r1)
        yref = Reference(ws, min_col=yc, min_row=r0, max_row=r1)
        s = Series(yref, xref, title="controlling (max ρv²)")
        s.marker = Marker(symbol="diamond", size=11)
        s.graphicalProperties.line.noFill = True
        chart.series.append(s)
        col += 3
    return chart, col


# ── Regime on every map ──────────────────────────────────────────────────
def _loglog_y(poly, x):
    """y on a polyline at x, interpolated linearly in log-log space (None if outside)."""
    for (x0, y0), (x1, y1) in zip(poly, poly[1:]):
        lo, hi = min(x0, x1), max(x0, x1)
        if lo <= x <= hi and hi > lo:
            t = (math.log10(x) - math.log10(x0)) / (math.log10(x1) - math.log10(x0))
            return 10 ** (math.log10(y0) + t * (math.log10(y1) - math.log10(y0)))
    return None


def _shell2_h_regime(frg, frl) -> str:
    """Shell DEP 31.22.05.11-Gen (horizontal): region of (gas Froude, liquid Froude),
    from the digitised boundary polylines of the DEP2-horizontal map."""
    P0, P1, P2, P3, P4, P5, P6, P7, P8 = BOUNDARIES["SHELL2_H"]
    x, y = frg, frl
    if not (x and y and x > 0 and y > 0):
        return "n/a"
    top = (_loglog_y(P5, x) if x <= 0.2 else _loglog_y(P6, x) if x <= 3.5
           else _loglog_y(P8, x))
    if top is not None and y >= top:
        return "Bubbly" if x <= 3.5 else "Bubbly / Intermittent (upper limit of map)"
    if x <= 0.3:
        low = _loglog_y(P1, x)
        if low is not None and y >= low:
            return "Intermittent (Slug-Plug)"
    else:
        up = _loglog_y(P7, x) if x <= 3.5 else _loglog_y(P8, x)
        if up is not None and y >= up:
            return "Intermittent (Slug-Plug)"
        lower = (_loglog_y(P2, x) if x <= 2.0 else _loglog_y(P3, x) if x <= 6.0
                 else _loglog_y(P4, x))
        if lower is not None and y >= lower:
            return "Annular"
        mist_line = _loglog_y(P4, x)
        if x >= 6.0 and mist_line is not None and y < mist_line:
            return "Mist flow"
        return "Stratified wavy flow"
    s0 = _loglog_y(P0, x)
    if x <= 0.08 and s0 is not None and y < s0:
        return "Stratified smooth flow"
    return "Stratified wavy flow"


def _station_regimes(st, sp, gdf, params) -> dict:
    """{method: regime} for one station (T&D, Taitel-Barnea, Baker, Hewitt-Roberts,
    Barnea, Beggs-Brill) plus Shell DEP2 (map polygons)."""
    out: dict = {}
    t = station_tp_inputs(st, sp, gdf)
    if t is None:
        return out
    sph = _is_single_phase(t)
    if sph:
        return {k: sph for k in ("Taitel-Dukler", "Taitel-Barnea", "Baker",
                                 "Hewitt-Roberts", "Barnea", "Beggs-Brill", "Shell DEP")}
    for name, fn in (("Taitel-Dukler", lambda: taitel_dukler(t)["regime"]),
                     ("Taitel-Barnea", lambda: taitel_barnea_vertical(t)),
                     ("Baker", lambda: baker(t)["regime"]),
                     ("Hewitt-Roberts", lambda: hewitt_roberts(t)["regime"]),
                     ("Barnea", lambda: barnea_unified(t)),
                     ("Beggs-Brill", lambda: shell_beggs_brill(t))):
        try:
            out[name] = fn()
        except Exception:
            out[name] = "n/a"
    out["Shell DEP"] = _shell2_h_regime(params.FrG, params.FrL)
    return out


# which classifier names each map (primary regime written on its charts)
_MAP_REGIME_METHOD = {
    "H Taitel-Dukler": "Taitel-Dukler", "H Gregory-Fogarasi": "Taitel-Dukler",
    "H Shell DEP": "Shell DEP", "H Shell DEP2": "Shell DEP", "H HTRI": "Taitel-Dukler",
    "H Baker (HEDH)": "Baker",
    "V Gregory-Fogarasi": "Taitel-Barnea", "V Shell DEP2": "Taitel-Barnea",
    "V Shell DEP": "Taitel-Barnea", "V HTRI": "Taitel-Barnea",
    "V HEDH up": "Hewitt-Roberts", "V Oshinowo up": "Taitel-Barnea",
    "V HEDH down": "Hewitt-Roberts", "V Oshinowo down": "Taitel-Barnea",
}


def build_map_sheets(wb, rows, ctrl_seq, regimes=None):
    """rows = [(station, FPParams)]; build one sheet per map-group with charts.
    ``regimes`` = {method: regime} at the controlling station — written on every
    chart title and in a banner above the charts."""
    regimes = regimes or {}
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
        ctrl_pr0 = next((pr for st, pr in rows if st.seq == ctrl_seq), None)
        primary = regimes.get(_MAP_REGIME_METHOD.get(sheet, "Taitel-Dukler"), "n/a")
        others = "   |   ".join(f"{k}: {v}" for k, v in regimes.items())
        ws.merge_cells("A1:R1")
        _cell(ws, 1, 1,
              f"OPERATING REGIME at controlling station Seq {ctrl_seq}:  {primary}   "
              f"({_MAP_REGIME_METHOD.get(sheet, 'Taitel-Dukler')})"
              + (f"   |   FrG = {ctrl_pr0.FrG:.4g}, FrL = {ctrl_pr0.FrL:.4g}"
                 if (ctrl_pr0 and ctrl_pr0.FrG and ctrl_pr0.FrL) else ""),
              bg=NAVY, fg=WHITE, sz=11, bold=True)
        _cell(ws, 2, 1, f"All methods:  {others}" if others else "", sz=9, bold=False)
        for j, mp in enumerate(maps):
            op_points = [(pr.get(mp.xkey), pr.get(mp.ykey)) for _, pr in rows]
            ctrl_pr = next((pr for st, pr in rows if st.seq == ctrl_seq), None)
            ctrl_point = (ctrl_pr.get(mp.xkey), ctrl_pr.get(mp.ykey)) if ctrl_pr else None
            chart, data_col = _add_scatter(ws, mp, data_col, op_points, ctrl_point,
                                           primary)
            anchor = f"A{4 + j * 23}"
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
    # per-station regimes next to the parameter table
    try:
        wsd = wb["Flow_Pattern_Data"]
        c0 = len(_TABLE_COLS) + 1
        # inputs behind the superficial velocities (FPS) so they can be checked by hand
        chk = ("Gas mass (lb/hr)", "Liq mass (lb/hr)", "Gas ρ (lb/ft3)", "Liq ρ (lb/ft3)",
               "Flow area (ft2)", "Vsg (ft/s)", "Vsl (ft/s)")
        c1 = c0 + 4
        for j, nm_ in enumerate(chk):
            _cell(wsd, 2, c1 + j, nm_, bg=GRNHDR, fg=WHITE, sz=9, bold=True,
                  ha="center", wrap=True)
            wsd.column_dimensions[get_column_letter(c1 + j)].width = 13
        for i_, (st_, pr_) in enumerate(rows):
            ph_ = _station_phase_state(st_, sp, gas_density_fn)
            area_ = math.pi / 4.0 * ((st_.id_in or 0.0) / 12.0) ** 2
            vals_ = [ph_["mg"], ph_["ml"], ph_["rho_g"], ph_["rho_l"], area_,
                     (pr_.Vsg / FT_TO_M) if pr_.Vsg is not None else None,
                     (pr_.Vsl / FT_TO_M) if pr_.Vsl is not None else None]
            for j, v_ in enumerate(vals_):
                _cell(wsd, 3 + i_, c1 + j, _fmt(v_), sz=9, ha="right")
        names_ = ("Regime (Taitel-Dukler)", "Regime (Shell DEP)", "Regime (Baker)",
                  "Regime (Hewitt-Roberts)")
        keys_ = ("Taitel-Dukler", "Shell DEP", "Baker", "Hewitt-Roberts")
        for j, nm_ in enumerate(names_):
            _cell(wsd, 2, c0 + j, nm_, bg=NAVY, fg=WHITE, sz=9, bold=True,
                  ha="center", wrap=True)
            wsd.column_dimensions[get_column_letter(c0 + j)].width = 26
        for i_, (st_, pr_) in enumerate(rows):
            rg_ = _station_regimes(st_, sp, gas_density_fn, pr_)
            for j, k_ in enumerate(keys_):
                _cell(wsd, 3 + i_, c0 + j, rg_.get(k_, ""), sz=9, ha="left")
    except Exception:
        pass
    ctrl_row = next((rp for rp in rows if rp[0].seq == ctrl_seq), rows[0])
    regimes = _station_regimes(ctrl_row[0], sp, gas_density_fn, ctrl_row[1])
    sheets = build_map_sheets(wb, rows, ctrl_seq, regimes)
    return ["Flow_Pattern_Data"] + sheets


if __name__ == "__main__":
    main()
