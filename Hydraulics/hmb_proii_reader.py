#!/usr/bin/env python3
"""Reader for PRO/II "Case N" transposed HMB exports.

A Case-N workbook has one worksheet per case ("Case 1", "Case 2", ...).  Each
case sheet is TRANSPOSED:

    col A = property label      col B = unit      col C..  = one stream each
    row 4 ("Stream Name")       carries the stream identifiers (133, 136, ...)

The rows above the first per-component block are scalar stream properties
(phase, rates, T, P, MW, phase densities/viscosities/Z/Cp, criticals ...);
below them are per-component sections (weight/molar rates & fractions, phase
mole fractions, ...), each an indented list of components with one value per
stream column.

This module implements exactly the interface ``hydraulics_XOM.py`` already
dispatches to (``HMBPROII``):

    list_cases(path)            -> ["Case 1", "Case 2", ...]
    list_streams(path, case)    -> [stream names in that case]
    get_stream(name, path, case)-> HMBStream(name, phase, case, props)

``streamprops_from_hmb(hs)`` consumes ``hs.props`` (a flat dict).  The
companion ``td_parser.parse_proii`` reuses ``parse_case`` here for the
per-component compositions used to build the flash feed.

Every result equals what the engine's per-stream path (``extract_stream`` /
``_from_td_dump``) reads from the expanded per-stream workbook — i.e. picking
a case shows the streams exactly as that per-stream workbook.
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field

from openpyxl import load_workbook


# ── scalar property label -> StreamProps prop key ──────────────────────────
# Matched against the stripped col-A label (case-insensitive, startswith).
# Order matters only where one label is a prefix of another (longest first).
_SCALAR_MAP: list[tuple[str, str]] = [
    ("stream phase",                    "phase"),
    ("total mass rate",                 "total_mass"),
    ("total molar rate",                "total_molar"),
    ("total std. liq.",                 "total_std_liq"),
    ("total std. vap.",                 "total_std_vap"),
    ("temperature",                     "temp_f"),
    ("pressure",                        "pres_psia"),
    ("total molecular",                 "mol_weight"),
    ("liquid weight fraction",          "liq_wt_frac"),
    ("total actual density",            "total_density"),
    ("total z (from actual density)",   "total_z"),
    ("true critical temperature",       "tc_f"),
    ("true critical pressure",          "pc_psia"),
    ("dry vapor molecular weight",      "vap_mw"),
    ("dry vapor act. density",          "vap_density"),
    ("dry vapor cp/cv ratio",           "vap_cp_cv"),
    ("dry vapor cp",                    "vap_cp"),
    ("dry vapor viscosity",             "vap_visc"),
    ("dry vapor z (from actual",        "vap_z"),
    ("dry vapor thermal conductivity",  "vap_therm_cond"),
    ("dry vapor sp. enthalpy",          "vap_sp_enthalpy"),
    ("dry liquid molecular weight",     "liq_mw"),
    ("dry liquid act. density",         "liq_density"),
    ("dry liquid cp",                   "liq_cp"),
    ("dry liquid viscosity",            "liq_visc"),
    ("dry liquid thermal conductivity", "liq_therm_cond"),
    ("dry liquid sp. enthalpy",         "liq_sp_enthalpy"),
    ("surface tension",                 "liq_surf_tens"),
    ("acentric",                        "acentric"),
]

# ── per-component sections we keep (header text -> composition key) ─────────
# Keys match what td_parser / _build_feed expect.
_COMP_SECTIONS: list[tuple[str, str]] = [
    ("total weight comp. rates",     "MASS_FLOW"),
    ("total molar comp. rates",      "MOLE_FLOW"),
    ("total molar comp. fractions",  "MOLE_FRAC"),      # z (total)
    ("liquid mole comp. fractions",  "LIQ_MOLE_FRAC"),  # x
    ("vapor mole comp. fractions",   "VAP_MOLE_FRAC"),  # y
]
# Any non-indented label containing all of these tokens is a section header
# (used to reset the "current section" so unwanted blocks are skipped).
_SECTION_TOKENS = ("comp.",)


@dataclass
class HMBStream:
    name: str
    case: str
    phase: str | None = None
    props: dict = field(default_factory=dict)


def _num(v):
    """Return float or None ('n/a', '', text -> None)."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s.lower() in ("n/a", "na", "-", "--"):
        return None
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return None


def _conv_density(val, unit):
    """-> LB/FT3.  HMB may print vapor density as 'LB/M FT3' (per 1000 ft3)."""
    if val is None:
        return None
    u = (unit or "").upper().replace(" ", "")
    return val / 1000.0 if ("MFT3" in u) else val


def _conv_massflow(val, unit):
    """-> LB/HR.  HMB may print mass flow as 'M LB/HR' (thousands)."""
    if val is None:
        return None
    return val * 1000.0 if "M LB" in (unit or "").upper() else val


# ── one-pass parse of a case sheet (cached per absolute path + mtime) ───────
_CACHE: dict[tuple, dict] = {}


def _case_sheet_name(wb, case) -> str | None:
    if case and case in wb.sheetnames:
        return case
    # tolerant: "Case 1", "case1", or the first sheet that looks like a case
    want = str(case or "").upper().replace(" ", "")
    for sn in wb.sheetnames:
        if sn.upper().replace(" ", "") == want:
            return sn
    for sn in wb.sheetnames:
        if sn.upper().startswith("CASE"):
            return sn
    return None


def list_cases(path) -> list[str]:
    wb = load_workbook(path, read_only=True, data_only=True)
    cases = [sn for sn in wb.sheetnames if sn.upper().startswith("CASE")]
    wb.close()
    return cases


def parse_case(path, case="Case 1") -> dict:
    """Parse one case sheet fully.  Returns a dict:
        {"streams": [names], "cols": {name: col_idx},
         "scalars": {name: {label_lower: (value, unit)}},
         "comp": {SECTION_KEY: {name: {component: value}}},
         "components": [component names in feed order]}
    Cached per (abspath, mtime, case)."""
    ap = os.path.abspath(path)
    key = (ap, os.path.getmtime(ap), str(case))
    if key in _CACHE:
        return _CACHE[key]

    wb = load_workbook(path, read_only=True, data_only=True)
    sheet = _case_sheet_name(wb, case)
    if sheet is None:
        wb.close()
        raise KeyError(f"no case sheet for {case!r} in {os.path.basename(path)}")
    ws = wb[sheet]

    cols: dict[str, int] = {}           # stream name -> 0-based column index
    scalars: dict[str, dict] = {}
    comp: dict[str, dict] = {k: {} for _, k in _COMP_SECTIONS}
    components: list[str] = []
    comp_seen: set[str] = set()

    section_lookup = {h: k for h, k in _COMP_SECTIONS}
    cur_key = None                      # current wanted comp section, else None
    in_scalar = True                    # until first comp section header

    for row in ws.iter_rows(values_only=True):
        a = row[0]
        if a is None:
            continue
        raw = str(a)
        label = raw.strip()
        if not label:
            continue
        indented = raw[:1].isspace()
        unit = row[1]

        # row 4: the stream-name header (also non-indented)
        if label.lower() == "stream name":
            for ci in range(2, len(row)):
                nm = row[ci]
                if nm is not None and str(nm).strip() != "":
                    cols[str(nm).strip()] = ci
            continue

        if not indented:
            low = label.lower()
            # is this a per-component SECTION header?
            is_section = any(tok in low for tok in _SECTION_TOKENS) and (
                "comp." in low)
            wanted = section_lookup.get(low)
            if wanted is not None:
                cur_key, in_scalar = wanted, False
                continue
            if is_section:
                cur_key, in_scalar = None, False   # a section we don't keep
                continue
            # otherwise a scalar property row (only meaningful before comps)
            if in_scalar:
                for nm, ci in cols.items():
                    if ci < len(row):
                        scalars.setdefault(nm, {})[low] = (row[ci], unit)
            continue

        # indented => a component row within the current section
        if cur_key is not None:
            comp_name = label
            if comp_name not in comp_seen:
                comp_seen.add(comp_name)
                components.append(comp_name)
            bucket = comp[cur_key]
            for nm, ci in cols.items():
                if ci < len(row):
                    v = _num(row[ci])
                    if v is not None:
                        bucket.setdefault(nm, {})[comp_name] = v

    wb.close()
    out = {"streams": list(cols.keys()), "cols": cols,
           "scalars": scalars, "comp": comp, "components": components,
           "sheet": sheet}
    _CACHE[key] = out
    return out


def list_streams(path, case="Case 1") -> list[str]:
    return parse_case(path, case)["streams"]


def get_stream(name, path, case="Case 1") -> HMBStream | None:
    """Return an HMBStream for `name` (props flat dict) or None if absent."""
    data = parse_case(path, case)
    sc = data["scalars"].get(str(name).strip())
    if sc is None:                       # case-insensitive fallback
        up = str(name).strip().upper()
        for nm, d in data["scalars"].items():
            if nm.upper() == up:
                sc, name = d, nm
                break
    if sc is None:
        return None

    def val(label_lower):
        e = sc.get(label_lower)
        return (e[0], e[1]) if e else (None, None)

    props: dict = {}
    phase = None
    for matcher, key in _SCALAR_MAP:
        # find the first scalar label that startswith the matcher
        hit = next(((v, u) for lbl, (v, u) in sc.items()
                    if lbl.startswith(matcher)), (None, None))
        v, u = hit
        if key == "phase":
            phase = None if v is None else str(v).strip()
            continue
        num = _num(v)
        if num is None:
            continue
        if key in ("vap_density", "liq_density", "total_density"):
            num = _conv_density(num, u)
        elif key == "total_mass":
            num = _conv_massflow(num, u)
        props[key] = num

    # phase mass split from the liquid weight fraction (robust for two-phase)
    tot = props.get("total_mass")
    lwf = props.pop("liq_wt_frac", None)
    if tot is not None and lwf is not None:
        props["liq_mass"] = tot * lwf
        props["vap_mass"] = tot * (1.0 - lwf)

    # ── fill gaps the transposed Case export leaves as 'n/a' ──────────────
    # This export omits Dry Vapor ACTUAL density & viscosity for every stream,
    # and the dry liquid density/viscosity for a few streams.  The per-stream
    # "TD Property Dump" carries them; here we reconstruct so the hydraulics
    # (which need a liquid density and both viscosities) still run.  Values are
    # marked estimated via props['_estimated'] for transparency.
    est: list[str] = []
    p, t = props.get("pres_psia"), props.get("temp_f")
    phz = (phase or "").upper()
    liq_present = ("LIQUID" in phz) or ("MIXED" in phz) or ("WET" in phz)
    vap_present = ("VAPOR" in phz) or ("GAS" in phz) or ("MIXED" in phz)

    # vapor actual density: ideal-gas with Z  (rho = P*MW/(Z*R*T_R))
    if props.get("vap_density") is None and vap_present:
        mw = props.get("vap_mw") or props.get("mol_weight")
        z = props.get("vap_z") or props.get("total_z") or 1.0
        if mw and p and t is not None:
            t_r = t + 459.67
            if z > 0 and t_r > 0:
                props["vap_density"] = p * mw / (z * 10.7316 * t_r)
                est.append("vap_density")
    # liquid density: fall back to the flowing bulk (total actual) density
    if props.get("liq_density") is None and liq_present:
        if props.get("total_density") is not None:
            props["liq_density"] = props["total_density"]
            est.append("liq_density")
    # viscosities: conservative defaults only when genuinely absent
    if props.get("vap_visc") is None and vap_present:
        props["vap_visc"] = 0.012                    # cP, typical light gas/vapour
        est.append("vap_visc")
    if props.get("liq_visc") is None and liq_present:
        props["liq_visc"] = 0.5                       # cP, generic hydrocarbon liquid
        est.append("liq_visc")
    if est:
        props["_estimated"] = est

    return HMBStream(name=str(name).strip(), case=data["sheet"],
                     phase=phase, props=props)


# ── raw unit strings (for the app's units-auto-propagation feature) ────────
_KEY_TO_QTY = {
    "temp_f": "T", "pres_psia": "P",
    "total_mass": "mflow", "total_molar": "molflow",
    "total_density": "rho", "vap_density": "rho", "liq_density": "rho",
    "vap_visc": "visc", "liq_visc": "visc",
    "liq_surf_tens": "st",
    "vap_mw": "MW", "liq_mw": "MW", "mol_weight": "MW",
    "total_std_liq": "qvol", "total_std_vap": "qvol",
    "vap_sp_enthalpy": "h", "liq_sp_enthalpy": "h",
}


def extract_units(path, case="Case 1") -> dict[str, str]:
    """Raw {quantity_code: unit_string} read from a Case-N HMB export's
    per-stream scalar unit column (col B). Picks the FIRST stream in the
    case — unit convention is per-case/workbook, not per-stream. Reuses
    _SCALAR_MAP's existing label matchers; does NOT normalize spelling
    (caller does). Unmapped/unknown labels are skipped."""
    data = parse_case(path, case)
    if not data["streams"]:
        return {}
    sc = data["scalars"].get(data["streams"][0], {})
    out: dict[str, str] = {}
    for matcher, key in _SCALAR_MAP:
        qty = _KEY_TO_QTY.get(key)
        if qty is None or qty in out:
            continue
        hit = next((u for lbl, (v, u) in sc.items()
                    if lbl.startswith(matcher) and u), None)
        if hit:
            out[qty] = str(hit).strip()
    return out
