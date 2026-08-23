#!/usr/bin/env python3
"""Live Aspen HYSYS connector via COM automation.

**Attaches to an already-running HYSYS instance only** — never launches
HYSYS or opens a case itself.  This is a deliberate product decision (not a
technical limitation — the underlying COM API does support launching a new
instance and opening a file, as the companion Excel tool this module was
reverse-engineered from optionally does too): attaching only is fast,
predictable, and never surprises the user with a hidden HYSYS process
consuming a license seat.  If the case isn't already open, this module
raises a clear error asking the user to open it first.

Reverse-engineered from the strings embedded in a companion "HSR" (HYSYS
Stream Reporter) Excel/VBA tool's ``xl/vbaProject.bin`` — readable via
``strings`` even without decompressing the MS-OVBA-compressed source.
Confirmed patterns: ``GetObject(,"HYSYS.Application")``,
``hyApp.SimulationCases``, a recursive flowsheet walk (the macro's
``FindFlowSheetsRecurs`` / ``AddAllFlowSheetsRecurs``) to reach
``Flowsheet.MaterialStreams``, and dozens of named stream properties
(``MassFlow``, ``MoleFlow``, ``VapourFraction``, ``ComponentMassFraction``,
density/viscosity/Cp/...).  The exact property-access calls below are
provisional, following the standard HYSYS "Basis Variable"
``<Stream>.<PropertyName>.GetValue(<unit>)`` pattern the macro strings
imply — this development sandbox has no HYSYS installed to verify against,
so expect to iterate against real COM errors on a real Windows+HYSYS
machine.

Implements the same interface ``hydraulics_XOM.py`` already dispatches to
for the Excel-based ``hmb_proii_reader`` (``HMBStream(name, case, phase,
props)``), extended with per-component mole-fraction composition so a
``FlashFeed`` can be built directly from one live read:

    list_cases(hsc_path)          -> [name(s) of the matching/open case(s)]
    list_streams(hsc_path, case)  -> [stream names, Main + sub-flowsheets]
    get_stream(name, path, case)  -> HMBStream(name, case, phase, props,
                                                composition)
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field

try:
    import win32com.client as _win32       # optional — Windows + pywin32
except ImportError:                          # pragma: no cover
    _win32 = None


@dataclass
class HMBStream:
    name: str
    case: str
    phase: str | None = None
    props: dict = field(default_factory=dict)
    composition: dict = field(default_factory=dict)   # mole fraction


class HysysComError(RuntimeError):
    """Raised when a live HYSYS COM attach cannot be established."""


def _attach():
    if _win32 is None:
        raise HysysComError(
            "pywin32 is not installed — a live HYSYS connection requires "
            "Windows with pywin32 and a running Aspen HYSYS instance")
    try:
        return _win32.GetObject(None, "HYSYS.Application")
    except Exception as exc:
        raise HysysComError(
            "could not attach to a running HYSYS instance — open the case "
            f"in HYSYS first, then try again ({exc})") from exc


def _find_case(hy_app, hsc_path):
    """Match an open SimulationCase to hsc_path by filename (case-
    insensitive, ignoring directory — HYSYS's own FullName reports the path
    it was opened from). Returns (matched_case_or_None, [all open names])."""
    want = os.path.basename(str(hsc_path)).upper()
    matched = None
    all_names: list[str] = []
    for case in hy_app.SimulationCases:
        try:
            nm = case.Name
            full = getattr(case, "FullName", "") or ""
        except Exception:
            continue
        all_names.append(nm)
        if os.path.basename(str(full)).upper() == want:
            matched = case
    return matched, all_names


def list_cases(hsc_path) -> list[str]:
    hy_app = _attach()
    matched, all_names = _find_case(hy_app, hsc_path)
    if matched is not None:
        return [matched.Name]
    if all_names:
        # None matched this exact file, but HYSYS is running with other
        # cases open — let the user pick, mirroring the reference macro's
        # case listbox rather than failing outright.
        return all_names
    raise HysysComError(
        f"HYSYS is running but no case matching "
        f"{os.path.basename(str(hsc_path))!r} is open — open it in HYSYS "
        "first, then try again")


def _resolve_case(hy_app, hsc_path, case):
    matched, all_names = _find_case(hy_app, hsc_path)
    if matched is not None and (not case or matched.Name == case):
        return matched
    for c in hy_app.SimulationCases:
        if c.Name == case:
            return c
    if matched is not None:
        return matched
    raise HysysComError(
        f"no open HYSYS case named {case!r} found (open cases: {all_names})")


def _iter_material_streams(flowsheet):
    """Recurse Flowsheet -> sub-Flowsheets, yielding (name, stream) for
    every MaterialStream — mirrors the HSR macro's FindFlowSheetsRecurs /
    AddAllFlowSheetsRecurs walk so streams inside columns/sub-flowsheets are
    found too, not just the top-level Main flowsheet."""
    try:
        for strm in flowsheet.MaterialStreams:
            yield (strm.Name, strm)
    except Exception:
        pass
    try:
        sub = flowsheet.Flowsheets
    except Exception:
        sub = None
    if sub is not None:
        try:
            count = sub.Count
        except Exception:
            count = 0
        for i in range(count):
            try:
                fs = sub.Item(i)
                yield from _iter_material_streams(fs.Flowsheet)
            except Exception:
                continue


def list_streams(hsc_path, case: str = "") -> list[str]:
    hy_app = _attach()
    sim_case = _resolve_case(hy_app, hsc_path, case)
    return [name for name, _ in _iter_material_streams(sim_case.Flowsheet)]


# Property access is provisional — see module docstring.  Each entry reads
# as strm.<PropertyName>.GetValue(<unit string>) (HYSYS "Basis Variable"
# pattern), except VapourFraction/MolecularWeight which are unitless.
_SCALAR_PROPS = [
    ("Temperature", "F", "temp_f"),
    ("Pressure", "psia", "pres_psia"),
    ("MassFlow", "lb/hr", "total_mass"),
    ("MolarFlow", "lbmole/hr", "total_molar"),
    ("VapourFraction", None, "_vap_mole_frac"),
    ("MolecularWeight", None, "mol_weight"),
    ("MassDensity", "lb/ft3", "total_density"),
    ("Viscosity", "cP", "vap_visc"),
]


def get_stream(name, path, case: str = "") -> HMBStream | None:
    """Return an HMBStream for `name`, or raise HysysComError on failure
    (never silently None — a live connect/read failure should surface to
    the user, unlike a missing row in a static Excel export)."""
    hy_app = _attach()
    sim_case = _resolve_case(hy_app, path, case)
    target = None
    for nm, strm in _iter_material_streams(sim_case.Flowsheet):
        if nm.upper() == str(name).strip().upper():
            target = strm
            break
    if target is None:
        raise HysysComError(
            f"stream {name!r} not found in case {sim_case.Name!r}")

    props: dict = {}
    for prop_name, unit, key in _SCALAR_PROPS:
        try:
            var = getattr(target, prop_name)
            v = var.GetValue(unit) if unit else var.Value
        except Exception:
            continue
        if v is not None:
            props[key] = float(v)

    vap_frac = props.pop("_vap_mole_frac", None)
    phase = None
    if vap_frac is not None:
        phase = ("Vapor" if vap_frac > 0.999 else
                 "Liquid" if vap_frac < 0.001 else "Mixed")

    composition: dict[str, float] = {}
    try:
        names = list(target.ComponentMoleFractionValue.Names)
        fracs = list(target.ComponentMoleFractionValue.GetValues())
        composition = {n: float(f) for n, f in zip(names, fracs)
                       if f is not None}
    except Exception:
        pass

    return HMBStream(name=str(name), case=sim_case.Name, phase=phase,
                     props=props, composition=composition)
