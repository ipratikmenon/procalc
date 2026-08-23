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
props)``), extended with per-component mole-fraction composition (total,
and now also vapor/liquid phase-split) so a ``FlashFeed`` can be built
directly from one live read using the accurate true-K (y/x) basis:

    list_cases(hsc_path)          -> [name(s) of the matching/open case(s)]
    list_streams(hsc_path, case)  -> [stream names, Main + sub-flowsheets]
    get_stream(name, path, case)  -> HMBStream(name, case, phase, props,
                                                composition, composition_y,
                                                composition_x)

Property display names (``Temperature``, ``Mass Flow``, ``Molecular
Weight``, ``Component Molar Fraction``, ...) are confirmed against the
companion HSR tool's own ``PropSets``/``Settings``/``Setup`` worksheets — a
real, exhaustive ~200-row ``Parameter -> Unit Type`` lookup table plus a
real saved example run — not just VBA-string guesswork.  The phase-
qualified COM *call shape* itself (``<Name>Value.GetValue(unit, phase)``)
is still the standard documented HYSYS Automation pattern but remains
unverified without real HYSYS installed — this sandbox has none.
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
    composition: dict = field(default_factory=dict)     # total mole fraction (z)
    composition_y: dict = field(default_factory=dict)   # vapor mole fraction
    composition_x: dict = field(default_factory=dict)   # liquid mole fraction


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


# Property display names confirmed against the HSR tool's PropSets/Settings
# sheets (real vendor vocabulary — see module docstring); the phase-
# qualified call shape (<Name>Value.GetValue(unit, phase)) is provisional.
_OVERALL_PROPS = [
    ("Temperature", "F", "temp_f"),
    ("Pressure", "psia", "pres_psia"),
    ("Mass Flow", "lb/hr", "total_mass"),
    ("Molar Flow", "lbmole/hr", "total_molar"),
    ("Molecular Weight", None, "mol_weight"),
    ("Mass Density", "lb/ft3", "total_density"),
    ("Vapour Fraction", None, "_vap_mole_frac"),
]
_VAPOUR_PROPS = [
    ("Mass Flow", "lb/hr", "vap_mass"),
    ("Molecular Weight", None, "vap_mw"),
    ("Mass Density", "lb/ft3", "vap_density"),
    ("Viscosity", "cP", "vap_visc"),
    ("Thermal Conductivity", "btu/hr-ft-F", "vap_therm_cond"),
    ("Z Factor", None, "vap_z"),
]
_LIGHT_LIQUID_PROPS = [
    ("Mass Flow", "lb/hr", "liq_mass"),
    ("Molecular Weight", None, "liq_mw"),
    ("Mass Density", "lb/ft3", "liq_density"),
    ("Viscosity", "cP", "liq_visc"),
    ("Thermal Conductivity", "btu/hr-ft-F", "liq_therm_cond"),
    ("Surface Tension", "dyne/cm", "liq_surf_tens"),
]
# Heavy Liquid (aqueous phase) -- only its Mass Flow is folded additively
# into liq_mass (documented v1 simplification); its own intensive
# properties (density/viscosity/...) aren't separately tracked, matching
# the reference HSR tool's own "LiqProps" range-name merge of Light+Heavy
# Liquid under one property set (see Settings sheet).
_HEAVY_LIQUID_MASS_PROP = ("Mass Flow", "lb/hr")


def _read_prop(target, prop_name, unit, phase=None):
    """<Stream>.<PropertyName>Value.GetValue(<unit>[, <phase>]) -- provisional,
    see module docstring. The <PropertyName>Value attribute strips spaces
    (e.g. "Mass Flow" -> MassFlowValue)."""
    attr = prop_name.replace(" ", "") + "Value"
    var = getattr(target, attr)
    if phase is not None:
        return var.GetValue(unit or "", phase)
    return var.GetValue(unit) if unit else var.Value


def _read_phase_props(target, prop_list, phase=None):
    out: dict = {}
    for prop_name, unit, key in prop_list:
        try:
            v = _read_prop(target, prop_name, unit, phase)
        except Exception:
            continue
        if v is not None:
            out[key] = float(v)
    return out


def _read_composition(target, phase=None) -> dict[str, float]:
    """Component Molar Fraction, optionally phase-qualified (y/x) --
    provisional call shape, see module docstring."""
    try:
        var = target.ComponentMolarFractionValue
        names = list(var.Names)
        fracs = list(var.GetValues(phase) if phase is not None else var.GetValues())
        return {n: float(f) for n, f in zip(names, fracs) if f is not None}
    except Exception:
        return {}


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

    props = _read_phase_props(target, _OVERALL_PROPS)
    props.update(_read_phase_props(target, _VAPOUR_PROPS, "Vapour"))
    props.update(_read_phase_props(target, _LIGHT_LIQUID_PROPS, "Light Liquid"))

    heavy = _read_phase_props(target, [_HEAVY_LIQUID_MASS_PROP + ("_heavy_mass",)],
                              "Heavy Liquid")
    if heavy.get("_heavy_mass"):
        props["liq_mass"] = (props.get("liq_mass") or 0.0) + heavy["_heavy_mass"]

    vap_frac = props.pop("_vap_mole_frac", None)
    phase = None
    if vap_frac is not None:
        phase = ("Vapor" if vap_frac > 0.999 else
                 "Liquid" if vap_frac < 0.001 else "Mixed")

    composition = _read_composition(target)
    composition_y = _read_composition(target, "Vapour")
    composition_x = _read_composition(target, "Light Liquid")

    return HMBStream(name=str(name), case=sim_case.Name, phase=phase,
                     props=props, composition=composition,
                     composition_y=composition_y, composition_x=composition_x)
