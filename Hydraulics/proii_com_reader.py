#!/usr/bin/env python3
"""Live PRO/II connector.

Two independent halves:

  * Stream/component **listing** is pure Python — no COM, no PRO/II license,
    no Windows required.  Every ``.prz`` archive is a zip that embeds a
    plain-text keyword (``.inp``) backup with a ``STREAM DATA`` section
    naming every stream and a ``COMPONENT DATA`` section giving the
    component index<->name table PRO/II's composition attribute is indexed
    against.  Verified against two real ``.prz`` files during development
    (14 streams / 78 components, and 40 streams / 1 component) — this
    listing can be stale relative to the live database (it reflects the
    flowsheet structure as of the last keyword-export/save), so treat it as
    a structure hint, never as authoritative live data.

  * **Live values** require PRO/II installed and licensed on this machine
    (Windows only), read via PRO/II's ``SimSciDbs.Database.<NNN>`` COM
    automation server.  Unlike a HYSYS attach, this is a *headless* COM
    server — PRO/II's own GUI need not be open — but ``RunCalcs()`` re-
    solves the whole flowsheet and consumes a real license seat, so it is
    neither instant nor side-effect-free.  The exact ProgID/version and
    attribute names below are grounded in one working public example
    (github.com/bryanpiguave/Air-Separation) and are provisional — this
    development sandbox has no PRO/II installed to verify against, so
    expect to iterate against real COM errors on a real Windows+PRO/II
    machine.

Implements the same interface ``hydraulics_XOM.py`` already dispatches to
for the Excel-based ``hmb_proii_reader`` (``HMBStream(name, case, phase,
props)``), extended with per-component mole-fraction composition so a
``FlashFeed`` can be built directly from one live read without a second
pass over an Excel sheet:

    list_cases(path)             -> ["Default"]
    list_streams(path, case)     -> [stream names]
    get_stream(name, path, case) -> HMBStream(name, case, phase, props,
                                               composition)
"""
from __future__ import annotations

import os
import re
import zipfile
from dataclasses import dataclass, field

try:
    import win32com.client as _win32       # optional — Windows + pywin32
except ImportError:                          # pragma: no cover
    _win32 = None

# ProgID version suffixes to probe, in order.  Only ".102"/".110" are
# confirmed from a working example; the rest are speculative fallbacks for
# other installed PRO/II releases — refine this list once tested against
# real installs.
_PROGID_VERSIONS = ("110", "102", "100", "95", "91", "90")

_RESERVED_STREAM_NAMES = {"ALL", "NONE"}


@dataclass
class HMBStream:
    name: str
    case: str
    phase: str | None = None
    props: dict = field(default_factory=dict)
    composition: dict = field(default_factory=dict)   # total mole fraction (z)


class ProiiComError(RuntimeError):
    """Raised when a live PRO/II COM connection cannot be established."""


# ════════════════════════════════════════════════════════════════════════
#  Pure-Python listing — parse the .prz's own embedded keyword backup
# ════════════════════════════════════════════════════════════════════════
_CACHE: dict[tuple, dict] = {}


def _find_inp_member(zf: zipfile.ZipFile) -> str | None:
    for n in zf.namelist():
        if n.lower().endswith(".inp"):
            return n
    return None


def _join_continuations(text: str) -> list[str]:
    """PRO/II keyword lines ending in '&' continue on the next line."""
    out: list[str] = []
    buf = ""
    for raw in text.splitlines():
        s = raw.rstrip()
        if s.endswith("&"):
            buf += s[:-1] + " "
        else:
            buf += s
            out.append(buf)
            buf = ""
    if buf:
        out.append(buf)
    return out


def _parse_inp(text: str) -> dict:
    """Extract ``{"streams": [...], "components": {index: name}}`` from a
    PRO/II keyword (.inp) file's text, tolerant of continuation lines.

    Three keyword forms carry stream names and are handled distinctly — a
    single blanket ``STREAM=`` regex both false-positives on ``PRINT``'s
    ``STREAM=ALL`` control keyword and misses names that only appear in
    comma-separated ``OUTPUT``-style lists (confirmed against real sample
    files):
        PROPERTY STREAM=<name>, ...
        ... REFSTREAM=<name>, ...
        OUTPUT FORMAT=..., STREAMS=<name1>,<name2>,..., ...
    """
    lines = _join_continuations(text)
    streams: set[str] = set()
    components: dict[int, str] = {}

    for ln in lines:
        for m in re.finditer(r'\bPROPERTY\s+STREAM\s*=\s*([A-Za-z0-9_\-]+)', ln, re.I):
            streams.add(m.group(1))
        for m in re.finditer(r'\bREFSTREAM\s*=\s*([A-Za-z0-9_\-]+)', ln, re.I):
            streams.add(m.group(1))
        m2 = re.search(r'\bSTREAMS\s*=\s*([A-Za-z0-9_,\-]+)', ln, re.I)
        if m2:
            for nm in m2.group(1).split(","):
                nm = nm.strip()
                if nm and nm.upper() not in _RESERVED_STREAM_NAMES:
                    streams.add(nm)
        if "LIBID" in ln.upper():
            tail = ln.split("LIBID", 1)[-1]
            for m in re.finditer(r'(\d+)\s*,\s*([A-Za-z0-9]+)', tail):
                components[int(m.group(1))] = m.group(2)

    return {"streams": sorted(streams), "components": components}


def _load_inp_data(prz_path) -> dict:
    ap = os.path.abspath(prz_path)
    key = (ap, os.path.getmtime(ap))
    if key in _CACHE:
        return _CACHE[key]
    with zipfile.ZipFile(prz_path) as zf:
        inp_name = _find_inp_member(zf)
        if inp_name is None:
            raise ProiiComError(
                f"{os.path.basename(prz_path)} has no embedded .inp keyword "
                "file backup — cannot list streams without one")
        text = zf.read(inp_name).decode("latin1", errors="replace")
    data = _parse_inp(text)
    _CACHE[key] = data
    return data


def list_cases(prz_path) -> list[str]:
    """A .prz is one database, not multiple cases — one pseudo-case name
    satisfies the app's existing case-combo contract with no UI change."""
    return ["Default"]


def list_streams(prz_path, case: str = "Default") -> list[str]:
    return _load_inp_data(prz_path)["streams"]


# ════════════════════════════════════════════════════════════════════════
#  Live COM read (Windows + licensed PRO/II required)
# ════════════════════════════════════════════════════════════════════════
_SESSION_CACHE: dict[tuple, object] = {}


def _connect(prz_path):
    """Open (or reuse) a solved PRO/II COM database session for prz_path."""
    if _win32 is None:
        raise ProiiComError(
            "pywin32 is not installed — a live PRO/II connection requires "
            "Windows with pywin32 and a licensed PRO/II install")
    ap = os.path.abspath(prz_path)
    key = (ap, os.path.getmtime(ap))
    if key in _SESSION_CACHE:
        return _SESSION_CACHE[key]

    last_err = None
    pro2 = None
    for ver in _PROGID_VERSIONS:
        try:
            pro2 = _win32.Dispatch(f"SimSciDbs.Database.{ver}")
            break
        except Exception as exc:                        # pragma: no cover
            last_err = exc
            continue
    if pro2 is None:
        raise ProiiComError(
            "could not dispatch a PRO/II COM server (tried SimSciDbs."
            f"Database.{{{','.join(_PROGID_VERSIONS)}}}) — is PRO/II "
            f"installed and licensed on this machine? ({last_err})")

    try:
        db = pro2.OpenDatabase(ap)
        db.RunCalcs()
    except Exception as exc:
        raise ProiiComError(
            f"PRO/II COM connection to {os.path.basename(prz_path)} failed "
            f"during open/solve — check for a free license seat and that "
            f"the file isn't already open elsewhere: {exc}") from exc

    _SESSION_CACHE[key] = db
    return db


# Attribute names are provisional — pinned down against the real PRO/II COM
# object model once tested on a Windows machine with PRO/II installed; this
# sandbox has no PRO/II to verify against.
_SCALAR_ATTRS = [
    ("TemperatureCalc", "temp_f"),
    ("PressureCalc", "pres_psia"),
    ("TotalMassRate", "total_mass"),
    ("TotalMolarRate", "total_molar"),
    ("MolecularWeight", "mol_weight"),
    ("VapourMoleFraction", "_vap_mole_frac"),
]


def get_stream(name, path, case: str = "Default") -> HMBStream | None:
    """Return an HMBStream for `name`, or raise ProiiComError on failure
    (never silently None — a live connect failure should surface to the
    user, unlike a missing row in a static Excel export)."""
    inp_data = _load_inp_data(path)
    comp_index = inp_data["components"]

    db = _connect(path)
    try:
        strm = db.ActivateObject("Stream", str(name))
    except Exception as exc:
        raise ProiiComError(
            f"PRO/II stream {name!r} not found in "
            f"{os.path.basename(path)}: {exc}") from exc

    props: dict = {}
    for attr, key in _SCALAR_ATTRS:
        try:
            v = strm.GetAttribute(attr)
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
    for idx, comp_name in comp_index.items():
        try:
            v = strm.GetAttribute("TotalComposition", idx)
        except Exception:
            continue
        if v is not None:
            composition[comp_name] = float(v)

    return HMBStream(name=str(name), case=case, phase=phase, props=props,
                     composition=composition)
