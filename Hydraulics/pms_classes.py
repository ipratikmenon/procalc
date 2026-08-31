"""
Project piping-material-class (PMS / GEMS) catalogue.

Classes come from two sources, merged into CLASSES:
  1. gems_extracted.json - the full catalogue bulk-extracted from the GEMS PDF by
     gems_pdf_extractor.py (~250 classes). Regenerate it whenever the PDF reissues.
  2. Curated overrides below (G1S-1, G1S-2) - hand-verified against the source
     sheets. Curated definitions WIN over the extracted copy of the same class.

Class names must match the PIPING-SPEC values in the PCF / ISOGEN files (see
pcf_parser.py) so pms_parser.py can reconcile drawings against the spec.

To add or correct a class by hand, append a PipingClass to CURATED.
"""

from __future__ import annotations

import json
import os
import shutil
import sys
from dataclasses import asdict, dataclass, field, replace

_HERE = os.path.dirname(os.path.abspath(__file__))
_JSON = os.path.join(_HERE, "gems_extracted.json")     # bundled catalogue (seed)


def user_json_path() -> str:
    """Writable per-user PMS catalogue location.

    The app edits / uploads the piping catalogue here so a project change never
    touches the bundled seed.  Windows: %LOCALAPPDATA%/Procalc/pms.json;
    macOS: ~/Library/Application Support/Procalc/pms.json; otherwise (Linux)
    ~/.local/share/Procalc/pms.json.  An env override (PROCALC_PMS_JSON) wins
    so tests / the app can point it anywhere.
    """
    env = os.environ.get("PROCALC_PMS_JSON")
    if env:
        return env
    if sys.platform == "darwin":
        base = os.path.join(os.path.expanduser("~"), "Library", "Application Support")
    else:
        base = (os.environ.get("LOCALAPPDATA")
                or os.path.join(os.path.expanduser("~"), ".local", "share"))
    return os.path.join(base, "Procalc", "pms.json")


def _active_json_path() -> str:
    """The catalogue actually loaded: the user file if present, else the seed.

    On first use the bundled seed is copied to the user path so the user always
    has an editable copy (best-effort; falls back to the seed if the copy
    fails, e.g. a read-only home)."""
    up = user_json_path()
    if os.path.exists(up):
        return up
    try:
        os.makedirs(os.path.dirname(up), exist_ok=True)
        if os.path.exists(_JSON):
            shutil.copyfile(_JSON, up)
            return up
    except OSError:
        pass
    return _JSON


@dataclass(frozen=True)
class PipeRule:
    """One GEMS Pipe-table row: a size range mapped to a schedule.

    schedule = a B36.10 designation ("80S", "40S", "STD", ...) when fixed, or
    "CAL" when the sheet says 'Calculate Thickness' (verify by ASME B31.3).
    """
    nps_low: str
    nps_high: str
    schedule: str
    ends: str
    description: str
    note: str = ""
    min_schedule: str = "STD"   # CAL floor: never thinner than this wall.
    # Explicit wall thickness (in) for this size range, only meaningful when
    # schedule == "CAL" -- takes precedence over the min_schedule table-
    # lookup floor when present. None (the default) reproduces today's
    # behavior exactly (min_schedule governs).
    special_thickness_in: float | None = None


@dataclass(frozen=True)
class PTPoint:
    """One point of a class's Temperature/Pressure derating curve, kept in
    both unit systems as entered (not cross-derived) since the source sheet
    hand-enters both independently."""
    temp_c: float | None = None
    pressure_kgcm2g: float | None = None
    temp_f: float | None = None
    pressure_psig: float | None = None


@dataclass(frozen=True)
class ComponentRow:
    """A generic GEMS component-table row (valves, flanges, fittings, etc.)."""
    item: str
    nps_low: str
    nps_high: str
    rating: str
    ends: str
    description: str
    pd_code: str = ""
    note: str = ""


@dataclass(frozen=True)
class PipingClass:
    name: str
    moc: str
    moc_tag: str
    material_group: str
    typical_service: str
    nominal_rating: str
    pt_limit_text: str
    design_pressure_psig: float
    design_temp_f: float
    corrosion_allow_in: float
    corrosion_allow_min_in: float
    nace_moc: str = ""
    special_moc: str = ""
    rev: str = ""
    gems_ref: str = ""
    pipe_rules: tuple[PipeRule, ...] = field(default_factory=tuple)
    components: tuple[ComponentRow, ...] = field(default_factory=tuple)
    # Up to ~10 Temperature/Pressure derating points (design_pressure_psig/
    # design_temp_f above stay the single "primary" point most code reads --
    # auto-derived from pt_curve[0] in _class_from_dict when a curve is
    # given, so nothing that only reads the scalars regresses).
    pt_curve: tuple[PTPoint, ...] = field(default_factory=tuple)


# ── curated, hand-verified classes (override the extracted copy) ────────────
_ALLOY20_PIPE_RULES = (
    PipeRule("3/4", "1-1/2", "80S", "SW", "Alloy 20 (UNS N08020), Seamless", "35"),
    PipeRule("2", "6", "40S", "BW", "Alloy 20 (UNS N08020), Seamless"),
    PipeRule("2", "6", "40S", "BW", "Alloy 20 (UNS N08020), Welded"),
    PipeRule("8", "18", "CAL", "BW", "Alloy 20 (UNS N08020), Calculate Thickness", min_schedule="STD"),
)

G1S_1 = PipingClass(
    name="G1S-1", moc="Alloy 20 (UNS N08020)", moc_tag="N08020", material_group="3.17",
    typical_service="Sulfuric acid 89% to 98% solution and acid emulsions.",
    nominal_rating="ASME B16.5 CLASS 150, M.G. 3.17",
    pt_limit_text="230 PSIG at -20 F to 100 F through 155 PSIG at 450 F",
    design_pressure_psig=230.0, design_temp_f=100.0,
    corrosion_allow_in=0.063, corrosion_allow_min_in=0.05,
    special_moc="Gaskets: Flexible Graphite Covered w/PTFE Inner Ring (GCT01).",
    rev="7", gems_ref="GEMS 3-23", pipe_rules=_ALLOY20_PIPE_RULES,
)

G1S_2 = PipingClass(
    name="G1S-2", moc="Alloy 20 (UNS N08020)", moc_tag="N08020", material_group="3.17",
    typical_service="Hydrocarbon systems subject to intergranular corrosion.",
    nominal_rating="ASME B16.5 CLASS 150, M.G. 3.17",
    pt_limit_text="230 PSIG at -20 F to 100 F through 155 PSIG at 450 F",
    design_pressure_psig=230.0, design_temp_f=100.0,
    corrosion_allow_in=0.063, corrosion_allow_min_in=0.05,
    special_moc="Gaskets: Flexible Graphite Filled, 1/8\" (3.2 mm) thick (GWG09).",
    rev="6", gems_ref="GEMS 3-23", pipe_rules=_ALLOY20_PIPE_RULES,
)

CURATED: list[PipingClass] = [G1S_1, G1S_2]


# ── pipe-rule patches ────────────────────────────────────────────────────
# Some extracted classes carry a single "CAL" row spanning a size range that
# the GEMS sheet actually splits across multiple footnotes with DIFFERENT
# schedule floors per sub-range; the bulk extractor has no way to tell them
# apart and flattens them to one row with min_schedule="STD". A patch here
# replaces ONLY pipe_rules for the named class — moc/components/etc. still
# come from gems_extracted.json.
#
# G1A-3: the bulk-extracted "26"-48" CAL" row collapsed footnotes 63 (26"
# only) and 1 (28"-48") into one min_schedule="STD" row. Per the project's
# pipe-class data sheet (A 671 Gr.CC60 CL.22, WELDED, size range 26"-36"):
#   26" only (note 63): P = 20.04 kgf/cm2g (285 psig) -> floor STD
#   28"-36"  (note 1):  P = 16.03 kgf/cm2g (228 psig) -> floor XS
# ASME B31.3 304.1.2(b): t = P*D / (2*(S*E*W + P*Y)), S = 20000 psig
# (M.G. 1.1 carbon steel), E = 1.0, Y = 0.4, W = 1.0; tm = t + 0.125 in
# corrosion allowance; required nominal T = tm + 0.3 mm (flat A671 welded-
# pipe mill tolerance, not the 12.5% seamless convention). Both computed T
# values land well under their schedule floor, so the floor governs:
#   26":     t=0.184 in, tm=0.309 in, T=0.321 in -> STD (0.375 in) wins
#   28"-36": t=0.159 in, tm=0.284 in, T=0.296 in -> XS  (0.500 in) wins
PIPE_RULE_PATCHES: dict[str, tuple[PipeRule, ...]] = {
    "G1A-3": (
        PipeRule("3/4", "1", "160", "–", "Carbon Steel, Seamless", "14"),
        PipeRule("1-1/2", "2", "80", "–", "Carbon Steel, Seamless"),
        PipeRule("3", "24", "Std", "–", "Carbon Steel, Seamless"),
        PipeRule("26", "26", "STD", "–",
                 "Carbon Steel, Welded (A671 Gr.CC60 CL.22), Calculate Thickness",
                 "63"),
        PipeRule("28", "36", "XS", "–",
                 "Carbon Steel, Welded (A671 Gr.CC60 CL.22), Calculate Thickness",
                 "1"),
    ),
}


# ── load extracted classes from JSON ────────────────────────────────────────
def _class_from_dict(d: dict) -> PipingClass:
    def _f(v):
        try:
            return float(v) if v is not None and v != "" else None
        except (TypeError, ValueError):
            return None

    pipe_rules = tuple(
        PipeRule(
            nps_low=r.get("nps_low", ""), nps_high=r.get("nps_high", ""),
            schedule=r.get("schedule", "STD"), ends=r.get("ends", ""),
            description=r.get("description", ""), note=r.get("note", ""),
            min_schedule=r.get("min_schedule", "STD"),
            special_thickness_in=_f(r.get("special_thickness_in")),
        )
        for r in d.get("pipe_rules", [])
        if r.get("nps_low") and r.get("nps_high")
    )
    components = tuple(
        ComponentRow(
            item=r.get("item") or r.get("section", ""),
            nps_low=r.get("nps_low", ""), nps_high=r.get("nps_high", ""),
            rating=r.get("rating", ""), ends=r.get("ends", ""),
            description=r.get("description", ""), pd_code=r.get("pd_code", ""),
            note=r.get("note", ""),
        )
        for r in d.get("components", [])
    )
    pt_curve = tuple(
        PTPoint(temp_c=_f(p.get("temp_c")), pressure_kgcm2g=_f(p.get("pressure_kgcm2g")),
               temp_f=_f(p.get("temp_f")), pressure_psig=_f(p.get("pressure_psig")))
        for p in d.get("pt_curve", [])
    )

    design_pressure_psig = _f(d.get("design_pressure_psig"))
    design_temp_f = _f(d.get("design_temp_f"))
    if pt_curve:
        # scalar "primary" point stays in sync with the curve when both are
        # given but disagree -- the curve is the richer, authoritative source.
        design_pressure_psig = pt_curve[0].pressure_psig or design_pressure_psig
        design_temp_f = pt_curve[0].temp_f or design_temp_f

    return PipingClass(
        name=d["name"], moc=d.get("moc", ""), moc_tag=d.get("moc_tag", ""),
        material_group=d.get("material_group", ""),
        typical_service=d.get("typical_service", ""),
        nominal_rating=d.get("nominal_rating", ""),
        pt_limit_text=d.get("pt_limit_text", ""),
        design_pressure_psig=design_pressure_psig or 0.0,
        design_temp_f=design_temp_f if design_temp_f is not None else 100.0,
        corrosion_allow_in=float(d.get("corrosion_allow_in", 0) or 0),
        corrosion_allow_min_in=float(d.get("corrosion_allow_min_in", 0) or 0),
        nace_moc=d.get("nace_moc", ""), special_moc=d.get("special_moc", ""),
        rev=d.get("rev", ""), gems_ref=d.get("gems_ref", ""),
        pipe_rules=pipe_rules, components=components, pt_curve=pt_curve,
    )


def load_classes(json_path: str = _JSON) -> list[PipingClass]:
    """Curated classes first, then every extracted class not already curated."""
    curated_names = {pc.name.upper() for pc in CURATED}
    out: list[PipingClass] = list(CURATED)
    if os.path.exists(json_path):
        with open(json_path, encoding="utf-8") as f:
            for d in json.load(f):
                name = d.get("name", "")
                if name.upper() in curated_names:
                    continue
                try:
                    pc = _class_from_dict(d)
                    patch = PIPE_RULE_PATCHES.get(name.upper())
                    if patch:
                        pc = replace(pc, pipe_rules=patch)
                    out.append(pc)
                except Exception as e:  # noqa: BLE001
                    print(f"  ! skipped class {d.get('name')}: {e}")
    return out


def _to_jsonable(pc: PipingClass) -> dict:
    """PipingClass -> plain dict for gems_extracted.json (round-trips through
    _class_from_dict).  pipe_rules / components / pt_curve become lists of
    dicts."""
    d = asdict(pc)
    d["pipe_rules"] = [dict(r) for r in d.get("pipe_rules", ())]
    d["components"] = [dict(c) for c in d.get("components", ())]
    d["pt_curve"] = [dict(p) for p in d.get("pt_curve", ())]
    return d


def save_user_classes(classes: list[PipingClass],
                      json_path: str | None = None) -> str:
    """Write the (non-curated) catalogue to the user pms.json.

    CURATED classes stay code-owned (authoritative overrides) and are not
    written — they are re-applied by load_classes on the next reload.  Returns
    the path written."""
    path = json_path or user_json_path()
    curated_names = {pc.name.upper() for pc in CURATED}
    payload = [_to_jsonable(pc) for pc in classes
               if pc.name.upper() not in curated_names]
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=1)
    return path


def install_catalogue(src_json: str, json_path: str | None = None) -> int:
    """Validate an uploaded gems_extracted.json and make it the active
    catalogue.  Parses every entry through _class_from_dict (raising on a bad
    file, so the app can reject it), then copies it to the user path and
    reloads.  Returns the number of classes loaded."""
    with open(src_json, encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, list):
        raise ValueError("PMS catalogue must be a JSON list of classes")
    bad = []
    for d in data:
        try:
            _class_from_dict(d)
        except Exception as e:  # noqa: BLE001
            bad.append(f"{d.get('name', '?')}: {e}")
    if bad:
        raise ValueError("invalid classes:\n  " + "\n  ".join(bad[:20]))
    path = json_path or user_json_path()
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    shutil.copyfile(src_json, path)
    reload(path)
    return len(CLASSES)


def reload(json_path: str | None = None) -> None:
    """Rebuild the module-level CLASSES in place from json_path (or the active
    user/seed path).  Engines holding their own name index must re-index (see
    hydraulics_XOM.reload_pms)."""
    global CLASSES
    CLASSES[:] = sorted(load_classes(json_path or _active_json_path()),
                        key=lambda pc: pc.name)


# The project class catalogue, sorted by name for stable output.
CLASSES: list[PipingClass] = sorted(load_classes(_active_json_path()),
                                    key=lambda pc: pc.name)
