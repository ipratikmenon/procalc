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
from dataclasses import dataclass, field

_HERE = os.path.dirname(os.path.abspath(__file__))
_JSON = os.path.join(_HERE, "gems_extracted.json")


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


# ── load extracted classes from JSON ────────────────────────────────────────
def _class_from_dict(d: dict) -> PipingClass:
    pipe_rules = tuple(
        PipeRule(
            nps_low=r.get("nps_low", ""), nps_high=r.get("nps_high", ""),
            schedule=r.get("schedule", "STD"), ends=r.get("ends", ""),
            description=r.get("description", ""), note=r.get("note", ""),
            min_schedule=r.get("min_schedule", "STD"),
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
    return PipingClass(
        name=d["name"], moc=d.get("moc", ""), moc_tag=d.get("moc_tag", ""),
        material_group=d.get("material_group", ""),
        typical_service=d.get("typical_service", ""),
        nominal_rating=d.get("nominal_rating", ""),
        pt_limit_text=d.get("pt_limit_text", ""),
        design_pressure_psig=float(d.get("design_pressure_psig", 0) or 0),
        design_temp_f=float(d.get("design_temp_f", 100) or 100),
        corrosion_allow_in=float(d.get("corrosion_allow_in", 0) or 0),
        corrosion_allow_min_in=float(d.get("corrosion_allow_min_in", 0) or 0),
        nace_moc=d.get("nace_moc", ""), special_moc=d.get("special_moc", ""),
        rev=d.get("rev", ""), gems_ref=d.get("gems_ref", ""),
        pipe_rules=pipe_rules, components=components,
    )


def load_classes(json_path: str = _JSON) -> list[PipingClass]:
    """Curated classes first, then every extracted class not already curated."""
    curated_names = {pc.name.upper() for pc in CURATED}
    out: list[PipingClass] = list(CURATED)
    if os.path.exists(json_path):
        with open(json_path, encoding="utf-8") as f:
            for d in json.load(f):
                if d.get("name", "").upper() in curated_names:
                    continue
                try:
                    out.append(_class_from_dict(d))
                except Exception as e:  # noqa: BLE001
                    print(f"  ! skipped class {d.get('name')}: {e}")
    return out


# The project class catalogue, sorted by name for stable output.
CLASSES: list[PipingClass] = sorted(load_classes(), key=lambda pc: pc.name)
