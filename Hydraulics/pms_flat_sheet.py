"""Bidirectional converter: the flat-sheet PMS master format (Excel) <->
the PipingClass/PipeRule/PTPoint schema used everywhere else in the app.

The flat sheet is the human-editable master: one row per (spec, item, NPS
range), grouped by spec, with one "CON" (condition) row per spec carrying
every spec-level scalar plus its Temperature/Pressure curve. See
FOLLOW-UP CHANGE 11 in the project plan for the full column layout this
was built against (row 3 = short field-key codes = the column names below).

Only Hydraulics/pms_classes.py is imported -- this stays a thin, testable
layer on top of the existing dataclasses, not a second source of truth.
"""
from __future__ import annotations

import re
from dataclasses import replace

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

import pms_classes as PMS

HEADER_ROW = 3          # field-key row (SPEC, ITEM, MOCS, ...)
DATA_START_ROW = 4
N_PT_POINTS = 10
PSI_PER_KGCM2 = 14.223343307120154
IN_PER_MM = 1.0 / 25.4

# Canonical NPS strings, smallest to largest -- must match exactly what
# Hydraulics/hydraulics_XOM.py's PIPE_OD_IN / expand_nps_range expect
# ("1/2", not ".5" -- see FOLLOW-UP CHANGE 11 postmortem: the HURL PDF
# extractor used ".5" and silently broke every small-bore CAL/schedule
# lookup for every class that used it).
_FRACTIONS = {0.25: "1/4", 0.5: "1/2", 0.75: "3/4", 1.25: "1-1/4",
             1.5: "1-1/2", 2.5: "2-1/2", 3.5: "3-1/2"}


def nps_to_string(value) -> str:
    """A DMIN/DMAX spreadsheet number -> the engine's canonical NPS string."""
    if value is None or value == "":
        return ""
    if isinstance(value, str):
        s = value.strip()
        # Real HURL sheets sometimes annotate a size with a parenthetical
        # note about it ("6 (Alt.)", "0.5 (Alt. 2)") -- strip the note (the
        # "alternate" distinction has no home in PipeRule's schema today,
        # a known simplification, not silently invented data) and parse
        # the leading token as the actual NPS value.
        s = re.split(r"\s*\(", s, 1)[0].strip()
        try:
            value = float(s)
        except ValueError:
            # not a bare number -- e.g. a hand-typed "1-1/2": pass through
            # as-is, trusting the sheet author over any further reformatting.
            return s
    f = float(value)
    if f in _FRACTIONS:
        return _FRACTIONS[f]
    if f == int(f):
        return str(int(f))
    return str(f)  # unexpected fractional value -- surface it, don't guess


def nps_to_float(nps: str) -> float | None:
    """Reverse of nps_to_string, for writing DMIN/DMAX back out as numbers."""
    if not nps:
        return None
    for f, s in _FRACTIONS.items():
        if s == nps:
            return f
    try:
        return float(nps)
    except ValueError:
        return None


# Bare tokens the engine's PIPE_WALL_IN/SCHEDULE_ALIASES actually recognize
# (Hydraulics/hydraulics_XOM.py:86-129). The source HURL sheets print
# "S-XS"/"S-STD"/"S-80" (a leading "S-" for "Schedule-"); the user's flat
# sheet compresses that to "SXS"/"S80"/"STD" (dash dropped). Only strip a
# LEADING "S" when what follows is one of these bare tokens -- must not
# touch genuine trailing-S stainless designations ("40S"/"80S"/"10S"),
# which are already correct as-is.
_BARE_SCHEDULES = {"XS", "STD", "XXS", "10", "40", "80", "120", "160"}


def normalize_schedule(value) -> str:
    """A flat-sheet SCHD value -> the engine's canonical schedule string."""
    s = str(value or "").strip().upper()
    if s == "CAL":
        return s
    if s.startswith("S") and s[1:] in _BARE_SCHEDULES:
        return s[1:]
    return s


# ── field-key layout (row 3) ────────────────────────────────────────────
_FIXED_FIELDS = ["SPEC", "ITEM", "MOCS", "MATG", "PLBS", "FACE", "PWHT",
                 "CAIN", "CAMM", "DMIN", "DMAX", "SCHD", "THIN", "THMM"]


def _pt_field_names():
    keys = []
    for i in range(1, N_PT_POINTS + 1):
        keys += [f"DT{i}S", f"DP{i}S", f"DT{i}F", f"DP{i}F"]
    return keys


ALL_FIELDS = _FIXED_FIELDS + _pt_field_names()


def _yn(v) -> bool:
    return str(v or "").strip().upper() == "YES"


_BLANK_PLACEHOLDERS = {"-", "—", "–", "N/A", "NA"}


def _is_blank(v) -> bool:
    """True for a genuinely empty cell OR a text placeholder real HURL
    sheets use for "not applicable" instead of leaving the cell empty
    (a bare "-"/"—", or a cached Excel formula-error literal like
    "#VALUE!"/"#N/A" from a since-broken source formula) -- checked
    WITHOUT coercing to float, so callers that still want the raw value
    (e.g. nps_to_string's hand-typed-fraction string pass-through) aren't
    forced through a numeric conversion."""
    if v is None:
        return True
    if isinstance(v, str):
        s = v.strip()
        if not s or s.upper() in _BLANK_PLACEHOLDERS or s.startswith("#"):
            return True
    return False


def _num(v) -> float | None:
    """A spreadsheet cell -> float, or None for a blank/placeholder cell
    (see _is_blank)."""
    if _is_blank(v):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    try:
        return float(s)
    except ValueError:
        if s.endswith(".") and s.count(".") == 2:
            # a stray trailing decimal-point typo in the source data
            # ("11.91." instead of "11.91") -- unambiguous to strip.
            return float(s[:-1])
        raise   # a genuinely malformed, non-placeholder value -- surface it


def flat_sheet_to_classes(xlsx_path: str, sheet_name: str = "PMS") -> list[PMS.PipingClass]:
    """Read the flat sheet -> list[PipingClass]. Raises ValueError with a
    row-numbered message list on anything malformed (never silently drops
    a spec)."""
    wb = load_workbook(xlsx_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"sheet '{sheet_name}' not found (have: {wb.sheetnames})")
    ws = wb[sheet_name]

    col_of = {}
    for c in range(1, ws.max_column + 1):
        key = ws.cell(HEADER_ROW, c).value
        if key:
            col_of[str(key).strip()] = c
    missing = [f for f in ("SPEC", "ITEM") if f not in col_of]
    if missing:
        raise ValueError(f"header row {HEADER_ROW} missing required field(s): {missing}")

    def cell(r, key):
        c = col_of.get(key)
        return ws.cell(r, c).value if c else None

    # group data rows by spec, preserving first-seen order
    groups: dict[str, list[int]] = {}
    order: list[str] = []
    for r in range(DATA_START_ROW, ws.max_row + 1):
        spec = cell(r, "SPEC")
        if not spec:
            continue
        spec = str(spec).strip()
        if spec not in groups:
            groups[spec] = []
            order.append(spec)
        groups[spec].append(r)

    errors = []
    classes = []
    for spec in order:
        rows = groups[spec]
        con_labeled = [r for r in rows if str(cell(r, "ITEM") or "").strip().upper() == "CON"]
        # A genuine CON (condition) row carries the P/T curve and has no
        # NPS range; a row labeled "CON" that DOES have DMIN/DMAX is a
        # source-data mislabel (seen in real HURL sheets: a pipe-schedule
        # row where someone typed "CON" instead of an item code) -- treat
        # it as an ordinary pipe_rule row below rather than hard-failing
        # the whole spec over one bad label, since which row is the real
        # CON is still mechanically unambiguous (DMIN/DMAX presence).
        real_con = [r for r in con_labeled
                   if _is_blank(cell(r, "DMIN")) and _is_blank(cell(r, "DMAX"))]
        if len(real_con) != 1:
            errors.append(f"{spec}: expected exactly 1 CON row (no NPS range), "
                          f"found {len(real_con)} among {len(con_labeled)} row(s) labeled CON")
            continue
        con = real_con[0]
        mislabeled_con = [r for r in con_labeled if r != con]

        ca_in = _num(cell(con, "CAIN"))
        if ca_in is None:
            ca_mm = _num(cell(con, "CAMM"))
            ca_in = round(ca_mm * IN_PER_MM, 4) if ca_mm is not None else 0.0

        pt_curve = []
        for i in range(1, N_PT_POINTS + 1):
            t_c = _num(cell(con, f"DT{i}S")); p_kg = _num(cell(con, f"DP{i}S"))
            t_f = _num(cell(con, f"DT{i}F")); p_psig = _num(cell(con, f"DP{i}F"))
            if t_c is None and p_kg is None and t_f is None and p_psig is None:
                continue
            pt_curve.append(PMS.PTPoint(
                temp_c=t_c, pressure_kgcm2g=p_kg,
                temp_f=t_f, pressure_psig=p_psig,
            ))

        pipe_rules = []
        for r in rows:
            if r == con:
                continue
            item = cell(r, "ITEM")
            dmin, dmax = cell(r, "DMIN"), cell(r, "DMAX")
            if not item or _is_blank(dmin) or _is_blank(dmax):
                errors.append(f"{spec} row {r}: item row missing ITEM/DMIN/DMAX")
                continue
            sched = normalize_schedule(cell(r, "SCHD") or "STD")
            thin = _num(cell(r, "THIN"))
            if thin is None:
                thmm = _num(cell(r, "THMM"))
                thin = round(thmm * IN_PER_MM, 4) if thmm is not None else None
            desc = str(item).strip()
            if r in mislabeled_con:
                desc = f"{desc} (mislabeled CON in source row {r})"
            pipe_rules.append(PMS.PipeRule(
                nps_low=nps_to_string(dmin), nps_high=nps_to_string(dmax),
                schedule=sched, ends="", description=desc,
                special_thickness_in=thin,
            ))

        design_p = pt_curve[0].pressure_psig if pt_curve else None
        design_t = pt_curve[0].temp_f if pt_curve else None

        try:
            pc = PMS.PipingClass(
                name=spec, moc=str(cell(con, "MOCS") or ""),
                moc_tag="", material_group=str(cell(con, "MATG") or ""),
                typical_service="", nominal_rating=f"{cell(con, 'PLBS') or ''} lbs".strip(),
                pt_limit_text="", design_pressure_psig=design_p or 0.0,
                design_temp_f=design_t if design_t is not None else 100.0,
                corrosion_allow_in=float(ca_in or 0.0),
                corrosion_allow_min_in=float(ca_in or 0.0),
                rev="", gems_ref=f"flat sheet '{sheet_name}'",
                pipe_rules=tuple(pipe_rules), pt_curve=tuple(pt_curve),
            )
            # facing/PWHT don't have dedicated PipingClass fields today --
            # fold into special_moc so nothing typed in is silently dropped.
            face = cell(con, "FACE")
            pwht = _yn(cell(con, "PWHT"))
            extra = []
            if face:
                extra.append(f"Face: {face}")
            extra.append(f"PWHT: {'YES' if pwht else 'NO'}")
            pc = replace(pc, special_moc="; ".join(extra))
            classes.append(pc)
        except Exception as e:  # noqa: BLE001
            errors.append(f"{spec}: {type(e).__name__}: {e}")

    if errors:
        raise ValueError("flat sheet has errors:\n  " + "\n  ".join(errors))
    return classes


def classes_to_flat_sheet(classes: list[PMS.PipingClass], xlsx_path: str,
                          sheet_name: str = "PMS") -> None:
    """Write the flat-sheet format from a PipingClass list.

    FPS T/P values are written as plain pre-computed numbers (not live
    formulas): openpyxl has no formula engine, so a formula written here
    reads back as an empty cell on the very next load (no application has
    ever opened the file to compute+cache a result) -- that would silently
    corrupt design_pressure_psig/design_temp_f on every round-trip that
    doesn't pass through real Excel first. They're meant as a starting
    suggestion only, not an enforced conversion -- real HURL data shows a
    spec's own FPS figures don't always match a straight SI conversion
    (e.g. a ~14.696 psi/(kgf/cm2) figure appears in practice, not the
    precise 14.223343 factor), so keep both independently editable.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]

    labels = {
        "SPEC": "Spec", "ITEM": "Item", "MOCS": "MOC", "MATG": "Material Group",
        "PLBS": "Rating", "FACE": "Face", "PWHT": "PWHT",
        "CAIN": "Corrosion Allowance", "CAMM": "Corrosion Allowance",
        "DMIN": "Min NPS", "DMAX": "Max NPS", "SCHD": "Schedule",
        "THIN": "Special Thickness (in)", "THMM": "Special Thickness (in)",
    }
    for i in range(1, N_PT_POINTS + 1):
        labels[f"DT{i}S"] = f"Design Temperature {i} (SI)"
        labels[f"DP{i}S"] = f"Design Pressure {i} (SI)"
        labels[f"DT{i}F"] = f"Design Temperature {i} (FPS)"
        labels[f"DP{i}F"] = f"Design Pressure {i} (FPS)"

    for i, key in enumerate(ALL_FIELDS, 2):  # col A reserved (free-text label)
        ws.cell(1, i, labels.get(key, key))
        if key.startswith(("DT", "DP")):
            ws.cell(2, i, "DEGC" if key[-1] == "S" and key[1] == "T" else
                    "KG/CM2G" if key[-1] == "S" else
                    "DEGF" if key[-1] == "F" and key[1] == "T" else "PSIG")
        ws.cell(HEADER_ROW, i, key)
    ws.freeze_panes = f"B{DATA_START_ROW}"

    col_of = {key: i for i, key in enumerate(ALL_FIELDS, 2)}

    dv_sched = DataValidation(type="list", formula1='"STD,XS,CAL,40,80,40S,80S,160"')
    dv_pwht = DataValidation(type="list", formula1='"YES,NO"')
    ws.add_data_validation(dv_sched)
    ws.add_data_validation(dv_pwht)

    r = DATA_START_ROW
    for pc in classes:
        con_row = r
        ws.cell(r, col_of["SPEC"], pc.name)
        ws.cell(r, col_of["ITEM"], "CON")
        ws.cell(r, col_of["MOCS"], pc.moc)
        ws.cell(r, col_of["MATG"], pc.material_group)
        rating_num = "".join(ch for ch in pc.nominal_rating if ch.isdigit()) or ""
        ws.cell(r, col_of["PLBS"], int(rating_num) if rating_num else pc.nominal_rating)
        face = ""
        for part in (pc.special_moc or "").split(";"):
            part = part.strip()
            if part.startswith("Face:"):
                face = part[len("Face:"):].strip()
        ws.cell(r, col_of["FACE"], face)
        ws.cell(r, col_of["PWHT"], "YES" if "PWHT: YES" in (pc.special_moc or "") else "NO")
        dv_pwht.add(ws.cell(r, col_of["PWHT"]))
        ws.cell(r, col_of["CAIN"], pc.corrosion_allow_in)
        ws.cell(r, col_of["CAMM"], round(pc.corrosion_allow_in / IN_PER_MM, 3))

        for i, pt in enumerate(pc.pt_curve[:N_PT_POINTS], 1):
            ws.cell(r, col_of[f"DT{i}S"], pt.temp_c)
            ws.cell(r, col_of[f"DP{i}S"], pt.pressure_kgcm2g)
            # Prefer the class's own stored FPS value (may not be a pure SI
            # conversion, see docstring above); only compute a fallback
            # suggestion from SI when the class has no FPS value of its own.
            temp_f = pt.temp_f if pt.temp_f is not None else (
                round(pt.temp_c * 1.8 + 32.0, 2) if pt.temp_c is not None else None)
            pressure_psig = pt.pressure_psig if pt.pressure_psig is not None else (
                round(pt.pressure_kgcm2g * PSI_PER_KGCM2, 2) if pt.pressure_kgcm2g is not None else None)
            ws.cell(r, col_of[f"DT{i}F"], temp_f)
            ws.cell(r, col_of[f"DP{i}F"], pressure_psig)
        r += 1

        for rule in pc.pipe_rules:
            ws.cell(r, col_of["SPEC"], pc.name)
            ws.cell(r, col_of["ITEM"], rule.description or "PIP")
            ws.cell(r, col_of["MOCS"], pc.moc)
            ws.cell(r, col_of["MATG"], pc.material_group)
            ws.cell(r, col_of["PLBS"], ws.cell(con_row, col_of["PLBS"]).value)
            ws.cell(r, col_of["FACE"], ws.cell(con_row, col_of["FACE"]).value)
            ws.cell(r, col_of["PWHT"], ws.cell(con_row, col_of["PWHT"]).value)
            ws.cell(r, col_of["CAIN"], pc.corrosion_allow_in)
            ws.cell(r, col_of["CAMM"], ws.cell(con_row, col_of["CAMM"]).value)
            ws.cell(r, col_of["DMIN"], nps_to_float(rule.nps_low))
            ws.cell(r, col_of["DMAX"], nps_to_float(rule.nps_high))
            ws.cell(r, col_of["SCHD"], rule.schedule)
            dv_sched.add(ws.cell(r, col_of["SCHD"]))
            if rule.special_thickness_in is not None:
                ws.cell(r, col_of["THIN"], rule.special_thickness_in)
                ws.cell(r, col_of["THMM"], round(rule.special_thickness_in / IN_PER_MM, 3))
            r += 1

    wb.save(xlsx_path)
