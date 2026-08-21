"""The single seam between the Qt app and the hydraulics engine.

No other app module imports the engine directly.  This inserts the engine
directories on sys.path, then exposes thin wrappers for everything the UI
needs: the input schema + drop-down lists, the HMB stream/case listers, a
run() that writes a temp input workbook and calls run_noiso, the PMS
catalogue upload/reload, and resolve_id passthroughs for live previews.
"""
from __future__ import annotations

import os
import sys
import tempfile

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
for _p in (os.path.join(_ROOT, "common"), os.path.join(_ROOT, "Hydraulics")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import hydraulics_XOM as H          # noqa: E402
import pms_classes as PMS           # noqa: E402
import hmb_proii_reader as HMBP     # noqa: E402

# wire the Case-N reader into the engine dispatch (it imports it as optional)
H.HMBPROII = HMBP

LOGO_PATH = os.path.join(_HERE, "resources", "ten_logo.png")


# ── schema / drop-downs ────────────────────────────────────────────────────
def input_headers() -> list[str]:
    return list(H.INPUT_HEADERS)


def field_meta() -> dict:
    return dict(H._FIELD_META)


def display_header(canon: str, unit_system: str = "FPS") -> str:
    return H._display_header(canon, H.UN.UnitSystem(unit_system))


def dropdown_lists() -> dict:
    """Column name -> list of allowed values (for combobox delegates)."""
    return {
        "Run Type": ["Main", "Branch"],
        "Fitting Name": list(H.FITTING_NAMES),
        "Control Valve Type": ["", "F", "P", "T", "L"],
        "Valve Body Style": ["", "Globe", "Angle", "Ball", "Butterfly", "Eccentric"],
        "Valve Characteristic": ["", "Linear", "Equal%", "Quick-Open"],
        "Seat Leakage Class": ["", "II", "III", "IV", "V", "VI"],
        "Direction": ["", "N", "S", "E", "W", "UP", "DOWN"],
        "Piping Spec": [c.name for c in PMS.CLASSES],
        "Bore (in)": [str(x) for x in H.PIPE_OD_IN.keys()],
    }


def numeric_columns() -> set[str]:
    """Columns whose values are numeric (quantity code present, or known num)."""
    out = set()
    fm = H._FIELD_META
    for k, (_b, qty) in fm.items():
        if qty is not None:
            out.add(k)
    out.update({"Seq", "Bore (in)", "Fixed K", "Design Opening %", "Rated Cv",
                "Min Flow Mult", "Max Flow Mult", "Noise Limit dBA",
                "Inlet Line Size (in)", "Outlet Line Size (in)",
                "Flow Fraction from Main", "Mass Vapor Fraction Carry Over",
                "Mass Liquid Fraction Carry Over"})
    return out


# ── HMB ────────────────────────────────────────────────────────────────────
def is_proii_export(path: str) -> bool:
    try:
        return bool(H.is_proii_export(path))
    except Exception:
        return False


def list_cases(path: str) -> list[str]:
    if is_proii_export(path):
        try:
            return HMBP.list_cases(path)
        except Exception:
            return ["Case 1"]
    return []                      # per-stream dump has no case sheets


def list_streams(path: str, case: str = "Case 1") -> list[str]:
    """Stream names in an HMB dump (both layouts)."""
    try:
        if is_proii_export(path):
            return HMBP.list_streams(path, case)
        from pathlib import Path
        return H.list_streams(Path(path))
    except Exception:
        return []


# ── PMS ────────────────────────────────────────────────────────────────────
def pms_class_names() -> list[str]:
    return [c.name for c in PMS.CLASSES]


def install_pms_catalogue(json_path: str) -> int:
    """Validate + activate an uploaded gems_extracted.json, reindex the engine."""
    n = PMS.install_catalogue(json_path)
    H.reload_pms()
    return n


def reload_pms(json_path: str | None = None) -> int:
    return H.reload_pms(json_path)


def resolve_id(spec, bore):
    return H.resolve_id(spec, bore)


def resolve_id_preview(working_class, bore):
    """resolve_id using an UNSAVED, edited PipingClass (for the PMS live
    preview) without mutating the active catalogue.  Temporarily swaps the
    working class into the engine's name index, resolves, then restores."""
    key = (working_class.name or "").upper()
    prev = H._CLASS_BY_NAME.get(key)
    try:
        H._CLASS_BY_NAME[key] = working_class
        return H.resolve_id(working_class.name, bore)
    finally:
        if prev is not None:
            H._CLASS_BY_NAME[key] = prev
        else:
            H._CLASS_BY_NAME.pop(key, None)


def pms_classes_module():
    """Expose the pms_classes module (dataclasses + CLASSES) for the manager."""
    return PMS


def wall_for(nps, schedule):
    return H.wall_for(nps, schedule)


def flange_class_for(spec):
    return H.flange_class_for(spec)


# ── run ────────────────────────────────────────────────────────────────────
def new_project_rows() -> list[dict]:
    """A small starter circuit (Source -> pipe -> elbow) for a blank project."""
    return [
        {"Circuit": "C1", "Line No": "L-01", "PID Number": "PID-001",
         "Run Type": "Main", "Seq": 1, "Comp ID": "SRC", "Fitting Name": "Source",
         "Bore (in)": 6, "Piping Spec": "G1A-5", "Set P (psia)": 150.0,
         "Notes": "Upstream source"},
        {"Circuit": "C1", "Line No": "L-01", "Run Type": "Main", "Seq": 2,
         "Comp ID": "P-01", "Fitting Name": "Straight Pipeline", "Bore (in)": 6,
         "Piping Spec": "G1A-5", "Length (ft)": 40.0},
        {"Circuit": "C1", "Line No": "L-01", "Run Type": "Main", "Seq": 3,
         "Comp ID": "EL-01", "Fitting Name": "Elbow 90 Long", "Bore (in)": 6,
         "Piping Spec": "G1A-5", "Length (ft)": 0.0},
    ]


def write_input_workbook(rows: list[dict], out_path: str,
                         unit_system: str = "FPS",
                         unit_overrides: dict | None = None) -> str:
    return H.write_input_workbook(rows, out_path, unit_system, unit_overrides)


def unit_systems() -> list[str]:
    return list(H.UN.SYSTEM_DEFAULTS.keys())


def unit_quantities() -> list[tuple[str, str, dict]]:
    """[(qty_code, display name, {unit: is_default_for_selected}) ...] for the
    per-quantity override editor.  Returns code, name, and the allowed units
    plus the FPS/SI defaults so the UI can show 'default' hints."""
    out = []
    conv = H.UN._CONVERTERS
    for qty, name in H.UN.QUANTITY_NAMES.items():
        units = list(conv.get(qty, {}).keys())
        defs = {sys: H.UN.SYSTEM_DEFAULTS[sys].get(qty) for sys in H.UN.SYSTEM_DEFAULTS}
        out.append((qty, name, {"units": units, "defaults": defs}))
    return out


def resolve_stream_for_snapshot(hmb_path: str, case: str, stream: str):
    """Resolve one (hmb, case, stream) to (StreamProps, feed|None), for
    building a portable .calc snapshot. Raises if the stream isn't found —
    the caller (Project.save) should surface that clearly rather than save
    a project that silently can't replay."""
    sp = H.load_stream_props(hmb_path, stream, case)
    if sp is None:
        raise ValueError(f"Stream '{stream}' not found in {hmb_path} (case '{case}')")
    is_case = H.is_proii_export(hmb_path)
    feed = H.read_feed(hmb_path, stream, case, sp, is_casesheet=is_case)
    return sp, feed


def snapshot_keys_for_rows(rows: list[dict], hmb_path: str, case: str) -> list[tuple]:
    """Every distinct (hmb, case, stream) the given grid rows actually
    reference — the circuit-level default plus any per-row overrides —
    so a caller can build a complete .calc snapshot without missing a
    stream a march would otherwise need."""
    seen, out = set(), []
    for r in rows:
        stream = r.get("Stream Lookup")
        if not stream:
            continue
        key = (r.get("HMB File") or hmb_path, r.get("Case") or case, stream)
        if key not in seen:
            seen.add(key)
            out.append(key)
    return out


def stream_snapshot_to_json(snapshot: dict) -> dict:
    """{(hmb, case, stream): (StreamProps, feed|None)} -> a JSON-safe dict
    (string keys, plain-dict values) for writing into a .calc file."""
    import dataclasses
    out = {}
    for (hmb, case, stream), (sp, feed) in snapshot.items():
        out[f"{hmb}||{case}||{stream}"] = {
            "sp": dataclasses.asdict(sp) if sp is not None else None,
            "feed": dataclasses.asdict(feed) if feed is not None else None,
        }
    return out


def stream_snapshot_from_json(d: dict) -> dict:
    """Reverse of stream_snapshot_to_json — rebuilds real StreamProps/
    FlashFeed instances (run_noiso accesses them by attribute, so plain
    dicts won't do) from a .calc file's JSON-safe snapshot payload.

    FlashFeed.k_special is a dict[int, tuple[str, float]] — JSON coerces
    int keys to strings and tuples to lists, so it needs an explicit
    fixup on the way back in (verified: every other field round-trips
    cleanly through dataclasses.asdict()/JSON as-is)."""
    out = {}
    for key, pair in d.items():
        hmb, case, stream = key.split("||", 2)
        sp_d, feed_d = pair.get("sp"), pair.get("feed")
        sp = H.StreamProps(**sp_d) if sp_d is not None else None
        feed = None
        if feed_d is not None:
            ks = feed_d.get("k_special")
            if ks:
                feed_d = dict(feed_d,
                              k_special={int(k): tuple(v) for k, v in ks.items()})
            feed = H.FlashFeed(**feed_d)
        out[(hmb, case, stream)] = (sp, feed)
    return out


def run(rows: list[dict], hmb_path: str, *, case: str = "Case 1",
        flash_mode: str = "isothermal", out_dir: str | None = None,
        meta: dict | None = None, unit_system: str = "FPS",
        unit_overrides: dict | None = None,
        stream_snapshot: dict | None = None) -> list[str]:
    """Write a temp input workbook from the grid rows and run the engine.

    ``stream_snapshot`` (optional): a pre-resolved ``{(hmb, case, stream):
    (StreamProps, feed)}`` dict (see ``resolve_stream_for_snapshot``) —
    lets a restored .calc project replay even if ``hmb_path`` no longer
    resolves to a real file on disk.

    Returns the list of produced workbook paths (one per circuit).  Raises on
    engine error (the caller runs this off the UI thread and shows the error).
    """
    out_dir = out_dir or tempfile.mkdtemp(prefix="procalc_out_")
    inp = os.path.join(out_dir, "pipeline_input.xlsx")
    write_input_workbook(rows, inp, unit_system, unit_overrides)
    results = H.run_noiso(inp, hmb_path, out_path=None, case=case,
                          flash_mode=flash_mode, meta=meta,
                          stream_snapshot=stream_snapshot)
    # run_noiso writes into cwd when out_path is None; move them into out_dir
    paths = []
    import shutil
    for tup in results:
        op = tup[0]
        if op and os.path.exists(op):
            dest = os.path.join(out_dir, os.path.basename(op))
            if os.path.abspath(op) != os.path.abspath(dest):
                shutil.move(op, dest)
            paths.append(dest)
    return paths
