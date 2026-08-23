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
import pms_flat_sheet as PFS        # noqa: E402
import hmb_proii_reader as HMBP     # noqa: E402

# wire the Case-N reader into the engine dispatch (it imports it as optional)
H.HMBPROII = HMBP

LOGO_PATH = os.path.join(_HERE, "resources", "ten_logo.png")
# ProCalc's own brand mark (blue pipe-fitting + water-drop icon) — separate
# from the T.EN company logo above. Every consumer guards with
# os.path.exists() since the asset may not be present in every checkout.
PROCALC_LOGO_PATH = os.path.join(_HERE, "resources", "procalc_logo.png")


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


def hmb_source_kind(path: str) -> str:
    """One of "hysys" | "proii_com" | "proii_xlsx" | "per_stream_xlsx" —
    extension-routed for the two live-connector sources, content-sniffed
    for the two Excel-based ones (see hydraulics_XOM._source_kind)."""
    try:
        return H._source_kind(path)
    except Exception:
        return "per_stream_xlsx"


def list_cases(path: str) -> list[str]:
    """Case/session names for an HMB source. For the two live-connector
    kinds (HYSYS/.hsc, PRO/II/.prz) a connection failure is NOT swallowed
    here — it propagates so the caller (the app's "Load HMB" action) can
    show the user a clear, actionable error instead of a silently-empty
    case list. The two Excel-based kinds keep their existing
    swallow-to-safe-default behavior unchanged."""
    kind = hmb_source_kind(path)
    if kind == "hysys":
        if H.HYSYS_COM is None:
            raise RuntimeError(
                "Live HYSYS connection requires Windows with pywin32 "
                "installed and a running Aspen HYSYS instance.")
        return H.HYSYS_COM.list_cases(path)
    if kind == "proii_com":
        if H.PROII_COM is None:
            raise RuntimeError(
                "Live PRO/II connection requires Windows with pywin32 "
                "installed and a licensed PRO/II install.")
        return H.PROII_COM.list_cases(path)
    if kind == "proii_xlsx":
        try:
            return HMBP.list_cases(path)
        except Exception:
            return ["Case 1"]
    return []                      # per-stream dump has no case sheets


def list_streams(path: str, case: str = "Case 1") -> list[str]:
    """Stream names in an HMB source (all four layouts). Live-connector
    failures propagate — see list_cases()'s docstring."""
    kind = hmb_source_kind(path)
    if kind == "hysys":
        if H.HYSYS_COM is None:
            raise RuntimeError(
                "Live HYSYS connection requires Windows with pywin32 "
                "installed and a running Aspen HYSYS instance.")
        return H.HYSYS_COM.list_streams(path, case)
    if kind == "proii_com":
        if H.PROII_COM is None:
            raise RuntimeError(
                "Live PRO/II connection requires Windows with pywin32 "
                "installed and a licensed PRO/II install.")
        return H.PROII_COM.list_streams(path, case)
    try:
        if kind == "proii_xlsx":
            return HMBP.list_streams(path, case)
        from pathlib import Path
        return H.list_streams(Path(path))
    except Exception:
        return []


def list_all_stream_data(hmb_path: str, case: str = "Case 1") -> dict:
    """{stream_name: {"props": StreamProps|None, "composition": dict|None}}
    for every stream in the HMB source, batched where the source format
    supports it cheaply:
      - per_stream_xlsx: a single-open batch reader (extract_all_streams/
        extract_all_compositions) -- reading each stream individually would
        reopen the whole workbook per stream, ~40x slower on a real
        400+-stream file (measured).
      - proii_xlsx: hmb_proii_reader's own per-case parse is already
        memoized (one parse serves every stream), so this just loops the
        already-cheap per-stream accessors.
      - hysys / proii_com (live): resolving 400+ streams over a live COM
        connection is impractical (COM round-trips, and for proii_com a
        RunCalcs()-backed session) -- returns names only (props/composition
        left None); the Streams table resolves those lazily per-column via
        resolve_stream_for_snapshot(), the same on-demand path already used
        elsewhere for live sources.
    """
    kind = hmb_source_kind(hmb_path)
    out: dict[str, dict] = {}

    if kind == "per_stream_xlsx":
        from pathlib import Path
        all_props = H.extract_all_streams(Path(hmb_path))
        all_comp = H.extract_all_compositions(Path(hmb_path))
        for name, sp in all_props.items():
            out[name] = {"props": sp, "composition": all_comp.get(name)}
        return out

    if kind == "proii_xlsx":
        try:
            import td_parser
            streams, comp_data, _names = td_parser.parse_proii(hmb_path, case)
        except Exception:
            streams, comp_data = [], {}
        for name in streams:
            try:
                hs = HMBP.get_stream(name, path=hmb_path, case=case)
                sp = H.streamprops_from_hmb(hs) if hs else None
            except Exception:
                sp = None
            out[name] = {"props": sp,
                        "composition": comp_data.get(name, {}).get("MOLE_FRAC") or None}
        return out

    try:
        names = list_streams(hmb_path, case)
    except Exception:
        names = []
    for name in names:
        out[name] = {"props": None, "composition": None}
    return out


# ── units ────────────────────────────────────────────────────────────────
def detect_hmb_units(hmb_path: str, case: str | None = None) -> tuple[str, dict]:
    """Best-guess (unit_system, unit_overrides) from an HMB workbook's own
    printed unit strings — for propagating them as fresh-project defaults.
    Returns ("FPS", {}) on any failure/ambiguity (safe no-op default).
    unit_overrides is {} when the HMB's units already match the chosen
    system's defaults exactly (no redundant overrides)."""
    try:
        if is_proii_export(hmb_path):
            raw = HMBP.extract_units(hmb_path, case or "Case 1")
        else:
            from pathlib import Path
            raw = H.extract_stream_units(Path(hmb_path))
    except Exception:
        return "FPS", {}
    if not raw:
        return "FPS", {}

    normalized = {}
    for qty, u in raw.items():
        key = " ".join(str(u).strip().upper().split())
        norm = H.UN.HMB_UNIT_ALIASES.get(qty, {}).get(key)
        if norm:
            normalized[qty] = norm
    if not normalized:
        return "FPS", {}

    best_system, best_score = "FPS", -1
    for sysname, defaults in H.UN.SYSTEM_DEFAULTS.items():
        score = sum(1 for qty, u in normalized.items() if defaults.get(qty) == u)
        if score > best_score:
            best_system, best_score = sysname, score

    defaults = H.UN.SYSTEM_DEFAULTS[best_system]
    overrides = {qty: u for qty, u in normalized.items() if defaults.get(qty) != u}
    return best_system, overrides


# ── PMS ────────────────────────────────────────────────────────────────────
def pms_class_names() -> list[str]:
    return [c.name for c in PMS.CLASSES]


def install_pms_catalogue(json_path: str) -> int:
    """Validate + activate an uploaded gems_extracted.json, reindex the engine."""
    n = PMS.install_catalogue(json_path)
    H.reload_pms()
    return n


def install_pms_catalogue_from_flat_sheet(xlsx_path: str, sheet_name: str = "PMS") -> int:
    """Validate + activate an uploaded flat-sheet PMS workbook (the human-
    editable master format -- see Hydraulics/pms_flat_sheet.py), reindex the
    engine.  Raises ValueError with a row-numbered message list on anything
    malformed, same contract as install_pms_catalogue()."""
    classes = PFS.flat_sheet_to_classes(xlsx_path, sheet_name)
    path = PMS.user_json_path()
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    PMS.save_user_classes(classes, path)
    H.reload_pms(path)
    return len(PMS.CLASSES)


def export_pms_catalogue_to_flat_sheet(xlsx_path: str, sheet_name: str = "PMS") -> int:
    """Write the active (non-curated) catalogue out as a flat-sheet workbook.
    Returns the number of classes written."""
    curated_names = {pc.name.upper() for pc in PMS.CURATED}
    classes = [c for c in PMS.CLASSES if c.name.upper() not in curated_names]
    PFS.classes_to_flat_sheet(classes, xlsx_path, sheet_name)
    return len(classes)


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


def units_for(qty: str) -> list[str]:
    """Every unit string common/units.py knows for one quantity code."""
    return list(H.UN._CONVERTERS.get(qty, {}).keys())


def internal_unit(qty: str) -> str | None:
    """The unit a quantity's value is in internally (the engine's FPS
    canonical basis, e.g. "psia" for pressure, "degF" for temperature) --
    every StreamProps/FlashFeed field is already expressed in this unit,
    so a widget seeding its initial display from a raw internal value
    should label it with this, not guess at dict ordering."""
    return H.UN.SYSTEM_DEFAULTS["FPS"].get(qty)


def convert(qty: str, value, from_unit: str, to_unit: str):
    """Convert a single value between two named units of the same quantity
    (e.g. qty="P", from_unit="psia", to_unit="bara") — the single-value
    primitive UnitSystem doesn't expose directly (it only converts against
    whichever unit is currently selected for a whole UnitSystem instance).
    Returns `value` unchanged if the quantity/unit pair is unknown or value
    is None (mirrors UnitSystem.to_internal/from_internal's own leniency)."""
    if value is None:
        return None
    conv = H.UN._CONVERTERS.get(qty)
    if not conv or from_unit not in conv or to_unit not in conv:
        return value
    try:
        v = float(value)
    except (TypeError, ValueError):
        return value
    internal = conv[from_unit][0](v)
    return conv[to_unit][1](internal)


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
