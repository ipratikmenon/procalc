# Establishing a live PRO/II COM interface to read `.prz` files

This documents how Procalc talks to a `.prz` (PRO/II database) file directly,
what's already implemented and verified, what's still provisional, and how
to finish verifying/tuning it against a real PRO/II install. The
implementation lives in `Hydraulics/proii_com_reader.py`.

## Why two halves

A `.prz` is a zip archive. Opening it doesn't require PRO/II at all — but
getting live, freshly-solved values out of it does. So the reader is split
into two independent pieces:

| | Requires | Gives you |
|---|---|---|
| **Listing** (`list_cases`, `list_streams`) | nothing — pure Python, any OS | stream names + component index, read straight from the `.prz`'s own embedded keyword backup |
| **Live values** (`get_stream`) | Windows, PRO/II installed + licensed, `pywin32` | actual resolved stream properties + composition, via COM |

This split matters because it's the difference between "the app can browse
what streams exist" (works everywhere, instantly) and "the app can pull a
live number out of a solved flowsheet" (needs the real software, a license
seat, and a `RunCalcs()` that re-solves the whole model).

## Part 1 — listing streams without COM (done, verified)

Every `.prz` embeds a plain-text keyword (`.inp`) backup of the flowsheet.
`_load_inp_data()` unzips it, finds the `.inp` member, and `_parse_inp()`
extracts:

- **Stream names**, from three distinct keyword forms (a single blanket
  regex both false-positives on `PRINT`'s `STREAM=ALL` control keyword and
  misses names that only appear in comma-separated lists — confirmed
  against two real sample files):
  ```
  PROPERTY STREAM=<name>, ...
  ... REFSTREAM=<name>, ...
  OUTPUT FORMAT=..., STREAMS=<name1>,<name2>,..., ...
  ```
- **Component index**, from `LIBID` lines (`LIBID 1,H2O/2,CO2/3,CO/...`) —
  the same index PRO/II's own `GetAttribute("TotalComposition", idx)` call
  is keyed against.

This was verified directly against two real user-supplied `.prz` files
during development (14 streams / 78 components, and 40 streams / 1
component) — the parsing logic itself is solid.

**Caveat to keep visible in the UI/logs:** this reflects the flowsheet
structure as of the file's *last save/keyword-export*, not necessarily
the live in-memory state if someone has since edited the case without
saving. Treat it as a structure hint for populating the stream picker, not
as authoritative data — live values always come from Part 2.

## Part 2 — live values via COM (implemented, NOT yet tested against real PRO/II)

This is the part that needs a real Windows + PRO/II machine to finish
verifying. Everything below is grounded in one working public example
(`github.com/bryanpiguave/Air-Separation`) plus cross-referencing PRO/II's
own report vocabulary (`Hydraulics/HMB.xlsx`'s "TD Property Dump" labels
and `hmb_proii_reader.py`'s Case-N label set) — not hands-on testing.

### Connecting

```python
import win32com.client as win32
pro2 = win32.Dispatch(f"SimSciDbs.Database.{version}")   # e.g. "110" or "102"
db = pro2.OpenDatabase(abspath_to_prz)
db.RunCalcs()          # re-solves the whole flowsheet — not instant, consumes a license seat
```

`_connect()` in `proii_com_reader.py` tries a list of ProgID version
suffixes in order (`_PROGID_VERSIONS = ("110", "102", "100", "95", "91",
"90")`) until one dispatches. Only `.102`/`.110` are confirmed from the
public example; the rest are speculative fallbacks. **First thing to check
on a real machine:** which version string actually matches your PRO/II
install. If none of the listed ones work, find the real ProgID via the
Windows registry (`HKEY_CLASSES_ROOT\SimSciDbs.Database.*`) or PRO/II's own
installation docs, and add it to `_PROGID_VERSIONS`.

The opened `db` is cached per `(path, mtime)` in `_SESSION_CACHE` — so
`RunCalcs()` only re-solves once per session unless the file changes on
disk, not once per stream read.

### Reading one stream

```python
strm = db.ActivateObject("Stream", stream_name)
value = strm.GetAttribute("TemperatureCalc")
composition = strm.GetAttribute("TotalComposition", component_index)
```

`_SCALAR_ATTRS` (in `proii_com_reader.py`) is the current attribute table —
17 scalar properties, PascalCase, mostly `Calc`-suffixed. Only
**`TotalMolarRate`** and **`TotalComposition`** are confirmed real (from the
public example); everything else is a plausible sibling name inferred from
that same naming convention plus PRO/II's own report vocabulary. Expect
some of these to be wrong on first real test.

### What "iterate against real errors" actually looks like

This is the expected, planned-for workflow — not a sign something is
broken:

1. Run Procalc on the Windows machine with PRO/II installed, `pywin32`
   installed (`pip install pywin32`), and licensed.
2. Load a `.prz` → open **Streams** → pick a stream.
3. If it fails, the exception message will name the exact COM call that
   errored (dispatch failure, `OpenDatabase`/`RunCalcs` failure, or one
   specific `GetAttribute` call). Common failure shapes and what they mean:
   - **Dispatch fails for every version in `_PROGID_VERSIONS`** → the
     installed ProgID isn't in the list. Find the real one (registry or
     PRO/II docs) and add it.
   - **`OpenDatabase`/`RunCalcs` fails** → usually a license-seat or
     file-lock issue (the message already checks for these), or the `.prz`
     needs to be opened differently than a raw path (e.g. via `Import()` on
     the extracted `.inp` instead — worth trying if `OpenDatabase` rejects
     the packed `.prz` directly).
   - **One specific `GetAttribute(attr)` raises** → that attribute name is
     wrong for this PRO/II version. PRO/II's **Object Browser**
     (accessible from within the PRO/II GUI, or via its COM type library
     if registered) is the authoritative source for real attribute names —
     browse to a Stream object and look at its actual property/attribute
     list, then correct the entry in `_SCALAR_ATTRS`.
   - **Composition reads return nothing** → try `VaporComposition`/
     `LiquidComposition`'s exact spelling against the Object Browser too;
     these are PascalCase guesses, not confirmed like `TotalComposition`.
4. Fix the specific name/call that failed, re-test, repeat. Each fix is
   small and isolated — the plumbing around it (dispatch, session caching,
   error surfacing, composition indexing) is already solid.

### A faster way to find real attribute names, if available

If the PRO/II install exposes a registered COM type library, `pywin32`'s
`makepy` utility (or simply `win32com.client.gencache.EnsureDispatch(...)`
instead of plain `Dispatch(...)`) can generate a wrapper with real method/
property names and autocomplete in an interactive session — much faster
than trial-and-error against `_SCALAR_ATTRS` one attribute at a time. Worth
trying first:
```python
import win32com.client
pro2 = win32com.client.gencache.EnsureDispatch("SimSciDbs.Database.110")
```
then inspect `pro2._prop_map_get_` / use `dir()` on the returned objects.

## How it's wired into the app

- `Hydraulics/hydraulics_XOM.py` imports this module as an optional slot
  (`PROII_COM`, guarded `try/except ImportError` so a non-Windows dev
  environment or missing `pywin32` degrades cleanly to "feature
  unavailable," never a crash) and dispatches `.prz` paths to it exactly
  like every other HMB source kind.
- `procalc_app/engine_api.py`'s `hmb_source_kind()` reports `"proii_com"`
  for a `.prz` path; the app's toolbar, error bar, and the Streams dialog
  (`procalc_app/streams/streams_view.py`) all branch on this the same way
  they already do for HYSYS live connections.
- The Streams dialog's live-connector view
  (`procalc_app/streams/live_sheet.py`) works against `get_stream()`'s
  return shape unchanged — once `_SCALAR_ATTRS` is correct, the resolved
  values flow straight through to the same styled, faithful-to-HMB.xlsx
  sheet view with no further wiring needed.

## Packaging note

`pywin32` is already declared as a Windows-only dependency
(`procalc_app/requirements.txt`, environment-marker-gated) and its COM
hidden imports are already in `build/procalc.spec`, guarded for
`sys.platform == "win32"`. No packaging changes needed to test this.
