#!/usr/bin/env python3
"""HyCalign Hydraulics — single front end (console menu by default, --gui for window).

    python run_hydraulics.py          # interactive console menu
    python run_hydraulics.py --gui    # tkinter GUI

  command line (no questions):
    python run_hydraulics.py build DUMP.xlsx [HMB_OUT.xlsx] [--sheet CASE2] [--constants C.xlsx]
    python run_hydraulics.py run   INPUT.xlsx HMB.xlsx [CIRCUIT] [--case "Case 1"] [--isenthalpic] [--light]
    python run_hydraulics.py all   DUMP.xlsx INPUT.xlsx [CIRCUIT]   # build the HMB, then run the hydraulics

Engine lives in hydraulics.py (one file: dP marching, flash VLE, FIV/AIV, flow-pattern maps).
"""
from __future__ import annotations



# ══════════════════════════════════════════════════════════════════════════
# ║  CONSOLE LAUNCHER (former hydraulics_launcher.py)
# ══════════════════════════════════════════════════════════════════════════
#!/usr/bin/env python3
"""
HyCalign — No-ISO Hydraulics interactive launcher.

A menu-driven front end so you never have to type a full command line:
  • lists the input workbooks and HMB workbooks in this folder,
  • lets you search/pick the HMB case,
  • lists the circuits found in the chosen input,
  • picks the flash mode,
  • then runs the engine and reports where the workbooks landed.

Run it with:   python3 hydraulics_launcher.py
(or double-click run_hydraulics.command on macOS / run_hydraulics.bat on Windows)
"""

import os
import re
import sys
import glob
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)

# ── pretty console helpers ──────────────────────────────────────────────────
_USE_COLOR = sys.stdout.isatty() and os.environ.get("NO_COLOR") is None


def _c(text: str, code: str) -> str:
    return f"\033[{code}m{text}\033[0m" if _USE_COLOR else text


def bold(t):   return _c(t, "1")
def cyan(t):   return _c(t, "36")
def green(t):  return _c(t, "32")
def yellow(t): return _c(t, "33")
def red(t):    return _c(t, "31")
def dim(t):    return _c(t, "2")


def rule(char="─", n=64):
    print(dim(char * n))


def banner():
    rule("═")
    print(bold(cyan("  HyCalign — No-ISO Flash Hydraulics  ·  Launcher")))
    print(dim(f"  {HERE}"))
    rule("═")


def ask(prompt: str, default: str | None = None) -> str:
    suffix = f" [{default}]" if default is not None else ""
    try:
        val = input(f"{prompt}{suffix}: ").strip()
    except (EOFError, KeyboardInterrupt):
        print()
        sys.exit(0)
    return val or (default or "")


# ── workbook discovery ──────────────────────────────────────────────────────
def list_xlsx() -> list[str]:
    files = sorted(glob.glob("*.xlsx"))
    # ignore generated output workbooks
    return [f for f in files if not f.startswith("hydraulics_flash")]


def is_input_wb(path: str) -> bool:
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True)
        ok = "Pipeline_Input" in wb.sheetnames
        wb.close()
        return ok
    except Exception:
        return False


def _sheetnames(path: str) -> list[str]:
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


_KIND_CACHE: dict = {}


def kind_of(path: str) -> str:
    """'input' (Pipeline_Input) | 'dump' (raw PRO/II property dump) | 'hmb' (built HMB) | 'other'."""
    if path in _KIND_CACHE:
        return _KIND_CACHE[path]
    k = "other"
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True)
        try:
            names = list(wb.sheetnames)
            if "Pipeline_Input" in names:
                k = "input"
            else:
                for n in names[:3]:                       # a PRO/II dump has a 'Stream Name' row in col A
                    hit = False
                    for r in wb[n].iter_rows(min_row=1, max_row=12, max_col=1, values_only=True):
                        if r and r[0] is not None and str(r[0]).strip() == "Stream Name":
                            hit = True
                            break
                    if hit:
                        k = "dump"
                        break
                if k == "other":
                    up = {x.upper() for x in names}
                    if "COMP_CONSTANTS" in up or ("UNITS" in up and "COMPONENTS" in up):
                        k = "hmb"
        finally:
            wb.close()
    except Exception:
        k = "other"
    _KIND_CACHE[path] = k
    return k


def categorise(files: list[str]) -> tuple[list[str], list[str]]:
    """(input workbooks, everything else)."""
    inputs, others = [], []
    for f in files:
        (inputs if kind_of(f) == "input" else others).append(f)
    return inputs, others


def ask_path(title: str) -> str | None:
    """Type a path to a workbook anywhere on disk (quotes allowed; Enter cancels)."""
    raw = ask(title + " (full path, Enter = cancel)")
    raw = raw.strip().strip('"').strip("'")
    if not raw:
        return None
    if not os.path.isfile(raw):
        print(red(f"  Not found: {raw}"))
        return None
    return raw


def choose_from(title: str, items: list[str], allow_blank=False,
                extra: list[tuple[str, str]] | None = None) -> str | None:
    """Numbered chooser.  ``extra`` = [(key,label)] for non-numeric actions."""
    print()
    print(bold(title))
    for i, it in enumerate(items, 1):
        print(f"  {cyan(str(i)):>3}  {it}")
    for key, label in (extra or []):
        print(f"  {yellow(key):>3}  {label}")
    while True:
        raw = ask("Choose")
        if raw == "" and allow_blank:
            return None
        for key, _ in (extra or []):
            if raw.lower() == key.lower():
                return key
        if raw.isdigit() and 1 <= int(raw) <= len(items):
            return items[int(raw) - 1]
        print(red("  Invalid choice — try again."))


# ── HMB case picker (search-driven; HMBs can hold hundreds of cases) ─────────
def choose_case(hmb_path: str) -> str:
    """Only workbooks with 'Case N' sheets (PRO/II case export) need a case; a built per-stream HMB does not."""
    try:
        import hmbbuilder as H
        cases = [c for c in H.list_cases(hmb_path) if re.match(r"(?i)case", c)]
    except Exception as exc:
        print(dim(f"  (could not read cases from {hmb_path}: {exc})"))
        cases = []

    if not cases:
        return "Case 1"

    print()
    print(bold(f"HMB cases available: {len(cases)}"))
    print(dim("  Press Enter for 'Case 1', or type part of a case name to search."))
    while True:
        q = ask("Case (search / exact / Enter=default)", default="Case 1")
        if q == "Case 1" and "Case 1" not in cases:
            # default kept even if not literally present (engine handles single-case HMBs)
            return q
        if q in cases:
            return q
        hits = [c for c in cases if q.lower() in c.lower()]
        if not hits:
            print(red(f"  No case matches '{q}'. Using it as typed."))
            return q
        if len(hits) == 1:
            print(green(f"  → {hits[0]}"))
            return hits[0]
        shown = hits[:40]
        sel = choose_from(f"{len(hits)} matches (showing {len(shown)})",
                          shown, allow_blank=True,
                          extra=[("s", "search again")])
        if sel is None or sel == "s":
            continue
        return sel


# ── circuit picker ──────────────────────────────────────────────────────────
def list_circuits(input_path: str):
    try:
        import hydraulics as N
        rows = N.read_pipeline_input(input_path)
        circuits = N._group_circuits(rows)
        out = []
        for cid, c in circuits.items():
            lines = ", ".join(c["line_order"])
            stream = c["stream_lookup"] or "(manual)"
            out.append((cid, lines, stream))
        return out
    except Exception as exc:
        print(dim(f"  (could not parse circuits: {exc})"))
        return []


def choose_circuit(input_path: str) -> str | None:
    circuits = list_circuits(input_path)
    if not circuits:
        return None
    print()
    print(bold(f"Circuits in {os.path.basename(input_path)}:"))
    labels = [f"{cid}   ·   lines: {lines}   ·   stream: {stream}"
              for cid, lines, stream in circuits]
    print(f"  {yellow('a'):>3}  ALL circuits")
    for i, lab in enumerate(labels, 1):
        print(f"  {cyan(str(i)):>3}  {lab}")
    while True:
        raw = ask("Choose circuit", default="a")
        if raw.lower() == "a":
            return None
        if raw.isdigit() and 1 <= int(raw) <= len(circuits):
            return circuits[int(raw) - 1][0]
        print(red("  Invalid choice — try again."))


# ── actions ─────────────────────────────────────────────────────────────────
def make_template():
    import hydraulics as N
    name = ask("Template file name", default="pipeline_input_noiso.xlsx")
    usys = ask("Unit system (FPS/SI)", default="FPS").strip().upper()
    usys = "SI" if usys == "SI" else "FPS"
    out = N.create_input_template(name, unit_system=usys)
    print(green(f"  Template created: {out}  (units: {usys})"))
    print(dim("  Fill the Pipeline_Input sheet (and Flow_Cases / Stream_Map / Instrument_Data if used), then run option 1."))
    print(dim("  Flip FPS/SI any time on the workbook's UNITS sheet."))


def default_hmb_name(dump: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9_-]", "_", os.path.splitext(os.path.basename(dump))[0])[:55]
    return f"HMB_{stem}.xlsx"


def build_hmb_action(dump: str | None = None, out: str | None = None,
                     sheet: str | None = None, constants: str | None = None,
                     interactive: bool = True) -> str | None:
    """PRO/II dump  →  HMB workbook (hmbbuilder.py).  Returns the HMB path (None if cancelled)."""
    try:
        import hmbbuilder as B
    except ImportError:
        print(red("  hmbbuilder.py was not found next to this launcher."))
        return None
    if dump is None:
        dumps = [f for f in list_xlsx() if kind_of(f) == "dump"]
        pick = choose_from("PRO/II dump workbook:", dumps, allow_blank=True,
                           extra=[("p", "type a path to the dump")])
        if pick is None:
            return None
        dump = ask_path("Path to the PRO/II dump") if pick == "p" else pick
        if not dump:
            return None
    if out is None:
        out = default_hmb_name(dump)
        if interactive:
            out = ask("HMB file to write", default=out)
            if os.path.exists(out) and ask(f"{out} exists — rebuild it? (y/n)", default="y").lower() not in ("y", "yes"):
                print(green(f"  Using the existing {out}"))
                return out
            sheet = sheet or (ask("Dump sheet name (Enter = auto-detect)") or None)
            if constants is None:
                constants = ask("Optional Constants.xlsx for pseudo-component overrides (Enter = none)") or None
    print(dim("  Building the HMB (about 80 s for ~460 streams) …"))
    try:
        res = B.build_hmb(dump, out, sheet, constants)
    except SystemExit as exc:
        print(red(f"  Stopped: {exc}"))
        return None
    except Exception:
        print(red("  Error while building the HMB:"))
        traceback.print_exc()
        return None
    print(green(f"  HMB ready: {res}"))
    return res


def choose_theme() -> str:
    t = choose_from("Output workbook colours:",
                    ["coloured  (default — dark banners, original colour coding)",
                     "light  (Technip Energies light theme + TEN logo)"])
    return "light" if t.startswith("light") else "colored"


def run_engine():
    files = list_xlsx()
    inputs, others = categorise(files)
    hmbs = [f for f in others if kind_of(f) == "hmb"]
    dumps = [f for f in others if kind_of(f) == "dump"]

    if not inputs:
        print(red("  No input workbook (with a 'Pipeline_Input' sheet) found here."))
        print(dim("  Create one from the template first (menu option 3)."))
        return

    input_path = (inputs[0] if len(inputs) == 1
                  else choose_from("Input workbook:", inputs))
    if len(inputs) == 1:
        print(green(f"  Input: {input_path}"))

    # HMB: an existing one, or build it from a PRO/II dump right now.  Optional for manual circuits.
    hmb_path = None
    labels = {}
    for f in hmbs:
        labels[f"{f}   (HMB workbook)"] = ("hmb", f)
    for f in dumps:
        labels[f"{f}   (PRO/II dump → build the HMB first)"] = ("dump", f)
    if labels:
        pick = choose_from("HMB for this run (Enter to skip if all circuits are manual):",
                           list(labels), allow_blank=True,
                           extra=[("p", "type a path to an HMB or a PRO/II dump"),
                                  ("n", "no HMB (manual only)")])
        kind = path = None
        if pick == "p":
            path = ask_path("Path to the HMB or dump")
            kind = kind_of(path) if path else None
            kind = "dump" if kind == "dump" else "hmb"
        elif pick not in (None, "n"):
            kind, path = labels[pick]
        if kind == "dump":
            hmb_path = build_hmb_action(path)
            if hmb_path is None:
                return
        elif kind == "hmb":
            hmb_path = path
    else:
        print(dim("  No HMB workbook or PRO/II dump found here — manual-property circuits only."))
        p_ = ask("Type a path to an HMB / dump, or press Enter to continue without one")
        if p_.strip().strip('"'):
            path = p_.strip().strip('"')
            if os.path.isfile(path):
                hmb_path = build_hmb_action(path) if kind_of(path) == "dump" else path

    case = choose_case(hmb_path) if hmb_path else "Case 1"
    circuit = choose_circuit(input_path)

    mode = choose_from("Flash mode:",
                       ["isothermal  (default)", "isenthalpic  (Joule-Thomson)"])
    flash_mode = "isenthalpic" if mode.startswith("isenthalpic") else "isothermal"
    theme = choose_theme()

    # ── summary ──
    print()
    rule()
    print(bold("  Run summary"))
    print(f"    Input   : {green(input_path)}")
    print(f"    HMB     : {green(hmb_path or '(none — manual)')}")
    print(f"    Case    : {green(case)}")
    print(f"    Circuit : {green(circuit or 'ALL')}")
    print(f"    Mode    : {green(flash_mode)}")
    print(f"    Colours : {green(theme)}")
    print(dim("    Flow cases (Normal / Max / Min) come from the Flow_Cases sheet of the input."))
    rule()
    if ask("Proceed? (y/n)", default="y").lower() not in ("y", "yes"):
        print(dim("  Cancelled."))
        return
    execute_run(input_path, hmb_path, circuit, case, flash_mode, theme)


def execute_run(input_path, hmb_path, circuit=None, case="Case 1",
                flash_mode="isothermal", theme="colored") -> bool:
    if hmb_path is None:
        # The engine needs a path argument; manual circuits resolve from the yellow cells.
        hmb_path = ""
    try:
        import hydraulics as N
        N.set_output_theme(theme)
        results = N.run_noiso(input_path, hmb_path,
                              circuit_id=circuit, case=case,
                              flash_mode=flash_mode)
        print()
        print(green(bold(f"  Done — {len(results)} workbook(s) written:")))
        for op, *_ in results:
            print(f"    • {op}")
        return True
    except SystemExit as exc:
        print(red(f"  Stopped: {exc}"))
    except Exception:
        print(red("  Error while running the engine:"))
        traceback.print_exc()
    return False


def build_and_run():
    """Menu option 3: build the HMB from a PRO/II dump, then run the hydraulics on it."""
    hmb = build_hmb_action()
    if hmb:
        print(dim("  Now the hydraulics run — pick the generated HMB in the next list."))
        run_engine()


# ── command line (no questions) ─────────────────────────────────────────────
def _opt(argv, name, flag=False):
    if name in argv:
        i = argv.index(name)
        if flag:
            return True, [a for j, a in enumerate(argv) if j != i]
        if i + 1 < len(argv):
            return argv[i + 1], [a for j, a in enumerate(argv) if j not in (i, i + 1)]
    return (False if flag else None), argv


def cli_main(argv: list[str]) -> int:
    cmd, rest = argv[0].lower(), argv[1:]
    sheet, rest = _opt(rest, "--sheet")
    consts, rest = _opt(rest, "--constants")
    case, rest = _opt(rest, "--case")
    iso, rest = _opt(rest, "--isenthalpic", flag=True)
    light, rest = _opt(rest, "--light", flag=True)
    args = [a for a in rest if not a.startswith("--")]
    mode = "isenthalpic" if iso else "isothermal"
    theme = "light" if light else "colored"
    case = case or "Case 1"
    if cmd == "build" and args:
        out = args[1] if len(args) > 1 else default_hmb_name(args[0])
        return 0 if build_hmb_action(args[0], out, sheet, consts, interactive=False) else 1
    if cmd == "run" and len(args) >= 2:
        return 0 if execute_run(args[0], args[1], args[2] if len(args) > 2 else None,
                                case, mode, theme) else 1
    if cmd == "all" and len(args) >= 2:
        hmb = build_hmb_action(args[0], default_hmb_name(args[0]), sheet, consts, interactive=False)
        if not hmb:
            return 1
        return 0 if execute_run(args[1], hmb, args[2] if len(args) > 2 else None,
                                case, mode, theme) else 1
    print(__doc__)
    return 2


def open_output_folder():
    if sys.platform == "darwin":
        os.system(f'open "{HERE}"')
    elif os.name == "nt":
        os.startfile(HERE)            # type: ignore[attr-defined]
    else:
        os.system(f'xdg-open "{HERE}" >/dev/null 2>&1')
    print(dim("  Opened the working folder."))


# ── main menu loop ──────────────────────────────────────────────────────────



def console_main():
    banner()
    actions = {
        "1": ("Run hydraulics (pick an HMB, or build it from a PRO/II dump first)", run_engine),
        "2": ("Build the HMB from a PRO/II dump  →  then run hydraulics", build_and_run),
        "3": ("Build the HMB from a PRO/II dump only", build_hmb_action),
        "4": ("Create blank input template", make_template),
        "5": ("Open this folder", open_output_folder),
        "q": ("Quit", None),
    }
    while True:
        print()
        print(bold("  Menu"))
        for key, (label, _) in actions.items():
            print(f"   {cyan(key)}  {label}")
        choice = ask("Select").lower()
        if choice in ("q", "quit", "exit", ""):
            print(dim("  Bye."))
            return
        entry = actions.get(choice)
        if not entry:
            print(red("  Unknown option."))
            continue
        try:
            entry[1]()
        except KeyboardInterrupt:
            print(dim("\n  (cancelled)"))


# ══════════════════════════════════════════════════════════════════════════
# ║  TKINTER GUI (former hydraulics_gui.py)
# ══════════════════════════════════════════════════════════════════════════
#!/usr/bin/env python3
"""
HyCalign — No-ISO Hydraulics: small GUI front end (Tkinter, stdlib only).

Pick an input workbook, an HMB workbook + case, a circuit and the flash mode
from drop-downs, press Run, and watch the log.  No command line needed.

Run with:   python3 hydraulics_gui.py
"""

import os
import sys
import glob
import queue
import threading
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)

try:
    import tkinter as tk
    from tkinter import ttk, messagebox
except Exception:                      # pragma: no cover
    tk = None                          # console mode still works without tkinter
    class _TtkStub:                    # lets the App class definition succeed
        Frame = object
    ttk = _TtkStub()
    messagebox = None


# ── discovery helpers ───────────────────────────────────────────────────────




def split_files():
    inputs, others = [], []
    for f in list_xlsx():
        (inputs if is_input_wb(f) else others).append(f)
    return inputs, others


def hmb_cases(path: str) -> list[str]:
    try:
        import hmbbuilder as H
        return [c for c in H.list_cases(path) if re.match(r"(?i)case", c)]
    except Exception:
        return []


def circuits_in(path: str) -> list[str]:
    try:
        import hydraulics as N
        rows = N.read_pipeline_input(path)
        return list(N._group_circuits(rows).keys())
    except Exception:
        return []


# ── stdout redirector → log queue ───────────────────────────────────────────
class QueueWriter:
    def __init__(self, q): self.q = q
    def write(self, s): self.q.put(s)
    def flush(self): pass


# ── the app ─────────────────────────────────────────────────────────────────
class App(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=14)
        self.grid(sticky="nsew")
        master.title("HyCalign — No-ISO Hydraulics")
        master.minsize(680, 520)
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(7, weight=1)

        self.q: queue.Queue = queue.Queue()
        self._build()
        self.refresh_files()
        self.after(80, self._drain_log)

    # -- layout -----------------------------------------------------------
    def _row(self, r, label, widget):
        ttk.Label(self, text=label).grid(row=r, column=0, sticky="w", pady=4, padx=(0, 10))
        widget.grid(row=r, column=1, sticky="ew", pady=4)

    def _build(self):
        head = ttk.Label(self, text="No-ISO Flash Hydraulics",
                         font=("Helvetica", 15, "bold"))
        head.grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(self, text=HERE, foreground="#666").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(0, 8))

        self.input_cb = ttk.Combobox(self, state="readonly")
        self.input_cb.bind("<<ComboboxSelected>>", lambda e: self.reload_circuits())
        self._row(2, "Input workbook", self.input_cb)

        self.hmb_cb = ttk.Combobox(self, state="readonly")
        self.hmb_cb.bind("<<ComboboxSelected>>", lambda e: self.reload_cases())
        self._row(3, "HMB workbook", self.hmb_cb)

        self.case_cb = ttk.Combobox(self)          # editable → type to filter
        self._row(4, "Case", self.case_cb)

        self.circuit_cb = ttk.Combobox(self, state="readonly")
        self._row(5, "Circuit", self.circuit_cb)

        self.mode_cb = ttk.Combobox(self, state="readonly",
                                    values=["isothermal", "isenthalpic"])
        self.mode_cb.set("isothermal")
        self._row(6, "Flash mode", self.mode_cb)

        # log box
        logwrap = ttk.Frame(self)
        logwrap.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=(10, 8))
        logwrap.columnconfigure(0, weight=1)
        logwrap.rowconfigure(0, weight=1)
        self.log = tk.Text(logwrap, height=14, wrap="word",
                           bg="#0f1b2d", fg="#d6e2f5", insertbackground="#d6e2f5",
                           font=("Menlo", 10), relief="flat")
        self.log.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(logwrap, command=self.log.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.log["yscrollcommand"] = sb.set

        # buttons
        btns = ttk.Frame(self)
        btns.grid(row=8, column=0, columnspan=3, sticky="ew")
        btns.columnconfigure(3, weight=1)
        ttk.Button(btns, text="Refresh", command=self.refresh_files).grid(row=0, column=0)
        ttk.Button(btns, text="New template", command=self.make_template).grid(row=0, column=1, padx=6)
        ttk.Button(btns, text="Open folder", command=self.open_folder).grid(row=0, column=2)
        self.run_btn = ttk.Button(btns, text="Run  ▶", command=self.run)
        self.run_btn.grid(row=0, column=4, sticky="e")

    # -- data wiring ------------------------------------------------------
    def refresh_files(self):
        inputs, others = split_files()
        self.input_cb["values"] = inputs
        self.hmb_cb["values"] = others
        if inputs and not self.input_cb.get():
            self.input_cb.set(inputs[0])
        if others and not self.hmb_cb.get():
            # prefer a file literally named like an HMB
            pref = next((o for o in others if "hmb" in o.lower()), others[0])
            self.hmb_cb.set(pref)
        self.reload_cases()
        self.reload_circuits()
        self._logln(f"Found {len(inputs)} input file(s), {len(others)} other workbook(s).")

    def reload_cases(self):
        path = self.hmb_cb.get()
        cases = hmb_cases(path) if path else []
        self.case_cb["values"] = cases
        if not self.case_cb.get():
            self.case_cb.set("Case 1")
        if cases:
            self._logln(f"{os.path.basename(path)}: {len(cases)} case(s) available.")

    def reload_circuits(self):
        path = self.input_cb.get()
        cids = circuits_in(path) if path else []
        vals = ["ALL"] + cids
        self.circuit_cb["values"] = vals
        self.circuit_cb.set("ALL")
        if cids:
            self._logln(f"{os.path.basename(path)}: circuits {', '.join(cids)}.")

    # -- actions ----------------------------------------------------------
    def make_template(self):
        try:
            import hydraulics as N
            usys = "SI" if messagebox.askyesno(
                "Units", "Create the template in SI units?\n"
                "(No = FPS.  You can flip FPS/SI later on the UNITS sheet.)") \
                else "FPS"
            out = N.create_input_template("pipeline_input_noiso.xlsx",
                                          unit_system=usys)
            self._logln(f"Template created: {out}  (units: {usys})")
            self.refresh_files()
            messagebox.showinfo("Template", f"Created {out}  (units: {usys})\n"
                                "Fill the Pipeline_Input sheet, then Run.")
        except Exception:
            self._logln(traceback.format_exc())

    def open_folder(self):
        if sys.platform == "darwin":
            os.system(f'open "{HERE}"')
        elif os.name == "nt":
            os.startfile(HERE)        # type: ignore[attr-defined]
        else:
            os.system(f'xdg-open "{HERE}" >/dev/null 2>&1')

    def run(self):
        input_path = self.input_cb.get()
        if not input_path:
            messagebox.showwarning("No input", "Choose an input workbook first.")
            return
        hmb_path = self.hmb_cb.get() or "HMB.xlsx"
        case = self.case_cb.get() or "Case 1"
        circuit = self.circuit_cb.get()
        circuit = None if circuit in ("", "ALL") else circuit
        flash_mode = self.mode_cb.get() or "isothermal"

        self.run_btn["state"] = "disabled"
        self.log.delete("1.0", "end")
        self._logln(f"Running  input={input_path}  hmb={hmb_path}  "
                    f"case={case}  circuit={circuit or 'ALL'}  mode={flash_mode}")
        self._logln("─" * 60)

        t = threading.Thread(target=self._run_worker,
                             args=(input_path, hmb_path, circuit, case, flash_mode),
                             daemon=True)
        t.start()

    def _run_worker(self, input_path, hmb_path, circuit, case, flash_mode):
        old = sys.stdout
        sys.stdout = QueueWriter(self.q)
        try:
            import hydraulics as N
            results = N.run_noiso(input_path, hmb_path, circuit_id=circuit,
                                  case=case, flash_mode=flash_mode)
            self.q.put(f"\nDone — {len(results)} workbook(s):\n")
            for op, *_ in results:
                self.q.put(f"  • {op}\n")
        except SystemExit as exc:
            self.q.put(f"\nStopped: {exc}\n")
        except Exception:
            self.q.put("\n" + traceback.format_exc())
        finally:
            sys.stdout = old
            self.q.put("__DONE__")

    # -- log plumbing -----------------------------------------------------
    def _logln(self, s):
        self.log.insert("end", s + "\n")
        self.log.see("end")

    def _drain_log(self):
        try:
            while True:
                s = self.q.get_nowait()
                if s == "__DONE__":
                    self.run_btn["state"] = "normal"
                    self.refresh_files()
                else:
                    self.log.insert("end", s)
                    self.log.see("end")
        except queue.Empty:
            pass
        self.after(80, self._drain_log)





def gui_main():
    if tk is None:
        sys.stderr.write("Tkinter is not available in this Python build.\n"
                         "Run the console menu instead:  python3 run_hydraulics.py\n")
        sys.exit(1)
    root = tk.Tk()
    try:
        ttk.Style().theme_use("clam")
    except Exception:
        pass
    App(root)
    root.mainloop()



def main():
    argv = sys.argv[1:]
    if argv and argv[0].lower() in ("build", "run", "all"):
        sys.exit(cli_main(argv))
    if any(a.lower() in ("--gui", "-g") for a in argv):
        gui_main()
    else:
        console_main()


if __name__ == "__main__":
    main()
