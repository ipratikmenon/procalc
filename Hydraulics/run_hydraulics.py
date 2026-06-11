#!/usr/bin/env python3
"""HyCalign Hydraulics — single front end (console menu by default, --gui for window).

    python run_hydraulics.py          # interactive console menu
    python run_hydraulics.py --gui    # tkinter GUI

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


def categorise(files: list[str]) -> tuple[list[str], list[str]]:
    inputs, others = [], []
    for f in files:
        (inputs if is_input_wb(f) else others).append(f)
    return inputs, others


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
    try:
        import hmb_proii_reader as H
        cases = H.list_cases(hmb_path)
    except Exception as exc:
        print(dim(f"  (could not read cases from {hmb_path}: {exc})"))
        cases = []

    if not cases:
        return ask("HMB case", default="Case 1")

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
    print(dim("  Fill the Pipeline_Input sheet, then run option 1."))
    print(dim("  Flip FPS/SI any time on the workbook's UNITS sheet."))


def run_engine():
    files = list_xlsx()
    inputs, others = categorise(files)

    if not inputs:
        print(red("  No input workbook (with a 'Pipeline_Input' sheet) found here."))
        print(dim("  Use option 2 to create one first."))
        return

    input_path = (inputs[0] if len(inputs) == 1
                  else choose_from("Input workbook:", inputs))
    if len(inputs) == 1:
        print(green(f"  Input: {input_path}"))

    # HMB is optional (manual circuits need none) — offer a 'none' choice.
    hmb_path = None
    if others:
        pick = choose_from("HMB workbook (Enter to skip if all circuits are manual):",
                           others, allow_blank=True,
                           extra=[("n", "no HMB (manual only)")])
        if pick not in (None, "n"):
            hmb_path = pick
    else:
        print(dim("  No HMB workbook found — manual-property circuits only."))

    case = choose_case(hmb_path) if hmb_path else "Case 1"
    circuit = choose_circuit(input_path)

    mode = choose_from("Flash mode:",
                       ["isothermal  (default)", "isenthalpic  (Joule-Thomson)"])
    flash_mode = "isenthalpic" if mode.startswith("isenthalpic") else "isothermal"

    # ── summary ──
    print()
    rule()
    print(bold("  Run summary"))
    print(f"    Input   : {green(input_path)}")
    print(f"    HMB     : {green(hmb_path or '(none — manual)')}")
    print(f"    Case    : {green(case)}")
    print(f"    Circuit : {green(circuit or 'ALL')}")
    print(f"    Mode    : {green(flash_mode)}")
    rule()
    if ask("Proceed? (y/n)", default="y").lower() not in ("y", "yes"):
        print(dim("  Cancelled."))
        return

    if hmb_path is None:
        # The engine needs a path argument; pass a placeholder that won't be
        # used because manual circuits resolve from the yellow cells.
        hmb_path = next((f for f in others), "HMB.xlsx")

    try:
        import hydraulics as N
        results = N.run_noiso(input_path, hmb_path,
                              circuit_id=circuit, case=case,
                              flash_mode=flash_mode)
        print()
        print(green(bold(f"  Done — {len(results)} workbook(s) written:")))
        for op, *_ in results:
            print(f"    • {op}")
    except SystemExit as exc:
        print(red(f"  Stopped: {exc}"))
    except Exception:
        print(red("  Error while running the engine:"))
        traceback.print_exc()


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
        "1": ("Run hydraulics", run_engine),
        "2": ("Create blank input template", make_template),
        "3": ("Open this folder", open_output_folder),
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
        import hmb_proii_reader as H
        return H.list_cases(path)
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
    if any(a.lower() in ("--gui", "-g") for a in sys.argv[1:]):
        gui_main()
    else:
        console_main()


if __name__ == "__main__":
    main()
