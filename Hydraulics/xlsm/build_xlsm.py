"""Build the Procalc macro workbook skeleton (sheets + data tables [+ VBA project]).

  python build_xlsm.py [--hmb HMB.xlsx] [--out Procalc_Hydraulics.xlsm] [--vba vbaProject.bin]

Everything the VBA needs (fitting K, ASME pipe tables, PMS classes) is written to the Data_* sheets, so
the macro code itself stays small.  `--hmb` fills the Streams sheet from a per-stream-sheet HMB workbook.
"""
import argparse
import os
import sys
import xlsxwriter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import hydraulics as H                      # noqa: E402
import openpyxl                             # noqa: E402

STREAM_KEYS = ["phase", "temp_f", "pres_psia", "total_mass", "vap_mass", "liq_mass", "mol_weight", "vap_mw",
               "vap_density", "liq_density", "vap_visc", "liq_visc", "vap_z", "total_z", "vap_cp_cv",
               "vap_cp", "liq_cp", "liq_surf_tens"]


def sample_rows():
    """Example circuits: L-810 liquid line with flow-control CV, and the pump back-calculation case."""
    rows, hdr = [], None
    for tag, f in (("C1", "pipeline_input_L810.xlsx"), ("C2", "pipeline_input_L810_pump.xlsx")):
        ws = openpyxl.load_workbook(os.path.join(os.path.dirname(HERE), f))["Pipeline_Input"]
        it = list(ws.iter_rows(min_row=2, values_only=True))
        hdr = it[0]
        for r in it[1:]:
            if r[1]:
                r = list(r)
                r[0] = tag
                rows.append(r)
    return hdr, rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--hmb")
    ap.add_argument("--out", default=os.path.join(HERE, "Procalc_Hydraulics.xlsm"))
    ap.add_argument("--vba")
    a = ap.parse_args()
    wb = xlsxwriter.Workbook(a.out)
    NSHEETS = 20
    if a.out.lower().endswith(".xlsm"):
        import tempfile
        import ovba
        mods = [(f"Sheet{i}", "sheet", "") for i in range(1, NSHEETS + 1)]
        mods.append(("Procalc", "std", open(os.path.join(HERE, "Procalc.bas"), encoding="utf-8").read()))
        vb = a.vba or os.path.join(tempfile.gettempdir(), "procalc_vbaProject.bin")
        if not a.vba:
            open(vb, "wb").write(ovba.make_vba_project(mods))
        a.vba = vb
        wb.add_vba_project(vb)
        wb.set_vba_name("ThisWorkbook")
    _orig_add = wb.add_worksheet

    def _add(name):
        ws = _orig_add(name)
        if a.vba:
            ws.set_vba_name(f"Sheet{len(wb.worksheets())}")
        return ws
    wb.add_worksheet = _add
    navy = wb.add_format({"bg_color": "#001742", "font_color": "white", "bold": True, "font_name": "Arial",
                          "font_size": 9, "text_wrap": True, "valign": "vcenter", "border": 1})
    yel = wb.add_format({"bg_color": "#FFF2CC", "font_name": "Arial", "font_size": 11, "bold": True, "border": 1})
    ttl = wb.add_format({"bold": True, "font_size": 16, "font_color": "#001742", "font_name": "Arial"})
    txt = wb.add_format({"font_name": "Arial", "font_size": 9, "text_wrap": True, "valign": "top"})
    lab = wb.add_format({"font_name": "Arial", "font_size": 10, "bold": True})

    # ── Control ──────────────────────────────────────────────────────────
    c = wb.add_worksheet("Control")
    c.hide_gridlines(2)
    c.set_column("A:A", 3); c.set_column("B:B", 30); c.set_column("C:C", 60)
    c.write("B1", "PROCALC  -  Hydraulics (macro workbook)", ttl)
    c.write("B2", "Edit the Input sheet, then press RUN.  Results are rebuilt in the output sheets (Pressure_Profile, CV_Summary, FIV, AIV, regime, instruments, composition).", txt)
    c.write("B4", "Flow case multiplier (%)", lab); c.write("C4", 100, yel)
    c.write("B5", "Streams used", lab)
    c.write("C5", "A sheet named like the stream (HMB per-stream sheet, with composition) is flashed; else the Streams table row is used.", txt)
    c.write("B7", "Flash from composition (ON/OFF)", lab); c.write("C7", "ON", yel)
    c.write("B9", "Isenthalpic T march, PR-covered streams (ON/OFF)", lab); c.write("C9", "OFF", yel)
    c.write("B8", "Last run", lab); c.write("C8", "not run yet", txt)
    if a.vba:
        c.insert_button("B6", {"macro": "RunHydraulics", "caption": "RUN HYDRAULICS", "width": 260, "height": 34})
    c.set_row(5, 28)
    notes = [
        "Units: FPS (psia, degF, lb/hr, ft, in, cP, lb/ft3) - the same as hydraulics.py.",
        "Input sheet: same 37 columns as Pipeline_Input.  Header row = the row with 'Circuit' in column A (row 1-3);",
        "unit suffixes in headers are ignored.  Rows sharing a Circuit ID are chained in series.",
        "Stream Lookup -> a sheet named like the stream (HMB per-stream sheet, flashed) or the Streams table; blank -> the manual-property columns.",
        "Pump discharge: Source row without Set P + Destination with Set P => Source P is back-calculated.",
        "Control valve: Fixed dP > Set P (outlet) > flow-control floats to Destination > Fixed K.",
        "Component_Constants (HMB COMP_CONSTANTS) feeds the flash. Run Type = Branch plus Tee Split / Join rows give branch runs.",
    ]
    for i, n in enumerate(notes):
        c.merge_range(10 + i, 1, 10 + i, 2, n, txt)

    # ── Input ────────────────────────────────────────────────────────────
    hdr, rows = sample_rows()
    w = wb.add_worksheet("Input")
    w.write(0, 0, "PIPELINE INPUT  -  rows with the same Circuit are chained in series (FPS units)", ttl)
    w.set_row(1, 36)
    for j, h in enumerate(hdr):
        w.write(1, j, h, navy)
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            if v is not None:
                w.write(2 + i, j, v)
    w.set_column(0, 0, 8); w.set_column(1, 1, 9); w.set_column(2, 12, 7); w.set_column(13, 26, 9)
    w.set_column(27, 28, 22); w.set_column(29, 41, 10); w.set_column(42, 42, 50)
    w.freeze_panes(2, 2)

    # ── Streams ──────────────────────────────────────────────────────────
    s = wb.add_worksheet("Streams")
    s.write_row(0, 0, ["Stream"] + STREAM_KEYS, navy)
    wanted = {str(r[4]).upper() for r in rows if r[4]}
    hmb = a.hmb or os.path.join(os.path.dirname(HERE), "HMB.xlsx")
    found = {}
    wbh = openpyxl.load_workbook(hmb, read_only=True, data_only=True)
    for wsh in wbh.worksheets:                          # single pass over the HMB workbook
        if wsh.title in H.SKIP_SHEETS:
            continue
        nm, _ = H._stream_title(wsh)
        if not nm or (a.hmb is None and nm.upper() not in wanted):
            continue
        found[nm] = H._extract_from_sheet(wsh)
        if a.hmb is None and len(found) == len(wanted):
            break
    wbh.close()
    ri = 1
    for nm in sorted(found):
        sp = found[nm]
        s.write(ri, 0, nm)
        for j, k in enumerate(STREAM_KEYS):
            v = sp.phase if k == "phase" else getattr(sp, k, None)
            if v is not None:
                s.write(ri, 1 + j, v)
        ri += 1
    print("Streams sheet:", len(found), "streams from", os.path.basename(hmb))
    s.set_column(0, 0, 18); s.set_column(1, 20, 12); s.freeze_panes(1, 1)

    # ── HMB sheets pasted by the user: COMP_CONSTANTS + per-stream sheets (sample: DX10-OUT) ─────────
    smp = openpyxl.load_workbook(os.path.join(HERE, "sample_DX10-OUT.xlsx"), data_only=True)

    def copy_sheet(src, dst, head_rows):
        for ri, row in enumerate(src.iter_rows(values_only=True)):
            for ci, v in enumerate(row):
                if v is not None:
                    dst.write(ri, ci, v, navy if ri in head_rows else None)
    cc = wb.add_worksheet("Component_Constants")
    copy_sheet(smp["COMP_CONSTANTS"], cc, (1,))
    cc.set_column(1, 1, 18); cc.set_row(1, 30)
    for nm_ in ("DX10-OUT",):
        w_ = wb.add_worksheet(nm_)
        copy_sheet(smp[nm_], w_, (1,))
        w_.set_column(1, 1, 38); w_.set_column(2, 5, 14)

    # ── Data sheets ──────────────────────────────────────────────────────
    d = wb.add_worksheet("Data_K"); d.write_row(0, 0, ["Fitting (lower case)", "K"], navy)
    for i, (k, v) in enumerate(sorted(H._K_TABLE.items())):
        d.write(1 + i, 0, k); d.write(1 + i, 1, v)
    d.set_column(0, 0, 44)
    d = wb.add_worksheet("Data_OD"); d.write_row(0, 0, ["NPS", "NPS (number)", "OD (in)"], navy)
    for i, n in enumerate(H.A.NPS_ORDER):
        num = H._NPS_BY_FLOAT_REV[n] if hasattr(H, "_NPS_BY_FLOAT_REV") else {v: k for k, v in H._NPS_BY_FLOAT.items()}[n]
        d.write(1 + i, 0, n); d.write(1 + i, 1, num); d.write(1 + i, 2, H.A.PIPE_OD_IN[n])
    d = wb.add_worksheet("Data_Pipe"); d.write_row(0, 0, ["NPS", "Schedule", "Wall (in)"], navy)
    i = 1
    for n in H.A.NPS_ORDER:
        for sch, wl in H.A.PIPE_WALL_IN.get(n, {}).items():
            d.write(i, 0, n); d.write(i, 1, str(sch).upper()); d.write(i, 2, wl); i += 1
    d = wb.add_worksheet("Data_PMS"); d.write_row(0, 0, ["Class", "NPS low", "NPS high", "Schedule", "Min schedule", "MOC"], navy)
    i = 1
    for name, (moc, rules) in H.PMS_CLASSES.items():
        for r in rules:
            d.write(i, 0, name); d.write(i, 1, r[0]); d.write(i, 2, r[1]); d.write(i, 3, str(r[2]))
            d.write(i, 4, r[3] if len(r) > 3 else "STD"); d.write(i, 5, moc); i += 1
    idn = wb.add_worksheet("Instrument_Data")
    idn.write_row(1, 0, H._INSTR_DATA_HEADERS, navy)
    idn.write(0, 0, "Optional per-tag data: TT thermowell geometry, FO/RO orifice d, control-valve rated Cv / FL / xT / characteristic / curve points ('10:4.2; 25:11; ...').", txt)
    ex = [("TT-101", "Reactor effluent TT", None, 6.0, 1.0, 0.75, 0.26, "SS316", None, None, None, None, None, None, "example"),
          ("FE-101", "Flow orifice", 2.5, None, None, None, None, None, None, None, None, None, None, None, "example"),
          ("FV-101", "Flow control valve", None, None, None, None, None, None, 20, 0.9, 0.7, "Equal %", 50, None, "example")]
    for i, rw in enumerate(ex):
        for j, v in enumerate(rw):
            if v is not None:
                idn.write(2 + i, j, v)
    idn.set_column(0, 14, 13); idn.set_row(1, 40)
    tw = wb.add_worksheet("Data_TWClass")
    tw.write_row(0, 0, ["Piping class", "TW root OD (in)", "TW tip OD (in)", "TW bore (in)", "TW insertion U (in)", "TW material"], navy)
    tw.write_row(1, 0, ["G9A-3", 1.0, 0.75, 0.26, 6.0, "SS316"])
    tw.set_column(0, 5, 16)
    pr = wb.add_worksheet("Data_PR")
    pr.write_row(0, 0, ["Component", "MW", "Tc (K)", "Pc (bar)", "omega", "Vc (cm3/mol)", "H offset (kJ/kmol)", "Hf (kJ/kmol)",
                        "H poly a", "b", "c", "d", "e", "f", "Peneloux c (cm3/mol)"], navy)
    for i, (k, d) in enumerate(H._PR_DATA.items()):
        pr.write_row(1 + i, 0, [k, d[0], d[1], d[2], d[3], d[4], d[5], d[6], *d[7], d[8]])
    pr.set_column(0, 0, 18)
    pa = wb.add_worksheet("Data_PRAlias")
    pa.write_row(0, 0, ["Normalised name", "PR component"], navy)
    for i, (k, v) in enumerate(H._PR_ALIAS.items()):
        pa.write_row(1 + i, 0, [k, v])
    pk = wb.add_worksheet("Data_PRKij")
    pk.write_row(0, 0, ["Component 1", "Component 2", "kij"], navy)
    for i, ((a_, b_), v) in enumerate(H._PR_KIJ.items()):
        pk.write_row(1 + i, 0, [a_, b_, v])
    wb.close()
    print("wrote", a.out, os.path.getsize(a.out), "bytes")


if __name__ == "__main__":
    main()
