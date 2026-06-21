#!/usr/bin/env python3
"""
Merge PRO/II "Custom Properties" exports (.xls) into one component property
lookup keyed by name (upper-case). Source files are direct PRO/II library
exports (Fixed-tab constants), not OCR'd or reverse-engineered.

Usage:
    python build_proii_db.py
Produces proii_props.json in this directory.
"""
import os, json
import xlrd

HERE = os.path.dirname(os.path.abspath(__file__))

# Later sources win on name collisions (REFORMER/HFALKYLATION reactor cuts
# are more specific to those units than the general PROCESS library).
SOURCES = ["PROCESSprop.xls", "LNGprop.xls", "REFORMERprop.xls", "HFALKYLATIONprop.xls"]

COLUMNS = [
    "MW", "NMP", "NBP", "SG60F", "MVOL25C", "TC", "PC", "VC", "ZC",
    "RADIUS", "DIELECTRIC", "UNIFAC_Q", "UNIFAC_R", "GFORMATION",
    "HFORMATION", "HCOMBUST", "HFUSIONNMP", "HVAPNBP", "ACENTRIC",
    "SOLUPARA", "RACKETT", "DIPOLE", "FLASHPOINT", "LOFLAMM", "HIFLAMM",
    "AUTOIGNITION", "GHV", "LHV", "TTP", "PTP", "CNUM", "ZNUM",
]


def _val(v):
    if v == "" or v is None:
        return None
    return v


def load_source(fname):
    wb = xlrd.open_workbook(os.path.join(HERE, fname))
    sh = wb.sheet_by_index(0)
    out = {}
    for r in range(2, sh.nrows):
        name = str(sh.cell_value(r, 0)).strip()
        if not name:
            continue
        row = {col: _val(sh.cell_value(r, c + 1)) for c, col in enumerate(COLUMNS)}
        row["SOURCE"] = fname
        out[name.upper()] = row
    return out


def build():
    merged = {}
    for fname in SOURCES:
        for key, row in load_source(fname).items():
            merged[key] = row
    return merged


if __name__ == "__main__":
    db = build()
    out_path = os.path.join(HERE, "proii_props.json")
    with open(out_path, "w") as f:
        json.dump(db, f, indent=1, sort_keys=True)
    print(f"{len(db)} components -> {out_path}")
