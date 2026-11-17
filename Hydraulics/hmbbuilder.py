#!/usr/bin/env python3
"""
hmbbuilder.py  —  HyCalign / Procalc  ·  PRO/II dump  →  HMB workbook for hydraulics.py
=======================================================================================
ONE file that replaces  td_parser.py · hmb_proii_reader.py · comp_constants.py ·
hmb_dump_parser.py.

    python hmbbuilder.py BTGOLD_HCU_CASE2.xlsx                 # → HMB_BTGOLD_HCU_CASE2_<ts>.xlsx
    python hmbbuilder.py DUMP.xlsx HMB.xlsx --sheet CASE2      # explicit output / dump sheet
    python hmbbuilder.py DUMP.xlsx --constants Constants.xlsx  # optional pseudo-component overrides
    python hmbbuilder.py compare HMB_a.xlsx HMB_b.xlsx ...     # multi-case comparison workbook
    python hmbbuilder.py wide DUMP.xlsx [OUT.xlsx]             # clean wide-table 'Case N' workbook

What it reads
-------------
The transposed PRO/II property dump: column A = property label, column B = unit,
columns C… = one stream each (stream name on the 'Stream Name' row).  Every value is
found by its LABEL (never by a fixed row number), so a dump with extra / re-ordered
rows still lands in the right property.

What it writes  (the HMB format hydraulics.py reads)
---------------------------------------------------
    UNITS · COMPONENTS · COMP_CONSTANTS · INPUT · OUTPUT · one sheet per stream
COMP_CONSTANTS is built from the dump itself:
    pure components   embedded databank (Tc, Pc, ω, Vc, Peneloux, parachor …)
    pseudo-components MW and standard liquid density measured from the dump's own component
                      flows; NBP from the cut name (PF<NBP °F>A<API>D) or a Søreide estimate;
                      Tc / Pc / Vc / ω by Twu (1984) + PRO/II SIMSCI/TWU Frost-Kalkwarf-Thodos ω
"""
from __future__ import annotations

import datetime as _dt
import glob
import math
import os
import re
import statistics
import sys
import tempfile
from collections import OrderedDict, defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ════════════════════════════════════════════════════════════════════════
#  Shared styling / helpers
# ════════════════════════════════════════════════════════════════════════
C = dict(
    NAVY="1F4973", STEEL="2E75B6", WHITE="FFFFFF", LGRAY="F2F2F2",
    MGRAY="A0A0A0", DGRAY="404040", ORANGE="C55A11", AMBER="FFF2CC",
    GREEN="E2EFDA", GRNHDR="375623", RED="FCE4D6", LTBLUE="DEEAF1",
    PURPLE="7030A0", TEAL="1F7070", WARN="FFF0C0", FORMULA="FFFDE7",
    PH_TOT="D6E4F7", PH_VAP="DEEAF1", PH_LIQ="E2EFDA",
    PH_AQ="FFF2CC", PH_WAT="E8F4FD", PH_DRY="F2F2F2",
    PURE="D6E4F7", PSEUDO="FFF2CC", ESTIM="FFFDE7",
)
PHASE_BG = {'TOTAL': C['PH_TOT'], 'VAPOR': C['PH_VAP'], 'LIQUID': C['PH_LIQ'],
            'WATER': C['PH_WAT'], 'DRY': C['PH_DRY']}
TABS = ["1F4973", "2E75B6", "375623", "C55A11", "7030A0", "1F7070",
        "808080", "C00000", "70AD47", "1F4973", "2E75B6", "375623",
        "C55A11", "7030A0", "1F7070", "808080"]
CASE_PALETTE = ["1F4973", "2E75B6", "375623", "C55A11", "7030A0", "1F7070",
                "808080", "C00000", "70AD47", "5B2C6F", "1A5276", "117A65"]


_STYLE_CACHE: dict = {}


def _fill(h):
    k = ("fill", h)
    v = _STYLE_CACHE.get(k)
    if v is None:
        v = _STYLE_CACHE[k] = PatternFill("solid", fgColor=h.lstrip('#'))
    return v


def _thin():
    v = _STYLE_CACHE.get("thin")
    if v is None:
        s = Side(style="thin", color="D0D0D0")
        v = _STYLE_CACHE["thin"] = Border(left=s, right=s, top=s, bottom=s)
    return v


def _c(ws, r, c, v=None, bg=None, fg=None, sz=9, bold=False, italic=False,
       wrap=False, ha='left', va='center', border=True):
    cl = ws.cell(row=r, column=c)
    if isinstance(cl, MergedCell):
        return cl
    if v is not None:
        cl.value = v
    k = ("font", sz, bold, italic, fg or C['DGRAY'])
    f = _STYLE_CACHE.get(k)
    if f is None:
        f = _STYLE_CACHE[k] = Font(name="Calibri", size=sz, bold=bold, italic=italic, color=(fg or C['DGRAY']))
    cl.font = f
    if bg:
        cl.fill = _fill(bg)
    k = ("al", ha, va, wrap)
    a_ = _STYLE_CACHE.get(k)
    if a_ is None:
        a_ = _STYLE_CACHE[k] = Alignment(horizontal=ha, vertical=va, wrap_text=wrap)
    cl.alignment = a_
    if border:
        cl.border = _thin()
    return cl


cell = _c


def cw(ws, col, w):
    ws.column_dimensions[get_column_letter(col) if isinstance(col, int) else col].width = w


def rh(ws, row, h):
    ws.row_dimensions[row].height = h


def sf(v):
    """Safe float — None if not numeric ('n/a', blanks, text, NaN)."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return None if (isinstance(v, float) and math.isnan(v)) else float(v)
    s = str(v).strip()
    if not s or s.lower() in ("n/a", "na", "nan", "none", "missing", "-", "--"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def fmtv(v, dp=6):
    if v is None or v == '':
        return ''
    if isinstance(v, str):
        return v
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return ''
        if v == 0.0:
            return 0.0
        a = abs(v)
        if a < 1e-3 or a >= 1e8:
            return float(f'{v:.{dp}E}')
        return round(v, dp)
    return v


fmt = fmtv


def _writable(d):
    try:
        tp = os.path.join(d, '.hmb_wtest')
        open(tp, 'w').close()
        os.remove(tp)
        return True
    except OSError:
        return False


def _out_folder(*paths):
    for p in paths:
        d = os.path.dirname(os.path.abspath(p))
        if _writable(d):
            return d
    for d in [os.getcwd(), os.path.expanduser('~'), tempfile.gettempdir()]:
        if _writable(d):
            return d
    return tempfile.gettempdir()


# ════════════════════════════════════════════════════════════════════════
#  1.  LABEL TABLES  (from the PRO/II dump)
# ════════════════════════════════════════════════════════════════════════
# Exact column-A label → internal key (first occurrence wins).
_EXACT = {
    "Stream Phase": "phase", "Temperature": "temp_f", "Pressure": "pres_psia",
    "Total Molar Rate": "total_molar", "Total Mass Rate": "total_mass",
    "Total Molecular Weight": "mol_weight", "Total Enthalpy": "total_enthalpy",
    "Total Sp. Enthalpy": "total_sp_enthalpy", "Total Act. Vol. Rate": "total_act_vol",
    "Liquid Mole Fraction": "liq_mole_frac", "Vapor Mole Fraction": "vap_mole_frac",
    "Liquid Weight Fraction": "liq_wt_frac", "Total Sp. Entropy": "total_sp_entropy",
    "True Critical Temperature": "tc_f", "True Critical Pressure": "pc_psia",
    "True Critical Volume": "tc_vol",
    "Total Crit. Temperature (Kay's Rule)": "tc_kays",
    "Total Crit. Pressure (Kay's Rule)": "pc_kays",
    "Total Acentric Factor": "acentric", "Total UOP K factor": "watson_k",
    "Total Std. Liq. Density": "total_std_liq_density", "Total Std. API": "total_std_api",
    "Total Actual Density": "total_density", "Total Z (from actual density)": "total_z",
    "Total Water Molar Fraction": "water_mole_frac",
    # genuine phase-segregated vapour block
    "Vapor Molar Rate": "vap_molar", "Vapor Mass Rate": "vap_mass",
    "Vapor Act. Vol. Rate": "vap_act_vol", "Vapor Molecular Weight": "vap_mw",
    "Vapor Z (from k)": "vap_z_k", "Vapor Z (from actual density)": "vap_z",
    "Vapor Act. Density": "vap_density", "Vapor Sp. Enthalpy": "vap_sp_enthalpy",
    "Vapor CP": "vap_cp", "Vapor Viscosity": "vap_visc",
    "Vapor Therm. Cond.": "vap_therm_cond", "Vapor CP/CV Ratio": "vap_cp_cv",
    "Vapor CV": "vap_cv",
    # genuine phase-segregated liquid block
    "Liquid Molar Rate": "liq_molar", "Liquid Mass Rate": "liq_mass",
    "Liquid Molecular Weight": "liq_mw", "Liquid Z (from k-value)": "liq_z_k",
    "Liquid CP": "liq_cp", "Liquid Std. API": "liq_std_api",
    "Liquid Std. Density": "liq_std_density", "Liquid Act. Density": "liq_density",
    "Liquid Sp. Enthalpy": "liq_sp_enthalpy", "Liquid Viscosity": "liq_visc",
    "Liquid Therm. Cond.": "liq_therm_cond", "Liquid Surface Tension.": "liq_surf_tens",
}
_PREFIX = {   # bracketed-suffix labels (startswith; first occurrence wins)
    "Total Std. Liq. Rate": "total_std_liq", "Total Std. Vap. Rate": "total_std_vap",
    "Vapor Std. Vol. Rate": "vap_std_vol", "Liquid Act. Rate": "liq_act_rate",
    "Liquid Std Rate": "liq_std_rate",
}
KEYS = tuple(_EXACT.values()) + tuple(_PREFIX.values())

# dump key → key used by the stream sheets / OUTPUT (the td_parser vocabulary)
PK2TD = {
    'temp_f': 'temperature', 'pres_psia': 'pressure', 'total_molar': 'molar_rate',
    'total_mass': 'mass_rate', 'total_std_liq': 'std_liq_rate', 'total_std_vap': 'std_vap_rate',
    'total_act_vol': 'act_vol_rate', 'mol_weight': 'mol_weight', 'total_enthalpy': 'enthalpy',
    'total_sp_enthalpy': 'sp_enthalpy', 'total_sp_entropy': 'sp_entropy',
    'liq_mole_frac': 'liq_mole_frac', 'vap_mole_frac': 'vap_mole_frac', 'liq_wt_frac': 'liq_wt_frac',
    'tc_f': 'tc_f', 'pc_psia': 'pc_psia', 'tc_vol': 'tc_vol', 'tc_kays': 'tc_kays',
    'pc_kays': 'pc_kays', 'acentric': 'acentric', 'watson_k': 'watson_k',
    'total_std_liq_density': 'std_liq_den', 'total_std_api': 'std_api',
    'total_density': 'act_density', 'total_z': 'z_total', 'water_mole_frac': 'water_mole_frac',
    'vap_molar': 'vap_molar_rate', 'vap_mass': 'vap_mass_rate', 'vap_act_vol': 'vap_act_vol',
    'vap_mw': 'vap_mw', 'vap_z_k': 'vap_z_k', 'vap_z': 'vap_z_den', 'vap_density': 'vap_act_density',
    'vap_sp_enthalpy': 'vap_sp_enthalpy', 'vap_cp': 'vap_cp', 'vap_visc': 'vap_viscosity',
    'vap_therm_cond': 'vap_therm_cond', 'vap_cp_cv': 'vap_cp_cv', 'vap_cv': 'vap_cv',
    'vap_std_vol': 'vap_std_vol', 'liq_molar': 'liq_molar_rate', 'liq_mass': 'liq_mass_rate',
    'liq_mw': 'liq_mw', 'liq_z_k': 'liq_z_k', 'liq_cp': 'liq_cp', 'liq_std_api': 'liq_std_api',
    'liq_std_density': 'liq_std_density', 'liq_density': 'liq_act_density',
    'liq_sp_enthalpy': 'liq_sp_enthalpy', 'liq_visc': 'liq_viscosity',
    'liq_therm_cond': 'liq_therm_cond', 'liq_surf_tens': 'liq_surf_tension',
    'liq_act_rate': 'liq_act_rate', 'liq_std_rate': 'liq_std_lv_rate',
}
# composition-block header (startswith) → section type
COMP_SECT_LABELS = {
    'Total Weight Comp. Rates': 'MASS_FLOW', 'Total Weight Comp. Fractions': 'MASS_FRAC',
    'Total Molar Comp. Rates': 'MOLE_FLOW', 'Total Molar Comp. Fractions': 'MOLE_FRAC',
    'Vapor Mole Comp. Fractions': 'VAP_MOLE_FRAC', 'Liquid Mole Comp. Fractions': 'LIQ_MOLE_FRAC',
}
_STDLIQ_COMP_RE = re.compile(r"^(Total|Toatl)\s+Std\.?\s+Liquid\s+Comp\.?\s+Rates", re.I)


def _label_key(lab: str):
    k = _EXACT.get(lab)
    if k is None:
        for pref, kk in _PREFIX.items():
            if lab.startswith(pref):
                return kk
    return k


def _is_header_label(s: str) -> bool:
    return (s in _EXACT or 'Comp.' in s or s.startswith(("Total", "Toatl", "Dry "))
            or any(s.startswith(p) for p in _PREFIX)
            or any(s.startswith(c) for c in COMP_SECT_LABELS))


# ════════════════════════════════════════════════════════════════════════
#  2.  DUMP READER
# ════════════════════════════════════════════════════════════════════════
def _find_dump_sheet(wb, sheet=None):
    if sheet:
        if sheet not in wb.sheetnames:
            raise KeyError(f"sheet '{sheet}' not in {wb.sheetnames}")
        return sheet
    for n in wb.sheetnames:
        ws = wb[n]
        for r in ws.iter_rows(min_row=1, max_row=12, max_col=1, values_only=True):
            if r and r[0] is not None and str(r[0]).strip() == "Stream Name":
                return n
    raise KeyError("no sheet with a 'Stream Name' row in column A — is this a PRO/II property dump? "
                   f"sheets: {wb.sheetnames}")


def read_dump(path, sheet=None):
    """→ (rows, sheet_name): the dump as a list of value-tuples."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        sn = _find_dump_sheet(wb, sheet)
        rows = [tuple(r) for r in wb[sn].iter_rows(values_only=True)]
    finally:
        wb.close()
    return rows, sn


def parse_dump(path, sheet=None, verbose=True):
    """Label-driven parse of a transposed PRO/II dump.

    Returns dict(streams, comp_data, stream_names, comp_names, comp_stats, sheet)
      streams[name]   = {td-key: value}          scalar properties
      comp_data[name] = {section: {comp: value}}  composition blocks
      comp_stats[comp]= {mw, sld, n}             measured from the component flows
    """
    rows, sn = read_dump(path, sheet)
    if verbose:
        print(f"  Dump sheet '{sn}': {len(rows)} rows × {max(len(r) for r in rows)} cols")
    labs = [(str(r[0]).strip() if (r and r[0] is not None) else None) for r in rows]
    name_ri = next(i for i, l in enumerate(labs) if l == "Stream Name")
    ncol = max(len(r) for r in rows)
    stream_names, col_map = [], {}
    for ci in range(2, ncol):
        v = rows[name_ri][ci] if ci < len(rows[name_ri]) else None
        if v is None or str(v).strip() in ("", "nan"):
            continue
        s0 = str(v).strip()
        base, n = s0, 1
        while s0 in col_map:
            s0 = f"{base}_{n}"
            n += 1
        col_map[s0] = ci
        stream_names.append(s0)
    if verbose:
        print(f"  Streams found: {len(stream_names)}")

    label_rows, comp_starts, stdliq_start = {}, {}, None
    for ri, lab in enumerate(labs):
        if not lab:
            continue
        key = _label_key(lab)
        if key is not None and key not in label_rows:
            label_rows[key] = ri
        for clab, ctype in COMP_SECT_LABELS.items():
            if ctype not in comp_starts and lab.startswith(clab):
                comp_starts[ctype] = ri
        if stdliq_start is None and _STDLIQ_COMP_RE.match(lab):
            stdliq_start = ri

    def comp_rows(start_ri):
        out = []
        for ri in range(start_ri + 1, len(rows)):
            lab = labs[ri]
            if not lab or lab == 'nan' or _is_header_label(lab):
                break
            out.append((ri, lab))
        return out

    comp_layout = {ct: comp_rows(ri) for ct, ri in comp_starts.items()}
    comp_names = [n for _, n in comp_layout.get('MASS_FLOW', comp_layout.get('MOLE_FLOW', []))]

    streams, comp_data = {}, {}
    for sn_ in stream_names:
        ci = col_map[sn_]
        props = {}
        for pk, ri in label_rows.items():
            val = rows[ri][ci] if ci < len(rows[ri]) else None
            if pk == 'phase':
                props['phase_str'] = str(val).strip() if val is not None else ''
            else:
                tk = PK2TD.get(pk)
                if tk is not None:
                    props[tk] = sf(val)
        props.setdefault('phase_str', '')
        streams[sn_] = props
        cd = {}
        for ct, lst in comp_layout.items():
            blk = {}
            for ri, nm in lst:
                v = sf(rows[ri][ci] if ci < len(rows[ri]) else None)
                if v is not None:
                    blk[nm] = v
            cd[ct] = blk
        comp_data[sn_] = cd

    # per-component MW / SLD measured from the component flows (identical in every stream)
    stats = {}
    mass = {n: ri for ri, n in comp_layout.get('MASS_FLOW', [])}
    mole = {n: ri for ri, n in comp_layout.get('MOLE_FLOW', [])}
    sliq = {}
    if stdliq_start is not None:
        sliq = {n: ri for ri, n in comp_rows(stdliq_start)}
    for nm, rm in mass.items():
        mws, slds = [], []
        rn, rs = mole.get(nm), sliq.get(nm)
        for ci in col_map.values():
            m = sf(rows[rm][ci] if ci < len(rows[rm]) else None)
            if not m or m <= 1e-9:
                continue
            if rn is not None:
                n_ = sf(rows[rn][ci] if ci < len(rows[rn]) else None)
                if n_ and n_ > 1e-12:
                    mws.append(m / n_)
            if rs is not None:
                s_ = sf(rows[rs][ci] if ci < len(rows[rs]) else None)
                if s_ and s_ > 1e-12:
                    slds.append(m / s_)
        stats[nm] = dict(mw=statistics.median(mws) if mws else None,
                         sld=statistics.median(slds) if slds else None, n=len(mws))
    return dict(streams=streams, comp_data=comp_data, stream_names=stream_names,
                comp_names=comp_names, comp_stats=stats, sheet=sn,
                label_rows=label_rows, n_label_keys=len(label_rows))

# ════════════════════════════════════════════════════════════════════════
#  3.  EMBEDDED PURE-COMPONENT DATABANK
# ════════════════════════════════════════════════════════════════════════
# 459 components (simulator databank 'Sheet2' + the 98-component PR/Therm set for the
# rest).  name: (CAS, MW, NBP °F, Tc °F, Pc psia, Vc ft3/lbmol, Zc, ω, PR Peneloux,
#                SRK Peneloux, parachor, std liquid density lb/ft3)
# Regenerate from the property workbooks rather than hand-editing.
PURE_DB = {
    'BENZENE': ('71-43-2', 78.11364, 176.16, 552.22, 710.39, 4.1488, 0.271, 0.210024, -0.069422, 0.13884, 245.99, 54.495),
    '12BD': ('590-19-2', 54.09164, 51.53, 353.93, 632.36, 3.5241, 0.255, 0.165877, -0.086275, 0.1036, 187.399, 41.364),
    '13BD': ('106-99-0', 54.09164, 24.06, 305.64, 620.33, 3.5305, 0.267, 0.189465, -0.070653, 0.11052, 178.812, 40.673),
    'BUTANE': ('106-97-8', 58.1234, 31.1, 305.55, 550.56, 4.0847, 0.274, 0.200164, -0.073426, 0.13022, 191.001, 37.608),
    '1BUTENE': ('106-98-9', 56.10752, 20.75, 296.24, 586.39, 3.8332, 0.277, 0.19054, -0.07325, 0.11601, 183.875, 39.078),
    'C2BUTENE': ('590-18-1', 56.10752, 38.7, 324.37, 615.39, 3.7451, 0.274, 0.204776, -0.064869, 0.12162, 188.511, 40.071),
    'T2BUTENE': ('624-64-6', 56.10752, 33.58, 311.86, 594.65, 3.8076, 0.273, 0.217664, -0.059122, 0.13027, 187.176, 39.203),
    'CO2': ('124-38-9', 44.0098, -109.26, 87.91, 1070.8, 1.5057, 0.274, 0.223621, -0.022038, 0.052514, 78.475, 73.708),
    'CO': ('630-08-0', 28.0104, -312.61, -220.41, 507.49, 1.5121, 0.299, 0.048162, -0.054642, 0.016658, 45.896, 49.337),
    'DECANE': ('124-18-5', 142.2847, 345.48, 652.19, 306.03, 9.6111, 0.247, 0.492328, 0.15097, 0.54648, 425.471, 45.36),
    '22MB': ('75-83-2', 86.17715, 121.51, 420.13, 446.86, 5.7506, 0.272, 0.234967, -0.07558, 0.21077, 255.328, 40.234),
    '23MB': ('79-29-8', 86.17715, 136.36, 440.29, 453.53, 5.7314, 0.269, 0.246104, -0.067004, 0.22091, 263.979, 41.019),
    'ETHANE': ('74-84-0', 30.06964, -127.48, 89.91, 706.62, 2.3307, 0.279, 0.099493, -0.073577, 0.042808, 113.382, 33.989),
    'EBENZENE': ('100-41-4', 106.1674, 277.16, 651.29, 523.01, 5.9909, 0.263, 0.302604, -0.021926, 0.28255, 350.474, 53.988),
    'ETHYLENE': ('74-85-1', 28.05376, -154.73, 48.54, 730.99, 2.0984, 0.281, 0.086452, -0.069534, 0.034785, 102.9, 35.509),
    'HEPTANE': ('142-82-5', 100.204, 209.17, 512.69, 397.4, 6.8559, 0.261, 0.349469, 0.022326, 0.36949, 313.277, 42.548),
    'HEXANE': ('110-54-3', 86.17715, 155.71, 454.01, 438.74, 5.9428, 0.266, 0.301261, -0.022656, 0.27594, 272.896, 40.955),
    'H2': ('1333-74-0', 2.01588, -422.97, -399.93, 190.43, 1.0275, 0.305, -0.215993, -0.070747, -0.02073, -57.023, 4.405),
    'H2S': ('7783-06-4', 34.08188, -76.63, 212.68, 1300.0, 1.5778, 0.284, 0.094168, -0.050072, 0.02741, 75.173, 59.332),
    'IBUTANE': ('75-28-5', 58.1234, 10.9, 274.98, 529.1, 4.208, 0.282, 0.180771, -0.084527, 0.11975, 183.155, 34.48),
    'IBUTENE': ('115-11-7', 56.10752, 19.58, 292.55, 580.15, 3.8268, 0.275, 0.194273, -0.071668, 0.11854, 182.809, 36.727),
    'IPENTANE': ('78-78-4', 72.15028, 82.12, 369.1, 490.37, 4.8984, 0.27, 0.227461, -0.070143, 0.17605, 224.129, 38.466),
    'METHANE': ('74-82-8', 16.04276, -258.68, -116.65, 667.03, 1.5794, 0.286, 0.011548, -0.067415, 0.010942, 71.588, 26.377),
    '2MP': ('107-83-5', 86.17715, 140.47, 435.83, 436.56, 5.8692, 0.267, 0.277438, -0.042593, 0.25303, 263.751, 40.472),
    '3MP': ('96-14-0', 86.17715, 145.89, 448.3, 453.1, 5.8692, 0.273, 0.273729, -0.044695, 0.24434, 269.357, 41.202),
    'N2': ('7727-37-9', 28.01348, -320.45, -232.51, 493.13, 1.429, 0.289, 0.037721, -0.055386, 0.01443, 43.48, 50.438),
    'NONANE': ('111-84-2', 128.2578, 303.48, 610.61, 332.14, 8.714, 0.252, 0.44346, 0.16818, 0.53939, 393.78, 44.591),
    'OCTANE': ('111-65-9', 114.2309, 258.22, 563.99, 361.14, 7.785, 0.256, 0.399552, 0.08477, 0.48253, 353.379, 43.65),
    'O2': ('7782-44-7', 31.9988, -297.33, -181.43, 731.43, 1.1758, 0.288, 0.02218, -0.048192, 0.0096469, 55.369, 71.286),
    'PENTANE': ('109-66-0', 72.15028, 96.93, 385.79, 488.78, 5.0138, 0.27, 0.251506, -0.05453, 0.19615, 231.923, 38.813),
    '1PENTENE': ('109-67-1', 70.1344, 85.93, 376.93, 509.52, 4.7254, 0.268, 0.23115, -0.065632, 0.17337, 224.158, 39.633),
    'PROPANE': ('74-98-6', 44.09652, -43.67, 206.02, 616.12, 3.2037, 0.276, 0.152291, -0.078543, 0.08137, 150.83, 36.365),
    'PROPENE': ('115-07-1', 42.08064, -53.84, 198.36, 676.6, 3.0179, 0.289, 0.139817, -0.075732, 0.068585, 145.621, 38.096),
    'TOLUENE': ('108-88-3', 92.14052, 231.13, 605.57, 595.52, 5.0618, 0.264, 0.262122, -0.04851, 0.21013, 299.626, 53.975),
    'H2O': ('7732-18-5', 18.01528, 212.0, 705.16, 3198.8, 0.8962, 0.229, 0.344861, 0.0026264, 0.054345, 52.6, 62.242),
    'OXYLENE': ('95-47-6', 106.1674, 291.97, 674.92, 541.57, 5.9108, 0.263, 0.310448, -0.014805, 0.28498, 354.831, 54.709),
    'MXYLENE': ('108-38-3', 106.1674, 282.42, 651.02, 512.85, 6.0229, 0.259, 0.325855, -0.0014595, 0.30741, 351.594, 53.715),
    'PXYLENE': ('106-42-3', 106.1674, 281.05, 649.54, 509.23, 6.071, 0.26, 0.321495, -0.0054087, 0.30554, 352.885, 53.485),
    'ACENAP': ('83-32-9', 154.2114, 531.3, 986.0, 449.62, 8.8582, 0.257, 0.381147, 0.18264, 0.61048, 460.297, 64.271),
    'ACENAPHT': ('208-96-8', 152.1955, 518.0, 965.93, 464.12, 8.714, 0.264, 0.398701, 0.17622, 0.59889, 454.411, 53.342),
    'ANTH': ('120-12-7', 178.2334, 647.65, 1111.73, 420.61, 8.8742, 0.221, 0.485671, 0.21323, 0.66182, 530.442, 60.835),
    'BIPHENYL': ('92-52-4', 154.2114, 491.0, 961.0, 558.4, 8.0413, 0.295, 0.365368, 0.18264, 0.61048, 460.297, 61.995),
    'PTBUTEBZ': ('7364-19-4', 162.2749, 413.78, 771.53, 345.19, 9.6912, 0.253, 0.434698, 0.1956, 0.63141, 483.843, 53.714),
    'BBENZENE': ('104-51-8', 134.22121, 361.95, 729.32, 418.72, 7.9612, 0.261, 0.393793, 0.1551, 0.5356, 448.899, 53.543),
    'SBBENZEN': ('135-98-8', 134.22121, 343.99, 736.5, 427.86, 7.9612, 0.265, 0.279149, 0.1551, 0.5356, 452.551, 53.543),
    'TBBENZEN': ('98-06-6', 134.22121, 336.47, 728.33, 430.76, 7.8811, 0.266, 0.267406, 0.1551, 0.5356, 446.007, 53.848),
    '1BNP': ('1634-09-9', 184.28101, 552.9, 965.93, 388.7, 10.108, 0.257, 0.495136, 0.22531, 0.68618, 548.101, 60.755),
    '2BNP': ('1134-62-9', 184.28101, 550.4, 968.0, 368.87, 9.7713, 0.2353, 0.45, 0.22531, 0.68618, 548.101, 60.067),
    '4TBUTSTY': ('1746-23-2', 160.259, 440.33, 816.53, 359.69, 9.515, 0.25, 0.421575, 0.19602, 0.63364, 477.956, 55.176),
    'CHRYSENE': ('218-01-9', 228.2933, 825.8, 1302.53, 346.64, 11.998, 0.22, 0.603008, 0.31733, 0.87118, 676.616, 63.878),
    'CUMENE': ('98-82-8', 120.1943, 306.34, 676.31, 465.43, 6.8399, 0.261, 0.325809, -0.0016914, 0.3464, 389.629, 53.653),
    'CHXBNZ': ('827-52-1', 160.259, 464.22, 879.53, 417.71, 8.4738, 0.246, 0.378294, 0.19602, 0.63364, 477.956, 58.621),
    'OCYMENE': ('527-84-4', 134.22121, 352.72, 731.93, 424.96, 7.833, 0.26, 0.337173, 0.1551, 0.5356, 446.338, 54.486),
    'MCYMENE': ('535-77-3', 134.22121, 347.14, 722.93, 424.96, 7.7689, 0.26, 0.341126, 0.1551, 0.5356, 439.933, 53.511),
    'PCYMENE': ('99-87-6', 134.22121, 350.83, 716.0, 406.11, 7.9612, 0.256, 0.366752, 0.1551, 0.5356, 442.041, 53.218),
    'C10BNZN': ('104-72-3', 218.3824, 568.2, 895.73, 256.72, 13.023, 0.23, 0.679718, 0.29637, 0.8292, 647.677, 53.169),
    'C10NAPH': ('26438-27-7', 268.44229, 713.93, 1086.53, 229.16, 17.14, 0.237, 0.641493, 0.40469, 1.0421, 793.851, 57.945),
    'DIBZTHIO': ('132-65-0', 184.2615, 628.63, 1154.93, 559.85, 8.2014, 0.265, 0.398277, 0.22527, 0.6861, 548.044, 70.224),
    'ODEBZ': ('135-01-3', 134.22121, 362.22, 742.73, 417.71, 8.0413, 0.26, 0.339547, 0.1551, 0.5356, 458.16, 54.666),
    '13DEBZ': ('141-93-5', 134.22121, 358.04, 733.73, 417.71, 7.817, 0.255, 0.354019, 0.1551, 0.5356, 446.73, 53.691),
    '14EZ': ('105-05-5', 134.22121, 362.82, 724.66, 406.54, 7.9612, 0.255, 0.402977, 0.1551, 0.5356, 446.511, 53.562),
    '26DENP': ('59919-41-4', 184.28101, 577.13, 992.93, 374.2, 10.092, 0.242, 0.511844, 0.22531, 0.68618, 548.101, 62.327),
    'MDIPBN': ('99-62-7', 162.2749, 397.72, 771.53, 355.34, 9.6111, 0.258, 0.358687, 0.1956, 0.63141, 483.843, 53.214),
    'PDIPBN': ('100-18-5', 162.2749, 410.9, 780.53, 355.34, 9.579, 0.256, 0.390023, 0.1956, 0.63141, 483.843, 53.252),
    'DM23DPB': ('1889-67-4', 238.3727, 600.53, 989.33, 288.62, 12.51, 0.232, 0.520782, 0.33803, 0.91287, 706.048, 63.779),
    '26DMNP': ('581-42-0', 156.22729, 503.6, 938.93, 459.77, 8.3296, 0.255, 0.417678, 0.18821, 0.6204, 466.184, 57.989),
    '27DMNP': ('582-16-1', 156.22729, 505.4, 940.73, 459.77, 8.3296, 0.255, 0.419584, 0.18821, 0.6204, 466.184, 58.736),
    'DPHACET': ('501-65-5', 178.2334, 571.73, 1037.93, 420.61, 9.7873, 0.256, 0.383595, 0.21323, 0.66182, 530.442, 59.039),
    'DPHENET': ('612-00-0', 182.2652, 522.73, 935.33, 388.7, 9.6751, 0.251, 0.456636, 0.22134, 0.67813, 542.214, 62.175),
    '12DPHNLE': ('103-29-7', 182.2652, 536.9, 944.33, 384.35, 9.8674, 0.252, 0.488475, 0.22134, 0.67813, 542.214, 60.263),
    'DPHM': ('101-81-5', 168.2383, 507.69, 922.73, 423.51, 8.7621, 0.25, 0.461542, 0.19739, 0.63179, 501.256, 62.554),
    'MDVBNZN': ('108-57-6', 130.18941, 391.1, 785.93, 452.52, 7.0481, 0.239, 0.373361, 0.16301, 0.53706, 440.996, 57.77),
    'DDBZ': ('123-01-3', 246.4362, 621.7, 944.33, 226.26, 14.785, 0.222, 0.733309, 0.3541, 0.94542, 729.594, 53.202),
    'DTG': ('62587-63-7', 208.0, 550.4, 1007.6, 382.09, 10.047, 0.244, 0.341, 0.27376, 0.78419, 617.36, None),
    'DTHT': (None, 240.0, 649.99, 1088.6, 293.92, None, None, 0.4571, 0.34133, 0.91952, None, None),
    'DTJ': ('92170-72-4', 134.0, 357.8, 721.49, 411.49, 7.8394, 0.255, 0.258, 0.15543, 0.53557, 441.328, None),
    'DTLF': ('81544-18-5', 175.39999, 527.0, 950.0, 404.14, None, None, 0.4387, 0.20739, 0.6502, None, None),
    '1ENP': ('1127-76-0', 156.22729, 496.99, 937.13, 435.11, 8.3296, 0.242, 0.362617, 0.18821, 0.6204, 466.184, 62.689),
    '2ENP': ('939-27-5', 156.22729, 496.22, 928.13, 459.77, 8.3296, 0.257, 0.421298, 0.18821, 0.6204, 466.184, 61.733),
    'M2EZ': ('611-14-3', 120.1943, 329.32, 712.4, 440.88, 7.3685, 0.258, 0.293233, -0.037778, 0.34406, 422.435, 54.726),
    'M3EZ': ('620-14-4', 120.1943, 322.39, 687.2, 411.91, 7.849, 0.263, 0.322552, -0.0056925, 0.39168, 423.995, 53.716),
    'M4EZ': ('622-96-8', 120.1943, 323.62, 692.74, 468.91, 6.8399, 0.259, 0.366325, 0.039615, 0.38703, 396.38, 53.503),
    '2EMXYL': ('2870-04-4', 134.22121, 374.07, 748.13, 438.01, 7.7209, 0.261, 0.406602, 0.1551, 0.5356, 450.574, 55.335),
    '2EPXYL': ('1758-88-9', 134.22121, 368.29, 733.73, 417.71, 7.7209, 0.252, 0.411418, 0.1551, 0.5356, 443.703, 54.511),
    '12DM3EBZ': ('933-98-2', 134.22121, 381.13, 764.33, 417.71, 8.1214, 0.258, 0.362112, 0.1551, 0.5356, 471.677, 55.445),
    '4EOXYL': ('934-80-5', 134.22121, 373.6, 740.93, 417.71, 7.849, 0.254, 0.411361, 0.1551, 0.5356, 451.277, 54.349),
    '4EMXYL': ('874-41-9', 134.22121, 371.19, 737.33, 417.71, 7.7209, 0.251, 0.413957, 0.1551, 0.5356, 445.432, 54.457),
    '5EMXYL': ('934-74-7', 134.22121, 362.8, 719.33, 398.85, 7.7209, 0.243, 0.416938, 0.1551, 0.5356, 436.708, 53.738),
    'FLUANTHE': ('206-44-0', 202.2554, 721.04, 1169.33, 378.55, 10.492, 0.227, 0.587526, 0.26079, 0.75863, 600.586, 68.364),
    'FLUORENE': ('86-73-7', 166.2224, 567.12, 1106.33, 681.68, 6.4074, 0.26, 0.349259, 0.19628, 0.6305, 495.369, 80.918),
    'C17BENZ': ('14752-75-1', 316.57059, 732.2, 1027.13, 175.5, 19.222, 0.211, 0.919461, 0.57669, 1.3197, 934.386, 52.88),
    'C7BZ': ('1078-71-3', 176.3018, 474.98, 825.53, 316.18, 10.38, 0.238, 0.527166, 0.20926, 0.65391, 524.801, 53.317),
    'C16BENZ': ('1459-09-2', 302.5437, 712.4, 1012.73, 184.2, 18.261, 0.213, 0.899552, 0.5132, 1.2243, 893.428, 53.03),
    'HXBNZN': ('1077-16-3', 162.2749, 439.0, 796.73, 345.19, 9.4989, 0.243, 0.478964, 0.1956, 0.63141, 483.843, 53.348),
    '1HNP': ('2876-53-1', 212.33479, 611.6, 1003.73, 326.33, 11.87, 0.247, 0.587382, 0.2833, 0.80314, 630.018, 59.136),
    '2HNP': ('2876-46-2', 212.33479, 615.2, 1004.0, 304.21, 11.629, 0.2252, 0.571, 0.2833, 0.80314, 630.018, 59.013),
    'NC6THNP': ('66325-11-9', 216.3665, 581.0, 942.53, 274.12, 12.35, 0.225, 0.588792, 0.29203, 0.82055, 641.79, 57.285),
    '4HSTYREN': ('2628-17-3', 120.1509, 443.93, 865.13, 655.57, 5.9909, 0.276, 0.50899, 0.1522, 0.42892, 409.912, 58.71),
    'INDANE': ('56573-11-6', 118.1784, 352.35, 773.15, 572.9, 6.3433, 0.275, 0.309213, -0.01631, 0.2917, 404.544, 59.916),
    'INDENE': ('95-13-6', 116.1625, 360.72, 776.93, 554.04, 5.8948, 0.246, 0.333766, 0.0058226, 0.32359, 382.865, 62.056),
    'IBBENZEN': ('538-93-2', 134.22121, 343.02, 710.6, 440.91, 7.3044, 0.256, 0.379669, 0.1551, 0.5356, 419.549, 53.01),
    '4IBUTSTY': ('63444-56-4', 160.259, 483.53, 863.33, 345.19, 9.9475, 0.242, 0.463134, 0.19602, 0.63364, 477.956, 54.806),
    'PISOPRST': ('16262-48-9', 144.21629, 467.33, 874.13, 403.2, 8.5058, 0.24, 0.406837, 0.15229, 0.55181, 431.112, 58.199),
    '4IPDPH': ('7116-95-2', 196.29201, 546.8, 998.6, 337.94, 10.551, 0.2278, 0.299, 0.24741, 0.73224, 583.173, 60.595),
    '135MBENZ': ('108-67-8', 120.1943, 328.53, 687.58, 453.53, 6.936, 0.256, 0.399043, 0.075117, 0.43013, 397.487, 53.773),
    '24DPH4MP': ('6362-80-7', 236.3568, 645.53, 1043.33, 290.08, 13.439, 0.242, 0.546694, 0.33394, 0.90461, 700.162, 61.49),
    '2PB1': ('2039-93-2', 132.20531, 359.6, 739.13, 436.56, 7.2724, 0.247, 0.353568, 0.15857, 0.53576, 430.962, 55.374),
    'AMST': ('98-83-9', 118.1784, 329.9, 717.53, 487.33, 6.3914, 0.247, 0.32297, -0.0045197, 0.34021, 388.631, 56.466),
    '1MINDENE': ('767-59-9', 130.18941, 389.3, 805.73, 501.83, 6.984, 0.258, 0.334861, 0.16301, 0.53706, 445.494, 60.362),
    '2MINDENE': ('2177-47-1', 130.18941, 403.34, 820.13, 501.83, 6.984, 0.255, 0.350814, 0.16301, 0.53706, 450.658, 60.621),
    '1MNP': ('90-12-0', 142.20039, 472.43, 930.0, 530.84, 7.4486, 0.265, 0.347775, 0.15093, 0.54627, 425.225, 63.454),
    '2MNP': ('91-57-6', 142.20039, 466.0, 910.13, 507.63, 7.4486, 0.257, 0.371584, 0.15093, 0.54627, 425.225, 62.029),
    '1M2PRBZ': ('1074-17-5', 134.22121, 364.64, 731.93, 426.41, 7.7209, 0.257, 0.407014, 0.1551, 0.5356, 442.835, 54.292),
    '1M3PRBZ': ('1074-43-7', 134.22121, 359.24, 717.53, 407.56, 7.7209, 0.249, 0.412785, 0.1551, 0.5356, 435.825, 53.498),
    '1M4PRBZ': ('1074-55-1', 134.22121, 361.94, 721.13, 407.56, 7.7209, 0.248, 0.413441, 0.1551, 0.5356, 437.589, 53.336),
    'OMSTYR': ('611-15-4', 118.1784, 337.66, 726.53, 503.28, 6.5195, 0.258, 0.341159, 0.013379, 0.3484, 397.05, 56.667),
    'MMSTYR': ('100-80-1', 118.1784, 340.88, 722.93, 477.17, 6.5195, 0.245, 0.348736, 0.02186, 0.37355, 395.781, 56.659),
    'PVTOL': ('622-97-9', 118.1784, 343.0, 737.33, 487.33, 6.904, 0.262, 0.317536, -0.010134, 0.3408, 416.443, 57.16),
    'NAPHTHLN': ('91-20-3', 128.17349, 424.39, 887.36, 587.55, 6.6156, 0.269, 0.302169, -0.024078, 0.30458, 452.891, 61.073),
    'C9BZ': ('1081-77-2', 204.3555, 539.69, 874.13, 274.85, 12.062, 0.232, 0.633063, 0.26556, 0.76802, 606.718, 53.202),
    '1NNP': ('26438-26-6', 254.4154, 690.53, 1068.53, 243.66, 16.018, 0.238, 0.616778, 0.37012, 0.97758, 752.893, 58.287),
    '2NNP': ('61886-67-7', 254.4154, 701.6, 1065.2, 241.01, 14.353, 0.2114, 0.686, 0.37012, 0.97758, 752.893, 57.868),
    'C18BENZ': ('4445-07-2', 330.59741, 752.0, 1041.53, 168.24, 20.023, 0.209, 0.950469, 0.69671, 1.4786, 975.344, 52.833),
    'DIPPHTH': ('2189-60-8', 190.32871, 507.92, 852.53, 292.98, 11.261, 0.234, 0.567041, 0.23598, 0.70868, 565.76, 53.239),
    'C15BENZ': ('2131-18-2', 288.51682, 690.8, 996.53, 192.9, 17.46, 0.216, 0.856716, 0.46392, 1.1447, 852.469, 53.11),
    'PNBZ': ('538-68-1', 148.248, 401.83, 764.15, 377.68, 8.8101, 0.253, 0.437782, 0.16116, 0.57132, 442.884, 53.357),
    '1PNAP': ('86-89-5', 198.30791, 582.8, 985.73, 353.89, 10.989, 0.251, 0.535388, 0.25172, 0.74084, 589.059, 60.112),
    'PHAN': ('85-01-8', 178.2334, 638.38, 1104.98, 420.61, 8.8742, 0.222, 0.469466, 0.21323, 0.66182, 530.442, 66.597),
    'C2PHBUT2': (None, 132.20531, 382.46, 773.33, 436.56, 7.5287, 0.248, 0.356388, 0.15857, 0.53576, 455.211, 57.378),
    'T2PHBUT2': ('768-00-3', 132.20531, 345.2, 717.53, 436.56, 7.4326, 0.257, 0.361021, 0.15857, 0.53576, 426.844, 55.672),
    '1PHNAP': ('605-02-7', 204.2713, 633.2, 1068.53, 381.45, 10.508, 0.244, 0.530878, 0.26537, 0.76765, 606.472, 67.144),
    'C1PRBENZ': ('766-90-5', 118.1784, 353.98, 748.13, 487.33, 6.5195, 0.245, 0.341093, 0.014001, 0.36629, 404.518, 56.469),
    'TBMS': ('873-66-5', 118.1784, 352.87, 746.33, 487.33, 6.5195, 0.245, 0.339306, 0.012145, 0.36405, 403.905, 56.431),
    'PRBENZEN': ('103-65-1', 120.1943, 318.63, 689.41, 464.12, 7.0481, 0.265, 0.34472, 0.017712, 0.36936, 401.926, 53.661),
    '1PRNP': ('2765-18-6', 170.2542, 523.0, 947.93, 430.76, 8.3296, 0.238, 0.455435, 0.19901, 0.63427, 507.142, 61.604),
    '2PNP': ('2027-19-2', 170.2542, 524.3, 944.6, 412.96, 9.3227, 0.2555, 0.454, 0.19901, 0.63427, 507.142, 60.768),
    'PYRENE': ('129-00-0', 202.2554, 742.64, 1225.13, 378.55, 10.572, 0.221, 0.507416, 0.26079, 0.75863, 600.586, 68.047),
    'CSTILBEN': ('645-49-8', 180.2493, 537.53, 951.53, 397.4, 9.547, 0.251, 0.475803, 0.21731, 0.67001, 536.328, 63.095),
    'TSTILBEN': ('103-30-0', 180.2493, 583.7, 1016.33, 397.4, 9.7232, 0.244, 0.487424, 0.21731, 0.67001, 536.328, 59.638),
    'STYRENE': ('100-42-5', 104.1515, 293.29, 685.13, 556.94, 5.6385, 0.256, 0.297097, -0.025914, 0.26907, 345.873, 56.213),
    'OTPH': ('84-15-1', 230.3092, 639.5, 1144.04, 565.65, 12.062, 0.396, 0.481734, 0.32153, 0.87961, 682.503, 65.632),
    'MTPH': ('92-06-8', 230.3092, 710.33, 1205.06, 509.08, 12.302, 0.351, 0.558777, 0.32153, 0.87961, 682.503, 65.021),
    'PTPH': ('92-94-4', 230.3092, 708.8, 1207.04, 481.53, 12.222, 0.329, 0.523983, 0.32153, 0.87961, 682.503, 60.062),
    'C14BENZ': ('1459-10-5', 274.48999, 669.2, 980.33, 203.05, 16.499, 0.217, 0.812954, 0.42149, 1.0719, 811.511, 53.184),
    'TETRALIN': ('119-64-2', 132.20531, 405.72, 836.6, 525.04, 7.0641, 0.267, 0.328416, 0.15857, 0.53576, 460.332, 60.373),
    '1234TMBZ': ('488-23-3', 134.22121, 401.07, 787.73, 451.07, 7.6088, 0.256, 0.417172, 0.1551, 0.5356, 464.573, 56.209),
    '1235TMBZ': ('527-53-7', 134.22121, 388.4, 762.53, 430.76, 7.7209, 0.254, 0.42416, 0.1551, 0.5356, 457.324, 55.349),
    'TMBZ': ('95-93-2', 134.22121, 386.31, 755.6, 426.41, 7.7209, 0.252, 0.434059, 0.1551, 0.5356, 454.09, 52.371),
    'TETPHET': ('632-50-8', 334.46069, 680.0, 1028.93, 242.21, 16.371, 0.248, 0.732152, 0.72248, 1.5143, 986.625, 60.411),
    'TETPNETE': ('632-51-9', 332.44479, 908.33, 1333.13, 248.01, 16.339, 0.211, 0.728463, 0.70982, 1.4966, 980.739, 73.516),
    'TETPHM': ('630-76-2', 320.43381, 877.73, 1309.73, 259.62, 15.474, 0.212, 0.679388, 0.59885, 1.3513, 945.667, 69.866),
    'T59': ('204336-88-9', 229.0, 598.6, 960.8, 309.0, 12.132, 0.3636, 0.6583, 0.31881, 0.87414, 678.68, None),
    'T60': ('72993-95-4', 250.0, 559.99, 998.6, 323.0, 12.077, 0.364, 0.3378, 0.36096, 0.95943, 740.0, None),
    'T66': ('54578-28-8', 252.0, 648.0, 1052.6, 338.0, 13.252, 0.3965, 0.6128, 0.36502, 0.96754, 745.84, None),
    'T75': ('143597-04-0', 255.0, 723.0, 1157.0, 367.0, 12.27, 0.3627, 0.6197, 0.37139, 0.98005, 754.6, None),
    'C13BNZN': ('123-02-4', 260.4631, 646.3, 962.33, 214.66, 15.666, 0.22, 0.779866, 0.3842, 1.0044, 770.552, 53.136),
    '123MBENZ': ('526-73-8', 120.1943, 349.02, 736.48, 500.98, 6.6316, 0.259, 0.366405, 0.038565, 0.37607, 405.211, 55.62),
    '124MBENZ': ('95-63-6', 120.1943, 336.88, 708.76, 468.76, 6.8879, 0.258, 0.377346, 0.051577, 0.40308, 404.546, 54.453),
    '123TMIND': ('4773-83-5', 158.2431, 456.53, 847.13, 401.75, 8.4738, 0.243, 0.441894, 0.19292, 0.62865, 472.07, 60.643),
    'TPHET': ('1520-42-9', 258.36291, 659.93, 1052.33, 295.88, 12.943, 0.236, 0.611614, 0.37912, 0.99486, 764.42, 69.93),
    'TRIPHET': ('58-72-0', 256.34699, 744.53, 1174.73, 304.58, 13.776, 0.239, 0.599841, 0.37437, 0.98581, 758.533, 65.035),
    'TRIPHM': ('519-73-3', 244.336, 678.2, 1097.33, 319.08, 12.062, 0.23, 0.573483, 0.35011, 0.93721, 723.461, 63.446),
    'C11BZ': ('6742-54-7', 232.4093, 595.85, 915.53, 242.5, 13.904, 0.228, 0.733305, 0.32588, 0.88835, 688.635, 53.154),
    'TCCP': ('15890-40-1', 112.215, 230.72, 570.2, 410.02, 6.6637, 0.2472, 0.272, -0.057823, 0.3046, 338.415, 46.778),
    'ADAMANTA': ('281-23-2', 136.237, 370.13, 805.73, 423.51, 7.7689, 0.242, 0.185299, 0.1526, 0.5366, 478.77, 44.614),
    'BICHXL': ('92-51-3', 166.3067, 462.27, 848.93, 371.3, 9.579, 0.253, 0.427556, 0.19631, 0.63053, 495.616, 55.114),
    'BCH': ('1678-93-9', 140.2688, 357.77, 740.93, 372.75, 8.5539, 0.247, 0.274326, 0.15054, 0.54204, 471.895, 49.679),
    'SBCH': ('7058-01-7', 140.2688, 354.83, 744.53, 387.97, 8.4658, 0.2542, 0.264, 0.15054, 0.54204, 471.508, 50.527),
    'TBCH': ('3178-22-1', 140.2688, 340.79, 726.53, 386.5, 8.2415, 0.2502, 0.252, 0.15054, 0.54204, 455.412, 50.524),
    'BCYP': ('2040-95-1', 126.2419, 313.88, 658.13, 394.5, 7.7369, 0.254, 0.37188, 0.052204, 0.45225, 405.913, 48.778),
    'CYBU': ('287-23-0', 56.10752, 54.52, 368.2, 722.29, 3.3639, 0.273, 0.18474, -0.067888, 0.1006, 183.373, 43.017),
    'CYCDECAN': ('293-96-9', 140.2688, 396.45, 823.28, 461.22, 7.9959, 0.2679, 0.279, 0.15054, 0.54204, 495.664, 53.73),
    'CHP': ('291-64-5', 98.18816, 245.82, 628.07, 556.98, 5.7506, 0.274, 0.243027, -0.068437, 0.21511, 333.888, 50.322),
    'CH': ('110-82-7', 84.16128, 177.3, 536.77, 590.74, 4.9337, 0.273, 0.209609, -0.0825, 0.16415, 277.171, 48.263),
    'CNON': ('293-55-0', 126.2419, 353.12, 699.8, 455.57, 8.1694, 0.2991, 0.499, 0.18114, 0.53043, 438.879, 52.828),
    'COCT': ('292-64-8', 112.215, 303.24, 692.33, 504.73, 6.3433, 0.259, 0.290432, -0.035087, 0.29293, 377.991, 51.949),
    'CP': ('287-92-3', 70.1344, 120.65, 461.5, 652.96, 4.1328, 0.273, 0.195884, -0.077035, 0.12985, 229.622, 46.278),
    'CYPR': ('75-19-4', 42.08064, -27.0, 256.57, 796.98, 2.6078, 0.27, 0.12688, -0.074805, 0.058913, 139.606, 37.582),
    'CISD': ('493-01-6', 138.2529, 384.47, 804.38, 469.92, 7.6889, 0.266, 0.293876, 0.15108, 0.53874, 475.007, 55.821),
    'TRND': ('493-02-7', 138.2529, 369.16, 777.02, 430.76, 7.6889, 0.25, 0.272435, 0.15108, 0.53874, 462.82, 54.172),
    'C10CYH': ('1795-16-0', 224.4301, 567.68, 892.58, 239.31, 13.744, 0.227, 0.662663, 0.30923, 0.85493, 665.336, 50.895),
    'DCCP': ('1795-21-7', 210.4032, 534.83, 843.17, 220.44, 13.039, 0.2056, 0.654, 0.27909, 0.79475, 624.377, 51.873),
    'DIAMTANE': ('2292-79-7', 188.3128, 492.53, 964.13, 327.79, 10.62, 0.228, 0.163289, 0.23239, 0.70112, 559.873, 52.753),
    '11CH': ('590-66-9', 112.215, 247.19, 604.4, 426.18, 7.2083, 0.269, 0.232569, -0.098329, 0.26498, 366.444, 48.49),
    '1C2H': ('2207-01-4', 112.215, 265.62, 631.4, 426.18, 7.3685, 0.268, 0.232443, -0.10096, 0.27158, 383.587, 49.458),
    '1T2H': ('6876-23-9', 112.215, 254.17, 613.4, 426.18, 7.3685, 0.273, 0.237864, -0.093628, 0.27234, 374.377, 48.198),
    '1C3H': ('638-04-0', 112.215, 248.16, 604.4, 426.18, 7.2083, 0.269, 0.236572, -0.094182, 0.26882, 366.444, 47.565),
    '1T3H': ('2207-03-6', 112.215, 256.03, 616.73, 426.18, 7.3685, 0.272, 0.23346, -0.098534, 0.26891, 376.097, 48.728),
    '1C4H': ('624-29-3', 112.215, 255.78, 617.0, 426.18, 7.3685, 0.272, 0.23114, -0.10099, 0.26673, 376.236, 48.612),
    '1T4H': ('2207-04-7', 112.215, 246.85, 602.6, 426.18, 7.2083, 0.269, 0.237028, -0.093551, 0.2688, 365.534, 47.368),
    '11CP': ('1638-26-2', 98.18816, 190.13, 524.93, 499.65, 5.7666, 0.273, 0.272354, -0.045075, 0.23923, 298.833, 46.819),
    '1C2C': ('1192-18-3', 98.18816, 211.15, 557.6, 499.66, 5.9268, 0.271, 0.266243, -0.051732, 0.24239, 315.311, 47.955),
    '1T2C': ('822-50-4', 98.18816, 197.37, 536.0, 499.66, 5.7666, 0.27, 0.269784, -0.047706, 0.23995, 302.97, 46.623),
    '1C3M': ('2532-58-3', 98.18816, 195.39, 532.13, 499.65, 5.7666, 0.271, 0.27426, -0.043834, 0.24243, 301.533, 46.211),
    '1T3M': ('1759-58-6', 98.18816, 197.11, 535.73, 499.65, 5.7666, 0.27, 0.269949, -0.047557, 0.24002, 302.87, 46.466),
    'DDCP': ('5634-30-0', 238.457, 591.71, 890.33, 188.11, 14.801, 0.1922, 0.719, 0.3382, 0.91321, 706.294, 52.884),
    'ECH': ('1678-91-7', 112.215, 269.23, 636.8, 440.91, 6.8879, 0.258, 0.245525, -0.084568, 0.27629, 374.074, 48.977),
    'ECP': ('1640-89-7', 98.18816, 218.25, 565.47, 492.84, 6.0069, 0.269, 0.271582, -0.048246, 0.25191, 320.478, 47.609),
    'HPCP': ('5617-42-5', 168.3225, 435.47, 762.53, 282.16, 10.396, 0.2236, 0.515, 0.19744, 0.63187, 501.502, 45.486),
    'HDCP': ('6812-39-1', 294.56451, 686.93, 964.13, 179.0, 18.325, 0.2146, 0.861, 0.48452, 1.1785, 870.128, 54.894),
    'HXCP': ('4457-00-5', 154.2957, 397.67, 728.51, 310.08, 9.515, 0.2314, 0.476, 0.18289, 0.61093, 460.543, 49.839),
    'IBCH': ('1678-98-4', 140.2688, 340.43, 726.53, 452.63, 8.4658, 0.301, 0.319, 0.15054, 0.54204, 461.382, 49.404),
    'IPCH': ('696-29-7', 126.2419, 310.57, 668.93, 413.36, 7.4326, 0.254, 0.329509, 0.0022998, 0.39138, 403.828, 49.833),
    'IPCP': ('3875-51-2', 112.215, 259.57, 607.73, 440.91, 6.7277, 0.259, 0.302996, -0.024594, 0.32238, 356.907, 48.223),
    'MCHPTAN': ('4126-78-7', 112.215, 273.2, 647.6, 471.37, 6.5363, 0.2593, 0.258, -0.067723, 0.27224, 368.172, 51.291),
    'MCH': ('108-87-2', 98.18816, 213.68, 570.27, 503.43, 5.8948, 0.269, 0.234982, -0.078524, 0.21902, 319.112, 47.817),
    'MCP': ('96-37-7', 84.16128, 161.26, 499.35, 548.97, 5.1099, 0.273, 0.230249, -0.070482, 0.18385, 272.371, 46.477),
    'MCYP': ('594-11-6', 56.10752, 40.41, 344.79, 649.26, 3.6842, 0.2771, 0.1495, -0.091505, 0.091984, 191.038, 39.869),
    '1M1E': ('16747-50-5', 112.215, 250.74, 587.93, 438.01, 6.8559, 0.267, 0.329811, 0.0023143, 0.34312, 350.858, 48.479),
    '1C2E': ('930-89-2', 112.215, 262.51, 613.4, 426.18, 6.7438, 0.2496, 0.29, -0.039158, 0.32274, 359.833, 48.766),
    '1T2E': ('930-90-5', 112.215, 250.16, 599.0, 426.18, 6.7438, 0.253, 0.267, -0.06234, 0.29647, 353.392, 47.753),
    '1C3E': ('2613-66-3', 112.215, 249.98, 597.2, 426.18, 6.7438, 0.2534, 0.281, -0.047828, 0.30929, 352.578, 47.627),
    '1T3E': ('2613-65-2', 112.215, 249.98, 597.2, 420.3, 6.7438, 0.2499, 0.276, -0.053714, 0.30879, 352.578, 47.627),
    'NOCP': ('2882-98-6', 196.3763, 503.87, 819.23, 239.54, 12.158, 0.2122, 0.61, 0.24758, 0.73258, 583.419, 51.368),
    'OCP': ('1795-20-6', 182.3494, 470.75, 789.53, 260.12, 11.277, 0.2188, 0.564, 0.22151, 0.67847, 542.46, 50.859),
    'PDCP': ('4669-01-6', 280.5376, 665.33, 944.33, 187.37, 17.444, 0.2169, 0.833, 0.4391, 1.1026, 829.17, 54.391),
    'PINENE': ('80-56-8', 136.237, 313.05, 699.53, 400.3, 7.2724, 0.234, 0.220949, 0.1526, 0.5366, 413.542, 53.506),
    'BPINENE': ('127-91-3', 136.237, 330.87, 697.73, 400.3, 7.9131, 0.255, 0.324932, 0.1526, 0.5366, 431.192, 54.088),
    'PCH': ('1678-92-8', 126.2419, 314.14, 690.8, 407.12, 7.6408, 0.252, 0.259535, -0.079671, 0.32916, 420.286, 49.331),
    'PRCP': ('2040-96-2', 112.215, 267.73, 613.13, 438.01, 6.8559, 0.261, 0.326642, -0.0008507, 0.34839, 362.566, 48.248),
    'TDCP': ('1795-22-8', 266.51071, 618.53, 929.93, 196.63, 16.563, 0.2183, 0.789, 0.39956, 1.0328, 788.211, 53.891),
    'TRCP': ('6006-34-4', 252.48379, 617.81, 910.13, 206.77, 15.682, 0.2205, 0.755, 0.36602, 0.96953, 747.253, 53.387),
    '1T35TMCH': ('1795-26-2', 126.2419, 284.99, 624.29, 385.8, 8.5324, 0.283, 0.333, 0.0064366, 0.40652, 400.616, 48.361),
    '112C': ('4259-00-1', 112.215, 236.71, 583.52, 426.18, 6.6797, 0.2543, 0.251, -0.077679, 0.2771, 344.868, 47.965),
    '113C': ('4516-69-2', 112.215, 220.82, 565.52, 410.02, 6.6797, 0.2489, 0.211, -0.12085, 0.24466, 336.582, 46.45),
    'CMCP': ('2613-69-6', 112.215, 253.4, 595.4, 426.18, 6.6637, 0.2508, 0.307, -0.021037, 0.33348, 349.826, 48.385),
    'CTCP': ('19374-46-0', 112.215, 243.5, 584.6, 420.3, 6.6637, 0.2499, 0.287, -0.041733, 0.3156, 344.987, 47.832),
    'CCCP': ('2613-72-1', 112.215, 242.24, 584.6, 418.83, 6.6637, 0.2491, 0.287, -0.041879, 0.31671, 344.987, 47.933),
    'CT4C': ('4850-28-6', 112.215, 242.13, 582.8, 417.36, 6.6637, 0.2486, 0.278, -0.051282, 0.30865, 344.174, 47.4),
    'TC4C': ('16883-48-0', 112.215, 228.72, 568.4, 407.81, 6.6637, 0.2463, 0.244, -0.087324, 0.27861, 337.584, 46.391),
    '33EP': ('1067-20-5', 128.2578, 295.14, 638.42, 387.97, 7.5767, 0.249, 0.33809, 0.16818, 0.53939, 391.88, 46.821),
    'DM3E': ('16747-32-3', 128.2578, 272.91, 602.33, 372.75, 8.1854, 0.268, 0.335271, 0.16818, 0.53939, 381.931, 45.64),
    'DMEP': ('1068-87-7', 128.2578, 278.1, 604.13, 366.95, 8.2014, 0.264, 0.353019, 0.16818, 0.53939, 383.27, 45.837),
    '22HP': ('1071-26-7', 128.2578, 270.84, 578.57, 340.84, 8.3136, 0.254, 0.389933, 0.16818, 0.53939, 368.186, 44.111),
    '26HP': ('1072-05-5', 128.2578, 275.38, 582.53, 333.59, 8.3296, 0.248, 0.392729, 0.16818, 0.53939, 370.948, 44.06),
    '22HX': ('590-73-8', 114.2309, 224.31, 529.97, 366.95, 7.6568, 0.265, 0.337807, 0.011558, 0.39519, 331.349, 43.187),
    '23HX': ('584-94-1', 114.2309, 240.1, 554.45, 381.45, 7.4998, 0.263, 0.347204, 0.02176, 0.39916, 344.322, 44.216),
    '24HX': ('589-43-5', 114.2309, 228.97, 536.63, 371.3, 7.5607, 0.263, 0.343599, 0.017948, 0.39915, 334.528, 43.273),
    '25HX': ('592-13-2', 114.2309, 228.4, 530.33, 361.14, 7.7209, 0.262, 0.357589, 0.03425, 0.4225, 332.085, 43.073),
    '33HX': ('563-16-6', 114.2309, 233.55, 551.93, 384.35, 7.093, 0.251, 0.320234, -0.0079126, 0.36791, 337.271, 44.127),
    '34HX': ('583-48-2', 114.2309, 243.91, 564.17, 390.15, 7.4678, 0.265, 0.338129, 0.011597, 0.38485, 349.45, 44.719),
    '22DMOCT': ('15869-87-1', 142.2847, 314.42, 623.93, 313.28, 9.1946, 0.248, 0.428759, 0.15097, 0.54648, 425.471, 44.997),
    '23DMOCT': ('7146-60-3', 142.2847, 327.76, 643.73, 320.53, 9.2106, 0.249, 0.432722, 0.15097, 0.54648, 425.471, 45.832),
    '24DMOCT': ('4032-94-4', 142.2847, 312.62, 620.33, 314.73, 9.2106, 0.25, 0.434479, 0.15097, 0.54648, 425.471, 45.097),
    '25DMOCT': ('15869-89-3', 142.2847, 317.3, 625.73, 314.73, 9.2106, 0.249, 0.44034, 0.15097, 0.54648, 425.471, 45.345),
    '26DMOCT': ('2051-30-1', 142.2847, 320.68, 631.13, 314.73, 9.2106, 0.248, 0.438586, 0.15097, 0.54648, 425.471, 45.174),
    '27DMOCT': ('1072-16-8', 142.2847, 319.77, 627.53, 308.93, 9.2106, 0.244, 0.441982, 0.15097, 0.54648, 425.471, 44.954),
    '22MP': ('590-35-2', 100.204, 174.54, 477.23, 402.23, 6.6637, 0.267, 0.287884, -0.03827, 0.29667, 292.439, 42.014),
    '23MP': ('565-59-3', 100.204, 193.61, 507.56, 421.77, 6.2953, 0.256, 0.295734, -0.030208, 0.29899, 302.56, 43.158),
    '24MP': ('108-08-7', 100.204, 176.89, 475.95, 396.97, 6.6957, 0.265, 0.301427, -0.025479, 0.31244, 292.054, 41.696),
    '33MP': ('562-49-2', 100.204, 186.91, 505.85, 427.21, 6.6316, 0.273, 0.267208, -0.056523, 0.26991, 306.849, 42.886),
    'NC22': ('629-97-0', 310.60721, 695.48, 956.93, 153.74, 20.664, 0.209, 0.97219, 0.54688, 1.2759, 916.973, 48.589),
    'NC12': ('112-40-3', 170.33839, 421.38, 724.73, 263.97, 11.469, 0.238, 0.576385, 0.19909, 0.6344, 507.388, 46.516),
    'NC32': ('544-85-4', 450.87601, 870.26, 1079.33, 108.78, 29.794, 0.196, 1.37655, 0.90614, 1.9004, 1326.558, 48.816),
    'NC20': ('112-95-8', 282.5535, 650.8, 922.73, 168.24, 18.742, 0.213, 0.906878, 0.44512, 1.113, 835.056, 48.5),
    '3EHP': ('15869-80-4', 128.2578, 289.76, 602.33, 346.64, 8.4577, 0.257, 0.407991, 0.16818, 0.53939, 385.341, 45.106),
    '3EHX': ('619-99-8', 114.2309, 245.37, 558.05, 378.55, 7.29, 0.253, 0.36278, 0.039381, 0.41972, 343.512, 44.3),
    '3EPN': ('617-78-7', 100.204, 200.25, 513.48, 419.27, 6.6637, 0.268, 0.309372, -0.017439, 0.31477, 311.127, 43.407),
    'NC21': ('629-94-7', 296.58041, 673.7, 940.73, 160.99, 19.703, 0.211, 0.942004, 0.49174, 1.1901, 876.015, 48.568),
    'NC27': ('593-49-7', 380.74161, 791.78, 1027.13, 128.07, 25.243, 0.203, 1.21357, 2.5216, 3.7066, 1121.765, 48.689),
    'NC17': ('629-78-7', 240.47279, 575.87, 865.13, 194.35, 16.018, 0.219, 0.769688, 0.34229, 0.92146, 712.181, 48.28),
    'NC26': ('630-01-3', 366.71481, 773.96, 1014.53, 131.98, 24.348, 0.203, 1.15444, 2.4257, 3.5689, 1080.807, 48.712),
    'NC16': ('544-76-3', 226.446, 548.36, 841.73, 203.05, 15.121, 0.22, 0.717404, 0.31347, 0.86343, 671.222, 48.049),
    'NC36': ('630-06-8', 506.98361, 926.6, 1113.53, 98.626, 33.479, 0.196, 1.52596, 3.7892, 5.3965, 1490.392, 48.809),
    '2M3E': ('609-26-7', 114.2309, 240.17, 560.93, 391.6, 7.093, 0.254, 0.329418, 0.0020968, 0.3735, 341.991, 44.726),
    '3M3E': ('1067-08-9', 114.2309, 244.89, 578.03, 407.56, 7.29, 0.267, 0.305027, -0.023721, 0.34105, 354.311, 45.179),
    '2MHP': ('592-27-8', 114.2309, 243.77, 547.68, 360.27, 7.817, 0.261, 0.377224, 0.057715, 0.45203, 343.712, 43.374),
    '3MHP': ('589-81-1', 114.2309, 246.07, 554.94, 369.27, 7.4326, 0.252, 0.371838, 0.050575, 0.43851, 343.742, 43.787),
    '4MHP': ('589-53-7', 114.2309, 243.88, 551.46, 368.69, 7.6248, 0.259, 0.370577, 0.049046, 0.43637, 344.047, 43.802),
    '2MHX': ('591-76-4', 100.204, 194.09, 495.0, 396.53, 6.7438, 0.261, 0.327728, 0.0002487, 0.34346, 302.641, 42.057),
    '3MHX': ('589-34-4', 100.204, 197.33, 503.78, 408.14, 6.4715, 0.255, 0.321624, -0.0057352, 0.33124, 303.566, 42.72),
    '2MNON': ('871-83-0', 142.2847, 332.6, 638.33, 307.48, 9.3388, 0.244, 0.472342, 0.15097, 0.54648, 425.471, 45.13),
    '3MN': ('5911-04-6', 142.2847, 334.04, 643.73, 313.28, 9.3388, 0.247, 0.464925, 0.15097, 0.54648, 425.471, 45.534),
    '4MNON': ('17301-94-9', 142.2847, 330.26, 638.33, 313.28, 9.3388, 0.248, 0.465073, 0.15097, 0.54648, 425.471, 45.453),
    '5MNON': ('15869-85-9', 142.2847, 329.27, 638.33, 313.28, 9.3388, 0.248, 0.456177, 0.15097, 0.54648, 425.471, 45.482),
    '2MO': ('3221-61-2', 128.2578, 289.9, 596.48, 332.14, 8.666, 0.254, 0.421169, 0.16818, 0.53939, 383.628, 44.302),
    '3MO': ('2216-33-3', 128.2578, 291.61, 602.6, 339.39, 8.4738, 0.252, 0.412346, 0.16818, 0.53939, 385.705, 44.745),
    '4MO': ('2216-34-4', 128.2578, 288.39, 598.1, 339.39, 8.3776, 0.25, 0.412925, 0.16818, 0.53939, 381.645, 44.722),
    '22PR': ('463-82-1', 72.15028, 49.1, 321.13, 463.98, 4.8632, 0.269, 0.196363, -0.091558, 0.15521, 205.428, 37.676),
    'NC29': ('630-03-5', 408.79541, 825.44, 1048.73, 119.8, 27.071, 0.2, 1.25453, 2.8124, 4.0939, 1203.683, 48.719),
    'NC19': ('629-92-5', 268.52661, 625.82, 904.73, 175.5, 17.941, 0.215, 0.852231, 0.40492, 1.0425, 794.098, 48.503),
    'NC28': ('630-02-4', 394.76849, 808.88, 1037.93, 123.28, 26.11, 0.2, 1.23752, 2.699, 3.937, 1162.724, 48.674),
    'NC18': ('593-45-3', 254.4997, 602.08, 884.93, 184.2, 16.98, 0.217, 0.811359, 0.37031, 0.97794, 753.139, 48.404),
    'NC25': ('629-99-2', 352.6879, 755.42, 1001.93, 137.79, 23.387, 0.205, 1.10526, 0.93046, 1.7891, 1039.849, 48.642),
    'NC15': ('629-62-9', 212.4191, 519.23, 814.73, 214.66, 14.24, 0.224, 0.68632, 0.28349, 0.80351, 630.264, 47.738),
    'PMHP': ('13475-82-6', 170.33839, 352.04, 665.42, 282.16, 11.117, 0.2598, 0.43, 0.19909, 0.6344, 507.388, 46.312),
    'SQUALANE': ('111-01-3', 422.8223, 836.33, 1093.73, 125.89, 26.27, 0.198, 1.05584, 2.9838, 4.3136, 1244.641, 50.341),
    'NC24': ('646-31-1', 338.66101, 736.34, 987.53, 142.14, 22.586, 0.207, 1.07102, 0.74337, 1.5452, 998.89, 48.609),
    'NC14': ('629-59-4', 198.3922, 488.44, 787.73, 227.71, 13.295, 0.226, 0.643017, 0.25191, 0.74122, 589.305, 47.403),
    'TEMB': ('594-82-1', 114.2309, 223.32, 562.73, 416.26, 7.3845, 0.28, 0.244953, -0.084109, 0.27235, 347.463, 40.82),
    'TEMH': ('13475-81-5', 142.2847, 320.63, 661.91, 364.46, 8.2953, 0.2512, 0.36, 0.15097, 0.54648, 425.471, 47.485),
    '2233': ('7154-79-2', 128.2578, 284.52, 639.86, 396.79, 7.6568, 0.257, 0.27997, 0.16818, 0.53939, 394.392, 47.008),
    '2234': ('1186-53-4', 128.2578, 271.45, 606.2, 371.81, 7.849, 0.255, 0.310598, 0.16818, 0.53939, 379.28, 45.898),
    '2244': ('1070-87-7', 128.2578, 252.12, 568.76, 342.42, 8.0733, 0.25, 0.315927, 0.16818, 0.53939, 359.457, 44.674),
    '2334': ('16747-38-9', 128.2578, 286.79, 634.19, 393.92, 7.8971, 0.265, 0.312682, 0.16818, 0.53939, 396.178, 46.888),
    'T55': ('69522-75-4', 340.0, 676.99, 944.6, 220.0, 17.107, 0.3713, 0.9433, 0.74847, 1.5533, 1002.8, None),
    'NC30': ('638-68-6', 422.8223, 841.46, 1059.53, 116.03, 28.032, 0.2, 1.30718, 2.9838, 4.3136, 1244.641, 48.836),
    'NC23': ('638-67-5', 324.63409, 716.36, 973.13, 147.94, 21.625, 0.208, 1.02617, 0.64457, 1.4098, 957.932, 48.615),
    'NC13': ('629-50-5', 184.3653, 455.84, 755.33, 243.66, 12.414, 0.232, 0.617397, 0.22547, 0.68652, 548.347, 47.049),
    '223B': ('464-06-2', 100.204, 177.58, 496.44, 428.39, 6.3753, 0.266, 0.250438, -0.071349, 0.25218, 298.626, 42.881),
    'TMHP': ('7154-80-5', 142.2847, 312.17, 637.61, 336.54, 8.6083, 0.246, 0.388, 0.15097, 0.54648, 425.471, 46.147),
    '224H': ('16747-26-5', 128.2578, 259.77, 572.72, 345.36, 8.1214, 0.2532, 0.35, 0.16818, 0.53939, 362.473, 44.421),
    '225H': ('3522-94-9', 128.2578, 255.36, 562.82, 337.94, 8.3136, 0.256, 0.359642, 0.16818, 0.53939, 357.72, 43.934),
    '233H': ('16747-28-7', 128.2578, 279.82, 613.22, 370.34, 7.8651, 0.253, 0.33, 0.16818, 0.53939, 383.61, 45.8),
    '235H': ('1069-53-0', 128.2578, 268.41, 582.8, 348.29, 8.1534, 0.2538, 0.364, 0.16818, 0.53939, 369.241, 44.823),
    '244H': ('16747-30-1', 128.2578, 267.19, 586.13, 356.79, 8.1854, 0.26, 0.352203, 0.16818, 0.53939, 371.723, 44.951),
    '334H': ('16747-31-2', 128.2578, 284.83, 624.38, 380.63, 7.7529, 0.2537, 0.325, 0.16818, 0.53939, 387.914, 46.292),
    '223P': ('564-02-3', 114.2309, 229.72, 554.63, 395.95, 6.984, 0.254, 0.297083, -0.03231, 0.33532, 336.93, 44.474),
    '224P': ('540-84-1', 114.2309, 210.63, 519.46, 372.46, 7.4966, 0.266, 0.302204, -0.027571, 0.34928, 323.631, 43.101),
    '233P': ('560-21-4', 114.2309, 238.59, 572.63, 409.01, 7.29, 0.269, 0.2903, -0.038938, 0.3238, 351.421, 45.09),
    '234P': ('565-75-3', 114.2309, 236.25, 559.67, 395.95, 7.3781, 0.267, 0.316137, -0.012116, 0.35581, 345.668, 44.705),
    'NC11': ('1120-21-4', 156.3116, 384.67, 690.53, 282.82, 10.556, 0.242, 0.530316, 0.18842, 0.62078, 466.43, 45.977),
    'BU1THIOL': ('109-79-5', 90.1894, 209.23, 566.51, 575.8, 4.9177, 0.257, 0.271361, -0.0415, 0.21568, 284.058, 52.223),
    '1MPR1THL': ('513-53-1', 90.1894, 184.96, 537.53, 588.85, 4.9177, 0.271, 0.25059, -0.05403, 0.19144, 276.777, 51.452),
    '2MPR2THL': ('75-66-1', 90.1894, 147.6, 494.33, 588.85, 4.9177, 0.283, 0.191395, -0.091484, 0.14633, 265.05, 49.595),
    'CS2': ('75-15-0', 76.14301, 115.2, 533.93, 1145.8, 2.563, 0.275, 0.110697, -0.078005, 0.051459, 137.191, 78.39),
    'PN1THIOL': ('110-66-7', 104.2163, 259.95, 616.73, 503.28, 5.7506, 0.251, 0.320705, -0.0060119, 0.29935, 330.292, 52.247),
    'BZTHIOL': ('108-98-5', 110.1796, 336.45, 780.53, 687.48, 5.0458, 0.261, 0.262789, -0.048424, 0.21239, 333.769, 66.974),
    'PR1THIOL': ('107-03-9', 76.16252, 153.9, 506.21, 671.52, 4.0687, 0.264, 0.231789, -0.057111, 0.15222, 234.851, 52.152),
    'THIOPHEN': ('110-02-1', 84.14175, 183.49, 583.16, 825.26, 3.508, 0.259, 0.192846, -0.070594, 0.11484, 210.917, 66.092),
    'ACETYLN': ('74-86-2', 26.03788, -118.84, 95.31, 890.39, 1.8101, 0.271, 0.1873, -0.036255, 0.055318, 95.908, 38.511),
    'CAMPHENE': ('79-92-5', 136.237, 320.9, 688.73, 398.85, 7.9932, 0.259, 0.295744, 0.1526, 0.5366, 428.451, 52.916),
    'CYHEPTEN': ('628-92-2', 96.17228, 237.83, 616.73, 581.6, 5.3822, 0.271, 0.251706, -0.058191, 0.21002, 316.055, 51.327),
    '13CD': ('592-57-4', 80.12952, 176.61, 544.73, 686.03, 4.4371, 0.282, 0.230835, -0.058713, 0.15441, 258.635, 52.266),
    '14CD': ('628-41-1', 80.12952, 188.6, 564.53, 686.03, 4.4371, 0.277, 0.20483, -0.07598, 0.14255, 262.349, 53.098),
    'CYHEXENE': ('110-83-8', 82.14539, 181.35, 549.05, 630.91, 4.6614, 0.272, 0.212302, -0.076412, 0.15724, 269.358, 50.337),
    '15COD': ('111-78-4', 108.1833, 302.22, 701.33, 565.65, 5.8628, 0.266, 0.285813, -0.035487, 0.25979, 360.633, 54.799),
    'CYLOCTEN': ('931-88-4', 110.1992, 289.4, 677.93, 533.74, 6.103, 0.267, 0.280425, -0.041615, 0.26536, 363.62, 52.788),
    'CD13': ('542-92-7', 66.10264, 106.7, 452.93, 746.94, 3.6042, 0.275, 0.202005, -0.063612, 0.11533, 205.106, 49.784),
    'CYPENTEN': ('142-29-0', 68.11852, 111.61, 452.93, 696.47, 3.9245, 0.279, 0.195859, -0.071564, 0.1206, 219.596, 47.853),
    '1DECENE': ('872-05-9', 140.2688, 339.08, 649.85, 321.69, 9.3548, 0.253, 0.479961, 0.15054, 0.54204, 428.69, 46.031),
    'DCPD': ('77-73-6', 132.20531, 337.73, 728.33, 443.82, 7.1282, 0.248, 0.288011, 0.15857, 0.53576, 421.203, 61.254),
    '2BUTYNE': ('503-17-3', 54.09164, 80.56, 392.09, 706.33, 3.5401, 0.274, 0.238542, -0.044503, 0.13075, 193.99, 42.797),
    '23DM13BD': ('513-81-5', 82.14539, 155.8, 487.13, 510.53, 5.0458, 0.254, 0.214241, -0.087141, 0.18377, 266.769, 45.122),
    '23M1BUTE': ('563-78-0', 84.16128, 132.1, 440.33, 467.02, 5.5904, 0.27, 0.226929, -0.080405, 0.20035, 261.94, 42.022),
    '23M2BUTE': ('563-79-1', 84.16128, 163.76, 483.53, 458.32, 5.9589, 0.27, 0.233283, -0.080438, 0.21897, 286.147, 43.914),
    '33M1BUTE': ('558-37-2', 84.16128, 106.24, 404.33, 477.17, 5.3341, 0.275, 0.225718, -0.076456, 0.1874, 243.588, 40.449),
    '25DM15HD': ('627-58-7', 110.1992, 237.74, 560.93, 397.4, 7.1282, 0.259, 0.323776, -0.0039462, 0.36248, 342.568, 46.053),
    '25DM24HD': ('764-13-6', 110.1992, 275.47, 614.93, 397.4, 6.8078, 0.235, 0.321841, -0.0063262, 0.37965, 362.167, 47.404),
    '23DM1HX': ('16746-86-4', 112.215, 230.9, 550.13, 400.3, 7.1282, 0.263, 0.325141, -0.0024474, 0.35737, 336.849, 44.773),
    'DDC1': ('112-41-4', 168.3225, 416.04, 724.01, 279.92, 10.957, 0.241, 0.570505, 0.19744, 0.63187, 501.502, 47.137),
    '1EICOSE': ('3452-07-1', 280.5376, 648.3, 928.13, 176.95, 18.581, 0.221, 0.880365, 0.4391, 1.1026, 829.17, 49.255),
    '3MT2PNTE': ('616-12-6', 84.16128, 158.79, 470.93, 462.67, 5.4943, 0.255, 0.262659, -0.0541, 0.2367, 272.276, 43.268),
    '1BUTYNE': ('107-00-6', 54.09164, 46.53, 338.09, 717.94, 3.5561, 0.298, 0.246864, -0.03717, 0.12403, 185.734, 41.715),
    '2E1BUTEN': ('760-21-4', 84.16128, 148.41, 461.93, 458.32, 5.8307, 0.27, 0.227716, -0.083241, 0.20966, 274.754, 42.75),
    '15ED': ('40085-10-7', 108.1833, 271.4, 640.4, 496.72, 6.1831, 0.2602, 0.292, -0.032604, 0.28557, 353.694, 52.031),
    '2E1HEXEN': ('1632-16-2', 112.215, 248.0, 573.53, 445.27, 6.3914, 0.257, 0.379909, 0.050482, 0.37752, 333.658, 45.124),
    '2E1PNTEN': ('3404-71-5', 98.18816, 201.2, 517.73, 427.86, 6.3753, 0.26, 0.308512, -0.017979, 0.30904, 308.611, 43.948),
    '3E1PNTEN': ('4038-04-4', 98.18816, 183.4, 494.33, 439.46, 6.3753, 0.274, 0.301588, -0.023322, 0.28791, 297.62, 43.182),
    'PHACETYL': ('536-74-3', 102.1356, 289.13, 710.33, 620.76, 5.3181, 0.263, 0.226363, -0.079082, 0.19554, 336.277, 57.684),
    '1HPTADCN': ('6765-39-5', 238.457, 572.59, 865.13, 204.5, 15.89, 0.229, 0.750312, 0.3382, 0.91321, 706.294, 48.786),
    '1HEPTENE': ('592-76-7', 98.18816, 200.55, 507.45, 410.46, 6.6156, 0.262, 0.331019, 0.003461, 0.33912, 307.433, 43.255),
    'C2HEPTEN': ('6443-92-1', 98.18816, 209.14, 528.53, 411.91, 6.7918, 0.264, 0.294154, -0.033175, 0.31134, 320.588, 43.871),
    'T2HEPTEN': ('14686-13-6', 98.18816, 208.31, 517.73, 413.36, 6.5035, 0.256, 0.337166, 0.0095045, 0.3459, 310.769, 43.508),
    'C3HEPTEN': ('7642-10-6', 98.18816, 204.35, 521.33, 411.91, 6.7438, 0.264, 0.294912, -0.032184, 0.30976, 316.232, 43.606),
    'T3HEPTEN': ('14686-14-7', 98.18816, 204.21, 512.33, 413.36, 6.5035, 0.258, 0.334056, 0.0064174, 0.34118, 308.179, 43.312),
    '1HEPTYNE': ('628-71-7', 96.17228, 211.6, 546.53, 455.42, 6.1831, 0.261, 0.272805, -0.050124, 0.26861, 317.352, 45.493),
    'HDC1': ('629-73-2', 224.4301, 544.77, 839.93, 214.66, 14.945, 0.23, 0.724224, 0.30923, 0.85493, 665.336, 48.593),
    '12HD': ('592-44-9', 82.14539, 168.8, 487.13, 485.88, 5.3021, 0.254, 0.276228, -0.041441, 0.23947, 273.668, 44.468),
    '14HD': ('592-45-0', 82.14539, 149.0, 458.33, 485.88, 5.3021, 0.262, 0.279951, -0.037262, 0.23489, 263.594, 43.401),
    '15HD': ('592-42-7', 82.14539, 139.03, 452.93, 485.88, 5.4303, 0.269, 0.232144, -0.074302, 0.19903, 264.181, 42.963),
    'CT24HDIE': ('5194-50-3', 82.14539, 182.3, 508.73, 485.88, 5.3021, 0.248, 0.275367, -0.043098, 0.24427, 280.87, 44.864),
    '24HD': ('5194-51-4', 82.14539, 179.42, 503.33, 485.88, 5.3021, 0.249, 0.282106, -0.037316, 0.24804, 279.097, 44.341),
    '1HEXENE': ('592-41-6', 84.16128, 146.26, 447.58, 455.42, 5.6705, 0.265, 0.280435, -0.038888, 0.24803, 266.186, 41.776),
    'C2HEXENE': ('7688-21-3', 84.16128, 155.98, 463.73, 458.32, 5.7506, 0.266, 0.272185, -0.046227, 0.24447, 274.217, 42.607),
    'T2HEXENE': ('4050-45-7', 84.16128, 154.17, 463.73, 458.32, 5.7666, 0.267, 0.261253, -0.055366, 0.23601, 274.484, 42.018),
    'C3HEXENE': ('7642-09-3', 84.16128, 151.61, 456.53, 459.77, 5.6225, 0.263, 0.278699, -0.040335, 0.24678, 269.079, 42.152),
    'T3HEXENE': ('13269-52-8', 84.16128, 152.76, 456.53, 459.77, 5.6225, 0.263, 0.285358, -0.034829, 0.25187, 269.079, 41.984),
    '1HEXYNE': ('693-02-7', 82.14539, 160.39, 469.49, 525.04, 5.1579, 0.272, 0.332699, 0.0038332, 0.25585, 264.133, 44.425),
    '2HEXYNE': ('764-35-2', 82.14539, 184.14, 528.53, 511.98, 5.3021, 0.256, 0.221387, -0.08497, 0.19656, 287.217, 45.374),
    '3HEXYNE': ('928-49-4', 82.14539, 178.16, 519.53, 511.98, 5.3021, 0.258, 0.218301, -0.086645, 0.1925, 284.361, 44.837),
    'ISOPRENE': ('78-79-5', 68.11852, 93.3, 411.53, 558.4, 4.4211, 0.264, 0.158323, -0.10951, 0.12111, 227.727, 42.197),
    'DLIMENE': ('5989-27-5', 136.237, 351.41, 715.73, 409.01, 7.5287, 0.244, 0.380863, 0.1526, 0.5366, 429.062, 52.427),
    'M12B': ('598-25-4', 68.11852, 105.53, 422.33, 555.49, 4.6614, 0.274, 0.187439, -0.092265, 0.141, 237.03, 42.492),
    '2M1BUTEN': ('563-46-2', 70.1344, 88.08, 377.33, 493.13, 4.6774, 0.257, 0.229982, -0.068668, 0.17845, 223.43, 40.272),
    '2M2BUTEN': ('513-35-9', 70.1344, 101.4, 388.13, 493.13, 4.6774, 0.254, 0.275316, -0.037213, 0.21068, 226.918, 40.819),
    '3M1BUTEN': ('563-45-1', 70.1344, 68.1, 351.0, 510.53, 4.8392, 0.284, 0.229705, -0.064422, 0.16678, 216.692, 39.159),
    '2M1B3YNE': ('78-80-8', 66.10264, 90.05, 425.93, 635.27, 3.9726, 0.266, 0.137046, -0.11016, 0.096828, 216.613, 43.638),
    '3M1BTYNE': ('598-23-2', 68.11852, 84.2, 374.09, 609.16, 4.4051, 0.3, 0.308085, -0.011015, 0.18494, 216.962, 41.204),
    'MCYPD': ('26519-91-5', 80.12952, 163.0, 514.13, 642.52, 4.4691, 0.275, 0.237559, -0.05655, 0.16375, 253.82, 50.254),
    '1MCPEN': ('693-89-0', 82.14539, 167.88, 515.93, 599.01, 4.8536, 0.278, 0.23179, -0.064669, 0.17237, 268.824, 48.392),
    '3MCPEN': ('1120-62-3', 82.14539, 148.82, 487.13, 599.01, 4.8536, 0.286, 0.229606, -0.064193, 0.16595, 261.014, 47.368),
    '4MCPEN': ('1759-81-5', 82.14539, 150.21, 488.93, 599.01, 4.8536, 0.286, 0.235746, -0.06028, 0.17, 261.516, 47.654),
    '2M1HPTEN': ('15870-10-7', 112.215, 246.6, 560.93, 377.1, 7.2564, 0.25, 0.35513, 0.031053, 0.41458, 344.58, 44.718),
    '2H24': ('28823-41-8', 96.17228, 232.7, 588.2, 495.25, 6.2472, 0.2751, 0.271, -0.049586, 0.25577, 335.944, 46.228),
    '2M1HEXEN': ('6094-02-6', 98.18816, 197.31, 508.73, 416.26, 6.3753, 0.255, 0.309375, -0.017477, 0.3155, 304.431, 43.604),
    '3M1HEXEN': ('3404-61-3', 98.18816, 183.02, 490.73, 427.86, 6.3753, 0.267, 0.305734, -0.020043, 0.29814, 295.894, 42.894),
    '4M1HEXEN': ('3769-23-1', 98.18816, 188.11, 501.53, 440.91, 6.3753, 0.273, 0.302408, -0.022679, 0.28981, 301.045, 43.339),
    '2M1PNTEN': ('763-29-1', 84.16128, 143.78, 452.93, 458.32, 5.7506, 0.269, 0.24062, -0.071767, 0.21748, 269.67, 42.125),
    '2M2PNTEN': ('625-27-4', 84.16128, 153.14, 465.53, 458.32, 5.8147, 0.268, 0.244478, -0.069526, 0.22347, 276.034, 42.534),
    '3M1PNTEN': ('760-20-3', 84.16128, 129.52, 431.33, 477.17, 5.4991, 0.274, 0.264031, -0.04916, 0.22072, 256.863, 41.375),
    '4M1PNTEN': ('691-37-2', 84.16128, 128.95, 433.13, 467.02, 5.5264, 0.269, 0.238917, -0.070252, 0.20755, 258.009, 41.147),
    '4MT2PNTE': ('674-76-0', 84.16128, 137.48, 442.13, 467.02, 5.5424, 0.267, 0.255237, -0.057884, 0.22174, 261.916, 41.455),
    'NONDE1': ('18435-45-5', 266.51071, 624.24, 908.33, 185.65, 17.62, 0.223, 0.832458, 0.39956, 1.0328, 788.211, 49.213),
    '1NONENE': ('124-11-8', 126.2419, 296.36, 608.18, 337.94, 8.4577, 0.249, 0.417123, 0.11754, 0.55925, 389.168, 45.28),
    'NORBRENE': ('498-66-8', 94.15639, 203.9, 589.73, 570.0, 5.3982, 0.273, 0.159, -0.12871, 0.14339, 309.029, 49.134),
    'OCD1': ('112-88-9', 252.48379, 598.68, 886.73, 194.35, 16.819, 0.226, 0.794275, 0.36602, 0.96953, 747.253, 48.927),
    '1OCTENE': ('111-66-0', 112.215, 250.32, 560.3, 372.46, 7.3685, 0.251, 0.376405, 0.055596, 0.44186, 345.883, 44.338),
    'C2OCTENE': ('7642-04-8', 112.215, 258.15, 569.93, 375.65, 7.3845, 0.251, 0.388932, 0.069892, 0.45542, 351.448, 44.954),
    'T2OCTENE': ('13389-42-9', 112.215, 257.0, 578.93, 374.2, 7.7529, 0.26, 0.338412, 0.012592, 0.40735, 361.896, 44.673),
    'C3OCTENE': ('14850-22-7', 112.215, 253.22, 564.53, 375.65, 7.3845, 0.252, 0.379994, 0.059413, 0.44368, 348.463, 44.761),
    'T3OCTENE': ('14919-01-8', 112.215, 253.94, 573.53, 374.2, 7.6889, 0.26, 0.343849, 0.018755, 0.41099, 357.858, 44.391),
    'C4OCTENE': ('7642-15-1', 112.215, 252.57, 562.73, 375.65, 7.3845, 0.253, 0.385401, 0.065415, 0.44855, 347.463, 44.761),
    'T4OCTENE': ('14850-23-8', 112.215, 252.07, 571.73, 374.2, 7.6889, 0.26, 0.339318, 0.01354, 0.40548, 356.802, 44.317),
    '1OCTYNE': ('629-05-0', 110.1992, 259.16, 593.33, 409.01, 7.0641, 0.256, 0.323356, -0.0044047, 0.36297, 357.909, 46.361),
    'PDC1': ('13360-61-7', 210.4032, 515.23, 814.73, 227.71, 14.016, 0.233, 0.681504, 0.27909, 0.79475, 624.377, 48.266),
    '12PD': ('591-95-7', 68.11852, 112.75, 440.33, 551.14, 4.4211, 0.252, 0.154153, -0.11745, 0.12415, 235.122, 42.926),
    'PPLN': ('1574-41-0', 68.11852, 111.33, 438.53, 542.44, 4.4211, 0.249, 0.146995, -0.12401, 0.12134, 234.675, 42.836),
    'TPND': ('2004-70-8', 68.11852, 107.64, 440.33, 542.44, 4.4211, 0.248, 0.116159, -0.14549, 0.10194, 235.122, 41.891),
    '14PD': ('591-93-5', 68.11852, 78.73, 402.53, 542.44, 4.8536, 0.285, 0.083652, -0.16082, 0.077819, 235.086, 40.792),
    '23PD': ('591-96-8', 68.11852, 118.85, 434.93, 551.14, 4.7254, 0.271, 0.218362, -0.073494, 0.16341, 242.368, 43.072),
    'C2PENTEN': ('627-20-3', 70.1344, 98.47, 395.69, 527.94, 4.6774, 0.269, 0.245236, -0.055292, 0.17983, 229.311, 40.571),
    'T2PENTEN': ('646-04-8', 70.1344, 97.41, 393.89, 530.84, 4.6934, 0.272, 0.248372, -0.052782, 0.18041, 229.075, 40.139),
    '1P3YNE': ('646-05-9', 66.10264, 138.65, 476.33, 638.17, 4.1007, 0.261, 0.252309, -0.045749, 0.16678, 231.118, 45.812),
    '1P4YNE': ('871-28-3', 66.10264, 108.5, 445.73, 638.17, 4.1007, 0.269, 0.178636, -0.087625, 0.1212, 225.282, 45.223),
    '1PENTYNE': ('627-19-0', 68.11852, 104.32, 406.49, 604.81, 4.4371, 0.289, 0.289925, -0.022317, 0.18353, 226.801, 42.955),
    '2PENTYNE': ('627-21-4', 68.11852, 133.02, 474.53, 584.5, 4.4211, 0.258, 0.175199, -0.10099, 0.13442, 243.253, 44.041),
    'APHELDR': ('99-83-2', 136.237, 347.0, 708.53, 409.01, 7.833, 0.256, 0.380709, 0.1526, 0.5366, 434.59, 52.474),
    'BPHELDR': ('555-10-2', 136.237, 345.2, 706.73, 409.01, 7.801, 0.255, 0.376605, 0.1526, 0.5366, 432.771, 52.23),
    '1PLIND': ('1961-96-2', 192.2603, 638.33, 1097.33, 400.3, 10.044, 0.241, 0.47032, 0.23955, 0.71612, 571.4, 67.151),
    'PROPDIEN': ('463-49-0', 40.06475, -30.1, 248.0, 738.24, 2.643, 0.257, 0.131481, -0.077961, 0.064533, 140.458, 41.386),
    'PROPYNE': ('74-99-7', 40.06475, -9.78, 264.63, 816.27, 2.627, 0.276, 0.216117, -0.041003, 0.088567, 141.046, 41.955),
    'ATERP': ('99-86-5', 136.237, 347.0, 708.53, 406.11, 7.833, 0.254, 0.372078, 0.1526, 0.5366, 434.59, 52.32),
    'GTERP': ('99-85-4', 136.237, 361.4, 730.13, 406.11, 7.833, 0.249, 0.376166, 0.1526, 0.5366, 445.447, 52.725),
    'TERPOLIN': ('586-62-9', 136.237, 368.33, 740.93, 409.01, 7.8971, 0.251, 0.376151, 0.1526, 0.5366, 452.792, 53.483),
    'TED1': ('1120-36-1', 196.3763, 483.98, 785.93, 240.76, 13.087, 0.236, 0.644885, 0.24758, 0.73258, 583.419, 47.961),
    'TDC1': ('2437-56-1', 182.3494, 451.0, 755.33, 256.72, 12.11, 0.238, 0.60634, 0.22151, 0.67847, 542.46, 47.636),
    '233M1BTE': ('594-56-9', 98.18816, 172.2, 496.13, 455.42, 6.103, 0.271, 0.240641, -0.075624, 0.22924, 294.129, 43.736),
    '244TM1P': ('107-39-1', 112.215, 214.59, 535.73, 381.45, 7.4486, 0.266, 0.269544, -0.062733, 0.31399, 332.859, 44.361),
    '244TM2P': ('107-40-4', 112.215, 220.84, 544.73, 381.45, 7.5287, 0.266, 0.26497, -0.068298, 0.31221, 339.0, 44.78),
    'UND1': ('821-95-4', 154.2957, 378.81, 688.73, 294.43, 10.284, 0.246, 0.517489, 0.18289, 0.61093, 460.543, 46.607),
    'VAC': ('689-97-4', 52.07576, 41.18, 357.53, 704.88, 3.2838, 0.264, 0.106852, -0.10614, 0.067086, 178.653, 43.761),
    'VCHXEN': ('100-40-3', 108.1833, 262.13, 618.53, 497.48, 6.071, 0.261, 0.32936, 0.0016916, 0.31055, 342.05, 51.582),
    'VINORBOR': ('3048-64-4', 120.1943, 284.9, 667.13, 465.57, 6.8399, 0.263, 0.248691, -0.079125, 0.27184, 385.789, 55.077),
    '3MC2PNTE': ('922-62-3', 84.16128, 153.86, 467.33, 477.2, 5.4943, 0.264, 0.258474, -0.055623, 0.22548, 270.913, 42.985),
    '4MC2PNTE': ('691-38-3', 84.16128, 133.48, 438.53, 467.02, 5.5424, 0.269, 0.244178, -0.066479, 0.21269, 260.457, 41.485),
    'NH3': (None, 17.03, -28.212, 270.318, 1635.576, 1.28788, None, 0.254999, None, None, None, 38.46),
    'HCL': (None, 36.4608, -121.09, 124.79, 1203.813, 1.29589, None, 0.133, None, None, None, 54.333),
    'SO2': (None, 64.063, 14.09, 315.768, 1141.447, 1.95698, None, 0.256, None, None, None, 87.122),
    'SO3': (None, 80.0578, 112.73, 424.13, 1190.76, 2.03835, None, 0.481, None, None, None, 117.376),
    'CL2': (None, 70.9057, -29.112, 290.746, 1122.592, 1.98309, None, 0.09, None, None, None, 97.47),
    'ARGON': (None, 39.948, -302.531, -188.397, 705.405, 1.19978, None, -0.004, None, None, None, 85.538),
    'CYCLOBUTENE': (None, 54.09, 36.768, 343.668, 764.349, 3.13161, None, 0.188997, None, None, None, 44.346),
    '12PENTADIENE': (None, 68.119, 112.717, 445.996, 590.792, 4.42029, None, 0.17294, None, None, None, 43.47),
    '15-HEXADIENE': (None, 82.1455, 139.01, 477.999, 499.697, 5.33847, None, 0.15997, None, None, None, 43.427),
    '26-OCTADIENE': (None, 110.194, 256.096, 650.93, 415.897, 7.60252, None, 0.12347, None, None, None, 46.62),
    'M-MERCAPTAN': (None, 48.1068, 42.703, 386.325, 1049.29, 2.32204, None, 0.155, None, None, None, 53.887),
    'E-MERCAPTAN': (None, 62.1338, 95.09, 438.525, 796.519, 3.31582, None, 0.18998, None, None, None, 52.221),
    'T-B-MERCAPTA': (None, 90.189, 147.592, 476.33, 588.847, 4.91767, None, 0.190696, None, None, None, 50.243),
    'NBMERCAPTAN': (None, 90.189, 209.226, 560.93, 575.8, 4.91767, None, 0.2784, None, None, None, 52.809),
    'BIACETYLENE': (None, 50.06, 50.63, 398.925, 849.921, 2.93939, None, 0.1, None, None, None, 44.777),
    '3-M-1-BUTYNE': (None, 68.12, 79.43, 397.13, 610.609, 4.349, None, 0.175999, None, None, None, 37.378),
    '1-PENTYNE': (None, 68.12, 104.268, 428.625, 587.403, 4.45312, None, 0.164, None, None, None, 43.381),
    '2-PENTYNE': (None, 68.12, 132.89, 479.93, 613.509, 4.44512, None, 0.186, None, None, None, 44.653),
    'HELIUM': (None, 4.00298, -452.092, -450.328, 32.919, 0.91786, None, -0.390032, None, None, None, 7.745),
    'NO': (None, 30.006, -241.15, -135.67, 938.394, 0.92427, None, 0.58797, None, None, None, 4.453),
    'NO2': (None, 46.006, 70.068, 316.125, 1450.377, 1.29269, None, 0.83398, None, None, None, 91.754),
    'N2O': (None, 44.013, -127.212, 97.61, 1050.073, 1.5602, None, 0.165, None, None, None, 50.996),
    'COS': (None, 60.0699, -58.27, 222.168, 896.305, 2.27462, None, 0.10407, None, None, None, 63.389),
}
# other names (PRO/II, databank, common) → PURE_DB key; looked up after stripping to A-Z 0-9 =
PURE_ALIAS = {
    '12BD': '1BUTYNE', '12BUTADIENE': '1BUTYNE', '12PENTADIENE': '12PENTADIENE', '13BD': '13BD',
    '13BUTADIENE': '13BD', '14PD': '14PD', '14PENTADIENE': '14PD', '15HEXADIENE': '15-HEXADIENE',
    '1BUTENE': '1BUTENE', '1BUTYNE': '1BUTYNE', '1CI3C5==': 'PPLN', '1HEPTENE': '1HEPTENE',
    '1HEPTYNE': '1HEPTYNE', '1HEXENE': '4MT2PNTE', '1HEXYNE': '4MCPEN', '1M1E': '1M1E',
    '1NONENE': '1NONENE', '1OCTENE': '1OCTENE', '1OCTYNE': '1OCTYNE', '1PENTENE': '1PENTENE',
    '1PENTYNE': '1-PENTYNE', '1TR3C5==': '23PD', '23PD': '23PD', '23PENTADIENE': 'CYPENTEN',
    '26OCTADIENE': '26-OCTADIENE', '2BUTYNE': '2BUTYNE', '2E1C4=': 'C3HEXENE', '2M1C5=': '2M1PNTEN',
    '2M1PNTEN': '2M1PNTEN', '2PENTYNE': '2-PENTYNE', '3M1BUTYNE': '3-M-1-BUTYNE',
    '3M1C5=': '3M1PNTEN', '3M1PNTEN': '3M1PNTEN', '3MT2PNTE': '3MT2PNTE', '4MCPEN': '4MCPEN',
    '4MT2PNTE': '4MT2PNTE', 'ACETYLENE': 'ACETYLN', 'ACETYLN': 'ACETYLN', 'AMMONIA': 'NH3',
    'AR': 'ARGON', 'ARGON': 'ARGON', 'BENZENE': 'BENZENE', 'BIACETYLENE': 'BIACETYLENE',
    'BUTADIENE': '13BD', 'BUTANE': 'BUTANE', 'C2BUTENE': 'C2BUTENE', 'C2H6': 'ETHANE',
    'C2PENTEN': 'C2PENTEN', 'C3H8': 'PROPANE', 'C3HEXENE': 'C3HEXENE', 'CARBONDIOXIDE': 'CO2',
    'CARBONDISULFIDE': 'CS2', 'CARBONMONOXIDE': 'CO', 'CARBONYLSULFIDE': 'COS', 'CH': 'CH',
    'CH4': 'METHANE', 'CIS2BUTENE': 'C2BUTENE', 'CIS2HEXENE': '3MT2PNTE', 'CIS2OCTENE': '1M1E',
    'CIS2PENTENE': 'C2PENTEN', 'CIS3HEXENE': '3MT2PNTE', 'CIS4OCTENE': 'T2OCTENE', 'CL2': 'CL2',
    'CO': 'CO', 'CO2': 'CO2', 'COS': 'COS', 'CP': 'CP', 'CS2': 'CS2', 'CT4C': 'CT4C',
    'CTCP': 'CTCP', 'CYBU': 'CYBU', 'CYCLOBUTANE': 'CYBU', 'CYCLOBUTENE': 'CYCLOBUTENE',
    'CYCLOHEXANE': 'CH', 'CYCLOPENTANE': 'CP', 'CYCLOPENTENE': 'CYPENTEN', 'CYCLOPROPANE': 'CYPR',
    'CYPENTEN': 'CYPENTEN', 'CYPR': 'CYPR', 'DMACETYLENE': '2BUTYNE', 'EACETYLENE': '1BUTYNE',
    'EBENZENE': 'MXYLENE', 'ECP': 'ECP', 'ECYCLOPENTAN': 'ECP', 'EMERCAPTAN': 'E-MERCAPTAN',
    'ETHANE': 'ETHANE', 'ETHANETHIOL': 'E-MERCAPTAN', 'ETHENE': 'ETHYLENE',
    'ETHYLBENZENE': 'MXYLENE', 'ETHYLENE': 'ETHYLENE', 'ETHYLMERCAPTAN': 'E-MERCAPTAN',
    'ETHYNE': 'ACETYLN', 'H2': 'H2', 'H2O': 'H2O', 'H2S': 'H2S', 'HCL': 'HCL', 'HE': 'HELIUM',
    'HELIUM': 'HELIUM', 'HEPTANE': 'HEPTANE', 'HEXANE': 'HEXANE', 'HYDROGEN': 'H2',
    'HYDROGENSULFIDE': 'H2S', 'IBUTANE': 'IBUTANE', 'IBUTENE': 'IBUTENE', 'IC4': 'IBUTANE',
    'IC5': 'IPENTANE', 'IPENTANE': 'IPENTANE', 'ISOBUTENE': 'IBUTENE', 'ISOBUTYLENE': 'IBUTENE',
    'ISOHEXANE': '2MP', 'MACETYLENE': 'PROPYNE', 'MCH': 'MCH', 'MCP': 'MCP', 'MCYCLOHEXANE': 'MCH',
    'MCYCLOPENTAN': 'MCP', 'METHANE': 'METHANE', 'METHANETHIOL': 'M-MERCAPTAN',
    'METHYLBENZENE': 'TOLUENE', 'METHYLCYCLOHEXANE': 'MCH', 'METHYLCYCLOPENTANE': 'MCP',
    'METHYLMERCAPTAN': 'M-MERCAPTAN', 'MMERCAPTAN': 'M-MERCAPTAN', 'MXYLENE': 'MXYLENE', 'N2': 'N2',
    'N2O': 'N2O', 'NBMERCAPTAN': 'NBMERCAPTAN', 'NBUTANE': 'BUTANE', 'NC4': 'BUTANE',
    'NC5': 'PENTANE', 'NC6': 'HEXANE', 'NC7': 'HEPTANE', 'NC8': 'OCTANE', 'NH3': 'NH3',
    'NHEPTANE': 'HEPTANE', 'NHEXANE': 'HEXANE', 'NITROGEN': 'N2', 'NITROUSOXIDE': 'N2O', 'NO': 'NO',
    'NO2': 'NO2', 'NOCTANE': 'OCTANE', 'NPBENZENE': 'PRBENZEN', 'NPENTANE': 'PENTANE',
    'NPMERCAPTAN': 'PR1THIOL', 'NPROPYLBENZENE': 'PRBENZEN', 'O2': 'O2', 'OCTANE': 'OCTANE',
    'OXYGEN': 'O2', 'OXYLENE': 'OXYLENE', 'PENTANE': 'PENTANE', 'PPLN': 'PPLN',
    'PR1THIOL': 'PR1THIOL', 'PRBENZEN': 'PRBENZEN', 'PROPADIENE': 'PROPDIEN', 'PROPANE': 'PROPANE',
    'PROPDIEN': 'PROPDIEN', 'PROPENE': 'PROPENE', 'PROPYLENE': 'PROPENE', 'PROPYNE': 'PROPYNE',
    'PXYLENE': 'PXYLENE', 'SO2': 'SO2', 'SO3': 'SO3', 'STYRENE': 'STYRENE', 'T2BUTENE': 'T2BUTENE',
    'T2OCTENE': 'T2OCTENE', 'TBMERCAPTA': 'T-B-MERCAPTA', 'TOLUENE': 'TOLUENE',
    'TR2BUTENE': 'T2BUTENE', 'TR2HEXENE': '3MT2PNTE', 'TR2OCTENE': 'CTCP', 'TR2PENTENE': 'C2PENTEN',
    'TR3HEXENE': '3MT2PNTE', 'TR3OCTENE': 'CT4C', 'TR4OCTENE': 'T2OCTENE', 'VAC': 'VAC',
    'VINYLBENZENE': 'STYRENE', 'VNYLACETLEN': 'VAC', 'WATER': 'H2O',
}


# ════════════════════════════════════════════════════════════════════════
#  4.  COMPONENT CONSTANTS  (pure from the databank, pseudo from the dump)
# ════════════════════════════════════════════════════════════════════════
WATER_DENSITY_60F = 62.428   # lb/ft³ at 60°F

# ── PRO/II "SIMSCI/TWU" acentric factor (generalized Frost-Kalkwarf-Thodos
# vapor pressure correlation) ───────────────────────────────────────────
# AVEVA PRO/II Simulation Reference Manual, "SIMSCI/TWU Characterization
# Method", Eqns (1-14)-(1-18) and Table 1-4:
#   ln(Pr) = (A1 + Ω·A4) + (A2 + Ω·A5)/Tr + (A3 + Ω·A6)·ln(Tr) + A7·Pr/Tr²
# Ω is back-solved from this equation at the (known) NBP boundary
# condition, Tr,b = Tb/Tc, Pr,b = 1 atm / Pc. Ω is then substituted back
# into the same equation to get Pr at Tr = 0.7 (solved implicitly, since
# Pr appears on both sides via the A7 term), and the Pitzer definition
# (Eqn 1-18) gives ω = -log10(Pr at Tr=0.7) - 1.
# Verified against accepted ω: n-heptane 0.349 (lit. 0.349), n-octane
# 0.394 (lit. ~0.398), benzene 0.217 (lit. 0.212).
_FK_A = (10.2005, -10.6317, -5.58058, 2.09167, -2.09167, -1.70214, 0.4312)

def _fk_f0(tr):
    return _FK_A[0] + _FK_A[1]/tr + _FK_A[2]*math.log(tr)

def _fk_f1(tr):
    return _FK_A[3] + _FK_A[4]/tr + _FK_A[5]*math.log(tr)

def estimate_acentric(tb_r, tc_r, pc_psia):
    """PRO/II SIMSCI/TWU acentric factor (Frost-Kalkwarf-Thodos vapor
    pressure correlation, Eqns 1-14-1-18)."""
    if tc_r <= tb_r or pc_psia <= 0: return None
    a7 = _FK_A[6]
    tr_b = tb_r / tc_r
    pr_b = 14.696 / pc_psia
    f0_b, f1_b = _fk_f0(tr_b), _fk_f1(tr_b)
    if f1_b == 0: return None
    omega_cap = (math.log(pr_b) - a7*pr_b/tr_b**2 - f0_b) / f1_b

    tr = 0.7
    rhs = _fk_f0(tr) + omega_cap*_fk_f1(tr)
    pr = math.exp(rhs)
    for _ in range(50):
        g  = math.log(pr) - a7*pr/tr**2 - rhs
        dg = 1.0/pr - a7/tr**2
        step = g/dg
        pr -= step
        if abs(step) < 1e-12: break
    if pr <= 0: return None
    return -math.log10(pr) - 1.0

# ── Twu (1984) correlation ─────────────────────────────────────────────
# Twu, C.H., "An internally consistent correlation for predicting the
# critical properties and molecular weights of petroleum and coal-tar
# liquids," Fluid Phase Equilibria, 16 (1984) 137-150.
# Sole correlation for all pseudo-component (petroleum-fraction) Tc/Vc/Pc/MW.

def _twu_tc0(tb_r):
    """Eqn (1): hypothetical n-alkane critical temperature (°R)."""
    return tb_r / (0.533272 + 0.191017e-3*tb_r + 0.779681e-7*tb_r**2
                   - 0.284376e-10*tb_r**3 + 0.959468e28/tb_r**13)

def _twu_vc0(alpha):
    """Eqn (2): hypothetical n-alkane critical volume (ft³/lbmol)."""
    inner = 0.419869 - 0.505839*alpha - 1.56436*alpha**3 - 9481.70*alpha**14
    return (1 - inner) ** -8

def _twu_sg0(alpha):
    """Eqn (3): hypothetical n-alkane specific gravity."""
    return 0.843593 - 0.128624*alpha - 3.36159*alpha**3 - 13749.5*alpha**12

def _twu_pc0(alpha):
    """Eqn (8): hypothetical n-alkane critical pressure (psia)."""
    bracket = (3.83354 + 1.19629*alpha**0.5 + 34.8888*alpha
               + 36.1952*alpha**2 + 104.193*alpha**4)
    return bracket ** 2

def _twu_mw0(tb_r):
    """Eqn (4)/(7): hypothetical n-alkane molecular weight, solved by
    bisection (eqn. 4 is explicit in Tb, not MW)."""
    def tb_from_mw(mw_):
        theta = math.log(mw_)
        return (math.exp(5.71419 + 2.71579*theta - 0.286590*theta**2
                          - 39.8544/theta - 0.122488/theta**2)
                - 24.7522*theta + 35.3155*theta**2)
    lo, hi = 2.0, 5000.0
    flo, fhi = tb_from_mw(lo) - tb_r, tb_from_mw(hi) - tb_r
    if flo * fhi > 0:
        return None
    for _ in range(100):
        mid = 0.5*(lo+hi)
        fm = tb_from_mw(mid) - tb_r
        if (fm > 0) == (flo > 0):
            lo, flo = mid, fm
        else:
            hi, fhi = mid, fm
    return 0.5*(lo+hi)

# Twu's re-summation g = g0*[(1+2f)/(1-2f)]^2 is singular as f -> 0.5
# (the ratio diverges). Real fractions far from the n-alkane reference
# (very aromatic/naphthenic, or outside the paper's fitted Watson K range
# of ~8-14) can push f past that point, producing nonphysical Tc/Vc/Pc/MW.
# Clamp f so (1+2f)/(1-2f), squared, stays within ~5.4x of g0 — generous
# enough for legitimate heavy-fraction deviations, tight enough to kill
# the runaway blow-up seen for out-of-range Watson K inputs.
TWU_F_CLAMP = 0.20

def _clamp_f(f):
    clamped = abs(f) > TWU_F_CLAMP
    return max(-TWU_F_CLAMP, min(TWU_F_CLAMP, f)), clamped

def estimate_twu_props(tb_r, sg):
    """
    Twu (1984) perturbation about the n-alkane reference system.
    Given Tb(°R) and SG(60°F), returns
    (Tc_R, Vc_ft3lbmol, Pc_psia, MW, was_clamped) or None if the
    reference solve fails. The perturbation terms f_T, f_V, f_P, f_M are
    clamped (see TWU_F_CLAMP) to avoid the singularity in the
    (1+2f)/(1-2f) re-summation for components far outside the
    correlation's fitted range; was_clamped is True if any of them hit
    the clamp, flagging the result as an extrapolation.
    """
    tc0 = _twu_tc0(tb_r)
    if tc0 <= 0:
        return None
    alpha = 1 - tb_r/tc0
    vc0 = _twu_vc0(alpha)
    sg0 = _twu_sg0(alpha)
    pc0 = _twu_pc0(alpha)
    mw0 = _twu_mw0(tb_r)
    if mw0 is None:
        return None

    # Eqns (11)-(13): critical temperature
    dsg_t = math.exp(5*(sg0 - sg)) - 1
    f_t, c_t = _clamp_f(dsg_t * (-0.362456/tb_r**0.5
                        + (0.0398285 - 0.948125/tb_r**0.5)*dsg_t))
    tc = tc0 * ((1 + 2*f_t)/(1 - 2*f_t))**2

    # Eqns (14)-(16): critical volume
    dsg_v = math.exp(4*(sg0**2 - sg**2)) - 1
    f_v, c_v = _clamp_f(dsg_v * (0.466590/tb_r**0.5
                        + (-0.182421 + 3.01721/tb_r**0.5)*dsg_v))
    vc = vc0 * ((1 + 2*f_v)/(1 - 2*f_v))**2

    # Eqns (17)-(19): critical pressure
    dsg_p = math.exp(0.5*(sg0 - sg)) - 1
    f_p, c_p = _clamp_f(dsg_p * ((2.53262 - 46.1955/tb_r**0.5 - 0.00127885*tb_r)
                        + (-11.4277 + 252.140/tb_r**0.5 + 0.00230535*tb_r)*dsg_p))
    pc = pc0 * (tc/tc0) * (vc0/vc) * ((1 + 2*f_p)/(1 - 2*f_p))**2

    # Eqns (20)-(23): molecular weight
    dsg_m = math.exp(5*(sg0 - sg)) - 1
    abs_x = abs(0.0123420 - 0.328086/tb_r**0.5)
    f_m, c_m = _clamp_f(dsg_m * (abs_x + (-0.0175691 + 0.193168/tb_r**0.5)*dsg_m))
    ln_mw = math.log(mw0) * ((1 + 2*f_m)/(1 - 2*f_m))**2
    mw = math.exp(ln_mw)

    return tc, vc, pc, mw, (c_t or c_v or c_p or c_m)

def estimate_pseudo_props(mw, nbp_f, sld_lbft3):
    """
    Given MW, NBP(°F), SLD(lb/ft³) for a pseudo-component,
    estimate Tc(°F), Pc(psia), Vc(ft³/lbmol), Zc, ω via Twu (1984).
    Returns dict (empty if the reference solve fails).
    """
    if None in (mw, nbp_f, sld_lbft3) or sld_lbft3 <= 0:
        return {}
    sg  = sld_lbft3 / WATER_DENSITY_60F
    tb_r = nbp_f + 459.67   # °F → °R
    if tb_r <= 0 or sg <= 0:
        return {}

    twu = estimate_twu_props(tb_r, sg)
    if twu is None:
        return {}
    tc_r, vc, pc_psia, mw_twu, was_clamped = twu
    tc_f  = tc_r - 459.67
    zc    = pc_psia * vc / (10.7316 * tc_r) if pc_psia and vc else 0.27
    omega = estimate_acentric(tb_r, tc_r, pc_psia)

    return {
        'Tc_F':     round(tc_f, 4),
        'Tc_R':     round(tc_r, 4),
        'Pc_psia':  round(pc_psia, 4),
        'Vc_ft3lbmol': round(vc, 4) if vc else None,
        'Zc':       round(zc, 4),
        'omega':    round(omega, 6) if omega is not None else None,
        'SG':       round(sg, 6),
        'extrapolated': was_clamped,
        'method':   'Twu (1984)',
    }


def soreide_tb_f(mw: float, sg: float) -> float | None:
    """Søreide (1989) normal boiling point (°F) from MW and SG — used only when the cut
    name carries no NBP."""
    try:
        tb_r = 1928.3 - 1.695e5 * mw ** -0.03522 * sg ** 3.266 * math.exp(
            -4.922e-3 * mw - 4.7685 * sg + 3.462e-3 * mw * sg)
        return tb_r - 459.67
    except (ValueError, OverflowError, ZeroDivisionError):
        return None


_PF_NAME_RE = re.compile(r"^PF(\d+(?:\.\d+)?)A(\d+(?:\.\d+)?)D", re.I)


def decode_pf_name(name: str):
    """'PF282A58D_6' → (NBP °F, API) — PRO/II petroleum-cut naming; None if not that pattern."""
    m = _PF_NAME_RE.match(str(name).strip())
    return (float(m.group(1)), float(m.group(2))) if m else None


_NORM_RE = re.compile(r"[^A-Z0-9=]")


def pure_key(name: str):
    """Resolve a component name to its PURE_DB key (None → not a databank pure)."""
    u = str(name).strip().upper()
    if u in PURE_DB:
        return u
    n = _NORM_RE.sub("", u)
    k = PURE_ALIAS.get(n)
    if k in PURE_DB:
        return k
    return n if n in PURE_DB else None


def read_constants_xlsx(path):
    """Optional external Constants.xlsx (name, CAS, MW, SLD, NBP, Tc, Pc, Vc, Zc, ω, PR, SRK,
    parachor …, header rows 1-2) → {NAME_UPPER: dict}.  'Missing' cells → None."""
    wb = load_workbook(path, read_only=True, data_only=True)
    out = {}
    try:
        for r in list(wb.worksheets[0].iter_rows(min_row=3, values_only=True)):
            if not r or r[0] is None or not str(r[0]).strip():
                continue

            def g(i, r=r):
                return sf(r[i]) if i < len(r) else None
            out[str(r[0]).strip().upper()] = dict(
                name=str(r[0]).strip(), cas=(str(r[1]).strip() if len(r) > 1 and r[1] else None),
                mw=g(2), sld=g(3), nbp=g(4), tc=g(5), pc=g(6), vc=g(7), zc=g(8), w=g(9),
                pr=g(10), srk=g(11), par=g(12))
    finally:
        wb.close()
    return out


def build_constants(comp_names, comp_stats, constants_path=None, log=print):
    """One COMP_CONSTANTS entry per dump component.

    pure    : embedded databank (MW/SLD cross-checked against the dump)
    pseudo  : MW + SLD measured from the dump; NBP from the cut name (or an optional
              Constants.xlsx, or Søreide); Tc/Pc/Vc/ω by Twu + SIMSCI ω.
    """
    ext = read_constants_xlsx(constants_path) if constants_path else {}
    out, warn = [], []
    for nm in comp_names:
        st = comp_stats.get(nm) or {}
        mw_d, sld_d = st.get("mw"), st.get("sld")
        e = dict(name=nm, CAS='', MW=None, SLD_lbft3=None, NBP_F=None, Tc_F=None, Pc_psia=None,
                 Vc_ft3lbmol=None, Zc=None, omega=None, PR_Peneloux=None, SRK_Peneloux=None,
                 Parachor=None, SG=None, is_pseudo=False, estimated=False, method='')
        key = pure_key(nm)
        if key is not None:
            cas, mw, nbp, tc, pc, vc, zc, w, pr, srk, par, sld = PURE_DB[key]
            e.update(CAS=cas or '', MW=mw_d or mw, NBP_F=nbp, Tc_F=tc, Pc_psia=pc, Vc_ft3lbmol=vc,
                     Zc=zc, omega=w, PR_Peneloux=pr, SRK_Peneloux=srk, Parachor=par,
                     SLD_lbft3=sld_d or sld, method=f'Library ({key})')
            if mw_d and mw and abs(mw_d - mw) / mw > 0.02:
                warn.append(f"{nm}: dump MW {mw_d:.3f} vs databank {key} MW {mw:.3f}")
            out.append(e)
            continue
        # ── pseudo-component ────────────────────────────────────────────
        e['is_pseudo'] = True
        x = ext.get(nm.upper(), {})
        mw = mw_d or x.get("mw")
        sld = sld_d or x.get("sld")
        pf = decode_pf_name(nm)
        nbp, nbp_src = x.get("nbp"), "Constants.xlsx"
        if nbp is None and pf:
            nbp, nbp_src = pf[0], "cut name"
        if sld is None and pf:                       # component absent from every stream
            sld = 62.428 * 141.5 / (131.5 + pf[1])
        if mw is None or sld is None or sld <= 0:
            e['method'] = '⚠ no data — component carries no flow in the dump'
            out.append(e)
            continue
        sg = sld / 62.428
        if nbp is None:
            nbp, nbp_src = soreide_tb_f(mw, sg), "Søreide"
        if pf:
            sg_name = 141.5 / (131.5 + pf[1])
            if abs(sg_name - sg) / sg > 0.03:
                warn.append(f"{nm}: dump SG {sg:.4f} vs cut-name API SG {sg_name:.4f}")
        if x.get("mw") and mw_d and abs(x["mw"] - mw_d) / mw_d > 0.02:
            warn.append(f"{nm}: Constants.xlsx MW {x['mw']:.2f} ignored — dump says {mw_d:.2f}")
        est = estimate_pseudo_props(mw, nbp, sld) if nbp is not None else {}
        e.update(MW=mw, SLD_lbft3=sld, NBP_F=nbp, SG=sg)
        meth = f"Twu (1984) + SIMSCI ω; MW, SLD from dump; NBP from {nbp_src}"
        if est:
            e.update(Tc_F=est.get('Tc_F'), Pc_psia=est.get('Pc_psia'), Vc_ft3lbmol=est.get('Vc_ft3lbmol'),
                     Zc=est.get('Zc'), omega=est.get('omega'), estimated=True)
            if est.get('Pc_psia') is not None and est['Pc_psia'] < 10:
                meth += " (⚠ unreliable — Pc < 10 psia)"
            elif est.get('extrapolated'):
                meth += " (⚠ extrapolated — Twu perturbation clamped)"
            elif est.get('omega') is None:
                meth += " (⚠ ω unavailable)"
        else:
            meth += " (⚠ Twu failed — no Tc/Pc/ω)"
        e['method'] = meth
        out.append(e)
    for w_ in warn:
        log(f"  ! {w_}")
    return out


COMP_HEADERS = [
    ('Name', 14, 'left'), ('Type', 8, 'center'), ('CAS', 14, 'center'),
    ('MW\n(g/mol)', 10, 'right'), ('NBP\n(°F)', 10, 'right'), ('SLD\n(lb/ft³)', 10, 'right'),
    ('SG\n(60°F)', 8, 'right'), ('Tc\n(°F)', 10, 'right'), ('Pc\n(psia)', 10, 'right'),
    ('Vc\n(ft³/lbmol)', 12, 'right'), ('Zc', 8, 'right'), ('ω\n(acentric)', 10, 'right'),
    ('PR Peneloux', 12, 'right'), ('SRK Peneloux', 12, 'right'), ('Parachor', 10, 'right'),
    ('Method', 28, 'left'),
]


def build_comp_constants_sheet(wb, comp_list):
    """COMP_CONSTANTS sheet — layout read by hydraulics.py (B name … E MW, F NBP, G SLD, I Tc,
    J Pc, K Vc, M ω)."""
    if 'COMP_CONSTANTS' in wb.sheetnames:
        del wb['COMP_CONSTANTS']
    ws = wb.create_sheet('COMP_CONSTANTS')
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = C['PURPLE']
    NC = len(COMP_HEADERS)
    LC = get_column_letter(NC + 1)
    cw(ws, 'A', 3)
    ws.merge_cells(f"B1:{LC}1")
    _c(ws, 1, 2, f"COMPONENT CONSTANTS  —  {len(comp_list)} components  |  pure: embedded databank  |  "
                 f"pseudo: MW & SLD measured from the dump, Twu (1984) + SIMSCI/TWU ω",
       bg=C['NAVY'], fg=C['WHITE'], sz=11, bold=True)
    rh(ws, 1, 22)
    for ci, (hdr, w, ha) in enumerate(COMP_HEADERS, 2):
        _c(ws, 2, ci, hdr, bg=C['NAVY'], fg=C['WHITE'], sz=8, bold=True, ha='center', wrap=True)
        cw(ws, ci, w)
    rh(ws, 2, 34)
    ws.auto_filter.ref = f"B2:{LC}2"
    ws.freeze_panes = "C3"
    n_pure = n_pseudo = n_est = 0
    for ri, comp in enumerate(comp_list):
        r = ri + 3
        is_p, is_e = comp.get('is_pseudo', False), comp.get('estimated', False)
        n_pseudo += is_p
        n_pure += (not is_p)
        n_est += is_e
        bg_name = C['PSEUDO'] if is_p else C['PURE']
        bg_data = C['ESTIM'] if is_e else (C['LGRAY'] if ri % 2 else C['WHITE'])
        vals = [comp['name'], 'PSEUDO' if is_p else 'PURE', comp.get('CAS', ''),
                fmtv(comp.get('MW'), 4), fmtv(comp.get('NBP_F'), 2), fmtv(comp.get('SLD_lbft3'), 3),
                fmtv(comp.get('SG'), 4), fmtv(comp.get('Tc_F'), 2), fmtv(comp.get('Pc_psia'), 2),
                fmtv(comp.get('Vc_ft3lbmol'), 4), fmtv(comp.get('Zc'), 4), fmtv(comp.get('omega'), 6),
                fmtv(comp.get('PR_Peneloux'), 6), fmtv(comp.get('SRK_Peneloux'), 6),
                fmtv(comp.get('Parachor'), 3), comp.get('method', '')]
        for ci, v in enumerate(vals, 2):
            _c(ws, r, ci, v, bg=(bg_name if ci == 2 else bg_data), sz=8, ha=COMP_HEADERS[ci - 2][2],
               bold=(ci == 2), italic=(is_e and 9 <= ci <= 13))
        rh(ws, r, 13)
    r = len(comp_list) + 4
    ws.merge_cells(f"B{r}:{LC}{r}")
    _c(ws, r, 2, f"  Total: {len(comp_list)} components  |  Pure: {n_pure} (databank)  |  "
                 f"Pseudo: {n_pseudo} ({n_est} estimated by Twu + SIMSCI ω)  |  estimated cells in yellow italic",
       bg=C['LGRAY'], fg=C['DGRAY'], sz=8, italic=True)
    rh(ws, r, 15)
    r += 2
    for lbl, bg, desc in [('PURE', C['PURE'], 'Embedded databank component — Tc/Pc/ω/Peneloux from the library'),
                          ('PSEUDO', C['PSEUDO'], 'Petroleum cut — MW and SLD measured from the dump'),
                          ('ESTIMATED', C['ESTIM'], 'Twu (1984) Tc/Pc/Vc + PRO/II SIMSCI/TWU ω — used by the flash')]:
        _c(ws, r, 2, lbl, bg=bg, sz=8, bold=True, ha='center')
        _c(ws, r, 3, desc, bg=C['WHITE'], sz=8, italic=True)
        ws.merge_cells(f"C{r}:{LC}{r}")
        rh(ws, r, 14)
        r += 1
    print(f"  COMP_CONSTANTS: {len(comp_list)} components ({n_pure} pure, {n_pseudo} pseudo, {n_est} estimated)")
    return ws


# ════════════════════════════════════════════════════════════════════════
#  5.  HMB WORKBOOK SHEETS  (UNITS · COMPONENTS · INPUT · OUTPUT · one sheet per stream)
# ════════════════════════════════════════════════════════════════════════
# ── UNIT sheet ────────────────────────────────────────────────────────
def build_units(wb):
    ws = wb.create_sheet("UNITS")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = C['TEAL']
    cw(ws,'A',3); cw(ws,'B',36); cw(ws,'C',22); cw(ws,'D',3)

    ws.merge_cells("B1:C1")
    cell(ws,1,2,"UNITS OF MEASURE  —  TD Property Dump",
         bg=C['NAVY'], fg=C['WHITE'], sz=12, bold=True, ha='center')
    rh(ws,1,22)

    units_data = [
        ("Temperature",      "°F"),
        ("Pressure",         "PSIA"),
        ("Molar Flow Rate",  "LB-MOL/HR"),
        ("Mass Flow Rate",   "LB/HR"),
        ("Std Liq Vol Rate", "FT3/HR  (at 1 atm, 60°F)"),
        ("Std Vap Vol Rate", "FT3/HR  (at 14.696 psia, 60°F)"),
        ("Act Vol Rate",     "FT3/HR"),
        ("Enthalpy",         "MM BTU/HR"),
        ("Sp Enthalpy",      "BTU/LB"),
        ("Sp Entropy",       "BTU/LB-F"),
        ("Density (Actual)", "LB/FT3"),
        ("Viscosity",        "cP"),
        ("Thermal Cond",     "BTU/HR-FT-F"),
        ("Surface Tension",  "DYNE/CM"),
        ("Cp",               "BTU/LB-F"),
        ("Crit Temperature", "°F  (True Critical)"),
        ("Crit Pressure",    "PSIA  (True Critical)"),
        ("Crit Volume",      "FT3/LB-MOL  (True Critical)"),
        ("Acentric Factor",  "—  (dimensionless)"),
        ("Mol Weight",       "g/mol  (dimensionless)"),
    ]
    cell(ws,2,2,"Property",  bg=C['NAVY'], fg=C['WHITE'], sz=9, bold=True)
    cell(ws,2,3,"Unit",      bg=C['NAVY'], fg=C['WHITE'], sz=9, bold=True)
    rh(ws,2,18)
    for ri, (prop, unit) in enumerate(units_data):
        r  = ri + 3
        bg = C['WHITE'] if ri%2==0 else C['LGRAY']
        cell(ws,r,2,prop, bg=bg, sz=9)
        cell(ws,r,3,unit, bg=bg, sz=9, bold=True)
        rh(ws,r,14)

# ── COMPONENTS sheet ──────────────────────────────────────────────────
def build_components(wb, comp_data, stream_names):
    ws = wb.create_sheet("COMPONENTS")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = C['ORANGE']
    cw(ws,'A',3); cw(ws,'B',8); cw(ws,'C',28); cw(ws,'D',3)

    ws.merge_cells("B1:C1")
    cell(ws,1,2,"COMPONENT INDEX", bg=C['NAVY'], fg=C['WHITE'],
         sz=12, bold=True, ha='center')
    rh(ws,1,22)
    cell(ws,2,2,"#",    bg=C['NAVY'], fg=C['WHITE'], sz=9, bold=True, ha='center')
    cell(ws,2,3,"Name", bg=C['NAVY'], fg=C['WHITE'], sz=9, bold=True)
    rh(ws,2,18)

    # Collect all unique component names from MOLE_FRAC section
    all_comps = []
    seen = set()
    for sn in stream_names:
        for cname in comp_data.get(sn,{}).get('MOLE_FRAC',{}).keys():
            if cname not in seen:
                seen.add(cname)
                all_comps.append(cname)

    for ri, name in enumerate(all_comps):
        r  = ri + 3
        bg = C['WHITE'] if ri%2==0 else C['LGRAY']
        cell(ws,r,2,ri+1, bg=bg, sz=9, ha='center')
        cell(ws,r,3,name, bg=bg, sz=9)
        rh(ws,r,13)

# ── INPUT sheet — TD input conditions ────────────────────────────────
def build_input(wb, streams, stream_names):
    ws = wb.create_sheet("INPUT")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = C['TEAL']

    for col,w in [('A',3),('B',24),('C',12),('D',12),('E',12),
                  ('F',16),('G',16),('H',14),('I',16),('J',3)]:
        cw(ws,col,w)

    ws.merge_cells("B1:I1")
    cell(ws,1,2,"TD PROPERTY DUMP — STREAM INPUT CONDITIONS",
         bg=C['NAVY'], fg=C['WHITE'], sz=12, bold=True, ha='center')
    rh(ws,1,22)

    hdrs = ['Stream','Phase','Temp (°F)','Press (PSIA)',
            'Total Flow\n(LB-MOL/HR)','Total Mass\n(LB/HR)',
            'Liq Mole Frac','Vap Mole Frac']
    for ci, h in enumerate(hdrs, 2):
        cell(ws,2,ci,h, bg=C['NAVY'], fg=C['WHITE'], sz=9, bold=True,
             ha='center', wrap=True)
    rh(ws,2,30)
    ws.auto_filter.ref = "B2:I2"
    ws.freeze_panes    = "B3"

    for ri, sn in enumerate(stream_names):
        row = ri + 3
        bg  = C['LTBLUE'] if ri%2==0 else C['WHITE']
        p   = streams[sn]
        vals = [
            sn,
            p.get('phase_str',''),
            fmt(p.get('temperature')),
            fmt(p.get('pressure')),
            fmt(p.get('molar_rate')),
            fmt(p.get('mass_rate')),
            fmt(p.get('liq_mole_frac')),
            fmt(p.get('vap_mole_frac')),
        ]
        for ci, v in enumerate(vals, 2):
            cell(ws,row,ci,v, bg=bg, sz=9,
                 bold=(ci==2), ha='left' if ci<=3 else 'right')
        rh(ws,row,14)

# ── OUTPUT summary sheet (streams as columns) ─────────────────────────
def build_output(wb, streams, stream_names):
    ws = wb.create_sheet("OUTPUT")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = C['GRNHDR']

    N  = len(stream_names)
    LC = get_column_letter(3 + N)

    ws.merge_cells(f"B1:{LC}1")
    cell(ws,1,2,f"TD PROPERTY DUMP — STREAM SUMMARY  ({N} streams)",
         bg=C['NAVY'], fg=C['WHITE'], sz=12, bold=True, ha='center')
    rh(ws,1,22)
    cw(ws,'A',3); cw(ws,'B',32); cw(ws,'C',14)
    for j in range(N): cw(ws, 4+j, 15)

    cell(ws,2,2,"Property",  bg=C['NAVY'],fg=C['WHITE'],sz=9,bold=True,ha='center')
    cell(ws,2,3,"Unit",      bg=C['NAVY'],fg=C['WHITE'],sz=9,bold=True,ha='center')
    for j, sn in enumerate(stream_names):
        ci    = 4+j
        phase = streams[sn].get('phase_str','')
        bg    = C['PH_VAP'] if 'Vapor' in phase else \
                C['PH_WAT'] if 'Water' in phase else \
                C['PH_LIQ'] if 'Liquid' in phase else C['LGRAY']
        cell(ws,2,ci,sn, bg=bg,fg=C['DGRAY'],sz=8,bold=True,ha='center',wrap=True)
    rh(ws,2,36)

    ROWS = [
        # (key, label, unit, phase_hint)
        ('phase_str',        'Phase',                     '-',           C['LGRAY']),
        ('temperature',      'Temperature',               '°F',          C['PH_TOT']),
        ('pressure',         'Pressure',                  'PSIA',        C['PH_TOT']),
        ('molar_rate',       'Total Molar Rate',          'LB-MOL/HR',   C['PH_TOT']),
        ('mass_rate',        'Total Mass Rate',           'LB/HR',       C['PH_TOT']),
        ('std_liq_rate',     'Std Liq Vol Rate',          'FT3/HR',      C['PH_TOT']),
        ('std_vap_rate',     'Std Vap Vol Rate',          'FT3/HR',      C['PH_TOT']),
        ('act_vol_rate',     'Actual Vol Rate',           'FT3/HR',      C['PH_TOT']),
        ('mol_weight',       'Molecular Weight',          '-',           C['PH_TOT']),
        ('enthalpy',         'Enthalpy',                  'MM BTU/HR',   C['PH_TOT']),
        ('sp_enthalpy',      'Specific Enthalpy',         'BTU/LB',      C['PH_TOT']),
        ('liq_mole_frac',    'Liquid Mole Fraction',      '-',           C['PH_LIQ']),
        ('vap_mole_frac',    'Vapor Mole Fraction',       '-',           C['PH_VAP']),
        ('liq_wt_frac',      'Liquid Weight Fraction',    '-',           C['PH_LIQ']),
        ('water_mole_frac',  'Water Mole Fraction',       '-',           C['PH_WAT']),
        ('act_density',      'Actual Density (total)',    'LB/FT3',      C['PH_TOT']),
        ('tc_f',             'True Critical Temp',        '°F',          C['AMBER']),
        ('pc_psia',          'True Critical Pressure',    'PSIA',        C['AMBER']),
        ('tc_vol',           'True Critical Volume',      'FT3/LB-MOL',  C['AMBER']),
        ('acentric',         'Acentric Factor',           '-',           C['AMBER']),
        ('watson_k',         'Watson K (UOP)',             '-',           C['AMBER']),
        ('std_api',          'Std API Gravity',           '°API',        C['PH_LIQ']),
        ('z_total',          'Z Factor (total)',          '-',           C['PH_VAP']),
        ('vap_molar_rate',   'Vapor Molar Rate',          'LB-MOL/HR',   C['PH_VAP']),
        ('vap_mass_rate',    'Vapor Mass Rate',           'LB/HR',       C['PH_VAP']),
        ('vap_act_density',  'Vapor Actual Density',      'LB/FT3',      C['PH_VAP']),
        ('vap_viscosity',    'Vapor Viscosity',           'cP',          C['PH_VAP']),
        ('vap_cp',           'Vapor Cp',                  'BTU/LB-F',    C['PH_VAP']),
        ('vap_cp_cv',        'Vapor Cp/Cv',               '-',           C['PH_VAP']),
        ('vap_z_den',        'Vapor Z (from density)',    '-',           C['PH_VAP']),
        ('vap_therm_cond',   'Vapor Thermal Cond',        'BTU/HR-FT-F', C['PH_VAP']),
        ('liq_molar_rate',   'Liquid Molar Rate',         'LB-MOL/HR',   C['PH_LIQ']),
        ('liq_mass_rate',    'Liquid Mass Rate',          'LB/HR',       C['PH_LIQ']),
        ('liq_act_density',  'Liquid Actual Density',     'LB/FT3',      C['PH_LIQ']),
        ('liq_std_density',  'Liquid Std Density',        'LB/FT3',      C['PH_LIQ']),
        ('liq_viscosity',    'Liquid Viscosity',          'cP',          C['PH_LIQ']),
        ('liq_cp',           'Liquid Cp',                 'BTU/LB-F',    C['PH_LIQ']),
        ('liq_surf_tension', 'Liquid Surface Tension',    'DYNE/CM',     C['PH_LIQ']),
        ('liq_therm_cond',   'Liquid Thermal Cond',       'BTU/HR-FT-F', C['PH_LIQ']),
        ('liq_std_lv_rate',  'Liquid Std LV Rate',        'FT3/HR',      C['PH_LIQ']),
    ]

    r = 3
    for key, label, unit, ph_bg in ROWS:
        cell(ws,r,2,label, bg=ph_bg, fg=C['DGRAY'], sz=9)
        cell(ws,r,3,unit,  bg=ph_bg, fg=C['MGRAY'], sz=8, italic=True, ha='center')
        for j, sn in enumerate(stream_names):
            ci = 4+j
            v  = streams[sn].get(key)
            if key == 'phase_str':
                cell(ws,r,ci, v or '', bg=ph_bg, sz=8, ha='center')
            else:
                cell(ws,r,ci, fmt(v), bg=ph_bg, sz=8, ha='right')
        rh(ws,r,14); r += 1

    ws.freeze_panes = "D3"
    ws.auto_filter.ref = f"B2:{LC}2"

# ── Per-stream sheet (vertical — matches proii_parser.py structure) ───
def build_stream_sheet(wb, sn, props, comp_data_s, tab_color):
    safe = re.sub(r'[/\\*\[\]:?]','_', str(sn))[:31]
    existing = {s.title for s in wb.worksheets}
    base = safe; n = 1
    while safe in existing:
        safe = f"{base[:29]}_{n}"; n += 1

    ws = wb.create_sheet(safe)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = tab_color

    cw(ws,'A',3); cw(ws,'B',30); cw(ws,'C',14)
    cw(ws,'D',20); cw(ws,'E',18); cw(ws,'F',18); cw(ws,'G',3)

    # Phase columns: TOTAL | VAPOR | LIQUID
    PHASES = [('TOTAL',C['PH_TOT']),('VAPOR',C['PH_VAP']),('LIQUID',C['PH_LIQ'])]
    LC = 'F'

    # Title
    ws.merge_cells(f"B1:{LC}1")
    phase = props.get('phase_str','')
    cell(ws,1,2,
         f"STREAM: {sn}  |  Phase: {phase}  |  "
         f"T: {fmt(props.get('temperature'))} °F  |  "
         f"P: {fmt(props.get('pressure'))} PSIA",
         bg=C['NAVY'], fg=C['WHITE'], sz=10, bold=True)
    rh(ws,1,22)

    # Column headers
    cell(ws,2,2,"Property",bg=C['NAVY'],fg=C['WHITE'],sz=9,bold=True,ha='center')
    cell(ws,2,3,"Unit",    bg=C['NAVY'],fg=C['WHITE'],sz=9,bold=True,ha='center')
    for i,(ph,bg) in enumerate(PHASES):
        cell(ws,2,4+i,ph, bg=bg,fg=C['DGRAY'],sz=9,bold=True,ha='center')
    rh(ws,2,22)
    ws.freeze_panes = "B3"

    row = 3
    def shdr(label, bg):
        nonlocal row
        ws.merge_cells(f"B{row}:{LC}{row}")
        cell(ws,row,2,f"  {label}", bg=bg, fg=C['WHITE'], sz=9, bold=True)
        for ci in range(3,7): cell(ws,row,ci,'', bg=bg)
        rh(ws,row,16); row += 1

    def prow(label, unit, total=None, vapor=None, liquid=None):
        nonlocal row
        bg = C['WHITE'] if row%2==0 else C['LGRAY']
        cell(ws,row,2,label, bg=bg, sz=9)
        cell(ws,row,3,unit,  bg=bg, fg=C['MGRAY'], sz=8, italic=True, ha='center')
        for ci, (ph,ph_bg), val in [(4,PHASES[0],total),
                                     (5,PHASES[1],vapor),
                                     (6,PHASES[2],liquid)]:
            bg2 = ph_bg if val is not None else bg
            cell(ws,row,ci,fmt(val) if val is not None else '', bg=bg2, sz=9, ha='right')
        rh(ws,row,14); row += 1

    # 1. Flow Rates
    shdr("1.  FLOW RATES", C['NAVY'])
    prow('Molar Rate','LB-MOL/HR', props.get('molar_rate'),
         props.get('vap_molar_rate'), props.get('liq_molar_rate'))
    prow('Mass Rate','LB/HR', props.get('mass_rate'),
         props.get('vap_mass_rate'), props.get('liq_mass_rate'))
    prow('Std Liq Vol Rate','FT3/HR', props.get('std_liq_rate'),
         None, props.get('liq_std_lv_rate'))
    prow('Std Vap Vol Rate','FT3/HR', props.get('std_vap_rate'),
         props.get('vap_std_vol'), None)
    prow('Actual Vol Rate','FT3/HR', props.get('act_vol_rate'),
         props.get('vap_act_vol'), props.get('liq_act_rate'))

    # 2. Conditions
    shdr("2.  CONDITIONS", C['DGRAY'])
    prow('Temperature','°F', props.get('temperature'))
    prow('Pressure','PSIA', props.get('pressure'))
    prow('Molecular Weight','-', props.get('mol_weight'),
         props.get('vap_mw'), props.get('liq_mw'))
    prow('Enthalpy','MM BTU/HR', props.get('enthalpy'))
    prow('Specific Enthalpy','BTU/LB', props.get('sp_enthalpy'),
         props.get('vap_sp_enthalpy'), props.get('liq_sp_enthalpy'))
    prow('Liquid Mole Fraction','-', props.get('liq_mole_frac'))
    prow('Vapor Mole Fraction','-', props.get('vap_mole_frac'))
    prow('Liquid Weight Fraction','-', props.get('liq_wt_frac'))
    prow('Water Mole Fraction','-', props.get('water_mole_frac'))

    # 3. Vapor Phase
    shdr("3.  VAPOR PHASE PROPERTIES", C['STEEL'])
    prow('Actual Density','LB/FT3', None, props.get('vap_act_density'))
    prow('Viscosity','cP', None, props.get('vap_viscosity'))
    prow('Cp','BTU/LB-F', None, props.get('vap_cp'))
    prow('Cp/Cv Ratio','-', None, props.get('vap_cp_cv'))
    prow('Z Factor (from density)','-', props.get('z_total'), props.get('vap_z_den'))
    prow('Z Factor (from K)','-', None, props.get('vap_z_k'))
    prow('Thermal Conductivity','BTU/HR-FT-F', None, props.get('vap_therm_cond'))

    # 4. Liquid Phase
    shdr("4.  LIQUID PHASE PROPERTIES", C['GRNHDR'])
    prow('Actual Density','LB/FT3', None, None, props.get('liq_act_density'))
    prow('Std Density','LB/FT3', None, None, props.get('liq_std_density'))
    prow('API Gravity','°API', props.get('std_api'), None, props.get('liq_std_api'))
    prow('Viscosity','cP', None, None, props.get('liq_viscosity'))
    prow('Cp','BTU/LB-F', None, None, props.get('liq_cp'))
    prow('Surface Tension','DYNE/CM', None, None, props.get('liq_surf_tension'))
    prow('Thermal Conductivity','BTU/HR-FT-F', None, None, props.get('liq_therm_cond'))

    # 5. Critical & Thermodynamic  ← KEY FOR EOS (no CSV needed)
    shdr("5.  CRITICAL PROPERTIES  (used by EOS in hydraulics engine)", C['PURPLE'])
    prow('True Critical Temperature','°F', props.get('tc_f'))
    prow('True Critical Pressure','PSIA', props.get('pc_psia'))
    prow('True Critical Volume','FT3/LB-MOL', props.get('tc_vol'))
    prow("Critical Temp (Kay's Rule)",'°F', props.get('tc_kays'))
    prow("Critical Press (Kay's Rule)",'PSIA', props.get('pc_kays'))
    prow('Acentric Factor','-', props.get('acentric'))
    prow('Molecular Weight (for EOS)','g/mol', props.get('mol_weight'))
    prow('Watson K (UOP)','-', props.get('watson_k'))
    prow('Actual Density (bulk)','LB/FT3', props.get('act_density'))

    # 6–9. Composition
    COMP_DISPLAY = [
        ('MOLE_FLOW', "6.  MOLE FLOW RATES (LB-MOL/HR)",  "LB-MOL/HR", C['NAVY']),
        ('MOLE_FRAC', "7.  MOLE FRACTIONS (total)",        "mol/mol",    C['STEEL']),
        ('MASS_FLOW', "8.  MASS FLOW RATES (LB/HR)",       "LB/HR",      C['GRNHDR']),
        ('MASS_FRAC', "9.  MASS FRACTIONS (total)",        "wt/wt",      C['TEAL']),
        ('VAP_MOLE_FRAC',"10. VAPOR MOLE FRACTIONS",       "mol/mol",    C['STEEL']),
        ('LIQ_MOLE_FRAC',"11. LIQUID MOLE FRACTIONS",      "mol/mol",    C['GRNHDR']),
    ]
    for ctype, title, unit, hdr_bg in COMP_DISPLAY:
        cd = comp_data_s.get(ctype, {})
        active = {k:v for k,v in cd.items() if v is not None and abs(v) > 1e-30}
        if not active: continue

        shdr(title, hdr_bg)
        # sub-header
        cell(ws,row,2,"  Component", bg=C['LGRAY'], sz=8, bold=True)
        cell(ws,row,3,unit, bg=C['LGRAY'], sz=8, italic=True, ha='center')
        cell(ws,row,4,"Value", bg=C['LGRAY'], sz=8, bold=True, ha='center')
        ec_pct = 5; ec_ppm = 6
        cell(ws,row,ec_pct,"Wt%/Mol%\n(calc)", bg=C['FORMULA'],sz=8,bold=True,
             ha='center',wrap=True)
        cell(ws,row,ec_ppm,"ppm (calc)", bg=C['FORMULA'],sz=8,bold=True,ha='center')
        rh(ws,row,18); row += 1

        total_sum = sum(v for v in active.values() if v > 0)
        for pi, (cname, val) in enumerate(sorted(active.items(),
                                                   key=lambda x: -abs(x[1]))):
            bg = C['WHITE'] if pi%2==0 else C['LGRAY']
            cell(ws,row,2,f"  {cname}", bg=bg, sz=8)
            cell(ws,row,3,unit, bg=bg, sz=7, italic=True, ha='center')
            cell(ws,row,4,fmt(val), bg=C['PH_TOT'], sz=8, ha='right')
            pct_v = val/total_sum*100  if total_sum else None
            ppm_v = val/total_sum*1e6  if total_sum else None
            cell(ws,row,ec_pct,
                 fmt(round(pct_v,4)) if pct_v is not None else '',
                 bg=C['FORMULA'], sz=8, ha='right')
            cell(ws,row,ec_ppm,
                 fmt(round(ppm_v,2)) if ppm_v is not None else '',
                 bg=C['FORMULA'], sz=8, ha='right')
            rh(ws,row,13); row += 1

    ws.auto_filter.ref = f"B2:{LC}2"


# ════════════════════════════════════════════════════════════════════════
#  6.  READER API  (what hydraulics.py imports as HMBPROII)
# ════════════════════════════════════════════════════════════════════════
_TD2PK = {v: k for k, v in PK2TD.items()}


@dataclass
class HMBStream:
    """Flat property bag for one stream; missing values are None."""
    name: str
    case: str
    props: dict = field(default_factory=dict)

    def get(self, key, default=None):
        v = self.props.get(key)
        return default if v is None else v

    @property
    def phase(self):
        return self.props.get("phase")


_CASE_CACHE: dict = {}


def _read_case(path_str: str, case: str) -> dict:
    """{stream: HMBStream} for one case.  Accepts the raw PRO/II dump (any sheet name) or the
    clean wide table written by ``hmbbuilder.py wide`` (row 1 = Stream, Phase, KEYS…)."""
    key = (path_str, case)
    if key in _CASE_CACHE:
        return _CASE_CACHE[key]
    wb = load_workbook(path_str, read_only=True, data_only=True)
    try:
        if case not in wb.sheetnames:
            raise KeyError(f"Case '{case}' not in {wb.sheetnames}")
        first = next(wb[case].iter_rows(min_row=1, max_row=1, values_only=True), (None,))
        wide = first[0] is not None and str(first[0]).strip().lower() == "stream"
        if wide:
            rows = list(wb[case].iter_rows(values_only=True))
    finally:
        wb.close()
    out: dict = {}
    if wide:
        header = rows[0]
        col_key = {ci: ("phase" if str(h).strip().lower() == "phase" else str(h).strip())
                   for ci, h in enumerate(header) if ci > 0 and h is not None}
        for row in rows[1:]:
            if not row or row[0] is None or not str(row[0]).strip():
                continue
            props = {}
            for ci, k in col_key.items():
                raw = row[ci] if ci < len(row) else None
                props[k] = (str(raw).strip() if raw is not None else None) if k == "phase" else sf(raw)
            nm, u, n = str(row[0]).strip(), str(row[0]).strip(), 1
            while u in out:
                u = f"{nm}_{n}"
                n += 1
            out[u] = HMBStream(name=nm, case=case, props=props)
    else:
        d = parse_dump(path_str, sheet=case, verbose=False)
        for sn_, tp in d["streams"].items():
            props = {"phase": tp.get("phase_str") or None}
            for tk, v in tp.items():
                if tk in _TD2PK:
                    props[_TD2PK[tk]] = v
            out[sn_] = HMBStream(name=sn_, case=case, props=props)
    _CASE_CACHE[key] = out
    return out


def load_streams(path, case: str = "Case 1") -> dict:
    return _read_case(str(path), case)


def list_cases(path) -> list:
    wb = load_workbook(str(path), read_only=True, data_only=True)
    try:
        skip = {"COMP_CONSTANTS", "UNITS", "COMPONENTS", "INPUT", "OUTPUT", "README", "LEGEND"}
        return [s for s in wb.sheetnames if s.strip().upper() not in skip]
    finally:
        wb.close()


def get_stream(name: str, path=None, case: str = "Case 1"):
    streams = load_streams(path, case)
    if name in streams:
        return streams[name]
    up = name.strip().upper()
    for k, v in streams.items():
        if k.upper() == up or v.name.upper() == up:
            return v
    return None


def build_wide_workbook(dump_path, out_path=None, sheet=None, constants_path=None):
    """Clean wide-table workbook: sheet per case, row 1 = Stream, Phase, KEYS…, one row per
    stream, plus COMP_CONSTANTS.  (The per-stream HMB is the default; this is the compact form.)"""
    d = parse_dump(dump_path, sheet=sheet)
    keys = [k for k in KEYS if k != "phase"]
    wb = Workbook()
    ws = wb.active
    ws.title = (d["sheet"] or "Case 1")[:31]
    ws.append(["Stream", "Phase", *keys])
    for sn_ in d["stream_names"]:
        tp = d["streams"][sn_]
        ws.append([sn_, tp.get("phase_str")] + [tp.get(PK2TD.get(k)) for k in keys])
    ws.freeze_panes = "A2"
    comp = build_constants(d["comp_names"], d["comp_stats"], constants_path)
    build_comp_constants_sheet(wb, comp)
    out = out_path or os.path.join(_out_folder(dump_path),
                                   f"HMBwide_{Path(dump_path).stem}_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    wb.save(out)
    return out


# ════════════════════════════════════════════════════════════════════════
#  7.  BUILD THE HMB WORKBOOK
# ════════════════════════════════════════════════════════════════════════
def build_hmb(dump_path, out_path=None, sheet=None, constants_path=None):
    """PRO/II dump → HMB workbook (UNITS · COMPONENTS · COMP_CONSTANTS · INPUT · OUTPUT · streams)."""
    print(f"\n{'=' * 64}\n  PRO/II dump: {os.path.basename(dump_path)}")
    d = parse_dump(dump_path, sheet=sheet)
    streams, comp_data, stream_names = d["streams"], d["comp_data"], d["stream_names"]
    print(f"  Components per stream: {len(d['comp_names'])}   labels matched: {d['n_label_keys']}/{len(KEYS)}")
    missing = [k for k in ("temp_f", "pres_psia", "total_mass", "vap_mass", "liq_mass", "vap_density",
                           "vap_visc", "vap_cp", "vap_cp_cv", "vap_therm_cond", "liq_density",
                           "liq_visc", "liq_cp", "liq_surf_tens") if k not in d["label_rows"]]
    if missing:
        print(f"  ! dump has no row labelled for: {', '.join(missing)}  (those properties stay blank)")

    wb = Workbook()
    del wb['Sheet']
    build_units(wb)
    build_components(wb, comp_data, stream_names)
    comp = build_constants(d["comp_names"], d["comp_stats"], constants_path)
    build_comp_constants_sheet(wb, comp)
    build_input(wb, streams, stream_names)
    build_output(wb, streams, stream_names)
    print(f"  Writing {len(stream_names)} stream sheets...", flush=True)
    for i, sn_ in enumerate(stream_names):
        build_stream_sheet(wb, sn_, streams[sn_], comp_data[sn_], TABS[i % len(TABS)])
        if (i + 1) % 100 == 0:
            print(f"    {i + 1}/{len(stream_names)}", flush=True)
    fixed = ["UNITS", "COMPONENTS", "COMP_CONSTANTS", "INPUT", "OUTPUT"]
    wb._sheets = [wb[n] for n in fixed] + [w for w in wb._sheets if w.title not in fixed]

    if out_path is None:
        safe = re.sub(r'[<>:"/\\|?*\s]', '_', Path(dump_path).stem)[:55]
        out_path = os.path.join(_out_folder(dump_path),
                                f"HMB_{safe}_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    wb.save(out_path)
    print(f"\n  Saved : {out_path}  ({os.path.getsize(out_path) / 1e6:.1f} MB)")
    print(f"  Sheets: {len(wb.worksheets)}  (5 fixed + {len(stream_names)} stream sheets)")
    print('=' * 64)
    return out_path


# ════════════════════════════════════════════════════════════════════════
#  8.  MULTI-CASE COMPARE  (former hmb_dump_parser.py v2)
# ════════════════════════════════════════════════════════════════════════
# Section colour map — used for row backgrounds in all sheets
SEC_COLORS = {
    "FLOW RATES":       C['NAVY'],
    "CONDITIONS":       C['DGRAY'],
    "VAPOR PROPERTIES": C['STEEL'],
    "LIQUID PROPERTIES":C['GRNHDR'],
    "WATER / AQUEOUS":  C['TEAL'],
    "THERMODYNAMIC":    C['PURPLE'],
    "REFINERY":         C['ORANGE'],
    "OTHER PROPERTIES": "505050",
}
SEC_ROW_BG = {
    "FLOW RATES":       C['PH_TOT'],
    "CONDITIONS":       C['PH_TOT'],
    "VAPOR PROPERTIES": C['PH_VAP'],
    "LIQUID PROPERTIES":C['PH_LIQ'],
    "WATER / AQUEOUS":  C['PH_WAT'],
    "THERMODYNAMIC":    C['AMBER'],
    "REFINERY":         "FFF0E8",
    "OTHER PROPERTIES": C['LGRAY'],
}
SEC_ORDER = ["FLOW RATES","CONDITIONS","VAPOR PROPERTIES","LIQUID PROPERTIES",
             "WATER / AQUEOUS","THERMODYNAMIC","REFINERY","OTHER PROPERTIES"]

def _classify(lbl):
    ll = lbl.lower()
    if any(x in ll for x in ['flow','rate','vol rate','std liq','std vap','actual vol']):
        return "FLOW RATES"
    if any(x in ll for x in ['temperature','pressure','molecular weight','mol weight',
                               'enthalpy','entropy','fraction','phase']):
        return "CONDITIONS"
    if 'vapor' in ll or 'vapour' in ll:
        return "VAPOR PROPERTIES"
    if 'liquid' in ll:
        return "LIQUID PROPERTIES"
    if 'water' in ll or 'aqueous' in ll:
        return "WATER / AQUEOUS"
    if any(x in ll for x in ['critical','acentric','watson','density','api','gravity',
                               'z-factor','z factor','reduced','flash','bubble','dew',
                               'bulk','spgr','kays','uopk']):
        return "THERMODYNAMIC"
    if any(x in ll for x in ['pour','cloud','flash point']):
        return "REFINERY"
    return "OTHER PROPERTIES"

# ── Skip sheets (not stream data) ─────────────────────────────────────
CMP_SKIP_SHEETS = {
    'UNITS','COMPONENTS','INPUT','OUTPUT','Summary','Component Index',
    'Legend','Stream_List','Line_Stream_Map','Stream_Properties',
    'Backend_Properties','Component_DB','Backend_Config',
    'FIV_LOF','Formula_Reference','README','VBA Code','CASES',
    'SUMMARY','STREAM_LOOKUP',
}

# ── Parser ─────────────────────────────────────────────────────────────
def parse_stream_sheet(ws):
    """
    Parse one per-stream sheet from proii_parser.py output.
    Returns:
      stream, phase, title_t, title_p, warned,
      props  : OrderedDict  {label: {phase_col: value}}
      comp   : OrderedDict  {section_label: {comp_name: {phase_col: value}}}
      prop_order : list of labels in the original row order (for section grouping)
    """
    title_cell = ws['B1'].value or ''
    sm = re.search(r'STREAM:\s*([^|]+)',  str(title_cell))
    pm = re.search(r'Phase:\s*([^|]+)',   str(title_cell))
    tm = re.search(r'T:\s*([\d\.\-]+)',   str(title_cell))
    prm= re.search(r'P:\s*([\d\.\-]+)',   str(title_cell))

    sname  = sm.group(1).strip()  if sm  else ws.title
    phase  = pm.group(1).strip()  if pm  else ''
    warned = '⚠' in str(title_cell) or 'UNCONVERGED' in str(title_cell).upper()

    # Column headers (row 2)
    hdr_row   = list(ws.iter_rows(min_row=2, max_row=2, values_only=True))[0]
    phase_cols = {}   # col_index (0-based) → phase name
    for ci, h in enumerate(hdr_row):
        if h and ci >= 2:
            phase_cols[ci] = str(h).strip()

    props      = OrderedDict()
    prop_order = []          # labels in sheet order (for section reconstruction)
    comp_block = None
    comp_data  = OrderedDict()

    for row_vals in ws.iter_rows(min_row=3, values_only=True):
        if not row_vals or row_vals[1] is None: continue
        label = str(row_vals[1]).strip()
        if not label: continue

        # ── Section header (numbered like "1.  FLOW RATES")
        is_section = bool(re.match(r'\s*\d+\.\s+', label))
        if is_section:
            comp_block = None
            upper = label.upper()
            if any(x in upper for x in ('COMPOSITION','MOLE FRAC','MOLE FLOW',
                                         'MASS FRAC','MASS FLOW','MASS PCT',
                                         'MOLE PCT','DRY LV','STD GAS')):
                comp_block = label.strip()
                comp_data[comp_block] = OrderedDict()
            continue

        # ── Composition data row  "  NNN  CompName"
        comp_m = re.match(r'\s{1,6}(\d{1,3})\s+(.+)$', label)
        if comp_block and comp_m:
            cname    = comp_m.group(2).strip()
            row_data = {}
            for ci, ph_name in phase_cols.items():
                if ci < len(row_vals):
                    v = sf(row_vals[ci])
                    if v is not None:
                        row_data[ph_name] = v
            if row_data:
                comp_data[comp_block][cname] = row_data
            continue

        # ── Normal property row
        if label in ('Unit','Property','Phase/Unit'): continue
        if label.startswith('#'): continue

        comp_block = None   # leaving any composition block

        row_data = {}
        for ci, ph_name in phase_cols.items():
            if ci < len(row_vals):
                v = row_vals[ci]
                if v is not None and str(v).strip() not in ('','N/A','MISSING'):
                    row_data[ph_name] = sf(v) if not isinstance(v, str) else str(v).strip()
        if row_data:
            props[label]    = row_data
            prop_order.append(label)

    return {
        'stream':     sname,
        'phase':      phase,
        'title_t':    sf(tm.group(1))  if tm  else None,
        'title_p':    sf(prm.group(1)) if prm else None,
        'warned':     warned,
        'props':      props,
        'prop_order': prop_order,
        'comp':       comp_data,
    }

def parse_hmb_workbook(path):
    wb         = load_workbook(path, read_only=True, data_only=True)
    base       = os.path.splitext(os.path.basename(path))[0]
    cm         = re.search(r'(Case\s*\d+|CASE\s*\d+|case\s*\d+)', base, re.I)
    case_label = cm.group(0).strip() if cm else base[:30]

    streams = []
    for ws in wb.worksheets:
        if ws.title in CMP_SKIP_SHEETS: continue
        try:
            data             = parse_stream_sheet(ws)
            data['case_label']  = case_label
            data['source_file'] = os.path.basename(path)
            streams.append(data)
        except Exception as e:
            print(f"    WARN sheet '{ws.title}' skipped — {e}")
    wb.close()
    return streams, case_label

# ── Shared helpers for sheet building ─────────────────────────────────
def _sheet_title(ws, text, last_col_ci, bg=C['NAVY']):
    LC = get_column_letter(last_col_ci)
    ws.merge_cells(f"B1:{LC}1")
    _c(ws,1,2,text, bg=bg,fg=C['WHITE'],sz=12,bold=True)
    rh(ws,1,22)

def _col_hdrs(ws, row, labels_and_bgs):
    """labels_and_bgs = [(text, bg), ...]  starting at col B"""
    for ci, (lbl, bg) in enumerate(labels_and_bgs, 2):
        _c(ws,row,ci,lbl, bg=bg,fg=C['WHITE'],sz=9,bold=True,ha='center',wrap=True)
    rh(ws,row,32)

def _sec_row(ws, r, label, last_ci, bg):
    ws.merge_cells(f"B{r}:{get_column_letter(last_ci)}{r}")
    _c(ws,r,2,f"  {label}", bg=bg,fg=C['WHITE'],sz=9,bold=True)
    for ci in range(3, last_ci+1):
        _c(ws,r,ci,'', bg=bg)
    rh(ws,r,15)

def _get_val(sdata, prop_label):
    """Get best value for a property from a stream data dict."""
    if not sdata: return None
    pdata = sdata['props'].get(prop_label)
    if pdata is None: return None
    return (pdata.get('TOTAL') or pdata.get('OVERALL')
            or pdata.get('Property') or next(iter(pdata.values()), None))

# ── Collect the master property list across ALL cases + ALL streams ────
def collect_all_props(all_cases):
    """
    Returns OrderedDict: {prop_label: section_name}
    in section-grouped order, preserving first-seen order within each section.
    """
    seen_in_sec = defaultdict(list)   # section → [labels in order]
    seen_global = set()

    for cl, sdict in all_cases.items():
        for sn, sdata in sdict.items():
            for lbl in sdata.get('prop_order', []):
                if lbl not in seen_global:
                    seen_global.add(lbl)
                    sec = _classify(lbl)
                    seen_in_sec[sec].append(lbl)

    result = OrderedDict()
    for sec in SEC_ORDER:
        for lbl in seen_in_sec.get(sec, []):
            result[lbl] = sec
    return result

# ── CASES sheet ────────────────────────────────────────────────────────
def build_cases_sheet(wb, all_cases):
    ws = wb.create_sheet("CASES")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = C['ORANGE']
    cw(ws,'A',3); cw(ws,'B',22); cw(ws,'C',10); cw(ws,'D',55); cw(ws,'E',18); cw(ws,'F',3)

    ws.merge_cells("B1:E1")
    _c(ws,1,2,"CASE REGISTRY", bg=C['NAVY'],fg=C['WHITE'],sz=12,bold=True,ha='center')
    rh(ws,1,22)

    for ci,h in enumerate(['Case Label','# Streams','Source File','Parsed'],2):
        _c(ws,2,ci,h, bg=C['NAVY'],fg=C['WHITE'],sz=9,bold=True,ha='center')
    rh(ws,2,18)

    for ri,(cl,sdict) in enumerate(all_cases.items()):
        r   = ri+3
        bg  = C['WHITE'] if ri%2==0 else C['LGRAY']
        src = next(iter(sdict.values()),{}).get('source_file','?') if sdict else '?'
        _c(ws,r,2,cl, bg=CASE_PALETTE[ri%len(CASE_PALETTE)],fg=C['WHITE'],sz=9,bold=True)
        _c(ws,r,3,len(sdict), bg=bg,sz=9,ha='center')
        _c(ws,r,4,src,  bg=bg,sz=9)
        _c(ws,r,5,datetime.now().strftime('%d-%b-%Y %H:%M'), bg=bg,sz=8,italic=True)
        rh(ws,r,15)

# ── SUMMARY sheet — ALL properties dynamically ────────────────────────
def build_summary_sheet(wb, all_cases, all_props):
    """
    Rows  = (stream_name, property_label)   — every property in the HMB
    Cols  = Case 1, Case 2, ...

    Section colour bands group properties visually.
    """
    ws = wb.active
    ws.title = "SUMMARY"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = C['NAVY']

    case_labels = list(all_cases.keys())
    N = len(case_labels)
    LC = get_column_letter(4+N)

    _sheet_title(ws,
        f"HMB MULTI-CASE SUMMARY  —  {N} case(s)  |  "
        f"{len(all_props)} properties  |  {datetime.now():%d-%b-%Y %H:%M}",
        4+N)

    cw(ws,'A',3); cw(ws,'B',28); cw(ws,'C',18)
    for j in range(N): cw(ws, 4+j, 18)

    _col_hdrs(ws, 2,
        [("Stream",   C['NAVY']),
         ("Property", C['NAVY'])] +
        [(cl, CASE_PALETTE[j % len(CASE_PALETTE)]) for j, cl in enumerate(case_labels)]
    )

    stream_order = list(dict.fromkeys(
        sn for cl in case_labels for sn in all_cases[cl].keys()
    ))

    row = 3
    prev_sec = None

    for sn in stream_order:
        first_in_stream = True

        for prop_lbl, sec in all_props.items():
            # Collect values for this property across all cases
            vals = {}
            for cl in case_labels:
                v = _get_val(all_cases[cl].get(sn), prop_lbl)
                if v is not None:
                    vals[cl] = v

            if not vals:
                continue    # property not present for this stream in any case

            # Section header band when section changes (first prop of stream resets it)
            if first_in_stream:
                # Dark stream name separator row
                ws.merge_cells(f"B{row}:{LC}{row}")
                _c(ws,row,2, f"  {sn}", bg=C['STEEL'],fg=C['WHITE'],sz=9,bold=True)
                for ci in range(3,4+N): _c(ws,row,ci,'',bg=C['STEEL'])
                rh(ws,row,15); row += 1
                first_in_stream = False
                prev_sec = None   # reset so section header prints fresh

            if sec != prev_sec:
                _sec_row(ws, row, sec, 3+N, SEC_COLORS[sec])
                row += 1
                prev_sec = sec

            # Property data row
            row_bg = SEC_ROW_BG.get(sec, C['LGRAY'])
            alt_bg = C['WHITE'] if row%2==0 else row_bg
            _c(ws,row,2,'', bg=alt_bg, sz=9, border=True)   # stream col blank
            _c(ws,row,3,prop_lbl, bg=alt_bg, sz=8, ha='left')

            for j, cl in enumerate(case_labels):
                v = vals.get(cl)
                if v is None:
                    _c(ws,row,4+j,'', bg=C['LGRAY'], sz=8, ha='center')
                elif isinstance(v, str):
                    _c(ws,row,4+j,v,  bg=alt_bg, sz=8, ha='center')
                else:
                    _c(ws,row,4+j, fmtv(v), bg=alt_bg, sz=9, ha='right')

            rh(ws,row,13); row += 1

    ws.freeze_panes = "C3"
    ws.auto_filter.ref = f"B2:{LC}2"

# ── STREAM_LOOKUP sheet — single-stream comparison via formula lookup ──
def build_stream_lookup(wb, all_cases, all_props, stream_names, preselect=None):
    """
    Interactive single-stream comparison sheet.

    Layout:
      B1: "TYPE STREAM NAME →"   C1: [editable input cell, yellow]
      Row 2: col headers (Property | Unit | Case1 | Case2 | ...)
      Row 3+: property rows with INDEX/XMATCH formulas pointing at S_* sheets

    The user types any stream name in C1.  All value cells update live via
    INDEX(XMATCH()) because the S_* sheets each have all properties in the
    same row order.
    """
    ws = wb.create_sheet("STREAM_LOOKUP")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = C['AMBER']

    case_labels = list(all_cases.keys())
    N = len(case_labels)
    LC = get_column_letter(3+N)

    cw(ws,'A',3); cw(ws,'B',30); cw(ws,'C',22)
    for j in range(N): cw(ws, 4+j, 22)

    # Row 1: instruction + input cell
    _c(ws,1,2,"▶  TYPE STREAM NAME IN CELL D1  ◀",
       bg=C['NAVY'],fg=C['WHITE'],sz=11,bold=True)
    _c(ws,1,3,"Stream name (editable):",
       bg=C['AMBER'],fg=C['DGRAY'],sz=9,bold=True)
    input_cell = ws.cell(row=1, column=4)
    input_cell.value       = preselect or (stream_names[0] if stream_names else '')
    input_cell.font        = Font(name="Calibri", size=11, bold=True, color=C['NAVY'])
    input_cell.fill        = _fill(C['WARN'])
    input_cell.alignment   = Alignment(horizontal='left', vertical='center')
    _sd = Side(style='medium', color=C['ORANGE'])
    input_cell.border      = Border(left=_sd, right=_sd, top=_sd, bottom=_sd)
    # Fill rest of row 1 with neutral colour
    for ci in range(5, 4+N):
        _c(ws,1,ci,'', bg=C['AMBER'])
    rh(ws,1,28)

    # Row 2: col headers
    _col_hdrs(ws, 2,
        [("Property", C['NAVY']),
         ("Section",  C['NAVY'])] +
        [(cl, CASE_PALETTE[j % len(CASE_PALETTE)]) for j,cl in enumerate(case_labels)]
    )

    # Build a reference map: S_<stream> sheet name for each case
    # We need to know the column in S_* sheets that corresponds to each case.
    # S_* sheets: row 3 = col headers, col B=Property, col C=Unit/Phase,
    #             cols D..D+N-1 = cases (same order as case_labels)

    # We'll write formulas that:
    #  1. Look up which S_ sheet exists for the typed stream name using STREAM_INDEX sheet
    #  2. Pull values via INDIRECT + XMATCH

    # Build STREAM_INDEX helper sheet (hidden): 2 cols — stream_name | sheet_name
    si_ws = wb.create_sheet("_STREAM_IDX")
    si_ws.sheet_state = 'hidden'
    si_ws['A1'] = 'Stream'
    si_ws['B1'] = 'SheetName'
    for ri, sn in enumerate(stream_names, 2):
        safe = re.sub(r'[/\\*\[\]:?]','_', sn)[:28]
        sheet_name = f"S_{safe}"
        si_ws.cell(ri,1).value = sn
        si_ws.cell(ri,2).value = sheet_name

    # Input cell address (absolute)
    INPUT = "$D$1"

    row = 3
    prev_sec = None

    # Collect all property row labels in section order (same as SUMMARY)
    for prop_lbl, sec in all_props.items():
        if sec != prev_sec:
            _sec_row(ws, row, sec, 3+N, SEC_COLORS[sec])
            row += 1; prev_sec = sec

        alt_bg = C['WHITE'] if row%2==0 else SEC_ROW_BG.get(sec, C['LGRAY'])
        _c(ws,row,2,prop_lbl, bg=alt_bg, sz=9, ha='left')
        _c(ws,row,3,sec,      bg=alt_bg, fg=C['MGRAY'], sz=7, italic=True, ha='left')

        # For each case column: pull value from S_<stream> sheet, row where Property=prop_lbl
        # Formula: =IFERROR(
        #     INDEX(
        #       INDIRECT("'"&IFERROR(INDEX(_STREAM_IDX!$B:$B,XMATCH($D$1,_STREAM_IDX!$A:$A)),"")&"'!$"&<case_col>&":$"&<case_col>),
        #       XMATCH(<prop_lbl>,
        #              INDIRECT("'"&IFERROR(INDEX(_STREAM_IDX!$B:$B,XMATCH($D$1,_STREAM_IDX!$A:$A)),"")&"'!$B:$B"))
        #     ), "")
        # The case column in S_* sheets is D for case 0, E for case 1, etc.

        sheet_ref = (
            f'IFERROR(INDEX(_STREAM_IDX!$B:$B,'
            f'XMATCH({INPUT},_STREAM_IDX!$A:$A)),"")'
        )

        for j in range(N):
            case_col = get_column_letter(4+j)  # D, E, F, ...
            f = (
                f'=IFERROR('
                f'INDEX('
                f'INDIRECT("\'"&{sheet_ref}&"\'!${case_col}:${case_col}"),'
                f'XMATCH("{prop_lbl}",'
                f'INDIRECT("\'"&{sheet_ref}&"\'!$B:$B"))'
                f'),"")'
            )
            cl = ws.cell(row=row, column=4+j)
            cl.value     = f
            cl.font      = Font(name='Calibri', size=9, color=C['DGRAY'])
            cl.fill      = _fill(CASE_PALETTE[j % len(CASE_PALETTE)].replace(
                                 CASE_PALETTE[j % len(CASE_PALETTE)],
                                 # light tint of the case colour for data rows
                                 alt_bg))
            cl.fill      = _fill(alt_bg)
            cl.alignment = Alignment(horizontal='right', vertical='center')
            cl.border    = _thin()

        rh(ws,row,13); row += 1

    # Composition sections
    all_comp_types = list(dict.fromkeys(
        ct for cl in case_labels
        for sn, sdata in all_cases[cl].items()
        for ct in sdata.get('comp',{}).keys()
    ))
    all_comps_global = list(dict.fromkeys(
        cname for cl in case_labels
        for sn, sdata in all_cases[cl].items()
        for ct in sdata.get('comp',{}).values()
        for cname in ct.keys()
    ))

    if all_comp_types:
        for ct in all_comp_types:
            _sec_row(ws, row, ct, 3+N, C['NAVY'])
            row += 1
            # comp header
            _c(ws,row,2,"  Component",bg=C['LGRAY'],sz=8,bold=True)
            _c(ws,row,3,"section",bg=C['LGRAY'],sz=7,italic=True)
            for j in range(N):
                _c(ws,row,4+j,case_labels[j],
                   bg=CASE_PALETTE[j%len(CASE_PALETTE)],fg=C['WHITE'],sz=8,bold=True,ha='center')
            rh(ws,row,16); row += 1

            # collect comps for this type
            ct_comps = list(dict.fromkeys(
                cname for cl in case_labels
                for sn, sdata in all_cases[cl].items()
                for cname in sdata.get('comp',{}).get(ct,{}).keys()
            ))
            for pi, cname in enumerate(ct_comps):
                alt_bg2 = C['WHITE'] if pi%2==0 else C['LGRAY']
                _c(ws,row,2,f"  {cname}",bg=alt_bg2,sz=8)
                _c(ws,row,3,'',bg=alt_bg2,sz=7)
                # For comps, also use INDIRECT/XMATCH into S_* sheets
                for j in range(N):
                    case_col = get_column_letter(4+j)
                    comp_lbl = f"  {cname}"   # matches what S_* sheet stores
                    f = (
                        f'=IFERROR('
                        f'INDEX('
                        f'INDIRECT("\'"&{sheet_ref}&"\'!${case_col}:${case_col}"),'
                        f'XMATCH("{comp_lbl}",'
                        f'INDIRECT("\'"&{sheet_ref}&"\'!$B:$B"))'
                        f'),"")'
                    )
                    cl = ws.cell(row=row, column=4+j)
                    cl.value = f
                    cl.font  = Font(name='Calibri', size=8, color=C['DGRAY'])
                    cl.fill  = _fill(alt_bg2)
                    cl.alignment = Alignment(horizontal='right')
                    cl.border = _thin()
                rh(ws,row,13); row += 1

    ws.freeze_panes = "B3"

    # Stream name dropdown for the input cell
    if stream_names:
        # Build a named range list in the hidden sheet for DV
        dv = DataValidation(
            type='list',
            formula1=f'_STREAM_IDX!$A$2:$A${len(stream_names)+1}',
            allow_blank=True
        )
        ws.add_data_validation(dv)
        dv.add("D1")

    # Note row at top
    _c(ws,1,4+N if N > 0 else 4,
       "← Edit this cell (dropdown or type stream name)",
       bg=C['WARN'],fg=C['ORANGE'],sz=8,italic=True,border=False)

# ── Per-stream comparison sheet (one per unique stream) ────────────────
def build_stream_compare_sheet(wb, all_cases, all_props, stream_name):
    safe = re.sub(r'[/\\*\[\]:?]','_', stream_name)[:28]
    sheet_name = f"S_{safe}"
    existing   = {s.title for s in wb.worksheets}
    base = sheet_name; n = 1
    while sheet_name in existing:
        sheet_name = f"{base[:26]}_{n}"; n += 1

    case_labels = list(all_cases.keys())
    N = len(case_labels)

    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = C['TEAL']

    cw(ws,'A',3); cw(ws,'B',30); cw(ws,'C',14)
    for j in range(N): cw(ws, 4+j, 20)

    LC = get_column_letter(3+N)

    # Row 1: title
    _sheet_title(ws, f"STREAM: {stream_name}  —  {N}-Case Comparison", 3+N)

    # Row 2: source files
    ws.merge_cells(f"B2:{LC}2")
    sources = " | ".join(
        all_cases[cl].get(stream_name,{}).get('source_file','?')
        for cl in case_labels)
    _c(ws,2,2,sources, bg=C['STEEL'],fg=C['WHITE'],sz=8,italic=True)
    rh(ws,2,14)

    # Row 3: col headers
    hdrs = [("Property", C['NAVY']), ("Phase/Unit", C['NAVY'])]
    for j, cl in enumerate(case_labels):
        sdata = all_cases[cl].get(stream_name)
        ph    = sdata.get('phase','') if sdata else ''
        warn  = ' ⚠' if (sdata and sdata.get('warned')) else ''
        hdrs.append((f"{cl}\n{ph}{warn}", CASE_PALETTE[j%len(CASE_PALETTE)]))
    _col_hdrs(ws, 3, hdrs)

    row = 4
    prev_sec = None

    for prop_lbl, sec in all_props.items():
        # Check if any case has this property for this stream
        vals = {}
        for cl in case_labels:
            v = _get_val(all_cases[cl].get(stream_name), prop_lbl)
            if v is not None: vals[cl] = v
        if not vals: continue

        if sec != prev_sec:
            _sec_row(ws, row, sec, 3+N, SEC_COLORS[sec])
            row += 1; prev_sec = sec

        alt_bg = C['WHITE'] if row%2==0 else SEC_ROW_BG.get(sec, C['LGRAY'])
        _c(ws,row,2,prop_lbl, bg=alt_bg, sz=9, ha='left')
        _c(ws,row,3,'',        bg=alt_bg, fg=C['MGRAY'], sz=8, ha='center')

        for j, cl in enumerate(case_labels):
            v = vals.get(cl)
            if v is None:
                _c(ws,row,4+j,'', bg=C['LGRAY'], sz=8, ha='center')
            elif isinstance(v, str):
                _c(ws,row,4+j,v,  bg=alt_bg, sz=8, ha='center')
            else:
                _c(ws,row,4+j,fmtv(v), bg=alt_bg, sz=9, ha='right')
        rh(ws,row,13); row += 1

    # Composition tables
    all_comp_types = list(dict.fromkeys(
        ct for cl in case_labels
        for ct in all_cases[cl].get(stream_name,{}).get('comp',{}).keys()
    ))
    for ct in all_comp_types:
        _sec_row(ws, row, ct, 3+N, C['NAVY']); row += 1

        _c(ws,row,2,"  Component",bg=C['LGRAY'],sz=8,bold=True)
        _c(ws,row,3,'value',bg=C['LGRAY'],sz=8,italic=True,ha='center')
        for j,cl in enumerate(case_labels):
            _c(ws,row,4+j,cl,
               bg=CASE_PALETTE[j%len(CASE_PALETTE)],fg=C['WHITE'],sz=8,bold=True,ha='center')
        rh(ws,row,16); row += 1

        all_comps = list(dict.fromkeys(
            cname for cl in case_labels
            for cname in all_cases[cl].get(stream_name,{}).get('comp',{}).get(ct,{}).keys()
        ))
        for pi, cname in enumerate(all_comps):
            alt_bg = C['WHITE'] if pi%2==0 else C['LGRAY']
            _c(ws,row,2,f"  {cname}",bg=alt_bg,sz=8)
            _c(ws,row,3,'',bg=alt_bg,sz=8)
            for j,cl in enumerate(case_labels):
                entry = all_cases[cl].get(stream_name,{}).get('comp',{}).get(ct,{}).get(cname,{})
                v = (entry.get('TOTAL') or entry.get('OVERALL')
                     or next(iter(entry.values()),None) if entry else None)
                _c(ws,row,4+j,fmtv(v) if v is not None else '', bg=alt_bg,sz=8,ha='right')
            rh(ws,row,13); row += 1

    ws.freeze_panes = "B4"
    ws.auto_filter.ref = f"B3:{LC}3"


def compare_main(args):
    preselect_stream = None
    if '--stream' in args:
        idx = args.index('--stream')
        if idx + 1 < len(args):
            preselect_stream = args[idx + 1]
            args = [a for i, a in enumerate(args) if i not in (idx, idx + 1)]
    files = []
    if not args:
        files = sorted(glob.glob('HMB_*.xlsx'))
    for a in args:
        if os.path.isdir(a):
            files += sorted(glob.glob(os.path.join(a, 'HMB_*.xlsx'))) or sorted(glob.glob(os.path.join(a, '*.xlsx')))
        elif os.path.isfile(a):
            files.append(a)
        else:
            files += sorted(glob.glob(a))
    if not files:
        print("ERROR: no HMB .xlsx files found.\nUsage: python hmbbuilder.py compare HMB_a.xlsx HMB_b.xlsx "
              "[--stream NAME]")
        return 1
    print(f"\n{'=' * 64}\nHMB multi-case compare — {len(files)} file(s)")
    all_cases, all_streams = OrderedDict(), set()
    for fpath in files:
        print(f"\n  Loading: {os.path.basename(fpath)} ...", flush=True)
        try:
            streams_list, case_label = parse_hmb_workbook(fpath)
            base_cl, n = case_label, 1
            while case_label in all_cases:
                case_label = f"{base_cl} ({n})"
                n += 1
            all_cases[case_label] = {}
            for sdata in streams_list:
                sdata['case_label'] = case_label
                sdata['source_file'] = os.path.basename(fpath)
                all_cases[case_label][sdata['stream']] = sdata
                all_streams.add(sdata['stream'])
            print(f"  OK — case '{case_label}', {len(streams_list)} streams")
        except Exception as e:                         # noqa: BLE001
            import traceback
            print(f"  ERROR: {e}")
            traceback.print_exc()
    if not all_cases:
        print("No cases loaded.")
        return 1
    all_props = collect_all_props(all_cases)
    stream_order = list(dict.fromkeys(sn_ for cl in all_cases for sn_ in all_cases[cl]))
    wb = Workbook()
    del wb['Sheet']
    build_cases_sheet(wb, all_cases)
    build_summary_sheet(wb, all_cases, all_props)
    build_stream_lookup(wb, all_cases, all_props, stream_order, preselect_stream)
    print(f"  Building {len(stream_order)} stream comparison sheets...")
    for sn_ in stream_order:
        build_stream_compare_sheet(wb, all_cases, all_props, sn_)
    front = ["CASES", "SUMMARY", "STREAM_LOOKUP"]
    wb._sheets = [wb[n] for n in front if n in wb.sheetnames] + [w for w in wb._sheets if w.title not in front]
    out_path = os.path.join(_out_folder(*files), f"HMB_MultiCase_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    wb.save(out_path)
    print(f"\n  Saved : {out_path}\n{'=' * 64}")
    return 0


# ════════════════════════════════════════════════════════════════════════
#  9.  CLI
# ════════════════════════════════════════════════════════════════════════
def _opt(argv, name):
    if name in argv:
        i = argv.index(name)
        if i + 1 < len(argv):
            v = argv[i + 1]
            return v, [a for j, a in enumerate(argv) if j not in (i, i + 1)]
    return None, argv


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv[:1] == ["compare"]:
        return compare_main(argv[1:])
    sheet, argv = _opt(argv, "--sheet")
    consts, argv = _opt(argv, "--constants")
    wide = argv[:1] == ["wide"]
    if wide:
        argv = argv[1:]
    paths = [a for a in argv if not a.startswith("-")]
    if not paths:
        cand = [p for p in glob.glob("*.xlsx") if not Path(p).name.startswith(("HMB_", "~$"))]
        if not cand:
            print(__doc__)
            return 1
        paths = [cand[0]]
    dump = paths[0]
    out = paths[1] if len(paths) > 1 else None
    if not os.path.isfile(dump):
        print(f"ERROR: {dump} not found")
        return 1
    if wide:
        print(f"  Saved : {build_wide_workbook(dump, out, sheet, consts)}")
    else:
        build_hmb(dump, out, sheet, consts)
    return 0


if __name__ == "__main__":
    sys.exit(main())
