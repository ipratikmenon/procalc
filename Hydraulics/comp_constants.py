#!/usr/bin/env python3
"""
Component Constants Builder  —  HyCalign
==========================================
Reads Constants.xlsx, estimates Tc/Pc/ω for pseudo-components using
Lee-Kesler (1975) + Edmister (1958) correlations, and either:
  (a) Enriches an existing HMB workbook by adding a COMP_CONSTANTS sheet
  (b) Produces a standalone enriched Constants file

For pseudo-components (petroleum fractions), given:
  MW          — molecular weight (g/mol)
  NBP (°F)    — normal boiling point
  SLD (lb/ft³)— standard liquid density at 60°F

Lee-Kesler (1975):
  SG  = SLD / 62.428
  Tb  = NBP + 459.67  (°R)
  Tc  = 341.7 + 811.1·SG + (0.4244 + 0.1174·SG)·Tb
        + (0.4669 - 3.2623·SG)·1e5/Tb                       [°R → °F]
  ln(Pc) = 8.3634 - 0.0566/SG
           - (0.24244 + 2.2898/SG + 0.11857/SG²)·1e-3·Tb
           + (1.4685 + 3.648/SG + 0.47227/SG²)·1e-7·Tb²
           - (0.42019 + 1.6977/SG²)·1e-10·Tb³               [psia]

Edmister (1958):
  ω = (3/7) · log10(Pc/14.696) / (Tc/Tb - 1) - 1

Usage:
    python comp_constants.py Constants.xlsx                  # standalone enriched xlsx
    python comp_constants.py Constants.xlsx HMB.xlsx         # patch HMB in-place
"""

import os, sys, re, math
from collections import OrderedDict
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell

# ── Colours ────────────────────────────────────────────────────────────
C = dict(
    NAVY="1F4973", STEEL="2E75B6", WHITE="FFFFFF", LGRAY="F2F2F2",
    MGRAY="A0A0A0", DGRAY="404040", ORANGE="C55A11", AMBER="FFF2CC",
    GREEN="E2EFDA", GRNHDR="375623", RED="FCE4D6", LTBLUE="DEEAF1",
    PURPLE="7030A0", TEAL="1F7070", FORMULA="FFFDE7",
    PURE="D6E4F7", PSEUDO="FFF2CC", ESTIM="FFFDE7",
)

def _fill(h):  return PatternFill("solid", fgColor=h.lstrip('#'))
def _thin():
    s = Side(style="thin", color="D0D0D0")
    return Border(left=s, right=s, top=s, bottom=s)

def _c(ws, r, c, v=None, bg=None, fg=None, sz=9,
       bold=False, italic=False, wrap=False, ha='left', va='center'):
    cl = ws.cell(row=r, column=c)
    if isinstance(cl, MergedCell): return cl
    if v is not None: cl.value = v
    cl.font = Font(name="Calibri", size=sz, bold=bold, italic=italic,
                   color=(fg or C['DGRAY']))
    if bg: cl.fill = _fill(bg)
    cl.alignment = Alignment(horizontal=ha, vertical=va, wrap_text=wrap)
    cl.border = _thin()
    return cl

def cw(ws, col, w):
    ws.column_dimensions[get_column_letter(col) if isinstance(col,int) else col].width = w

def rh(ws, row, h): ws.row_dimensions[row].height = h

def sf(v):
    if v is None: return None
    if isinstance(v, bool): return None
    s = str(v).strip()
    if s.lower() in ('nan','missing','','none','n/a'): return None
    try: return float(s)
    except: return None

def fmtv(v, dp=6):
    if v is None: return ''
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v): return ''
        if v == 0.0: return 0.0
        a = abs(v)
        if a < 1e-3 or a >= 1e8: return float(f'{v:.{dp}E}')
        return round(v, dp)
    return v

# ── Lee-Kesler (1975) correlations ────────────────────────────────────
WATER_DENSITY_60F = 62.428   # lb/ft³ at 60°F

def estimate_tc_r(tb_r, sg):
    """Lee-Kesler Tc in °R."""
    return (341.7 + 811.1*sg
            + (0.4244 + 0.1174*sg)*tb_r
            + (0.4669 - 3.2623*sg)*1e5/tb_r)

def estimate_pc_psia(tb_r, sg):
    """Lee-Kesler Pc in psia."""
    ln_pc = (8.3634
             - 0.0566/sg
             - (0.24244 + 2.2898/sg + 0.11857/(sg**2))*1e-3*tb_r
             + (1.4685 + 3.648/sg + 0.47227/(sg**2))*1e-7*(tb_r**2)
             - (0.42019 + 1.6977/(sg**2))*1e-10*(tb_r**3))
    return math.exp(ln_pc)

def estimate_vc_ftlbmol(tc_r, pc_psia):
    """Rough Vc from Tc and Pc using Zc≈0.27."""
    # Vc = Zc·R·Tc/Pc
    R_psia_ft3 = 10.7316   # psia·ft³/(lbmol·°R)
    Zc = 0.27
    return Zc * R_psia_ft3 * tc_r / pc_psia if pc_psia > 0 else None

def estimate_acentric(tb_r, tc_r, pc_psia):
    """Edmister (1958) acentric factor."""
    if tc_r <= tb_r or pc_psia <= 14.696: return None
    theta = tc_r / tb_r - 1.0
    if theta <= 0: return None
    return (3.0/7.0) * math.log10(pc_psia / 14.696) / theta - 1.0

def estimate_pseudo_props(mw, nbp_f, sld_lbft3):
    """
    Given MW, NBP(°F), SLD(lb/ft³) for a pseudo-component,
    estimate Tc(°F), Pc(psia), Vc(ft³/lbmol), Zc, ω.
    Returns dict.
    """
    if None in (mw, nbp_f, sld_lbft3) or sld_lbft3 <= 0:
        return {}
    sg  = sld_lbft3 / WATER_DENSITY_60F
    tb_r = nbp_f + 459.67   # °F → °R
    if tb_r <= 0 or sg <= 0:
        return {}

    tc_r    = estimate_tc_r(tb_r, sg)
    tc_f    = tc_r - 459.67
    pc_psia = estimate_pc_psia(tb_r, sg)
    vc      = estimate_vc_ftlbmol(tc_r, pc_psia) if pc_psia > 0 else None
    zc      = 0.27   # assumed
    omega   = estimate_acentric(tb_r, tc_r, pc_psia)

    return {
        'Tc_F':     round(tc_f, 4),
        'Tc_R':     round(tc_r, 4),
        'Pc_psia':  round(pc_psia, 4),
        'Vc_ft3lbmol': round(vc, 4) if vc else None,
        'Zc':       round(zc, 4),
        'omega':    round(omega, 6) if omega is not None else None,
        'SG':       round(sg, 6),
        'method':   'Lee-Kesler (1975) + Edmister (1958)',
    }

# ── Read Constants.xlsx ────────────────────────────────────────────────
def read_constants(path):
    """
    Returns list of dicts, one per component:
      name, CAS, MW, SLD, NBP_F, Tc_F, Pc_psia, Vc, Zc, omega,
      is_pseudo, estimated (bool), method
    """
    df = pd.read_excel(path, header=None, dtype=object)
    results = []

    for ri in range(2, len(df)):
        name = str(df.iloc[ri, 0]).strip()
        if name in ('nan','') or not name: continue

        mw   = sf(df.iloc[ri, 2])
        sld  = sf(df.iloc[ri, 3])
        nbp  = sf(df.iloc[ri, 4])
        tc   = sf(df.iloc[ri, 5])
        pc   = sf(df.iloc[ri, 6])
        vc   = sf(df.iloc[ri, 7])
        zc   = sf(df.iloc[ri, 8])
        omega= sf(df.iloc[ri, 9])
        cas  = str(df.iloc[ri, 1]).strip() if df.iloc[ri, 1] is not None else ''
        pr_pen = sf(df.iloc[ri, 10])
        srk_pen= sf(df.iloc[ri, 11])
        parachor = sf(df.iloc[ri, 12])

        is_pseudo = (tc is None or str(df.iloc[ri, 5]).strip().lower() == 'missing')

        entry = {
            'name': name, 'CAS': cas, 'MW': mw,
            'SLD_lbft3': sld, 'NBP_F': nbp,
            'Tc_F': tc, 'Pc_psia': pc,
            'Vc_ft3lbmol': vc, 'Zc': zc, 'omega': omega,
            'PR_Peneloux': pr_pen, 'SRK_Peneloux': srk_pen,
            'Parachor': parachor,
            'is_pseudo': is_pseudo,
            'estimated': False, 'method': 'Library',
        }

        # Estimate missing properties for pseudos
        if is_pseudo and mw and nbp is not None and sld:
            est = estimate_pseudo_props(mw, nbp, sld)
            if est:
                entry['Tc_F']          = est.get('Tc_F')
                entry['Pc_psia']       = est.get('Pc_psia')
                entry['Vc_ft3lbmol']   = est.get('Vc_ft3lbmol')
                entry['Zc']            = est.get('Zc')
                entry['omega']         = est.get('omega')
                entry['SG']            = est.get('SG')
                entry['estimated']     = True
                # Flag unreliable estimates (very heavy fractions where LK breaks down)
                if est.get('Pc_psia') is not None and est['Pc_psia'] < 10:
                    entry['method'] = 'Lee-Kesler (⚠ unreliable — Pc < 10 psia, use PROII Extracted)'
                elif est.get('omega') is None:
                    entry['method'] = 'Lee-Kesler (⚠ partial — Edmister failed, ω unavailable)'
                else:
                    entry['method'] = est.get('method', 'Lee-Kesler + Edmister')
        elif not is_pseudo:
            # Pure components: SLD from Constants.xlsx is in PRO/II internal units
            # (not lb/ft3), so SG conversion not applicable. Tc/Pc/omega used directly.
            entry['SG'] = None

        results.append(entry)

    return results

# ── Build COMP_CONSTANTS sheet ─────────────────────────────────────────
COMP_HEADERS = [
    ('Name',            14, 'left'),
    ('Type',             8, 'center'),
    ('CAS',             14, 'center'),
    ('MW\n(g/mol)',      10, 'right'),
    ('NBP\n(°F)',        10, 'right'),
    ('SLD\n(lb/ft³)',    10, 'right'),
    ('SG\n(60°F)',       8,  'right'),
    ('Tc\n(°F)',         10, 'right'),
    ('Pc\n(psia)',       10, 'right'),
    ('Vc\n(ft³/lbmol)',  12, 'right'),
    ('Zc',               8,  'right'),
    ('ω\n(acentric)',    10, 'right'),
    ('PR Peneloux',      12, 'right'),
    ('SRK Peneloux',     12, 'right'),
    ('Parachor',         10, 'right'),
    ('Method',           28, 'left'),
]

def build_comp_constants_sheet(wb, comp_list, position='after_components'):
    """
    Adds a COMP_CONSTANTS sheet to an openpyxl Workbook.
    comp_list: list of dicts from read_constants().
    """
    # Remove existing sheet if present
    if 'COMP_CONSTANTS' in [s.title for s in wb.worksheets]:
        del wb['COMP_CONSTANTS']

    ws = wb.create_sheet('COMP_CONSTANTS')
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = C['PURPLE']

    NC = len(COMP_HEADERS)
    LC = get_column_letter(NC + 1)

    # Title
    cw(ws, 'A', 3)
    ws.merge_cells(f"B1:{LC}1")
    _c(ws,1,2,
       f"COMPONENT CONSTANTS  —  {len(comp_list)} components  |  "
       f"Lee-Kesler (1975) + Edmister (1958) for pseudo-fractions",
       bg=C['NAVY'], fg=C['WHITE'], sz=11, bold=True)
    rh(ws,1,22)

    # Column headers
    for ci, (hdr, w, ha) in enumerate(COMP_HEADERS, 2):
        _c(ws,2,ci,hdr, bg=C['NAVY'],fg=C['WHITE'],sz=8,bold=True,ha='center',wrap=True)
        cw(ws, ci, w)
    rh(ws,2,34)
    ws.auto_filter.ref = f"B2:{LC}2"
    ws.freeze_panes = "C3"

    # Data rows
    n_pure, n_pseudo, n_estimated = 0, 0, 0
    for ri, comp in enumerate(comp_list):
        r = ri + 3
        is_p = comp.get('is_pseudo', False)
        is_e = comp.get('estimated', False)
        if is_p: n_pseudo += 1
        else:    n_pure += 1
        if is_e: n_estimated += 1

        bg_name = C['PSEUDO'] if is_p else C['PURE']
        bg_data = C['ESTIM']  if is_e else (C['LGRAY'] if ri%2 else C['WHITE'])

        vals = [
            comp['name'],
            'PSEUDO' if is_p else 'PURE',
            comp.get('CAS',''),
            fmtv(comp.get('MW'), 4),
            fmtv(comp.get('NBP_F'), 2),
            fmtv(comp.get('SLD_lbft3'), 3),
            fmtv(comp.get('SG'), 4),
            fmtv(comp.get('Tc_F'), 2),
            fmtv(comp.get('Pc_psia'), 2),
            fmtv(comp.get('Vc_ft3lbmol'), 4),
            fmtv(comp.get('Zc'), 4),
            fmtv(comp.get('omega'), 6),
            fmtv(comp.get('PR_Peneloux'), 6),
            fmtv(comp.get('SRK_Peneloux'), 6),
            fmtv(comp.get('Parachor'), 3),
            comp.get('method',''),
        ]
        for ci, v in enumerate(vals, 2):
            bg = bg_name if ci == 2 else bg_data
            _c(ws,r,ci, v, bg=bg, sz=8,
               ha=COMP_HEADERS[ci-2][2],
               bold=(ci==2), italic=is_e and ci >= 9 and ci <= 13)
        rh(ws,r,13)

    # Summary row
    r = len(comp_list) + 4
    ws.merge_cells(f"B{r}:{LC}{r}")
    _c(ws,r,2,
       f"  Total: {len(comp_list)} components  |  "
       f"Pure: {n_pure} (library)  |  "
       f"Pseudo: {n_pseudo} ({n_estimated} estimated by LK+Edmister)  |  "
       f"Estimated cells shown in yellow italic",
       bg=C['LGRAY'], fg=C['DGRAY'], sz=8, italic=True)
    rh(ws,r,15)

    # Legend
    r += 2
    for lbl, bg, desc in [
        ('PURE',     C['PURE'],   'Library component — Tc/Pc/ω from PRO/II database'),
        ('PSEUDO',   C['PSEUDO'], 'Petroleum fraction — estimated from MW + NBP + SLD'),
        ('ESTIMATED',C['ESTIM'],  'Lee-Kesler (Tc, Pc) + Edmister (ω) — use for EOS flash'),
    ]:
        _c(ws,r,2,lbl,  bg=bg, sz=8, bold=True, ha='center')
        _c(ws,r,3,desc, bg=C['WHITE'], sz=8, italic=True)
        ws.merge_cells(f"C{r}:{LC}{r}")
        rh(ws,r,14); r += 1

    print(f"  COMP_CONSTANTS: {len(comp_list)} components "
          f"({n_pure} pure, {n_pseudo} pseudo, {n_estimated} estimated)")

    return ws

# ── Patch an existing HMB workbook ─────────────────────────────────────
def patch_hmb(hmb_path, constants_path, out_path=None):
    """
    Reads constants, adds COMP_CONSTANTS sheet to existing HMB workbook.
    Saves to out_path (default: overwrites in-place).
    """
    print(f"\n  Constants: {os.path.basename(constants_path)}")
    comp_list = read_constants(constants_path)
    print(f"  Read {len(comp_list)} components")

    print(f"  HMB: {os.path.basename(hmb_path)}")
    wb = load_workbook(hmb_path)

    build_comp_constants_sheet(wb, comp_list)

    # Position after COMPONENTS sheet
    target_idx = None
    for i, ws in enumerate(wb.worksheets):
        if ws.title == 'COMPONENTS':
            target_idx = i + 1
            break
    if target_idx:
        wb.move_sheet('COMP_CONSTANTS', offset=target_idx - len(wb.worksheets) + 1)

    if out_path is None:
        out_path = hmb_path
    wb.save(out_path)
    print(f"  Saved: {out_path}")
    return out_path, comp_list

# ── Build dict for lookup by name ──────────────────────────────────────
def comp_dict(comp_list):
    """Returns {name.upper(): comp_entry} for fast lookup."""
    return {c['name'].upper(): c for c in comp_list}

# ── Main ───────────────────────────────────────────────────────────────
def main():
    args = sys.argv[1:]
    if not args:
        print("Usage:")
        print("  python comp_constants.py Constants.xlsx              # standalone")
        print("  python comp_constants.py Constants.xlsx HMB.xlsx     # patch HMB")
        print("  python comp_constants.py Constants.xlsx HMB.xlsx out.xlsx")
        sys.exit(1)

    constants_path = args[0]
    if not os.path.exists(constants_path):
        print(f"ERROR: {constants_path} not found"); sys.exit(1)

    if len(args) >= 2:
        hmb_path = args[1]
        out_path = args[2] if len(args) >= 3 else None
        patch_hmb(hmb_path, constants_path, out_path)
    else:
        # Standalone: produce enriched constants xlsx
        comp_list = read_constants(constants_path)
        wb = Workbook()
        del wb['Sheet']
        build_comp_constants_sheet(wb, comp_list)
        base = os.path.splitext(os.path.basename(constants_path))[0]
        out_dir = os.getcwd() if os.access(os.getcwd(), os.W_OK) else os.path.expanduser('~')
        out_path = os.path.join(out_dir, base + '_enriched.xlsx')
        wb.save(out_path)
        print(f"\n  Saved: {out_path}")

if __name__ == "__main__":
    main()
