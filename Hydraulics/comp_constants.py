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

# ── Twu (1984) correlation ─────────────────────────────────────────────
# Twu, C.H., "An internally consistent correlation for predicting the
# critical properties and molecular weights of petroleum and coal-tar
# liquids," Fluid Phase Equilibria, 16 (1984) 137-150.
# Used for heavy fractions (MW >= TWU_MW_THRESHOLD) where Lee-Kesler
# degrades (flagged via the Pc < 10 psia warning below).
TWU_MW_THRESHOLD = 265.0

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

def estimate_twu_props(tb_r, sg):
    """
    Twu (1984) perturbation about the n-alkane reference system.
    Given Tb(°R) and SG(60°F), returns (Tc_R, Vc_ft3lbmol, Pc_psia, MW)
    or None if the reference solve fails.
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
    f_t = dsg_t * (-0.362456/tb_r**0.5
                   + (0.0398285 - 0.948125/tb_r**0.5)*dsg_t)
    tc = tc0 * ((1 + 2*f_t)/(1 - 2*f_t))**2

    # Eqns (14)-(16): critical volume
    dsg_v = math.exp(4*(sg0**2 - sg**2)) - 1
    f_v = dsg_v * (0.466590/tb_r**0.5
                   + (-0.182421 + 3.01721/tb_r**0.5)*dsg_v)
    vc = vc0 * ((1 + 2*f_v)/(1 - 2*f_v))**2

    # Eqns (17)-(19): critical pressure
    dsg_p = math.exp(0.5*(sg0 - sg)) - 1
    f_p = dsg_p * ((2.53262 - 46.1955/tb_r**0.5 - 0.00127885*tb_r)
                   + (-11.4277 + 252.140/tb_r**0.5 + 0.00230535*tb_r)*dsg_p)
    pc = pc0 * (tc/tc0) * (vc0/vc) * ((1 + 2*f_p)/(1 - 2*f_p))**2

    # Eqns (20)-(23): molecular weight
    dsg_m = math.exp(5*(sg0 - sg)) - 1
    abs_x = abs(0.0123420 - 0.328086/tb_r**0.5)
    f_m = dsg_m * (abs_x + (-0.0175691 + 0.193168/tb_r**0.5)*dsg_m)
    ln_mw = math.log(mw0) * ((1 + 2*f_m)/(1 - 2*f_m))**2
    mw = math.exp(ln_mw)

    return tc, vc, pc, mw

def estimate_pseudo_props(mw, nbp_f, sld_lbft3):
    """
    Given MW, NBP(°F), SLD(lb/ft³) for a pseudo-component,
    estimate Tc(°F), Pc(psia), Vc(ft³/lbmol), Zc, ω.
    Uses Twu (1984) for MW >= TWU_MW_THRESHOLD (more reliable for heavy
    fractions), Lee-Kesler (1975) + Edmister (1958) otherwise.
    Returns dict.
    """
    if None in (mw, nbp_f, sld_lbft3) or sld_lbft3 <= 0:
        return {}
    sg  = sld_lbft3 / WATER_DENSITY_60F
    tb_r = nbp_f + 459.67   # °F → °R
    if tb_r <= 0 or sg <= 0:
        return {}

    if mw >= TWU_MW_THRESHOLD:
        twu = estimate_twu_props(tb_r, sg)
        if twu is not None:
            tc_r, vc, pc_psia, mw_twu = twu
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
                'method':   'Twu (1984)',
            }

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
                base_method = est.get('method', 'Lee-Kesler + Edmister')
                # Flag unreliable estimates (very heavy fractions where the
                # correlation in use breaks down)
                if est.get('Pc_psia') is not None and est['Pc_psia'] < 10:
                    entry['method'] = f'{base_method} (⚠ unreliable — Pc < 10 psia, use PROII Extracted)'
                elif est.get('omega') is None:
                    entry['method'] = f'{base_method} (⚠ partial — Edmister failed, ω unavailable)'
                else:
                    entry['method'] = base_method
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
       f"Lee-Kesler (1975)+Edmister (1958) below MW {TWU_MW_THRESHOLD:.0f}, "
       f"Twu (1984) at/above, for pseudo-fractions",
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
        ('ESTIMATED',C['ESTIM'],  f'Lee-Kesler+Edmister (MW<{TWU_MW_THRESHOLD:.0f}) or Twu 1984 (MW>={TWU_MW_THRESHOLD:.0f}) — use for EOS flash'),
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
