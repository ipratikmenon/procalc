#!/usr/bin/env python3
"""
Component Constants Builder  —  HyCalign
==========================================
Reads Constants.xlsx, estimates Tc/Pc/Vc/ω for pseudo-components using
the Twu (1984) correlation, and either:
  (a) Enriches an existing HMB workbook by adding a COMP_CONSTANTS sheet
  (b) Produces a standalone enriched Constants file

For pseudo-components (petroleum fractions), given:
  MW          — molecular weight (g/mol)
  NBP (°F)    — normal boiling point
  SLD (lb/ft³)— standard liquid density at 60°F

Twu, C.H. (1984), Fluid Phase Equilibria 16: 137-150. Perturbs a
hypothetical n-alkane reference (same Tb) by the real component's
specific-gravity deviation to get Tc, Vc, Pc, MW; ω then follows from
PRO/II's SIMSCI/TWU generalized Frost-Kalkwarf-Thodos vapor pressure
correlation (back-solving Ω at the NBP and applying Pitzer's definition
at Tr = 0.7).

The Twu perturbation terms (f_T, f_V, f_P, f_M) are clamped (see
TWU_F_CLAMP) to avoid the (1+2f)/(1-2f) re-summation singularity for
components far outside the correlation's fitted range.

Usage:
    python comp_constants.py Constants.xlsx                  # standalone enriched xlsx
    python comp_constants.py Constants.xlsx HMB.xlsx         # patch HMB in-place
"""

import os, sys, re, math
from collections import OrderedDict
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

# ── Direct empirical curve fit (single formula, Watson K 6.76-32.42) ────
# Twu (1984) is fitted/tested over Watson K ~ 8-14, but real assay
# pseudo-components routinely fall well outside that (dense/aromatic
# cuts below K~9, or highly paraffinic heavy ends above K~14, where raw
# Twu's perturbation re-summation runs away -- up to 600-800% error, see
# TWU_F_CLAMP).
#
# Earlier iterations fit three separate Watson-K bands. This is the same
# underlying regression -- ln(Tc_R)/ln(Pc_psia) by OLS against Tb(R),
# SG, and MW, bypassing Twu's n-alkane reference machinery entirely --
# but fit ONCE against the union of all three bands' PRO/II ground truth
# (220 points total: 5 low-K + 192 mid-K + 23 high-K, spanning Watson K
# 7.08-32.42), so one continuous formula covers the whole range with no
# band-boundary discontinuity. Getting there needed a richer feature set
# (quadratic MW cross-terms) than any single band alone required, since
# one polynomial now has to track curvature across a >2x wider K range
# at once. 14-term fits, 206 DOF: max error 0.23% (Tc) / 1.55% (Pc) --
# comparable to the old per-band fits' accuracy (which were 0.04-0.74%)
# despite covering 4x the K range in one formula. This is curve-fitting
# against real assay data, not a physical model -- only trustworthy
# inside K=6.76-32.42 (the range its fit data spans); outside it, falls
# back to raw Twu (1984), flagged as unverified.
def _watson_k(tb_r, sg):
    return tb_r**(1/3) / sg

def _ln_poly_tb_sg_mw(tb_r, sg, mw, c):
    return (c[0] + c[1]*tb_r + c[2]*tb_r**2 + c[3]*tb_r**3
            + c[4]*sg + c[5]*sg**2 + c[6]*tb_r*sg + c[7]*tb_r**2*sg
            + c[8]*mw + c[9]*mw**2 + c[10]*tb_r*mw + c[11]*sg*mw
            + c[12]*tb_r**2*mw + c[13]*sg**2*mw)

CURVEFIT_K_MIN = 6.7607782905215785
CURVEFIT_K_MAX = 32.42
_CF_TC_COEF = (5.007883101013148, 0.003450517505203951,
               -5.350561845008301e-08, -1.074024144669105e-09,
               1.6490962015029442, -0.5550533538284068,
               -0.0015136220956506788, 6.726083603733643e-07,
               -0.01062975296484001, -1.8505676528363376e-06,
               1.1442752460598282e-05, -0.002262849729260199,
               -1.5958487089281692e-09, 0.0015542330963586107)
_CF_PC_COEF = (5.701626621482552, -0.0073214380294048045,
               2.2394905038568905e-05, -1.4116473210398495e-08,
               9.73621911814537, -3.3041596315868413,
               -0.00870625333062654, 3.7745792200831124e-06,
               -0.055143811415915074, -1.7498645134846417e-05,
               4.5674217570408565e-05, -0.011989311917111075,
               4.171268477379032e-09, 0.008258835869442238)

def estimate_pseudo_props(mw, nbp_f, sld_lbft3):
    """
    Given MW, NBP(°F), SLD(lb/ft³) for a pseudo-component, estimate
    Tc(°F), Pc(psia), Vc(ft³/lbmol), Zc, ω. Tc/Pc come from a single
    direct Tb/SG/MW curve fit within the PRO/II-verified Watson-K range
    (see CURVEFIT_K_MIN/MAX above); outside it, falls back to raw Twu
    (1984), flagged as unverified. Vc/MW/omega always come from Twu.
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

    K = _watson_k(tb_r, sg)
    in_curvefit = CURVEFIT_K_MIN <= K <= CURVEFIT_K_MAX

    if in_curvefit:
        tc_r    = math.exp(_ln_poly_tb_sg_mw(tb_r, sg, mw, _CF_TC_COEF))
        pc_psia = math.exp(_ln_poly_tb_sg_mw(tb_r, sg, mw, _CF_PC_COEF))
        method  = 'Direct Tb/SG/MW curve fit (Watson K 6.76-32.42)'
    else:
        method = 'Twu (1984), Watson K outside verified curve-fit range'

    tc_f = tc_r - 459.67
    zc    = pc_psia * vc / (10.7316 * tc_r) if pc_psia and vc else 0.27
    omega = estimate_acentric(tb_r, tc_r, pc_psia)
    extrapolated = was_clamped or not in_curvefit

    return {
        'Tc_F':     round(tc_f, 4),
        'Tc_R':     round(tc_r, 4),
        'Pc_psia':  round(pc_psia, 4),
        'Vc_ft3lbmol': round(vc, 4) if vc else None,
        'Zc':       round(zc, 4),
        'omega':    round(omega, 6) if omega is not None else None,
        'SG':       round(sg, 6),
        'extrapolated': extrapolated,
        'method':   method,
    }

# ── Read Constants.xlsx ────────────────────────────────────────────────
def read_constants(path):
    """
    Returns list of dicts, one per component:
      name, CAS, MW, SLD, NBP_F, Tc_F, Pc_psia, Vc, Zc, omega,
      is_pseudo, estimated (bool), method
    """
    from openpyxl import load_workbook
    _wb = load_workbook(path, read_only=True, data_only=True)
    _ws = _wb.active
    rows = list(_ws.iter_rows(values_only=True))
    _wb.close()
    results = []

    def _cell(r, c):
        return r[c] if (r is not None and c < len(r)) else None

    for ri in range(2, len(rows)):
        row = rows[ri]
        c0 = _cell(row, 0)
        name = str(c0).strip() if c0 is not None else ''
        if name in ('nan','') or not name: continue

        mw   = sf(_cell(row, 2))
        sld  = sf(_cell(row, 3))
        nbp  = sf(_cell(row, 4))
        tc   = sf(_cell(row, 5))
        pc   = sf(_cell(row, 6))
        vc   = sf(_cell(row, 7))
        zc   = sf(_cell(row, 8))
        omega= sf(_cell(row, 9))
        cas  = str(_cell(row, 1)).strip() if _cell(row, 1) is not None else ''
        pr_pen = sf(_cell(row, 10))
        srk_pen= sf(_cell(row, 11))
        parachor = sf(_cell(row, 12))

        is_pseudo = (tc is None or str(_cell(row, 5)).strip().lower() == 'missing')

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
                base_method = est.get('method', 'Twu (1984)')
                # Flag unreliable estimates (very heavy fractions where the
                # correlation in use breaks down, or inputs whose Watson K
                # lies far outside Twu's fitted ~8-14 range, where the
                # perturbation terms had to be clamped)
                if est.get('Pc_psia') is not None and est['Pc_psia'] < 10:
                    entry['method'] = f'{base_method} (⚠ unreliable — Pc < 10 psia, use PROII Extracted)'
                elif est.get('extrapolated'):
                    entry['method'] = f'{base_method} (⚠ extrapolated — Watson K outside Twu\'s fitted/calibrated range)'
                elif est.get('omega') is None:
                    entry['method'] = f'{base_method} (⚠ partial — Frost-Kalkwarf-Thodos failed, ω unavailable)'
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
       f"Twu (1984) + PRO/II SIMSCI/TWU vapor pressure ω for pseudo-fractions",
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
       f"Pseudo: {n_pseudo} ({n_estimated} estimated by Twu+SIMSCI-ω)  |  "
       f"Estimated cells shown in yellow italic",
       bg=C['LGRAY'], fg=C['DGRAY'], sz=8, italic=True)
    rh(ws,r,15)

    # Legend
    r += 2
    for lbl, bg, desc in [
        ('PURE',     C['PURE'],   'Library component — Tc/Pc/ω from PRO/II database'),
        ('PSEUDO',   C['PSEUDO'], 'Petroleum fraction — estimated from MW + NBP + SLD'),
        ('ESTIMATED',C['ESTIM'],  'Twu (1984) + PRO/II SIMSCI/TWU vapor pressure ω — use for EOS flash'),
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
