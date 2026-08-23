"""Build an in-memory, single-sheet openpyxl workbook for one live-resolved
stream (HYSYS/PRO-II COM connection) -- there is no on-disk HMB file for a
live connection to render faithfully, so this synthesizes a sheet in the
SAME shape as a real per-stream HMB sheet (title row, then a
Property | Unit | TOTAL | VAPOR | LIQUID table) so
streams.unit_sheet_model.UnitAwareSheetModel's existing shape-1 detector
picks it up unchanged -- no separate "live" rendering path needed anywhere
else, and the same clickable unit dropdowns work identically.
"""
from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

import engine_api as api

_TITLE_FILL = PatternFill("solid", fgColor="0D0E10")
_HEADER_FILL = PatternFill("solid", fgColor="EAF3FE")
_TITLE_FONT = Font(bold=True, color="FFFFFF", size=10)
_HEADER_FONT = Font(bold=True, size=9)

# (label, total_field, vapor_field, liquid_field, quantity code | None)
# Mirrors the real per-stream sheet's own label vocabulary (so
# engine_api.property_quantity_for recognizes every row exactly as it
# would for a file-based sheet) against whatever StreamProps fields a
# live-resolved stream actually carries.
_ROWS = [
    ("Mass Rate", "total_mass", "vap_mass", "liq_mass", "mflow"),
    ("Std Liq Vol Rate", "total_std_liq", None, None, "qvol"),
    ("Std Vap Vol Rate", "total_std_vap", None, None, "qvol"),
    ("Temperature", "temp_f", None, None, "T"),
    ("Pressure", "pres_psia", None, None, "P"),
    ("Molecular Weight", "mol_weight", "vap_mw", "liq_mw", "MW"),
    ("Specific Enthalpy", None, "vap_sp_enthalpy", "liq_sp_enthalpy", "h"),
    ("Actual Density", "total_density", "vap_density", "liq_density", "rho"),
    ("Viscosity", None, "vap_visc", "liq_visc", "visc"),
    ("Surface Tension", None, None, "liq_surf_tens", "st"),
    ("True Critical Temperature", "tc_f", None, None, "T"),
    ("True Critical Pressure", "pc_psia", None, None, "P"),
    ("Total Z Factor", "total_z", None, None, None),
    ("Vapor Z Factor", None, "vap_z", None, None),
    ("Cp/Cv Ratio", "vap_cp_cv", None, None, None),
    ("Vapor Cp", None, "vap_cp", None, None),
    ("Liquid Cp", None, None, "liq_cp", None),
    ("Vapor Thermal Conductivity", None, "vap_therm_cond", None, None),
    ("Liquid Thermal Conductivity", None, None, "liq_therm_cond", None),
    ("Acentric Factor", "acentric", None, None, None),
]


def build_synthetic_stream_sheet(wb: Workbook, name: str, sp) -> None:
    """Adds one sheet named `name` to `wb`, in the same shape a real
    per-stream HMB sheet uses. `sp` is a StreamProps (already resolved via
    engine_api.resolve_stream_for_snapshot) -- every value is already in
    the engine's internal FPS unit, from engine_api.internal_unit(qty)."""
    ws = wb.create_sheet(title=str(name)[:31])

    title = f"STREAM: {name}  |  Phase: {sp.phase or '—'}  |  live connection"
    ws.cell(row=1, column=2, value=title)
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=6)
    for c in range(2, 7):
        cell = ws.cell(row=1, column=c)
        cell.fill = _TITLE_FILL
        cell.font = _TITLE_FONT
        cell.alignment = Alignment(horizontal="left" if c == 2 else "center",
                                   vertical="center")

    headers = ["", "Property", "Unit", "TOTAL", "VAPOR", "LIQUID"]
    for c, h in enumerate(headers, start=1):
        if not h:
            continue
        cell = ws.cell(row=2, column=c, value=h)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

    r = 3
    for label, tot_f, vap_f, liq_f, qty in _ROWS:
        ws.cell(row=r, column=2, value=label)
        unit = api.raw_hmb_unit_for(qty) if qty else "-"
        ws.cell(row=r, column=3, value=unit or "-")
        for col, field in ((4, tot_f), (5, vap_f), (6, liq_f)):
            if field is None:
                continue
            val = getattr(sp, field, None)
            if val is not None:
                ws.cell(row=r, column=col, value=val)
        r += 1

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 10
    for col in "DEF":
        ws.column_dimensions[col].width = 14
