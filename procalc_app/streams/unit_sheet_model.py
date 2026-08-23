"""A SheetTableModel that additionally makes each qty-bearing row's Unit
cell a clickable, live-converting unit dropdown -- layered directly on top
of the real HMB sheet, not a hand-built reconstruction.

Only rows whose Property label maps to a real quantity code
(engine_api.property_quantity_for) AND whose own printed Unit cell
normalizes to a unit engine_api actually knows (engine_api.
normalize_hmb_unit) get a dropdown; every other cell renders byte-
identical to the raw file. Conversion is a pure per-row display overlay --
the underlying openpyxl worksheet (the source of truth) is never mutated.

Whether a sheet has any interactive rows at all falls out naturally from
its own shape: a sheet needs a "Property"/"Unit" header (columns B/C) AND
at least one value column past it (column D+) -- which is true for every
per-stream sheet and OUTPUT, and false for UNITS (Property/Unit header but
no value columns), COMPONENTS/COMP_CONSTANTS, and INPUT (no Property/Unit
header at all). No sheet-name special-casing anywhere in this module.
"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QComboBox

import engine_api as api
from results.workbook_model import SheetTableModel, _fmt

_PROPERTY_COL = 2      # 1-indexed column B
_UNIT_COL = 3          # 1-indexed column C
_FIRST_VALUE_COL = 4   # 1-indexed column D


class UnitAwareSheetModel(SheetTableModel):
    def __init__(self, ws, parent=None):
        super().__init__(ws, parent)
        self.unit_col = _UNIT_COL - 1   # 0-indexed, for the view layer
        self._native: dict[int, tuple[str, str]] = {}   # row0 -> (qty, native_unit)
        self._chosen: dict[int, str] = {}                # row0 -> currently-picked unit
        if self.ncols < _FIRST_VALUE_COL:
            return   # no value columns at all (e.g. UNITS) -- nothing to make interactive
        header_prop = ws.cell(row=2, column=_PROPERTY_COL).value
        header_unit = ws.cell(row=2, column=_UNIT_COL).value
        if (str(header_prop or "").strip().lower() != "property"
                or str(header_unit or "").strip().lower() != "unit"):
            return   # not a Property/Unit-shaped sheet (e.g. INPUT) -- stay plain
        for r in range(3, self.nrows + 1):    # data rows start at Excel row 3
            if (r, _PROPERTY_COL) in self.covered:
                continue
            label = ws.cell(row=r, column=_PROPERTY_COL).value
            if not label:
                continue
            qty = api.property_quantity_for(str(label))
            if not qty:
                continue
            raw_unit = ws.cell(row=r, column=_UNIT_COL).value
            native = api.normalize_hmb_unit(qty, raw_unit) if raw_unit else None
            if not native:
                continue
            row0 = r - 1
            self._native[row0] = (qty, native)
            self._chosen[row0] = native

    def unit_rows(self) -> dict[int, tuple[str, str]]:
        """{0-indexed row: (qty, native_unit)} for every row with a working
        clickable unit dropdown."""
        return dict(self._native)

    def set_row_unit(self, row0: int, new_unit: str):
        entry = self._native.get(row0)
        if not entry or new_unit == self._chosen.get(row0):
            return
        self._chosen[row0] = new_unit
        first = self.index(row0, _FIRST_VALUE_COL - 1)
        last = self.index(row0, self.ncols - 1)
        self.dataChanged.emit(first, last)

    def data(self, index, role=Qt.DisplayRole):
        r0, c0 = index.row(), index.column()
        entry = self._native.get(r0)
        if entry:
            if c0 == self.unit_col and role == Qt.DisplayRole:
                # the combo widget IS this cell's visible content -- blank
                # the model's own text so it doesn't show through/behind it
                # (setIndexWidget overlays a widget, it doesn't hide data()).
                return ""
            if role == Qt.DisplayRole and c0 >= _FIRST_VALUE_COL - 1:
                qty, native = entry
                chosen = self._chosen.get(r0, native)
                if chosen != native:
                    cell = self.ws.cell(row=r0 + 1, column=c0 + 1)
                    if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                        converted = api.convert(qty, cell.value, native, chosen)
                        return _fmt(converted, cell.number_format)
        return super().data(index, role)


def install_unit_dropdowns(view, model: UnitAwareSheetModel):
    """Put a real, always-visible unit-picker QComboBox in every row's Unit
    cell that model.unit_rows() reports -- the same "persistent real widget
    in a cell" technique the retired reconstructed Streams table used
    (QTableWidget.setCellWidget), via QTableView's equivalent
    (setIndexWidget)."""
    for row0, (qty, native) in model.unit_rows().items():
        combo = QComboBox()
        combo.setObjectName("UnitPicker")
        combo.addItems(api.units_for(qty))
        i = combo.findText(native)
        if i >= 0:
            combo.setCurrentIndex(i)
        combo.currentTextChanged.connect(
            lambda u, r=row0: model.set_row_unit(r, u))
        view.setIndexWidget(model.index(row0, model.unit_col), combo)
