"""A SheetTableModel that additionally makes each qty-bearing row's Unit
cell a clickable, live-converting unit dropdown -- layered directly on top
of the real HMB sheet, not a hand-built reconstruction.

Only rows whose Property label maps to a real quantity code
(engine_api.property_quantity_for) AND whose own printed Unit cell
normalizes to a unit engine_api actually knows (engine_api.
normalize_hmb_unit) get a dropdown; every other cell renders byte-
identical to the raw file. Conversion is a pure per-row display overlay --
the underlying openpyxl worksheet (the source of truth) is never mutated.

Two on-disk shapes are recognized, tried in order, so this covers both
HMB export formats without any sheet-name special-casing:
  1. Per-stream / OUTPUT shape: a "Property"/"Unit" header at row 2,
     columns B/C, value columns starting at D. True for every per-stream
     sheet and OUTPUT; false for UNITS (header but no value columns),
     COMPONENTS/COMP_CONSTANTS, and INPUT (no such header at all).
  2. PRO/II "Case N" transposed-export shape (hmb_proii_reader.py's own
     format): property label in column A, unit in column B, one stream
     per column from C on, with a "Stream Name" row (not a fixed row
     number) instead of a "Property"/"Unit" header.
A sheet matching neither shape (or a synthetic live-connector sheet built
elsewhere, which always uses shape 1) has no interactive rows and renders
exactly like a plain SheetTableModel.
"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QComboBox

import engine_api as api
from results.workbook_model import SheetTableModel, _fmt

# shape 1: per-stream / OUTPUT
_S1_PROPERTY_COL, _S1_UNIT_COL, _S1_FIRST_VALUE_COL = 2, 3, 4   # cols B, C, D
_S1_HEADER_ROW, _S1_DATA_START_ROW = 2, 3

# shape 2: PRO/II Case-N transposed export
_S2_PROPERTY_COL, _S2_UNIT_COL, _S2_FIRST_VALUE_COL = 1, 2, 3   # cols A, B, C
_S2_STREAM_NAME_SCAN_ROWS = 10   # "Stream Name" row is near the top, not fixed


class UnitAwareSheetModel(SheetTableModel):
    def __init__(self, ws, parent=None):
        super().__init__(ws, parent)
        self._native: dict[int, tuple[str, str]] = {}   # row0 -> (qty, native_unit)
        self._chosen: dict[int, str] = {}                # row0 -> currently-picked unit
        self._property_col = self._unit_col_1 = self._first_value_col = None
        data_start_row = self._detect_shape(ws)
        if data_start_row is None:
            self.unit_col = _S1_UNIT_COL - 1   # arbitrary but harmless; no rows use it
            return
        self.unit_col = self._unit_col_1 - 1   # 0-indexed, for the view layer
        for r in range(data_start_row, self.nrows + 1):
            if (r, self._property_col) in self.covered:
                continue
            label = ws.cell(row=r, column=self._property_col).value
            if not label:
                continue
            qty = api.property_quantity_for(str(label))
            if not qty:
                continue
            raw_unit = ws.cell(row=r, column=self._unit_col_1).value
            native = api.normalize_hmb_unit(qty, raw_unit) if raw_unit else None
            if not native:
                continue
            row0 = r - 1
            self._native[row0] = (qty, native)
            self._chosen[row0] = native

    def _detect_shape(self, ws) -> int | None:
        """Returns the first data row to scan, or None if neither known
        shape matches (sheet stays plain/non-interactive)."""
        if self.ncols >= _S1_FIRST_VALUE_COL:
            header_prop = ws.cell(row=_S1_HEADER_ROW, column=_S1_PROPERTY_COL).value
            header_unit = ws.cell(row=_S1_HEADER_ROW, column=_S1_UNIT_COL).value
            if (str(header_prop or "").strip().lower() == "property"
                    and str(header_unit or "").strip().lower() == "unit"):
                self._property_col = _S1_PROPERTY_COL
                self._unit_col_1 = _S1_UNIT_COL
                self._first_value_col = _S1_FIRST_VALUE_COL
                return _S1_DATA_START_ROW
        if self.ncols >= _S2_FIRST_VALUE_COL:
            for r in range(1, min(_S2_STREAM_NAME_SCAN_ROWS, self.nrows) + 1):
                v = ws.cell(row=r, column=_S2_PROPERTY_COL).value
                if str(v or "").strip().lower() == "stream name":
                    self._property_col = _S2_PROPERTY_COL
                    self._unit_col_1 = _S2_UNIT_COL
                    self._first_value_col = _S2_FIRST_VALUE_COL
                    return r + 1
        return None

    def unit_rows(self) -> dict[int, tuple[str, str]]:
        """{0-indexed row: (qty, native_unit)} for every row with a working
        clickable unit dropdown."""
        return dict(self._native)

    def set_row_unit(self, row0: int, new_unit: str):
        entry = self._native.get(row0)
        if not entry or new_unit == self._chosen.get(row0):
            return
        self._chosen[row0] = new_unit
        first = self.index(row0, self._first_value_col - 1)
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
            if role == Qt.DisplayRole and c0 >= self._first_value_col - 1:
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
