"""Reconstructed "Streams" table: every stream in the loaded HMB, laid out
with streams as columns and properties as rows — the same row-group order
Hydraulics/HMB.xlsx's own per-stream sheets use (Conditions, Flow Rates,
Vapor/Liquid Phase Properties, Critical Properties, Composition), just
transposed into one wide table instead of one sheet per stream.

Each numeric property row gets ONE clickable unit dropdown in its row
header (not one per cell) — changing it re-renders every stream's value in
that row at once, via engine_api.convert(). A per-cell UnitValueLabel per
row would mean thousands of composite widgets for a large HMB file; a
single dropdown per row achieves the same "click the unit, see it convert"
requirement far more cheaply and matches the sketch (the unit shown once,
on the left, not repeated per stream).
"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QComboBox, QDialog, QHBoxLayout, QLabel,
                               QLineEdit, QTableWidget, QTableWidgetItem,
                               QVBoxLayout, QWidget)

import engine_api as api

# (row label, StreamProps field, quantity code | None, decimals)
_ROWS = [
    ("Phase", "phase", None, None),
    ("Temperature", "temp_f", "T", 2),
    ("Pressure", "pres_psia", "P", 3),
    ("Total Mass Flow", "total_mass", "mflow", 1),
    ("Vapor Mass Flow", "vap_mass", "mflow", 1),
    ("Liquid Mass Flow", "liq_mass", "mflow", 1),
    ("Molecular Weight", "mol_weight", "MW", 3),
    ("Vapor Density", "vap_density", "rho", 4),
    ("Liquid Density", "liq_density", "rho", 4),
    ("Vapor Viscosity", "vap_visc", "visc", 4),
    ("Liquid Viscosity", "liq_visc", "visc", 4),
    ("Vapor Z Factor", "vap_z", None, 4),
    ("Critical Temperature", "tc_f", "T", 2),
    ("Critical Pressure", "pc_psia", "P", 3),
]


class _RowHeader(QWidget):
    """Row-header cell: a caption + (for numeric rows) a clickable unit
    dropdown, styled the same as UnitValueLabel's #UnitPicker combo."""

    def __init__(self, caption, qty, unit, on_unit_changed, parent=None):
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(4, 0, 4, 0)
        lay.setSpacing(4)
        lay.addWidget(QLabel(caption))
        self.combo = None
        if qty:
            self.combo = QComboBox()
            self.combo.setObjectName("UnitPicker")
            self.combo.addItems(api.units_for(qty))
            i = self.combo.findText(unit)
            if i >= 0:
                self.combo.setCurrentIndex(i)
            self.combo.currentTextChanged.connect(on_unit_changed)
            lay.addWidget(self.combo)
        lay.addStretch(1)


class StreamsTableWidget(QWidget):
    def __init__(self, hmb_path: str, case: str = "Case 1", parent=None):
        super().__init__(parent)
        self._hmb_path = hmb_path
        self._case = case
        self._row_units: dict[int, str] = {}   # row index -> current unit
        self._stream_names: list[str] = []
        # {row_idx: {stream_name: raw_internal_value}} for numeric rows only
        self._raw: dict[int, dict[str, float]] = {}
        self._comp_rows_start = None
        self._components: list[str] = []

        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 10, 10, 10)
        outer.setSpacing(8)

        top = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Filter streams by name…")
        self.search.textChanged.connect(self._apply_filter)
        self.status_lbl = QLabel("")
        self.status_lbl.setObjectName("Muted")
        top.addWidget(self.search, 1)
        top.addWidget(self.status_lbl)
        outer.addLayout(top)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        outer.addWidget(self.table, 1)

        self.reload()

    def reload(self):
        self.status_lbl.setText("Loading…")
        data = api.list_all_stream_data(self._hmb_path, self._case)
        self._stream_names = list(data.keys())

        # union of components across every stream that has composition,
        # first-appearance order (mirrors the dict.fromkeys idiom used
        # elsewhere in this codebase for stable ordered de-dup)
        comp_names: dict[str, None] = {}
        for d in data.values():
            comp = d.get("composition")
            if comp:
                for name in comp:
                    comp_names.setdefault(name, None)
        self._components = list(comp_names.keys())
        self._comp_rows_start = len(_ROWS)

        n_streams = len(self._stream_names)
        # column 0 is the row-header (label + unit-picker) column; stream
        # data occupies columns 1..n_streams
        self.table.clear()
        self.table.setRowCount(len(_ROWS) + len(self._components))
        self.table.setColumnCount(n_streams + 1)
        self.table.setHorizontalHeaderItem(0, QTableWidgetItem(""))
        for c, name in enumerate(self._stream_names, start=1):
            self.table.setHorizontalHeaderItem(c, QTableWidgetItem(name))
        self.table.setColumnWidth(0, 190)
        for c in range(1, n_streams + 1):
            self.table.setColumnWidth(c, 100)
        self.table.verticalHeader().setVisible(False)

        self._raw = {}
        for r, (caption, field, qty, decimals) in enumerate(_ROWS):
            unit = api.internal_unit(qty) if qty else None
            self._row_units[r] = unit
            self.table.setCellWidget(
                r, 0, _RowHeader(caption, qty, unit,
                                 lambda u, row=r: self._on_row_unit_changed(row, u)))
            if qty:
                self._raw[r] = {}
            for c, name in enumerate(self._stream_names, start=1):
                sp = data[name].get("props")
                val = getattr(sp, field, None) if sp is not None else None
                if qty and val is not None:
                    self._raw[r][name] = val
                item = QTableWidgetItem(self._fmt(val, decimals, qty, unit))
                item.setTextAlignment(int(Qt.AlignRight | Qt.AlignVCenter)
                                      if qty else int(Qt.AlignLeft | Qt.AlignVCenter))
                self.table.setItem(r, c, item)

        for i, comp_name in enumerate(self._components):
            r = self._comp_rows_start + i
            self.table.setCellWidget(r, 0, _RowHeader(comp_name, None, None, None))
            for c, name in enumerate(self._stream_names, start=1):
                comp = data[name].get("composition") or {}
                frac = comp.get(comp_name)
                item = QTableWidgetItem("—" if frac is None else f"{frac:.4f}")
                item.setTextAlignment(int(Qt.AlignRight | Qt.AlignVCenter))
                self.table.setItem(r, c, item)

        self.status_lbl.setText(f"{n_streams} streams · {len(self._components)} components")

    def _fmt(self, value, decimals, qty, unit):
        if value is None:
            return "—"
        if not qty:
            return str(value)
        try:
            return f"{float(value):,.{decimals}f}"
        except (TypeError, ValueError):
            return str(value)

    def _on_row_unit_changed(self, row, new_unit):
        old_unit = self._row_units.get(row)
        if not new_unit or new_unit == old_unit:
            return
        self._row_units[row] = new_unit
        qty = _ROWS[row][2]
        decimals = _ROWS[row][3]
        raw = self._raw.get(row, {})
        for c, name in enumerate(self._stream_names, start=1):
            # raw[] stores the value as last displayed (already in old_unit)
            current = raw.get(name)
            if current is None:
                continue
            converted = api.convert(qty, current, old_unit, new_unit)
            item = self.table.item(row, c)
            if item is not None:
                item.setText(self._fmt(converted, decimals, qty, new_unit))
            raw[name] = converted
        self._raw[row] = raw

    def _apply_filter(self, term):
        term = (term or "").strip().lower()
        for c, name in enumerate(self._stream_names, start=1):
            self.table.setColumnHidden(c, bool(term) and term not in name.lower())


class StreamsDialog(QDialog):
    def __init__(self, hmb_path: str, case: str = "Case 1", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Streams")
        self.resize(1200, 700)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(StreamsTableWidget(hmb_path, case, self))
