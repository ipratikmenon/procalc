"""Reconstructed "Streams" table: every stream in the loaded HMB, laid out
with streams as columns and properties as rows — the same property
vocabulary/order Hydraulics/hmb_proii_reader.py's own "Case N" transposed
PRO/II report uses (its _SCALAR_MAP, PRO/II's native col-A label list),
just transposed back the other way: PRO/II's report already has properties
as rows and streams as columns per case sheet — this table reproduces that
same row list/order for every source kind (Case-N, per-stream dump, or
live), not just Case-N inputs.

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
from PySide6.QtGui import QFontMetrics
from PySide6.QtWidgets import (QComboBox, QDialog, QHBoxLayout, QLabel,
                               QLineEdit, QTableWidget, QTableWidgetItem,
                               QVBoxLayout, QWidget)

import engine_api as api

# (row label, StreamProps field, quantity code | None, decimals)
# Order/labels mirror Hydraulics/hmb_proii_reader.py's _SCALAR_MAP (PRO/II's
# own "Case N" transposed-report vocabulary) and _KEY_TO_QTY (quantity-code
# assignments for the unit dropdown) — the PRO/II-native reference the
# Streams table's property list/order is built from. Fields with no entry
# in _KEY_TO_QTY have no modeled unit in common/units.py and render as
# plain unitless numbers (qty=None), matching "Phase"/"Vapor Z Factor"'s
# existing qty=None convention.
#
# Two _SCALAR_MAP entries don't survive to StreamProps and are adapted
# rather than shown as permanently-blank rows: "Total Molar Rate"
# (total_molar) is parsed but dropped at the HMBStream->StreamProps
# boundary (streamprops_from_hmb, hydraulics_XOM.py) with no field to
# carry it, so it's omitted here; "Liquid Weight Fraction" (liq_wt_frac)
# is consumed there to derive the real per-phase mass split, so it's shown
# as that split (Vapor/Liquid Mass Flow) instead of the fraction itself.
_ROWS = [
    ("Phase", "phase", None, None),
    ("Total Mass Rate", "total_mass", "mflow", 1),
    ("Vapor Mass Flow", "vap_mass", "mflow", 1),
    ("Liquid Mass Flow", "liq_mass", "mflow", 1),
    ("Total Std. Liq. Vol Rate", "total_std_liq", "qvol", 1),
    ("Total Std. Vap. Vol Rate", "total_std_vap", "qvol", 1),
    ("Temperature", "temp_f", "T", 2),
    ("Pressure", "pres_psia", "P", 3),
    ("Total Molecular Weight", "mol_weight", "MW", 3),
    ("Total Actual Density", "total_density", "rho", 4),
    ("Total Z (from actual density)", "total_z", None, 4),
    ("True Critical Temperature", "tc_f", "T", 2),
    ("True Critical Pressure", "pc_psia", "P", 3),
    ("Dry Vapor Molecular Weight", "vap_mw", "MW", 3),
    ("Dry Vapor Act. Density", "vap_density", "rho", 4),
    ("Dry Vapor Cp/Cv Ratio", "vap_cp_cv", None, 4),
    ("Dry Vapor Cp", "vap_cp", None, 4),
    ("Dry Vapor Viscosity", "vap_visc", "visc", 4),
    ("Dry Vapor Z (from actual)", "vap_z", None, 4),
    ("Dry Vapor Thermal Conductivity", "vap_therm_cond", None, 5),
    ("Dry Vapor Sp. Enthalpy", "vap_sp_enthalpy", "h", 2),
    ("Dry Liquid Molecular Weight", "liq_mw", "MW", 3),
    ("Dry Liquid Act. Density", "liq_density", "rho", 4),
    ("Dry Liquid Cp", "liq_cp", None, 4),
    ("Dry Liquid Viscosity", "liq_visc", "visc", 4),
    ("Dry Liquid Thermal Conductivity", "liq_therm_cond", None, 5),
    ("Dry Liquid Sp. Enthalpy", "liq_sp_enthalpy", "h", 2),
    ("Surface Tension", "liq_surf_tens", "st", 3),
    ("Acentric Factor", "acentric", None, 4),
]


# Fixed caption AND combo widths so every row header's unit combo starts
# (and ends) at the same x offset, column-0-wide (long captions like "Dry
# Vapor Thermal Conductivity" are elided with the full text kept as a
# tooltip). Fixing the combo's width too, not just the caption's, avoids a
# subtler alignment break: with only the caption pinned, a row whose combo
# has a wider natural sizeHint (e.g. Pressure's longer unit strings like
# "kg/cm2g") can still pull the whole row's layout into a tighter squeeze
# that shifts its own start position left of the fixed-width rows.
_ROW_HEADER_CAP_W = 132
_ROW_HEADER_COMBO_W = 78
_ROW_HEADER_COL_W = 240


class _RowHeader(QWidget):
    """Row-header cell: a caption + (for numeric rows) a clickable unit
    dropdown, styled the same as UnitValueLabel's #UnitPicker combo."""

    def __init__(self, caption, qty, unit, on_unit_changed, parent=None):
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(4, 0, 4, 0)
        lay.setSpacing(4)
        cap_lbl = QLabel()
        cap_lbl.setFixedWidth(_ROW_HEADER_CAP_W)
        # ensurePolished() forces the app stylesheet's font (e.g. Inter,
        # wider than the pre-style default) to apply before measuring --
        # eliding against the pre-style font underestimates the real
        # rendered width and the caption clips without an ellipsis.
        cap_lbl.ensurePolished()
        fm = QFontMetrics(cap_lbl.font())
        cap_lbl.setText(fm.elidedText(caption, Qt.ElideRight, _ROW_HEADER_CAP_W))
        cap_lbl.setToolTip(caption)
        lay.addWidget(cap_lbl)
        self.combo = None
        if qty:
            self.combo = QComboBox()
            self.combo.setObjectName("UnitPicker")
            self.combo.setFixedWidth(_ROW_HEADER_COMBO_W)
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
        self.table.setColumnWidth(0, _ROW_HEADER_COL_W)
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
