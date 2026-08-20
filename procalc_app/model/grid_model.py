"""The circuit grid: a QAbstractTableModel over the engine's INPUT_HEADERS.

Rows are plain dicts keyed by the canonical header names (values in internal
FPS).  Editing any cell emits ``gridEdited`` so the run controller can debounce
an autocalc.
"""
from __future__ import annotations

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, Signal
from PySide6.QtGui import QColor

import engine_api as api

# columns worth showing first / grouping colour by circuit
_CIRCUIT_TINTS = ["#EAF3FE", "#EAF7F0", "#FDF6E3", "#F3EEFA", "#FDECEC"]


class CircuitGridModel(QAbstractTableModel):
    gridEdited = Signal()

    def __init__(self, rows=None, unit_system="FPS", parent=None):
        super().__init__(parent)
        self.headers = api.input_headers()
        self.units = unit_system
        self.numeric = api.numeric_columns()
        self.rows: list[dict] = rows or []

    # ── shape ──
    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.headers)

    def headerData(self, section, orient, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orient == Qt.Horizontal:
            return api.display_header(self.headers[section], self.units)
        return str(section + 1)

    # ── data ──
    def _circuit_tint(self, row_idx) -> QColor | None:
        cid = self.rows[row_idx].get("Circuit")
        if not cid:
            return None
        order = list(dict.fromkeys(r.get("Circuit") for r in self.rows if r.get("Circuit")))
        try:
            i = order.index(cid)
        except ValueError:
            return None
        return QColor(_CIRCUIT_TINTS[i % len(_CIRCUIT_TINTS)])

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        col = self.headers[index.column()]
        row = self.rows[index.row()]
        v = row.get(col)
        if role in (Qt.DisplayRole, Qt.EditRole):
            return "" if v is None else str(v)
        if role == Qt.BackgroundRole:
            return self._circuit_tint(index.row())
        if role == Qt.TextAlignmentRole:
            if col in self.numeric:
                return int(Qt.AlignRight | Qt.AlignVCenter)
            return int(Qt.AlignLeft | Qt.AlignVCenter)
        return None

    def flags(self, index):
        if not index.isValid():
            return Qt.NoItemFlags
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable

    def setData(self, index, value, role=Qt.EditRole):
        if role != Qt.EditRole or not index.isValid():
            return False
        col = self.headers[index.column()]
        s = "" if value is None else str(value).strip()
        if s == "":
            self.rows[index.row()][col] = None
        elif col in self.numeric:
            try:
                self.rows[index.row()][col] = float(s) if ("." in s or "e" in s.lower()) else int(s)
            except ValueError:
                self.rows[index.row()][col] = s        # keep text if not numeric
        else:
            self.rows[index.row()][col] = s
        self.dataChanged.emit(index, index, [Qt.DisplayRole])
        self.gridEdited.emit()
        return True

    # ── row ops ──
    def add_row(self, at=None, template=None):
        at = len(self.rows) if at is None else at
        self.beginInsertRows(QModelIndex(), at, at)
        self.rows.insert(at, dict(template or {}))
        self.endInsertRows()
        self.gridEdited.emit()

    def duplicate_row(self, at):
        if 0 <= at < len(self.rows):
            self.add_row(at + 1, dict(self.rows[at]))

    def delete_row(self, at):
        if 0 <= at < len(self.rows):
            self.beginRemoveRows(QModelIndex(), at, at)
            self.rows.pop(at)
            self.endRemoveRows()
            self.gridEdited.emit()

    def set_rows(self, rows):
        self.beginResetModel()
        self.rows = rows or []
        self.endResetModel()
        self.gridEdited.emit()

    def to_rows(self) -> list[dict]:
        return [dict(r) for r in self.rows]
