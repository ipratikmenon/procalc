"""Render an openpyxl worksheet in a QTableView as close to Excel as possible:
cell values (number-formatted), fills, bold/colour fonts, alignment, merges,
and approximate column widths."""
from __future__ import annotations

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PySide6.QtGui import QColor, QFont


def _argb_to_qcolor(rgb):
    if not rgb or not isinstance(rgb, str):
        return None
    s = rgb[-6:] if len(rgb) >= 6 else rgb
    try:
        return QColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
    except ValueError:
        return None


def _fmt(value, number_format):
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        nf = number_format or "General"
        try:
            if nf.endswith("%"):
                dec = nf.count("0", nf.find(".")) if "." in nf else 0
                return f"{value * 100:.{dec}f}%"
            if "." in nf:
                dec = len(nf.split(".")[-1].replace("_", "").replace(")", "").replace("%", ""))
                dec = min(max(dec, 0), 6)
                return f"{value:,.{dec}f}" if "," in nf else f"{value:.{dec}f}"
            if nf in ("#,##0", "0"):
                return f"{value:,.0f}" if "," in nf else f"{value:.0f}"
        except Exception:
            pass
        if isinstance(value, float):
            return f"{value:.4g}"
    return str(value)


class SheetTableModel(QAbstractTableModel):
    def __init__(self, ws, parent=None):
        super().__init__(parent)
        self.ws = ws
        self.nrows = ws.max_row
        self.ncols = ws.max_column
        # merged ranges: anchor -> (rows, cols); covered cells -> anchor
        self.spans = {}
        self.covered = set()
        for rng in ws.merged_cells.ranges:
            self.spans[(rng.min_row, rng.min_col)] = (
                rng.max_row - rng.min_row + 1, rng.max_col - rng.min_col + 1)
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    if (r, c) != (rng.min_row, rng.min_col):
                        self.covered.add((r, c))

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else self.nrows

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else self.ncols

    def _cell(self, index):
        return self.ws.cell(row=index.row() + 1, column=index.column() + 1)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        r, c = index.row() + 1, index.column() + 1
        if (r, c) in self.covered:
            return "" if role == Qt.DisplayRole else None
        cell = self.ws.cell(row=r, column=c)
        if role == Qt.DisplayRole:
            return _fmt(cell.value, cell.number_format)
        if role == Qt.BackgroundRole:
            fill = cell.fill
            if fill is not None and fill.fgColor is not None and fill.patternType:
                return _argb_to_qcolor(getattr(fill.fgColor, "rgb", None))
            return None
        if role == Qt.ForegroundRole:
            f = cell.font
            if f is not None and f.color is not None:
                return _argb_to_qcolor(getattr(f.color, "rgb", None))
            return None
        if role == Qt.FontRole:
            f = cell.font
            if f is None:
                return None
            qf = QFont()
            if f.bold:
                qf.setBold(True)
            if f.italic:
                qf.setItalic(True)
            if f.size:
                qf.setPointSizeF(float(f.size))
            return qf
        if role == Qt.TextAlignmentRole:
            a = cell.alignment
            h = {"center": Qt.AlignHCenter, "right": Qt.AlignRight,
                 "left": Qt.AlignLeft}.get(getattr(a, "horizontal", None) or "", Qt.AlignLeft)
            return int(h | Qt.AlignVCenter)
        return None

    def headerData(self, section, orient, role=Qt.DisplayRole):
        return None      # sheets carry their own header rows
