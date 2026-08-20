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


def _decimals(nf):
    """Count the '0'/'#' placeholders after the decimal point in a format."""
    if "." not in nf:
        return 0
    frac = nf.split(".", 1)[1]
    n = 0
    for ch in frac:
        if ch in "0#":
            n += 1
        elif ch in "%_)eE ":
            continue
        else:
            break
    return min(max(n, 0), 6)


def _fmt(value, number_format):
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        nf = (number_format or "General").strip()
        try:
            if nf in ("General", "@", ""):
                if isinstance(value, int) or float(value).is_integer():
                    return str(int(value))
                return f"{value:.6g}"
            grp = "," in nf                       # thousands separator, e.g. #,##0
            if nf.endswith("%") or "0.0%" in nf or "0%" in nf:
                dec = _decimals(nf)
                return f"{value * 100:,.{dec}f}%" if grp else f"{value * 100:.{dec}f}%"
            dec = _decimals(nf)                   # covers 0.0 / 0.00 / 0.000 / #,##0.00
            return f"{value:,.{dec}f}" if grp else f"{value:.{dec}f}"
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
