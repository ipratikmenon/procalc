"""Results panel: a circuit selector + a tab per output sheet, each rendered
from the engine-produced .xlsx."""
from __future__ import annotations

import os

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QComboBox,
                               QLabel, QTabWidget, QTableView, QHeaderView,
                               QAbstractItemView)

from results.workbook_model import SheetTableModel


class _SheetTab(QTableView):
    def __init__(self, ws, parent=None):
        super().__init__(parent)
        self.setModel(SheetTableModel(ws))
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.horizontalHeader().setVisible(False)
        self.verticalHeader().setVisible(False)
        self.setShowGrid(True)
        self.setAlternatingRowColors(False)
        m = self.model()
        for (ar, ac), (rs, cs) in m.spans.items():
            self.setSpan(ar - 1, ac - 1, rs, cs)
        # approximate column widths from the sheet
        for c in range(m.ncols):
            letter = get_column_letter(c + 1)
            w = ws.column_dimensions[letter].width if letter in ws.column_dimensions else None
            self.setColumnWidth(c, int((w or 10) * 7) + 6)


class ResultsView(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._paths: list[str] = []
        lay = QVBoxLayout(self)
        lay.setContentsMargins(6, 6, 6, 6)
        top = QHBoxLayout()
        top.addWidget(QLabel("Circuit:"))
        self.circuit_cb = QComboBox()
        self.circuit_cb.currentIndexChanged.connect(self._show_current)
        top.addWidget(self.circuit_cb)
        top.addStretch(1)
        self.info = QLabel("")
        self.info.setStyleSheet("color:#555;")
        top.addWidget(self.info)
        lay.addLayout(top)
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        lay.addWidget(self.tabs, 1)
        self._placeholder()

    def _placeholder(self):
        self.tabs.clear()
        w = QLabel("Run a calculation to see results here.\n"
                   "The output sheets (Cover, Line List, Pressure Profile, "
                   "Flash, CV datasheets, …) appear as tabs.")
        w.setAlignment(Qt.AlignCenter)
        w.setStyleSheet("color:#7a8a99; font-size:13px;")
        self.tabs.addTab(w, "Results")

    def load(self, paths: list[str]):
        self._paths = [p for p in (paths or []) if p and os.path.exists(p)]
        self.circuit_cb.blockSignals(True)
        self.circuit_cb.clear()
        for p in self._paths:
            name = os.path.basename(p).replace("hydraulics_flash_noiso_", "").replace(".xlsx", "")
            self.circuit_cb.addItem(name, p)
        self.circuit_cb.blockSignals(False)
        if self._paths:
            self._show_current()
        else:
            self._placeholder()

    def _show_current(self):
        idx = self.circuit_cb.currentIndex()
        if idx < 0 or idx >= len(self._paths):
            return
        prev = self.tabs.currentIndex()
        prev_name = self.tabs.tabText(prev) if prev >= 0 else None
        path = self.circuit_cb.itemData(idx)
        self.tabs.clear()
        wb = load_workbook(path, data_only=False)
        for ws in wb.worksheets:
            self.tabs.addTab(_SheetTab(ws), ws.title)
        self.info.setText(f"{self.tabs.count()} sheets  ·  {os.path.basename(path)}")
        # restore the previously-open tab by name if present
        if prev_name:
            for i in range(self.tabs.count()):
                if self.tabs.tabText(i) == prev_name:
                    self.tabs.setCurrentIndex(i)
                    break
