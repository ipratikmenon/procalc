"""Streams / HMB viewer: a faithful, sheet-by-sheet rendering of the loaded
HMB workbook -- exactly as the source Excel file shows it, not a hand-built
reconstruction -- with a searchable sheet-name list standing in for tabs
(a real HMB workbook can have 400+ sheets, far too many for a literal
scrolling tab bar to stay usable), and a clickable unit dropdown layered
directly onto any sheet shaped like a real Property/Unit table (every
per-stream sheet, and the workbook's own OUTPUT summary sheet) via
streams.unit_sheet_model.UnitAwareSheetModel.

Every sheet renders through the same results/workbook_model.build_sheet_view
the Results tab uses for engine output -- so a per-stream sheet's merged
title row, column widths, fills and fonts come out identical to opening the
file in Excel. A full-fidelity load of a large HMB workbook takes real time
(openpyxl's non-read_only parser materializes every sheet up front; there is
no partial/lazy-but-full-fidelity mode in its public API), so it runs on a
background QThread once per (path, mtime) and is cached for the app session.
"""
from __future__ import annotations

import os

from openpyxl import load_workbook
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (QDialog, QHBoxLayout, QLabel, QLineEdit,
                               QListWidget, QListWidgetItem, QSplitter,
                               QVBoxLayout, QWidget)

from results.workbook_model import build_sheet_view
from streams.unit_sheet_model import UnitAwareSheetModel, install_unit_dropdowns

# Non-per-stream sheets, pinned as a small group at the top of the sheet
# list ahead of the (potentially hundreds of) per-stream sheets below.
_META_SHEETS = ("UNITS", "COMPONENTS", "COMP_CONSTANTS", "INPUT", "OUTPUT")


class _WorkbookLoadWorker(QThread):
    finished_ok = Signal(object)   # openpyxl.Workbook
    failed = Signal(str)

    def __init__(self, path, parent=None):
        super().__init__(parent)
        self._path = path

    def run(self):
        try:
            wb = load_workbook(self._path, read_only=False, data_only=True)
        except Exception as e:  # noqa: BLE001
            self.failed.emit(f"{type(e).__name__}: {e}")
            return
        self.finished_ok.emit(wb)


class StreamsDialog(QDialog):
    def __init__(self, hmb_path: str, case: str = "Case 1", parent=None,
                workbook_cache: dict | None = None):
        super().__init__(parent)
        self.setWindowTitle("Streams — HMB Workbook")
        self.resize(1300, 800)
        self._hmb_path = hmb_path
        # shared with the caller (typically MainWindow) so re-opening this
        # dialog for the same file/session is instant instead of re-paying
        # the full-workbook parse; a fresh dict if the caller doesn't pass one.
        self._workbook_cache = workbook_cache if workbook_cache is not None else {}
        self._wb = None
        self._current_view = None
        self._worker = None
        self._closed = False

        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 10, 10, 10)
        lay.setSpacing(8)

        self.status_lbl = QLabel("Loading HMB workbook…")
        self.status_lbl.setObjectName("Muted")
        lay.addWidget(self.status_lbl)

        split = QSplitter(Qt.Horizontal)
        left = QWidget()
        llay = QVBoxLayout(left)
        llay.setContentsMargins(0, 0, 0, 0)
        llay.setSpacing(6)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Filter sheets…")
        self.search.textChanged.connect(self._apply_filter)
        llay.addWidget(self.search)
        self.sheet_list = QListWidget()
        self.sheet_list.currentItemChanged.connect(self._on_sheet_selected)
        llay.addWidget(self.sheet_list, 1)
        left.setMinimumWidth(180)
        left.setMaximumWidth(280)
        split.addWidget(left)

        content_host = QWidget()
        self._content_slot = QVBoxLayout(content_host)
        self._content_slot.setContentsMargins(0, 0, 0, 0)
        split.addWidget(content_host)
        split.setStretchFactor(0, 0)
        split.setStretchFactor(1, 1)
        lay.addWidget(split, 1)

        self._load()

    # ── loading ──────────────────────────────────────────────────────────
    def _cache_key(self):
        try:
            return (self._hmb_path, os.path.getmtime(self._hmb_path))
        except OSError:
            return None

    def _load(self):
        key = self._cache_key()
        cached = self._workbook_cache.get(key) if key else None
        if cached is not None:
            self._on_loaded(cached)
            return
        self._worker = _WorkbookLoadWorker(self._hmb_path, self)
        self._worker.finished_ok.connect(self._on_loaded)
        self._worker.failed.connect(self._on_load_failed)
        self._worker.start()

    def _on_loaded(self, wb):
        if self._closed:
            return
        self._wb = wb
        key = self._cache_key()
        if key:
            self._workbook_cache[key] = wb
        self.status_lbl.setText(
            f"{len(wb.sheetnames)} sheets · {os.path.basename(self._hmb_path)}")
        self._populate_list()

    def _on_load_failed(self, msg):
        if self._closed:
            return
        self.status_lbl.setText(f"Could not load HMB workbook: {msg}")

    def closeEvent(self, event):
        self._closed = True
        super().closeEvent(event)

    # ── sheet list ───────────────────────────────────────────────────────
    def _populate_list(self):
        self.sheet_list.clear()
        names = self._wb.sheetnames
        meta = [n for n in names if n in _META_SHEETS]
        streams = [n for n in names if n not in _META_SHEETS]
        for n in meta:
            self.sheet_list.addItem(QListWidgetItem(n))
        if meta and streams:
            sep = QListWidgetItem(f"── {len(streams)} streams ──")
            sep.setFlags(Qt.NoItemFlags)
            self.sheet_list.addItem(sep)
        for n in streams:
            self.sheet_list.addItem(QListWidgetItem(n))

        default_row = 0
        for i in range(self.sheet_list.count()):
            if self.sheet_list.item(i).text() == "OUTPUT":
                default_row = i
                break
        if self.sheet_list.count():
            self.sheet_list.setCurrentRow(default_row)

    def _apply_filter(self, term):
        term = (term or "").strip().lower()
        for i in range(self.sheet_list.count()):
            item = self.sheet_list.item(i)
            item.setHidden(bool(term) and term not in item.text().lower())

    def _on_sheet_selected(self, current, _previous):
        if current is None or self._wb is None:
            return
        name = current.text()
        if name not in self._wb.sheetnames:
            return
        self._show_sheet(name)

    # ── sheet content ────────────────────────────────────────────────────
    def _show_sheet(self, name):
        if self._current_view is not None:
            self._content_slot.removeWidget(self._current_view)
            self._current_view.setParent(None)
            self._current_view.deleteLater()
            self._current_view = None
        ws = self._wb[name]
        view = build_sheet_view(ws, model_cls=UnitAwareSheetModel)
        install_unit_dropdowns(view, view.model())
        self._content_slot.addWidget(view)
        self._current_view = view
