"""Procalc Hydraulics — main window (P1: grid + HMB loader + Run + results)."""
from __future__ import annotations

import os

from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QAction, QPixmap, QIcon
from PySide6.QtWidgets import (QMainWindow, QWidget, QToolBar, QSplitter,
                               QTableView, QTabWidget, QPlainTextEdit, QLabel,
                               QFileDialog, QComboBox, QLineEdit, QFormLayout,
                               QGroupBox, QVBoxLayout, QHBoxLayout, QMessageBox,
                               QAbstractItemView, QStatusBar, QWidgetAction)

import engine_api as api
from model.grid_model import CircuitGridModel
from model.delegates import install_delegates
from run.controller import RunController
from results.results_view import ResultsView
from resources import theme

_HERE = os.path.dirname(os.path.abspath(__file__))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Procalc Hydraulics — T.EN")
        self.resize(1360, 820)
        self.hmb_path = None
        self.hmb_case = "Case 1"

        self.model = CircuitGridModel(api.new_project_rows())
        self.model.gridEdited.connect(self._on_grid_edited)

        self.controller = RunController(self._rows, self._context, self)
        self.controller.log.connect(self._log)
        self.controller.started.connect(lambda: self._set_running(True))
        self.controller.finished.connect(self._on_finished)
        self.controller.error.connect(self._on_error)
        self.controller.state.connect(self._on_state)

        self._build_toolbar()
        self._build_central()
        self._build_statusbar()
        self._refresh_hmb_label()

    # ── UI construction ──
    def _build_toolbar(self):
        tb = QToolBar("Main")
        tb.setIconSize(QSize(18, 18))
        tb.setMovable(False)
        self.addToolBar(tb)

        def act(text, slot):
            a = QAction(text, self)
            a.triggered.connect(slot)
            tb.addAction(a)
            return a

        act("＋ Row", lambda: self.model.add_row(template={"Circuit": "C1", "Run Type": "Main"}))
        act("⧉ Duplicate", self._dup_row)
        act("🗑 Delete", self._del_row)
        tb.addSeparator()
        act("📂 Load HMB", self._load_hmb)
        self.run_act = act("▶ Run", self.controller.run_now)
        self.auto_act = QAction("Auto", self, checkable=True)
        self.auto_act.toggled.connect(self._toggle_auto)
        tb.addAction(self.auto_act)
        tb.addSeparator()
        act("PMS…", self._open_pms)
        act("Export…", self._export)

    def _build_central(self):
        split = QSplitter(Qt.Horizontal)

        # left: project meta + grid
        left = QWidget()
        lv = QVBoxLayout(left)
        lv.setContentsMargins(6, 6, 6, 6)
        meta_box = QGroupBox("Project")
        form = QFormLayout(meta_box)
        self.ed_project = QLineEdit("BTGOLD HCU")
        self.ed_area = QLineEdit("Unit 24")
        self.ed_unit = QLineEdit("HCU")
        self.cb_case = QComboBox(); self.cb_case.addItem("Case 1")
        self.cb_mode = QComboBox(); self.cb_mode.addItems(["isothermal", "isenthalpic"])
        row = QHBoxLayout()
        for w in (QLabel("Case"), self.cb_case, QLabel("Flash"), self.cb_mode):
            row.addWidget(w)
        rw = QWidget(); rw.setLayout(row)
        form.addRow("Project", self.ed_project)
        form.addRow("Area", self.ed_area)
        form.addRow("Unit", self.ed_unit)
        form.addRow(rw)
        lv.addWidget(meta_box)

        self.view = QTableView()
        self.view.setModel(self.model)
        self.view.setAlternatingRowColors(True)
        self.view.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.view.horizontalHeader().setStretchLastSection(True)
        self.view.setColumnWidth(0, 60)
        install_delegates(self.view, self.model, self._streams_for_delegate)
        # hide the rarely-used columns for a clean first view
        self._apply_column_visibility()
        lv.addWidget(QLabel("Circuit builder"))
        lv.addWidget(self.view, 1)
        split.addWidget(left)

        # right: results + log tabs
        right = QTabWidget()
        self.results = ResultsView()
        right.addTab(self.results, "Results")
        self.log_view = QPlainTextEdit(); self.log_view.setReadOnly(True)
        right.addTab(self.log_view, "Log")
        self.right_tabs = right
        split.addWidget(right)
        split.setSizes([720, 640])
        self.setCentralWidget(split)

    def _build_statusbar(self):
        sb = QStatusBar()
        self.setStatusBar(sb)
        # logo bottom-left
        logo = QLabel()
        if os.path.exists(api.LOGO_PATH):
            pm = QPixmap(api.LOGO_PATH).scaledToHeight(26, Qt.SmoothTransformation)
            logo.setPixmap(pm)
        logo.setContentsMargins(8, 0, 8, 0)
        sb.addWidget(logo)
        self.hmb_lbl = QLabel("  No HMB loaded")
        sb.addWidget(self.hmb_lbl)
        self.state_lbl = QLabel("idle")
        self.state_lbl.setStyleSheet("color:#555; margin-right:10px;")
        sb.addPermanentWidget(self.state_lbl)

    def _apply_column_visibility(self):
        keep = {"Circuit", "Line No", "PID Number", "Stream Lookup", "Run Type",
                "Seq", "Start P (psia)", "Comp ID", "Fitting Name", "Bore (in)",
                "Piping Spec", "Length (ft)", "Elev Change (ft)",
                "Control Valve Type", "Set P (psia)", "Notes"}
        for i, h in enumerate(self.model.headers):
            self.view.setColumnHidden(i, h not in keep)

    # ── helpers ──
    def _rows(self):
        return self.model.to_rows()

    def _context(self):
        return {
            "hmb_path": self.hmb_path,
            "case": self.cb_case.currentText() or "Case 1",
            "flash_mode": self.cb_mode.currentText() or "isothermal",
            "units": "FPS",
            "meta": {"project": self.ed_project.text(), "area": self.ed_area.text(),
                     "unit": self.ed_unit.text()},
        }

    def _streams_for_delegate(self):
        if not self.hmb_path:
            return []
        return api.list_streams(self.hmb_path, self.cb_case.currentText() or "Case 1")

    def _dup_row(self):
        idx = self.view.currentIndex()
        if idx.isValid():
            self.model.duplicate_row(idx.row())

    def _del_row(self):
        idx = self.view.currentIndex()
        if idx.isValid():
            self.model.delete_row(idx.row())

    def _load_hmb(self):
        path, _ = QFileDialog.getOpenFileName(self, "Load HMB dump", "",
                                              "Excel (*.xlsx)")
        if not path:
            return
        self.hmb_path = path
        cases = api.list_cases(path) or ["Case 1"]
        self.cb_case.clear(); self.cb_case.addItems(cases)
        kind = "PRO/II Case-N export" if api.is_proii_export(path) else "per-stream dump"
        n = len(api.list_streams(path, cases[0]))
        self._refresh_hmb_label(f"{os.path.basename(path)}  ·  {kind}  ·  {n} streams")
        self._log(f"Loaded HMB: {path}  ({kind}, {n} streams, cases={cases})")

    def _refresh_hmb_label(self, text=None):
        if hasattr(self, "hmb_lbl"):
            self.hmb_lbl.setText("  " + (text or ("No HMB loaded" if not self.hmb_path
                                                  else os.path.basename(self.hmb_path))))

    def _toggle_auto(self, on):
        self.controller.auto = on
        self.auto_act.setText("Auto ●" if on else "Auto")
        if on:
            self.controller.schedule()

    def _on_grid_edited(self):
        self.controller.schedule()

    def _open_pms(self):
        try:
            from pms.pms_manager import PMSManagerDialog
            PMSManagerDialog(self).exec()
        except Exception as e:  # pragma: no cover
            QMessageBox.information(self, "PMS", f"PMS manager: {e}")

    def _export(self):
        QMessageBox.information(self, "Export",
                                "Export to Excel/PDF arrives in the next build step.")

    # ── run signals ──
    def _set_running(self, running):
        self.run_act.setEnabled(not running)

    def _on_state(self, s):
        self.state_lbl.setText(s)

    def _log(self, line):
        self.log_view.appendPlainText(line)

    def _on_finished(self, paths):
        self._set_running(False)
        self.results.load(paths)
        self.right_tabs.setCurrentWidget(self.results)
        self._log(f"Done — {len(paths)} workbook(s).")

    def _on_error(self, msg):
        self._set_running(False)
        self._log("ERROR: " + msg)
        self.right_tabs.setCurrentWidget(self.log_view)
