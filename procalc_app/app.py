"""Procalc Hydraulics — main window (P1: grid + HMB loader + Run + results)."""
from __future__ import annotations

import os

from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QAction, QPixmap, QIcon, QFont
from PySide6.QtWidgets import (QMainWindow, QWidget, QToolBar, QSplitter,
                               QTableView, QTabWidget, QPlainTextEdit, QLabel,
                               QFileDialog, QComboBox, QLineEdit, QFormLayout,
                               QGroupBox, QVBoxLayout, QHBoxLayout, QMessageBox,
                               QAbstractItemView, QStatusBar, QWidgetAction, QMenu,
                               QPushButton)

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

        brand = QLabel("Procalc")
        brand.setObjectName("Brand")
        bf = brand.font()
        bf.setLetterSpacing(QFont.AbsoluteSpacing, -0.6)
        brand.setFont(bf)
        brand.setContentsMargins(4, 0, 12, 0)
        tb.addWidget(brand)
        tb.addSeparator()

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
        # make the primary Run button read as a solid blue pill
        rb = tb.widgetForAction(self.run_act)
        if rb is not None:
            rb.setObjectName("RunBtn")

    def _build_central(self):
        split = QSplitter(Qt.Horizontal)

        # left: project meta + grid
        left = QWidget()
        lv = QVBoxLayout(left)
        lv.setContentsMargins(12, 12, 12, 12)
        lv.setSpacing(10)
        meta_box = QGroupBox("Project")
        form = QFormLayout(meta_box)
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)
        self.ed_project = QLineEdit("BTGOLD HCU")
        self.ed_area = QLineEdit("Unit 24")
        self.ed_unit = QLineEdit("HCU")
        self.cb_case = QComboBox(); self.cb_case.addItem("Case 1")
        # repopulated after the window is shown (on HMB load) — without this,
        # AdjustToContentsOnFirstShow freezes the box at its "Case 1" width
        # and clips longer case names.
        self.cb_case.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.cb_mode = QComboBox(); self.cb_mode.addItems(["isothermal", "isenthalpic"])
        row = QHBoxLayout()
        row.setSpacing(8)
        for w in (QLabel("Case"), self.cb_case, QLabel("Flash"), self.cb_mode):
            row.addWidget(w)
        rw = QWidget(); rw.setLayout(row)
        # units: system + per-quantity overrides (per project)
        self.unit_overrides: dict = {}
        self.cb_units = QComboBox(); self.cb_units.addItems(api.unit_systems())
        self.cb_units.currentTextChanged.connect(lambda *_: self.unit_overrides.clear()
                                                 or self.controller.schedule())
        self.btn_units = QPushButton("Overrides…"); self.btn_units.setObjectName("Secondary")
        self.btn_units.clicked.connect(self._edit_unit_overrides)
        urow = QHBoxLayout()
        urow.setSpacing(8)
        for w in (QLabel("Units"), self.cb_units, self.btn_units):
            urow.addWidget(w)
        urow.addStretch(1)
        uw = QWidget(); uw.setLayout(urow)
        form.addRow("Project", self.ed_project)
        form.addRow("Area", self.ed_area)
        form.addRow("Unit", self.ed_unit)
        form.addRow(rw)
        form.addRow(uw)
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
        builder_lbl = QLabel("Circuit builder")
        builder_lbl.setObjectName("SectionTitle")
        lv.addWidget(builder_lbl)
        lv.addWidget(self.view, 1)
        split.addWidget(left)

        # right: results + log tabs, with a dismissible error banner on top
        right_wrap = QWidget()
        rv = QVBoxLayout(right_wrap)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.setSpacing(0)
        self.err_bar = QLabel("")
        self.err_bar.setWordWrap(True)
        self.err_bar.setStyleSheet(
            "background:#FDECEC; color:#C0392B; padding:8px 16px; font-weight:600;"
            "font-size:13px; border-left:3px solid #E84242; border-bottom:1px solid #F3C6C6;")
        self.err_bar.hide()
        rv.addWidget(self.err_bar)
        right = QTabWidget()
        rv.addWidget(right, 1)
        self.results = ResultsView()
        right.addTab(self.results, "Results")
        self.log_view = QPlainTextEdit(); self.log_view.setReadOnly(True)
        right.addTab(self.log_view, "Log")
        self.right_tabs = right
        split.addWidget(right_wrap)
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
        logo.setContentsMargins(8, 0, 12, 0)
        sb.addWidget(logo)
        self.hmb_lbl = QLabel("  No HMB loaded")
        self.hmb_lbl.setObjectName("Muted")
        sb.addWidget(self.hmb_lbl)
        self.state_lbl = QLabel("idle")
        self.state_lbl.setObjectName("StateBadge")
        self.state_lbl.setContentsMargins(0, 0, 10, 0)
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
            "units": self.cb_units.currentText() or "FPS",
            "unit_overrides": dict(self.unit_overrides),
            "meta": {"project": self.ed_project.text(), "area": self.ed_area.text(),
                     "unit": self.ed_unit.text()},
        }

    def _edit_unit_overrides(self):
        from PySide6.QtWidgets import (QDialog, QVBoxLayout, QGridLayout,
                                       QDialogButtonBox, QScrollArea)
        sysname = self.cb_units.currentText() or "FPS"
        dlg = QDialog(self)
        dlg.setWindowTitle(f"Per-quantity unit overrides — base: {sysname}")
        dlg.resize(500, 560)
        outer = QVBoxLayout(dlg)
        outer.setContentsMargins(16, 16, 16, 16)
        outer.setSpacing(10)
        hint = QLabel("Leave a quantity on its system default, or pick a "
                      "unit to override it for this project.")
        hint.setObjectName("Muted")
        hint.setWordWrap(True)
        outer.addWidget(hint)
        area = QScrollArea(); area.setWidgetResizable(True)
        inner = QWidget(); grid = QGridLayout(inner)
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(8)
        combos = {}
        for i, (qty, name, info) in enumerate(api.unit_quantities()):
            default = info["defaults"].get(sysname)
            grid.addWidget(QLabel(f"{name}"), i, 0)
            cb = QComboBox()
            cb.addItem(f"(default: {default})", "")
            for un in info["units"]:
                cb.addItem(un, un)
            cur = self.unit_overrides.get(qty)
            if cur:
                j = cb.findData(cur)
                if j >= 0:
                    cb.setCurrentIndex(j)
            combos[qty] = cb
            grid.addWidget(cb, i, 1)
        area.setWidget(inner)
        outer.addWidget(area, 1)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject)
        outer.addWidget(bb)
        if dlg.exec():
            self.unit_overrides = {q: cb.currentData() for q, cb in combos.items()
                                   if cb.currentData()}
            n = len(self.unit_overrides)
            self.btn_units.setText(f"Overrides… ({n})" if n else "Overrides…")
            self.controller.schedule()

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
            rows = self.model.rows
            heavy = len(rows) > 60 or any(
                str(r.get("Fitting Name", "")).lower() == "control valve" for r in rows)
            if heavy:
                QMessageBox.information(
                    self, "Auto-calc",
                    "This circuit has control valves or many rows — each edit "
                    "triggers a full re-march (control valves run 3×). Auto-calc "
                    "is on; switch it off if edits feel slow.")
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
        path = self.results.current_workbook_path()
        if not path:
            QMessageBox.information(self, "Export", "Run a calculation first.")
            return
        sheet = self.results.current_sheet_title()
        menu = QMenu(self)
        menu.addAction("Excel — whole workbook (.xlsx)",
                       lambda: self._do_export("xlsx_wb", path, sheet))
        menu.addAction(f"Excel — current sheet ({sheet}) (.xlsx)",
                       lambda: self._do_export("xlsx_sheet", path, sheet))
        menu.addSeparator()
        menu.addAction("PDF — whole workbook (.pdf)",
                       lambda: self._do_export("pdf_wb", path, sheet))
        menu.addAction(f"PDF — current sheet ({sheet}) (.pdf)",
                       lambda: self._do_export("pdf_sheet", path, sheet))
        menu.exec(self.cursor().pos())

    def _do_export(self, kind, src, sheet):
        from results import export as EX
        logo = api.LOGO_PATH
        try:
            if kind == "xlsx_wb":
                dest, _ = QFileDialog.getSaveFileName(self, "Export workbook", "results.xlsx", "Excel (*.xlsx)")
                if dest:
                    EX.export_excel_workbook(src, dest)
            elif kind == "xlsx_sheet":
                dest, _ = QFileDialog.getSaveFileName(self, "Export sheet", f"{sheet}.xlsx", "Excel (*.xlsx)")
                if dest:
                    EX.export_excel_sheet(src, sheet, dest)
            elif kind == "pdf_wb":
                dest, _ = QFileDialog.getSaveFileName(self, "Export PDF", "results.pdf", "PDF (*.pdf)")
                if dest:
                    EX.export_pdf_workbook(src, dest, logo_path=logo)
            elif kind == "pdf_sheet":
                dest, _ = QFileDialog.getSaveFileName(self, "Export PDF", f"{sheet}.pdf", "PDF (*.pdf)")
                if dest:
                    EX.export_pdf_sheet(src, sheet, dest, logo_path=logo)
            else:
                return
            if dest:
                self._log(f"Exported: {dest}")
                self.statusBar().showMessage(f"Exported {os.path.basename(dest)}", 5000)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Export failed", str(e))

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
