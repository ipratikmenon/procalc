"""Procalc Hydraulics — main window (P1: grid + HMB loader + Run + results)."""
from __future__ import annotations

import base64
import os
import tempfile

from PySide6.QtCore import Qt, QSize, QTimer
from PySide6.QtGui import QAction, QPixmap, QIcon, QFont
from PySide6.QtWidgets import (QMainWindow, QWidget, QToolButton,
                               QSplitter, QDialog, QFrame, QScrollArea,
                               QTableView, QTabWidget, QPlainTextEdit, QLabel,
                               QFileDialog, QComboBox, QLineEdit, QFormLayout,
                               QGroupBox, QVBoxLayout, QHBoxLayout, QMessageBox,
                               QAbstractItemView, QStatusBar, QMenu,
                               QPushButton, QStackedWidget)

import engine_api as api
from model.grid_model import CircuitGridModel
from model.delegates import install_delegates
from model.project import Project
from run.controller import RunController
from results.results_view import ResultsView
from resources import theme
from widgets.scroll_toolbar import HScrollToolbar
from widgets.run_control import RunAutoControl
from stream_details.stream_details_panel import StreamDetailsPanel

_HERE = os.path.dirname(os.path.abspath(__file__))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Procalc Hydraulics — T.EN")
        self.resize(1360, 820)
        self.hmb_path = None
        self.hmb_case = "Case 1"
        self._project_path: str | None = None   # None until New/Open/Save
        self._stream_snapshot = None             # set on Open, cleared on new HMB load
        self.client_logo_path: str | None = None
        self._project_active = False             # True once New/Open has run (dashboard state)

        # start with an empty grid — the landing page ("Build a Project")
        # is what actually seeds it via New/Open, not an auto-populated sample
        self.model = CircuitGridModel()
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
        self._refresh_dashboard()
        self._central_stack.setCurrentWidget(self._dashboard)

        container = QWidget()
        cl = QVBoxLayout(container)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(0)
        cl.addWidget(self.toolbar)
        cl.addWidget(self._central_stack, 1)
        self.setCentralWidget(container)

    # ── UI construction ──
    def _build_toolbar(self):
        self.toolbar = HScrollToolbar(self)

        brand = QLabel("Procalc")
        brand.setObjectName("Brand")
        bf = brand.font()
        bf.setLetterSpacing(QFont.AbsoluteSpacing, -0.6)
        brand.setFont(bf)
        brand.setContentsMargins(4, 0, 12, 0)
        self.toolbar.add_widget(brand)
        self.toolbar.add_separator()

        def act(text, slot, checkable=False):
            a = QAction(text, self)
            a.setCheckable(checkable)
            a.triggered.connect(slot)
            btn = QToolButton()
            btn.setDefaultAction(a)
            self.toolbar.add_widget(btn)
            return a

        act("New", self._new_project)
        act("Open…", self._open_project)
        act("Save", self._save_project)
        act("Save As…", self._save_project_as)
        self.toolbar.add_separator()
        act("🏠 Dashboard", lambda: self._central_stack.setCurrentWidget(self._dashboard))
        act("Project…", self._open_project_dialog)
        self.toolbar.add_separator()
        act("＋ Row", lambda: self.model.add_row(template={"Circuit": "C1", "Run Type": "Main"}))
        act("⧉ Duplicate", self._dup_row)
        act("🗑 Delete", self._del_row)
        self.toolbar.add_separator()
        act("📂 Load HMB", self._load_hmb)

        self.run_auto = RunAutoControl(self)
        self.run_auto.run_btn.clicked.connect(self.controller.run_now)
        self.run_auto.toggle.toggled.connect(self._toggle_auto)
        self.toolbar.add_widget(self.run_auto)

        self.toolbar.add_separator()
        act("PMS…", self._open_pms)
        act("Export…", self._export)

    def _build_project_meta_boxes(self):
        """The Project + Document-header fields — no longer shown inline in
        the Circuit Builder screen (moved to a separate dialog, see
        _open_project_dialog); the widgets themselves stay MainWindow
        attributes since _context()/_meta_dict()/_save_to()/_open_project()
        all read/write them directly by name."""
        meta_box = QGroupBox("Project")
        form = QFormLayout(meta_box)
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)
        self.ed_project = QLineEdit()
        self.ed_area = QLineEdit()
        self.ed_unit = QLineEdit()
        self.ed_client = QLineEdit()
        self.ed_site = QLineEdit()
        self.ed_circuit_name = QLineEdit()
        self.ed_prep_by = QLineEdit()
        self.ed_chk_by = QLineEdit()
        self.ed_appr_by = QLineEdit()
        self.ed_revision = QLineEdit()
        self.ed_page = QLineEdit()
        self.btn_client_logo = QPushButton("Client Logo…")
        self.btn_client_logo.setObjectName("Secondary")
        self.btn_client_logo.clicked.connect(self._pick_client_logo)
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
        self._units_customized = False
        self._applying_units_programmatically = False
        self.cb_units = QComboBox(); self.cb_units.addItems(api.unit_systems())
        self.cb_units.currentTextChanged.connect(self._on_units_changed)
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
        form.addRow("Client", self.ed_client)
        form.addRow("Site", self.ed_site)
        form.addRow("Circuit Name", self.ed_circuit_name)
        form.addRow(rw)
        form.addRow(uw)

        # document header-block admin fields (Prep/Chk/Appr By, Revision,
        # Page) + the client-logo upload — collapsed into their own card so
        # the Project box above stays focused on the run-affecting fields
        admin_box = QGroupBox("Document header")
        aform = QFormLayout(admin_box)
        aform.setHorizontalSpacing(10)
        aform.setVerticalSpacing(8)
        pcrow = QHBoxLayout(); pcrow.setSpacing(8)
        for w in (QLabel("Prep"), self.ed_prep_by, QLabel("Chk"), self.ed_chk_by,
                  QLabel("Appr"), self.ed_appr_by):
            pcrow.addWidget(w)
        pcw = QWidget(); pcw.setLayout(pcrow)
        rprow = QHBoxLayout(); rprow.setSpacing(8)
        for w in (QLabel("Revision"), self.ed_revision, QLabel("Page"), self.ed_page):
            rprow.addWidget(w)
        rprow.addStretch(1)
        rpw = QWidget(); rpw.setLayout(rprow)
        aform.addRow(pcw)
        aform.addRow(rpw)
        aform.addRow(self.btn_client_logo)

        return meta_box, admin_box

    def _open_project_dialog(self):
        self._project_dlg.show()
        self._project_dlg.raise_()
        self._project_dlg.activateWindow()

    def _build_central(self):
        self._central_stack = QStackedWidget()
        self._dashboard = self._build_dashboard()
        self._central_stack.addWidget(self._dashboard)

        meta_box, admin_box = self._build_project_meta_boxes()
        self._project_dlg = QDialog(self)
        self._project_dlg.setWindowTitle("Project")
        self._project_dlg.resize(480, 520)
        pdl = QVBoxLayout(self._project_dlg)
        pdl.setContentsMargins(16, 16, 16, 16)
        pdl.setSpacing(12)
        pdl.addWidget(meta_box)
        pdl.addWidget(admin_box)
        pdl.addStretch(1)

        # ── outer split: [grid + Stream Details | Results/Log] ──
        outer = QSplitter(Qt.Horizontal)

        grid_wrap = QWidget()
        gv = QVBoxLayout(grid_wrap)
        gv.setContentsMargins(12, 12, 12, 12)
        gv.setSpacing(8)
        builder_lbl = QLabel("Circuit builder")
        builder_lbl.setObjectName("SectionTitle")
        gv.addWidget(builder_lbl)

        self.view = QTableView()
        self.view.setModel(self.model)
        self.view.setAlternatingRowColors(True)
        self.view.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.view.horizontalHeader().setStretchLastSection(True)
        self.view.setColumnWidth(0, 60)
        install_delegates(self.view, self.model, self._streams_for_delegate)
        # hide the rarely-used columns for a clean first view
        self._apply_column_visibility()
        gv.addWidget(self.view, 1)

        self.stream_details = StreamDetailsPanel()
        self.view.selectionModel().currentChanged.connect(self._on_grid_row_selected)
        self._stream_details_timer = QTimer(self)
        self._stream_details_timer.setSingleShot(True)
        self._stream_details_timer.timeout.connect(self._resolve_stream_details)
        self._pending_stream_row = None

        left_col = QSplitter(Qt.Vertical)
        left_col.addWidget(grid_wrap)
        left_col.addWidget(self.stream_details)
        left_col.setSizes([560, 220])
        outer.addWidget(left_col)

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
        outer.addWidget(right_wrap)
        outer.setSizes([760, 600])
        self._workspace = outer
        self._central_stack.addWidget(self._workspace)

    def _build_dashboard(self):
        w = QWidget()
        outer = QVBoxLayout(w)
        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(20)

        header = QVBoxLayout()
        header.setAlignment(Qt.AlignHCenter)
        header.setSpacing(8)

        logo = QLabel()
        logo_path = getattr(api, "PROCALC_LOGO_PATH", None) or api.LOGO_PATH
        if logo_path and os.path.exists(logo_path):
            pm = QPixmap(logo_path).scaledToHeight(56, Qt.SmoothTransformation)
            logo.setPixmap(pm)
        logo.setAlignment(Qt.AlignCenter)
        header.addWidget(logo)

        heading = QLabel("Dashboard")
        heading.setObjectName("Brand")
        hf = heading.font(); hf.setPointSize(20)
        heading.setFont(hf)
        heading.setAlignment(Qt.AlignCenter)
        header.addWidget(heading)

        sub = QLabel("Start a new hydraulic circuit, or open an existing .calc project.")
        sub.setObjectName("Muted")
        sub.setAlignment(Qt.AlignCenter)
        header.addWidget(sub)

        btn_row = QHBoxLayout()
        btn_row.setAlignment(Qt.AlignCenter)
        btn_row.setSpacing(10)
        btn_new = QPushButton("New Project")
        btn_new.clicked.connect(self._new_project)
        btn_open = QPushButton("Open Project…")
        btn_open.setObjectName("Secondary")
        btn_open.clicked.connect(self._open_project)
        btn_row.addWidget(btn_new)
        btn_row.addWidget(btn_open)
        btn_row_w = QWidget(); btn_row_w.setLayout(btn_row)
        header.addWidget(btn_row_w)
        header_w = QWidget(); header_w.setLayout(header)
        outer.addWidget(header_w)

        cards_row = QHBoxLayout()
        cards_row.setSpacing(16)
        self._card_circuits = self._make_dashboard_card("Recent Circuits")
        self._card_streams = self._make_dashboard_card("Recent Streams")
        self._card_valves = self._make_dashboard_card("Control Valves")
        for card in (self._card_circuits, self._card_streams, self._card_valves):
            cards_row.addWidget(card["frame"], 1)
        outer.addLayout(cards_row)
        outer.addStretch(1)
        return w

    def _make_dashboard_card(self, title):
        frame = QFrame()
        frame.setObjectName("Card")
        lay = QVBoxLayout(frame)
        lay.setContentsMargins(16, 14, 16, 14)
        lay.setSpacing(6)
        title_lbl = QLabel(title)
        title_lbl.setObjectName("CardTitle")
        lay.addWidget(title_lbl)
        body = QLabel("")
        body.setObjectName("Muted")
        body.setWordWrap(True)
        body.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setWidget(body)
        lay.addWidget(scroll, 1)
        return {"frame": frame, "body": body}

    def _refresh_dashboard(self):
        cards = (self._card_circuits, self._card_streams, self._card_valves)
        if not getattr(self, "_project_active", False):
            msg = "Start a new project or open one to see project data here."
            for card in cards:
                card["body"].setText(msg)
            return
        rows = self.model.to_rows()
        if not rows:
            msg = "No rows yet — add circuit rows in the builder."
            for card in cards:
                card["body"].setText(msg)
            return
        circuits = list(dict.fromkeys(r.get("Circuit") for r in rows if r.get("Circuit")))
        self._card_circuits["body"].setText(
            "\n".join(circuits) if circuits else "No circuits in this project.")
        streams = list(dict.fromkeys(
            r.get("Stream Lookup") for r in rows if r.get("Stream Lookup")))
        self._card_streams["body"].setText(
            "\n".join(streams) if streams else "No stream lookups set yet.")
        valves = [r for r in rows
                 if str(r.get("Fitting Name", "")).lower() == "control valve"]
        if valves:
            lines = [f"{r.get('Comp ID', '?')} — {r.get('Control Valve Type') or 'F'}"
                    for r in valves]
            self._card_valves["body"].setText("\n".join(lines))
        else:
            self._card_valves["body"].setText("No control valves in this circuit.")

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
            "meta": self._meta_dict(),
            "stream_snapshot": self._stream_snapshot,
        }

    def _meta_dict(self):
        return {
            "project": self.ed_project.text(), "area": self.ed_area.text(),
            "unit": self.ed_unit.text(), "client": self.ed_client.text(),
            "site": self.ed_site.text(), "circuit_name": self.ed_circuit_name.text(),
            "prep_by": self.ed_prep_by.text(), "chk_by": self.ed_chk_by.text(),
            "appr_by": self.ed_appr_by.text(), "revision": self.ed_revision.text(),
            "page": self.ed_page.text(), "client_logo_path": self.client_logo_path,
        }

    def _set_meta_dict(self, m):
        m = m or {}
        self.ed_project.setText(m.get("project", "") or "")
        self.ed_area.setText(m.get("area", "") or "")
        self.ed_unit.setText(m.get("unit", "") or "")
        self.ed_client.setText(m.get("client", "") or "")
        self.ed_site.setText(m.get("site", "") or "")
        self.ed_circuit_name.setText(m.get("circuit_name", "") or "")
        self.ed_prep_by.setText(m.get("prep_by", "") or "")
        self.ed_chk_by.setText(m.get("chk_by", "") or "")
        self.ed_appr_by.setText(m.get("appr_by", "") or "")
        self.ed_revision.setText(m.get("revision", "") or "")
        self.ed_page.setText(m.get("page", "") or "")

    def _on_units_changed(self, *_):
        self.unit_overrides.clear()
        if not self._applying_units_programmatically:
            self._units_customized = True
        self.controller.schedule()

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
            self._units_customized = True
            self.controller.schedule()

    def _streams_for_delegate(self):
        if not self.hmb_path:
            return []
        try:
            return api.list_streams(self.hmb_path, self.cb_case.currentText() or "Case 1")
        except Exception:
            # a transient live-connector hiccup (HYSYS/PRO-II) shouldn't crash
            # the dropdown editor — _load_hmb() is where a connect failure
            # gets a real, actionable error message
            return []

    def _on_grid_row_selected(self, current, previous):
        if not current.isValid() or current.row() >= len(self.model.rows):
            self.stream_details.show_empty(
                "Select a row with a Stream Lookup value to see resolved "
                "stream properties.")
            self._pending_stream_row = None
            return
        # debounced so fast arrow-key navigation doesn't fire one live
        # resolve per row — the connector's own session cache (Change 12/13)
        # keeps a settled selection cheap either way
        self._pending_stream_row = current.row()
        self._stream_details_timer.start(200)

    def _resolve_stream_details(self):
        row_idx = self._pending_stream_row
        if row_idx is None or row_idx >= len(self.model.rows):
            return
        row = self.model.rows[row_idx]
        stream = row.get("Stream Lookup")
        if not stream:
            self.stream_details.show_empty(
                "Select a row with a Stream Lookup value to see resolved "
                "stream properties.")
            return
        hmb = row.get("HMB File") or self.hmb_path
        case = row.get("Case") or (self.cb_case.currentText() or "Case 1")
        sp = None
        if hmb:
            try:
                sp, _feed = api.resolve_stream_for_snapshot(hmb, case, stream)
            except Exception:
                sp = None
        if sp is None and self._stream_snapshot:
            entry = self._stream_snapshot.get((hmb, case, stream))
            if entry:
                sp = entry[0]
        if sp is None:
            self.stream_details.show_empty(
                f"Could not resolve stream {stream!r} — check the HMB "
                "connection and case.")
            return
        self.stream_details.show_props(sp)

    def _dup_row(self):
        idx = self.view.currentIndex()
        if idx.isValid():
            self.model.duplicate_row(idx.row())

    def _del_row(self):
        idx = self.view.currentIndex()
        if idx.isValid():
            self.model.delete_row(idx.row())

    def _load_hmb(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Load HMB / connect to a simulation", "",
            "All supported (*.xlsx *.hsc *.prz);;Excel HMB dump (*.xlsx);;"
            "HYSYS case — live, must already be open (*.hsc);;"
            "PRO/II database — live (*.prz)")
        if not path:
            return
        self.err_bar.hide()
        try:
            cases = api.list_cases(path) or ["Case 1"]
        except Exception as exc:
            self._show_error(f"Could not connect to {os.path.basename(path)}: {exc}")
            return
        self.hmb_path = path
        # a freshly-loaded HMB supersedes any snapshot restored from a .calc —
        # live resolution takes over; the next Save rebuilds a fresh snapshot
        self._stream_snapshot = None

        if not self._units_customized:
            try:
                sysname, overrides = api.detect_hmb_units(path, cases[0])
            except Exception:
                sysname, overrides = None, None
            if sysname:
                self._applying_units_programmatically = True
                try:
                    ui = self.cb_units.findText(sysname)
                    self.cb_units.setCurrentIndex(ui if ui >= 0 else 0)
                    self.unit_overrides = dict(overrides or {})
                    n = len(self.unit_overrides)
                    self.btn_units.setText(f"Overrides… ({n})" if n else "Overrides…")
                finally:
                    self._applying_units_programmatically = False
                self.controller.schedule()
                self._log(f"Units auto-detected from HMB: {sysname} "
                         f"({len(self.unit_overrides)} override"
                         f"{'s' if len(self.unit_overrides) != 1 else ''})")

        self.cb_case.clear(); self.cb_case.addItems(cases)
        kind_label = {
            "hysys": "HYSYS (live)",
            "proii_com": "PRO/II (live, re-solved on connect)",
            "proii_xlsx": "PRO/II Case-N export",
            "per_stream_xlsx": "per-stream dump",
        }.get(api.hmb_source_kind(path), "per-stream dump")
        try:
            n = len(api.list_streams(path, cases[0]))
        except Exception as exc:
            self._show_error(f"Connected, but could not list streams: {exc}")
            n = 0
        self._refresh_hmb_label(f"{os.path.basename(path)}  ·  {kind_label}  ·  {n} streams")
        self._log(f"Loaded HMB: {path}  ({kind_label}, {n} streams, cases={cases})")

    def _refresh_hmb_label(self, text=None):
        if hasattr(self, "hmb_lbl"):
            self.hmb_lbl.setText("  " + (text or ("No HMB loaded" if not self.hmb_path
                                                  else os.path.basename(self.hmb_path))))

    # ── project: New / Open / Save / Save As ──
    def _pick_client_logo(self):
        path, _ = QFileDialog.getOpenFileName(self, "Client logo image", "",
                                              "Images (*.png *.jpg *.jpeg)")
        if not path:
            return
        self.client_logo_path = path
        self.btn_client_logo.setText(f"Client Logo… ({os.path.basename(path)})")
        self.controller.schedule()

    def _new_project(self):
        self._project_path = None
        self._stream_snapshot = None
        self.hmb_path = None
        self.client_logo_path = None
        self.btn_client_logo.setText("Client Logo…")
        self.model.set_rows([])
        self._set_meta_dict({})
        self.unit_overrides = {}
        self._units_customized = False
        self.cb_units.setCurrentIndex(0)
        self._refresh_hmb_label()
        self.setWindowTitle("Procalc Hydraulics — T.EN — (untitled)")
        self._project_active = True
        self._refresh_dashboard()
        self._central_stack.setCurrentWidget(self._workspace)

    def _open_project(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Project", "",
                                              "Procalc Project (*.calc)")
        if not path:
            return
        try:
            proj = Project.load(path)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Open failed", str(e))
            return

        self._project_path = path
        self.model.set_rows(proj.grid_rows)
        self.hmb_path = proj.hmb_path
        self._stream_snapshot = (api.stream_snapshot_from_json(proj.stream_snapshot)
                                 if proj.stream_snapshot else None)

        cases = []
        if self.hmb_path and os.path.exists(self.hmb_path):
            try:
                cases = api.list_cases(self.hmb_path) or []
            except Exception:
                cases = []
        self.cb_case.clear()
        self.cb_case.addItems(cases or [proj.case or "Case 1"])
        idx = self.cb_case.findText(proj.case or "Case 1")
        if idx >= 0:
            self.cb_case.setCurrentIndex(idx)
        self.cb_mode.setCurrentText(proj.flash_mode or "isothermal")
        ui = self.cb_units.findText(proj.unit_system or "FPS")
        self.cb_units.setCurrentIndex(ui if ui >= 0 else 0)
        self.unit_overrides = dict(proj.unit_overrides or {})
        n = len(self.unit_overrides)
        self.btn_units.setText(f"Overrides… ({n})" if n else "Overrides…")
        self._units_customized = True

        self._set_meta_dict(proj.meta)
        self.client_logo_path = None
        if proj.client_logo_b64:
            try:
                fd, tmp_path = tempfile.mkstemp(prefix="procalc_client_logo_", suffix=".png")
                with os.fdopen(fd, "wb") as f:
                    f.write(base64.b64decode(proj.client_logo_b64))
                self.client_logo_path = tmp_path
                self.btn_client_logo.setText("Client Logo… (restored)")
            except Exception:
                pass
        else:
            self.btn_client_logo.setText("Client Logo…")

        self._refresh_hmb_label(
            None if (self.hmb_path and os.path.exists(self.hmb_path))
            else (f"{os.path.basename(self.hmb_path)}  ·  from saved snapshot (file not found)"
                  if self.hmb_path else None))
        self.setWindowTitle(f"Procalc Hydraulics — T.EN — {os.path.basename(path)}")
        self._project_active = True
        self._refresh_dashboard()
        self._central_stack.setCurrentWidget(self._workspace)
        self._log(f"Opened project: {path}")

    def _save_project(self):
        if not self._project_path:
            self._save_project_as()
            return
        self._save_to(self._project_path)

    def _save_project_as(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Project As", "project.calc",
                                              "Procalc Project (*.calc)")
        if not path:
            return
        if not path.lower().endswith(".calc"):
            path += ".calc"
        self._save_to(path)

    def _save_to(self, path):
        rows = self.model.to_rows()
        ctx = self._context()
        snapshot_json = {}
        if self.hmb_path:
            try:
                keys = api.snapshot_keys_for_rows(rows, self.hmb_path, ctx["case"])
                snap = {k: api.resolve_stream_for_snapshot(*k) for k in keys}
                snapshot_json = api.stream_snapshot_to_json(snap) if snap else {}
            except Exception as e:  # noqa: BLE001
                QMessageBox.critical(
                    self, "Save failed",
                    f"Could not resolve HMB stream data for saving: {e}")
                return
        client_logo_b64 = None
        if self.client_logo_path and os.path.exists(self.client_logo_path):
            with open(self.client_logo_path, "rb") as f:
                client_logo_b64 = base64.b64encode(f.read()).decode("ascii")
        proj = Project(
            grid_rows=rows, hmb_path=self.hmb_path, case=ctx["case"],
            flash_mode=ctx["flash_mode"], unit_system=ctx["units"],
            unit_overrides=ctx["unit_overrides"], meta=ctx["meta"],
            stream_snapshot=snapshot_json, client_logo_b64=client_logo_b64,
        )
        try:
            proj.save(path)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Save failed", str(e))
            return
        self._project_path = path
        self.setWindowTitle(f"Procalc Hydraulics — T.EN — {os.path.basename(path)}")
        self.statusBar().showMessage(f"Saved {os.path.basename(path)}", 5000)
        self._log(f"Saved project: {path}")

    def _toggle_auto(self, on):
        self.controller.auto = on
        self.run_auto.auto_lbl.setText("Auto ●" if on else "Auto")
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
        self._refresh_dashboard()

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
        self.run_auto.run_btn.setEnabled(not running)

    def _on_state(self, s):
        self.state_lbl.setText(s)

    def _log(self, line):
        self.log_view.appendPlainText(line)

    def _show_error(self, msg):
        self.err_bar.setText(msg)
        self.err_bar.show()
        self._log("ERROR: " + msg)

    def _on_finished(self, paths):
        self._set_running(False)
        self.results.load(paths)
        self.right_tabs.setCurrentWidget(self.results)
        self._log(f"Done — {len(paths)} workbook(s).")

    def _on_error(self, msg):
        self._set_running(False)
        self._log("ERROR: " + msg)
        self.right_tabs.setCurrentWidget(self.log_view)
