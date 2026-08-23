"""PMS (piping-material-spec) manager dialog.

A self-contained QDialog for browsing, editing and live-previewing the
piping-material-class catalogue exposed by ``engine_api`` / ``pms_classes``.

The main window opens it with::

    from pms.pms_manager import PMSManagerDialog
    PMSManagerDialog(self).exec()

It never mutates the active catalogue while previewing (that goes through
``engine_api.resolve_id_preview`` on a throwaway working class); only an
explicit Save rebuilds ``pms_classes.CLASSES`` and reloads the engine.
"""
from __future__ import annotations

import dataclasses
import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QWidget, QSplitter, QVBoxLayout, QHBoxLayout, QFormLayout,
    QGroupBox, QLabel, QLineEdit, QListWidget, QListWidgetItem, QPushButton,
    QDoubleSpinBox, QComboBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QFileDialog, QMessageBox, QInputDialog, QAbstractItemView,
)

import engine_api as api


# short material code from the full MOC text (KCS / LTCS / SS304 / A20 …)
_MOC_RULES = [
    ("low temp", "LTCS"), ("lt carbon", "LTCS"), ("impact tested", "LTCS"),
    ("killed carbon", "KCS"), ("carbon steel", "CS"), ("carbon stl", "CS"),
    ("304l", "SS304L"), ("304", "SS304"), ("316l", "SS316L"), ("316", "SS316"),
    ("321", "SS321"), ("347", "SS347"), ("317", "SS317"),
    ("duplex", "DSS"), ("2205", "DSS"), ("2507", "SDSS"),
    ("alloy 20", "A20"), ("n08020", "A20"), ("825", "A825"), ("625", "IN625"),
    ("inconel", "INC"), ("incoloy", "INC"), ("monel", "MONEL"),
    ("hastelloy", "HAST"), ("nickel", "NI"),
    ("5cr", "5Cr"), ("5 cr", "5Cr"), ("9cr", "9Cr"), ("9 cr", "9Cr"),
    ("1.25cr", "1¼Cr"), ("2.25cr", "2¼Cr"), ("chrome", "Cr-Mo"),
    ("ductile", "DI"), ("nodular", "DI"), ("cast iron", "CI"),
    ("galvan", "GALV"), ("copper", "Cu"), ("cupro", "CuNi"),
    ("titanium", "Ti"), ("gre", "GRE"), ("frp", "FRP"), ("pvc", "PVC"),
    ("pvdf", "PVDF"), ("ptfe", "PTFE"),
]


def _short_moc(moc: str, moc_tag: str = "") -> str:
    s = (moc or "").lower()
    for key, code in _MOC_RULES:
        if key in s:
            return code
    if (" cr" in s or s.strip().endswith("cr")) and "chrome" not in s:
        return "Cr-Mo"
    if "high density" in s or "hdpe" in s:
        return "HDPE"
    if moc_tag and str(moc_tag).strip():
        return str(moc_tag).strip()[:8]
    # fall back to the first word that looks like a material name, skipping
    # extraction noise (numbers, 'allowance', 'note', punctuation)
    first = (moc or "").split(",")[0].strip()
    low = first.lower()
    if not first or first[0].isdigit() or any(
            k in low for k in ("allowance", "note", "n/a", "see ")):
        return "—"
    return first[:10]

try:
    from resources import theme
    _TEN = theme.TEN
except Exception:  # pragma: no cover - theme is best-effort styling only
    _TEN = {"blue": "#0070EF", "navy": "#0D0E10", "amber": "#FDC300",
            "salmon": "#EE7766", "panel": "#F7F8F8", "bg": "#FFFFFF",
            "text": "#0D0E10", "lgray": "#E4E5E8", "gray": "#6B6E76",
            "border": "#E4E5E8", "primary_hover": "#0059C1"}

# Sensible NPS bore list for the live-preview drop-down.
_BORES = ["0.5", "0.75", "1", "1.5", "2", "3", "4", "6", "8", "10", "12",
          "16", "20", "24"]

_RULE_COLS = ["NPS low", "NPS high", "Schedule", "Min schedule", "Ends",
              "Description", "Special Thickness (in)"]

_PT_COLS = ["Temp (°C)", "Pressure (kg/cm2g)", "Temp (°F)", "Pressure (psig)"]


class PMSManagerDialog(QDialog):
    """Browse / edit / preview the piping-material-spec catalogue."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PMS — Piping Material Specs")
        self.resize(1260, 720)

        self._pms = api.pms_classes_module()
        self._loading = False          # guard: suppress recompute while populating
        self._current_name = None      # name of the class shown in the editor

        self._build_ui()
        self._style()
        self._refresh_class_list()
        if self.list.rowCount():
            self.list.setCurrentCell(0, 0)
        self._update_preview()

    # ── construction ────────────────────────────────────────────────────
    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        # top: prominent catalogue loader + status
        top = QHBoxLayout()
        top.setSpacing(10)
        self.btn_load = QPushButton("Load catalogue (JSON)…")
        self.btn_load.setObjectName("LoadCatalogue")
        self.btn_load.clicked.connect(self._on_load_catalogue)
        self.btn_load_xlsx = QPushButton("Load catalogue (Excel)…")
        self.btn_load_xlsx.clicked.connect(self._on_load_catalogue_xlsx)
        self.btn_export_xlsx = QPushButton("Save as Excel…")
        self.btn_export_xlsx.setObjectName("Secondary")
        self.btn_export_xlsx.clicked.connect(self._on_export_xlsx)
        self.status_lbl = QLabel("")
        self.status_lbl.setObjectName("Status")
        top.addWidget(self.btn_load)
        top.addWidget(self.btn_load_xlsx)
        top.addWidget(self.btn_export_xlsx)
        top.addWidget(self.status_lbl, 1)
        root.addLayout(top)

        split = QSplitter(Qt.Horizontal)
        split.addWidget(self._build_left())
        split.addWidget(self._build_right())
        split.setSizes([440, 760])
        root.addWidget(split, 1)

        # bottom bar
        bottom = QHBoxLayout()
        self.btn_new = QPushButton("New code…")
        self.btn_new.clicked.connect(self._on_new_code)
        self.btn_save = QPushButton("Save")
        self.btn_save.clicked.connect(self._on_save)
        self.btn_close = QPushButton("Close")
        self.btn_close.clicked.connect(self.accept)
        bottom.addWidget(self.btn_new)
        bottom.addStretch(1)
        bottom.addWidget(self.btn_save)
        bottom.addWidget(self.btn_close)
        root.addLayout(bottom)

    def _build_left(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search code / MOC…")
        self.search.textChanged.connect(self._apply_filter)
        self.list = QTableWidget(0, 4)
        self.list.setHorizontalHeaderLabels(["Code", "MOC", "C.A. (in)", "Rating"])
        self.list.verticalHeader().setVisible(False)
        self.list.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.list.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.list.setAlternatingRowColors(True)
        hh = self.list.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.Stretch)
        hh.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.list.currentCellChanged.connect(
            lambda r, *_: self._on_row_selected(r))
        lay.addWidget(self.search)
        lay.addWidget(self.list, 1)
        return w

    def _build_right(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(8)

        self.header_lbl = QLabel("—")
        self.header_lbl.setObjectName("EditorHeader")
        lay.addWidget(self.header_lbl)

        self.warn_lbl = QLabel("code-owned override — edit in source")
        self.warn_lbl.setObjectName("Warn")
        self.warn_lbl.setVisible(False)
        lay.addWidget(self.warn_lbl)

        # numeric fields
        fields_box = QGroupBox("Design parameters")
        form = QFormLayout(fields_box)
        self.sp_pressure = self._spin(0.0, 1_000_000.0, 1.0)
        self.sp_temp = self._spin(-460.0, 3000.0, 1.0)
        self.sp_corr = self._spin(0.0, 5.0, 0.001, decimals=3)
        self.sp_corr_min = self._spin(0.0, 5.0, 0.001, decimals=3)
        form.addRow("Design pressure (psig)", self.sp_pressure)
        form.addRow("Design temp (°F)", self.sp_temp)
        form.addRow("Corrosion allow (in)", self.sp_corr)
        form.addRow("Corrosion allow, min (in)", self.sp_corr_min)
        lay.addWidget(fields_box)

        # pipe rules table
        rules_box = QGroupBox("Pipe rules")
        rlay = QVBoxLayout(rules_box)
        self.table = QTableWidget(0, len(_RULE_COLS))
        self.table.setHorizontalHeaderLabels(_RULE_COLS)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.itemChanged.connect(self._on_rule_changed)
        rlay.addWidget(self.table)
        rbtns = QHBoxLayout()
        self.btn_add_row = QPushButton("Add row")
        self.btn_add_row.clicked.connect(self._on_add_row)
        self.btn_rm_row = QPushButton("Remove row")
        self.btn_rm_row.clicked.connect(self._on_remove_row)
        rbtns.addWidget(self.btn_add_row)
        rbtns.addWidget(self.btn_rm_row)
        rbtns.addStretch(1)
        rlay.addLayout(rbtns)
        lay.addWidget(rules_box, 1)

        # temperature/pressure curve -- shown for the selected spec (see
        # PipingClass.pt_curve), up to 10 points, SI and FPS independently
        # (real HURL sheets don't always use a pure unit conversion between
        # the two, so both are stored and edited as given).
        pt_box = QGroupBox("Temperature / Pressure curve")
        ptlay = QVBoxLayout(pt_box)
        self.pt_table = QTableWidget(0, len(_PT_COLS))
        self.pt_table.setHorizontalHeaderLabels(_PT_COLS)
        self.pt_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.pt_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.pt_table.itemChanged.connect(self._on_rule_changed)
        ptlay.addWidget(self.pt_table)
        ptbtns = QHBoxLayout()
        self.btn_add_pt = QPushButton("Add point")
        self.btn_add_pt.clicked.connect(self._on_add_pt)
        self.btn_rm_pt = QPushButton("Remove point")
        self.btn_rm_pt.clicked.connect(self._on_remove_pt)
        ptbtns.addWidget(self.btn_add_pt)
        ptbtns.addWidget(self.btn_rm_pt)
        ptbtns.addStretch(1)
        ptlay.addLayout(ptbtns)
        lay.addWidget(pt_box)

        # live preview
        prev_box = QGroupBox("Live ID preview")
        play = QHBoxLayout(prev_box)
        play.addWidget(QLabel("Bore (in)"))
        self.cb_bore = QComboBox()
        self.cb_bore.addItems(_BORES)
        self.cb_bore.setCurrentText("4")
        self.cb_bore.currentTextChanged.connect(self._update_preview)
        play.addWidget(self.cb_bore)
        self.preview_lbl = QLabel("—")
        self.preview_lbl.setObjectName("Preview")
        self.preview_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        play.addWidget(self.preview_lbl, 1)
        lay.addWidget(prev_box)

        # recompute preview on numeric edits
        for sp in (self.sp_pressure, self.sp_temp, self.sp_corr,
                   self.sp_corr_min):
            sp.valueChanged.connect(self._update_preview)

        return w

    def _spin(self, lo, hi, step, decimals=2) -> QDoubleSpinBox:
        sp = QDoubleSpinBox()
        sp.setRange(lo, hi)
        sp.setSingleStep(step)
        sp.setDecimals(decimals)
        sp.setKeyboardTracking(False)
        return sp

    def _style(self):
        t = _TEN
        blue = t.get('blue', '#5E6AD2'); ink = t.get('navy', '#0D0E10')
        hover = t.get('primary_hover', '#828FFF')
        panel = t.get('panel', '#F7F8F8'); border = t.get('border', '#E4E5E8')
        self.setStyleSheet(f"""
        QDialog {{ background: #FFFFFF; color: {ink};
            font-family: 'Inter','Segoe UI',system-ui,Arial; font-size: 14px; }}
        QGroupBox {{ font-weight: 600; font-size: 13px; border: 1px solid {border};
            border-radius: 12px; margin-top: 14px; padding: 16px 10px 10px 10px;
            background: #FFFFFF; }}
        QGroupBox::title {{ subcontrol-origin: margin; left: 12px; padding: 0 4px;
            color: {ink}; }}
        QPushButton {{ background: {blue}; color: white;
            border: 0; padding: 8px 14px; border-radius: 8px; font-weight: 500; }}
        QPushButton:hover {{ background: {hover}; }}
        QPushButton:disabled {{ background: #C7CAE8; }}
        QPushButton#LoadCatalogue {{ background: {blue}; font-size: 14px;
            padding: 9px 18px; font-weight: 600; }}
        QPushButton#LoadCatalogue:hover {{ background: {hover}; }}
        QLineEdit, QComboBox, QDoubleSpinBox {{
            background: white; border: 1px solid #D9DBE0;
            border-radius: 8px; padding: 6px 10px; min-height: 20px; }}
        QLineEdit:focus, QComboBox:focus, QDoubleSpinBox:focus {{
            border: 1px solid {blue}; }}
        QListWidget, QTableWidget {{ background: white; border: 1px solid {border};
            border-radius: 12px; gridline-color: #EEEFF1;
            alternate-background-color: #FAFAFB;
            selection-background-color: {blue}; selection-color: white; }}
        QLabel#EditorHeader {{ color: {ink}; font-weight: 700; font-size: 15px; }}
        QLabel#Warn {{ color: {t.get('salmon', '#EE7766')}; font-weight: 600; }}
        QLabel#Status {{ color: {ink}; font-size: 12px; }}
        QLabel#Preview {{ font-family: 'JetBrains Mono','Consolas','DejaVu Sans Mono',monospace;
            font-size: 13px; }}
        QHeaderView::section {{ background: {panel}; color: {ink};
            padding: 8px 10px; border: 0; border-right: 1px solid {border};
            border-bottom: 1px solid #D3D5DA; font-weight: 600; font-size: 13px; }}
        """)

    # ── class list ──────────────────────────────────────────────────────
    def _classes(self):
        return self._pms.CLASSES

    def _refresh_class_list(self, select_name=None):
        self._loading = True
        classes = self._classes()
        self.list.setRowCount(len(classes))
        for i, c in enumerate(classes):
            rating = api.flange_class_for(c.name)
            ca = c.corrosion_allow_in
            cells = [
                c.name,
                _short_moc(c.moc, c.moc_tag),
                (f"{float(ca):.3f}" if ca else "—"),
                (f"{rating}#" if rating else "—"),
            ]
            for col, text in enumerate(cells):
                it = QTableWidgetItem(text)
                if col == 0:
                    it.setData(Qt.UserRole, c.name)
                    f = it.font(); f.setBold(True); it.setFont(f)
                if col >= 2:
                    it.setTextAlignment(int(Qt.AlignRight | Qt.AlignVCenter))
                it.setToolTip(c.moc or "")
                self.list.setItem(i, col, it)
        self._loading = False
        self._apply_filter()
        if select_name is not None:
            self._select_by_name(select_name)
        elif self.list.rowCount():
            self.list.setCurrentCell(0, 0)

    def _select_by_name(self, name):
        for i in range(self.list.rowCount()):
            it = self.list.item(i, 0)
            if it is not None and it.data(Qt.UserRole) == name:
                self.list.setCurrentCell(i, 0)
                return

    def _apply_filter(self):
        term = (self.search.text() or "").strip().lower()
        for i in range(self.list.rowCount()):
            hay = " ".join((self.list.item(i, c).text() if self.list.item(i, c) else "")
                           for c in range(self.list.columnCount())).lower()
            self.list.setRowHidden(i, bool(term) and term not in hay)

    def _find_class(self, name):
        for c in self._classes():
            if c.name == name:
                return c
        return None

    def _curated_names(self):
        return {pc.name.upper() for pc in getattr(self._pms, "CURATED", [])}

    def _is_readonly(self, name):
        if not name:
            return False
        up = name.upper()
        return (up in self._curated_names()
                or up in getattr(self._pms, "PIPE_RULE_PATCHES", {}))

    # ── selection / editor population ───────────────────────────────────
    def _on_row_selected(self, row):
        if row is None or row < 0:
            self._current_name = None
            return
        it = self.list.item(row, 0)
        if it is None:
            return
        name = it.data(Qt.UserRole)
        cls = self._find_class(name)
        if cls is None:
            return
        self._current_name = name
        self._populate_editor(cls)

    def _populate_editor(self, cls):
        self._loading = True
        try:
            self.header_lbl.setText(f"{cls.name}   ·   {cls.moc or ''}".rstrip(" ·"))
            self.sp_pressure.setValue(float(cls.design_pressure_psig or 0.0))
            self.sp_temp.setValue(float(cls.design_temp_f or 0.0))
            self.sp_corr.setValue(float(cls.corrosion_allow_in or 0.0))
            self.sp_corr_min.setValue(float(cls.corrosion_allow_min_in or 0.0))

            rules = cls.pipe_rules or ()
            self.table.setRowCount(0)
            self.table.setRowCount(len(rules))
            for r, rule in enumerate(rules):
                vals = [rule.nps_low, rule.nps_high, rule.schedule,
                        rule.min_schedule, rule.ends, rule.description,
                        rule.special_thickness_in]
                for c, v in enumerate(vals):
                    self.table.setItem(r, c, QTableWidgetItem(str(v) if v not in (None, "") else ""))

            curve = cls.pt_curve or ()
            self.pt_table.setRowCount(0)
            self.pt_table.setRowCount(len(curve))
            for r, pt in enumerate(curve):
                vals = [pt.temp_c, pt.pressure_kgcm2g, pt.temp_f, pt.pressure_psig]
                for c, v in enumerate(vals):
                    self.pt_table.setItem(r, c, QTableWidgetItem(str(v) if v is not None else ""))

            ro = self._is_readonly(cls.name)
            self._set_editor_readonly(ro)
            self.warn_lbl.setVisible(ro)
        finally:
            self._loading = False
        self._update_preview()

    def _set_editor_readonly(self, ro):
        for sp in (self.sp_pressure, self.sp_temp, self.sp_corr,
                   self.sp_corr_min):
            sp.setReadOnly(ro)
            sp.setEnabled(not ro)
        self.btn_add_row.setEnabled(not ro)
        self.btn_rm_row.setEnabled(not ro)
        self.btn_add_pt.setEnabled(not ro)
        self.btn_rm_pt.setEnabled(not ro)
        self.btn_save.setEnabled(not ro)
        # table cells: toggle editability
        trigger = (QAbstractItemView.NoEditTriggers if ro
                   else QAbstractItemView.AllEditTriggers)
        self.table.setEditTriggers(trigger)
        self.pt_table.setEditTriggers(trigger)

    # ── pipe-rule table actions ─────────────────────────────────────────
    def _on_add_row(self):
        self._loading = True
        try:
            r = self.table.rowCount()
            self.table.insertRow(r)
            defaults = ["", "", "STD", "STD", "", ""]
            for c, v in enumerate(defaults):
                self.table.setItem(r, c, QTableWidgetItem(v))
        finally:
            self._loading = False
        self._update_preview()

    def _on_remove_row(self):
        r = self.table.currentRow()
        if r < 0:
            r = self.table.rowCount() - 1
        if r >= 0:
            self.table.removeRow(r)
            self._update_preview()

    def _on_rule_changed(self, _item):
        if not self._loading:
            self._update_preview()

    # ── T/P curve table actions ──────────────────────────────────────────
    def _on_add_pt(self):
        self._loading = True
        try:
            r = self.pt_table.rowCount()
            self.pt_table.insertRow(r)
            for c in range(len(_PT_COLS)):
                self.pt_table.setItem(r, c, QTableWidgetItem(""))
        finally:
            self._loading = False
        self._update_preview()

    def _on_remove_pt(self):
        r = self.pt_table.currentRow()
        if r < 0:
            r = self.pt_table.rowCount() - 1
        if r >= 0:
            self.pt_table.removeRow(r)
            self._update_preview()

    # ── build a working (unsaved) PipingClass from editor state ─────────
    def _cell(self, r, c):
        it = self.table.item(r, c)
        return (it.text() if it is not None else "").strip()

    def _rules_from_table(self):
        PipeRule = self._pms.PipeRule
        rules = []
        for r in range(self.table.rowCount()):
            nps_low = self._cell(r, 0)
            nps_high = self._cell(r, 1)
            if not nps_low and not nps_high:
                continue
            thin_txt = self._cell(r, 6)
            try:
                thin = float(thin_txt) if thin_txt else None
            except ValueError:
                thin = None
            rules.append(PipeRule(
                nps_low=nps_low,
                nps_high=nps_high or nps_low,
                schedule=self._cell(r, 2) or "STD",
                ends=self._cell(r, 4),
                description=self._cell(r, 5),
                note="",
                min_schedule=self._cell(r, 3) or "STD",
                special_thickness_in=thin,
            ))
        return tuple(rules)

    def _pt_cell(self, r, c):
        it = self.pt_table.item(r, c)
        return (it.text() if it is not None else "").strip()

    def _pt_curve_from_table(self):
        PTPoint = self._pms.PTPoint
        points = []
        for r in range(self.pt_table.rowCount()):
            vals = [self._pt_cell(r, c) for c in range(4)]
            if not any(vals):
                continue

            def _f(s):
                try:
                    return float(s) if s else None
                except ValueError:
                    return None

            points.append(PTPoint(temp_c=_f(vals[0]), pressure_kgcm2g=_f(vals[1]),
                                  temp_f=_f(vals[2]), pressure_psig=_f(vals[3])))
        return tuple(points)

    def _working_class(self):
        """Rebuild the edited class as a frozen PipingClass, or None."""
        if not self._current_name:
            return None
        base = self._find_class(self._current_name)
        if base is None:
            return None
        try:
            return dataclasses.replace(
                base,
                design_pressure_psig=float(self.sp_pressure.value()),
                design_temp_f=float(self.sp_temp.value()),
                corrosion_allow_in=float(self.sp_corr.value()),
                corrosion_allow_min_in=float(self.sp_corr_min.value()),
                pipe_rules=self._rules_from_table(),
                pt_curve=self._pt_curve_from_table(),
            )
        except Exception:
            return base

    # ── live preview ────────────────────────────────────────────────────
    def _update_preview(self, *_):
        if self._loading:
            return
        wc = self._working_class()
        if wc is None:
            self.preview_lbl.setText("—")
            return
        bore = self.cb_bore.currentText()

        def fmt(v, nd=4):
            return "—" if v is None else f"{v:.{nd}f}"

        try:
            idr = api.resolve_id_preview(wc, bore)
            parts = [
                f"sched {idr.schedule or '—'}",
                f"ID {fmt(idr.id_in)}\"",
                f"OD {fmt(idr.od_in)}\"",
                f"wall {fmt(idr.wall_in)}\"",
            ]
            line = "   ".join(parts)
            if idr.basis:
                line += f"    ·  {idr.basis}"
        except Exception as e:  # never crash on a bad edit
            line = f"—   (preview error: {e})"

        try:
            fc = api.flange_class_for(wc.name)
        except Exception:
            fc = None
        line += f"    ·  flange class {fc if fc is not None else '—'}"

        self.preview_lbl.setText(line)

    # ── catalogue upload ────────────────────────────────────────────────
    def _on_load_catalogue(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Load PMS catalogue", "", "JSON (*.json)")
        if not path:
            return
        try:
            n = api.install_pms_catalogue(path)
        except ValueError as e:
            QMessageBox.critical(self, "Invalid catalogue", str(e))
            return
        except Exception as e:  # unexpected; still surface it, don't crash
            QMessageBox.critical(self, "Load failed", str(e))
            return
        self.status_lbl.setText(
            f"Loaded {n} classes from {os.path.basename(path)}")
        keep = self._current_name
        self._refresh_class_list(select_name=keep)
        if self.list.currentRow() < 0 and self.list.rowCount():
            self.list.setCurrentCell(0, 0)
        self._update_preview()

    def _on_load_catalogue_xlsx(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Load PMS catalogue (Excel)", "", "Excel (*.xlsx)")
        if not path:
            return
        try:
            n = api.install_pms_catalogue_from_flat_sheet(path)
        except ValueError as e:
            QMessageBox.critical(self, "Invalid catalogue", str(e))
            return
        except Exception as e:  # unexpected; still surface it, don't crash
            QMessageBox.critical(self, "Load failed", str(e))
            return
        self.status_lbl.setText(
            f"Loaded {n} classes from {os.path.basename(path)}")
        keep = self._current_name
        self._refresh_class_list(select_name=keep)
        if self.list.currentRow() < 0 and self.list.rowCount():
            self.list.setCurrentCell(0, 0)
        self._update_preview()

    def _on_export_xlsx(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save PMS catalogue as Excel", "pms_catalogue.xlsx",
            "Excel (*.xlsx)")
        if not path:
            return
        try:
            n = api.export_pms_catalogue_to_flat_sheet(path)
        except Exception as e:
            QMessageBox.critical(self, "Export failed", str(e))
            return
        self.status_lbl.setText(f"Exported {n} classes to {os.path.basename(path)}")

    # ── new code ────────────────────────────────────────────────────────
    def _on_new_code(self):
        name, ok = QInputDialog.getText(self, "New PMS code", "Class name:")
        name = (name or "").strip()
        if not ok or not name:
            return
        if self._find_class(name) is not None:
            QMessageBox.information(self, "New code",
                                    f"A class named {name!r} already exists.")
            self._select_by_name(name)
            return
        PipingClass = self._pms.PipingClass
        blank = PipingClass(
            name=name, moc="", moc_tag="", material_group="",
            typical_service="", nominal_rating="", pt_limit_text="",
            design_pressure_psig=0.0, design_temp_f=100.0,
            corrosion_allow_in=0.0, corrosion_allow_min_in=0.0,
        )
        self._classes().append(blank)
        self._classes().sort(key=lambda c: c.name)
        self._refresh_class_list(select_name=name)

    # ── save ────────────────────────────────────────────────────────────
    def _on_save(self):
        wc = self._working_class()
        if wc is None:
            QMessageBox.information(self, "Save", "No class selected.")
            return
        if self._is_readonly(wc.name):
            QMessageBox.information(
                self, "Save",
                "This class is a code-owned override — edit it in source.")
            return
        updated = [wc if c.name == wc.name else c for c in self._classes()]
        try:
            self._pms.save_user_classes(updated)
            self._pms.CLASSES[:] = sorted(updated, key=lambda c: c.name)
            api.reload_pms()
        except Exception as e:
            QMessageBox.critical(self, "Save failed", str(e))
            return
        self.status_lbl.setText(f"Saved & reloaded ({wc.name})")
        self._refresh_class_list(select_name=wc.name)
        self._update_preview()
