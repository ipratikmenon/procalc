"""Bottom-left panel of the Circuit Builder screen: shows the resolved
stream properties for whichever grid row is currently selected.

Resolution goes through engine_api.resolve_stream_for_snapshot() when a
live HMB is loaded, or the .calc-restored stream snapshot otherwise — both
paths raise/return None on a miss rather than silently showing stale data,
so this panel always reflects either real data or a clear empty state.
"""
from __future__ import annotations

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (QCheckBox, QFormLayout, QFrame, QLabel,
                               QScrollArea, QVBoxLayout, QWidget)

# Core fields shown by default; extended fields revealed by "More…".
_CORE_FIELDS = [
    ("stream", "Stream"), ("phase", "Phase"), ("temp_f", "Temperature (°F)"),
    ("pres_psia", "Pressure (psia)"), ("total_mass", "Total mass (lb/hr)"),
    ("vap_mass", "Vapor mass (lb/hr)"), ("liq_mass", "Liquid mass (lb/hr)"),
    ("mol_weight", "Mol. weight"), ("vap_density", "Vapor density (lb/ft3)"),
    ("liq_density", "Liquid density (lb/ft3)"), ("vap_visc", "Vapor visc. (cP)"),
    ("liq_visc", "Liquid visc. (cP)"),
]
_EXTENDED_FIELDS = [
    ("total_z", "Z factor"), ("vap_z", "Vapor Z"), ("vap_cp", "Vapor Cp (BTU/lb-F)"),
    ("liq_cp", "Liquid Cp (BTU/lb-F)"), ("vap_therm_cond", "Vapor k (BTU/hr-ft-F)"),
    ("liq_therm_cond", "Liquid k (BTU/hr-ft-F)"), ("liq_surf_tens", "Surface tension (dyn/cm)"),
    ("vap_sp_enthalpy", "Vapor h (BTU/lb)"), ("liq_sp_enthalpy", "Liquid h (BTU/lb)"),
    ("tc_f", "Tc (°F)"), ("pc_psia", "Pc (psia)"), ("acentric", "Acentric factor"),
]


def _fmt(v):
    if v is None:
        return "—"
    if isinstance(v, float):
        return f"{v:,.4g}"
    return str(v)


class StreamDetailsPanel(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(12, 8, 12, 8)
        outer.setSpacing(6)

        title = QLabel("Stream details")
        title.setObjectName("SectionTitle")
        outer.addWidget(title)

        self._empty_lbl = QLabel("Select a row with a Stream Lookup value to "
                                 "see resolved stream properties.")
        self._empty_lbl.setObjectName("Muted")
        self._empty_lbl.setWordWrap(True)
        outer.addWidget(self._empty_lbl)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        host = QWidget()
        self._form = QFormLayout(host)
        self._form.setHorizontalSpacing(10)
        self._form.setVerticalSpacing(4)
        scroll.setWidget(host)
        self._scroll = scroll
        outer.addWidget(scroll, 1)

        self._more_chk = QCheckBox("More…")
        self._more_chk.toggled.connect(self._on_more_toggled)
        outer.addWidget(self._more_chk)

        self._core_labels: dict[str, QLabel] = {}
        self._ext_labels: dict[str, QLabel] = {}
        self._ext_rows: list = []
        for key, caption in _CORE_FIELDS:
            val = QLabel("—")
            self._form.addRow(caption, val)
            self._core_labels[key] = val
        for key, caption in _EXTENDED_FIELDS:
            cap_lbl = QLabel(caption)
            val = QLabel("—")
            self._form.addRow(cap_lbl, val)
            self._ext_labels[key] = val
            self._ext_rows.append((cap_lbl, val))

        self._scroll.setVisible(False)
        self._on_more_toggled(False)

    def _on_more_toggled(self, on):
        for cap_lbl, val in self._ext_rows:
            cap_lbl.setVisible(on)
            val.setVisible(on)

    def show_empty(self, message):
        self._empty_lbl.setText(message)
        self._empty_lbl.setVisible(True)
        self._scroll.setVisible(False)

    def show_props(self, sp):
        self._empty_lbl.setVisible(False)
        self._scroll.setVisible(True)
        for key, lbl in self._core_labels.items():
            lbl.setText(_fmt(getattr(sp, key, None)))
        for key, lbl in self._ext_labels.items():
            lbl.setText(_fmt(getattr(sp, key, None)))
