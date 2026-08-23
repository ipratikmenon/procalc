"""A composite <label> <value> <unit ▾> widget: the unit segment is a
clickable dropdown that live-converts the displayed value between every
unit `common/units.py` knows for that quantity, instead of the unit being
baked into the label text (e.g. "Pressure" + a separate "psia ▾", not one
static "Pressure (psia)" string).
"""
from __future__ import annotations

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QComboBox, QHBoxLayout, QLabel, QWidget

import engine_api as api


def _fmt(value, decimals):
    if value is None:
        return "—"
    try:
        return f"{float(value):,.{decimals}f}"
    except (TypeError, ValueError):
        return str(value)


class UnitValueLabel(QWidget):
    """qty: a common/units.py quantity code (e.g. "P", "T", "mflow").
    value: the raw number, already expressed in `unit`.
    unit: the unit `value` is currently expressed in (must be one of
    api.units_for(qty); defaults to that quantity's FPS unit if omitted).
    """

    unit_changed = Signal(str)   # new unit string

    def __init__(self, qty: str, value=None, unit: str | None = None,
                caption: str | None = None, decimals: int = 4, parent=None):
        super().__init__(parent)
        self._qty = qty
        self._decimals = decimals
        self._value = value
        self._unit = unit

        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(4)

        self._caption_lbl = QLabel(caption or "")
        self._caption_lbl.setObjectName("Muted")
        if caption:
            lay.addWidget(self._caption_lbl)

        self._value_lbl = QLabel()
        lay.addWidget(self._value_lbl)

        self._unit_combo = QComboBox()
        self._unit_combo.setObjectName("UnitPicker")
        units = api.units_for(qty)
        if self._unit and self._unit not in units:
            units = [self._unit] + units
        self._unit_combo.addItems(units)
        if self._unit:
            i = self._unit_combo.findText(self._unit)
            if i >= 0:
                self._unit_combo.setCurrentIndex(i)
        elif units:
            self._unit = units[0]
        self._unit_combo.currentTextChanged.connect(self._on_unit_changed)
        lay.addWidget(self._unit_combo)
        lay.addStretch(1)

        self._render()

    def _render(self):
        self._value_lbl.setText(_fmt(self._value, self._decimals))

    def _on_unit_changed(self, new_unit):
        if not new_unit or new_unit == self._unit:
            return
        if self._value is not None:
            self._value = api.convert(self._qty, self._value, self._unit, new_unit)
        self._unit = new_unit
        self._render()
        self.unit_changed.emit(new_unit)

    def set_value(self, value, unit: str | None = None):
        """Replace the displayed value. If `unit` is given and differs from
        the current unit, the combo is switched (without re-converting —
        the caller is supplying a fresh reading already in that unit)."""
        self._value = value
        if unit and unit != self._unit:
            self._unit = unit
            i = self._unit_combo.findText(unit)
            if i >= 0:
                self._unit_combo.blockSignals(True)
                self._unit_combo.setCurrentIndex(i)
                self._unit_combo.blockSignals(False)
        self._render()

    @property
    def value(self):
        return self._value

    @property
    def unit(self):
        return self._unit
