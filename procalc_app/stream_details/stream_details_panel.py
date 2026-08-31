"""Bottom-left panel of the Circuit Builder screen: one compact card per
distinct stream referenced anywhere in the CURRENTLY selected circuit
(not just the single grid row last clicked) — Stream ID + phase + T/P/mass
flow, each value's unit independently clickable via UnitValueLabel.

Resolution goes through engine_api.resolve_stream_for_snapshot() when a
live HMB is loaded, or the .calc-restored stream snapshot otherwise — both
paths raise/return None on a miss rather than silently showing stale data.
"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QFrame, QHBoxLayout, QLabel, QScrollArea,
                               QVBoxLayout, QWidget)

import engine_api as api
from widgets.unit_field import UnitValueLabel

# Shared fixed caption/value widths so every UnitValueLabel row in a card
# lines its unit dropdown up in one column, regardless of caption length
# ("T" vs "Liquid") -- see UnitValueLabel's caption_width/value_width.
_CAP_W, _VAL_W = 46, 56
_CARD_W = 210


class _StreamCard(QFrame):
    def __init__(self, name, sp, parent=None):
        super().__init__(parent)
        self.setObjectName("Card")
        self.setFixedWidth(_CARD_W)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 10, 12, 10)
        lay.setSpacing(3)

        title = QLabel(name)
        title.setObjectName("CardTitle")
        lay.addWidget(title)

        if sp is None:
            unresolved = QLabel("Could not resolve this stream.")
            unresolved.setObjectName("Muted")
            unresolved.setWordWrap(True)
            lay.addWidget(unresolved)
            return

        phase_lbl = QLabel(sp.phase or "—")
        phase_lbl.setObjectName("Muted")
        lay.addWidget(phase_lbl)

        lay.addWidget(UnitValueLabel("T", sp.temp_f, api.internal_unit("T"),
                                     caption="T", decimals=1,
                                     caption_width=_CAP_W, value_width=_VAL_W))
        lay.addWidget(UnitValueLabel("P", sp.pres_psia, api.internal_unit("P"),
                                     caption="P", decimals=2,
                                     caption_width=_CAP_W, value_width=_VAL_W))
        lay.addWidget(UnitValueLabel("mflow", sp.total_mass, api.internal_unit("mflow"),
                                     caption="Total", decimals=0,
                                     caption_width=_CAP_W, value_width=_VAL_W))
        lay.addWidget(UnitValueLabel("mflow", sp.vap_mass, api.internal_unit("mflow"),
                                     caption="Vapor", decimals=0,
                                     caption_width=_CAP_W, value_width=_VAL_W))
        lay.addWidget(UnitValueLabel("mflow", sp.liq_mass, api.internal_unit("mflow"),
                                     caption="Liquid", decimals=0,
                                     caption_width=_CAP_W, value_width=_VAL_W))


class CircuitStreamsPanel(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(12, 8, 12, 8)
        outer.setSpacing(6)

        title = QLabel("Streams in this circuit")
        title.setObjectName("SectionTitle")
        outer.addWidget(title)

        self._empty_lbl = QLabel(
            "Select a circuit with a Stream Lookup value to see its streams here.")
        self._empty_lbl.setObjectName("Muted")
        self._empty_lbl.setWordWrap(True)
        outer.addWidget(self._empty_lbl)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        host = QWidget()
        self._row = QHBoxLayout(host)
        self._row.setContentsMargins(0, 0, 4, 0)
        self._row.setSpacing(10)
        self._row.addStretch(1)
        scroll.setWidget(host)
        self._scroll = scroll
        outer.addWidget(scroll, 1)
        self._scroll.setVisible(False)

    def show_empty(self, message):
        self._clear_cards()
        self._empty_lbl.setText(message)
        self._empty_lbl.setVisible(True)
        self._scroll.setVisible(False)

    def refresh(self, streams: list[tuple[str, object]]):
        """`streams`: [(stream_name, StreamProps|None), ...] for every
        distinct stream referenced by the current circuit."""
        self._clear_cards()
        if not streams:
            self.show_empty("No streams used in this circuit yet.")
            return
        self._empty_lbl.setVisible(False)
        self._scroll.setVisible(True)
        for name, sp in streams:
            card = _StreamCard(name, sp)
            self._row.insertWidget(self._row.count() - 1, card)

    def _clear_cards(self):
        while self._row.count() > 1:
            item = self._row.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
