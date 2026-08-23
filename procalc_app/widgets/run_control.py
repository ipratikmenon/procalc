"""Combined Run + Auto control: a one-shot Run button next to a green
SlideToggle for continuous Auto-run.

The Run button is visible only while Auto is off — turning Auto on already
performs an implied first run (RunController.schedule() fires shortly
after), so a separate one-shot affordance only makes sense while Auto-run
isn't already going to fire on its own.
"""
from __future__ import annotations

from PySide6.QtWidgets import QHBoxLayout, QLabel, QToolButton, QWidget

from widgets.slide_toggle import SlideToggle


class RunAutoControl(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)

        self.run_btn = QToolButton()
        self.run_btn.setText("▶ Run")
        self.run_btn.setObjectName("RunBtn")
        self.run_btn.setToolTip("Run once")

        self.toggle = SlideToggle()
        self.toggle.setToolTip("Auto-run on every edit")

        self.auto_lbl = QLabel("Auto")

        lay.addWidget(self.run_btn)
        lay.addWidget(self.toggle)
        lay.addWidget(self.auto_lbl)

        self.toggle.toggled.connect(self._on_toggle)

    def _on_toggle(self, on):
        self.run_btn.setVisible(not on)
