"""A toolbar that fits any window shape by scrolling horizontally instead of
overflowing behind Qt's default chevron.

A bare QToolBar can't be dropped into a QScrollArea, so this replaces it with
a plain QWidget row of QToolButtons inside a QScrollArea — the content
widget's natural sizeHint (sum of its children's widths) makes the
horizontal scrollbar appear automatically once the window narrows, in both
landscape and portrait, with no manual wrap/measurement logic needed.
"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLayout, QScrollArea, QWidget


class HScrollToolbar(QScrollArea):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("MainToolbarArea")
        self.setWidgetResizable(False)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setFrameShape(QFrame.NoFrame)
        self.setFixedHeight(44)

        self._content = QWidget()
        self._content.setObjectName("MainToolbar")
        self._row = QHBoxLayout(self._content)
        self._row.setContentsMargins(8, 4, 8, 4)
        self._row.setSpacing(4)
        # SetMinAndMaxSize forces the content widget to size itself exactly
        # to its layout's natural sizeHint (sum of children's widths) rather
        # than being squeezed to the viewport — without this, a
        # non-resizable QScrollArea child never actually grows past the
        # viewport and the horizontal scrollbar never appears.
        self._row.setSizeConstraint(QLayout.SetMinAndMaxSize)
        self.setWidget(self._content)

    def add_widget(self, w):
        self._row.addWidget(w)

    def add_separator(self):
        line = QFrame()
        line.setObjectName("TBSep")
        line.setFrameShape(QFrame.VLine)
        self.add_widget(line)

    def wheelEvent(self, ev):
        delta = ev.angleDelta()
        if delta.x() == 0 and delta.y() != 0:
            bar = self.horizontalScrollBar()
            bar.setValue(bar.value() - delta.y())
            ev.accept()
        else:
            super().wheelEvent(ev)
