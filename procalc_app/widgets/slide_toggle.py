"""A green, animated pill-style toggle switch (iOS-style), used where a
binary on/off state should read as a slider rather than a checkable button.

Pure click-to-toggle — no drag gesture. This is a desktop mouse app; the
ask is a slider-*styled* switch, not touch-drag behavior, so a drag gesture
would only add hit-testing/accessibility complexity with no real benefit.
"""
from __future__ import annotations

from PySide6.QtCore import (
    Property, QEasingCurve, QPropertyAnimation, QRectF, Qt,
)
from PySide6.QtGui import QColor, QPainter
from PySide6.QtWidgets import QAbstractButton

from resources import theme

_W, _H = 44, 24
_THUMB_D = 20
_THUMB_MARGIN = 2
_THUMB_ON = float(_W - _THUMB_D - _THUMB_MARGIN)
_THUMB_OFF = float(_THUMB_MARGIN)


class SlideToggle(QAbstractButton):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("SlideToggle")
        self.setCheckable(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedSize(_W, _H)
        self._thumb_pos = _THUMB_OFF
        self._anim = QPropertyAnimation(self, b"thumb_pos", self)
        self._anim.setDuration(150)
        self._anim.setEasingCurve(QEasingCurve.OutCubic)
        self.toggled.connect(self._animate_to)

    def _animate_to(self, checked):
        self._anim.stop()
        self._anim.setStartValue(self._thumb_pos)
        self._anim.setEndValue(_THUMB_ON if checked else _THUMB_OFF)
        self._anim.start()

    def get_thumb_pos(self):
        return self._thumb_pos

    def set_thumb_pos(self, v):
        self._thumb_pos = v
        self.update()

    thumb_pos = Property(float, get_thumb_pos, set_thumb_pos)

    def sizeHint(self):
        return self.size()

    def paintEvent(self, ev):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        track = QColor(theme.TEN["green"] if self.isChecked() else theme.TEN["border2"])
        p.setPen(Qt.NoPen)
        p.setBrush(track)
        p.drawRoundedRect(QRectF(0, 0, _W, _H), _H / 2, _H / 2)
        p.setBrush(QColor("white"))
        p.drawEllipse(QRectF(self._thumb_pos, _THUMB_MARGIN, _THUMB_D, _THUMB_D))
