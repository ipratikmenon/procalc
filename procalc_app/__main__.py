"""Entry point: python -m procalc_app  (or the packaged exe target)."""
from __future__ import annotations

import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from PySide6.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve   # noqa: E402
from PySide6.QtGui import QPixmap, QPainter, QColor                       # noqa: E402
from PySide6.QtWidgets import (QApplication, QSplashScreen,               # noqa: E402
                               QGraphicsOpacityEffect)

_SPLASH_W, _SPLASH_H = 480, 320
_FADE_IN_MS = 700
_HOLD_MS = 900
_FADE_OUT_MS = 500


def _splash_pixmap(logo_path: str) -> QPixmap:
    """A plain white canvas with the T.EN logo centered — the splash screen
    doesn't need its own asset, just the app's existing logo scaled up."""
    canvas = QPixmap(_SPLASH_W, _SPLASH_H)
    canvas.fill(QColor("#FFFFFF"))
    if os.path.exists(logo_path):
        logo = QPixmap(logo_path).scaledToWidth(220, Qt.SmoothTransformation)
        p = QPainter(canvas)
        x = (_SPLASH_W - logo.width()) // 2
        y = (_SPLASH_H - logo.height()) // 2
        p.drawPixmap(x, y, logo)
        p.end()
    return canvas


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Procalc Hydraulics")
    from resources import theme
    app.setStyleSheet(theme.qss())
    import engine_api as api

    splash = QSplashScreen(_splash_pixmap(api.LOGO_PATH))
    splash.setWindowFlag(Qt.FramelessWindowHint)
    effect = QGraphicsOpacityEffect(splash)
    splash.setGraphicsEffect(effect)
    fade_in = QPropertyAnimation(effect, b"opacity", splash)
    fade_in.setStartValue(0.0)
    fade_in.setEndValue(1.0)
    fade_in.setDuration(_FADE_IN_MS)
    fade_in.setEasingCurve(QEasingCurve.OutCubic)
    splash.show()
    fade_in.start()

    def _show_main():
        from app import MainWindow
        win = MainWindow()
        win.show()
        app._main_window = win   # keep a live reference

        fade_out = QPropertyAnimation(effect, b"opacity", splash)
        fade_out.setStartValue(1.0)
        fade_out.setEndValue(0.0)
        fade_out.setDuration(_FADE_OUT_MS)
        fade_out.setEasingCurve(QEasingCurve.InCubic)
        fade_out.finished.connect(splash.close)
        app._splash_fade_out = fade_out   # keep alive until it finishes
        fade_out.start()

    QTimer.singleShot(_FADE_IN_MS + _HOLD_MS, _show_main)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
