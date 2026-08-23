"""Entry point: python -m procalc_app  (or the packaged exe target)."""
from __future__ import annotations

import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from PySide6.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve   # noqa: E402
from PySide6.QtGui import QPixmap, QPainter, QColor, QFont                # noqa: E402
from PySide6.QtWidgets import (QApplication, QSplashScreen,               # noqa: E402
                               QGraphicsOpacityEffect)

_SPLASH_W, _SPLASH_H = 480, 360
_FADE_IN_MS = 700
_HOLD_MS = 900
_FADE_OUT_MS = 500


def _splash_pixmap(procalc_logo_path: str, ten_logo_path: str) -> QPixmap:
    """ProCalc's own mark + wordmark, prominent; the T.EN logo demoted to a
    small confidentiality footer. Falls back gracefully — a missing asset
    just leaves that region blank — so this renders correctly whether or
    not procalc_logo_path has actually been supplied yet."""
    from resources import theme

    canvas = QPixmap(_SPLASH_W, _SPLASH_H)
    canvas.fill(QColor(theme.TEN["bg"]))
    p = QPainter(canvas)
    p.setRenderHint(QPainter.Antialiasing)

    # ProCalc logo, centered
    if os.path.exists(procalc_logo_path):
        logo = QPixmap(procalc_logo_path).scaledToWidth(160, Qt.SmoothTransformation)
        x = (_SPLASH_W - logo.width()) // 2
        p.drawPixmap(x, 55, logo)

    # "ProCalc" wordmark
    word_font = QFont("Inter", 22, QFont.Bold)
    p.setFont(word_font)
    p.setPen(QColor(theme.TEN["navy"]))
    p.drawText(0, 205, _SPLASH_W, 34, Qt.AlignHCenter, "ProCalc")

    # hairline separator
    p.setPen(QColor(theme.TEN["border"]))
    p.drawLine(60, 255, _SPLASH_W - 60, 255)

    # bottom band: T.EN logo (small) + confidentiality caption
    if os.path.exists(ten_logo_path):
        ten_logo = QPixmap(ten_logo_path).scaledToWidth(70, Qt.SmoothTransformation)
        p.drawPixmap((_SPLASH_W - ten_logo.width()) // 2, 272, ten_logo)

    cap_font = QFont("Inter", 9)
    p.setFont(cap_font)
    p.setPen(QColor(theme.TEN["text_subtle"]))
    p.drawText(0, 320, _SPLASH_W, 20, Qt.AlignHCenter,
              "Confidential — for internal T.EN use only")

    p.end()
    return canvas


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Procalc Hydraulics")
    from resources import theme
    app.setStyleSheet(theme.qss())
    import engine_api as api

    splash = QSplashScreen(_splash_pixmap(api.PROCALC_LOGO_PATH, api.LOGO_PATH))
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
