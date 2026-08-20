"""Entry point: python -m procalc_app  (or the packaged exe target)."""
from __future__ import annotations

import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from PySide6.QtWidgets import QApplication      # noqa: E402


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Procalc Hydraulics")
    from resources import theme
    app.setStyleSheet(theme.qss())
    from app import MainWindow
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
