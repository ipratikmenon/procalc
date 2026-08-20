"""T.EN brand palette + a Qt stylesheet, mirroring the engine output theme."""

# hex, no leading '#'
TEN = {
    "blue":   "#0070EF",
    "navy":   "#004C84",
    "teal":   "#3D98B7",
    "green":  "#80C7A0",
    "lime":   "#A2C61C",
    "amber":  "#FDC300",
    "salmon": "#EE7766",
    "red":    "#E84242",
    "gray":   "#878787",
    "lgray":  "#DEDEDE",
    "bg":     "#F4F7FA",
    "panel":  "#FFFFFF",
    "text":   "#20303A",
}


def qss() -> str:
    t = TEN
    return f"""
    QMainWindow, QWidget {{ background: {t['bg']}; color: {t['text']};
        font-family: 'Segoe UI', 'Noto Sans', Arial; font-size: 12px; }}
    QToolBar {{ background: {t['navy']}; spacing: 6px; padding: 5px; border: 0; }}
    QToolBar QToolButton {{ color: white; background: transparent;
        padding: 6px 12px; border-radius: 5px; font-weight: 600; }}
    QToolBar QToolButton:hover {{ background: {t['blue']}; }}
    QToolBar QToolButton:checked {{ background: {t['teal']}; }}
    QToolBar QToolButton:disabled {{ color: #9fb6c9; }}
    QLabel#Brand {{ color: {t['navy']}; font-weight: 700; font-size: 13px; }}
    QHeaderView::section {{ background: {t['navy']}; color: white;
        padding: 5px 8px; border: 0; border-right: 1px solid #ffffff33;
        font-weight: 600; }}
    QTableView {{ background: {t['panel']}; gridline-color: {t['lgray']};
        selection-background-color: {t['blue']}; selection-color: white;
        alternate-background-color: #F0F4F8; }}
    QTabWidget::pane {{ border: 1px solid {t['lgray']}; background: {t['panel']}; }}
    QTabBar::tab {{ background: {t['lgray']}; color: {t['text']};
        padding: 6px 14px; border-top-left-radius: 5px; border-top-right-radius: 5px; }}
    QTabBar::tab:selected {{ background: {t['blue']}; color: white; }}
    QPlainTextEdit, QTextEdit {{ background: #0e1b26; color: #cfe6ff;
        font-family: 'Consolas','DejaVu Sans Mono',monospace; font-size: 11px;
        border: 1px solid {t['navy']}; }}
    QPushButton {{ background: {t['blue']}; color: white; border: 0;
        padding: 6px 14px; border-radius: 5px; font-weight: 600; }}
    QPushButton:hover {{ background: {t['navy']}; }}
    QPushButton:disabled {{ background: {t['gray']}; }}
    QComboBox, QLineEdit {{ background: white; border: 1px solid {t['gray']};
        border-radius: 4px; padding: 3px 6px; }}
    QStatusBar {{ background: {t['panel']}; border-top: 1px solid {t['lgray']}; }}
    QGroupBox {{ font-weight: 600; border: 1px solid {t['lgray']};
        border-radius: 6px; margin-top: 8px; padding-top: 8px; }}
    QGroupBox::title {{ subcontrol-origin: margin; left: 10px; color: {t['navy']}; }}
    """
