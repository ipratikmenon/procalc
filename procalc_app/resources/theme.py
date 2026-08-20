"""T.EN brand palette + a light "shadcn/bag-ui"-style Qt stylesheet.

White surfaces, hairline borders, ~8-10px radius, generous padding, an
Inter/system font stack, ghost toolbar buttons and minimal tabs — the layout
& typography of bag-ui, colored with the T.EN palette (accents only).
The engine's output-workbook theme is separate and unchanged.
"""

# hex, with leading '#'
TEN = {
    # brand accents
    "blue":   "#0070EF",   # primary
    "navy":   "#004C84",   # emphasis / headings / hover
    "teal":   "#3D98B7",
    "green":  "#80C7A0",
    "lime":   "#A2C61C",
    "amber":  "#FDC300",
    "salmon": "#EE7766",
    "red":    "#E84242",   # FAIL / alarm
    "gray":   "#6B7A87",   # muted text
    "lgray":  "#DEDEDE",
    # light surfaces / structure
    "bg":     "#FFFFFF",   # window
    "panel":  "#F7F9FB",   # muted panel
    "card":   "#FFFFFF",
    "border": "#E6E9EF",   # hairline
    "input":  "#D8DEE6",   # input border
    "hover":  "#EAF3FE",   # soft blue hover tint
    "sel":    "#0070EF",
    "grid":   "#EDF0F3",
    "alt":    "#FAFBFC",   # zebra
    "text":   "#1B2A33",
}


def qss() -> str:
    t = TEN
    return f"""
    * {{ font-family: 'Inter','Segoe UI','Noto Sans',system-ui,Arial; }}
    QMainWindow, QWidget, QDialog {{ background: {t['bg']}; color: {t['text']};
        font-size: 12.5px; }}

    /* toolbar — white bar, ghost buttons */
    QToolBar {{ background: {t['bg']}; spacing: 4px; padding: 6px 8px;
        border: 0; border-bottom: 1px solid {t['border']}; }}
    QToolBar::separator {{ background: {t['border']}; width: 1px; margin: 4px 6px; }}
    QToolBar QToolButton {{ color: {t['blue']}; background: transparent;
        padding: 6px 12px; border-radius: 8px; font-weight: 600; }}
    QToolBar QToolButton:hover {{ background: {t['hover']}; }}
    QToolBar QToolButton:checked {{ background: #DCEBFB; color: {t['navy']}; }}
    QToolBar QToolButton:disabled {{ color: #A9B6C0; }}
    QToolButton#RunBtn {{ background: {t['blue']}; color: white; }}
    QToolButton#RunBtn:hover {{ background: {t['navy']}; }}

    /* labels */
    QLabel#Brand {{ color: {t['navy']}; font-weight: 700; font-size: 13px; }}

    /* tables — light header, hairline grid */
    QHeaderView::section {{ background: {t['hover']}; color: {t['navy']};
        padding: 5px 8px; border: 0; border-right: 1px solid {t['border']};
        border-bottom: 1px solid {t['border']}; font-weight: 600; }}
    QTableView, QTableWidget {{ background: {t['card']};
        gridline-color: {t['grid']}; alternate-background-color: {t['alt']};
        selection-background-color: {t['sel']}; selection-color: white;
        border: 1px solid {t['border']}; border-radius: 8px; }}
    QTableView::item, QTableWidget::item {{ padding: 2px 4px; }}

    /* tabs — minimal, blue underline on active */
    QTabWidget::pane {{ border: 1px solid {t['border']}; border-radius: 8px;
        background: {t['card']}; top: -1px; }}
    QTabBar::tab {{ background: transparent; color: {t['gray']};
        padding: 7px 14px; margin-right: 2px; border: 0;
        border-bottom: 2px solid transparent; }}
    QTabBar::tab:hover {{ color: {t['navy']}; }}
    QTabBar::tab:selected {{ color: {t['blue']}; font-weight: 600;
        border-bottom: 2px solid {t['blue']}; }}

    /* log / plain text — light */
    QPlainTextEdit, QTextEdit {{ background: {t['panel']}; color: #24333D;
        font-family: 'Consolas','DejaVu Sans Mono',monospace; font-size: 11px;
        border: 1px solid {t['border']}; border-radius: 8px; }}

    /* buttons */
    QPushButton {{ background: {t['blue']}; color: white; border: 0;
        padding: 7px 14px; border-radius: 8px; font-weight: 600; }}
    QPushButton:hover {{ background: {t['navy']}; }}
    QPushButton:disabled {{ background: #AEB9C2; }}
    QPushButton#Secondary {{ background: transparent; color: {t['blue']};
        border: 1px solid {t['input']}; }}
    QPushButton#Secondary:hover {{ background: {t['hover']}; }}

    /* inputs */
    QComboBox, QLineEdit, QDoubleSpinBox, QSpinBox {{ background: {t['card']};
        border: 1px solid {t['input']}; border-radius: 8px; padding: 4px 8px;
        min-height: 20px; }}
    QComboBox:focus, QLineEdit:focus, QDoubleSpinBox:focus, QSpinBox:focus {{
        border: 1px solid {t['blue']}; }}
    QComboBox::drop-down {{ border: 0; width: 18px; }}

    /* group cards */
    QGroupBox {{ font-weight: 600; border: 1px solid {t['border']};
        border-radius: 10px; margin-top: 10px; padding: 10px 8px 8px 8px;
        background: {t['card']}; }}
    QGroupBox::title {{ subcontrol-origin: margin; left: 12px; padding: 0 4px;
        color: {t['navy']}; }}

    /* misc */
    QStatusBar {{ background: {t['bg']}; border-top: 1px solid {t['border']}; }}
    QSplitter::handle {{ background: {t['border']}; }}
    QScrollBar:vertical {{ background: {t['bg']}; width: 12px; margin: 0; }}
    QScrollBar::handle:vertical {{ background: #C7D0DA; border-radius: 6px;
        min-height: 24px; }}
    QScrollBar::handle:vertical:hover {{ background: {t['gray']}; }}
    QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; }}
    QScrollBar:horizontal {{ background: {t['bg']}; height: 12px; }}
    QScrollBar::handle:horizontal {{ background: #C7D0DA; border-radius: 6px;
        min-width: 24px; }}
    """
