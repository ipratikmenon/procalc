"""Qt stylesheet derived from the Linear marketing design-system spec, with
the canvas inverted to white and the accent swapped to T.EN light blue
(per explicit instructions: "only white bd instead of black, strict
implementation", then "use light blue instead of lavender" — everything
else, i.e. the ink/surface ladder, the 4px spacing scale, the radius scale
and the type scale, follows the spec as given).

Token source: Linear's documented system, with the accent re-colored —
  primary #0070ef (T.EN blue, was Linear's lavender #5e6ad2) /
    hover #0059c1 / focus #0064d6
  radius xs4 sm6 md8 lg12 xl16 xxl24 pill/full 9999
  spacing base 4px: xxs4 xs8 sm12 md16 lg24 xl32 xxl48
  type: display-xl 80/600, display-lg 56/600, display-md 40/600,
        headline 28/600, card-title 22/500, subhead 20/400, body-lg 18/400,
        body 16/400, body-sm 14/400, caption 12/400, button 14/500,
        eyebrow 13/500(+0.4 tracking), mono 13/400
Since the spec is a *dark* canvas (#010102) with light ink, every canvas /
surface / ink token below is the light-mode mirror of the spec's dark
value; the accent, success and semantic colors are literal (they don't
depend on canvas darkness) — the accent hover/focus shades are darkened
rather than lightened, since the spec's lighter-on-hover only reads
correctly against a dark canvas. Font stack substitutes Inter for
the proprietary Linear Display/Text faces (the spec's own recommendation);
JetBrains Mono substitutes Linear Mono. Qt Style Sheets have no
letter-spacing property, so the spec's negative tracking is applied via
QFont.setLetterSpacing() in app.py for the couple of labels that carry it,
not here.

The engine's output-workbook theme (Excel/PDF branding) is separate and
unchanged — this file only styles the Qt app chrome.
"""

# hex, with leading '#'
TEN = {
    # accent — reserved for primary CTA / focus ring / active-tab emphasis
    # (Linear's own rule: never used for section titles or body text)
    "blue":   "#0070EF",   # primary (T.EN light blue)
    "navy":   "#0D0E10",   # ink — main text / section-title emphasis
    "teal":   "#3D98B7",
    "green":  "#80C7A0",
    "lime":   "#A2C61C",
    "amber":  "#FDC300",
    "salmon": "#EE7766",
    "red":    "#E84242",   # FAIL / alarm
    "gray":   "#6B6E76",   # ink-subtle (muted text)
    "lgray":  "#E4E5E8",   # hairline

    # light surface ladder (canvas -> surface-1 -> surface-2 -> surface-3)
    "bg":       "#FFFFFF",  # canvas
    "panel":    "#F7F8F8",  # surface-1
    "card":     "#FFFFFF",
    "surface2": "#F1F2F4",  # surface-2 — hovered/lifted rows, sub-nav
    "surface3": "#E9EAEC",  # surface-3 — dropdown/menu
    "border":   "#E4E5E8",  # hairline
    "border2":  "#D3D5DA",  # hairline-strong
    "input":    "#D9DBE0",
    "hover":    "#EEF0FB",  # faint lavender tint — scarce, active/checked only
    "sel":      "#5E6AD2",
    "grid":     "#EEEFF1",
    "alt":      "#FAFAFB",  # zebra
    "text":       "#0D0E10",  # ink
    "text_muted": "#40434A",  # ink-muted
    "text_subtle":"#6B6E76",  # ink-subtle
    "primary_hover": "#0059C1",  # darker on hover — reads correctly on white
    "primary_focus": "#0064D6",
}

# radius scale (literal)
R_MD, R_LG, R_PILL = "8px", "12px", "9999px"
# spacing scale (literal, 4px base)
SP_XXS, SP_XS, SP_SM, SP_MD, SP_LG = "4px", "8px", "12px", "16px", "24px"


def qss() -> str:
    t = TEN
    return f"""
    * {{ font-family: 'Inter','Segoe UI','Noto Sans',system-ui,Arial;
        letter-spacing: 0; }}
    QMainWindow, QWidget, QDialog {{ background: {t['bg']}; color: {t['text']};
        font-size: 14px; }}

    /* toolbar — white top-nav, ghost buttons (ink text, lavender reserved
       for the primary Run action and the active/checked state). A scroll
       area (HScrollToolbar) replaces the native QToolBar so the bar can
       scroll horizontally instead of overflowing behind Qt's chevron. */
    QScrollArea#MainToolbarArea {{ background: {t['bg']}; border: 0;
        border-bottom: 1px solid {t['border']}; }}
    QWidget#MainToolbar {{ background: {t['bg']}; }}
    QFrame#TBSep {{ background: {t['border']}; max-width: 1px; min-width: 1px;
        margin: {SP_XXS} {SP_XS}; }}
    QWidget#MainToolbar QToolButton {{ color: {t['text']}; background: transparent;
        padding: {SP_XS} {SP_SM}; border-radius: {R_MD}; font-weight: 500;
        font-size: 13px; }}
    QWidget#MainToolbar QToolButton:hover {{ background: {t['surface2']}; }}
    QWidget#MainToolbar QToolButton:checked {{ background: {t['hover']}; color: {t['blue']}; }}
    QWidget#MainToolbar QToolButton:disabled {{ color: {t['text_subtle']}; }}
    QToolButton#RunBtn {{ background: {t['blue']}; color: white; font-weight: 600;
        padding: {SP_XS} {SP_MD}; }}
    QToolButton#RunBtn:hover {{ background: {t['primary_hover']}; }}

    /* labels */
    QLabel#Brand {{ color: {t['text']}; font-weight: 700; font-size: 16px; }}
    QLabel#SectionTitle {{ color: {t['text']}; font-weight: 600; font-size: 14px; }}
    QLabel#Muted {{ color: {t['text_subtle']}; font-size: 12px; }}
    QLabel#StateBadge {{ background: {t['surface2']}; color: {t['text_muted']};
        font-size: 12px; border-radius: {R_PILL}; padding: 2px {SP_XS}; }}

    /* tables — light header, hairline grid, generous row padding */
    QHeaderView::section {{ background: {t['panel']}; color: {t['text']};
        padding: {SP_XS} {SP_SM}; border: 0; border-right: 1px solid {t['border']};
        border-bottom: 1px solid {t['border2']}; font-weight: 600; font-size: 13px; }}
    QTableView, QTableWidget {{ background: {t['card']};
        gridline-color: {t['grid']}; alternate-background-color: {t['alt']};
        selection-background-color: {t['sel']}; selection-color: white;
        border: 1px solid {t['border']}; border-radius: {R_LG}; }}
    QTableView::item, QTableWidget::item {{ padding: {SP_XXS} {SP_XS}; }}

    /* tabs — minimal, lavender underline on active (Linear's "link emphasis") */
    QTabWidget::pane {{ border: 1px solid {t['border']}; border-radius: {R_LG};
        background: {t['card']}; top: -1px; }}
    QTabBar::tab {{ background: transparent; color: {t['text_subtle']};
        padding: {SP_SM} {SP_MD}; margin-right: 2px; border: 0;
        font-size: 13px; border-bottom: 2px solid transparent; }}
    QTabBar::tab:hover {{ color: {t['text']}; }}
    QTabBar::tab:selected {{ color: {t['blue']}; font-weight: 600;
        border-bottom: 2px solid {t['blue']}; }}

    /* log / plain text */
    QPlainTextEdit, QTextEdit {{ background: {t['panel']}; color: {t['text']};
        font-family: 'JetBrains Mono','Consolas','DejaVu Sans Mono',monospace;
        font-size: 13px; border: 1px solid {t['border']}; border-radius: {R_LG}; }}

    /* buttons */
    QPushButton {{ background: {t['blue']}; color: white; border: 0;
        padding: {SP_XS} {SP_MD}; border-radius: {R_MD}; font-weight: 500;
        font-size: 14px; }}
    QPushButton:hover {{ background: {t['primary_hover']}; }}
    QPushButton:disabled {{ background: #C7CAE8; }}
    QPushButton#Secondary {{ background: {t['panel']}; color: {t['text']};
        border: 1px solid {t['border']}; }}
    QPushButton#Secondary:hover {{ background: {t['surface2']}; }}

    /* inputs */
    QComboBox, QLineEdit, QDoubleSpinBox, QSpinBox {{ background: {t['card']};
        border: 1px solid {t['input']}; border-radius: {R_MD}; padding: {SP_XS} {SP_SM};
        min-height: 22px; font-size: 14px; }}
    QComboBox:focus, QLineEdit:focus, QDoubleSpinBox:focus, QSpinBox:focus {{
        border: 1px solid {t['blue']}; }}
    QComboBox::drop-down {{ border: 0; width: 20px; }}

    /* group cards */
    QGroupBox {{ font-weight: 600; font-size: 13px; border: 1px solid {t['border']};
        border-radius: {R_LG}; margin-top: 14px; padding: {SP_MD} {SP_SM} {SP_SM} {SP_SM};
        background: {t['card']}; }}
    QGroupBox::title {{ subcontrol-origin: margin; left: {SP_SM}; padding: 0 {SP_XXS};
        color: {t['text']}; }}

    /* dashboard / PMS card tiles */
    QFrame#Card {{ background: {t['card']}; border: 1px solid {t['border']};
        border-radius: {R_LG}; }}
    QFrame#Card:hover {{ border: 1px solid {t['border2']}; }}
    QLabel#CardTitle {{ color: {t['text']}; font-weight: 600; font-size: 14px; }}

    /* misc */
    QStatusBar {{ background: {t['bg']}; border-top: 1px solid {t['border']};
        font-size: 12px; color: {t['text_subtle']}; }}
    QSplitter::handle {{ background: {t['border']}; }}
    QScrollBar:vertical {{ background: {t['bg']}; width: 12px; margin: 0; }}
    QScrollBar::handle:vertical {{ background: {t['border2']}; border-radius: 6px;
        min-height: 24px; }}
    QScrollBar::handle:vertical:hover {{ background: {t['text_subtle']}; }}
    QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; }}
    QScrollBar:horizontal {{ background: {t['bg']}; height: 12px; }}
    QScrollBar::handle:horizontal {{ background: {t['border2']}; border-radius: 6px;
        min-width: 24px; }}
    """
