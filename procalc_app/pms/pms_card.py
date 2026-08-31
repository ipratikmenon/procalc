"""Card-tile widget for the PMS Manager's alternate browse view.

Shows the same at-a-glance fields as the list's flat table (name/MOC/
corrosion-allowance/rating) plus a few that don't fit a table column
(nominal rating text, typical service, material group) — the actual
value-add of a card view over the table.
"""
from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QFrame, QLabel, QVBoxLayout

import engine_api as api

# short material code from the full MOC text (KCS / LTCS / SS304 / A20 …)
_MOC_RULES = [
    ("low temp", "LTCS"), ("lt carbon", "LTCS"), ("impact tested", "LTCS"),
    ("killed carbon", "KCS"), ("carbon steel", "CS"), ("carbon stl", "CS"),
    ("304l", "SS304L"), ("304", "SS304"), ("316l", "SS316L"), ("316", "SS316"),
    ("321", "SS321"), ("347", "SS347"), ("317", "SS317"),
    ("duplex", "DSS"), ("2205", "DSS"), ("2507", "SDSS"),
    ("alloy 20", "A20"), ("n08020", "A20"), ("825", "A825"), ("625", "IN625"),
    ("inconel", "INC"), ("incoloy", "INC"), ("monel", "MONEL"),
    ("hastelloy", "HAST"), ("nickel", "NI"),
    ("5cr", "5Cr"), ("5 cr", "5Cr"), ("9cr", "9Cr"), ("9 cr", "9Cr"),
    ("1.25cr", "1¼Cr"), ("2.25cr", "2¼Cr"), ("chrome", "Cr-Mo"),
    ("ductile", "DI"), ("nodular", "DI"), ("cast iron", "CI"),
    ("galvan", "GALV"), ("copper", "Cu"), ("cupro", "CuNi"),
    ("titanium", "Ti"), ("gre", "GRE"), ("frp", "FRP"), ("pvc", "PVC"),
    ("pvdf", "PVDF"), ("ptfe", "PTFE"),
]


def short_moc(moc: str, moc_tag: str = "") -> str:
    s = (moc or "").lower()
    for key, code in _MOC_RULES:
        if key in s:
            return code
    if (" cr" in s or s.strip().endswith("cr")) and "chrome" not in s:
        return "Cr-Mo"
    if "high density" in s or "hdpe" in s:
        return "HDPE"
    if moc_tag and str(moc_tag).strip():
        return str(moc_tag).strip()[:8]
    # fall back to the first word that looks like a material name, skipping
    # extraction noise (numbers, 'allowance', 'note', punctuation)
    first = (moc or "").split(",")[0].strip()
    low = first.lower()
    if not first or first[0].isdigit() or any(
            k in low for k in ("allowance", "note", "n/a", "see ")):
        return "—"
    return first[:10]


class PMSCardWidget(QFrame):
    clicked = Signal(str)   # emits the class name

    def __init__(self, cls, parent=None):
        super().__init__(parent)
        self.setObjectName("Card")
        self._name = cls.name
        self.setCursor(Qt.PointingHandCursor)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 12, 14, 12)
        lay.setSpacing(3)

        title = QLabel(cls.name)
        title.setObjectName("CardTitle")
        lay.addWidget(title)

        moc_lbl = QLabel(short_moc(cls.moc, cls.moc_tag))
        moc_lbl.setObjectName("Muted")
        lay.addWidget(moc_lbl)

        rating = api.flange_class_for(cls.name)
        ca = cls.corrosion_allow_in
        bits = []
        if rating:
            bits.append(f"{rating}#")
        if ca:
            bits.append(f"C.A. {float(ca):.3f} in")
        lay.addWidget(QLabel("  ·  ".join(bits) if bits else "—"))

        if cls.nominal_rating:
            lay.addWidget(QLabel(str(cls.nominal_rating)))
        if cls.material_group:
            lay.addWidget(QLabel(f"Group: {cls.material_group}"))
        if cls.typical_service:
            svc = QLabel(str(cls.typical_service))
            svc.setObjectName("Muted")
            svc.setWordWrap(True)
            lay.addWidget(svc)

    def mousePressEvent(self, ev):
        self.clicked.emit(self._name)
        super().mousePressEvent(ev)
