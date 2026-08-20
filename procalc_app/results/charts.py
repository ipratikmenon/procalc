"""QtCharts re-plots of the two engine chart-sheet types.

- ``pressure_profile_chart``: P_out vs cumulative length from a Pressure_Profile
  worksheet.
- ``flow_pattern_chart``: the regime boundaries and operating point(s) of an
  'H ...' / 'V ...' flow-pattern map sheet.

Both are defensive: they return ``None`` (never raise for expected shapes) when
there is nothing plottable, so the caller can fall back to a table-only tab.
Colours come from the T.EN brand palette (``resources.theme.TEN``).
"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPainter, QPen
from PySide6.QtCharts import (QChart, QChartView, QLineSeries, QScatterSeries,
                              QValueAxis, QLogValueAxis)

from resources.theme import TEN


def _qcolor(key, default="#0070EF"):
    return QColor(TEN.get(key, default))


def _is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _themed_view(chart):
    """Wrap a configured QChart in an antialiased, on-brand QChartView."""
    chart.setBackgroundBrush(QColor(TEN.get("panel", "#FFFFFF")))
    chart.setTitleBrush(_qcolor("navy", "#004C84"))
    tf = chart.titleFont()
    tf.setBold(True)
    tf.setPointSizeF(max(tf.pointSizeF(), 11.0))
    chart.setTitleFont(tf)
    chart.legend().setVisible(True)
    chart.legend().setAlignment(Qt.AlignBottom)
    chart.legend().setLabelColor(_qcolor("text", "#20303A"))
    view = QChartView(chart)
    view.setRenderHint(QPainter.Antialiasing, True)
    view.setBackgroundBrush(QColor(TEN.get("panel", "#FFFFFF")))
    return view


def _style_value_axis(axis):
    axis.setLabelsColor(_qcolor("text", "#20303A"))
    axis.setTitleBrush(_qcolor("navy", "#004C84"))
    axis.setGridLineColor(QColor(TEN.get("lgray", "#DEDEDE")))
    return axis


# --------------------------------------------------------------------------
# Pressure profile
# --------------------------------------------------------------------------
def pressure_profile_chart(ws):
    """QLineSeries of P Out (y) vs Cum Length (x) from a Pressure_Profile sheet.

    Row 2 holds the headers; data begins at row 3.  Merged LINE divider rows
    (all-blank) and the trailing TOTAL row are skipped.  Returns a themed
    QChartView, or None if the needed columns / data cannot be found.
    """
    try:
        ncols = ws.max_column
        x_col = y_col = None
        for c in range(1, ncols + 1):
            h = ws.cell(row=2, column=c).value
            if not isinstance(h, str):
                continue
            hl = h.lower()
            if x_col is None and "cum" in hl:
                x_col = c
            if y_col is None and "p out" in hl:
                y_col = c
        if x_col is None or y_col is None:
            return None

        series = QLineSeries()
        series.setName("P Out (psia)")
        pen = QPen(_qcolor("blue", "#0070EF"))
        pen.setWidth(2)
        series.setPen(pen)
        series.setPointsVisible(True)

        xs, ys = [], []
        for r in range(3, ws.max_row + 1):
            a = ws.cell(row=r, column=1).value
            if isinstance(a, str) and a.strip().upper().startswith("TOTAL"):
                continue
            x = ws.cell(row=r, column=x_col).value
            y = ws.cell(row=r, column=y_col).value
            if not (_is_number(x) and _is_number(y)):
                continue  # blank LINE divider rows land here
            xs.append(float(x))
            ys.append(float(y))
            series.append(float(x), float(y))
        if len(xs) < 2:
            return None

        chart = QChart()
        title = ws.cell(row=1, column=1).value
        chart.setTitle(str(title) if title else "Pressure Profile")
        chart.addSeries(series)

        ax = _style_value_axis(QValueAxis())
        ax.setTitleText("Cumulative Length (ft)")
        ay = _style_value_axis(QValueAxis())
        ay.setTitleText("P Out (psia)")
        chart.addAxis(ax, Qt.AlignBottom)
        chart.addAxis(ay, Qt.AlignLeft)
        series.attachAxis(ax)
        series.attachAxis(ay)
        xmin, xmax = min(xs), max(xs)
        ymin, ymax = min(ys), max(ys)
        pad = (ymax - ymin) * 0.08 or max(abs(ymax) * 0.02, 0.5)
        ax.setRange(xmin, xmax if xmax > xmin else xmin + 1)
        ay.setRange(ymin - pad, ymax + pad)
        return _themed_view(chart)
    except Exception:
        return None


# --------------------------------------------------------------------------
# Flow-pattern map
# --------------------------------------------------------------------------
def _paired_columns(ws):
    """Yield (kind, label, x_col, y_col) for helper columns in the map sheet.

    The engine lays out map data as adjacent X/Y column pairs whose row-1
    headers look like 'bnd0 X'/'bnd0 Y' (regime boundaries) or
    'path X'/'path Y' (the operating path).  We pair a '... X' header with the
    next column iff that column's header is the matching '... Y'.
    """
    ncols = ws.max_column
    for c in range(1, ncols):
        h = ws.cell(row=1, column=c).value
        h2 = ws.cell(row=1, column=c + 1).value
        if not (isinstance(h, str) and isinstance(h2, str)):
            continue
        hl, hl2 = h.lower(), h2.lower()
        if hl.endswith(" x") and hl2.endswith(" y") and hl[:-2] == hl2[:-2]:
            label = h[:-2].strip()
            kind = "path" if "path" in hl else ("bnd" if "bnd" in hl else "other")
            yield kind, label, c, c + 1


def _read_pair(ws, x_col, y_col):
    xs, ys = [], []
    for r in range(2, ws.max_row + 1):
        x = ws.cell(row=r, column=x_col).value
        y = ws.cell(row=r, column=y_col).value
        if _is_number(x) and _is_number(y):
            xs.append(float(x))
            ys.append(float(y))
    return xs, ys


def _use_log(values):
    """Log axis is appropriate when every value is > 0 and spans >2 decades."""
    vals = [v for v in values if v is not None]
    if not vals or any(v <= 0 for v in vals):
        return False
    return (max(vals) / min(vals)) > 100.0


def flow_pattern_chart(ws, data_ws=None):
    """Re-plot an 'H ...' / 'V ...' flow-pattern map.

    The map sheet carries its regime boundaries and (when non-degenerate) its
    operating path as adjacent X/Y helper-column pairs, already in the map's
    own coordinate system -- the reliable source, so it is used first.  The
    operating points are drawn as a QScatterSeries.  If the map sheet has no
    usable helper columns, we fall back to scanning ``data_ws``
    ('Flow_Pattern_Data') for its first two numeric columns.  Returns a themed
    QChartView, or None if nothing is plottable.
    """
    try:
        chart = QChart()
        chart.setTitle(ws.title)

        all_x, all_y = [], []
        boundary_pen = QPen(QColor(TEN.get("gray", "#878787")))
        boundary_pen.setWidth(2)
        bnd_labelled = False
        op_labelled = False
        n_series = 0

        for kind, label, xc, yc in _paired_columns(ws):
            xs, ys = _read_pair(ws, xc, yc)
            if len(xs) < 1:
                continue
            if kind == "path":
                if len(xs) < 1:
                    continue
                s = QScatterSeries()
                s.setName("operating point" if not op_labelled else "")
                s.setColor(_qcolor("blue", "#0070EF"))
                s.setBorderColor(_qcolor("navy", "#004C84"))
                s.setMarkerSize(11.0)
                op_labelled = True
            else:
                if len(xs) < 2:
                    continue
                s = QLineSeries()
                s.setName("regime boundary" if not bnd_labelled else "")
                s.setPen(boundary_pen)
                bnd_labelled = True
            for x, y in zip(xs, ys):
                s.append(x, y)
            chart.addSeries(s)
            all_x.extend(xs)
            all_y.extend(ys)
            n_series += 1

        # Fallback: operating points from the plain Flow_Pattern_Data table.
        if n_series == 0 and data_ws is not None:
            hdr_row = 2
            num_cols = []
            for c in range(1, data_ws.max_column + 1):
                col_vals = [data_ws.cell(row=r, column=c).value
                            for r in range(hdr_row + 1, data_ws.max_row + 1)]
                if any(_is_number(v) for v in col_vals):
                    num_cols.append((c, data_ws.cell(row=hdr_row, column=c).value))
            if len(num_cols) >= 2:
                xc, xhdr = num_cols[0]
                yc, yhdr = num_cols[1]
                s = QScatterSeries()
                s.setName("operating point")
                s.setColor(_qcolor("blue", "#0070EF"))
                s.setBorderColor(_qcolor("navy", "#004C84"))
                s.setMarkerSize(11.0)
                for r in range(hdr_row + 1, data_ws.max_row + 1):
                    x = data_ws.cell(row=r, column=xc).value
                    y = data_ws.cell(row=r, column=yc).value
                    if _is_number(x) and _is_number(y):
                        s.append(float(x), float(y))
                        all_x.append(float(x))
                        all_y.append(float(y))
                if s.count() > 0:
                    chart.addSeries(s)
                    n_series += 1
                    ax_title = str(xhdr) if xhdr else "x"
                    ay_title = str(yhdr) if yhdr else "y"
                else:
                    return None
            else:
                return None

        if n_series == 0 or not all_x:
            return None

        # Axes -- log when the data spans several decades (typical for maps).
        ax = QLogValueAxis() if _use_log(all_x) else QValueAxis()
        ay = QLogValueAxis() if _use_log(all_y) else QValueAxis()
        for a in (ax, ay):
            a.setLabelsColor(_qcolor("text", "#20303A"))
            a.setTitleBrush(_qcolor("navy", "#004C84"))
            a.setGridLineColor(QColor(TEN.get("lgray", "#DEDEDE")))
        chart.addAxis(ax, Qt.AlignBottom)
        chart.addAxis(ay, Qt.AlignLeft)
        for s in chart.series():
            s.attachAxis(ax)
            s.attachAxis(ay)
        return _themed_view(chart)
    except Exception:
        return None
