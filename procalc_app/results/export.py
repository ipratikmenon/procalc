"""Self-contained export helpers for the engine .xlsx workbooks.

The toolbar calls these to save the currently-loaded engine workbook (or a
single sheet of it) either as a styled .xlsx or as a branded, paginated PDF.

Everything here is standalone: the small style/format helpers below are
intentionally NOT imported from results.workbook_model, so this module carries
no coupling to the on-screen renderer. Only openpyxl, PySide6 and the stdlib
are used - no reportlab, no new pip dependencies.

Public API
----------
    list_sheets(src_xlsx) -> list[str]
    export_excel_workbook(src_xlsx, dest_xlsx)
    export_excel_sheet(src_xlsx, sheet_title, dest_xlsx)
    export_pdf_workbook(src_xlsx, dest_pdf, logo_path=None)
    export_pdf_sheet(src_xlsx, sheet_title, dest_pdf, logo_path=None)
"""
from __future__ import annotations

import os
import shutil
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

from PySide6.QtCore import QMarginsF, QRectF, Qt
from PySide6.QtGui import (QColor, QFont, QGuiApplication, QPageLayout,
                           QPageSize, QPainter, QPdfWriter, QPixmap)


# ── brand palette ────────────────────────────────────────────────────────────
NAVY = "#004C84"
GRID = "#C9D2DA"          # thin light-gray cell border
HEADER_TEXT = "#004C84"


# ── style / value helpers (standalone; do not import from workbook_model) ─────
def _argb_to_qcolor(rgb):
    """openpyxl ARGB hex (e.g. '00FFFFFF' or 'FF1F4973') -> QColor, or None."""
    if not rgb or not isinstance(rgb, str):
        return None
    s = rgb[-6:] if len(rgb) >= 6 else rgb
    try:
        return QColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
    except (ValueError, TypeError):
        return None


def _fmt(value, number_format):
    """Format a cell value with a small subset of Excel number formats:
    '0.0', '0.00', '0.000', '#,##0', '0.0%', 'General' (and close relatives)."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        nf = number_format or "General"
        try:
            if nf.endswith("%"):
                dec = 0
                if "." in nf:
                    dec = len(nf.split(".")[-1].rstrip("%_) "))
                dec = min(max(dec, 0), 6)
                return f"{value * 100:.{dec}f}%"
            if "." in nf:
                frac = nf.split(".")[-1]
                dec = len(frac.replace("_", "").replace(")", "").replace("%", "").replace(",", ""))
                dec = min(max(dec, 0), 6)
                return f"{value:,.{dec}f}" if "," in nf else f"{value:.{dec}f}"
            if nf in ("#,##0", "0", "#,##0;-#,##0"):
                return f"{value:,.0f}" if "," in nf else f"{value:.0f}"
        except Exception:
            pass
        if isinstance(value, float):
            return f"{value:.4g}"
        return str(value)
    return str(value)


def _ensure_app():
    """QPdfWriter / QPixmap need a running QGuiApplication."""
    app = QGuiApplication.instance()
    if app is None:
        app = QGuiApplication([])
    return app


# ── public: introspection ─────────────────────────────────────────────────────
def list_sheets(src_xlsx: str) -> list[str]:
    """Return the worksheet titles of the workbook, in order."""
    wb = load_workbook(src_xlsx, data_only=False, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _apply_print_setup(ws):
    """Landscape + 'fit to 1 page wide' print setup, so the sheet prints/
    previews nicely without the user having to set it up by hand."""
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5,
                                  header=0.2, footer=0.2)
    ws.print_options.horizontalCentered = True


# ── public: Excel exports ─────────────────────────────────────────────────────
def export_excel_workbook(src_xlsx: str, dest_xlsx: str) -> str:
    """Copy the styled workbook, with print setup set to fit each sheet to
    one page wide. Falls back to a verbatim byte copy if re-saving through
    openpyxl fails for any reason (fidelity over features)."""
    if not os.path.exists(src_xlsx):
        raise Exception(f"Source workbook not found: {src_xlsx}")
    try:
        wb = load_workbook(src_xlsx)
        for ws in wb.worksheets:
            _apply_print_setup(ws)
        wb.save(dest_xlsx)
    except Exception:
        shutil.copyfile(src_xlsx, dest_xlsx)
    return dest_xlsx


def export_excel_sheet(src_xlsx: str, sheet_title: str, dest_xlsx: str) -> str:
    """Deep-copy a single worksheet (values + styles + geometry) into a new
    single-sheet workbook."""
    if not os.path.exists(src_xlsx):
        raise Exception(f"Source workbook not found: {src_xlsx}")
    src_wb = load_workbook(src_xlsx, data_only=False)
    if sheet_title not in src_wb.sheetnames:
        raise Exception(f"Sheet '{sheet_title}' not found in {src_xlsx}")
    src = src_wb[sheet_title]

    out = Workbook()
    dst = out.active
    dst.title = sheet_title[:31]  # Excel sheet-name limit

    # An openpyxl cell's `_style` is an index-based StyleArray pointing into the
    # workbook's shared style tables.  For `copy(src_cell._style)` to resolve in
    # the new workbook, that workbook must carry the same style tables - so copy
    # the source's shared collections across first.
    for attr in ("_fonts", "_fills", "_borders", "_alignments",
                 "_number_formats", "_protections", "_cell_styles",
                 "_named_styles", "_differential_styles"):
        if hasattr(src_wb, attr) and hasattr(out, attr):
            src_tbl = getattr(src_wb, attr)
            try:
                # openpyxl IndexedList: copy() drops contents, so rebuild it
                # from its elements to preserve the index -> style mapping.
                setattr(out, attr, type(src_tbl)(list(src_tbl)))
            except Exception:
                setattr(out, attr, copy(src_tbl))

    # cell values + styles
    for row in src.iter_rows():
        for cell in row:
            new_cell = dst.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell._style = copy(cell._style)
            new_cell.number_format = cell.number_format

    # column widths
    for key, dim in src.column_dimensions.items():
        d = dst.column_dimensions[key]
        if dim.width is not None:
            d.width = dim.width
        d.hidden = dim.hidden
    # row heights
    for key, dim in src.row_dimensions.items():
        d = dst.row_dimensions[key]
        if dim.height is not None:
            d.height = dim.height
        d.hidden = dim.hidden

    # merged cells
    for rng in list(src.merged_cells.ranges):
        dst.merge_cells(str(rng))

    # freeze panes
    dst.freeze_panes = src.freeze_panes

    _apply_print_setup(dst)
    out.save(dest_xlsx)
    return dest_xlsx


# ── PDF rendering ─────────────────────────────────────────────────────────────
# geometry constants (points; QPdfWriter paints in device pixels at `dpi`)
_EXCEL_WIDTH_TO_PT = 5.25      # 1 Excel char-width -> points (approx)
_DEFAULT_COL_PT = 60.0
_DEFAULT_ROW_PT = 16.0
_HEADER_BAND_PT = 40.0
_CELL_PAD_PT = 3.0
_LOGO_H_PT = 26.0
_MIN_COL_PT = 22.0
_MAX_COL_PT = 240.0
_FROZEN_HEADER_ROWS = 2        # rows repeated at the top of every tile
_MIN_FIT_SCALE = 0.55          # floor for "fit to page width" shrinking
_MIN_FONT_PT = 6.0             # never shrink text below this, even scaled


class _TablePainter:
    """Renders openpyxl worksheets onto an open QPainter/QPdfWriter, with
    branded page headers and 2-D (row x column) pagination for wide tables."""

    def __init__(self, painter: QPainter, writer: QPdfWriter, logo_path=None):
        self.p = painter
        self.writer = writer
        self.dpi = writer.resolution()
        self.scale = self.dpi / 72.0            # points -> device px
        self._logo = None
        if logo_path and os.path.exists(logo_path):
            pm = QPixmap(logo_path)
            if not pm.isNull():
                h = int(round(_LOGO_H_PT * self.scale))
                self._logo = pm.scaledToHeight(h, Qt.SmoothTransformation)
        # printable region in device px
        pr = writer.pageLayout().paintRectPixels(self.dpi)
        self.page_x = float(pr.x())
        self.page_y = float(pr.y())
        self.page_w = float(pr.width())
        self.page_h = float(pr.height())
        self._first_page = True
        self._font_scale = 1.0

    # -- geometry -------------------------------------------------------------
    def _col_widths(self, ws):
        widths = []
        for c in range(1, ws.max_column + 1):
            letter = get_column_letter(c)
            w = None
            dim = ws.column_dimensions.get(letter)
            if dim is not None and dim.width:
                w = dim.width * _EXCEL_WIDTH_TO_PT
            if not w:
                w = _DEFAULT_COL_PT
            w = min(max(w, _MIN_COL_PT), _MAX_COL_PT)
            widths.append(w * self.scale)
        return widths

    def _row_heights(self, ws):
        heights = []
        for r in range(1, ws.max_row + 1):
            h = None
            dim = ws.row_dimensions.get(r)
            if dim is not None and dim.height:
                h = dim.height
            if not h:
                h = _DEFAULT_ROW_PT
            heights.append(h * self.scale)
        return heights

    def _spans(self, ws):
        """merged ranges -> (anchor->(rowspan,colspan)), covered-cell set."""
        spans, covered = {}, set()
        for rng in ws.merged_cells.ranges:
            spans[(rng.min_row, rng.min_col)] = (rng, )
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    if (r, c) != (rng.min_row, rng.min_col):
                        covered.add((r, c))
        return spans, covered

    # -- header band ----------------------------------------------------------
    def _draw_header(self, title):
        p = self.p
        p.save()
        band_h = _HEADER_BAND_PT * self.scale
        x = self.page_x
        y = self.page_y
        text_x = x
        if self._logo is not None:
            lg = self._logo
            ly = y + (band_h - lg.height()) / 2.0
            p.drawPixmap(int(round(x)), int(round(ly)), lg)
            text_x = x + lg.width() + 12 * self.scale
        f = QFont("Segoe UI, Arial")
        f.setBold(True)
        f.setPointSizeF(13.0)
        p.setFont(f)
        p.setPen(QColor(HEADER_TEXT))
        rect = QRectF(text_x, y, (self.page_w - (text_x - x)), band_h)
        p.drawText(rect, int(Qt.AlignLeft | Qt.AlignVCenter), title or "")
        # underline rule
        p.setPen(QColor(NAVY))
        ry = y + band_h
        p.drawLine(int(x), int(round(ry)), int(round(x + self.page_w)), int(round(ry)))
        p.restore()

    def _new_page(self):
        if not self._first_page:
            self.writer.newPage()
        self._first_page = False

    # -- pagination -----------------------------------------------------------
    def _bands(self, sizes, avail):
        """Greedily group consecutive sizes into bands each <= avail (points).
        A single oversize element still gets its own band. A small tolerance
        absorbs float rounding from the fit-to-page scale factor (avail/total
        can land a hair over `avail` after summation) so a page-width scale
        doesn't spuriously spill the last column into its own extra band."""
        bands, start, acc = [], 0, 0.0
        eps = 0.5
        for i, s in enumerate(sizes):
            if acc > 0 and acc + s > avail + eps:
                bands.append((start, i))       # [start, i)
                start, acc = i, 0.0
            acc += s
        bands.append((start, len(sizes)))
        return bands

    def _col_bands(self, widths, avail):
        return self._bands(widths, avail)

    def _row_bands(self, heights, avail, header_rows):
        """Row bands over the *data* rows (after the repeated header rows).
        header_rows is the count of leading rows repeated on every tile."""
        data = heights[header_rows:]
        raw = self._bands(data, avail)
        return [(a + header_rows, b + header_rows) for (a, b) in raw]

    # -- main entry -----------------------------------------------------------
    def render_sheet(self, ws):
        title = None
        try:
            a1 = ws["A1"].value
            if a1 not in (None, ""):
                title = str(a1)
        except Exception:
            title = None
        if not title:
            title = ws.title

        max_row = ws.max_row or 1
        max_col = ws.max_column or 1
        widths = self._col_widths(ws)
        heights = self._row_heights(ws)
        spans, covered = self._spans(ws)

        band_h = _HEADER_BAND_PT * self.scale
        content_top = self.page_y + band_h + 4 * self.scale
        avail_w = self.page_w
        avail_h = (self.page_y + self.page_h) - content_top

        # "fit columns to page": shrink column widths (and text) so the whole
        # row fits on one page width, down to a floor scale; only past that
        # floor do we still fall back to column-band tiling across pages.
        total_w = sum(widths)
        self._font_scale = 1.0
        if total_w > avail_w > 0:
            self._font_scale = max(avail_w / total_w, _MIN_FIT_SCALE)
            widths = [w * self._font_scale for w in widths]

        header_rows = min(_FROZEN_HEADER_ROWS, max_row)
        header_h = sum(heights[:header_rows])

        col_bands = self._col_bands(widths, avail_w)
        row_bands = self._row_bands(heights, max(avail_h - header_h, 1.0), header_rows)

        for (c0, c1) in col_bands:
            for (r0, r1) in row_bands:
                self._new_page()
                self._draw_header(title)
                # y cursor
                x_left = self.page_x
                y = content_top
                # repeated header rows
                if header_rows:
                    self._draw_rows(ws, 1, header_rows + 1, c0 + 1, c1 + 1,
                                    x_left, y, widths, heights, spans, covered)
                    y += header_h
                # data rows for this tile
                self._draw_rows(ws, r0 + 1, r1 + 1, c0 + 1, c1 + 1,
                                x_left, y, widths, heights, spans, covered)

    # -- draw a contiguous block of rows/cols starting at (x_left, y_top) ------
    def _draw_rows(self, ws, r_start, r_stop, c_start, c_stop,
                   x_left, y_top, widths, heights, spans, covered):
        """r_start..r_stop, c_start..c_stop are 1-based, stop-exclusive."""
        # precompute x positions for the visible column band
        x_of = {}
        x = x_left
        for c in range(c_start, c_stop):
            x_of[c] = x
            x += widths[c - 1]
        band_right = x

        y = y_top
        for r in range(r_start, r_stop):
            y_of_r = y
            row_h = heights[r - 1]
            for c in range(c_start, c_stop):
                if (r, c) in covered:
                    continue
                try:
                    self._draw_cell(ws, r, c, x_of, y_of_r, widths, heights,
                                    spans, covered, c_start, c_stop, r_start, r_stop,
                                    band_right, x_left)
                except Exception:
                    # one bad cell must not abort the page
                    continue
            y += row_h

    def _draw_cell(self, ws, r, c, x_of, y_top, widths, heights, spans, covered,
                   c_start, c_stop, r_start, r_stop, band_right, x_left):
        cell = ws.cell(row=r, column=c)
        x = x_of[c]
        w = widths[c - 1]
        h = heights[r - 1]

        # merged anchor: expand rect across the merged range, clipped to tile
        merge = spans.get((r, c))
        if merge:
            rng = merge[0]
            cc0 = max(rng.min_col, c_start)
            cc1 = min(rng.max_col, c_stop - 1)
            rr0 = max(rng.min_row, r_start)
            rr1 = min(rng.max_row, r_stop - 1)
            if cc1 >= cc0 and cc1 >= c and cc0 <= c:
                w = sum(widths[cc - 1] for cc in range(c, cc1 + 1))
            if rr1 >= rr0:
                h = sum(heights[rr - 1] for rr in range(r, rr1 + 1))

        rect = QRectF(x, y_top, w, h)
        p = self.p

        # background fill
        fill = cell.fill
        if fill is not None and getattr(fill, "patternType", None):
            col = _argb_to_qcolor(getattr(getattr(fill, "fgColor", None), "rgb", None))
            if col is not None:
                p.fillRect(rect, col)

        # thin border
        p.setPen(QColor(GRID))
        p.drawRect(rect)

        # text
        text = _fmt(cell.value, cell.number_format)
        if text:
            f = QFont("Segoe UI, Arial")
            font = cell.font
            base_pt = 9.0
            if font is not None:
                if font.bold:
                    f.setBold(True)
                if font.italic:
                    f.setItalic(True)
                if font.size:
                    base_pt = float(font.size)
            f.setPointSizeF(max(_MIN_FONT_PT, base_pt * self._font_scale))
            p.setFont(f)
            fg = None
            if font is not None and font.color is not None:
                fg = _argb_to_qcolor(getattr(font.color, "rgb", None))
            p.setPen(fg if fg is not None else QColor("#20303A"))

            align = cell.alignment
            hz = getattr(align, "horizontal", None) or "left"
            flag = {"center": Qt.AlignHCenter, "right": Qt.AlignRight,
                    "left": Qt.AlignLeft}.get(hz, Qt.AlignLeft)
            pad = _CELL_PAD_PT * self.scale
            trect = QRectF(rect.x() + pad, rect.y(), rect.width() - 2 * pad, rect.height())
            p.save()
            p.setClipRect(rect)
            p.drawText(trect, int(flag | Qt.AlignVCenter), text)
            p.restore()


def _make_writer(dest_pdf: str) -> QPdfWriter:
    writer = QPdfWriter(dest_pdf)
    writer.setResolution(220)          # dpi in the 200-300 band
    layout = QPageLayout(
        QPageSize(QPageSize.PageSizeId.A3),
        QPageLayout.Orientation.Landscape,
        QMarginsF(24, 20, 24, 20),     # points
        QPageLayout.Unit.Point,
    )
    writer.setPageLayout(layout)
    return writer


# ── public: PDF exports ───────────────────────────────────────────────────────
def export_pdf_workbook(src_xlsx: str, dest_pdf: str, logo_path=None) -> str:
    """Render every worksheet to a multi-page A3-landscape PDF."""
    if not os.path.exists(src_xlsx):
        raise Exception(f"Source workbook not found: {src_xlsx}")
    _ensure_app()
    try:
        wb = load_workbook(src_xlsx, data_only=False)
        writer = _make_writer(dest_pdf)
        painter = QPainter()
        if not painter.begin(writer):
            raise Exception("Could not open QPdfWriter for painting.")
        try:
            tp = _TablePainter(painter, writer, logo_path=logo_path)
            for ws in wb.worksheets:
                tp.render_sheet(ws)
        finally:
            painter.end()
    except Exception as e:
        raise Exception(f"PDF export failed: {e}") from e
    return dest_pdf


def export_pdf_sheet(src_xlsx: str, sheet_title: str, dest_pdf: str, logo_path=None) -> str:
    """Render a single worksheet to an A3-landscape PDF."""
    if not os.path.exists(src_xlsx):
        raise Exception(f"Source workbook not found: {src_xlsx}")
    _ensure_app()
    try:
        wb = load_workbook(src_xlsx, data_only=False)
        if sheet_title not in wb.sheetnames:
            raise Exception(f"Sheet '{sheet_title}' not found in {src_xlsx}")
        ws = wb[sheet_title]
        writer = _make_writer(dest_pdf)
        painter = QPainter()
        if not painter.begin(writer):
            raise Exception("Could not open QPdfWriter for painting.")
        try:
            tp = _TablePainter(painter, writer, logo_path=logo_path)
            tp.render_sheet(ws)
        finally:
            painter.end()
    except Exception as e:
        raise Exception(f"PDF export failed: {e}") from e
    return dest_pdf
