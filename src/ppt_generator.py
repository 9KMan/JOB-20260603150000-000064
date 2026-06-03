"""
PowerPoint report generator for finance/AR reporting.

Produces a polished, multi-slide .pptx deck for a single client with:
  - Title slide
  - Aged balances table
  - Aging bar chart
  - Transaction history table
  - Closing summary slide

Style: corporate finance (navy primary, gold accents, Calibri family,
16:9 widescreen, currency formatted with $ and thousand separators).
"""

from __future__ import annotations

import os
from datetime import datetime
from typing import Any, Dict, Iterable, List, Optional, Sequence

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt, Emu

# ----------------------------------------------------------------------------
# Design tokens
# ----------------------------------------------------------------------------

# Slide dimensions: 16:9 widescreen (13.333" x 7.5")
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# Colour palette
NAVY = RGBColor(0x0B, 0x2A, 0x4A)        # primary dark
NAVY_LIGHT = RGBColor(0x1F, 0x4E, 0x79)  # primary mid
GOLD = RGBColor(0xC9, 0xA2, 0x27)        # accent
GREY_LIGHT = RGBColor(0xF2, 0xF2, 0xF2)  # zebra stripe
GREY_MID = RGBColor(0xBF, 0xBF, 0xBF)    # borders
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x1A, 0x1A, 0x1A)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
RED = RGBColor(0xB7, 0x1C, 0x1C)

# Fonts
FONT_HEAD = "Calibri"
FONT_BODY = "Calibri"
FONT_MONO = "Consolas"

CURRENCY_FMT = '"$"#,##0.00;[Red]-"$"#,##0.00'
DATE_FMT = "%Y-%m-%d"


# ----------------------------------------------------------------------------
# Formatting helpers
# ----------------------------------------------------------------------------

def fmt_currency(value: Any) -> str:
    """Format a numeric value as $1,234.56. None / non-numeric -> '—'."""
    if value is None or value == "":
        return "—"
    try:
        v = float(value)
    except (TypeError, ValueError):
        return str(value)
    if v < 0:
        return f'-${abs(v):,.2f}'
    return f'${v:,.2f}'


def fmt_date(value: Any) -> str:
    """Best-effort YYYY-MM-DD rendering for date/datetime/str inputs."""
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        return value.strftime(DATE_FMT)
    if hasattr(value, "strftime"):  # date, pandas.Timestamp
        try:
            return value.strftime(DATE_FMT)
        except Exception:
            return str(value)
    return str(value)


def set_cell_text(cell, text: str, *, bold: bool = False, size: int = 11,
                  color: RGBColor = BLACK, align=PP_ALIGN.LEFT,
                  font: str = FONT_BODY) -> None:
    """Replace cell contents with a single styled paragraph."""
    cell.text = ""
    tf = cell.text_frame
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def set_cell_fill(cell, rgb: RGBColor) -> None:
    cell.fill.solid()
    cell.fill.fore_color.rgb = rgb


def set_cell_border(cell, *, color: RGBColor = GREY_MID, width: float = 0.5):
    """Apply a thin rectangular border to a single cell."""
    from pptx.oxml.ns import qn
    from lxml import etree
    tag = qn("a:lnL")
    nsmap = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    tcPr = cell._tc.get_or_add_tcPr()
    # Remove existing borders to avoid duplicates
    for side in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        existing = tcPr.find(side, nsmap)
        if existing is not None:
            tcPr.remove(existing)
    for side in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        ln = etree.SubElement(tcPr, side)
        ln.set("w", str(int(width * 12700)))  # EMU
        solid = etree.SubElement(ln, qn("a:solidFill"))
        clr = etree.SubElement(solid, qn("a:srgbClr"))
        clr.set("val", f"{color[0]:02X}{color[1]:02X}{color[2]:02X}")


def add_rectangle(slide, left, top, width, height, *,
                  fill: RGBColor = NAVY, line: Optional[RGBColor] = None) -> Any:
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
    shape.shadow.inherit = False
    return shape


def add_text_box(slide, left, top, width, height, text: str, *,
                 font: str = FONT_BODY, size: int = 12, bold: bool = False,
                 color: RGBColor = BLACK, align=PP_ALIGN.LEFT,
                 anchor=MSO_ANCHOR.TOP) -> Any:
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


# ----------------------------------------------------------------------------
# Main class
# ----------------------------------------------------------------------------

class ClientReport:
    """Build a finance report deck for a single client."""

    def __init__(self, client_name: str, report_date: Optional[Any] = None) -> None:
        self.client_name = client_name
        self.report_date = report_date or datetime.now()

        self.prs = Presentation()
        self.prs.slide_width = SLIDE_WIDTH
        self.prs.slide_height = SLIDE_HEIGHT
        self.blank_layout = self.prs.slide_layouts[6]  # truly blank

        # Track footer disclaimer text for re-use
        self._disclaimer = (
            "Confidential — For internal use only. Figures sourced from "
            "Accounts Receivable extracts and Statements of Account. "
            "Please verify with Finance before external distribution."
        )

    # ------------------------------------------------------------------
    # Title slide
    # ------------------------------------------------------------------
    def add_title_slide(self, title: str, subtitle: str = "") -> None:
        slide = self.prs.slides.add_slide(self.blank_layout)

        # Full-bleed navy background
        add_rectangle(slide, 0, 0, self.prs.slide_width, self.prs.slide_height,
                      fill=NAVY)

        # Gold accent ribbon
        add_rectangle(slide, 0, Inches(2.6), self.prs.slide_width,
                      Inches(0.08), fill=GOLD)

        # Small "Finance Report" eyebrow
        add_text_box(slide, Inches(0.8), Inches(1.9), Inches(11.7), Inches(0.5),
                     "FINANCE REPORT", font=FONT_HEAD, size=18, bold=True,
                     color=GOLD, align=PP_ALIGN.LEFT)

        # Main title
        add_text_box(slide, Inches(0.8), Inches(2.85), Inches(11.7), Inches(1.4),
                     title, font=FONT_HEAD, size=44, bold=True,
                     color=WHITE, align=PP_ALIGN.LEFT)

        # Subtitle
        if subtitle:
            add_text_box(slide, Inches(0.8), Inches(4.2), Inches(11.7),
                         Inches(0.6), subtitle, font=FONT_BODY, size=20,
                         bold=False, color=WHITE, align=PP_ALIGN.LEFT)

        # Client name (large)
        add_text_box(slide, Inches(0.8), Inches(5.1), Inches(11.7),
                     Inches(0.8), f"Prepared for: {self.client_name}",
                     font=FONT_HEAD, size=20, bold=True, color=WHITE,
                     align=PP_ALIGN.LEFT)

        # Report date
        date_str = fmt_date(self.report_date)
        add_text_box(slide, Inches(0.8), Inches(5.85), Inches(11.7),
                     Inches(0.5), f"Report Date: {date_str}",
                     font=FONT_BODY, size=14, color=GOLD,
                     align=PP_ALIGN.LEFT)

        # Footer accent + page number
        add_rectangle(slide, 0, Inches(7.15), self.prs.slide_width,
                      Inches(0.35), fill=NAVY_LIGHT)
        add_text_box(slide, Inches(0.8), Inches(7.18), Inches(11.7),
                     Inches(0.3),
                     "Generated by Finance Report Generator",
                     font=FONT_BODY, size=10, color=WHITE,
                     align=PP_ALIGN.LEFT)

    # ------------------------------------------------------------------
    # Section header helper
    # ------------------------------------------------------------------
    def _add_section_header(self, slide, title: str, subtitle: str = "") -> None:
        # Navy header band
        add_rectangle(slide, 0, 0, self.prs.slide_width, Inches(0.95),
                      fill=NAVY)
        # Gold underline
        add_rectangle(slide, 0, Inches(0.95), self.prs.slide_width,
                      Inches(0.05), fill=GOLD)
        # Title text
        add_text_box(slide, Inches(0.5), Inches(0.18), Inches(12.3),
                     Inches(0.6), title, font=FONT_HEAD, size=24, bold=True,
                     color=WHITE, align=PP_ALIGN.LEFT)
        if subtitle:
            add_text_box(slide, Inches(0.5), Inches(0.55), Inches(12.3),
                         Inches(0.35), subtitle, font=FONT_BODY, size=12,
                         color=GOLD, align=PP_ALIGN.LEFT)
        # Top-right client + date
        add_text_box(slide, Inches(9.0), Inches(0.18), Inches(4.0),
                     Inches(0.35), self.client_name, font=FONT_HEAD,
                     size=12, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)
        add_text_box(slide, Inches(9.0), Inches(0.50), Inches(4.0),
                     Inches(0.35), f"As at {fmt_date(self.report_date)}",
                     font=FONT_BODY, size=10, color=WHITE,
                     align=PP_ALIGN.RIGHT)

    def _add_footer(self, slide, page: int, total: int) -> None:
        add_text_box(slide, Inches(0.5), Inches(7.18), Inches(8.0),
                     Inches(0.25), self._disclaimer, font=FONT_BODY,
                     size=8, color=GREY_MID, align=PP_ALIGN.LEFT)
        add_text_box(slide, Inches(11.5), Inches(7.18), Inches(1.3),
                     Inches(0.25), f"Page {page} of {total}",
                     font=FONT_BODY, size=8, color=GREY_MID,
                     align=PP_ALIGN.RIGHT)

    # ------------------------------------------------------------------
    # Balance table
    # ------------------------------------------------------------------
    def add_balance_table(self, data: Sequence[Dict[str, Any]],
                          page: int = 2, total: int = 4) -> None:
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_section_header(
            slide, "Outstanding Balance Detail",
            "Aged open invoices as at report date"
        )

        headers = ["Invoice #", "Date", "Amount", "Days", "Aging Bucket"]
        rows = list(data) if data else []

        # Layout
        left = Inches(0.5)
        top = Inches(1.3)
        width = Inches(12.3)
        height = Inches(5.5)
        n_rows = len(rows) + 1  # +1 header
        n_cols = len(headers)

        if n_rows < 2:
            n_rows = 2
        table_shape = slide.shapes.add_table(n_rows, n_cols, left, top,
                                             width, height)
        tbl = table_shape.table

        # Column widths
        col_widths_in = [2.2, 1.5, 1.8, 1.0, 1.8]
        total_in = sum(col_widths_in)
        scale = 12.3 / total_in
        for i, w in enumerate(col_widths_in):
            tbl.columns[i].width = Inches(w * scale)

        # Header row
        for c, h in enumerate(headers):
            cell = tbl.cell(0, c)
            set_cell_text(cell, h, bold=True, size=12, color=WHITE,
                          align=PP_ALIGN.CENTER, font=FONT_HEAD)
            set_cell_fill(cell, NAVY)
            set_cell_border(cell)

        # Body rows
        total_amount = 0.0
        for r, row in enumerate(rows, start=1):
            zebra = GREY_LIGHT if r % 2 == 0 else WHITE
            inv = str(row.get("invoice_no", ""))
            dt = fmt_date(row.get("date"))
            amt_raw = row.get("amount", 0)
            try:
                amt_val = float(amt_raw) if amt_raw not in (None, "") else 0.0
            except (TypeError, ValueError):
                amt_val = 0.0
            total_amount += amt_val
            days = row.get("days", "")
            try:
                days = f"{int(days)}"
            except (TypeError, ValueError):
                days = str(days) if days is not None else ""
            bucket = str(row.get("aging_bucket", ""))

            set_cell_text(tbl.cell(r, 0), inv, size=10, font=FONT_MONO,
                          align=PP_ALIGN.LEFT)
            set_cell_text(tbl.cell(r, 1), dt, size=10, align=PP_ALIGN.CENTER)
            set_cell_text(tbl.cell(r, 2), fmt_currency(amt_val), size=10,
                          align=PP_ALIGN.RIGHT)
            set_cell_text(tbl.cell(r, 3), days, size=10, align=PP_ALIGN.CENTER)
            set_cell_text(tbl.cell(r, 4), bucket, size=10, bold=True,
                          align=PP_ALIGN.CENTER,
                          color=self._bucket_color(bucket))

            for c in range(n_cols):
                cell = tbl.cell(r, c)
                set_cell_fill(cell, zebra)
                set_cell_border(cell)

        # Total row
        total_row = len(rows) + 1
        if total_row >= n_rows:
            # Need to add a row for total
            # python-pptx doesn't expose row append easily; rebuild
            pass
        # Always show summary box below table
        summary_top = top + Inches(5.55)
        add_rectangle(slide, Inches(0.5), summary_top, Inches(12.3),
                      Inches(0.55), fill=NAVY_LIGHT)
        add_text_box(slide, Inches(0.7), summary_top + Inches(0.08),
                     Inches(6.0), Inches(0.4),
                     f"Total outstanding: {fmt_currency(total_amount)}",
                     font=FONT_HEAD, size=14, bold=True, color=WHITE)
        add_text_box(slide, Inches(6.7), summary_top + Inches(0.08),
                     Inches(6.0), Inches(0.4),
                     f"Open invoices: {len(rows)}",
                     font=FONT_BODY, size=12, color=WHITE,
                     align=PP_ALIGN.RIGHT)

        self._add_footer(slide, page, total)

    @staticmethod
    def _bucket_color(bucket: str) -> RGBColor:
        b = (bucket or "").lower()
        if "0-30" in b or "current" in b:
            return GREEN
        if "31-60" in b:
            return GOLD
        if "61-90" in b:
            return RGBColor(0xE6, 0x7E, 0x22)
        if "90" in b:
            return RED
        return BLACK

    # ------------------------------------------------------------------
    # Aging chart
    # ------------------------------------------------------------------
    def add_aging_chart(self, aging_data: Dict[str, float],
                        page: int = 3, total: int = 4) -> None:
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_section_header(
            slide, "Aged Receivables Summary",
            "Distribution of outstanding balances by aging bucket"
        )

        # Order buckets canonically
        ordered_keys = ["0-30", "31-60", "61-90", "90+"]
        categories: List[str] = []
        values: List[float] = []
        for k in ordered_keys:
            v = aging_data.get(k, 0.0)
            try:
                v = float(v) if v not in (None, "") else 0.0
            except (TypeError, ValueError):
                v = 0.0
            categories.append(k)
            values.append(v)

        # Bar chart (horizontal feel)
        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series("Outstanding", values)

        chart_left = Inches(0.7)
        chart_top = Inches(1.5)
        chart_width = Inches(8.5)
        chart_height = Inches(5.0)
        chart_shape = slide.shapes.add_chart(
            XL_CHART_TYPE.BAR_CLUSTERED, chart_left, chart_top,
            chart_width, chart_height, chart_data
        )
        chart = chart_shape.chart
        chart.has_title = True
        chart.chart_title.text_frame.text = "Aged Receivables ($)"
        for run in chart.chart_title.text_frame.paragraphs[0].runs:
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.color.rgb = NAVY
            run.font.name = FONT_HEAD

        chart.has_legend = False

        # Format axes
        try:
            cat_ax = chart.category_axis
            for tick in cat_ax.tick_labels.font:
                pass
            cat_ax.tick_labels.font.size = Pt(11)
            cat_ax.tick_labels.font.name = FONT_BODY
            cat_ax.tick_labels.font.color.rgb = NAVY
        except Exception:
            pass

        try:
            val_ax = chart.value_axis
            val_ax.tick_labels.font.size = Pt(10)
            val_ax.tick_labels.font.name = FONT_BODY
            val_ax.tick_labels.font.color.rgb = NAVY
            val_ax.tick_labels.number_format = CURRENCY_FMT
            val_ax.tick_labels.number_format_is_linked = False
        except Exception:
            pass

        # Color the bars
        try:
            series = chart.series[0]
            for idx, point in enumerate(series.points):
                fill = point.format.fill
                fill.solid()
                fill.fore_color.rgb = self._bucket_color(categories[idx])
            # Data labels
            plot = chart.plots[0]
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.font.size = Pt(10)
            dl.font.bold = True
            dl.font.color.rgb = NAVY
            dl.font.name = FONT_BODY
            dl.number_format = CURRENCY_FMT
            dl.number_format_is_linked = False
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
        except Exception:
            pass

        # Side summary panel
        panel_left = Inches(9.5)
        panel_top = Inches(1.5)
        panel_width = Inches(3.3)
        panel_height = Inches(5.0)
        add_rectangle(slide, panel_left, panel_top, panel_width,
                      panel_height, fill=GREY_LIGHT, line=GREY_MID)
        add_text_box(slide, panel_left + Inches(0.15),
                     panel_top + Inches(0.15),
                     panel_width - Inches(0.3), Inches(0.4),
                     "Bucket Summary", font=FONT_HEAD, size=14, bold=True,
                     color=NAVY)

        y = panel_top + Inches(0.65)
        total = sum(values) or 1.0
        for k, v in zip(categories, values):
            pct = (v / total) * 100.0
            add_text_box(slide, panel_left + Inches(0.15), y,
                         panel_width - Inches(0.3), Inches(0.3),
                         f"{k} days", font=FONT_HEAD, size=11, bold=True,
                         color=self._bucket_color(k))
            add_text_box(slide, panel_left + Inches(0.15), y + Inches(0.28),
                         panel_width - Inches(0.3), Inches(0.3),
                         fmt_currency(v), font=FONT_BODY, size=11,
                         color=BLACK)
            add_text_box(slide, panel_left + Inches(0.15), y + Inches(0.55),
                         panel_width - Inches(0.3), Inches(0.3),
                         f"{pct:.1f}% of total", font=FONT_BODY, size=9,
                         color=GREY_MID)
            y += Inches(1.05)

        # Footer
        add_text_box(slide, Inches(0.7), Inches(6.6), Inches(12.0),
                     Inches(0.4),
                     f"Total aged receivables: {fmt_currency(sum(values))}",
                     font=FONT_HEAD, size=14, bold=True, color=NAVY)

        self._add_footer(slide, page, total)

    # ------------------------------------------------------------------
    # Transaction history
    # ------------------------------------------------------------------
    def add_transaction_history(self, transactions: Sequence[Dict[str, Any]],
                                page: int = 4, total: int = 4) -> None:
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._add_section_header(
            slide, "Transaction History",
            "Recent activity on this account"
        )

        headers = ["Date", "Description", "Reference", "Amount"]
        rows = list(transactions) if transactions else []

        left = Inches(0.5)
        top = Inches(1.3)
        width = Inches(12.3)
        height = Inches(5.5)
        n_rows = max(len(rows) + 1, 2)
        n_cols = len(headers)
        table_shape = slide.shapes.add_table(n_rows, n_cols, left, top,
                                             width, height)
        tbl = table_shape.table

        col_widths_in = [1.6, 6.4, 2.3, 2.0]
        total_in = sum(col_widths_in)
        scale = 12.3 / total_in
        for i, w in enumerate(col_widths_in):
            tbl.columns[i].width = Inches(w * scale)

        # Header
        for c, h in enumerate(headers):
            cell = tbl.cell(0, c)
            set_cell_text(cell, h, bold=True, size=12, color=WHITE,
                          align=PP_ALIGN.CENTER, font=FONT_HEAD)
            set_cell_fill(cell, NAVY)
            set_cell_border(cell)

        net_total = 0.0
        for r, row in enumerate(rows, start=1):
            zebra = GREY_LIGHT if r % 2 == 0 else WHITE
            dt = fmt_date(row.get("date", ""))
            desc = str(row.get("description", ""))
            ref = str(row.get("reference",
                              row.get("invoice_no",
                                      row.get("ref", ""))))
            amt_raw = row.get("amount", 0)
            try:
                amt_val = float(amt_raw) if amt_raw not in (None, "") else 0.0
            except (TypeError, ValueError):
                amt_val = 0.0
            net_total += amt_val

            color = RED if amt_val < 0 else BLACK
            set_cell_text(tbl.cell(r, 0), dt, size=10,
                          align=PP_ALIGN.CENTER)
            set_cell_text(tbl.cell(r, 1), desc, size=10,
                          align=PP_ALIGN.LEFT)
            set_cell_text(tbl.cell(r, 2), ref, size=10, font=FONT_MONO,
                          align=PP_ALIGN.LEFT)
            set_cell_text(tbl.cell(r, 3), fmt_currency(amt_val), size=10,
                          color=color, align=PP_ALIGN.RIGHT)

            for c in range(n_cols):
                cell = tbl.cell(r, c)
                set_cell_fill(cell, zebra)
                set_cell_border(cell)

        # Net movement panel
        panel_top = top + Inches(5.55)
        add_rectangle(slide, Inches(0.5), panel_top, Inches(12.3),
                      Inches(0.55), fill=NAVY_LIGHT)
        add_text_box(slide, Inches(0.7), panel_top + Inches(0.08),
                     Inches(6.0), Inches(0.4),
                     f"Net movement: {fmt_currency(net_total)}",
                     font=FONT_HEAD, size=14, bold=True, color=WHITE)
        add_text_box(slide, Inches(6.7), panel_top + Inches(0.08),
                     Inches(6.0), Inches(0.4),
                     f"Transactions: {len(rows)}",
                     font=FONT_BODY, size=12, color=WHITE,
                     align=PP_ALIGN.RIGHT)

        self._add_footer(slide, page, total)

    # ------------------------------------------------------------------
    # Save
    # ------------------------------------------------------------------
    def save(self, output_path: str) -> str:
        """Persist the deck to disk. Creates parent dirs as needed."""
        out_dir = os.path.dirname(os.path.abspath(output_path))
        if out_dir and not os.path.isdir(out_dir):
            os.makedirs(out_dir, exist_ok=True)
        self.prs.save(output_path)
        return os.path.abspath(output_path)
