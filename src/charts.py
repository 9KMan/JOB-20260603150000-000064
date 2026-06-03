"""
Chart generation helpers built on top of python-pptx's native chart API.

These functions wrap python-pptx's ChartData + BarChart/ColumnChart so the
rest of the codebase can drop in categorised financial data without having
to wrestle with the low-level chart XML.
"""

from __future__ import annotations

from typing import Iterable, Sequence

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.slide import Slide
from pptx.util import Emu, Inches, Pt
from lxml import etree


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def create_bar_chart(
    slide: Slide,
    left: Emu,
    top: Emu,
    width: Emu,
    height: Emu,
    categories: Sequence[str],
    values: Sequence[float],
    title: str,
    color: str = "4472C4",
) -> "Chart":
    """Add a horizontal bar chart to ``slide`` and return the chart object."""
    chart_data = CategoryChartData()
    chart_data.categories = list(categories)
    chart_data.add_series("Amount", [float(v) for v in values])

    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        left,
        top,
        width,
        height,
        chart_data,
    )
    chart = graphic_frame.chart

    _apply_default_styling(chart, title, has_legend=False)
    set_bar_color(chart, color)
    format_axis_with_currency(chart.value_axis)
    _style_category_axis(chart.category_axis)
    return chart


def create_column_chart(
    slide: Slide,
    left: Emu,
    top: Emu,
    width: Emu,
    height: Emu,
    categories: Sequence[str],
    values: Sequence[float],
    title: str,
    color: str = "4472C4",
) -> "Chart":
    """Add a vertical column chart to ``slide`` and return the chart object."""
    chart_data = CategoryChartData()
    chart_data.categories = list(categories)
    chart_data.add_series("Amount", [float(v) for v in values])

    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        left,
        top,
        width,
        height,
        chart_data,
    )
    chart = graphic_frame.chart

    _apply_default_styling(chart, title, has_legend=False)
    set_bar_color(chart, color)
    format_axis_with_currency(chart.value_axis)
    _style_category_axis(chart.category_axis)
    return chart


def create_stacked_column_chart(
    slide: Slide,
    left: Emu,
    top: Emu,
    width: Emu,
    height: Emu,
    categories: Sequence[str],
    series: dict[str, Sequence[float]],
    title: str,
) -> "Chart":
    """Add a stacked column chart with multiple named series."""
    chart_data = CategoryChartData()
    chart_data.categories = list(categories)
    for name, values in series.items():
        chart_data.add_series(name, [float(v) for v in values])

    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_STACKED,
        left,
        top,
        width,
        height,
        chart_data,
    )
    chart = graphic_frame.chart
    _apply_default_styling(chart, title, has_legend=True)
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    format_axis_with_currency(chart.value_axis)
    _style_category_axis(chart.category_axis)
    return chart


def create_pie_chart(
    slide: Slide,
    left: Emu,
    top: Emu,
    width: Emu,
    height: Emu,
    categories: Sequence[str],
    values: Sequence[float],
    title: str,
) -> "Chart":
    """Add a pie chart to ``slide`` for proportional breakdowns."""
    chart_data = CategoryChartData()
    chart_data.categories = list(categories)
    chart_data.add_series("Amount", [float(v) for v in values])

    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE,
        left,
        top,
        width,
        height,
        chart_data,
    )
    chart = graphic_frame.chart
    _apply_default_styling(chart, title, has_legend=True)
    chart.legend.position = XL_LEGEND_POSITION.RIGHT
    chart.legend.include_in_layout = False
    return chart


def format_axis_with_currency(ax) -> None:
    """Format a value axis with a leading ``$`` and thousands separators.

    Silently no-ops on category axes or axes that don't expose
    ``tick_labels.number_format`` (some embedded Excel axes don't).
    """
    if ax is None:
        return
    try:
        tick_labels = ax.tick_labels
    except (AttributeError, ValueError):
        return
    try:
        tick_labels.number_format = '"$"#,##0;[Red]-"$"#,##0'
    except (AttributeError, NotImplementedError, TypeError):
        # Category axis or read-only tick labels - nothing we can do.
        return
    try:
        tick_labels.font.size = Pt(9)
    except (AttributeError, ValueError):
        pass


def set_bar_color(chart, hex_color: str) -> None:
    """Set the fill color of the first series on ``chart``."""
    if not getattr(chart, "series", None):
        return
    try:
        rgb = RGBColor.from_string(hex_color)
    except (ValueError, TypeError):
        return
    series = chart.series[0]
    fill = series.format.fill
    fill.solid()
    fill.fore_color.rgb = rgb
    # Also push the color down into the underlying series XML so the
    # resulting .pptx renders correctly in PowerPoint and LibreOffice.
    _set_series_color_xml(series, hex_color)


def set_series_palette(chart, hex_colors: Sequence[str]) -> None:
    """Apply a list of colors across all series in ``chart`` (cycles if short)."""
    if not getattr(chart, "series", None):
        return
    for index, series in enumerate(chart.series):
        color = hex_colors[index % len(hex_colors)]
        set_bar_color(chart, color)
        # Direct XML assignment for the index-th series.
        try:
            _set_series_color_xml_by_index(chart, index, color)
        except Exception:
            pass


def add_data_labels(chart, position: str = "outside_end") -> None:
    """Turn on data labels for every series in ``chart``."""
    if not getattr(chart, "series", None):
        return
    for series in chart.series:
        try:
            series.data_labels.show_value = True
            series.data_labels.number_format = '"$"#,##0'
            series.data_labels.font.size = Pt(8)
        except (AttributeError, NotImplementedError):
            continue
        try:
            series.data_labels.position = getattr(
                __import__("pptx.enum.chart", fromlist=["XL_LABEL_POSITION"]).XL_LABEL_POSITION,
                position.upper(),
            )
        except (AttributeError, KeyError):
            pass


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _apply_default_styling(chart, title: str, has_legend: bool) -> None:
    chart.has_title = True
    title_frame = chart.chart_title.text_frame
    title_frame.text = title
    for paragraph in title_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.LEFT
        for run in paragraph.runs:
            run.font.size = Pt(12)
            run.font.bold = True
    chart.has_legend = has_legend


def _style_category_axis(ax) -> None:
    if ax is None:
        return
    try:
        ax.tick_labels.font.size = Pt(9)
    except (AttributeError, ValueError):
        pass
    # Prevent PowerPoint from auto-rotating long labels into illegibility.
    try:
        ax_xml = ax._element
        txPr = ax_xml.find(qn("c:txPr"))
        if txPr is not None:
            bodyPr = txPr.find(qn("a:bodyPr"))
            if bodyPr is not None:
                bodyPr.set("rot", "0")
    except Exception:
        # Best-effort only; ignore any structural differences.
        pass


def _set_series_color_xml(series, hex_color: str) -> None:
    """Write a solidFill / srgbClr element on the first series spPr."""
    spPr = series.format._get_or_add_ln() if False else None  # placeholder
    try:
        spPr = series.format.fill._xPr  # underlying spPr element
    except (AttributeError, KeyError):
        return
    nsmap = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    # Drop any existing solidFill children, then append a fresh one.
    for existing in spPr.findall(qn("a:solidFill")):
        spPr.remove(existing)
    solid = etree.SubElement(spPr, qn("a:solidFill"))
    srgb = etree.SubElement(solid, qn("a:srgbClr"))
    srgb.set("val", hex_color.upper())


def _set_series_color_xml_by_index(chart, index: int, hex_color: str) -> None:
    ser_elements = chart._chartSpace.findall(".//" + qn("c:ser"))
    if index >= len(ser_elements):
        return
    ser = ser_elements[index]
    spPr = ser.find(qn("c:spPr"))
    if spPr is None:
        return
    for existing in spPr.findall(qn("a:solidFill")):
        spPr.remove(existing)
    solid = etree.SubElement(spPr, qn("a:solidFill"))
    srgb = etree.SubElement(solid, qn("a:srgbClr"))
    srgb.set("val", hex_color.upper())


# ---------------------------------------------------------------------------
# Convenience factory used by ppt_generator.py
# ---------------------------------------------------------------------------


def build_chart(
    slide: Slide,
    chart_type: str,
    *,
    left: Emu,
    top: Emu,
    width: Emu,
    height: Emu,
    categories: Sequence[str],
    values: Sequence[float],
    title: str,
    color: str = "4472C4",
):
    """Dispatch to the right ``create_*`` function based on a string name."""
    chart_type = chart_type.lower()
    if chart_type in {"bar", "bar_chart", "horizontal_bar"}:
        return create_bar_chart(
            slide, left, top, width, height, categories, values, title, color
        )
    if chart_type in {"column", "column_chart", "vertical_bar"}:
        return create_column_chart(
            slide, left, top, width, height, categories, values, title, color
        )
    if chart_type in {"pie", "pie_chart"}:
        return create_pie_chart(
            slide, left, top, width, height, categories, values, title
        )
    raise ValueError(f"Unsupported chart_type: {chart_type!r}")
