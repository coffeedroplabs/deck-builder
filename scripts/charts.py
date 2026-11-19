"""Native, editable PowerPoint charts from a spec's chart block.

Every colour is a theme slot, so a recoloured theme recolours the charts.
One accent marks what the title talks about; everything else is the quiet
series colour. Labels go on the data where they fit, and the value axis is
dropped when they do, because a labelled bar does not need a scale too.
"""

from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import (XL_AXIS_CROSSES, XL_CHART_TYPE, XL_LABEL_POSITION,
                             XL_LEGEND_POSITION, XL_MARKER_STYLE, XL_TICK_MARK)
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Pt

import ooxml
import themes
from numfmt import first_section, format_number

SCHEME = {m.xml_value: m for m in MSO_THEME_COLOR if m.xml_value}
ROLE = {role: SCHEME[val] for role, val in themes.ROLE_SCHEME.items()}
# Order of data colours when nothing is highlighted. accent4 holds the theme's
# muted colour (see ooxml.apply_theme), so a fourth series stays on-palette.
SERIES_ORDER = ("accent1", "accent2", "accent3", "accent4")
PIE_ORDER = ("accent1", "accent2", "accent3", "accent4", "tx1")
XL_TYPE = {
    "bar": XL_CHART_TYPE.BAR_CLUSTERED, "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "stacked-bar": XL_CHART_TYPE.BAR_STACKED, "stacked-column": XL_CHART_TYPE.COLUMN_STACKED,
    "line": XL_CHART_TYPE.LINE, "pie": XL_CHART_TYPE.PIE,
    "waterfall": XL_CHART_TYPE.COLUMN_STACKED, "scatter": XL_CHART_TYPE.XY_SCATTER,
}
TYPE_NAME = {
    "bar": "Bar chart", "column": "Column chart", "stacked-bar": "Stacked bar chart",
    "stacked-column": "Stacked column chart", "line": "Line chart", "pie": "Pie chart",
    "waterfall": "Waterfall chart", "scatter": "Scatter chart",
}
LINE_WIDTH_PT = 2.5
MARKER_SIZE = 7
ALT_MAX_POINTS = 16
# Share of the tallest bar a Base point needs before a step label fits under the step.
WATERFALL_LABEL_ROOM = 0.12
PLOT_TOP = 0.03  # a manual plot area's top and bottom inset, as PowerPoint's own layout leaves


def _signed(fmt, sign):
    """A one-section format that always shows `sign`, for waterfall rises and falls."""
    return sign + first_section(fmt or "General").replace("General", "0")


# -- styling helpers ----------------------------------------------------------

def _fill(fmt, scheme_val):
    fmt.fill.solid()
    fmt.fill.fore_color.theme_color = SCHEME[scheme_val]


def _font(font, size, scheme_val="tx1", bold=False):
    font.size = Pt(size)
    font.name = "+mn-lt"
    font.bold = bold
    font.color.theme_color = SCHEME[scheme_val]


def _readable_on(fill_hex, colors):
    """tx1 or bg1, whichever reads better on a filled bar."""
    on_text = themes.contrast_ratio(colors["text"], fill_hex)
    on_bg = themes.contrast_ratio(colors["bg"], fill_hex)
    return "tx1" if on_text >= on_bg else "bg1"


def _style_category_axis(axis, size):
    axis.major_tick_mark = XL_TICK_MARK.NONE
    axis.minor_tick_mark = XL_TICK_MARK.NONE
    axis.format.line.color.theme_color = SCHEME["bg2"]
    axis.format.line.width = Pt(1)
    axis.has_major_gridlines = False
    _font(axis.tick_labels.font, size, "tx2")


def _style_value_axis(axis, size, fmt, visible):
    axis.has_minor_gridlines = False
    if not visible:
        axis.visible = False
        axis.has_major_gridlines = False
        return
    axis.major_tick_mark = XL_TICK_MARK.NONE
    axis.minor_tick_mark = XL_TICK_MARK.NONE
    axis.format.line.fill.background()
    axis.has_major_gridlines = True
    axis.major_gridlines.format.line.color.theme_color = SCHEME["bg2"]
    axis.major_gridlines.format.line.width = Pt(0.75)
    if fmt:
        axis.tick_labels.number_format = fmt
        axis.tick_labels.number_format_is_linked = False
    _font(axis.tick_labels.font, size, "tx2")


def _axis_title(axis, text, size):
    axis.has_title = True
    tf = axis.axis_title.text_frame
    tf.text = text
    _font(tf.paragraphs[0].runs[0].font, size, "tx2")


def _data_labels(plot, size, fmt, position, scheme_val="tx1"):
    plot.has_data_labels = True
    dl = plot.data_labels
    _font(dl.font, size, scheme_val)
    if fmt:
        dl.number_format = fmt
        dl.number_format_is_linked = False
    dl.show_value = True
    if position is not None:
        dl.position = position
    return dl


def _point_label(point, text, size, position, scheme_val="tx1", bold=False):
    dl = point.data_label
    tf = dl.text_frame
    tf.text = text
    _font(tf.paragraphs[0].runs[0].font, size, scheme_val, bold)
    dl.position = position


def _legend(chart, size):
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.TOP
    chart.legend.include_in_layout = False
    _font(chart.legend.font, size, "tx2")


def _highlights(spec):
    hl = spec.get("highlight")
    if hl is None:
        return set()
    return {str(h) for h in (hl if isinstance(hl, list) else [hl])}


# -- chart data ---------------------------------------------------------------

def _category_data(spec, fmt):
    cd = CategoryChartData(number_format=fmt or "General")
    cd.categories = [str(c) for c in spec["categories"]]
    for s in spec["series"]:
        cd.add_series(s["name"], s["values"])
    return cd


def waterfall_series(categories, values, totals):
    """Split a waterfall into stacked Base/Up/Down/Total series.

    Base is an invisible bar that lifts each change to the running level, so
    a rise or fall floats where the previous bar ended.
    """
    totals = {str(t) for t in totals or []}
    base, up, down, total = [], [], [], []
    level = 0.0
    for cat, v in zip(categories, values):
        if str(cat) in totals:
            base.append(None), up.append(None), down.append(None), total.append(v)
            level = v
        elif v >= 0:
            base.append(level), up.append(v), down.append(None), total.append(None)
            level += v
        else:
            level += v
            base.append(level), up.append(None), down.append(-v), total.append(None)
    return {"Base": base, "Up": up, "Down": down, "Total": total}


def _zero_baseline(axis, values):
    # A bar's length is its value, so its axis starts at zero. Left alone,
    # PowerPoint starts it near the smallest value when the values sit close
    # together, and a rise from 38% to 42% draws as a tripling.
    values = [v for v in values if v is not None]
    if all(v >= 0 for v in values):
        axis.minimum_scale = 0
    elif all(v <= 0 for v in values):
        axis.maximum_scale = 0


# -- builders per type --------------------------------------------------------

def _build_bar_like(chart, spec, size, colors):
    ctype = spec["type"]
    plot = chart.plots[0]
    plot.gap_width = themes.GAP_WIDTH
    plot.vary_by_categories = False
    series = spec["series"]
    single = len(series) == 1
    stacked = ctype.startswith("stacked")
    labels = spec.get("labels", single and not stacked)
    hl = _highlights(spec)
    fmt = spec.get("number_format")

    if single:
        s = plot.series[0]
        _fill(s.format, "accent2" if hl else "accent1")
        for i, cat in enumerate(spec["categories"]):
            if str(cat) in hl:
                _fill(s.points[i].format, "accent1")
    else:
        names = [x["name"] for x in series]
        for i, s in enumerate(plot.series):
            if hl:
                _fill(s.format, "accent1" if names[i] in hl else "accent2")
            else:
                _fill(s.format, SERIES_ORDER[i % len(SERIES_ORDER)])
        _legend(chart, size)

    if labels:
        _data_labels(plot, size, fmt, XL_LABEL_POSITION.OUTSIDE_END)
    _style_category_axis(chart.category_axis, size)
    _style_value_axis(chart.value_axis, size, fmt, visible=not labels)
    _zero_baseline(chart.value_axis, [v for x in series for v in x["values"]])
    if ctype in ("bar", "stacked-bar"):
        # Bars read top-down in the order given, as a ranked list does.
        chart.category_axis.reverse_order = True
        if not labels:
            chart.category_axis.crosses = XL_AXIS_CROSSES.MAXIMUM


def _build_line(chart, spec, size, colors):
    plot = chart.plots[0]
    plot.vary_by_categories = False
    series = spec["series"]
    names = [x["name"] for x in series]
    hl = _highlights(spec)
    fmt = spec.get("number_format")
    labels = spec.get("labels", False)
    single = len(series) == 1

    for i, s in enumerate(plot.series):
        s.smooth = False
        s.marker.style = XL_MARKER_STYLE.NONE
        if single:
            role = "accent1"
        elif hl:
            role = "accent1" if names[i] in hl else "accent2"
        else:
            role = SERIES_ORDER[i % len(SERIES_ORDER)]
        s.format.line.color.theme_color = SCHEME[role]
        s.format.line.width = Pt(LINE_WIDTH_PT)

    if labels:
        _data_labels(plot, size, fmt, XL_LABEL_POSITION.ABOVE)
    if single:
        s = plot.series[0]
        values = series[0]["values"]
        for i, cat in enumerate(spec["categories"]):
            if str(cat) in hl:
                p = s.points[i]
                p.marker.style = XL_MARKER_STYLE.CIRCLE
                p.marker.size = MARKER_SIZE
                _fill(p.marker.format, "accent1")
                p.marker.format.line.color.theme_color = SCHEME["accent1"]
                if not labels:
                    _point_label(p, format_number(values[i], fmt), size, XL_LABEL_POSITION.ABOVE, bold=True)
    else:
        # Name each line at its end instead of a legend, so the eye never has
        # to travel to a key and back.
        last = len(spec["categories"]) - 1
        for i, s in enumerate(plot.series):
            _point_label(s.points[last], names[i], size, XL_LABEL_POSITION.RIGHT)
    _style_category_axis(chart.category_axis, size)
    _style_value_axis(chart.value_axis, size, fmt, visible=not labels)


def _build_pie(chart, spec, size, colors):
    plot = chart.plots[0]
    plot.vary_by_categories = True
    s = plot.series[0]
    hl = _highlights(spec)
    for i, cat in enumerate(spec["categories"]):
        point = s.points[i]
        if hl:
            role = "accent1" if str(cat) in hl else "accent2"
        else:
            role = PIE_ORDER[i % len(PIE_ORDER)]
        _fill(point.format, role)
        point.format.line.color.theme_color = SCHEME["bg1"]
        point.format.line.width = Pt(1.5)
    dl = _data_labels(plot, size, spec.get("number_format"), XL_LABEL_POSITION.OUTSIDE_END)
    dl.show_category_name = True
    chart.has_legend = False


def _build_waterfall(chart, spec, size, colors):
    plot = chart.plots[0]
    plot.gap_width = themes.GAP_WIDTH
    plot.overlap = 100
    plot.vary_by_categories = False
    fmt = spec.get("number_format") or "General"
    fills = {"Up": "accent1", "Down": "accent3", "Total": "accent2"}
    by_name = {s.name: s for s in plot.series}
    for name, s in by_name.items():
        if name == "Base":
            s.format.fill.background()
            s.format.line.fill.background()
        else:
            _fill(s.format, fills[name])
    total = by_name["Total"]
    total.data_labels.show_value = True
    total.data_labels.number_format = first_section(fmt)
    total.data_labels.number_format_is_linked = False
    total.data_labels.position = XL_LABEL_POSITION.INSIDE_END
    _font(total.data_labels.font, size, _readable_on(colors["series"], colors), bold=True)
    # A rise or fall is often too thin to hold its label, so the label goes
    # on the invisible Base point just under it, where the bar ends.
    data = waterfall_series(spec["categories"], spec["series"][0]["values"], spec.get("totals"))
    peak = max(sum(data[k][i] or 0 for k in data) for i in range(len(data["Base"]))) or 1
    for i, base in enumerate(data["Base"]):
        step = data["Up"][i] if data["Up"][i] is not None else data["Down"][i]
        if step is None:
            continue
        sign, name = ("+", "Up") if data["Up"][i] is not None else ("-", "Down")
        text = sign + format_number(step, fmt)
        if base >= WATERFALL_LABEL_ROOM * peak:
            _point_label(by_name["Base"].points[i], text, size, XL_LABEL_POSITION.INSIDE_END, "tx1", bold=True)
        else:
            role_hex = colors["accent"] if name == "Up" else colors["third"]
            _point_label(by_name[name].points[i], text, size, XL_LABEL_POSITION.INSIDE_END,
                         _readable_on(role_hex, colors), bold=True)
    chart.has_legend = False
    _style_category_axis(chart.category_axis, size)
    _style_value_axis(chart.value_axis, size, fmt, visible=False)
    _zero_baseline(chart.value_axis, [v for k in data for v in data[k]])


def _build_scatter(chart, spec, size, colors):
    plot = chart.plots[0]
    names = [x["name"] for x in spec["series"]]
    hl = _highlights(spec)
    for i, s in enumerate(plot.series):
        if hl:
            role = "accent1" if names[i] in hl else "accent2"
        else:
            role = SERIES_ORDER[i % len(SERIES_ORDER)]
        s.format.line.fill.background()
        s.marker.style = XL_MARKER_STYLE.CIRCLE
        s.marker.size = MARKER_SIZE
        _fill(s.marker.format, role)
        s.marker.format.line.color.theme_color = SCHEME[role]
    if len(names) > 1:
        _legend(chart, size)
    fmt = spec.get("number_format")
    x_axis, y_axis = chart.category_axis, chart.value_axis
    for axis in (x_axis, y_axis):
        _style_value_axis(axis, size, fmt, visible=True)
    x_axis.has_major_gridlines = False
    x_axis.format.line.color.theme_color = SCHEME["bg2"]
    _axis_title(x_axis, spec["x_axis_title"], size)
    _axis_title(y_axis, spec["value_axis_title"], size)


BUILDERS = {
    "bar": _build_bar_like, "column": _build_bar_like, "stacked-bar": _build_bar_like,
    "stacked-column": _build_bar_like, "line": _build_line, "pie": _build_pie,
    "waterfall": _build_waterfall, "scatter": _build_scatter,
}


def chart_data(spec):
    ctype = spec["type"]
    fmt = spec.get("number_format")
    if ctype == "scatter":
        cd = XyChartData(number_format=fmt or "General")
        for s in spec["series"]:
            xy = cd.add_series(s["name"])
            for x, y in s["values"]:
                xy.add_data_point(x, y)
        return cd
    if ctype == "waterfall":
        cd = CategoryChartData(number_format=fmt or "General")
        cd.categories = [str(c) for c in spec["categories"]]
        for name, values in waterfall_series(spec["categories"], spec["series"][0]["values"],
                                             spec.get("totals")).items():
            cd.add_series(name, values)
        return cd
    return _category_data(spec, fmt)


def alt_text(spec):
    """A sentence a screen reader can speak: what kind of chart, then the data."""
    if spec.get("alt"):
        return spec["alt"]
    ctype = spec["type"]
    fmt = spec.get("number_format")
    parts = []
    for s in spec["series"]:
        if ctype == "scatter":
            pts = [f"({format_number(x, fmt)}, {format_number(y, fmt)})" for x, y in s["values"]]
        else:
            pts = [f"{c} {format_number(v, fmt)}" for c, v in zip(spec["categories"], s["values"])]
        more = len(pts) - ALT_MAX_POINTS
        pts = pts[:ALT_MAX_POINTS] + ([f"and {more} more"] if more > 0 else [])
        parts.append(f"{s['name']}: " + ", ".join(pts))
    if ctype == "scatter":
        lead = f"{TYPE_NAME[ctype]} of {spec['value_axis_title']} against {spec['x_axis_title']}"
    else:
        lead = TYPE_NAME[ctype]
    return f"{lead}. " + "; ".join(parts) + "."


def add_chart(slide, element, colors):
    """Draw one planned chart element; returns the graphic frame."""
    spec = element["chart"]
    x, y, w, h = element["box"]
    size = element["size"]
    frame = slide.shapes.add_chart(XL_TYPE[spec["type"]], x, y, w, h, chart_data(spec))
    frame.name = element["name"]
    chart = frame.chart
    chart.has_title = False
    chart.has_legend = False
    _font(chart.font, size, "tx2")
    BUILDERS[spec["type"]](chart, spec, size, colors)
    if element.get("label_room"):
        _reserve_right(chart, element["label_room"])
    ooxml.set_alt_text(frame, alt_text(spec))
    return frame


def _reserve_right(chart, share):
    """Pin the plot area, axis labels included, to the left of the chart and
    leave `share` of the width free on the right for the line-end names."""
    plot_area = chart._chartSpace.chart.plotArea
    old = plot_area.find(qn("c:layout"))
    if old is not None:
        plot_area.remove(old)
    plot_area.insert(0, parse_xml(
        f'<c:layout {nsdecls("c")}><c:manualLayout>'
        '<c:layoutTarget val="outer"/><c:xMode val="edge"/><c:yMode val="edge"/>'
        f'<c:x val="0"/><c:y val="{PLOT_TOP}"/><c:w val="{1 - share:.4f}"/><c:h val="{1 - 2 * PLOT_TOP}"/>'
        '</c:manualLayout></c:layout>'))
