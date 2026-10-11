"""Chart tools with verified ranges and scoped workbook ownership."""

import logging
import re
from collections.abc import Iterator
from contextlib import contextmanager
from typing import Literal

from pydantic import Field, model_validator
from pywintypes import com_error

from src.common.colors import hex_to_int
from src.models import DocumentInput, InputModel, JsonObject, NativeObject
from src.native.runtime import get_session
from src.native.state import temporary_attributes
from src.presentation.navigation import goto_slide
from src.presentation.shape_lookup import resolve_shape
from src.protocol.definitions import tool

logger = logging.getLogger(__name__)
CHART_TYPE_MAP: dict[str, int] = {
    "column": 51,
    "column_stacked": 52,
    "column_stacked_100": 53,
    "bar": 57,
    "bar_stacked": 58,
    "line": 4,
    "line_markers": 65,
    "line_stacked": 63,
    "pie": 5,
    "pie_exploded": 69,
    "doughnut": -4120,
    "area": 1,
    "area_stacked": 76,
    "scatter": -4169,
    "scatter_lines": 74,
    "radar": -4151,
    "bubble": 15,
    "3d_column": 54,
    "3d_pie": -4102,
    "3d_line": -4101,
}
LEGEND_POSITION_MAP: dict[str, int] = {
    "bottom": -4107,
    "left": -4131,
    "right": -4152,
    "top": -4160,
    "corner": 2,
}
AXIS_TYPE_MAP: dict[str, int] = {
    "category": 1,
    "value": 2,
    "secondary_value": 2,
    "series": 3,
}
AXIS_GROUP_PRIMARY = 1
AXIS_GROUP_SECONDARY = 2
TICK_MARK_MAP: dict[str, int] = {
    "none": -4142,
    "inside": 2,
    "outside": 3,
    "cross": 4,
}
SCALE_LINEAR = -4132
SCALE_LOGARITHMIC = -4133
CHART_TYPE_NAMES: dict[int, str] = {v: k for k, v in CHART_TYPE_MAP.items()}


class ChartInput(DocumentInput):
    """Locate a chart by shape name or one-based shape index."""

    slide_index: int = Field(ge=1)
    shape_name_or_index: str | int


class ChartSeriesInput(InputModel):
    """A named series containing finite numeric values."""

    name: str
    values: list[float] = Field(min_length=1)


class AddChartInput(DocumentInput):
    """Create a chart with dimensions in points."""

    slide_index: int = Field(ge=1)
    chart_type: int | str = "column"
    left: float = 50.0
    top: float = 50.0
    width: float = Field(default=500.0, gt=0)
    height: float = Field(default=350.0, gt=0)


class SetChartDataInput(ChartInput):
    """Replace data within the existing, verified column data range."""

    categories: list[str] = Field(min_length=1)
    series: list[ChartSeriesInput] = Field(min_length=1)

    @model_validator(mode="after")
    def check_dimensions(self) -> "SetChartDataInput":
        """Require one numeric value per category in every series."""
        if any(
            len(item.values) != len(self.categories) for item in self.series
        ):
            raise ValueError("Each series must have one value per category")
        return self


class GetChartDataInput(ChartInput):
    """Read actual plotted categories and series."""


class FormatChartInput(ChartInput):
    """Set chart styles; coordinates override position presets."""

    title: str | None = None
    has_legend: bool | None = None
    legend_position: (
        Literal[
            "bottom",
            "left",
            "right",
            "top",
            "corner",
            "top-left",
            "top-center",
            "top-right",
            "middle-left",
            "middle-right",
            "bottom-left",
            "bottom-center",
            "bottom-right",
        ]
        | None
    ) = None
    chart_style: int | None = Field(default=None, ge=1)
    legend_font_size: float | None = Field(default=None, gt=0)
    legend_top: float | None = Field(default=None, ge=0)
    legend_left: float | None = Field(default=None, ge=0)
    title_position: Literal["top", "bottom", "center"] | None = None
    title_top: float | None = Field(default=None, ge=0)
    title_left: float | None = Field(default=None, ge=0)


class FormatChartAxisInput(ChartInput):
    """Format scales on value axes and spacing on category axes."""

    axis: Literal["category", "value", "secondary_value", "series"] = (
        "category"
    )
    title: str | None = None
    min_scale: float | None = None
    max_scale: float | None = None
    major_unit: float | None = Field(default=None, gt=0)
    minor_unit: float | None = Field(default=None, gt=0)
    tick_label_spacing: int | None = Field(default=None, ge=1)
    tick_mark_spacing: int | None = Field(default=None, ge=1)
    major_tick_mark: Literal["none", "inside", "outside", "cross"] | None = (
        None
    )
    minor_tick_mark: Literal["none", "inside", "outside", "cross"] | None = (
        None
    )
    reverse_order: bool | None = None
    log_scale: bool | None = None
    log_base: float | None = Field(default=None, gt=1)
    tick_label_font_size: float | None = Field(default=None, gt=0)
    number_format: str | None = None

    @model_validator(mode="after")
    def check_axis_properties(self) -> "FormatChartAxisInput":
        """Reject settings incompatible with the selected axis."""
        if self.axis != "category" and (
            self.tick_label_spacing is not None
            or self.tick_mark_spacing is not None
        ):
            raise ValueError(
                "Tick spacing is only valid for the category axis"
            )
        if self.axis not in {"value", "secondary_value"} and any(
            value is not None
            for value in (
                self.min_scale,
                self.max_scale,
                self.major_unit,
                self.minor_unit,
                self.log_scale,
                self.log_base,
            )
        ):
            raise ValueError("Scale properties are only valid for value axes")
        if self.log_base is not None and self.log_scale is not True:
            raise ValueError("log_base requires log_scale=true")
        if (
            self.min_scale is not None
            and self.max_scale is not None
            and self.min_scale >= self.max_scale
        ):
            raise ValueError("min_scale must be less than max_scale")
        if self.log_scale and any(
            value is not None and value <= 0
            for value in (self.min_scale, self.max_scale)
        ):
            raise ValueError("Logarithmic scale bounds must be positive")
        return self


class SetChartSeriesInput(ChartInput):
    """Format one plotted series."""

    series_index: int = Field(ge=1)
    color: str | None = None
    show_data_labels: bool | None = None
    line_weight: float | None = Field(default=None, gt=0)


class ChangeChartTypeInput(ChartInput):
    """Change the chart type without changing data."""

    chart_type: int | str


@contextmanager
def chart_workbook(
    chart: NativeObject, *, own_existing: bool = False
) -> Iterator[NativeObject]:
    """Close only workbooks owned by the current operation.

    A Workbook accessible before Activate is treated as borrowed. New
    charts explicitly own any workbook WPS opened automatically during
    their creation.
    """
    data = chart.ChartData
    try:
        workbook = data.Workbook
    except (com_error, AttributeError):
        workbook = None
    owned = own_existing or workbook is None
    if workbook is None:
        data.Activate()
        workbook = data.Workbook
    failed = False
    try:
        yield workbook
    except BaseException:
        failed = True
        raise
    finally:
        if owned:
            try:
                workbook.Close(False)
            except Exception:
                if not failed:
                    raise
                logger.exception(
                    "Cannot close the chart workbook after an error"
                )


def _resolve_chart_type(chart_type: str | int) -> int:
    """Resolve a chart type name to its XlChartType integer.

    If chart_type is already an int, returns it unchanged. If it is a
    string, looks it up in CHART_TYPE_MAP.

    Raises:
        ValueError: If the string name is not recognized.
    """
    if isinstance(chart_type, int):
        return chart_type
    key = chart_type.strip().lower()
    if key not in CHART_TYPE_MAP:
        raise ValueError(
            f"Unknown chart type {chart_type!r}. Valid "
            f"names: {', '.join(sorted(CHART_TYPE_MAP))}"
        )
    return CHART_TYPE_MAP[key]


def _get_chart_shape(
    params: ChartInput, *, navigate: bool = True
) -> NativeObject:
    """Resolve the selected document, slide and required chart shape."""
    session = get_session("wpp")
    document = session.get_document(params.document)
    if navigate:
        goto_slide(session.get_application(), params.slide_index)
    shape = resolve_shape(
        document.Slides(params.slide_index), params.shape_name_or_index
    )
    if not shape.HasChart:
        raise ValueError(f"Shape '{shape.Name}' is not a chart")
    return shape


def _validate_chart_layout(
    chart: NativeObject,
    worksheet: NativeObject,
    category_count: int,
    series_count: int,
) -> None:
    """Validate all chart references before writing cells."""
    actual_count = chart.SeriesCollection().Count
    if actual_count != series_count:
        raise NotImplementedError(
            f"Cannot resize chart series ({actual_count} existing, "
            f"{series_count} requested). Nothing was written."
        )
    sheet_name = worksheet.Name
    quoted_sheet = "'" + sheet_name.replace("'", "''") + "'"
    valid_sheets = {sheet_name.casefold(), quoted_sheet.casefold()}
    sheet_pattern = r"(?:'(?:[^']|'')+'|[^'!,()]+)"
    for index in range(1, actual_count + 1):
        item = chart.SeriesCollection(index)
        if (
            len(tuple(item.XValues)) != category_count
            or len(tuple(item.Values)) != category_count
        ):
            raise NotImplementedError(
                "Cannot resize chart categories. Nothing was written."
            )
        column_index = index + 1
        column = ""
        while column_index:
            column_index, remainder = divmod(column_index - 1, 26)
            column = chr(65 + remainder) + column
        final_row = category_count + 1
        pattern = (
            rf"=SERIES\((?P<header>{sheet_pattern})!\$?{column}\$?1,"
            rf"(?P<categories>{sheet_pattern})!\$?A\$?2:\$?A\$?{final_row},"
            rf"(?P<values>{sheet_pattern})!\$?{column}\$?2:"
            rf"\$?{column}\$?{final_row},{index}\)"
        )
        match = re.fullmatch(
            pattern, item.Formula.strip(), flags=re.IGNORECASE
        )
        if match is None or any(
            value.casefold() not in valid_sheets
            for value in match.groupdict().values()
        ):
            raise NotImplementedError(
                "Unsupported chart source: expected first-sheet column layout "
                "with headers in row 1, categories in column A and values in "
                "columns B onward. Nothing was written."
            )


def _read_chart_series(
    chart: NativeObject,
) -> tuple[list[str], list[JsonObject]]:
    """Read plotted values, rejecting mismatched categories."""
    categories: list[str] = []
    series_list: list[JsonObject] = []
    for index in range(1, chart.SeriesCollection().Count + 1):
        item = chart.SeriesCollection(index)
        item_categories = [str(value) for value in item.XValues]
        if index == 1:
            categories = item_categories
        elif item_categories != categories:
            raise ValueError("Chart series have different category lists")
        series_list.append({"name": item.Name, "values": list(item.Values)})
    return categories, series_list


@tool(
    "wps_ppt_charts",
    component="wpp",
    annotations={"title": "Add Chart"},
    operation="add",
)
def add_chart(params: AddChartInput) -> JsonObject:
    """Add a chart to a slide.

    Creates a chart with the specified type and position. Supported
    types: column, bar, line, pie, scatter, area, doughnut, radar,
    bubble, 3d_column, 3d_pie, 3d_line, and variants (e.g.
    column_stacked, line_markers, pie_exploded). You can also pass an
    XlChartType integer. All positions and sizes are in points (72
    points = 1 inch).
    """
    app = get_session("wpp").get_application()
    goto_slide(app, params.slide_index)
    pres = get_session("wpp").get_document(params.document)
    slide = pres.Slides(params.slide_index)
    type_int = _resolve_chart_type(params.chart_type)
    chart_shape = slide.Shapes.AddChart2(
        -1, type_int, params.left, params.top, params.width, params.height
    )
    chart = chart_shape.Chart
    with chart_workbook(chart, own_existing=True):
        pass
    type_name = CHART_TYPE_NAMES.get(type_int, str(type_int))
    return {
        "success": True,
        "shape_name": chart_shape.Name,
        "shape_index": chart_shape.ZOrderPosition,
        "chart_type": type_name,
        "chart_type_int": type_int,
    }


@tool(
    "wps_ppt_charts",
    component="wpp",
    annotations={"title": "Set Chart Data", "idempotentHint": True},
    operation="set_data",
)
def set_chart_data(params: SetChartDataInput) -> JsonObject:
    """Update the chart data range and verify plotted values.

    This WPS version cannot resize chart series or categories. Counts
    must match the existing chart, whose source must be a first-sheet
    column layout: headers in row 1, categories in column A, values in B
    onward. Unsupported dimensions/layouts are refused before any cells
    are written.

    Provide category labels and one or more series, each with a name and
    numeric values. The number of values in each series should match the
    number of categories.
    Example series:
        [{"name": "Revenue", "values": [100, 200, 150, 300]}]
    """
    shape = _get_chart_shape(params)
    chart = shape.Chart
    with chart_workbook(chart) as workbook:
        worksheet = workbook.Worksheets(1)
        _validate_chart_layout(
            chart, worksheet, len(params.categories), len(params.series)
        )
        for index, category in enumerate(params.categories, start=2):
            worksheet.Cells(index, 1).Value2 = "'" + category
        for series_index, item in enumerate(params.series, start=2):
            worksheet.Cells(1, series_index).Value2 = "'" + item.name
            for row_index, value in enumerate(item.values, start=2):
                worksheet.Cells(row_index, series_index).Value2 = value
        actual_categories, actual_series = _read_chart_series(chart)
        expected_series = [item.model_dump() for item in params.series]
        if (
            actual_categories != list(params.categories)
            or actual_series != expected_series
        ):
            raise RuntimeError(
                "WPS plotted data differs from the request; "
                "embedded worksheet "
                "cells may already have changed."
            )
    return {
        "success": True,
        "shape_name": shape.Name,
        "categories_count": len(params.categories),
        "series_count": len(params.series),
        "verified": True,
    }


@tool(
    "wps_ppt_charts",
    component="wpp",
    annotations={
        "title": "Get Chart Data",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="get_data",
)
def get_chart_data(params: GetChartDataInput) -> JsonObject:
    """Read chart data from the underlying WPS Spreadsheets workbook.

    Returns the category labels and all series (name + values). Identify
    the chart by shape name or 1-based shape index.
    """
    shape = _get_chart_shape(params, navigate=False)
    chart = shape.Chart
    with chart_workbook(chart):
        categories, series_list = _read_chart_series(chart)
    return {
        "success": True,
        "shape_name": shape.Name,
        "categories": categories,
        "series": series_list,
    }


@tool(
    "wps_ppt_charts",
    component="wpp",
    annotations={"title": "Format Chart", "idempotentHint": True},
    operation="format",
)
def format_chart(params: FormatChartInput) -> JsonObject:
    """Format chart properties: title, legend, and chart style.

    Set a chart title (enables HasTitle automatically), show/hide the
    legend, set legend position using WPS Presentation presets
    ('bottom', 'left', 'right', 'top', 'corner') or 8-direction
    coordinate presets ('top-left', 'top-center', 'top-right',
    'middle-left', 'middle-right', 'bottom-left', 'bottom-center',
    'bottom-right'), or use legend_top/legend_left for explicit
    coordinate positioning. Apply a built-in chart style by index
    number, and set the legend font size in points via legend_font_size.
    Control chart title position via title_position preset ('top',
    'bottom', 'center') or explicit title_top/title_left
    coordinates in points.
    """
    shape = _get_chart_shape(params)
    chart = shape.Chart
    has_legend = (
        chart.HasLegend if params.has_legend is None else params.has_legend
    )
    legend_requested = any(
        value is not None
        for value in (
            params.legend_position,
            params.legend_top,
            params.legend_left,
            params.legend_font_size,
        )
    )
    title_requested = any(
        value is not None
        for value in (
            params.title_position,
            params.title_top,
            params.title_left,
        )
    )
    if legend_requested and not has_legend:
        raise ValueError("Legend formatting requires has_legend=true")
    if title_requested and params.title is None and not chart.HasTitle:
        raise ValueError("Title positioning requires a chart title")
    if params.title is not None:
        chart.HasTitle = True
        chart.ChartTitle.Text = params.title
    if params.has_legend is not None:
        chart.HasLegend = params.has_legend
    if params.chart_style is not None:
        chart.ChartStyle = params.chart_style
    if legend_requested:
        legend = chart.Legend
        if params.legend_font_size is not None:
            legend.Font.Size = params.legend_font_size
        if params.legend_position is not None:
            position = params.legend_position
            if position in LEGEND_POSITION_MAP:
                legend.Position = LEGEND_POSITION_MAP[position]
            else:
                row, column = position.split("-")
                legend.Top = {
                    "top": 5,
                    "middle": (chart.ChartArea.Height - legend.Height) / 2,
                    "bottom": chart.ChartArea.Height - legend.Height - 5,
                }[row]
                legend.Left = {
                    "left": 5,
                    "center": (chart.ChartArea.Width - legend.Width) / 2,
                    "right": chart.ChartArea.Width - legend.Width - 5,
                }[column]
        if params.legend_top is not None:
            legend.Top = params.legend_top
        if params.legend_left is not None:
            legend.Left = params.legend_left
    warnings: list[str] = []
    if title_requested:
        title = chart.ChartTitle
        if params.title_position == "top":
            # WPS restores automatic plot layout when the title is recreated.
            title_text = title.Text
            font = title.Font
            saved_font: JsonObject = {}
            for name in ("Name", "Size", "Bold", "Color"):
                try:
                    value = getattr(font, name)
                    saved_font[name] = value.RGB if name == "Color" else value
                except (com_error, AttributeError):
                    warnings.append(f"Cannot preserve chart title font {name}")
            with temporary_attributes(chart, HasTitle=False):
                pass
            title = chart.ChartTitle
            title.Text = title_text
            for name, value in saved_font.items():
                if name == "Color":
                    title.Font.Color.RGB = value
                else:
                    setattr(title.Font, name, value)
        elif params.title_position == "bottom":
            plot = chart.PlotArea
            if chart.HasLegend and (
                chart.Legend.Position == LEGEND_POSITION_MAP["bottom"]
                or (params.legend_position or "").startswith("bottom-")
            ):
                title_top = chart.Legend.Top - title.Height - 5
            else:
                title_top = chart.ChartArea.Height - title.Height - 5
            available_height = title_top - 5 - plot.Top
            if available_height <= 0:
                raise ValueError("Chart has no room for a bottom title")
            if plot.Height > available_height:
                plot.Height = available_height
            title.Top = title_top
            title.Left = (chart.ChartArea.Width - title.Width) / 2
        elif params.title_position == "center":
            title.Top = (chart.ChartArea.Height - title.Height) / 2
            title.Left = (chart.ChartArea.Width - title.Width) / 2
        if params.title_top is not None:
            title.Top = params.title_top
        if params.title_left is not None:
            title.Left = params.title_left
    result: JsonObject = {
        "success": True,
        "shape_name": shape.Name,
        "has_title": bool(chart.HasTitle),
        "has_legend": bool(chart.HasLegend),
    }
    if warnings:
        result["warnings"] = warnings
    return result


@tool(
    "wps_ppt_charts",
    component="wpp",
    annotations={"title": "Format Chart Axis", "idempotentHint": True},
    operation="format_axis",
)
def format_chart_axis(params: FormatChartAxisInput) -> JsonObject:
    """Format axis scales, spacing, tick marks and labels.

    Select the axis via the `axis` parameter:
    - 'category' — horizontal x-axis on column/line/bar charts (default)
    - 'value' — primary y-axis
    - 'secondary_value' — secondary y-axis
    - 'series' — 3D charts only

    `tick_label_spacing` / `tick_mark_spacing` are category-axis only:
    for example, on a 0–60 day axis, `tick_label_spacing=3` renders
    labels at 0, 3, 6, …, 60.

    `min_scale` / `max_scale` / `major_unit` / `minor_unit` /
    `log_scale` are value-axis only.
    `log_base` requires `log_scale=true`.
    """
    shape = _get_chart_shape(params)
    chart = shape.Chart
    axis_key = params.axis
    is_secondary = axis_key == "secondary_value"
    axis_type = AXIS_TYPE_MAP[axis_key]
    axis_group = AXIS_GROUP_SECONDARY if is_secondary else AXIS_GROUP_PRIMARY
    try:
        ax = chart.Axes(axis_type, axis_group)
    except com_error as exc:
        raise ValueError(
            f"Axis {params.axis!r} is unavailable for this chart type"
        ) from exc
    applied: list[str] = []
    if params.title is not None:
        ax.HasTitle = True
        ax.AxisTitle.Text = params.title
        applied.append("title")
    properties = {
        "min_scale": "MinimumScale",
        "max_scale": "MaximumScale",
        "major_unit": "MajorUnit",
        "minor_unit": "MinorUnit",
        "tick_label_spacing": "TickLabelSpacing",
        "tick_mark_spacing": "TickMarkSpacing",
        "reverse_order": "ReversePlotOrder",
    }
    for field, attribute in properties.items():
        value = getattr(params, field)
        if value is not None:
            setattr(ax, attribute, value)
            applied.append(field)
    if params.major_tick_mark is not None:
        ax.MajorTickMark = TICK_MARK_MAP[params.major_tick_mark]
        applied.append("major_tick_mark")
    if params.minor_tick_mark is not None:
        ax.MinorTickMark = TICK_MARK_MAP[params.minor_tick_mark]
        applied.append("minor_tick_mark")
    if params.log_scale is not None:
        ax.ScaleType = SCALE_LOGARITHMIC if params.log_scale else SCALE_LINEAR
        applied.append("log_scale")
    if params.log_base is not None:
        ax.LogBase = params.log_base
        applied.append("log_base")
    if params.tick_label_font_size is not None:
        ax.TickLabels.Font.Size = params.tick_label_font_size
        applied.append("tick_label_font_size")
    if params.number_format is not None:
        ax.TickLabels.NumberFormat = params.number_format
        applied.append("number_format")
    return {
        "success": True,
        "shape_name": shape.Name,
        "axis": axis_key,
        "applied": applied,
    }


@tool(
    "wps_ppt_charts",
    component="wpp",
    annotations={"title": "Set Chart Series Format", "idempotentHint": True},
    operation="set_series",
)
def set_chart_series(params: SetChartSeriesInput) -> JsonObject:
    """Format an individual chart series.

    Set the fill color (as '#RRGGBB'), toggle data labels, and set line
    weight (in points, for line/scatter charts). Series
    are 1-based indexed.
    """
    shape = _get_chart_shape(params)
    chart = shape.Chart
    if params.series_index > chart.SeriesCollection().Count:
        raise ValueError("series_index exceeds the number of chart series")
    series = chart.SeriesCollection(params.series_index)
    if params.color is not None:
        series.Format.Fill.ForeColor.RGB = hex_to_int(params.color)
    if params.show_data_labels is not None:
        series.HasDataLabels = params.show_data_labels
    if params.line_weight is not None:
        series.Format.Line.Weight = params.line_weight
    return {
        "success": True,
        "shape_name": shape.Name,
        "series_index": params.series_index,
    }


@tool(
    "wps_ppt_charts",
    component="wpp",
    annotations={"title": "Change Chart Type", "idempotentHint": True},
    operation="set_type",
)
def change_chart_type(params: ChangeChartTypeInput) -> JsonObject:
    """Change the type of an existing chart.

    Accepts a friendly name (e.g. 'bar', 'line', 'pie') or XlChartType
    integer. The chart data is preserved when changing types.
    """
    shape = _get_chart_shape(params)
    chart = shape.Chart
    type_int = _resolve_chart_type(params.chart_type)
    chart.ChartType = type_int
    type_name = CHART_TYPE_NAMES.get(type_int, str(type_int))
    return {
        "success": True,
        "shape_name": shape.Name,
        "new_chart_type": type_name,
        "new_chart_type_int": type_int,
    }
