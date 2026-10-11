"""Create and configure native worksheet chart objects."""

from typing import Literal

from pydantic import (
    Field,
    field_validator,
)

from src.models import JsonObject, NativeObject
from src.native.runtime import get_session
from src.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)
from src.spreadsheet.access import get_range, get_sheet, resolve_range
from src.spreadsheet.addresses import A1Address
from src.spreadsheet.models import RangeInput, SheetInput, SheetReference

CHART_TYPES = {
    "column": 51,
    "stacked_column": 52,
    "bar": 57,
    "line": 4,
    "pie": 5,
    "doughnut": -4120,
    "area": 1,
    "scatter": -4169,
}


class AddChartInput(RangeInput):
    """Create a chart from a rectangle with explicit point geometry."""

    chart_type: Literal[
        "column",
        "stacked_column",
        "bar",
        "line",
        "pie",
        "doughnut",
        "area",
        "scatter",
    ] = "column"
    name: str | None = Field(default=None, min_length=1)
    title: str | None = None
    left: float = Field(default=30, ge=0)
    top: float = Field(default=30, ge=0)
    width: float = Field(default=480, gt=0)
    height: float = Field(default=300, gt=0)
    series_in_rows: bool = False


class ChartInput(SheetInput):
    """Select a worksheet chart by name or one-based index."""

    chart: SheetReference = Field(
        default=1, description="Chart object name or one-based index"
    )

    @field_validator("chart")
    @classmethod
    def validate_chart(cls, value: SheetReference) -> SheetReference:
        """Reject empty chart names and nonpositive chart indices."""
        if isinstance(value, int) and value < 1:
            raise ValueError("Chart index must be at least 1")
        if isinstance(value, str) and (not value):
            raise ValueError("Chart name must not be empty")
        return value


class SetChartInput(ChartInput):
    """Change an existing chart source, type, title, or legend."""

    address: A1Address | None = None
    chart_type: (
        Literal[
            "column",
            "stacked_column",
            "bar",
            "line",
            "pie",
            "doughnut",
            "area",
            "scatter",
        ]
        | None
    ) = None
    title: str | None = None
    show_legend: bool | None = None
    series_in_rows: bool = False


def get_chart(sheet: NativeObject, reference: SheetReference) -> NativeObject:
    """Resolve a native chart by case-insensitive name or one-based index."""
    charts = sheet.ChartObjects()
    if isinstance(reference, int):
        if reference < 1 or reference > charts.Count:
            raise ValueError(
                f"Chart index {reference} is outside 1..{charts.Count}"
            )
        return charts.Item(reference)
    for index in range(1, charts.Count + 1):
        chart = charts.Item(index)
        if str(chart.Name).casefold() == reference.casefold():
            return chart
    raise ValueError(f"Chart not found: {reference}")


@tool(
    "wps_et_charts",
    operation="list",
    component="et",
    annotations=READ_ANNOTATIONS,
)
def list_charts(params: SheetInput) -> JsonObject:
    """List chart objects on a worksheet."""
    sheet = get_sheet(
        get_session("et").get_document(params.document), params.sheet
    )
    charts = sheet.ChartObjects()
    return {
        "sheet": str(sheet.Name),
        "charts": [
            {
                "index": index,
                "name": str(charts.Item(index).Name),
                "chart_type": int(charts.Item(index).Chart.ChartType),
            }
            for index in range(1, charts.Count + 1)
        ],
    }


@tool(
    "wps_et_charts",
    operation="add",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def add_chart(params: AddChartInput) -> JsonObject:
    """Create a native WPS chart from a worksheet range.

    Geometry is in points.
    """
    sheet, source = resolve_range(params)
    charts = sheet.ChartObjects()
    if params.name is not None and any(
        str(charts.Item(index).Name).casefold() == params.name.casefold()
        for index in range(1, charts.Count + 1)
    ):
        raise ValueError(f"A chart already uses the name: {params.name}")
    chart_object = charts.Add(
        params.left, params.top, params.width, params.height
    )
    if params.name is not None:
        chart_object.Name = params.name
    chart_object.Chart.ChartType = CHART_TYPES[params.chart_type]
    chart_object.Chart.SetSourceData(source, 1 if params.series_in_rows else 2)
    if params.title is not None:
        chart_object.Chart.HasTitle = bool(params.title)
        if params.title:
            chart_object.Chart.ChartTitle.Text = params.title
    return {"sheet": str(sheet.Name), "chart": str(chart_object.Name)}


@tool(
    "wps_et_charts",
    operation="configure",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def set_chart(params: SetChartInput) -> JsonObject:
    """Update a chart's source range, chart type, title or legend."""
    sheet = get_sheet(
        get_session("et").get_document(params.document), params.sheet
    )
    chart_object = get_chart(sheet, params.chart)
    source = (
        get_range(sheet, params.address)
        if params.address is not None
        else None
    )
    chart = chart_object.Chart
    if source is not None:
        chart.SetSourceData(source, 1 if params.series_in_rows else 2)
    if params.chart_type is not None:
        chart.ChartType = CHART_TYPES[params.chart_type]
    if params.title is not None:
        chart.HasTitle = bool(params.title)
        if params.title:
            chart.ChartTitle.Text = params.title
    if params.show_legend is not None:
        chart.HasLegend = params.show_legend
    return {"sheet": str(sheet.Name), "chart": str(chart_object.Name)}


@tool(
    "wps_et_charts",
    operation="delete",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def delete_chart(params: ChartInput) -> JsonObject:
    """Delete a chart object from a worksheet."""
    sheet = get_sheet(
        get_session("et").get_document(params.document), params.sheet
    )
    chart_object = get_chart(sheet, params.chart)
    chart_name = str(chart_object.Name)
    chart_object.Delete()
    return {"sheet": str(sheet.Name), "deleted": chart_name}
