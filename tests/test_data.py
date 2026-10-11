"""Data safety and resource ownership tests for presentation tables/charts."""

from types import SimpleNamespace
from unittest.mock import Mock, PropertyMock

import pytest
from pydantic import ValidationError
from pywintypes import com_error

from src.models import JsonObject
from src.presentation import charts, tables


@pytest.fixture
def chart_environment(
    monkeypatch: pytest.MonkeyPatch,
) -> tuple[Mock, Mock, Mock]:
    chart = Mock()
    workbook = chart.ChartData.Workbook
    worksheet = workbook.Worksheets.return_value
    worksheet.Name = "Sheet1"
    series = Mock(Name="Revenue", XValues=["A", "B"], Values=[10.0, 20.0])
    series.Formula = "=SERIES(Sheet1!$B$1,Sheet1!$A$2:$A$3,Sheet1!$B$2:$B$3,1)"
    collection = Mock(Count=1)
    chart.SeriesCollection.side_effect = lambda index=None: (
        collection if index is None else series
    )
    shape = Mock(Name="Chart 1", HasChart=True, Chart=chart)
    session = Mock()
    monkeypatch.setattr(charts, "get_session", lambda component: session)
    monkeypatch.setattr(charts, "goto_slide", Mock())
    monkeypatch.setattr(
        charts, "resolve_shape", lambda slide, identifier: shape
    )
    return chart, workbook, worksheet


@pytest.fixture
def table_environment(monkeypatch: pytest.MonkeyPatch) -> tuple[Mock, Mock]:
    table = Mock()
    table.Rows.Count = 2
    table.Columns.Count = 2
    shape = Mock(Name="Table 1", HasTable=True, Table=table)
    session = Mock()
    monkeypatch.setattr(tables, "get_session", lambda component: session)
    monkeypatch.setattr(tables, "goto_slide", Mock())
    monkeypatch.setattr(
        tables, "resolve_shape", lambda slide, identifier: shape
    )
    return table, session


@pytest.mark.parametrize(
    "value", [True, "12", None, float("nan"), float("inf")]
)
def test_chart_values_must_be_finite_numbers(value: object) -> None:
    with pytest.raises(ValidationError):
        charts.ChartSeriesInput.model_validate(
            {"name": "Sales", "values": [value]}
        )


def test_chart_series_dimensions_are_validated_before_execution() -> None:
    with pytest.raises(ValidationError, match="one value per category"):
        charts.SetChartDataInput.model_validate(
            {
                "slide_index": 1,
                "shape_name_or_index": 1,
                "categories": ["A", "B"],
                "series": [{"name": "Sales", "values": [1]}],
            }
        )


def test_chart_workbook_borrowed_is_never_closed() -> None:
    chart = Mock()
    with charts.chart_workbook(chart) as workbook:
        assert workbook is chart.ChartData.Workbook
    chart.ChartData.Activate.assert_not_called()
    workbook.Close.assert_not_called()


def test_chart_workbook_owned_closes_on_error() -> None:
    data, workbook = Mock(), Mock()
    type(data).Workbook = PropertyMock(
        side_effect=[com_error(-1, "inactive"), workbook]
    )
    chart = SimpleNamespace(ChartData=data)
    with pytest.raises(ValueError, match="operation"):
        with charts.chart_workbook(chart):
            raise ValueError("operation")
    data.Activate.assert_called_once_with()
    workbook.Close.assert_called_once_with(False)


def test_chart_workbook_cleanup_failure_is_reported() -> None:
    data, workbook = Mock(), Mock()
    type(data).Workbook = PropertyMock(
        side_effect=[com_error(-1, "inactive"), workbook]
    )
    workbook.Close.side_effect = RuntimeError("cleanup")
    with pytest.raises(RuntimeError, match="cleanup"):
        with charts.chart_workbook(SimpleNamespace(ChartData=data)):
            pass


def test_chart_workbook_cleanup_preserves_original_error(
    caplog: pytest.LogCaptureFixture,
) -> None:
    data, workbook = Mock(), Mock()
    type(data).Workbook = PropertyMock(
        side_effect=[com_error(-1, "inactive"), workbook]
    )
    workbook.Close.side_effect = RuntimeError("cleanup")
    with pytest.raises(ValueError, match="operation"):
        with charts.chart_workbook(SimpleNamespace(ChartData=data)):
            raise ValueError("operation")
    assert "Cannot close the chart workbook" in caplog.text


@pytest.mark.parametrize(
    "formula",
    [
        "=SERIES(Sheet2!$B$1,Sheet2!$A$2:$A$3,Sheet2!$B$2:$B$3,1)",
        "=SERIES(Sheet1!$B$1,Sheet1!$A$2:$A$3,Sheet1!$C$2:$C$3,1)",
        '=SERIES("Revenue",{1,2},{10,20},1)',
    ],
)
def test_chart_unsupported_sources_are_rejected_before_writing(
    chart_environment: tuple[Mock, Mock, Mock],
    formula: str,
) -> None:
    chart, workbook, worksheet = chart_environment
    chart.SeriesCollection(1).Formula = formula
    params = charts.SetChartDataInput.model_validate(
        {
            "slide_index": 1,
            "shape_name_or_index": "Chart 1",
            "categories": ["A", "B"],
            "series": [{"name": "Revenue", "values": [10, 20]}],
        }
    )
    with pytest.raises(NotImplementedError, match="Unsupported chart source"):
        charts.set_chart_data(params)
    worksheet.Cells.assert_not_called()
    workbook.Close.assert_not_called()


def test_chart_resize_is_rejected_before_writing(
    chart_environment: tuple[Mock, Mock, Mock],
) -> None:
    _, _, worksheet = chart_environment
    params = charts.SetChartDataInput.model_validate(
        {
            "slide_index": 1,
            "shape_name_or_index": 1,
            "categories": ["A"],
            "series": [{"name": "Revenue", "values": [10]}],
        }
    )
    with pytest.raises(NotImplementedError, match="resize chart categories"):
        charts.set_chart_data(params)
    worksheet.Cells.assert_not_called()


def test_chart_write_escapes_labels_and_verifies_plotted_values(
    chart_environment: tuple[Mock, Mock, Mock],
) -> None:
    chart, workbook, worksheet = chart_environment
    cells: dict[tuple[int, int], Mock] = {}

    def get_cell(row: int, column: int) -> Mock:
        return cells.setdefault((row, column), Mock())

    worksheet.Cells.side_effect = get_cell
    chart.SeriesCollection(1).XValues = ["=1+1", " B "]
    chart.SeriesCollection(1).Name = "=SUM(A1:A2)"
    params = charts.SetChartDataInput.model_validate(
        {
            "slide_index": 1,
            "shape_name_or_index": 1,
            "categories": ["=1+1", " B "],
            "series": [{"name": "=SUM(A1:A2)", "values": [10, 20]}],
        }
    )
    assert charts.set_chart_data(params)["verified"] is True
    assert cells[(2, 1)].Value2 == "'=1+1"
    assert cells[(3, 1)].Value2 == "' B "
    assert cells[(1, 2)].Value2 == "'=SUM(A1:A2)"
    assert cells[(2, 2)].Value2 == 10
    workbook.Close.assert_not_called()


def test_chart_stale_plotted_data_does_not_report_success(
    chart_environment: tuple[Mock, Mock, Mock],
) -> None:
    params = charts.SetChartDataInput.model_validate(
        {
            "slide_index": 1,
            "shape_name_or_index": 1,
            "categories": ["A", "B"],
            "series": [{"name": "Revenue", "values": [99, 20]}],
        }
    )
    with pytest.raises(RuntimeError, match="plotted data differs"):
        charts.set_chart_data(params)


def test_chart_layout_accepts_quoted_sheet_names(
    chart_environment: tuple[Mock, Mock, Mock],
) -> None:
    chart, _, worksheet = chart_environment
    worksheet.Name = "Q1's sales"
    sheet = "'Q1''s sales'"
    chart.SeriesCollection(
        1
    ).Formula = f"=SERIES({sheet}!$B$1,{sheet}!$A$2:$A$3,{sheet}!$B$2:$B$3,1)"
    charts._validate_chart_layout(chart, worksheet, 2, 1)


@pytest.mark.parametrize(
    "properties",
    [
        {"axis": "category", "min_scale": 1},
        {"axis": "value", "tick_label_spacing": 2},
        {"axis": "value", "log_base": 10},
        {"axis": "value", "log_scale": True, "min_scale": 0},
        {"axis": "value", "min_scale": 10, "max_scale": 1},
    ],
)
def test_invalid_chart_axis_combinations(properties: JsonObject) -> None:
    with pytest.raises(ValidationError):
        charts.FormatChartAxisInput.model_validate(
            {
                "slide_index": 1,
                "shape_name_or_index": 1,
                **properties,
            }
        )


@pytest.mark.parametrize("data", [[], [[]], [["A"], []], [["A"], ["B", "C"]]])
def test_table_batch_requires_rectangular_nonempty_data(
    data: list[list[str]],
) -> None:
    with pytest.raises(ValidationError):
        tables.SetTableDataInput(
            slide_index=1, shape_name_or_index=1, data=data
        )


@pytest.mark.parametrize("start_row,start_col", [(2, 1), (1, 2), (3, 3)])
def test_table_batch_overflow_performs_no_writes(
    table_environment: tuple[Mock, Mock],
    start_row: int,
    start_col: int,
) -> None:
    table, _ = table_environment
    params = tables.SetTableDataInput(
        slide_index=1,
        shape_name_or_index=1,
        start_row=start_row,
        start_col=start_col,
        data=[["A", "B"], ["C", "D"]],
    )
    with pytest.raises(ValueError, match="exceeds table bounds"):
        tables.set_table_data(params)
    table.Cell.assert_not_called()


def test_table_batch_preserves_whitespace_and_normalizes_newlines(
    table_environment: tuple[Mock, Mock],
) -> None:
    table, session = table_environment
    cells = {(row, column): Mock() for row in (1, 2) for column in (1, 2)}
    table.Cell.side_effect = lambda row, column: cells[(row, column)]
    params = tables.SetTableDataInput(
        document="selected.pptx",
        slide_index=1,
        shape_name_or_index=1,
        data=[[" A ", "B\r\nC"], ["D", "E\nF"]],
        bold_first_row=True,
    )
    result = tables.set_table_data(params)
    assert result["cells_set"] == 4
    assert cells[(1, 1)].Shape.TextFrame.TextRange.Text == " A "
    assert cells[(1, 2)].Shape.TextFrame.TextRange.Text == "B\rC"
    assert cells[(2, 2)].Shape.TextFrame.TextRange.Text == "E\rF"
    assert cells[(1, 1)].Shape.TextFrame.TextRange.Font.Bold == -1
    session.get_document.assert_called_once_with("selected.pptx")


def test_table_layout_checks_both_dimensions_before_editing(
    table_environment: tuple[Mock, Mock],
) -> None:
    table, _ = table_environment
    params = tables.SetTableLayoutInput(
        slide_index=1,
        shape_name_or_index=1,
        row_heights=[40],
        col_widths=[10, 20, 30],
    )
    with pytest.raises(ValueError, match="col_widths exceeds"):
        tables.set_table_layout(params)
    table.Rows.assert_not_called()
    table.Columns.assert_not_called()


def test_table_creation_rejects_invalid_dimensions() -> None:
    with pytest.raises(ValidationError):
        tables.AddTableInput(
            slide_index=1, rows=2, cols=2, row_heights=[10, -1]
        )
    with pytest.raises(ValidationError, match="row_heights exceeds"):
        tables.AddTableInput(
            slide_index=1, rows=2, cols=2, row_heights=[10, 20, 30]
        )


def test_table_diagonal_visibility_limit_is_reported(
    table_environment: tuple[Mock, Mock],
) -> None:
    table, _ = table_environment
    border = SimpleNamespace(Visible=0)
    table.Cell.return_value.Borders.Item.return_value = border
    result = tables.set_table_borders(
        tables.SetTableBordersInput(
            slide_index=1,
            shape_name_or_index=1,
            sides=["diagonal_down"],
            visible=True,
        )
    )
    assert border.Visible == 0
    assert result["cells_updated"] == 4
    assert result["warnings"]


def test_new_chart_owns_automatically_opened_workbook() -> None:
    chart = Mock()
    with charts.chart_workbook(chart, own_existing=True) as workbook:
        assert workbook is chart.ChartData.Workbook
    chart.ChartData.Activate.assert_not_called()
    workbook.Close.assert_called_once_with(False)


def test_chart_title_auto_layout_preserves_font(
    chart_environment: tuple[Mock, Mock, Mock],
) -> None:
    chart, _, _ = chart_environment
    chart.HasTitle = True
    chart.HasLegend = True
    title = chart.ChartTitle
    title.Text = "Quarterly sales"
    title.Font = SimpleNamespace(
        Name="Arial", Size=24, Bold=-1, Color=SimpleNamespace(RGB=255)
    )
    result = charts.format_chart(
        charts.FormatChartInput(
            slide_index=1, shape_name_or_index=1, title_position="top"
        )
    )
    assert result["has_title"] is True
    assert title.Text == "Quarterly sales"
    assert title.Font.Name == "Arial"
    assert title.Font.Size == 24
    assert title.Font.Color.RGB == 255


def test_chart_legend_validation_precedes_title_mutation(
    chart_environment: tuple[Mock, Mock, Mock],
) -> None:
    chart, _, _ = chart_environment
    chart.HasLegend = False
    chart.ChartTitle.Text = "Old title"
    with pytest.raises(ValueError, match="has_legend=true"):
        charts.format_chart(
            charts.FormatChartInput(
                slide_index=1,
                shape_name_or_index=1,
                title="New title",
                legend_position="bottom",
            )
        )
    assert chart.ChartTitle.Text == "Old title"
