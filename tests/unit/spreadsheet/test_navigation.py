"""Verify chart and worksheet preflight before native UI or data mutations."""

from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from wps_mcp.spreadsheet import charts, sheets


@pytest.fixture
def workbook(monkeypatch: pytest.MonkeyPatch) -> Mock:
    book = Mock()
    sheet = Mock()
    sheet.Name = "Data"
    sheet.Visible = -1
    sheet.Rows.Count = 100
    sheet.Columns.Count = 20
    book.Worksheets.Count = 1
    book.Worksheets.Item.return_value = sheet
    book.Windows.Count = 1
    session = SimpleNamespace(get_document=lambda document: book)
    for module in (sheets, charts):
        monkeypatch.setattr(module, "get_session", lambda component: session)
    return book


@pytest.mark.parametrize(
    ("rows", "columns", "visible", "windows", "message"),
    [
        (100, 0, -1, 1, "scrollable"),
        (0, 20, -1, 1, "scrollable"),
        (1, 0, 0, 1, "hidden"),
        (1, 0, -1, 0, "window"),
    ],
)
def test_freeze_failures_do_not_activate_native_ui(
    workbook: Mock,
    rows: int,
    columns: int,
    visible: int,
    windows: int,
    message: str,
) -> None:
    sheet = workbook.Worksheets.Item.return_value
    sheet.Visible = visible
    workbook.Windows.Count = windows
    with pytest.raises(ValueError, match=message):
        sheets.freeze_panes(sheets.FreezeInput(rows=rows, columns=columns))
    workbook.Windows.Item.assert_not_called()
    sheet.Activate.assert_not_called()


def test_freeze_zero_counts_unfreezes_selected_window(workbook: Mock) -> None:
    sheet = workbook.Worksheets.Item.return_value
    window = workbook.Windows.Item.return_value
    window.FreezePanes = True
    result = sheets.freeze_panes(sheets.FreezeInput())
    assert result == {"sheet": "Data", "rows": 0, "columns": 0}
    window.Activate.assert_called_once_with()
    sheet.Activate.assert_called_once_with()
    assert window.SplitRow == window.SplitColumn == 0
    assert window.FreezePanes is False


def test_chart_duplicate_name_precedes_native_add(
    workbook: Mock, monkeypatch: pytest.MonkeyPatch
) -> None:
    sheet = workbook.Worksheets.Item.return_value
    collection = sheet.ChartObjects.return_value
    collection.Count = 1
    collection.Item.return_value.Name = "SUMMARY"
    monkeypatch.setattr(
        charts, "resolve_range", lambda params: (sheet, Mock())
    )
    with pytest.raises(ValueError, match="already uses"):
        charts.add_chart(charts.AddChartInput(address="A1:B3", name="Summary"))
    collection.Add.assert_not_called()


def test_chart_source_bounds_precede_configuration(workbook: Mock) -> None:
    sheet = workbook.Worksheets.Item.return_value
    collection = sheet.ChartObjects.return_value
    collection.Count = 1
    chart = collection.Item.return_value.Chart
    chart.ChartType = 4
    chart.HasTitle = False
    with pytest.raises(ValueError, match="dimensions"):
        charts.set_chart(
            charts.SetChartInput(
                address="A101", chart_type="pie", title="Changed"
            )
        )
    sheet.Range.assert_not_called()
    chart.SetSourceData.assert_not_called()
    assert chart.ChartType == 4
    assert chart.HasTitle is False
