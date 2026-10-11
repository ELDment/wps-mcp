"""ET uses strict bounded rectangles and explicit destructive operations."""

from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from pydantic import ValidationError

from src.spreadsheet import addresses, matrices, models, ranges, sheets


@pytest.fixture
def target(monkeypatch: pytest.MonkeyPatch) -> Mock:
    current = Mock()
    current.Address = "$A$1:$B$1"
    current.Rows.Count = 1
    current.Columns.Count = 2
    current.MergeCells = False
    sheet = SimpleNamespace(Name="Sheet1")
    monkeypatch.setattr(
        ranges, "resolve_range", lambda *args: (sheet, current)
    )
    return current


@pytest.mark.parametrize(
    "address", ["A:A", "B2:A1", "A0", "XFE1", "A1048577", "A1,B2", "Sheet1!A1"]
)
def test_invalid_addresses_are_rejected(address: str) -> None:
    with pytest.raises(ValidationError):
        ranges.ReadRangeInput(address=address)


def test_address_and_matrix_boundaries() -> None:
    assert addresses.parse_address("$Z$1:$AA$2") == (1, 26, 2, 27)
    with pytest.raises(ValidationError, match="exact target range"):
        ranges.WriteValuesInput(address="A1:B2", values=[[1, 2]])
    with pytest.raises(ValidationError, match="limited"):
        ranges.WriteValuesInput(address="A1:A100001", values=[[1]])
    with pytest.raises(ValidationError):
        ranges.WriteValuesInput(address="A1", values=[[float("inf")]])


def test_strict_selectors_and_counts_reject_bool() -> None:
    with pytest.raises(ValidationError):
        models.SheetInput.model_validate({"sheet": True})
    with pytest.raises(ValidationError):
        ranges.EditBandInput.model_validate(
            {"axis": "rows", "action": "insert", "start": True}
        )
    with pytest.raises(ValidationError):
        ranges.ReadRangeInput.model_validate(
            {"address": "A1", "max_cells": "100"}
        )


def test_write_literal_values_never_promotes_formulas(target: Mock) -> None:
    result = ranges.write_values(
        ranges.WriteValuesInput(address="A1:B1", values=[["=1+1", " 001 "]])
    )
    assert target.Value2 == (("'=1+1", "' 001 "),)
    assert result["mode"] == "literal"


def test_write_values_retains_false_zero_and_null(target: Mock) -> None:
    ranges.write_values(
        ranges.WriteValuesInput(address="A1:C1", values=[[False, 0, None]])
    )
    assert target.Value2 == ((False, 0, None),)


def test_formulas_are_explicit(target: Mock) -> None:
    with pytest.raises(ValidationError, match="begin with"):
        ranges.WriteFormulasInput(address="A1", formulas=[["1+1"]])
    ranges.write_formulas(
        ranges.WriteFormulasInput(address="A1", formulas=[["=1+1"]])
    )
    assert target.Formula == "=1+1"


def test_merged_write_is_rejected(target: Mock) -> None:
    target.MergeCells = None
    before = target.Value2
    with pytest.raises(ValueError, match="merged cells"):
        ranges.write_values(
            ranges.WriteValuesInput(address="A1:B1", values=[[1, 2]])
        )
    assert target.Value2 is before


def test_merge_checks_formulas_even_if_display_is_empty(target: Mock) -> None:
    target.Formula = (("keep", '=IF(TRUE,"",1)'),)
    with pytest.raises(ValueError, match="discard"):
        ranges.merge_range(ranges.MergeRangeInput(address="A1:B1"))
    target.Merge.assert_not_called()


def test_merge_restores_alerts_after_native_failure(
    monkeypatch: pytest.MonkeyPatch, target: Mock
) -> None:
    application = SimpleNamespace(DisplayAlerts=7)
    session = SimpleNamespace(get_application=lambda: application)
    for module in (ranges, sheets):
        monkeypatch.setattr(module, "get_session", lambda component: session)
    target.Merge.side_effect = RuntimeError("native merge failure")
    with pytest.raises(RuntimeError, match="native merge failure"):
        ranges.merge_range(
            ranges.MergeRangeInput(address="A1:B1", discard_other_cells=True)
        )
    assert application.DisplayAlerts == 7
    assert target.Merge.call_count == 1


def test_delete_last_visible_sheet_is_rejected(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    sheet = Mock()
    sheet.Visible = -1
    workbook = Mock()
    workbook.Sheets.Count = 1
    workbook.Sheets.Item.return_value = sheet
    session = SimpleNamespace(get_document=lambda document: workbook)
    for module in (ranges, sheets):
        monkeypatch.setattr(module, "get_session", lambda component: session)
    monkeypatch.setattr(sheets, "get_sheet", lambda workbook, reference: sheet)
    with pytest.raises(ValueError, match="last visible"):
        sheets.delete_sheet(models.SheetInput())
    sheet.Delete.assert_not_called()


def test_sort_and_filter_reject_invalid_relative_columns() -> None:
    with pytest.raises(ValidationError, match="key_column"):
        ranges.SortRangeInput(address="A1:B3", key_column=3)
    with pytest.raises(ValidationError, match="outside"):
        ranges.FilterRangeInput(address="A1:B3", field=3)
    with pytest.raises(ValidationError, match="header"):
        ranges.FilterRangeInput(address="A1:B1")


def test_find_stops_at_wraparound(target: Mock) -> None:
    first = SimpleNamespace(Address="$A$1", Value2="text", Formula="text")
    second = SimpleNamespace(Address="$B$1", Value2="text", Formula="text")
    target.Find.side_effect = [first, second, first]
    result = ranges.find_range(
        ranges.FindRangeInput(address="A1:B1", text="text")
    )
    assert len(result["matches"]) == 2
    assert result["truncated"] is False


def test_find_reports_truncation(target: Mock) -> None:
    first = SimpleNamespace(Address="$A$1", Value2="text", Formula="text")
    second = SimpleNamespace(Address="$B$1", Value2="text", Formula="text")
    target.Find.side_effect = [first, second]
    result = ranges.find_range(
        ranges.FindRangeInput(address="A1:B1", text="text", max_matches=1)
    )
    assert len(result["matches"]) == 1
    assert result["truncated"] is True


def test_matrix_result_keeps_scalar_false_and_empty_shape() -> None:
    assert matrices.matrix_result(False, 1, 1) == [[False]]
    assert matrices.matrix_result(None, 2, 2) == [[None, None], [None, None]]
    assert matrices.matrix_result(((False,),), 1, 1) == [[False]]


def test_matrix_result_rejects_invalid_native_shapes() -> None:
    with pytest.raises(RuntimeError, match="scalar"):
        matrices.matrix_result("unexpected", 1, 2)
    with pytest.raises(RuntimeError, match="do not match"):
        matrices.matrix_result(((1, 2),), 2, 2)


def test_color_contract_matches_writer() -> None:
    assert (
        ranges.FormatRangeInput(address="A1", font_color="#123AbC").font_color
        == "#123AbC"
    )
    with pytest.raises(ValidationError):
        ranges.FormatRangeInput(address="A1", font_color="#ABC")


def test_sheet_name_validation_preserves_significant_spaces() -> None:
    assert sheets.AddSheetInput(name=" Sheet ").name == " Sheet "
    with pytest.raises(ValidationError):
        sheets.AddSheetInput(name="'Sheet")
    with pytest.raises(ValidationError):
        sheets.AddSheetInput(name="bad/name")


@pytest.mark.parametrize("row", ["ab", b"ab", 1, None])
def test_matrix_result_rejects_scalar_rows(row: object) -> None:
    with pytest.raises(RuntimeError, match="invalid matrix row"):
        matrices.matrix_result((row,), 1, 2)
    assert matrices.matrix_result((("ab",),), 1, 1) == [["ab"]]


@pytest.mark.parametrize("value", [object(), b"ab", bytearray(b"ab")])
def test_matrix_result_rejects_invalid_native_containers(
    value: object,
) -> None:
    with pytest.raises(RuntimeError, match="invalid cell matrix"):
        matrices.matrix_result(value, 1, 2)


def test_matrix_result_rejects_native_objects_inside_cells() -> None:
    with pytest.raises(RuntimeError, match="invalid cell value"):
        matrices.matrix_result(((SimpleNamespace(),),), 1, 1)


@pytest.mark.parametrize(
    ("address", "max_cells"), [("A11", 100), ("K1", 100), ("A1:B2", 3)]
)
def test_native_range_limits_precede_lookup(
    address: str, max_cells: int
) -> None:
    from src.spreadsheet.access import get_range

    sheet = Mock()
    sheet.Rows.Count = sheet.Columns.Count = 10
    with pytest.raises(ValueError, match="dimensions|cell limit"):
        get_range(sheet, address, max_cells)
    sheet.Range.assert_not_called()


def test_read_range_does_not_read_unrequested_formulas(target: Mock) -> None:
    from unittest.mock import PropertyMock

    formula = PropertyMock(side_effect=RuntimeError("unexpected formula read"))
    type(target).Formula = formula
    target.Value2 = ((0, False),)
    result = ranges.read_range(ranges.ReadRangeInput(address="A1:B1"))
    assert result["values"] == [[0, False]]
    assert "formulas" not in result
    formula.assert_not_called()
