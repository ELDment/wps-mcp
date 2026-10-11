"""Verify et analysis behavior and native preflight boundaries."""

from collections.abc import Sequence
from types import SimpleNamespace
from unittest.mock import Mock, PropertyMock

import pytest
from pydantic import ValidationError

from src.spreadsheet import addresses, models, pivots, table_access, tables


def native_collection(items: Sequence[Mock | SimpleNamespace]) -> Mock:
    collection = Mock()
    collection.Count = len(items)
    collection.Item.side_effect = lambda index: items[index - 1]
    return collection


def native_range(
    sheet: Mock, top: int, left: int, bottom: int, right: int
) -> Mock:
    target = Mock()
    target.Row = top
    target.Column = left
    target.Rows.Count = bottom - top + 1
    target.Columns.Count = right - left + 1
    target.Parent = sheet
    target.MergeCells = False
    target.Address = f"${chr(64 + left)}${top}:${chr(64 + right)}${bottom}"
    matrix = tuple(
        tuple(
            sheet.grid.get((row, column)) for column in range(left, right + 1)
        )
        for row in range(top, bottom + 1)
    )
    target.Formula = target.Value2 = (
        matrix[0][0] if top == bottom and left == right else matrix
    )
    target.Rows.Item.side_effect = lambda index: native_range(
        sheet, top + index - 1, left, top + index - 1, right
    )
    return target


def native_sheet(name: str, values: list[list[str | int]]) -> Mock:
    sheet = Mock()
    sheet.Name = name
    sheet.grid = {
        (row_index, column_index): value
        for row_index, row in enumerate(values, 1)
        for column_index, value in enumerate(row, 1)
    }
    sheet.Rows.Count = 1048576
    sheet.Columns.Count = 16384
    sheet.ListObjects = native_collection([])
    sheet.PivotTables.return_value = native_collection([])
    sheet.Cells.side_effect = lambda row, column: SimpleNamespace(
        Row=row, Column=column
    )

    def select(
        first: str | SimpleNamespace, last: SimpleNamespace | None = None
    ) -> Mock:
        if isinstance(first, str):
            return native_range(sheet, *addresses.parse_address(first))
        assert last is not None
        return native_range(
            sheet, first.Row, first.Column, last.Row, last.Column
        )

    sheet.Range.side_effect = select
    sheet.UsedRange = native_range(sheet, 1, 1, 30, 26)
    return sheet


@pytest.fixture
def workbook(monkeypatch: pytest.MonkeyPatch) -> Mock:
    workbook = Mock()
    workbook.Name = "Book.xlsx"
    workbook.data = native_sheet(
        "Data",
        [
            ["Region", "Quarter", "Amount"],
            ["North", "Q1", 10],
            ["South", "Q1", 20],
        ],
    )
    workbook.report = native_sheet("Report", [])
    workbook.Worksheets = native_collection([workbook.data, workbook.report])
    workbook.Names = native_collection([])
    session = SimpleNamespace(get_document=lambda document: workbook)
    for module in (pivots, tables):
        monkeypatch.setattr(module, "get_session", lambda component: session)
    return workbook


def native_pivot(workbook: Mock, *, sheet: Mock | None = None) -> Mock:
    target_sheet: Mock = workbook.report if sheet is None else sheet
    pivot = Mock()
    pivot.Name = "Summary"
    pivot.TableRange2 = native_range(target_sheet, 1, 8, 4, 9)
    pivot.TableRange1 = pivot.TableRange2
    pivot.ManualUpdate = False
    pivot.RowGrand = True
    pivot.ColumnGrand = True
    pivot.CacheIndex = 1
    pivot.PivotCache.return_value.SourceType = 1
    pivot.PivotCache.return_value.SourceData = "'Data'!R1C1:R3C3"
    pivot.fields = {
        name: SimpleNamespace(
            Name=name, SourceName=name, Orientation=0, Position=0
        )
        for name in ("Region", "Quarter", "Amount")
    }
    fields = native_collection(list(pivot.fields.values()))
    pivot.PivotFields.side_effect = lambda name=None: (
        fields if name is None else pivot.fields[name]
    )
    pivot.DataFields = native_collection([])
    pivot.RefreshTable.return_value = True
    target_sheet.PivotTables.return_value = native_collection([pivot])
    cache = workbook.PivotCaches.return_value.Create.return_value
    cache.CreatePivotTable.return_value = pivot
    return pivot


def native_table(workbook: Mock) -> Mock:
    table = Mock()
    table.Name = "Sales"
    table.Range = native_range(workbook.data, 1, 1, 3, 3)
    table.ListColumns = native_collection(
        [
            SimpleNamespace(Name=name)
            for name in ("Region", "Quarter", "Amount")
        ]
    )
    table.ShowHeaders = True
    table.ShowTotals = False
    workbook.data.ListObjects = native_collection([table])
    return table


def native_shifted_table(workbook: Mock) -> Mock:
    """Reproduce WPS inserting generated headers above the source."""
    table = native_table(workbook)
    sheet = workbook.data
    sheet.ListObjects = native_collection([])
    generated = ["Column1", "Column2", "Column3"]
    headers = ["Region", "Quarter", "Amount"]
    original = dict(sheet.grid)
    first_row = native_range(sheet, 2, 1, 2, 3)
    first_row.Value2 = first_row.Formula = (tuple(headers),)
    list_row = Mock()
    list_row.Range = first_row
    table.ListRows = native_collection([list_row, Mock(), Mock()])
    table.DataBodyRange.Rows.Item.return_value = first_row
    table.HeaderRowRange = native_range(sheet, 1, 1, 1, 3)

    def create(**arguments: object) -> Mock:
        for (row, column), value in original.items():
            sheet.grid[(row + 1, column)] = value
        for column, value in enumerate(generated, 1):
            sheet.grid[(1, column)] = value
        table.ListColumns = native_collection(
            [SimpleNamespace(Name=name) for name in generated]
        )
        table.Range = native_range(sheet, 1, 1, 4, 3)
        return table

    def cut(**arguments: object) -> None:
        assert arguments["Destination"] is table.HeaderRowRange
        for column, value in enumerate(headers, 1):
            sheet.grid[(1, column)] = value
            sheet.grid[(2, column)] = None
        first_row.Formula = ((None, None, None),)
        table.ListColumns = native_collection(
            [SimpleNamespace(Name=name) for name in headers]
        )

    def delete() -> None:
        assert first_row.Formula == ((None, None, None),)
        for row in range(2, max(key[0] for key in sheet.grid)):
            for column in range(1, 4):
                sheet.grid[(row, column)] = sheet.grid.get((row + 1, column))
        table.Range = native_range(sheet, 1, 1, 3, 3)

    sheet.ListObjects.Add.side_effect = create
    first_row.Cut.side_effect = cut
    list_row.Delete.side_effect = delete
    return table


@pytest.mark.parametrize(
    "name", ["", "bad name", "A1", "R1C1", "R", "C", "1Table", "x" * 256]
)
def test_invalid_object_names_fail_before_native(name: str) -> None:
    with pytest.raises(ValidationError):
        tables.CreateTableInput(name=name, address="A1:C3")


@pytest.mark.parametrize("name", ["Sales", "sales_2026", "Sales.East", "XFE1"])
def test_valid_object_names_are_preserved(name: str) -> None:
    assert tables.CreateTableInput(name=name, address="A1:C3").name == name


@pytest.mark.parametrize(
    "changes",
    [
        {"row_fields": ["Region", "region"]},
        {"row_fields": ["Region"], "page_fields": ["Region"]},
        {"row_fields": [" "]},
        {"data_fields": [{"field": "Amount"}, {"field": "Amount"}]},
        {"data_fields": [{"field": "Amount", "caption": " "}]},
        {"data_fields": [{"field": "Amount", "aggregate": "median"}]},
        {"data_fields": []},
    ],
)
def test_invalid_layouts_fail_before_native(
    changes: dict[str, object],
) -> None:
    with pytest.raises(ValidationError):
        pivots.PivotLayout.model_validate(
            {"data_fields": [{"field": "Amount"}], **changes}
        )


@pytest.mark.parametrize("destination", ["A1:B2", "A:A", "A0", "XFE1"])
def test_pivot_destination_is_one_bounded_cell(destination: str) -> None:
    with pytest.raises(ValidationError):
        pivots.CreatePivotInput(
            name="Summary",
            address="A1:C3",
            destination_sheet=2,
            destination_cell=destination,
            data_fields=[pivots.PivotDataField(field="Amount")],
        )


def test_page_fields_require_space_above_destination() -> None:
    with pytest.raises(ValidationError, match="above"):
        pivots.CreatePivotInput(
            name="Summary",
            address="A1:C3",
            destination_sheet=2,
            destination_cell="H1",
            page_fields=["Quarter"],
            data_fields=[pivots.PivotDataField(field="Amount")],
        )


@pytest.mark.parametrize(
    "header", [None, 12, "", " ", "Region", "region", "=1+1"]
)
def test_invalid_headers_prevent_cache_or_table_creation(
    workbook: Mock, header: object
) -> None:
    workbook.data.grid[(1, 3)] = header
    with pytest.raises(ValueError, match="header|Column"):
        tables.create_table(
            tables.CreateTableInput(name="Sales", address="A1:C3")
        )
    workbook.data.ListObjects.Add.assert_not_called()
    with pytest.raises(ValueError, match="header|Column"):
        pivots.create_pivot(
            pivots.CreatePivotInput(
                name="Summary",
                address="A1:C3",
                destination_sheet="Report",
                destination_cell="H1",
                data_fields=[pivots.PivotDataField(field="Amount")],
            )
        )
    workbook.PivotCaches.assert_not_called()


def test_header_only_table_is_rejected_without_native_auto_row(
    workbook: Mock,
) -> None:
    with pytest.raises(ValueError, match="data row"):
        tables.create_table(
            tables.CreateTableInput(name="Sales", address="A1:C1")
        )
    workbook.data.ListObjects.Add.assert_not_called()


def test_duplicate_names_checked_across_sheets_and_defined_names(
    workbook: Mock,
) -> None:
    table = native_table(workbook)
    with pytest.raises(ValueError, match="already uses"):
        table_access.ensure_available_name(workbook, "sales")
    workbook.Names = native_collection([SimpleNamespace(Name="Data!Taken")])
    with pytest.raises(ValueError, match="defined name"):
        table_access.ensure_available_name(workbook, "Taken")
    table_access.ensure_available_name(workbook, "sales", current=table.Name)


def test_create_pivot_uses_explicit_source_and_native_aggregation(
    workbook: Mock,
) -> None:
    pivot = native_pivot(workbook)
    workbook.report.PivotTables.return_value = native_collection([])
    result = pivots.create_pivot(
        pivots.CreatePivotInput(
            name="Summary",
            address="A1:C3",
            destination_sheet="Report",
            destination_cell="H1",
            row_fields=["Region"],
            column_fields=["Quarter"],
            data_fields=[
                pivots.PivotDataField(
                    field="Amount", aggregate="avg", caption="Average"
                )
            ],
        )
    )
    workbook.PivotCaches.return_value.Create.assert_called_once_with(
        1, "'Data'!R1C1:R3C3"
    )
    assert pivot.fields["Region"].Orientation == 1
    assert pivot.fields["Quarter"].Orientation == 2
    pivot.AddDataField.assert_called_once_with(
        pivot.fields["Amount"], "Average", -4106
    )
    assert pivot.ManualUpdate is False
    assert result["name"] == "Summary"


@pytest.mark.parametrize("value", ["existing", 0, False, '=IF(TRUE,"",1)'])
def test_pivot_create_protects_existing_destination_values(
    workbook: Mock, value: object
) -> None:
    workbook.report.grid[(20, 15)] = value
    with pytest.raises(ValueError, match="existing content"):
        pivots.create_pivot(
            pivots.CreatePivotInput(
                name="Summary",
                address="A1:C3",
                destination_sheet="Report",
                destination_cell="H1",
                data_fields=[pivots.PivotDataField(field="Amount")],
            )
        )
    workbook.PivotCaches.assert_not_called()


def test_same_sheet_source_is_allowed_only_left_of_reserved_columns(
    workbook: Mock,
) -> None:
    with pytest.raises(ValueError, match="overlaps reserved"):
        pivots.create_pivot(
            pivots.CreatePivotInput(
                name="Summary",
                address="A1:C3",
                destination_sheet="Data",
                destination_cell="B20",
                data_fields=[pivots.PivotDataField(field="Amount")],
            )
        )
    workbook.PivotCaches.assert_not_called()
    pivot = native_pivot(workbook, sheet=workbook.data)
    workbook.data.PivotTables.return_value = native_collection([])
    assert (
        pivots.create_pivot(
            pivots.CreatePivotInput(
                name="Summary",
                address="A1:C3",
                destination_sheet="Data",
                destination_cell="H1",
                data_fields=[pivots.PivotDataField(field="Amount")],
            )
        )["name"]
        == pivot.Name
    )


@pytest.mark.parametrize(
    "field,caption", [("Missing", None), ("Amount", "region")]
)
def test_invalid_pivot_fields_fail_before_native_create(
    workbook: Mock,
    field: str,
    caption: str | None,
) -> None:
    with pytest.raises(ValueError, match="field"):
        pivots.create_pivot(
            pivots.CreatePivotInput(
                name="Summary",
                address="A1:C3",
                destination_sheet="Report",
                destination_cell="H1",
                data_fields=[
                    pivots.PivotDataField(field=field, caption=caption)
                ],
            )
        )
    workbook.PivotCaches.assert_not_called()


def test_configure_replaces_layout_and_restores_manual_update_on_failure(
    workbook: Mock,
) -> None:
    pivot = native_pivot(workbook)
    pivot.fields["Region"].Orientation = 1
    pivot.AddDataField.side_effect = RuntimeError("native failure")
    with pytest.raises(RuntimeError, match="native failure"):
        pivots.configure_pivot(
            pivots.ConfigurePivotInput(
                sheet="Report",
                pivot="Summary",
                column_fields=["Quarter"],
                data_fields=[pivots.PivotDataField(field="Amount")],
            )
        )
    assert pivot.ManualUpdate is False
    assert pivot.fields["Region"].Orientation == 0
    assert pivot.fields["Quarter"].Orientation == 2
    assert pivot.AddDataField.call_count == 1


def test_configure_replaces_separate_native_data_fields_in_reverse_order(
    workbook: Mock,
) -> None:
    pivot = native_pivot(workbook)
    active: list[Mock | SimpleNamespace] = []
    removed: list[str] = []

    def add_existing(caption: str) -> Mock:
        field = Mock()
        field.Name = caption

        def hide(orientation: int) -> None:
            assert orientation == 0
            assert active[-1] is field
            active.remove(field)
            removed.append(caption)
            pivot.DataFields.Count = len(active)

        type(field).Orientation = PropertyMock(side_effect=hide)
        return field

    active.extend(
        [add_existing("sum of Amount"), add_existing("count of Amount")]
    )
    pivot.DataFields = native_collection(active)

    def add_field(
        field: SimpleNamespace, caption: str, aggregate: int
    ) -> None:
        assert removed == ["count of Amount", "sum of Amount"]
        active.append(
            SimpleNamespace(
                SourceName=field.SourceName, Name=caption, Function=aggregate
            )
        )
        pivot.DataFields.Count = len(active)

    pivot.AddDataField.side_effect = add_field
    pivots.configure_pivot(
        pivots.ConfigurePivotInput(
            sheet="Report",
            pivot="Summary",
            row_fields=["Region"],
            data_fields=[
                pivots.PivotDataField(field="Amount", aggregate="avg")
            ],
        )
    )
    result = pivots.read_pivot(
        pivots.ReadPivotInput(sheet="Report", pivot="Summary")
    )
    assert result["data_fields"] == [
        {"field": "Amount", "caption": "avg of Amount", "aggregate": "avg"},
    ]
    assert pivot.ManualUpdate is False


def test_data_field_enumeration_failure_cannot_modify_layout(
    workbook: Mock,
) -> None:
    pivot = native_pivot(workbook)
    pivot.fields["Region"].Orientation = 1
    pivot.DataFields.Count = 2
    pivot.DataFields.Item.side_effect = [
        Mock(),
        RuntimeError("native data field unavailable"),
    ]
    with pytest.raises(RuntimeError, match="native data field unavailable"):
        pivots.configure_pivot(
            pivots.ConfigurePivotInput(
                sheet="Report",
                pivot="Summary",
                data_fields=[pivots.PivotDataField(field="Amount")],
            )
        )
    assert pivot.fields["Region"].Orientation == 1
    assert pivot.ManualUpdate is False
    pivot.AddDataField.assert_not_called()


def test_data_field_removal_failure_restores_state_without_retry(
    workbook: Mock,
) -> None:
    pivot = native_pivot(workbook)
    field = Mock()
    orientation = PropertyMock(side_effect=RuntimeError("native hide failure"))
    type(field).Orientation = orientation
    pivot.DataFields = native_collection([field])
    with pytest.raises(RuntimeError, match="native hide failure"):
        pivots.configure_pivot(
            pivots.ConfigurePivotInput(
                sheet="Report",
                pivot="Summary",
                data_fields=[pivots.PivotDataField(field="Amount")],
            )
        )
    orientation.assert_called_once_with(0)
    assert pivot.ManualUpdate is False
    pivot.AddDataField.assert_not_called()


@pytest.mark.parametrize(
    "position", ["row_fields", "column_fields", "page_fields", "data_fields"]
)
def test_stale_cache_field_cannot_clear_existing_layout(
    workbook: Mock, position: str
) -> None:
    pivot = native_pivot(workbook)
    pivot.fields["Region"].Orientation = 1
    pivot.fields["Region"].Position = 1
    pivot.TableRange1 = native_range(workbook.report, 5, 8, 8, 9)
    workbook.data.grid[(1, 1)] = "Area"
    layout: dict[str, object] = {
        "sheet": "Report",
        "pivot": "Summary",
        "data_fields": [{"field": "Amount"}],
    }
    layout[position] = (
        [{"field": "Area"}] if position == "data_fields" else ["Area"]
    )
    before = {name: vars(field).copy() for name, field in pivot.fields.items()}
    with pytest.raises(
        ValueError, match="cache field.*Area.*refresh operation first"
    ):
        pivots.configure_pivot(
            pivots.ConfigurePivotInput.model_validate(layout)
        )
    assert {
        name: vars(field) for name, field in pivot.fields.items()
    } == before
    assert pivot.ManualUpdate is False
    pivot.AddDataField.assert_not_called()
    pivot.RefreshTable.assert_not_called()


def test_cache_caption_collision_fails_before_layout_changes(
    workbook: Mock,
) -> None:
    pivot = native_pivot(workbook)
    pivot.fields["Region"].Orientation = 1
    workbook.data.grid[(1, 1)] = "Area"
    with pytest.raises(
        ValueError, match="caption conflicts with a cached field"
    ):
        pivots.configure_pivot(
            pivots.ConfigurePivotInput(
                sheet="Report",
                pivot="Summary",
                data_fields=[
                    pivots.PivotDataField(field="Amount", caption="Region")
                ],
            )
        )
    assert pivot.fields["Region"].Orientation == 1
    assert pivot.ManualUpdate is False
    pivot.AddDataField.assert_not_called()


def test_all_native_field_lookups_finish_before_hiding_existing_fields(
    workbook: Mock,
) -> None:
    pivot = native_pivot(workbook)
    pivot.fields["Region"].Orientation = 1
    original_lookup = pivot.PivotFields.side_effect

    def resolve(name: str | None = None) -> object:
        if name == "Amount":
            raise RuntimeError("native field lookup failure")
        return original_lookup(name)

    pivot.PivotFields.side_effect = resolve
    with pytest.raises(
        ValueError, match="Cannot resolve pivot cache field.*Amount.*refresh"
    ):
        pivots.configure_pivot(
            pivots.ConfigurePivotInput(
                sheet="Report",
                pivot="Summary",
                column_fields=["Quarter"],
                data_fields=[pivots.PivotDataField(field="Amount")],
            )
        )
    assert pivot.fields["Region"].Orientation == 1
    assert pivot.fields["Quarter"].Orientation == 0
    assert pivot.ManualUpdate is False
    pivot.AddDataField.assert_not_called()


def test_refresh_excludes_own_cells_but_not_neighbor_formulas(
    workbook: Mock,
) -> None:
    pivot = native_pivot(workbook)
    workbook.report.grid[(1, 8)] = "Region"
    result = pivots.refresh_pivot(
        pivots.PivotInput(sheet="Report", pivot="Summary")
    )
    assert result["name"] == "Summary"
    pivot.RefreshTable.assert_called_once()
    workbook.report.grid[(8, 12)] = '=IF(TRUE,"",1)'
    with pytest.raises(ValueError, match="existing content"):
        pivots.refresh_pivot(
            pivots.PivotInput(sheet="Report", pivot="Summary")
        )
    assert pivot.RefreshTable.call_count == 1


@pytest.mark.parametrize(
    "source",
    [
        "'[Other.xlsx]Data'!R1C1:R3C3",
        "'C:\\Data'!A1:C3",
        "NamedRange",
        "Data!A:A",
    ],
)
def test_refresh_rejects_external_and_unbounded_sources(
    workbook: Mock, source: str
) -> None:
    pivot = native_pivot(workbook)
    pivot.PivotCache.return_value.SourceData = source
    with pytest.raises(ValueError):
        pivots.refresh_pivot(
            pivots.PivotInput(sheet="Report", pivot="Summary")
        )
    pivot.RefreshTable.assert_not_called()


def test_refresh_rejects_non_range_cache_and_failed_native_result(
    workbook: Mock,
) -> None:
    pivot = native_pivot(workbook)
    pivot.PivotCache.return_value.SourceType = 2
    with pytest.raises(ValueError, match="local"):
        pivots.refresh_pivot(
            pivots.PivotInput(sheet="Report", pivot="Summary")
        )
    pivot.RefreshTable.assert_not_called()
    pivot.PivotCache.return_value.SourceType = 1
    pivot.RefreshTable.return_value = False
    with pytest.raises(RuntimeError, match="did not refresh"):
        pivots.refresh_pivot(
            pivots.PivotInput(sheet="Report", pivot="Summary")
        )
    pivot.RefreshTable.assert_called_once()


def test_shared_cache_refresh_protects_other_report_output(
    workbook: Mock,
) -> None:
    pivot = native_pivot(workbook)
    sibling = Mock()
    sibling.Name = "Shared"
    sibling.CacheIndex = pivot.CacheIndex
    sibling.TableRange2 = native_range(workbook.data, 1, 8, 4, 9)
    workbook.data.PivotTables.return_value = native_collection([sibling])
    workbook.data.grid[(10, 15)] = "Keep this"
    with pytest.raises(ValueError, match="existing content"):
        pivots.refresh_pivot(
            pivots.PivotInput(sheet="Report", pivot="Summary")
        )
    pivot.RefreshTable.assert_not_called()


def test_existing_pivot_source_inside_own_report_is_never_exempted(
    workbook: Mock,
) -> None:
    pivot = native_pivot(workbook, sheet=workbook.data)
    pivot.TableRange2 = native_range(workbook.data, 1, 1, 4, 3)
    with pytest.raises(ValueError, match="Source overlaps"):
        pivots.refresh_pivot(pivots.PivotInput(sheet="Data", pivot="Summary"))
    pivot.RefreshTable.assert_not_called()


@pytest.mark.parametrize(
    "source",
    [
        "'[Book.xlsx]Data'!$A$1:$C$3",
        "=Data!R1C1:R3C3",
        "='Data'!A1:C3",
        "='[Book.xlsx]Data'!$A$1:$C$3",
    ],
)
def test_local_source_accepts_native_reference_syntax(
    workbook: Mock, source: str
) -> None:
    pivot = native_pivot(workbook)
    pivot.PivotCache.return_value.SourceData = source
    target = pivots.resolve_pivot_source(workbook, pivot)
    assert table_access.range_bounds(target) == (1, 1, 3, 3)


@pytest.mark.parametrize(
    "source",
    [
        "==Data!R1C1:R3C3",
        "=SUM(Data!A1:C3)",
        "=Data!A1:C3+1",
        "=OFFSET(Data!A1,0,0,3,3)",
        "='[Other.xlsx]Data'!A1:C3",
    ],
)
def test_optional_source_equals_does_not_enable_arbitrary_formulas(
    workbook: Mock,
    source: str,
) -> None:
    pivot = native_pivot(workbook)
    pivot.PivotCache.return_value.SourceData = source
    with pytest.raises(ValueError):
        pivots.refresh_pivot(
            pivots.PivotInput(sheet="Report", pivot="Summary")
        )
    pivot.RefreshTable.assert_not_called()


def test_pivot_read_is_bounded_and_detaches_field_metadata(
    workbook: Mock,
) -> None:
    pivot = native_pivot(workbook)
    pivot.DataFields = native_collection(
        [
            SimpleNamespace(SourceName="Amount", Name="Total", Function=-4157),
        ]
    )
    with pytest.raises(ValueError, match="cell limit"):
        pivots.read_pivot(
            pivots.ReadPivotInput(sheet="Report", pivot="Summary", max_cells=1)
        )
    result = pivots.read_pivot(
        pivots.ReadPivotInput(sheet="Report", pivot="Summary")
    )
    assert result["data_fields"] == [
        {"field": "Amount", "caption": "Total", "aggregate": "sum"}
    ]
    assert len(result["values"]) == 4
    assert (
        pivots.list_pivots(models.SheetInput(sheet="Report"))["pivots"][0][
            "name"
        ]
        == "Summary"
    )


def test_delete_pivot_clears_exact_report_not_source(workbook: Mock) -> None:
    pivot = native_pivot(workbook)
    result = pivots.delete_pivot(
        pivots.PivotInput(sheet="Report", pivot="Summary")
    )
    pivot.TableRange2.Clear.assert_called_once()
    assert result["deleted"] is True
    assert workbook.data.grid[(2, 3)] == 10


def test_create_table_uses_existing_headers_and_rejects_rewritten_headers(
    workbook: Mock,
) -> None:
    table = native_table(workbook)
    workbook.data.ListObjects = native_collection([])
    workbook.data.ListObjects.Add.return_value = table
    result = tables.create_table(
        tables.CreateTableInput(name="Sales", address="A1:C3")
    )
    arguments = workbook.data.ListObjects.Add.call_args.kwargs
    assert arguments["SourceType"] == 1
    assert arguments["XlListObjectHasHeaders"] == 1
    assert result["headers"] == ["Region", "Quarter", "Amount"]
    table.ListColumns.Item.side_effect = lambda index: SimpleNamespace(
        Name=f"Changed{index}"
    )
    with pytest.raises(RuntimeError, match="changed table headers"):
        tables.create_table(
            tables.CreateTableInput(name="Sales", address="A1:C3")
        )


def test_native_extra_header_is_cut_back_and_only_empty_row_deleted(
    workbook: Mock,
) -> None:
    workbook.data.grid[(4, 1)] = "Neighbor"
    workbook.data.grid[(4, 2)] = "Keep"
    workbook.data.grid[(4, 3)] = 50
    original = dict(workbook.data.grid)
    table = native_shifted_table(workbook)
    result = tables.create_table(
        tables.CreateTableInput(
            name="Sales",
            address="A1:C3",
            style="TableStyleMedium2",
        )
    )
    assert result["address"] == "$A$1:$C$3"
    assert result["headers"] == ["Region", "Quarter", "Amount"]
    assert all(
        workbook.data.grid[position] == value
        for position, value in original.items()
    )
    table.DataBodyRange.Rows.Item.return_value.Cut.assert_called_once_with(
        Destination=table.HeaderRowRange,
    )
    table.ListRows.Item(1).Delete.assert_called_once_with()
    assert table.TableStyle == "TableStyleMedium2"


def test_unrecognized_inserted_header_never_cuts_or_deletes(
    workbook: Mock,
) -> None:
    table = native_shifted_table(workbook)
    first_row = table.DataBodyRange.Rows.Item.return_value
    first_row.Value2 = (("Different", "Quarter", "Amount"),)
    with pytest.raises(RuntimeError, match="did not preserve"):
        tables.create_table(
            tables.CreateTableInput(name="Sales", address="A1:C3")
        )
    first_row.Cut.assert_not_called()
    table.ListRows.Item(1).Delete.assert_not_called()


def test_cut_failure_does_not_delete_any_table_row(workbook: Mock) -> None:
    table = native_shifted_table(workbook)
    first_row = table.DataBodyRange.Rows.Item.return_value
    first_row.Cut.side_effect = RuntimeError("native cut failure")
    with pytest.raises(RuntimeError, match="native cut failure"):
        tables.create_table(
            tables.CreateTableInput(name="Sales", address="A1:C3")
        )
    first_row.Cut.assert_called_once()
    table.ListRows.Item(1).Delete.assert_not_called()


@pytest.mark.parametrize("remaining", ["Region", 0, False, '=IF(TRUE,"",1)'])
def test_nonempty_moved_header_row_is_never_deleted(
    workbook: Mock, remaining: object
) -> None:
    table = native_shifted_table(workbook)
    first_row = table.DataBodyRange.Rows.Item.return_value
    first_row.Cut.side_effect = lambda **arguments: setattr(
        first_row,
        "Formula",
        ((remaining, None, None),),
    )
    with pytest.raises(RuntimeError, match="refusing to delete"):
        tables.create_table(
            tables.CreateTableInput(name="Sales", address="A1:C3")
        )
    table.ListRows.Item(1).Delete.assert_not_called()


def test_header_repair_must_restore_exact_bounds_before_style(
    workbook: Mock,
) -> None:
    table = native_shifted_table(workbook)
    table.ListRows.Item(1).Delete.side_effect = None
    style = PropertyMock()
    type(table).TableStyle = style
    with pytest.raises(RuntimeError, match="did not restore"):
        tables.create_table(
            tables.CreateTableInput(
                name="Sales",
                address="A1:C3",
                style="TableStyleMedium2",
            )
        )
    table.ListRows.Item(1).Delete.assert_called_once()
    style.assert_not_called()


def test_unexpected_native_table_bounds_never_trigger_header_repair(
    workbook: Mock,
) -> None:
    table = native_table(workbook)
    table.Range = native_range(workbook.data, 1, 1, 5, 3)
    workbook.data.ListObjects = native_collection([])
    workbook.data.ListObjects.Add.return_value = table
    with pytest.raises(
        RuntimeError, match="changed the table range unexpectedly"
    ):
        tables.create_table(
            tables.CreateTableInput(name="Sales", address="A1:C3")
        )
    table.DataBodyRange.Rows.Item.assert_not_called()
    table.ListRows.Item.assert_not_called()


def test_table_creation_checks_overlap_before_native_write(
    workbook: Mock,
) -> None:
    table = native_table(workbook)
    with pytest.raises(ValueError, match="overlaps structured"):
        tables.create_table(
            tables.CreateTableInput(name="Other", address="A1:C3")
        )
    workbook.data.ListObjects.Add.assert_not_called()
    table.Delete.assert_not_called()


def test_table_creation_rejects_pivot_overlap(workbook: Mock) -> None:
    pivot = native_pivot(workbook, sheet=workbook.data)
    pivot.TableRange2 = native_range(workbook.data, 2, 1, 3, 3)
    with pytest.raises(ValueError, match="overlaps pivot"):
        tables.create_table(
            tables.CreateTableInput(name="Sales", address="A1:C3")
        )
    workbook.data.ListObjects.Add.assert_not_called()


def test_table_read_respects_bounds_and_keeps_formulas(workbook: Mock) -> None:
    native_table(workbook)
    with pytest.raises(ValueError, match="cell limit"):
        tables.read_table(tables.ReadTableInput(table="Sales", max_cells=2))
    result = tables.read_table(
        tables.ReadTableInput(table="sales", include_formulas=True)
    )
    assert result["values"][1] == ["North", "Q1", 10]
    assert result["formulas"][2] == ["South", "Q1", 20]
    assert (
        tables.list_tables(models.SheetInput())["tables"][0]["name"] == "Sales"
    )


@pytest.mark.parametrize("address", ["B1:C3", "A2:C4", "A1:C1"])
def test_invalid_resize_does_not_touch_table(
    workbook: Mock, address: str
) -> None:
    table = native_table(workbook)
    with pytest.raises(ValueError):
        tables.resize_table(
            tables.ResizeTableInput(table="Sales", address=address)
        )
    table.Resize.assert_not_called()


def test_resize_preserves_headers_and_released_cells(workbook: Mock) -> None:
    table = native_table(workbook)
    table.Resize.side_effect = lambda target: setattr(table, "Range", target)
    result = tables.resize_table(
        tables.ResizeTableInput(table="Sales", address="A1:C2")
    )
    assert result["rows"] == 2
    assert result["excluded_cells_preserved"] is True
    assert workbook.data.grid[(3, 3)] == 20
    assert result["previous_address"] == "$A$1:$C$3"


def test_resize_rejects_header_rewrite_or_totals(workbook: Mock) -> None:
    table = native_table(workbook)
    workbook.data.grid[(1, 2)] = "Renamed"
    with pytest.raises(ValueError, match="rename"):
        tables.resize_table(
            tables.ResizeTableInput(table="Sales", address="A1:C4")
        )
    table.Resize.assert_not_called()
    table.ShowTotals = True
    with pytest.raises(ValueError, match="totals"):
        tables.resize_table(
            tables.ResizeTableInput(table="Sales", address="A1:C4")
        )
    table.Resize.assert_not_called()


def test_configure_validates_style_before_renaming_and_preserves_false(
    workbook: Mock,
) -> None:
    table = native_table(workbook)
    workbook.TableStyles.Item.side_effect = ValueError("unknown style")
    with pytest.raises(ValueError, match="unknown style"):
        tables.configure_table(
            tables.ConfigureTableInput(
                table="Sales",
                name="Renamed",
                style="Missing",
            )
        )
    assert table.Name == "Sales"
    workbook.TableStyles.Item.side_effect = None
    tables.configure_table(
        tables.ConfigureTableInput(
            table="Sales",
            name="Renamed",
            banded_rows=False,
            first_column=True,
        )
    )
    assert table.Name == "Renamed"
    assert table.ShowTableStyleRowStripes is False
    assert table.ShowTableStyleFirstColumn is True


def test_unlist_and_delete_have_distinct_data_contracts(
    workbook: Mock,
) -> None:
    table = native_table(workbook)
    result = tables.unlist_table(tables.TableInput(table="Sales"))
    assert result["data_preserved"] is True
    table.Unlist.assert_called_once()
    table.Delete.assert_not_called()
    with pytest.raises(ValidationError):
        tables.DeleteTableInput.model_validate({"table": "Sales"})
    with pytest.raises(ValidationError):
        tables.DeleteTableInput(table="Sales", delete_data=False)
    with pytest.raises(ValidationError):
        tables.DeleteTableInput.model_validate(
            {"table": "Sales", "delete_data": 1}
        )
    result = tables.delete_table(
        tables.DeleteTableInput(table="Sales", delete_data=True)
    )
    table.Delete.assert_called_once()
    assert result["data_preserved"] is False


def test_missing_objects_and_empty_configuration_fail_explicitly(
    workbook: Mock,
) -> None:
    with pytest.raises(ValueError, match="not found"):
        tables.read_table(tables.ReadTableInput(table="Missing"))
    with pytest.raises(ValueError, match="not found"):
        pivots.read_pivot(pivots.ReadPivotInput(pivot="Missing"))
    with pytest.raises(ValidationError, match="at least one"):
        tables.ConfigureTableInput(table="Sales")
