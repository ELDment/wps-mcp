"""Manage structured tables while preserving native headers."""

from typing import Self

from pydantic import Field, field_validator, model_validator

from src.models import JsonObject, NativeObject
from src.native.runtime import get_session
from src.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)
from src.spreadsheet.access import get_range, get_sheet, range_result
from src.spreadsheet.matrices import matrix_result, range_values
from src.spreadsheet.models import MAX_CELLS, RangeInput, SheetInput
from src.spreadsheet.object_names import ObjectName
from src.spreadsheet.table_access import (
    ensure_available_name,
    ensure_table_space,
    get_object,
    range_bounds,
    read_headers,
)


class TableInput(SheetInput):
    """Identify an existing native table in the selected document."""

    table: str = Field(min_length=1)


class CreateTableInput(RangeInput):
    """Create a structured table with validated name and style."""

    name: ObjectName
    style: str | None = Field(default=None, min_length=1)


class ReadTableInput(TableInput):
    """Read bounded structured-table data with optional formulas."""

    max_cells: int = Field(default=10000, ge=1, le=MAX_CELLS)
    include_formulas: bool = False


class ResizeTableInput(TableInput, RangeInput):
    """Resize an existing table to an explicit local rectangle."""

    pass


class ConfigureTableInput(TableInput):
    """Require explicit table naming or presentation changes."""

    name: ObjectName | None = None
    style: str | None = Field(default=None, min_length=1)
    banded_rows: bool | None = None
    banded_columns: bool | None = None
    first_column: bool | None = None
    last_column: bool | None = None

    @model_validator(mode="after")
    def validate_changes(self) -> Self:
        """Reject updates that do not change any supported property."""
        if not any(
            value is not None
            for value in (
                self.name,
                self.style,
                self.banded_rows,
                self.banded_columns,
                self.first_column,
                self.last_column,
            )
        ):
            raise ValueError("Supply at least one table property to configure")
        return self


class DeleteTableInput(TableInput):
    """Acknowledge that removing a table also deletes its cell data."""

    delete_data: bool = Field(
        description=(
            "Required acknowledgement that deleting a table "
            "also clears its cells"
        )
    )

    @field_validator("delete_data")
    @classmethod
    def validate_delete_data(cls, value: bool) -> bool:
        """Require explicit acknowledgement before removing table data."""
        if not value:
            raise ValueError(
                "Deleting table cells requires delete_data=true; "
                "use unlist to keep data"
            )
        return value


def describe_table(sheet: NativeObject, table: NativeObject) -> JsonObject:
    """Detach table headers, bounds, and header or totals visibility."""
    columns = table.ListColumns
    return {
        "name": str(table.Name),
        **range_result(sheet, table.Range),
        "headers": [
            str(columns.Item(index).Name)
            for index in range(1, int(columns.Count) + 1)
        ],
        "show_totals": bool(table.ShowTotals),
        "show_headers": bool(table.ShowHeaders),
    }


@tool(
    "wps_et_tables",
    operation="list",
    component="et",
    annotations=READ_ANNOTATIONS,
)
def list_tables(params: SheetInput) -> JsonObject:
    """List native structured tables and their exact column headers."""
    sheet = get_sheet(
        get_session("et").get_document(params.document), params.sheet
    )
    return {
        "sheet": str(sheet.Name),
        "tables": [
            describe_table(sheet, sheet.ListObjects.Item(index))
            for index in range(1, int(sheet.ListObjects.Count) + 1)
        ],
    }


@tool(
    "wps_et_tables",
    operation="create",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def create_table(params: CreateTableInput) -> JsonObject:
    """Create a structured table while preserving source cell references.

    Literal headers and neighboring cells are preserved.

    WPS builds that insert an extra header row are handled only after proving
    the original headers moved intact. A native cut restores their references;
    only the confirmed empty inserted data row is then deleted.
    """
    workbook = get_session("et").get_document(params.document)
    sheet = get_sheet(workbook, params.sheet)
    source = get_range(sheet, params.address)
    headers = read_headers(source)
    original_bounds = range_bounds(source)
    ensure_available_name(workbook, params.name)
    ensure_table_space(sheet, source)
    if params.style is not None:
        workbook.TableStyles.Item(params.style)
    table = sheet.ListObjects.Add(
        SourceType=1, Source=source, XlListObjectHasHeaders=1
    )
    if range_bounds(table.Range) != original_bounds:
        top, left, bottom, right = original_bounds
        if (
            range_bounds(table.Range) != (top, left, bottom + 1, right)
            or not bool(table.ShowHeaders)
            or bool(table.ShowTotals)
            or int(table.ListRows.Count) != bottom - top + 1
        ):
            raise RuntimeError(
                "WPS changed the table range unexpectedly; "
                "inspect before retrying"
            )
        first_row = table.DataBodyRange.Rows.Item(1)
        header_row = table.HeaderRowRange
        if (
            range_bounds(first_row) != (top + 1, left, top + 1, right)
            or range_bounds(header_row) != (top, left, top, right)
            or matrix_result(first_row.Value2, 1, len(headers))[0] != headers
        ):
            raise RuntimeError(
                "WPS did not preserve the original headers; "
                "inspect before retrying"
            )
        first_row.Cut(Destination=header_row)
        empty_row = table.ListRows.Item(1)
        if range_bounds(empty_row.Range) != (
            top + 1,
            left,
            top + 1,
            right,
        ) or any(
            value is not None and value != ""
            for value in matrix_result(
                empty_row.Range.Formula, 1, len(headers)
            )[0]
        ):
            raise RuntimeError(
                "Moved header row is not empty; refusing to delete table data"
            )
        empty_row.Delete()
    if range_bounds(table.Range) != original_bounds:
        raise RuntimeError(
            "WPS did not restore the original table range; "
            "inspect before retrying"
        )
    result = describe_table(sheet, table)
    if result["headers"] != headers:
        raise RuntimeError(
            "WPS changed table headers unexpectedly; inspect "
            "the table before retrying"
        )
    table.Name = params.name
    if params.style is not None:
        table.TableStyle = params.style
    return {**result, "name": str(table.Name)}


@tool(
    "wps_et_tables",
    operation="read",
    component="et",
    annotations=READ_ANNOTATIONS,
)
def read_table(params: ReadTableInput) -> JsonObject:
    """Read native table data while preserving cell text and boundaries.

    Read the whole table, including headers/totals, with a strict cell
    limit.
    """
    sheet = get_sheet(
        get_session("et").get_document(params.document), params.sheet
    )
    table = get_object(sheet, params.table)
    target = get_range(sheet, str(table.Range.Address), params.max_cells)
    result = describe_table(sheet, table)
    result.update(
        range_values(target, include_formulas=params.include_formulas)
    )
    return result


@tool(
    "wps_et_tables",
    operation="resize",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def resize_table(params: ResizeTableInput) -> JsonObject:
    """Resize a table while retaining its header origin and shared names.

    Cells excluded by shrinking remain ordinary worksheet cells. Tables with
    hidden headers or totals must first be adjusted in WPS to avoid implicit
    header insertion or totals-row movement.
    """
    sheet = get_sheet(
        get_session("et").get_document(params.document), params.sheet
    )
    table = get_object(sheet, params.table)
    if not bool(table.ShowHeaders) or bool(table.ShowTotals):
        raise ValueError("Resize requires visible headers and no totals row")
    target = get_range(sheet, params.address)
    if range_bounds(target)[:2] != range_bounds(table.Range)[:2]:
        raise ValueError(
            "Resize must keep the existing table's upper-left header cell"
        )
    headers = read_headers(target)
    previous = describe_table(sheet, table)
    shared = min(len(headers), len(previous["headers"]))
    if headers[:shared] != previous["headers"][:shared]:
        raise ValueError(
            "Resize cannot silently rename existing table columns"
        )
    ensure_table_space(sheet, target, current=str(table.Name))
    table.Resize(target)
    result = describe_table(sheet, table)
    if result["headers"] != headers:
        raise RuntimeError(
            "WPS changed table headers unexpectedly; inspect before retrying"
        )
    return {
        **result,
        "previous_address": previous["address"],
        "excluded_cells_preserved": True,
    }


@tool(
    "wps_et_tables",
    operation="configure",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def configure_table(params: ConfigureTableInput) -> JsonObject:
    """Set validated table names, styles, and column or row emphasis."""
    workbook = get_session("et").get_document(params.document)
    sheet = get_sheet(workbook, params.sheet)
    table = get_object(sheet, params.table)
    if params.name is not None:
        ensure_available_name(workbook, params.name, current=str(table.Name))
    if params.style is not None:
        workbook.TableStyles.Item(params.style)
    for attribute, value in (
        ("Name", params.name),
        ("TableStyle", params.style),
        ("ShowTableStyleRowStripes", params.banded_rows),
        ("ShowTableStyleColumnStripes", params.banded_columns),
        ("ShowTableStyleFirstColumn", params.first_column),
        ("ShowTableStyleLastColumn", params.last_column),
    ):
        if value is not None:
            setattr(table, attribute, value)
    return describe_table(sheet, table)


@tool(
    "wps_et_tables",
    operation="unlist",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def unlist_table(params: TableInput) -> JsonObject:
    """Convert a table to ordinary cells while preserving its data.

    Native cell formatting is preserved with the data.
    """
    sheet = get_sheet(
        get_session("et").get_document(params.document), params.sheet
    )
    table = get_object(sheet, params.table)
    result = describe_table(sheet, table)
    table.Unlist()
    return {**result, "unlisted": True, "data_preserved": True}


@tool(
    "wps_et_tables",
    operation="delete",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def delete_table(params: DeleteTableInput) -> JsonObject:
    """Delete the selected native table and its cell contents.

    Use unlist to preserve ordinary cell data and formatting.
    """
    sheet = get_sheet(
        get_session("et").get_document(params.document), params.sheet
    )
    table = get_object(sheet, params.table)
    result = describe_table(sheet, table)
    table.Delete()
    return {**result, "deleted": True, "data_preserved": False}
