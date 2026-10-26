"""Create, edit, and format native Writer tables."""

from typing import Literal, Self

from pydantic import Field, model_validator

from wps_mcp.common.colors import RgbColor, hex_to_int
from wps_mcp.models import CellValue, DocumentInput, JsonObject
from wps_mcp.native.runtime import get_session
from wps_mcp.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)
from wps_mcp.writer.access import resolve_item, resolve_range
from wps_mcp.writer.formatting import ALIGNMENT_VALUES, apply_character
from wps_mcp.writer.models import CharacterFormat, Index, Position


class TableInput(DocumentInput):
    """Identify an existing native table in the selected document."""

    table_index: Index


class TableDataInput(DocumentInput):
    """Require rectangular literal table data with valid cell text."""

    data: list[list[CellValue]] = Field(min_length=1)

    @model_validator(mode="after")
    def validate_data(self) -> Self:
        """Require rectangular rows and reject native cell-end markers."""
        column_count = len(self.data[0])
        if column_count == 0 or column_count > 63:
            raise ValueError("Table data must have between 1 and 63 columns")
        if any(len(row) != column_count for row in self.data):
            raise ValueError("Table data must be rectangular")
        if any(
            isinstance(value, str) and "\x07" in value
            for row in self.data
            for value in row
        ):
            raise ValueError(
                "Table text cannot include the native end-of-cell marker"
            )
        return self


class AddTableInput(TableDataInput):
    """Create a Writer table at an explicit UTF-16 position."""

    position: Position
    style_name: str | None = Field(default=None, min_length=1)


class WriteTableInput(TableDataInput):
    """Write literal data from one-based table coordinates."""

    table_index: Index
    start_row: Index = 1
    start_column: Index = 1


class ChangeTableInput(TableInput):
    """Insert or delete an explicitly selected table row or column."""

    action: Literal[
        "insert_row", "delete_row", "insert_column", "delete_column"
    ]
    position: Index | None = Field(
        default=None,
        description=(
            "Insert before this index; omit to append. Required for delete."
        ),
    )

    @model_validator(mode="after")
    def validate_action(self) -> Self:
        """Require a position before deleting a table row or column."""
        if self.action.startswith("delete") and self.position is None:
            raise ValueError("position is required for deletion")
        return self


class FormatTableCellInput(TableInput):
    """Format one native table cell with optional style changes."""

    row: Index
    column: Index
    character: CharacterFormat | None = None
    alignment: (
        Literal["left", "center", "right", "justify", "distribute"] | None
    ) = None
    vertical_alignment: Literal["top", "center", "bottom"] | None = None
    fill_color: RgbColor | None = None

    @model_validator(mode="after")
    def validate_format(self) -> Self:
        """Require at least one explicit formatting property."""
        character_values = (
            self.character.model_dump(exclude_none=True)
            if self.character
            else {}
        )
        if (
            not character_values
            and self.alignment is None
            and (self.vertical_alignment is None)
            and (self.fill_color is None)
        ):
            raise ValueError("Specify at least one cell formatting property")
        return self


def cell_text(value: CellValue) -> str:
    """Encode literal table text while preserving zero and false values."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    return str(value)


@tool(
    "wps_writer_tables",
    operation="list",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def read_tables(params: DocumentInput) -> JsonObject:
    """List native table bounds including tables with merged cells."""
    document = get_session("writer").get_document(params.document)
    result = []
    for index in range(1, int(document.Tables.Count) + 1):
        table = document.Tables.Item(index)
        result.append(
            {
                "index": index,
                "start": int(table.Range.Start),
                "end": int(table.Range.End),
                "cell_count": int(table.Range.Cells.Count),
            }
        )
    return {"tables": result}


@tool(
    "wps_writer_tables",
    operation="read",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def read_table(params: TableInput) -> JsonObject:
    """Read every table cell with its row, column and text.

    Preserves meaningful trailing whitespace and handles merged cells.
    """
    document = get_session("writer").get_document(params.document)
    table = resolve_item(document.Tables, params.table_index, "table")
    cells = table.Range.Cells
    result = []
    for index in range(1, int(cells.Count) + 1):
        cell = cells.Item(index)
        text = str(cell.Range.Text)
        if text.endswith("\r\x07"):
            text = text[:-2]
        result.append(
            {
                "row": int(cell.RowIndex),
                "column": int(cell.ColumnIndex),
                "text": text,
                "start": int(cell.Range.Start),
                "end": int(cell.Range.End),
            }
        )
    return {"table_index": params.table_index, "cells": result}


@tool(
    "wps_writer_tables",
    operation="add",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def add_table(params: AddTableInput) -> JsonObject:
    """Insert a rectangular table at a native position.

    Zero/false are retained; null creates an empty cell. Indices start at
    one.
    """
    document = get_session("writer").get_document(params.document)
    text_range = resolve_range(document, params.position, params.position)
    style = (
        document.Styles.Item(params.style_name)
        if params.style_name is not None
        else None
    )
    data = [[cell_text(value) for value in row] for row in params.data]
    table = document.Tables.Add(text_range, len(data), len(data[0]))
    if style is not None:
        table.Style = style
    for row_index, row in enumerate(data, start=1):
        for column_index, value in enumerate(row, start=1):
            table.Cell(row_index, column_index).Range.Text = value
    return {
        "start": int(table.Range.Start),
        "end": int(table.Range.End),
        "rows": len(data),
        "columns": len(data[0]),
    }


@tool(
    "wps_writer_tables",
    operation="write",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def write_table(params: WriteTableInput) -> JsonObject:
    """Preflight every target cell before writing literal table data.

    The target rectangle must fit an existing unmerged table.
    """
    document = get_session("writer").get_document(params.document)
    table = resolve_item(document.Tables, params.table_index, "table")
    row_count = int(table.Rows.Count)
    column_count = int(table.Columns.Count)
    if (
        params.start_row + len(params.data) - 1 > row_count
        or params.start_column + len(params.data[0]) - 1 > column_count
    ):
        raise ValueError(f"Data exceeds table size {row_count}x{column_count}")
    changes = []
    for row_offset, row in enumerate(params.data):
        for column_offset, value in enumerate(row):
            cell = table.Cell(
                params.start_row + row_offset,
                params.start_column + column_offset,
            )
            if (
                int(cell.RowIndex) != params.start_row + row_offset
                or int(cell.ColumnIndex) != params.start_column + column_offset
            ):
                raise ValueError("The target rectangle includes a merged cell")
            changes.append((cell.Range, cell_text(value)))
    for text_range, text in changes:
        text_range.Text = text
    return {"table_index": params.table_index, "cells_written": len(changes)}


@tool(
    "wps_writer_tables",
    operation="change",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def change_table(params: ChangeTableInput) -> JsonObject:
    """Insert or delete a table row/column.

    Insert before position or omit to append; deletion requires a position.
    """
    document = get_session("writer").get_document(params.document)
    table = resolve_item(document.Tables, params.table_index, "table")
    collection = (
        table.Rows if params.action.endswith("_row") else table.Columns
    )
    count = int(collection.Count)
    if params.action.startswith("insert"):
        position = count + 1 if params.position is None else params.position
        if position > count + 1:
            raise ValueError(
                f"Insertion position must be within 1..{count + 1}"
            )
        if params.action == "insert_column" and count >= 63:
            raise ValueError("A WPS Writer table supports at most 63 columns")
        if position == count + 1:
            collection.Add()
        else:
            collection.Add(collection.Item(position))
    else:
        if params.position is None:
            raise ValueError("position is required for deletion")
        position = params.position
        item = resolve_item(
            collection,
            position,
            "row" if params.action.endswith("_row") else "column",
        )
        if count == 1:
            raise ValueError(
                "Cannot delete the final row or column; use "
                "wps_writer_tables/delete"
            )
        item.Delete()
    return {
        "table_index": params.table_index,
        "action": params.action,
        "position": position,
    }


@tool(
    "wps_writer_tables",
    operation="delete",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def delete_table(params: TableInput) -> JsonObject:
    """Delete the selected native table and its cell contents."""
    document = get_session("writer").get_document(params.document)
    table = resolve_item(document.Tables, params.table_index, "table")
    table.Delete()
    return {"table_index": params.table_index, "deleted": True}


@tool(
    "wps_writer_tables",
    operation="format_cell",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def format_table_cell(params: FormatTableCellInput) -> JsonObject:
    """Format one validated table cell and its text.

    Supports font, paragraph alignment, vertical alignment, and fill color.
    """
    document = get_session("writer").get_document(params.document)
    table = resolve_item(document.Tables, params.table_index, "table")
    if params.row > int(table.Rows.Count) or params.column > int(
        table.Columns.Count
    ):
        raise ValueError("Cell coordinates exceed the table bounds")
    cell = table.Cell(params.row, params.column)
    if (
        int(cell.RowIndex) != params.row
        or int(cell.ColumnIndex) != params.column
    ):
        raise ValueError("The coordinates refer to part of a merged cell")
    apply_character(cell.Range, params.character)
    if params.alignment is not None:
        cell.Range.ParagraphFormat.Alignment = ALIGNMENT_VALUES[
            params.alignment
        ]
    if params.vertical_alignment is not None:
        cell.VerticalAlignment = {"top": 0, "center": 1, "bottom": 3}[
            params.vertical_alignment
        ]
    if params.fill_color is not None:
        cell.Shading.BackgroundPatternColor = hex_to_int(params.fill_color)
    return {
        "table_index": params.table_index,
        "row": params.row,
        "column": params.column,
    }
