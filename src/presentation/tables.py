"""Table operations for native WPS Presentation automation.

Handles creating tables, setting cell text/formatting, merging cells,
adding/deleting rows/columns, and applying table styles.
"""

from operator import attrgetter
from typing import Annotated, Literal

from pydantic import Field, field_validator, model_validator

from src.common.colors import hex_to_int, int_to_hex
from src.common.text import font_size_warning, normalize_text
from src.errors import native_errors
from src.models import DocumentInput, JsonObject, NativeObject
from src.native.runtime import get_session
from src.presentation.constants import (
    DASH_STYLE_DESCRIPTION,
    MSO_FALSE,
    MSO_TRUE,
    PP_BORDER_BOTTOM,
    PP_BORDER_DIAGONAL_DOWN,
    PP_BORDER_DIAGONAL_UP,
    PP_BORDER_LEFT,
    PP_BORDER_RIGHT,
    PP_BORDER_TOP,
    check_dash_style,
    dash_style_value,
)
from src.presentation.navigation import goto_slide
from src.presentation.shape_lookup import resolve_shape
from src.presentation.text_enums import (
    ALIGNMENT_MAP,
    VERTICAL_ANCHOR_MAP,
    VERTICAL_ANCHOR_NAMES,
)
from src.presentation.text_formatting import apply_font_format
from src.presentation.text_models import TextFormatSpec
from src.protocol.definitions import tool

BORDER_SIDE_MAP: dict[str, int] = {
    "top": PP_BORDER_TOP,
    "left": PP_BORDER_LEFT,
    "bottom": PP_BORDER_BOTTOM,
    "right": PP_BORDER_RIGHT,
    "diagonal_down": PP_BORDER_DIAGONAL_DOWN,
    "diagonal_up": PP_BORDER_DIAGONAL_UP,
}
_DIAGONAL_SIDES: set[int] = {PP_BORDER_DIAGONAL_DOWN, PP_BORDER_DIAGONAL_UP}
ALIGNMENT_NAMES: dict[int, str] = {v: k for k, v in ALIGNMENT_MAP.items()}


class TableInput(DocumentInput):
    """Locate a table on a one-based slide index."""

    slide_index: int = Field(ge=1)
    shape_name_or_index: str | int


class TableLayoutInput(DocumentInput):
    """Optional dimensions in points, starting at the first row or column."""

    row_heights: list[Annotated[float, Field(gt=0)]] | None = None
    col_widths: list[Annotated[float, Field(gt=0)]] | None = None


class AddTableInput(TableLayoutInput):
    """Create a table with dimensions in points."""

    slide_index: int = Field(ge=1)
    rows: int = Field(ge=1)
    cols: int = Field(ge=1)
    left: float = 50.0
    top: float = 100.0
    width: float = Field(default=600.0, gt=0)
    height: float = Field(default=300.0, gt=0)

    @model_validator(mode="after")
    def check_layout_size(self) -> "AddTableInput":
        """Reject row or column dimensions beyond the new table size."""
        if len(self.row_heights or []) > self.rows:
            raise ValueError("row_heights exceeds the number of rows")
        if len(self.col_widths or []) > self.cols:
            raise ValueError("col_widths exceeds the number of columns")
        return self


class GetTableDataInput(TableInput):
    """Read cell text and optionally formatting."""

    include_format: bool = False


class SetTableCellInput(TableInput):
    """Set one cell's text, font, fill and alignment."""

    row: int = Field(ge=1)
    col: int = Field(ge=1)
    text: str | None = None
    font_name: str | None = None
    font_name_fareast: str | None = None
    font_size: float | None = Field(default=None, gt=0)
    bold: bool | None = None
    italic: bool | None = None
    color: str | None = None
    fill_color: str | None = None
    alignment: Literal["left", "center", "right", "justify"] | None = None
    vertical_alignment: Literal["top", "middle", "bottom"] | None = None


class SetTableDataInput(TableInput):
    """Write an entire rectangular batch after checking its bounds."""

    data: list[list[str]] = Field(min_length=1)
    start_row: int = Field(default=1, ge=1)
    start_col: int = Field(default=1, ge=1)
    bold_first_row: bool = False

    @model_validator(mode="after")
    def check_data_shape(self) -> "SetTableDataInput":
        """Require a nonempty rectangular matrix of cell values."""
        width = len(self.data[0])
        if not width or any(len(row) != width for row in self.data):
            raise ValueError("data must be a nonempty rectangular matrix")
        return self


class MergeTableCellsInput(TableInput):
    """Merge a rectangular range of existing cells."""

    start_row: int = Field(ge=1)
    start_col: int = Field(ge=1)
    end_row: int = Field(ge=1)
    end_col: int = Field(ge=1)

    @model_validator(mode="after")
    def check_range_order(self) -> "MergeTableCellsInput":
        """Require the cell range end to follow its start."""
        if self.end_row < self.start_row or self.end_col < self.start_col:
            raise ValueError("Range end must follow its start")
        return self


class TableRowInput(TableInput):
    """Insert before position, or append when position is omitted."""

    position: int | None = Field(default=None, ge=1)
    height: float | None = Field(default=None, gt=0)


class DeleteTableRowInput(TableInput):
    """Delete the row at a required one-based position."""

    position: int = Field(ge=1)


class TableColumnInput(TableInput):
    """Insert before position, or append when position is omitted."""

    position: int | None = Field(default=None, ge=1)
    width: float | None = Field(default=None, gt=0)


class DeleteTableColumnInput(TableInput):
    """Delete the column at a required one-based position."""

    position: int = Field(ge=1)


class SetTableStyleInput(TableInput):
    """Apply a built-in style GUID and optional table banding."""

    style_id: str | None = None
    first_row: bool | None = None
    last_row: bool | None = None
    first_col: bool | None = None
    last_col: bool | None = None
    banding_rows: bool | None = None
    banding_cols: bool | None = None


class SetTableLayoutInput(TableInput, TableLayoutInput):
    """Apply partial row and column dimensions; excess items are rejected."""


class SplitTableCellsInput(TableInput):
    """Split a merged cell into num_rows by num_cols cells."""

    row: int = Field(ge=1)
    col: int = Field(ge=1)
    num_rows: int = Field(default=1, ge=1)
    num_cols: int = Field(default=1, ge=1)


class SetTableBordersInput(TableInput):
    """Style selected borders of a rectangular cell range."""

    start_row: int = Field(default=1, ge=1)
    start_col: int = Field(default=1, ge=1)
    end_row: int | None = Field(default=None, ge=1)
    end_col: int | None = Field(default=None, ge=1)
    sides: list[
        Literal[
            "top", "bottom", "left", "right", "diagonal_down", "diagonal_up"
        ]
    ] = Field(min_length=1)
    visible: bool | None = None
    color: str | None = None
    weight: float | None = Field(default=None, gt=0)
    dash_style: str | None = Field(
        default=None, description=DASH_STYLE_DESCRIPTION
    )

    @field_validator("dash_style")
    @classmethod
    def check_style(cls, value: str | None) -> str | None:
        """Validate the requested native dash style."""
        return check_dash_style(value)

    @model_validator(mode="after")
    def check_range_and_properties(self) -> "SetTableBordersInput":
        """Validate range order and require a border property."""
        if (self.end_row is not None and self.end_row < self.start_row) or (
            self.end_col is not None and self.end_col < self.start_col
        ):
            raise ValueError("Range end must follow its start")
        if all(
            value is None
            for value in (
                self.visible,
                self.color,
                self.weight,
                self.dash_style,
            )
        ):
            raise ValueError("At least one border property must be provided")
        return self


def _get_table_shape(
    params: TableInput, *, navigate: bool = True
) -> NativeObject:
    """Resolve the selected document, slide and required table shape."""
    session = get_session("wpp")
    document = session.get_document(params.document)
    if navigate:
        goto_slide(session.get_application(), params.slide_index)
    shape = resolve_shape(
        document.Slides(params.slide_index), params.shape_name_or_index
    )
    if not shape.HasTable:
        raise ValueError(f"Shape '{shape.Name}' is not a table")
    return shape


def _validate_table_range(
    table: NativeObject,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> None:
    """Reject the full operation when any cell lies outside the table."""
    if not (
        1 <= start_row <= end_row <= table.Rows.Count
        and 1 <= start_col <= end_col <= table.Columns.Count
    ):
        raise ValueError(
            "Cell range exceeds table bounds; nothing was changed"
        )


def _get_cell_format(cell: NativeObject) -> JsonObject:
    """Extract formatting details from a table cell."""
    properties = {
        "fill_color_hex": "Fill.ForeColor.RGB",
        "font_name": "TextFrame.TextRange.Font.Name",
        "font_name_fareast": "TextFrame.TextRange.Font.NameFarEast",
        "font_size": "TextFrame.TextRange.Font.Size",
        "bold": "TextFrame.TextRange.Font.Bold",
        "italic": "TextFrame.TextRange.Font.Italic",
        "font_color_hex": "TextFrame.TextRange.Font.Color.RGB",
        "alignment": "TextFrame.TextRange.ParagraphFormat.Alignment",
        "vertical_alignment": "TextFrame.VerticalAnchor",
    }
    result: JsonObject = {}
    for name, path in properties.items():
        try:
            value = attrgetter(path)(cell.Shape)
            if name in {"fill_color_hex", "font_color_hex"}:
                value = int_to_hex(value)
            elif name in {"bold", "italic"}:
                value = value == MSO_TRUE
            elif name == "alignment":
                value = ALIGNMENT_NAMES.get(value)
            elif name == "vertical_alignment":
                value = VERTICAL_ANCHOR_NAMES.get(value)
            result[name] = value
        except native_errors() + (AttributeError,):
            # WPS versions differ in the format properties they expose.
            result[name] = None
    return result


@tool(
    "wps_ppt_tables",
    component="wpp",
    annotations={"title": "Add Table"},
    operation="add",
)
def add_table(params: AddTableInput) -> JsonObject:
    """Add a table to a slide.

    Creates a table with the specified number of rows and columns. All
    positions and sizes are in points (72 points = 1 inch).
    """
    app = get_session("wpp").get_application()
    goto_slide(app, params.slide_index)
    pres = get_session("wpp").get_document(params.document)
    slide = pres.Slides(params.slide_index)
    shape = slide.Shapes.AddTable(
        NumRows=params.rows,
        NumColumns=params.cols,
        Left=params.left,
        Top=params.top,
        Width=params.width,
        Height=params.height,
    )
    table = shape.Table
    if params.row_heights:
        for i, h in enumerate(params.row_heights, 1):
            if i <= table.Rows.Count:
                table.Rows(i).Height = h
    if params.col_widths:
        for i, w in enumerate(params.col_widths, 1):
            if i <= table.Columns.Count:
                table.Columns(i).Width = w
    return {
        "success": True,
        "shape_name": shape.Name,
        "shape_index": shape.ZOrderPosition,
        "rows": table.Rows.Count,
        "columns": table.Columns.Count,
    }


@tool(
    "wps_ppt_tables",
    component="wpp",
    annotations={
        "title": "Get Table Data",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="get_data",
)
def get_table_data(params: GetTableDataInput) -> JsonObject:
    """Get all cell text values from a table.

    Returns a 2D array of cell text, plus row and column counts. Identify
    the table by shape name or 1-based shape index.
    """
    shape = _get_table_shape(params, navigate=False)
    table = shape.Table
    rows_count = table.Rows.Count
    cols_count = table.Columns.Count
    data: list[list[str]] = []
    formats: list[list[JsonObject]] = []
    for r in range(1, rows_count + 1):
        row_data: list[str] = []
        row_formats: list[JsonObject] = []
        for c in range(1, cols_count + 1):
            cell = table.Cell(r, c)
            row_data.append(cell.Shape.TextFrame.TextRange.Text)
            if params.include_format:
                row_formats.append(_get_cell_format(cell))
        data.append(row_data)
        if params.include_format:
            formats.append(row_formats)
    result = {
        "success": True,
        "shape_name": shape.Name,
        "rows": rows_count,
        "columns": cols_count,
        "data": data,
    }
    if params.include_format:
        result["format"] = formats
    return result


@tool(
    "wps_ppt_tables",
    component="wpp",
    annotations={"title": "Set Table Cell", "idempotentHint": True},
    operation="set_cell",
)
def set_table_cell(params: SetTableCellInput) -> JsonObject:
    """Set text and/or formatting for a table cell.

    Access cell by 1-based row and column. Optionally set font properties,
    text alignment, and cell background color.
    """
    shape = _get_table_shape(params)
    table = shape.Table
    _validate_table_range(
        table, params.row, params.col, params.row, params.col
    )
    cell = table.Cell(params.row, params.col)
    font_options = TextFormatSpec(
        font_name=params.font_name,
        font_name_fareast=params.font_name_fareast,
        font_size=params.font_size,
        bold=params.bold,
        italic=params.italic,
        color=params.color,
    )
    fill_color = (
        hex_to_int(params.fill_color)
        if params.fill_color is not None
        else None
    )
    if params.text is not None:
        cell.Shape.TextFrame.TextRange.Text = normalize_text(params.text)
    tr = cell.Shape.TextFrame.TextRange
    apply_font_format(tr.Font, font_options)
    if params.alignment is not None:
        tr.ParagraphFormat.Alignment = ALIGNMENT_MAP[params.alignment]
    if params.fill_color is not None:
        cell.Shape.Fill.Visible = MSO_TRUE
        cell.Shape.Fill.ForeColor.RGB = fill_color
    if params.vertical_alignment is not None:
        cell.Shape.TextFrame.VerticalAnchor = VERTICAL_ANCHOR_MAP[
            params.vertical_alignment
        ]
    result: JsonObject = {
        "success": True,
        "row": params.row,
        "col": params.col,
        "text": cell.Shape.TextFrame.TextRange.Text,
    }
    warning = font_size_warning(params.font_size)
    if warning:
        result["warning"] = warning
    return result


@tool(
    "wps_ppt_tables",
    component="wpp",
    annotations={"title": "Set Table Data (Batch)", "idempotentHint": True},
    operation="set_data",
)
def set_table_data(params: SetTableDataInput) -> JsonObject:
    """Batch-set table cell text from a 2D array.

    Writes multiple cells at once: data=[["Q1","$1.2M"],["Q2","$1.5M"]]
    sets 2 rows x 2 cols starting from (start_row, start_col). The entire
    batch is rejected before writing if it exceeds the table boundary. Use
    bold_first_row=True to auto-bold the header row.
    """
    shape = _get_table_shape(params)
    table = shape.Table
    rows_count = table.Rows.Count
    cols_count = table.Columns.Count
    _validate_table_range(
        table,
        params.start_row,
        params.start_col,
        params.start_row + len(params.data) - 1,
        params.start_col + len(params.data[0]) - 1,
    )
    for row_index, row_data in enumerate(params.data, start=params.start_row):
        for column_index, text in enumerate(row_data, start=params.start_col):
            cell = table.Cell(row_index, column_index)
            cell.Shape.TextFrame.TextRange.Text = normalize_text(text)
            if params.bold_first_row and row_index == params.start_row:
                cell.Shape.TextFrame.TextRange.Font.Bold = MSO_TRUE
    return {
        "success": True,
        "shape_name": shape.Name,
        "cells_set": len(params.data) * len(params.data[0]),
        "rows_written": len(params.data),
        "table_rows": rows_count,
        "table_columns": cols_count,
    }


@tool(
    "wps_ppt_tables",
    component="wpp",
    annotations={"title": "Merge Table Cells"},
    operation="merge",
)
def merge_table_cells(params: MergeTableCellsInput) -> JsonObject:
    """Merge a range of table cells.

    Merges from (start_row, start_col) to (end_row, end_col). Uses
    Cell.Merge() which merges the entire rectangular range.
    """
    shape = _get_table_shape(params)
    table = shape.Table
    _validate_table_range(
        table,
        params.start_row,
        params.start_col,
        params.end_row,
        params.end_col,
    )
    cell_from = table.Cell(params.start_row, params.start_col)
    cell_to = table.Cell(params.end_row, params.end_col)
    cell_from.Merge(cell_to)
    return {
        "success": True,
        "merged": (
            f"Cell({params.start_row},{params.start_col}) "
            f"to Cell({params.end_row},{params.end_col})"
        ),
    }


@tool(
    "wps_ppt_tables",
    component="wpp",
    annotations={"title": "Add Table Row"},
    operation="add_row",
)
def add_table_row(params: TableRowInput) -> JsonObject:
    """Add a row to a table.

    If position is provided, inserts before that row (1-based). If omitted,
    appends at the end.
    """
    shape = _get_table_shape(params)
    table = shape.Table
    if params.position is not None and params.position > table.Rows.Count + 1:
        raise ValueError("position exceeds the table row insertion range")
    if params.position is not None:
        new_row = table.Rows.Add(BeforeRow=params.position)
    else:
        new_row = table.Rows.Add()
    if params.height is not None:
        new_row.Height = params.height
    return {
        "success": True,
        "new_row_count": table.Rows.Count,
        "new_column_count": table.Columns.Count,
    }


@tool(
    "wps_ppt_tables",
    component="wpp",
    annotations={"title": "Delete Table Row", "destructiveHint": True},
    operation="delete_row",
)
def delete_table_row(params: DeleteTableRowInput) -> JsonObject:
    """Delete a row from a table.

    Removes the row at the specified 1-based position. Remaining rows re-
    index automatically.
    """
    shape = _get_table_shape(params)
    table = shape.Table
    if params.position > table.Rows.Count:
        raise ValueError("position exceeds the number of table rows")
    table.Rows(params.position).Delete()
    return {
        "success": True,
        "new_row_count": table.Rows.Count,
        "new_column_count": table.Columns.Count,
    }


@tool(
    "wps_ppt_tables",
    component="wpp",
    annotations={"title": "Add Table Column"},
    operation="add_column",
)
def add_table_column(params: TableColumnInput) -> JsonObject:
    """Add a column to a table.

    If position is provided, inserts before that column (1-based). If
    omitted, appends at the end.
    """
    shape = _get_table_shape(params)
    table = shape.Table
    if (
        params.position is not None
        and params.position > table.Columns.Count + 1
    ):
        raise ValueError("position exceeds the table column insertion range")
    if params.position is not None:
        new_col = table.Columns.Add(BeforeColumn=params.position)
    else:
        new_col = table.Columns.Add()
    if params.width is not None:
        new_col.Width = params.width
    return {
        "success": True,
        "new_row_count": table.Rows.Count,
        "new_column_count": table.Columns.Count,
    }


@tool(
    "wps_ppt_tables",
    component="wpp",
    annotations={"title": "Delete Table Column", "destructiveHint": True},
    operation="delete_column",
)
def delete_table_column(params: DeleteTableColumnInput) -> JsonObject:
    """Delete a column from a table.

    Removes the column at the specified 1-based position. Remaining columns
    re-index automatically.
    """
    shape = _get_table_shape(params)
    table = shape.Table
    if params.position > table.Columns.Count:
        raise ValueError("position exceeds the number of table columns")
    table.Columns(params.position).Delete()
    return {
        "success": True,
        "new_row_count": table.Rows.Count,
        "new_column_count": table.Columns.Count,
    }


@tool(
    "wps_ppt_tables",
    component="wpp",
    annotations={"title": "Set Table Style", "idempotentHint": True},
    operation="set_style",
)
def set_table_style(params: SetTableStyleInput) -> JsonObject:
    """Apply a table style and configure banding options.

    Use a style GUID to apply a built-in style (e.g.
    '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' for Medium Style 2 - Accent
    1). Optionally toggle header row, total row, banding rows/columns.
    """
    shape = _get_table_shape(params)
    table = shape.Table
    if params.style_id is not None:
        table.ApplyStyle(params.style_id, False)
    if params.first_row is not None:
        table.FirstRow = MSO_TRUE if params.first_row else MSO_FALSE
    if params.last_row is not None:
        table.LastRow = MSO_TRUE if params.last_row else MSO_FALSE
    if params.first_col is not None:
        table.FirstCol = MSO_TRUE if params.first_col else MSO_FALSE
    if params.last_col is not None:
        table.LastCol = MSO_TRUE if params.last_col else MSO_FALSE
    if params.banding_rows is not None:
        table.HorizBanding = MSO_TRUE if params.banding_rows else MSO_FALSE
    if params.banding_cols is not None:
        table.VertBanding = MSO_TRUE if params.banding_cols else MSO_FALSE
    return {"success": True, "style_applied": params.style_id is not None}


@tool(
    "wps_ppt_tables",
    component="wpp",
    annotations={"title": "Set Table Layout", "idempotentHint": True},
    operation="set_layout",
)
def set_table_layout(params: SetTableLayoutInput) -> JsonObject:
    """Set row heights and/or column widths for an existing table.

    Provide row_heights (list of points per row) and/or col_widths (list of
    points per column), both 1-based and indexed from the first row/column.
    If either list is shorter than the table dimension, remaining
    rows/columns are left unchanged. Returns actual values after setting
    (WPS Presentation may clamp to a minimum).
    """
    shape = _get_table_shape(params)
    table = shape.Table
    if len(params.row_heights or []) > table.Rows.Count:
        raise ValueError(
            "row_heights exceeds the number of rows; nothing was changed"
        )
    if len(params.col_widths or []) > table.Columns.Count:
        raise ValueError(
            "col_widths exceeds the number of columns; nothing was changed"
        )
    for index, height in enumerate(params.row_heights or [], start=1):
        table.Rows(index).Height = height
    for index, width in enumerate(params.col_widths or [], start=1):
        table.Columns(index).Width = width
    return {
        "success": True,
        "row_heights": [
            table.Rows(i).Height for i in range(1, table.Rows.Count + 1)
        ],
        "col_widths": [
            table.Columns(i).Width for i in range(1, table.Columns.Count + 1)
        ],
    }


@tool(
    "wps_ppt_tables",
    component="wpp",
    annotations={"title": "Split Table Cells"},
    operation="split",
)
def split_table_cells(params: SplitTableCellsInput) -> JsonObject:
    """Split (unmerge) a merged table cell.

    num_rows and num_cols are the grid the cell becomes, not how many
    pieces to add, so the 1 by 1 default splits nothing. To undo a merge,
    pass the size the merge covered, num_cols=2 for two cells merged across
    or num_rows=3 for three merged down. Uses Cell.Split(NumRows,
    NumColumns) — the inverse of wps_ppt_tables/merge.
    """
    shape = _get_table_shape(params)
    table = shape.Table
    _validate_table_range(
        table, params.row, params.col, params.row, params.col
    )
    cell = table.Cell(params.row, params.col)
    cell.Split(params.num_rows, params.num_cols)
    return {
        "success": True,
        "row": params.row,
        "col": params.col,
        "num_rows": params.num_rows,
        "num_cols": params.num_cols,
    }


@tool(
    "wps_ppt_tables",
    component="wpp",
    annotations={"title": "Set Table Borders", "idempotentHint": True},
    operation="set_borders",
)
def set_table_borders(params: SetTableBordersInput) -> JsonObject:
    """Set border style for a range of table cells.

    By default applies to the entire table (start_row=1, start_col=1,
    end_row/col=last row/col). Specify sides as a list of strings: 'top',
    'bottom', 'left', 'right', 'diagonal_down', 'diagonal_up'. Optionally
    set visible, color ('#RRGGBB'), weight (points), and dash_style.
    """
    shape = _get_table_shape(params)
    table = shape.Table
    actual_end_row = (
        params.end_row if params.end_row is not None else table.Rows.Count
    )
    actual_end_col = (
        params.end_col if params.end_col is not None else table.Columns.Count
    )
    _validate_table_range(
        table,
        params.start_row,
        params.start_col,
        actual_end_row,
        actual_end_col,
    )
    side_constants = [BORDER_SIDE_MAP[side] for side in params.sides]
    color_int = hex_to_int(params.color) if params.color is not None else None
    dash_style_int = None
    if params.dash_style is not None:
        dash_style_int = dash_style_value(params.dash_style)
    diagonal_visible_skipped = False
    cells_updated = 0
    for r in range(params.start_row, actual_end_row + 1):
        for c in range(params.start_col, actual_end_col + 1):
            cell = table.Cell(r, c)
            for border_type in side_constants:
                border = cell.Borders.Item(border_type)
                if params.visible is not None:
                    if border_type in _DIAGONAL_SIDES:
                        diagonal_visible_skipped = True
                    else:
                        border.Visible = (
                            MSO_TRUE if params.visible else MSO_FALSE
                        )
                if color_int is not None:
                    border.ForeColor.RGB = color_int
                if params.weight is not None:
                    border.Weight = params.weight
                if dash_style_int is not None:
                    border.DashStyle = dash_style_int
            if side_constants:
                cells_updated += 1
    result = {
        "success": True,
        "cells_updated": cells_updated,
        "rows": f"{params.start_row}-{actual_end_row}",
        "cols": f"{params.start_col}-{actual_end_col}",
    }
    if diagonal_visible_skipped:
        result["warnings"] = [
            "WPS cannot toggle diagonal border visibility; use a color "
            "matching the cell background to hide them."
        ]
    return result
