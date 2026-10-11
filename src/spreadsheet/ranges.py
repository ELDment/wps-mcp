"""Read, write, search, and format native spreadsheet ranges."""

from typing import Literal, Self

from pydantic import (
    Field,
    model_validator,
)

from src.common.colors import hex_to_int
from src.models import CellValue, Color, JsonObject
from src.native.runtime import get_session
from src.native.state import temporary_attributes
from src.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)
from src.spreadsheet.access import get_sheet, range_result, resolve_range
from src.spreadsheet.addresses import parse_address
from src.spreadsheet.matrices import (
    matrix_result,
    range_values,
    validate_matrix,
)
from src.spreadsheet.models import MAX_CELLS, RangeInput, SheetInput

HORIZONTAL_ALIGNMENTS = {"left": -4131, "center": -4108, "right": -4152}


VERTICAL_ALIGNMENTS = {"top": -4160, "center": -4108, "bottom": -4107}


class ReadRangeInput(RangeInput):
    """Read a bounded rectangle with optional formula values."""

    max_cells: int = Field(default=10000, ge=1, le=MAX_CELLS)
    include_formulas: bool = False


class WriteValuesInput(RangeInput):
    """Write a literal matrix matching the exact target rectangle."""

    values: list[list[CellValue]] = Field(min_length=1)

    @model_validator(mode="after")
    def validate_values(self) -> Self:
        """Require literal matrix dimensions to match the target range."""
        validate_matrix(self.address, self.values)
        return self


class WriteFormulasInput(RangeInput):
    """Write an explicit formula matrix or clear cells with null."""

    formulas: list[list[str | None]] = Field(min_length=1)

    @model_validator(mode="after")
    def validate_formulas(self) -> Self:
        """Require exact matrix dimensions and explicit formula strings."""
        validate_matrix(self.address, self.formulas)
        if any(
            value is not None and (not value.startswith("="))
            for row in self.formulas
            for value in row
        ):
            raise ValueError(
                "Each formula must begin with '='; use null to clear a cell"
            )
        return self


class FormatRangeInput(RangeInput):
    """Set optional cell formatting and row or column dimensions."""

    font_name: str | None = Field(default=None, min_length=1)
    font_size: float | None = Field(default=None, gt=0, le=409)
    bold: bool | None = None
    italic: bool | None = None
    font_color: Color | None = None
    fill_color: Color | None = None
    number_format: str | None = Field(default=None, min_length=1)
    horizontal_alignment: Literal["left", "center", "right"] | None = None
    vertical_alignment: Literal["top", "center", "bottom"] | None = None
    wrap_text: bool | None = None
    border_style: Literal["none", "continuous", "dash", "dot"] | None = None
    border_color: Color | None = None
    column_width: float | None = Field(default=None, gt=0, le=255)
    row_height: float | None = Field(default=None, gt=0, le=409)
    auto_fit_columns: bool = False
    auto_fit_rows: bool = False

    @model_validator(mode="after")
    def validate_dimensions(self) -> Self:
        """Reject conflicting explicit and automatic sizing options."""
        if self.auto_fit_columns and self.column_width is not None:
            raise ValueError(
                "Choose auto_fit_columns or column_width, not both"
            )
        if self.auto_fit_rows and self.row_height is not None:
            raise ValueError("Choose auto_fit_rows or row_height, not both")
        return self


class MergeRangeInput(RangeInput):
    """Merge or unmerge cells with explicit data-loss acknowledgement."""

    merge: bool = True
    discard_other_cells: bool = Field(
        default=False,
        description=(
            "Allow merging to discard values outside the upper-left cell"
        ),
    )


class ClearRangeInput(RangeInput):
    """Choose which contents or formats to clear from a rectangle."""

    mode: Literal["contents", "formats", "all"] = "contents"


class EditBandInput(SheetInput):
    """Insert or delete a bounded run of complete rows or columns."""

    axis: Literal["rows", "columns"]
    action: Literal["insert", "delete"]
    start: int = Field(ge=1)
    count: int = Field(default=1, ge=1, le=10000)


class SortRangeInput(RangeInput):
    """Sort a rectangle by a validated relative key column."""

    key_column: int = Field(
        default=1,
        ge=1,
        description="Sort key column relative to the supplied range",
    )
    descending: bool = False
    has_header: bool = True

    @model_validator(mode="after")
    def validate_key(self) -> Self:
        """Require a sort key and data rows inside the selected rectangle."""
        first_row, first_column, last_row, last_column = parse_address(
            self.address
        )
        if self.key_column > last_column - first_column + 1:
            raise ValueError("key_column is outside the selected range")
        if last_row == first_row and self.has_header:
            raise ValueError(
                "The selected range has a header but no data rows"
            )
        return self


class FilterRangeInput(RangeInput):
    """Filter one relative field in a rectangle containing headers."""

    field: int = Field(
        default=1,
        ge=1,
        description="One-based column relative to the selected range",
    )
    criteria: str | int | float | None = Field(
        default=None,
        description="WPS filter criterion; null clears this field's filter",
    )

    @model_validator(mode="after")
    def validate_field(self) -> Self:
        """Require a filter field and data rows inside the target range."""
        first_row, first_column, last_row, last_column = parse_address(
            self.address
        )
        if self.field > last_column - first_column + 1:
            raise ValueError("field is outside the selected range")
        if first_row == last_row:
            raise ValueError(
                "Filtering requires a header row and at least one data row"
            )
        return self


class FindRangeInput(RangeInput):
    """Bound native value or formula search results."""

    text: str = Field(min_length=1)
    match_case: bool = False
    whole_cell: bool = False
    in_formulas: bool = False
    max_matches: int = Field(default=100, ge=1, le=1000)


@tool(
    "wps_et_ranges",
    operation="read",
    component="et",
    annotations=READ_ANNOTATIONS,
)
def read_range(params: ReadRangeInput) -> JsonObject:
    """Read cell values and optionally formulas, bounded by max_cells."""
    sheet, target = resolve_range(params, params.max_cells)
    result = range_result(sheet, target)
    result.update(
        range_values(target, include_formulas=params.include_formulas)
    )
    return result


@tool(
    "wps_et_ranges",
    operation="write_values",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def write_values(params: WriteValuesInput) -> JsonObject:
    """Write an exact rectangular matrix of literal values.

    Strings never become formulas; null clears cells.
    """
    sheet, target = resolve_range(params)
    if target.MergeCells not in (False, 0):
        raise ValueError(
            "The target contains merged cells; explicitly "
            "unmerge the range before writing values"
        )
    literal_values = tuple(
        tuple(
            "'" + value if isinstance(value, str) else value for value in row
        )
        for row in params.values
    )
    target.Value2 = (
        literal_values[0][0]
        if len(literal_values) == 1 and len(literal_values[0]) == 1
        else literal_values
    )
    return {**range_result(sheet, target), "mode": "literal"}


@tool(
    "wps_et_ranges",
    operation="write_formulas",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def write_formulas(params: WriteFormulasInput) -> JsonObject:
    """Write an exact matrix of explicit formulas and nullable cells.

    A1 formulas must begin with =; null clears a cell. WPS evaluates
    expressions.
    """
    sheet, target = resolve_range(params)
    if target.MergeCells not in (False, 0):
        raise ValueError(
            "The target contains merged cells; explicitly "
            "unmerge the range before writing formulas"
        )
    formulas = tuple(tuple(row) for row in params.formulas)
    target.Formula = (
        formulas[0][0]
        if len(formulas) == 1 and len(formulas[0]) == 1
        else formulas
    )
    return {**range_result(sheet, target), "mode": "formula"}


@tool(
    "wps_et_ranges",
    operation="format",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def format_range(params: FormatRangeInput) -> JsonObject:
    """Set fonts, fills, borders, number formats, alignment and dimensions.

    Width/height affect entire columns/rows.
    """
    sheet, target = resolve_range(params)
    for attribute, value in (
        ("Name", params.font_name),
        ("Size", params.font_size),
        ("Bold", params.bold),
        ("Italic", params.italic),
    ):
        if value is not None:
            setattr(target.Font, attribute, value)
    if params.font_color is not None:
        target.Font.Color = hex_to_int(params.font_color)
    if params.fill_color is not None:
        target.Interior.Color = hex_to_int(params.fill_color)
    for attribute, value in (
        ("NumberFormat", params.number_format),
        ("WrapText", params.wrap_text),
        ("ColumnWidth", params.column_width),
        ("RowHeight", params.row_height),
    ):
        if value is not None:
            setattr(target, attribute, value)
    if params.horizontal_alignment is not None:
        target.HorizontalAlignment = HORIZONTAL_ALIGNMENTS[
            params.horizontal_alignment
        ]
    if params.vertical_alignment is not None:
        target.VerticalAlignment = VERTICAL_ALIGNMENTS[
            params.vertical_alignment
        ]
    if params.border_style is not None:
        target.Borders.LineStyle = {
            "none": -4142,
            "continuous": 1,
            "dash": -4115,
            "dot": -4118,
        }[params.border_style]
    if params.border_color is not None:
        target.Borders.Color = hex_to_int(params.border_color)
    if params.auto_fit_columns:
        target.Columns.AutoFit()
    if params.auto_fit_rows:
        target.Rows.AutoFit()
    return range_result(sheet, target)


@tool(
    "wps_et_ranges",
    operation="merge",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def merge_range(params: MergeRangeInput) -> JsonObject:
    """Merge or unmerge cells.

    Merging keeps only the upper-left value; data loss requires explicit
    opt-in.
    """
    sheet, target = resolve_range(params)
    if params.merge:
        if not params.discard_other_cells:
            values = matrix_result(
                target.Formula,
                int(target.Rows.Count),
                int(target.Columns.Count),
            )
            if any(
                (
                    value is not None and value != ""
                    for row_index, row in enumerate(values)
                    for column_index, value in enumerate(row)
                    if row_index or column_index
                )
            ):
                raise ValueError(
                    "Merge would discard cell values; set "
                    "discard_other_cells=true to allow it"
                )
        if params.discard_other_cells:
            application = get_session("et").get_application()
            with temporary_attributes(application, DisplayAlerts=False):
                target.Merge()
        else:
            target.Merge()
    else:
        target.UnMerge()
    return {**range_result(sheet, target), "merged": params.merge}


@tool(
    "wps_et_ranges",
    operation="clear",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def clear_range(params: ClearRangeInput) -> JsonObject:
    """Clear cell contents, formats, or both within an explicit range."""
    sheet, target = resolve_range(params)
    if params.mode == "contents":
        target.ClearContents()
    elif params.mode == "formats":
        target.ClearFormats()
    else:
        target.Clear()
    return {**range_result(sheet, target), "cleared": params.mode}


@tool(
    "wps_et_ranges",
    operation="edit_bands",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def edit_band(params: EditBandInput) -> JsonObject:
    """Insert or delete complete native worksheet rows or columns.

    Native indices and formula references may shift.
    """
    workbook = get_session("et").get_document(params.document)
    sheet = get_sheet(workbook, params.sheet)
    last = params.start + params.count - 1
    maximum = int(
        sheet.Rows.Count if params.axis == "rows" else sheet.Columns.Count
    )
    if last > maximum:
        raise ValueError(
            f"Requested {params.axis} exceed this worksheet's dimensions"
        )
    if params.axis == "rows":
        band = sheet.Range(
            sheet.Cells(params.start, 1), sheet.Cells(last, 1)
        ).EntireRow
    else:
        band = sheet.Range(
            sheet.Cells(1, params.start), sheet.Cells(1, last)
        ).EntireColumn
    if params.action == "insert":
        band.Insert()
    else:
        band.Delete()
    return {
        "sheet": str(sheet.Name),
        "axis": params.axis,
        "action": params.action,
        "start": params.start,
        "count": params.count,
    }


@tool(
    "wps_et_ranges",
    operation="sort",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def sort_range(params: SortRangeInput) -> JsonObject:
    """Sort rows by one column within the selected rectangular range."""
    sheet, target = resolve_range(params)
    target.Sort(
        Key1=target.Columns.Item(params.key_column),
        Order1=2 if params.descending else 1,
        Header=1 if params.has_header else 2,
        Orientation=1,
    )
    return range_result(sheet, target)


@tool(
    "wps_et_ranges",
    operation="filter",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def filter_range(params: FilterRangeInput) -> JsonObject:
    """Set or clear a native filter criterion on one relative column.

    A null criterion clears the selected field filter.
    """
    sheet, target = resolve_range(params)
    if params.criteria is None:
        target.AutoFilter(Field=params.field)
    else:
        target.AutoFilter(Field=params.field, Criteria1=params.criteria)
    return {
        **range_result(sheet, target),
        "field": params.field,
        "criteria": params.criteria,
    }


@tool(
    "wps_et_ranges",
    operation="clear_filter",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def clear_filter(params: SheetInput) -> JsonObject:
    """Show all rows hidden by worksheet filters."""
    sheet = get_sheet(
        get_session("et").get_document(params.document), params.sheet
    )
    if sheet.FilterMode:
        sheet.ShowAllData()
    return {"sheet": str(sheet.Name), "filtered": False}


@tool(
    "wps_et_ranges",
    operation="find",
    component="et",
    annotations=READ_ANNOTATIONS,
)
def find_range(params: FindRangeInput) -> JsonObject:
    """Find bounded values or formulas using native wildcard matching.

    Return at most max_matches and report truncation.
    """
    sheet, target = resolve_range(params)
    results: list[JsonObject] = []
    seen = set()
    after = target.Cells.Item(
        int(target.Rows.Count), int(target.Columns.Count)
    )
    while True:
        found = target.Find(
            What=params.text,
            After=after,
            LookIn=-4123 if params.in_formulas else -4163,
            LookAt=1 if params.whole_cell else 2,
            SearchOrder=1,
            SearchDirection=1,
            MatchCase=params.match_case,
            SearchFormat=False,
        )
        if found is None or str(found.Address) in seen:
            return {
                "sheet": str(sheet.Name),
                "matches": results,
                "truncated": False,
            }
        if len(results) == params.max_matches:
            return {
                "sheet": str(sheet.Name),
                "matches": results,
                "truncated": True,
            }
        address = str(found.Address)
        seen.add(address)
        results.append(
            {
                "address": address,
                "value": found.Value2,
                "formula": found.Formula,
            }
        )
        after = found
