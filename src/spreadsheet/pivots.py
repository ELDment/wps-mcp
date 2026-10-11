"""Build and refresh bounded local native pivot reports."""

import re
from typing import Literal, Self

from pydantic import Field, field_validator, model_validator

from src.models import InputModel, JsonObject, NativeObject
from src.native.runtime import get_session
from src.native.state import temporary_attributes
from src.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)
from src.spreadsheet.access import get_range, get_sheet, range_result
from src.spreadsheet.addresses import parse_address
from src.spreadsheet.matrices import matrix_result
from src.spreadsheet.models import (
    MAX_CELLS,
    RangeInput,
    SheetInput,
    SheetReference,
)
from src.spreadsheet.object_names import ObjectName
from src.spreadsheet.table_access import (
    ensure_available_name,
    get_object,
    range_bounds,
    read_headers,
)

AGGREGATIONS = {
    "sum": -4157,
    "count": -4112,
    "avg": -4106,
    "min": -4139,
    "max": -4136,
}


class PivotDataField(InputModel):
    """Select a source field, aggregation, and optional result caption."""

    field: str = Field(min_length=1, max_length=255)
    aggregate: Literal["sum", "count", "avg", "min", "max"] = "sum"
    caption: str | None = Field(default=None, min_length=1, max_length=255)


class PivotLayout(InputModel):
    """Define unique dimension fields and aggregate result columns."""

    row_fields: list[str] = Field(default_factory=list, max_length=32)
    column_fields: list[str] = Field(default_factory=list, max_length=32)
    page_fields: list[str] = Field(default_factory=list, max_length=32)
    data_fields: list[PivotDataField] = Field(min_length=1, max_length=32)
    row_grand_totals: bool = True
    column_grand_totals: bool = True

    @model_validator(mode="after")
    def validate_layout(self) -> Self:
        """Require unique dimension fields and aggregate captions."""
        dimensions = self.row_fields + self.column_fields + self.page_fields
        if any(not name.strip() or len(name) > 255 for name in dimensions):
            raise ValueError(
                "Dimension fields require nonblank names of at "
                "most 255 characters"
            )
        if len({name.casefold() for name in dimensions}) != len(dimensions):
            raise ValueError(
                "A field can appear in only one row, column or page position"
            )
        captions = [
            item.caption or f"{item.aggregate} of {item.field}"
            for item in self.data_fields
        ]
        if any(
            not caption.strip() or len(caption) > 255 for caption in captions
        ):
            raise ValueError(
                "Data captions require nonblank names of at most "
                "255 characters"
            )
        if len({caption.casefold() for caption in captions}) != len(captions):
            raise ValueError(
                "Data captions must be unique; supply distinct captions"
            )
        return self


class CreatePivotInput(RangeInput, PivotLayout):
    """Create a pivot from a local source and one destination cell."""

    name: ObjectName
    destination_sheet: SheetReference
    destination_cell: str

    @field_validator("destination_sheet")
    @classmethod
    def validate_destination_sheet(
        cls, value: SheetReference
    ) -> SheetReference:
        """Validate the destination using the common worksheet selector."""
        return SheetInput(sheet=value).sheet

    @model_validator(mode="after")
    def validate_destination(self) -> Self:
        """Require one destination cell and enough room for page fields."""
        first_row, first_column, last_row, last_column = parse_address(
            self.destination_cell
        )
        if first_row != last_row or first_column != last_column:
            raise ValueError("destination_cell must identify one cell")
        if self.page_fields and first_row < len(self.page_fields) + 2:
            raise ValueError(
                "Leave one row per page field plus a blank row "
                "above destination_cell"
            )
        return self


class PivotInput(SheetInput):
    """Select a named pivot on an explicit worksheet."""

    pivot: str = Field(min_length=1)


class ConfigurePivotInput(PivotInput, PivotLayout):
    """Replace the layout of an existing native pivot."""

    pass


class ReadPivotInput(PivotInput):
    """Bound the cell count returned from a pivot report."""

    max_cells: int = Field(default=10000, ge=1, le=MAX_CELLS)


def ensure_pivot_space(
    sheet: NativeObject,
    column: int,
    *,
    current: NativeObject | None = None,
) -> None:
    """Protect all columns at or right of growing native pivot output."""
    own_bounds = (
        range_bounds(current.TableRange2) if current is not None else None
    )
    for collection, attribute in (
        (sheet.ListObjects, "Range"),
        (sheet.PivotTables(), "TableRange2"),
    ):
        for index in range(1, int(collection.Count) + 1):
            item = collection.Item(index)
            if (
                current is not None
                and attribute == "TableRange2"
                and item.Name == current.Name
            ):
                continue
            if range_bounds(getattr(item, attribute))[3] >= column:
                raise ValueError(
                    f"Pivot output columns contain another table: {item.Name}"
                )
    used = sheet.UsedRange
    top, left, bottom, right = range_bounds(used)
    if right < column:
        return
    left = max(left, column)
    if (bottom - top + 1) * (right - left + 1) > MAX_CELLS:
        raise ValueError(
            "Pivot output safety scan exceeds the cell limit; "
            "use an unused worksheet"
        )
    target = sheet.Range(sheet.Cells(top, left), sheet.Cells(bottom, right))
    if target.MergeCells not in (False, 0):
        raise ValueError("Pivot output columns contain merged cells")
    values = matrix_result(target.Formula, bottom - top + 1, right - left + 1)
    for row_index, row in enumerate(values, top):
        for column_index, value in enumerate(row, left):
            inside = own_bounds is not None and (
                own_bounds[0] <= row_index <= own_bounds[2]
                and own_bounds[1] <= column_index <= own_bounds[3]
            )
            if not inside and value is not None and value != "":
                raise ValueError(
                    "Pivot output requires empty columns at/right of its "
                    "destination; "
                    f"existing content at row {row_index}, "
                    f"column {column_index}"
                )


def resolve_pivot_source(
    workbook: NativeObject, pivot: NativeObject
) -> NativeObject:
    """Resolve a bounded source from the same native workbook."""
    cache = pivot.PivotCache()
    if int(cache.SourceType) != 1:
        raise ValueError(
            "Only local worksheet-range pivot sources are supported"
        )
    source = cache.SourceData
    if not isinstance(source, str) or "!" not in source:
        raise ValueError(
            "Pivot source must be an explicit local worksheet rectangle"
        )
    source = source.removeprefix("=")
    if source.startswith("="):
        raise ValueError(
            "Pivot source permits one optional leading '=' "
            "before a local rectangle"
        )
    sheet_name, address = source.rsplit("!", 1)
    if sheet_name.startswith("'") and sheet_name.endswith("'"):
        sheet_name = sheet_name[1:-1].replace("''", "'")
    if sheet_name.startswith("["):
        book_name, separator, sheet_name = sheet_name[1:].partition("]")
        if (
            not separator
            or book_name.casefold() != str(workbook.Name).casefold()
        ):
            raise ValueError(
                "External workbook pivot sources cannot be refreshed"
            )
    if any(character in sheet_name for character in "[]:/\\"):
        raise ValueError("External workbook pivot sources cannot be refreshed")
    sheet = get_sheet(workbook, sheet_name)
    r1c1 = re.fullmatch(
        r"R([1-9][0-9]*)C([1-9][0-9]*):R([1-9][0-9]*)C([1-9][0-9]*)",
        address,
        re.I,
    )
    if r1c1 is not None:
        top, left, bottom, right = map(int, r1c1.groups())
        if (
            top > bottom
            or left > right
            or bottom > int(sheet.Rows.Count)
            or right > int(sheet.Columns.Count)
        ):
            raise ValueError(
                "Pivot source rectangle exceeds worksheet dimensions"
            )
        if (bottom - top + 1) * (right - left + 1) > MAX_CELLS:
            raise ValueError("Pivot source exceeds the cell limit")
        return sheet.Range(sheet.Cells(top, left), sheet.Cells(bottom, right))
    return get_range(sheet, address)


def ensure_pivot_source_space(
    sheet: NativeObject,
    pivot: NativeObject,
    source: NativeObject,
) -> None:
    """Protect a pivot source and the columns reserved for its output."""
    column = int(pivot.TableRange2.Column)
    if (
        str(source.Parent.Name) == str(sheet.Name)
        and range_bounds(source)[3] >= column
    ):
        raise ValueError("Source overlaps reserved pivot output columns")
    ensure_pivot_space(sheet, column, current=pivot)


def validate_pivot_fields(layout: PivotLayout, headers: list[str]) -> None:
    """Resolve exact source headers and reject conflicting captions."""
    requested = layout.row_fields + layout.column_fields + layout.page_fields
    requested += [item.field for item in layout.data_fields]
    for field in requested:
        if field not in headers:
            raise ValueError(
                f"Pivot field not found (names are exact): {field}"
            )
    for item in layout.data_fields:
        caption = item.caption or f"{item.aggregate} of {item.field}"
        if caption.casefold() in {header.casefold() for header in headers}:
            raise ValueError(
                f"Data caption conflicts with a source field: {caption}"
            )


def apply_pivot_layout(pivot: NativeObject, layout: PivotLayout) -> None:
    """Resolve all fields before replacing a native pivot layout."""
    fields = pivot.PivotFields()
    existing = [
        fields.Item(index) for index in range(1, int(fields.Count) + 1)
    ]
    data_fields = pivot.DataFields
    existing_data = [
        data_fields.Item(index)
        for index in range(1, int(data_fields.Count) + 1)
    ]
    cached_sources = {str(field.SourceName) for field in existing}
    cached_names = cached_sources | {
        str(field.Name) for field in existing if int(field.Orientation) != 4
    }
    requested = layout.row_fields + layout.column_fields + layout.page_fields
    requested += [item.field for item in layout.data_fields]
    resolved: dict[str, NativeObject] = {}
    for name in dict.fromkeys(requested):
        if name not in cached_sources:
            raise ValueError(
                f"Pivot cache field is unavailable: {name}; "
                "run the refresh operation first"
            )
        try:
            resolved[name] = pivot.PivotFields(name)
        except Exception as error:
            raise ValueError(
                f"Cannot resolve pivot cache field: {name}; "
                "run the refresh operation first"
            ) from error
    for item in layout.data_fields:
        caption = item.caption or f"{item.aggregate} of {item.field}"
        if caption.casefold() in {name.casefold() for name in cached_names}:
            raise ValueError(
                f"Data caption conflicts with a cached field: {caption}"
            )
    with temporary_attributes(pivot, ManualUpdate=True):
        for field in reversed(existing_data):
            field.Orientation = 0
        for field in existing:
            if int(field.Orientation) != 0:
                field.Orientation = 0
        for orientation, names in (
            (1, layout.row_fields),
            (2, layout.column_fields),
            (3, layout.page_fields),
        ):
            for position, name in enumerate(names, 1):
                field = resolved[name]
                field.Orientation = orientation
                field.Position = position
        for item in layout.data_fields:
            pivot.AddDataField(
                resolved[item.field],
                item.caption or f"{item.aggregate} of {item.field}",
                AGGREGATIONS[item.aggregate],
            )
        pivot.RowGrand = layout.row_grand_totals
        pivot.ColumnGrand = layout.column_grand_totals


def describe_pivot(sheet: NativeObject, pivot: NativeObject) -> JsonObject:
    """Detach the pivot name and its complete native report bounds."""
    return {"name": str(pivot.Name), **range_result(sheet, pivot.TableRange2)}


@tool(
    "wps_et_pivot_tables",
    operation="list",
    component="et",
    annotations=READ_ANNOTATIONS,
)
def list_pivots(params: SheetInput) -> JsonObject:
    """List native pivot tables on the selected worksheet."""
    sheet = get_sheet(
        get_session("et").get_document(params.document), params.sheet
    )
    pivots = sheet.PivotTables()
    return {
        "sheet": str(sheet.Name),
        "pivots": [
            describe_pivot(sheet, pivots.Item(index))
            for index in range(1, int(pivots.Count) + 1)
        ],
    }


@tool(
    "wps_et_pivot_tables",
    operation="create",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def create_pivot(params: CreatePivotInput) -> JsonObject:
    """Create a native report from this workbook's literal-header rectangle.

    Destination columns and all columns to their right must contain no data
    or other tables. This reserves space for native report growth. On the
    same worksheet, source data must lie entirely to the left of the
    destination.
    """
    workbook = get_session("et").get_document(params.document)
    source_sheet = get_sheet(workbook, params.sheet)
    source = get_range(source_sheet, params.address)
    headers = read_headers(source)
    validate_pivot_fields(params, headers)
    sheet = get_sheet(workbook, params.destination_sheet)
    destination = get_range(sheet, params.destination_cell)
    if destination.MergeCells not in (False, 0):
        raise ValueError("Pivot destination cannot contain merged cells")
    ensure_available_name(workbook, params.name)
    if str(source_sheet.Name) == str(sheet.Name) and range_bounds(source)[
        3
    ] >= int(destination.Column):
        raise ValueError(
            "Source overlaps reserved pivot output columns; "
            "put it left of destination"
        )
    ensure_pivot_space(sheet, int(destination.Column))
    top, left, bottom, right = range_bounds(source)
    escaped = str(source_sheet.Name).replace("'", "''")
    cache = workbook.PivotCaches().Create(
        1, f"'{escaped}'!R{top}C{left}:R{bottom}C{right}"
    )
    pivot = cache.CreatePivotTable(destination, params.name)
    apply_pivot_layout(pivot, params)
    return describe_pivot(sheet, pivot)


@tool(
    "wps_et_pivot_tables",
    operation="configure",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def configure_pivot(params: ConfigurePivotInput) -> JsonObject:
    """Replace validated pivot fields and grand-total settings.

    Uses the existing local cache source. All output columns must be free of
    content outside this pivot; source data must stay outside those columns.
    Refresh explicitly after source-header changes; stale cache fields are
    rejected before any existing layout assignments are changed.
    """
    workbook = get_session("et").get_document(params.document)
    sheet = get_sheet(workbook, params.sheet)
    pivot = get_object(sheet, params.pivot, pivot=True)
    source = resolve_pivot_source(workbook, pivot)
    validate_pivot_fields(params, read_headers(source))
    if (
        params.page_fields
        and int(pivot.TableRange1.Row) < len(params.page_fields) + 2
    ):
        raise ValueError(
            "Not enough rows above the pivot for the requested page fields"
        )
    ensure_pivot_source_space(sheet, pivot, source)
    apply_pivot_layout(pivot, params)
    return describe_pivot(sheet, pivot)


@tool(
    "wps_et_pivot_tables",
    operation="refresh",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def refresh_pivot(params: PivotInput) -> JsonObject:
    """Refresh after protecting all reports sharing the native cache."""
    workbook = get_session("et").get_document(params.document)
    sheet = get_sheet(workbook, params.sheet)
    pivot = get_object(sheet, params.pivot, pivot=True)
    source = resolve_pivot_source(workbook, pivot)
    read_headers(source)
    for sheet_index in range(1, int(workbook.Worksheets.Count) + 1):
        sibling_sheet = workbook.Worksheets.Item(sheet_index)
        siblings = sibling_sheet.PivotTables()
        for index in range(1, int(siblings.Count) + 1):
            sibling = siblings.Item(index)
            if int(sibling.CacheIndex) == int(pivot.CacheIndex):
                ensure_pivot_source_space(sibling_sheet, sibling, source)
    if pivot.RefreshTable() is False:
        raise RuntimeError(
            "WPS did not refresh the pivot table; inspect before retrying"
        )
    return describe_pivot(sheet, pivot)


@tool(
    "wps_et_pivot_tables",
    operation="read",
    component="et",
    annotations=READ_ANNOTATIONS,
)
def read_pivot(params: ReadPivotInput) -> JsonObject:
    """Read bounded pivot results, field assignments, and aggregations."""
    sheet = get_sheet(
        get_session("et").get_document(params.document), params.sheet
    )
    pivot = get_object(sheet, params.pivot, pivot=True)
    target = get_range(sheet, str(pivot.TableRange2.Address), params.max_cells)
    result = describe_pivot(sheet, pivot)
    result["values"] = matrix_result(
        target.Value2, result["rows"], result["columns"]
    )
    fields = pivot.PivotFields()
    result["fields"] = [
        {
            "name": str(fields.Item(index).Name),
            "orientation": int(fields.Item(index).Orientation),
        }
        for index in range(1, int(fields.Count) + 1)
    ]
    data_fields = pivot.DataFields
    result["data_fields"] = [
        {
            "field": str(data_fields.Item(index).SourceName),
            "caption": str(data_fields.Item(index).Name),
            "aggregate": next(
                (
                    key
                    for key, value in AGGREGATIONS.items()
                    if value == int(data_fields.Item(index).Function)
                ),
                "other",
            ),
        }
        for index in range(1, int(data_fields.Count) + 1)
    ]
    return result


@tool(
    "wps_et_pivot_tables",
    operation="delete",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def delete_pivot(params: PivotInput) -> JsonObject:
    """Remove exactly this pivot and its report cells; preserve source data."""
    sheet = get_sheet(
        get_session("et").get_document(params.document), params.sheet
    )
    pivot = get_object(sheet, params.pivot, pivot=True)
    result = describe_pivot(sheet, pivot)
    pivot.TableRange2.Clear()
    return {**result, "deleted": True}
