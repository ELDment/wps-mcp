"""Preflight shared native table and pivot source boundaries."""

from wps_mcp.models import NativeObject
from wps_mcp.spreadsheet.matrices import matrix_result


def range_bounds(target: NativeObject) -> tuple[int, int, int, int]:
    """Return inclusive native row and column bounds for a rectangle."""
    return (
        int(target.Row),
        int(target.Column),
        int(target.Row) + int(target.Rows.Count) - 1,
        int(target.Column) + int(target.Columns.Count) - 1,
    )


def ranges_overlap(first: NativeObject, second: NativeObject) -> bool:
    """Check whether two native rectangles share at least one cell."""
    top, left, bottom, right = range_bounds(first)
    other_top, other_left, other_bottom, other_right = range_bounds(second)
    return (
        top <= other_bottom
        and other_top <= bottom
        and left <= other_right
        and other_left <= right
    )


def get_object(
    sheet: NativeObject, name: str, *, pivot: bool = False
) -> NativeObject:
    """Resolve a named table or pivot from its worksheet collection."""
    collection = sheet.PivotTables() if pivot else sheet.ListObjects
    for index in range(1, int(collection.Count) + 1):
        item = collection.Item(index)
        if str(item.Name).casefold() == name.casefold():
            return item
    raise ValueError(
        f"{'Pivot table' if pivot else 'Table'} not found: {name}"
    )


def ensure_available_name(
    workbook: NativeObject,
    name: str,
    *,
    current: str | None = None,
) -> None:
    """Reject object-name conflicts across tables, pivots, and names."""
    if current is not None and name.casefold() == current.casefold():
        return
    for sheet_index in range(1, int(workbook.Worksheets.Count) + 1):
        sheet = workbook.Worksheets.Item(sheet_index)
        for collection in (sheet.ListObjects, sheet.PivotTables()):
            for index in range(1, int(collection.Count) + 1):
                if (
                    str(collection.Item(index).Name).casefold()
                    == name.casefold()
                ):
                    raise ValueError(
                        f"A table or pivot table already uses the name: {name}"
                    )
    for index in range(1, int(workbook.Names.Count) + 1):
        if (
            str(workbook.Names.Item(index).Name).rsplit("!", 1)[-1].casefold()
            == name.casefold()
        ):
            raise ValueError(f"A defined name already uses the name: {name}")


def read_headers(source: NativeObject) -> list[str]:
    """Require literal headers that native table creation can preserve."""
    if source.MergeCells not in (False, 0):
        raise ValueError("Source ranges cannot contain merged cells")
    if int(source.Rows.Count) < 2:
        raise ValueError(
            "Source requires a header row and at least one data row"
        )
    header = source.Rows.Item(1)
    values = matrix_result(header.Value2, 1, int(source.Columns.Count))[0]
    formulas = matrix_result(header.Formula, 1, int(source.Columns.Count))[0]
    if any(
        not isinstance(value, str) or not value.strip() or len(value) > 255
        for value in values
    ):
        raise ValueError(
            "Every column header must be nonblank text of at "
            "most 255 characters"
        )
    if any(
        isinstance(formula, str) and formula.startswith("=")
        for formula in formulas
    ):
        raise ValueError(
            "Column headers must contain literal text, not formulas"
        )
    headers = [str(value) for value in values]
    if len({header.casefold() for header in headers}) != len(headers):
        raise ValueError(
            "Column headers must be unique without native automatic renaming"
        )
    return headers


def ensure_table_space(
    sheet: NativeObject,
    target: NativeObject,
    *,
    current: str | None = None,
) -> None:
    """Reject table or pivot overlap before creating or resizing a table."""
    for index in range(1, int(sheet.ListObjects.Count) + 1):
        item = sheet.ListObjects.Item(index)
        if str(item.Name) != current and ranges_overlap(target, item.Range):
            raise ValueError(f"Range overlaps structured table: {item.Name}")
    pivots = sheet.PivotTables()
    for index in range(1, int(pivots.Count) + 1):
        item = pivots.Item(index)
        if ranges_overlap(target, item.TableRange2):
            raise ValueError(f"Range overlaps pivot table: {item.Name}")
