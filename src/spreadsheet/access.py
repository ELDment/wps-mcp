"""Resolve native worksheets and ranges before accessing cells."""

from src.models import JsonObject, NativeObject
from src.native.runtime import get_session
from src.spreadsheet.addresses import parse_address
from src.spreadsheet.models import MAX_CELLS, RangeInput, SheetReference


def get_sheet(
    workbook: NativeObject, reference: SheetReference
) -> NativeObject:
    """Resolve an existing worksheet by case-insensitive name or index."""
    sheets = workbook.Worksheets
    if isinstance(reference, bool):
        raise TypeError(
            "Worksheet reference must be a name or one-based index"
        )
    if isinstance(reference, int):
        if reference < 1 or reference > sheets.Count:
            raise ValueError(
                f"Worksheet index {reference} is outside 1..{sheets.Count}"
            )
        return sheets.Item(reference)
    for index in range(1, sheets.Count + 1):
        sheet = sheets.Item(index)
        if str(sheet.Name).casefold() == reference.casefold():
            return sheet
    raise ValueError(f"Worksheet not found: {reference}")


def get_range(
    sheet: NativeObject, address: str, max_cells: int | None = MAX_CELLS
) -> NativeObject:
    """Validate worksheet bounds and cell limits before native range lookup."""
    first_row, first_column, last_row, last_column = parse_address(address)
    if last_row > sheet.Rows.Count or last_column > sheet.Columns.Count:
        raise ValueError("Range exceeds this workbook's worksheet dimensions")
    if (
        max_cells is not None
        and (last_row - first_row + 1) * (last_column - first_column + 1)
        > max_cells
    ):
        raise ValueError(
            f"Range exceeds the {max_cells}-cell limit; use smaller ranges"
        )
    return sheet.Range(address)


def resolve_range(
    params: RangeInput, max_cells: int | None = MAX_CELLS
) -> tuple[NativeObject, NativeObject]:
    """Resolve the selected document and validate its native range bounds."""
    workbook = get_session("et").get_document(params.document)
    sheet = get_sheet(workbook, params.sheet)
    return (sheet, get_range(sheet, params.address, max_cells))


def range_result(sheet: NativeObject, target: NativeObject) -> JsonObject:
    """Detach a range address, dimensions, and worksheet name."""
    return {
        "sheet": str(sheet.Name),
        "address": str(target.Address),
        "rows": int(target.Rows.Count),
        "columns": int(target.Columns.Count),
    }
