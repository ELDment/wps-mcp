"""Validate and normalize bounded spreadsheet cell matrices."""

from collections.abc import Sequence

from src.models import CellValue, JsonObject, NativeObject
from src.spreadsheet.addresses import parse_address
from src.spreadsheet.models import MAX_CELLS


def validate_matrix(
    address: str, matrix: Sequence[Sequence[CellValue]]
) -> None:
    """Require a bounded matrix matching the exact target dimensions."""
    first_row, first_column, last_row, last_column = parse_address(address)
    row_count = last_row - first_row + 1
    column_count = last_column - first_column + 1
    if row_count * column_count > MAX_CELLS:
        raise ValueError(f"Writes are limited to {MAX_CELLS} cells per call")
    if len(matrix) != row_count or any(
        len(row) != column_count for row in matrix
    ):
        raise ValueError(
            "The matrix must be rectangular and match the "
            "exact target range dimensions"
        )


def matrix_result(
    value: object,
    row_count: int,
    column_count: int,
) -> list[list[CellValue]]:
    """Normalize native cell matrices while rejecting malformed shapes."""
    if value is None:
        return [
            [None for column in range(column_count)]
            for row in range(row_count)
        ]
    if isinstance(value, (str, int, float, bool)):
        if row_count == 1 and column_count == 1:
            return [[value]]
        raise RuntimeError("WPS returned a scalar for a multi-cell range")
    if isinstance(value, (bytes, bytearray)) or not isinstance(
        value, Sequence
    ):
        raise RuntimeError("WPS returned an invalid cell matrix")
    matrix: list[list[CellValue]] = []
    for row in value:
        if isinstance(row, (str, bytes, bytearray)) or not isinstance(
            row, Sequence
        ):
            raise RuntimeError("WPS returned an invalid matrix row")
        cells: list[CellValue] = []
        for cell in row:
            if cell is not None and not isinstance(
                cell, (str, int, float, bool)
            ):
                raise RuntimeError("WPS returned an invalid cell value")
            cells.append(cell)
        matrix.append(cells)
    if len(matrix) != row_count or any(
        len(row) != column_count for row in matrix
    ):
        raise RuntimeError(
            "WPS returned values that do not match the requested range"
        )
    return matrix


def range_values(
    target: NativeObject, *, include_formulas: bool = False
) -> JsonObject:
    """Read cell matrices after the caller has validated the range bounds."""
    rows = int(target.Rows.Count)
    columns = int(target.Columns.Count)
    result: JsonObject = {
        "values": matrix_result(target.Value2, rows, columns)
    }
    if include_formulas:
        result["formulas"] = matrix_result(target.Formula, rows, columns)
    return result
