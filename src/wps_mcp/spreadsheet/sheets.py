"""Manage worksheet names, recalculation, and frozen panes."""

from pydantic import (
    Field,
)

from wps_mcp.models import DocumentInput, JsonObject, NativeObject
from wps_mcp.native.runtime import get_session
from wps_mcp.native.state import temporary_attributes
from wps_mcp.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)
from wps_mcp.spreadsheet.access import get_sheet
from wps_mcp.spreadsheet.models import SheetInput
from wps_mcp.spreadsheet.object_names import SheetName, validate_sheet_name


class AddSheetInput(DocumentInput):
    """Create a worksheet with a validated native sheet name."""

    name: SheetName = Field(min_length=1, max_length=31)


class RenameSheetInput(SheetInput):
    """Rename a selected worksheet with a validated native name."""

    name: SheetName = Field(min_length=1, max_length=31)


class FreezeInput(SheetInput):
    """Freeze leading rows and columns or unfreeze with zero counts."""

    rows: int = Field(default=0, ge=0)
    columns: int = Field(default=0, ge=0)


def ensure_sheet_name(
    workbook: NativeObject, name: str, current: NativeObject | None = None
) -> None:
    """Reject sheet-name collisions before creating or renaming a sheet."""
    validate_sheet_name(name)
    for index in range(1, workbook.Sheets.Count + 1):
        existing = workbook.Sheets.Item(index)
        if str(existing.Name).casefold() == name.casefold() and (
            current is None or str(existing.Name) != str(current.Name)
        ):
            raise ValueError(f"A sheet already uses the name: {name}")


@tool(
    "wps_et_sheets",
    operation="list",
    component="et",
    annotations=READ_ANNOTATIONS,
)
def list_sheets(params: DocumentInput) -> JsonObject:
    """List worksheets in an open WPS ET workbook."""
    workbook = get_session("et").get_document(params.document)
    sheets = workbook.Worksheets
    return {
        "document": str(workbook.Name),
        "sheets": [
            {
                "index": index,
                "name": str(sheets.Item(index).Name),
                "visible": int(sheets.Item(index).Visible) == -1,
            }
            for index in range(1, sheets.Count + 1)
        ],
    }


@tool(
    "wps_et_sheets",
    operation="add",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def add_sheet(params: AddSheetInput) -> JsonObject:
    """Append a named worksheet to a WPS ET workbook."""
    workbook = get_session("et").get_document(params.document)
    ensure_sheet_name(workbook, params.name)
    sheet = workbook.Worksheets.Add(
        After=workbook.Sheets.Item(workbook.Sheets.Count)
    )
    sheet.Name = params.name
    return {"name": str(sheet.Name), "index": int(sheet.Index)}


@tool(
    "wps_et_sheets",
    operation="rename",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def rename_sheet(params: RenameSheetInput) -> JsonObject:
    """Rename a worksheet without changing its contents."""
    workbook = get_session("et").get_document(params.document)
    sheet = get_sheet(workbook, params.sheet)
    ensure_sheet_name(workbook, params.name, sheet)
    previous_name = str(sheet.Name)
    sheet.Name = params.name
    return {"previous_name": previous_name, "name": str(sheet.Name)}


@tool(
    "wps_et_sheets",
    operation="delete",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def delete_sheet(params: SheetInput) -> JsonObject:
    """Delete one worksheet while retaining a visible workbook sheet.

    The selected sheet and its cell contents are permanently removed.
    """
    workbook = get_session("et").get_document(params.document)
    sheet = get_sheet(workbook, params.sheet)
    visible_count = sum(
        int(workbook.Sheets.Item(index).Visible) == -1
        for index in range(1, workbook.Sheets.Count + 1)
    )
    if int(sheet.Visible) == -1 and visible_count <= 1:
        raise ValueError("Cannot delete the workbook's last visible sheet")
    sheet_name = str(sheet.Name)
    application = get_session("et").get_application()
    with temporary_attributes(application, DisplayAlerts=False):
        sheet.Delete()
    return {"deleted": sheet_name}


@tool(
    "wps_et_sheets",
    operation="recalculate",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def recalculate(params: DocumentInput) -> JsonObject:
    """Recalculate each worksheet of the selected workbook."""
    workbook = get_session("et").get_document(params.document)
    for index in range(1, workbook.Worksheets.Count + 1):
        workbook.Worksheets.Item(index).Calculate()
    return {
        "document": str(workbook.Name),
        "recalculated_sheets": int(workbook.Worksheets.Count),
    }


@tool(
    "wps_et_sheets",
    operation="freeze",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def freeze_panes(params: FreezeInput) -> JsonObject:
    """Freeze leading worksheet rows or columns in its native window.

    Zero row and column counts unfreeze panes. The selected worksheet and
    window are activated.
    """
    workbook = get_session("et").get_document(params.document)
    sheet = get_sheet(workbook, params.sheet)
    if (
        params.rows >= sheet.Rows.Count
        or params.columns >= sheet.Columns.Count
    ):
        raise ValueError(
            "Frozen rows/columns must leave at least one scrollable row/column"
        )
    if int(sheet.Visible) != -1:
        raise ValueError("Cannot freeze panes on a hidden worksheet")
    if workbook.Windows.Count < 1:
        raise ValueError("The workbook has no available window")
    window = workbook.Windows.Item(1)
    window.Activate()
    sheet.Activate()
    window.FreezePanes = False
    window.SplitRow = params.rows
    window.SplitColumn = params.columns
    window.FreezePanes = bool(params.rows or params.columns)
    return {
        "sheet": str(sheet.Name),
        "rows": params.rows,
        "columns": params.columns,
    }
