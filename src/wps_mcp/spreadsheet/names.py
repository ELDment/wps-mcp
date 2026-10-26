"""Manage explicit workbook and worksheet scoped range names."""

import re
from typing import Literal, Self

from pydantic import Field, field_validator, model_validator

from wps_mcp.models import DocumentInput, JsonObject, NativeObject
from wps_mcp.native.runtime import get_session
from wps_mcp.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)
from wps_mcp.spreadsheet.access import get_range, get_sheet
from wps_mcp.spreadsheet.addresses import A1Address, parse_address
from wps_mcp.spreadsheet.matrices import matrix_result
from wps_mcp.spreadsheet.models import MAX_CELLS, SheetReference

LOCAL_REFERENCE = re.compile(
    r"=(?:'((?:[^']|'')+)'|([^'!]+))!"
    r"(\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?)"
)


class NameScopeInput(DocumentInput):
    """Choose workbook scope or one explicit worksheet scope."""

    scope: Literal["workbook", "sheet"] = "workbook"
    sheet: SheetReference | None = None

    @model_validator(mode="after")
    def validate_scope(self) -> Self:
        """Require a worksheet only when worksheet scope is selected."""
        if (self.scope == "sheet") != (self.sheet is not None):
            raise ValueError("sheet is required only for sheet scope")
        return self


class ListNamesInput(NameScopeInput):
    """Bound enumeration of names in one explicit native scope."""

    max_items: int = Field(default=200, ge=1, le=1000)


class NameInput(NameScopeInput):
    """Identify a supported defined name in its exact scope."""

    name: str = Field(min_length=1, max_length=255)

    @field_validator("name")
    @classmethod
    def validate_name(cls, value: str) -> str:
        """Reject blank or unsupported native object names."""
        if (
            not re.fullmatch(r"(?:[^\W\d]|_)[\w.]*", value)
            or value.casefold() in {"r", "c"}
            or re.fullmatch(
                r"(?i)(?:[a-z]{1,3}[1-9][0-9]*|r[1-9][0-9]*c[1-9][0-9]*)",
                value,
            )
            or value.casefold().startswith(("_xlnm.", "_xlfn."))
        ):
            raise ValueError(
                "Use a simple defined name, not a cell address or "
                "reserved name"
            )
        return value


class SetNameInput(NameInput):
    """Bind a scoped name to a local worksheet rectangle."""

    source_sheet: SheetReference
    address: A1Address
    visible: bool = True


class ReadNameInput(NameInput):
    """Inspect a scoped range name and optionally its bounded values."""

    include_values: bool = False
    max_cells: int = Field(default=10000, ge=1, le=MAX_CELLS)


def scoped_names(
    workbook: NativeObject, params: NameScopeInput
) -> list[NativeObject]:
    """Select names from exactly one workbook or worksheet scope."""
    sheet = (
        get_sheet(workbook, params.sheet) if params.sheet is not None else None
    )
    collection = sheet.Names if sheet is not None else workbook.Names
    if int(collection.Count) > 10000:
        raise ValueError(
            "Name enumeration is limited to 10000 names per scope"
        )
    matches = []
    for index in range(1, int(collection.Count) + 1):
        name = collection.Item(index)
        qualifier, separator, _ = str(name.Name).rpartition("!")
        if sheet is None and separator:
            continue
        if sheet is not None and (
            not separator
            or qualifier.strip("'").replace("''", "'").casefold()
            != str(sheet.Name).casefold()
        ):
            continue
        matches.append(name)
    return matches


def describe_name(name: NativeObject) -> JsonObject:
    """Classify literal local range names without evaluating formulas."""
    full_name = str(name.Name)
    qualifier, separator, simple_name = full_name.rpartition("!")
    reference = str(name.RefersTo)
    match = LOCAL_REFERENCE.fullmatch(reference)
    source_sheet = (
        (match[1].replace("''", "'") if match[1] else match[2])
        if match
        else None
    )
    address = match[3] if match else None
    is_range = source_sheet is not None and not any(
        c in source_sheet for c in "[]:/\\"
    )
    if is_range and address is not None:
        try:
            parse_address(address)
        except ValueError:
            is_range = False
    return {
        "name": simple_name,
        "full_name": full_name,
        "scope": "sheet" if separator else "workbook",
        "sheet": qualifier.strip("'").replace("''", "'")
        if separator
        else None,
        "refers_to": reference,
        "kind": "range" if is_range else "formula_or_external",
        "source_sheet": source_sheet if is_range else None,
        "address": address if is_range else None,
        "visible": bool(name.Visible),
    }


def find_name(
    workbook: NativeObject, params: NameInput
) -> NativeObject | None:
    """Find a case-insensitive name only within its explicitly chosen scope."""
    for name in scoped_names(workbook, params):
        if (
            str(name.Name).rpartition("!")[2].casefold()
            == params.name.casefold()
        ):
            return name
    return None


@tool(
    "wps_et_names",
    operation="list",
    component="et",
    annotations=READ_ANNOTATIONS,
)
def list_names(params: ListNamesInput) -> JsonObject:
    """List scoped names while distinguishing formula and external names."""
    workbook = get_session("et").get_document(params.document)
    names = scoped_names(workbook, params)
    return {
        "document": str(workbook.Name),
        "scope": params.scope,
        "names": [describe_name(name) for name in names[: params.max_items]],
        "count": len(names),
        "truncated": len(names) > params.max_items,
    }


@tool(
    "wps_et_names",
    operation="set",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def set_name(params: SetNameInput) -> JsonObject:
    """Bind a scoped name to a local rectangular worksheet range.

    Existing formula or external names cannot be overwritten.
    """
    workbook = get_session("et").get_document(params.document)
    source = get_sheet(workbook, params.source_sheet)
    target = get_range(source, params.address)
    scope_sheet = (
        get_sheet(workbook, params.sheet) if params.sheet is not None else None
    )
    existing = find_name(workbook, params)
    if existing is not None and describe_name(existing)["kind"] != "range":
        raise ValueError(
            "Cannot overwrite a formula or external name with a named range"
        )
    address = re.sub(
        r"\$?([A-Za-z]+)\$?([0-9]+)", r"$\1$\2", params.address.upper()
    )
    reference = "='" + str(source.Name).replace("'", "''") + "'!" + address
    if existing is None:
        collection = (
            scope_sheet.Names if scope_sheet is not None else workbook.Names
        )
        name = (
            "'" + str(scope_sheet.Name).replace("'", "''") + "'!" + params.name
            if scope_sheet is not None
            else params.name
        )
        existing = collection.Add(
            Name=name, RefersTo=reference, Visible=params.visible
        )
    else:
        existing.RefersTo = reference
        existing.Visible = params.visible
    return {
        "document": str(workbook.Name),
        "address": str(target.Address),
        **describe_name(existing),
    }


@tool(
    "wps_et_names",
    operation="read",
    component="et",
    annotations=READ_ANNOTATIONS,
)
def read_name(params: ReadNameInput) -> JsonObject:
    """Read local range identity and optional bounded cell values."""
    workbook = get_session("et").get_document(params.document)
    name = find_name(workbook, params)
    if name is None:
        raise ValueError(
            f"Named range not found in {params.scope} scope: {params.name}"
        )
    result = describe_name(name)
    if result["kind"] != "range":
        raise ValueError(
            "The requested name is a formula or external "
            "reference, not a local range"
        )
    sheet = get_sheet(workbook, result["source_sheet"])
    target = get_range(sheet, result["address"], params.max_cells)
    if params.include_values:
        result["values"] = matrix_result(
            target.Value2, int(target.Rows.Count), int(target.Columns.Count)
        )
    return result


@tool(
    "wps_et_names",
    operation="delete",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def delete_name(params: NameInput) -> JsonObject:
    """Remove a local range name while preserving its cell contents."""
    workbook = get_session("et").get_document(params.document)
    name = find_name(workbook, params)
    if name is None:
        raise ValueError(
            f"Named range not found in {params.scope} scope: {params.name}"
        )
    if describe_name(name)["kind"] != "range":
        raise ValueError(
            "Only local range names can be deleted by this operation"
        )
    full_name = str(name.Name)
    name.Delete()
    return {
        "document": str(workbook.Name),
        "deleted": full_name,
        "scope": params.scope,
    }
