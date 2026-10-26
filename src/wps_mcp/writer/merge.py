"""Merge local CSV records into a new native Writer document."""

import re
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Self

from pydantic import Field, field_validator, model_validator

from wps_mcp.common.paths import local_path
from wps_mcp.documents.service import describe_document
from wps_mcp.errors import WpsError
from wps_mcp.models import DocumentInput, JsonObject, NativeObject
from wps_mcp.native.context import CALL_DOCUMENT
from wps_mcp.native.runtime import get_session
from wps_mcp.native.selection import document_key
from wps_mcp.native.state import temporary_attributes
from wps_mcp.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)
from wps_mcp.writer.access import resolve_range
from wps_mcp.writer.models import Index, Position

MERGE_STATES = {
    0: "normal_document",
    1: "main_document",
    2: "main_and_source",
    3: "main_and_header",
    4: "main_source_and_header",
    5: "data_source",
}


CONNECTED_STATES = {2, 4}


CONNECTION_PROPERTY = re.compile(
    r"\s*([^=;]+?)\s*=\s*"
    r"(?:\"((?:[^\"]|\"\")*)\"|'((?:[^']|'')*)'|"
    r"\{((?:[^}]|}})*)\}|([^;]*))"
    r"\s*(?:;|$)"
)


SOURCE_QUERY = re.compile(
    r"\s*SELECT\s+\*\s+FROM\s+"
    r'(?:`([^`]+)`|\[([^\]]+)\]|"([^\"]+)"|([\w.-]+))'
    r"\s*;?\s*",
    re.IGNORECASE,
)


class MergeFieldInput(DocumentInput):
    """Insert a named merge field without accepting arbitrary field code."""

    position: Position
    name: str = Field(
        min_length=1, max_length=255, pattern=r'^[^"{}\\\x00-\x1f]+$'
    )

    @model_validator(mode="after")
    def validate_name(self) -> Self:
        """Reject blank or unsupported native object names."""
        if not self.name.strip():
            raise ValueError("Merge field name cannot be blank")
        return self


class MergeSourceInput(DocumentInput):
    """Attach an existing local CSV as the mail-merge data source."""

    file_path: str = Field(min_length=1)

    @field_validator("file_path")
    @classmethod
    def validate_source(cls, value: str) -> str:
        """Require a supported local source selector."""
        extension = Path(value).suffix.lower()
        if extension == ".xlsx":
            raise ValueError(
                "Export the desired worksheet to CSV before connecting; "
                "native XLSX selection cannot run unattended"
            )
        if extension != ".csv":
            raise ValueError(
                "Mail merge supports only local .csv data sources"
            )
        return value


class MergeExecuteInput(DocumentInput):
    """Select a bounded record interval for new-document output."""

    first_record: Index = 1
    last_record: Index | None = None
    suppress_blank_lines: bool = True

    @model_validator(mode="after")
    def validate_records(self) -> Self:
        """Reject reversed mail-merge record intervals."""
        if (
            self.last_record is not None
            and self.last_record < self.first_record
        ):
            raise ValueError(
                "last_record must be greater than or equal to first_record"
            )
        return self


@contextmanager
def native_mail_merge(document: NativeObject) -> Iterator[NativeObject]:
    """Report unsupported native mail-merge operations explicitly.

    Other native errors propagate without retrying the operation.
    """
    try:
        yield document.MailMerge
    except Exception as error:
        if isinstance(error, (AttributeError, NotImplementedError)) or getattr(
            error, "hresult", None
        ) in {-2147467263, -2147352573, -2147352570}:
            raise WpsError(
                "This WPS build does not implement the requested native "
                "mail-merge operation. "
                "Inspect the document before retrying; a native write may "
                "have partially completed."
            ) from error
        raise


def source_path(source: NativeObject) -> Path:
    """Resolve a local CSV source from native connection metadata.

    Parse unnamed OLE DB or ODBC connections without evaluating SQL or
    persisting state.
    """
    name = str(source.Name)
    if name:
        path = local_path(name, exists=True)
    else:
        connection = str(source.ConnectString)
        properties: dict[str, str] = {}
        position = 0
        while position < len(connection) and connection[position:].strip():
            match = CONNECTION_PROPERTY.match(connection, position)
            if match is None:
                raise ValueError(
                    "WPS returned an unsupported mail-merge connection string"
                )
            key = match[1].strip().casefold()
            if not key:
                raise ValueError(
                    "WPS returned an empty mail-merge connection property"
                )
            if key in properties:
                raise ValueError(
                    "WPS returned duplicate mail-merge connection properties"
                )
            value = next(
                item for item in match.groups()[1:] if item is not None
            )
            if match[5] is not None:
                value = value.strip()
                if value.startswith(("'", '"', "{")):
                    raise ValueError(
                        "WPS returned an unterminated quoted connection value"
                    )
            elif match[4] is not None:
                value = value.replace("}}", "}")
            else:
                value = (
                    value.replace('""', '"')
                    if match[2] is not None
                    else value.replace("''", "'")
                )
            properties[key] = value
            position = match.end()
        if "data source" in properties and "defaultdir" in properties:
            raise ValueError(
                "WPS returned ambiguous mail-merge source locations"
            )
        path = local_path(
            properties.get("data source", properties.get("defaultdir", ""))
        )
        if "defaultdir" in properties and not path.is_dir():
            raise ValueError(
                "ODBC mail merge requires an existing local source directory"
            )
        if path.is_dir():
            query = SOURCE_QUERY.fullmatch(str(source.QueryString))
            if query is None:
                raise ValueError(
                    "CSV mail merge requires a single SELECT * FROM "
                    "table query"
                )
            table = next(item for item in query.groups() if item is not None)
            if (
                not table.strip()
                or table in {".", ".."}
                or any(character in table for character in '/\\:*?[]<>|"`\x00')
                or any(ord(character) < 32 for character in table)
            ):
                raise ValueError(
                    "CSV query table must identify a file in its "
                    "local source directory"
                )
            candidate = local_path(
                str(
                    path
                    / (
                        table
                        if table.lower().endswith(".csv")
                        else table + ".csv"
                    )
                ),
                exists=True,
            )
            if candidate.parent != path:
                raise ValueError(
                    "CSV query table must stay within its local "
                    "source directory"
                )
            path = candidate
        else:
            path = local_path(str(path), exists=True)
    if path.suffix.lower() != ".csv":
        raise ValueError("Mail merge supports only local .csv data sources")
    return path


def source_details(source: NativeObject) -> JsonObject:
    """Read merge source schema without exposing connection secrets.

    Preserve the active record and expose only schema and record bounds.
    """
    fields = source.FieldNames
    return {
        "file_path": str(source_path(source)),
        "record_count": int(source.RecordCount),
        "first_record": int(source.FirstRecord),
        "last_record": int(source.LastRecord),
        "fields": [
            str(fields.Item(index).Name)
            for index in range(1, int(fields.Count) + 1)
        ],
    }


@tool(
    "wps_writer_mail_merge",
    operation="inspect",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def inspect_merge(params: DocumentInput) -> JsonObject:
    """Inspect native merge state, fields, and source schema.

    Connected data source record values are omitted.
    """
    document = get_session("writer").get_document(params.document)
    with native_mail_merge(document) as merge:
        state = int(merge.State)
        fields = []
        for index in range(1, int(merge.Fields.Count) + 1):
            field = merge.Fields.Item(index)
            code = field.Code
            fields.append(
                {
                    "index": index,
                    "type": int(field.Type),
                    "code": str(code.Text),
                    "start": int(code.Start),
                    "end": int(code.End),
                }
            )
        return {
            **describe_document(document),
            "state": MERGE_STATES.get(state, "unknown"),
            "native_state": state,
            "main_document_type": int(merge.MainDocumentType),
            "fields": fields,
            "source": source_details(merge.DataSource)
            if state in CONNECTED_STATES
            else None,
        }


@tool(
    "wps_writer_mail_merge",
    operation="field",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def insert_merge_field(params: MergeFieldInput) -> JsonObject:
    """Insert one validated native merge field at a UTF-16 position.

    The position is zero-based; existing text is not replaced.
    """
    document = get_session("writer").get_document(params.document)
    text_range = resolve_range(document, params.position, params.position)
    with native_mail_merge(document) as merge:
        if int(merge.State) in CONNECTED_STATES:
            names = source_details(merge.DataSource)["fields"]
            if params.name not in names:
                raise ValueError(
                    "Merge field is absent from the connected data source: "
                    f"{params.name}"
                )
        field = merge.Fields.Add(text_range, params.name)
        return {
            "document": str(document.FullName),
            "name": params.name,
            "code": str(field.Code.Text),
            "start": int(field.Code.Start),
            "end": int(field.Code.End),
        }


@tool(
    "wps_writer_mail_merge",
    operation="connect",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def connect_merge_source(params: MergeSourceInput) -> JsonObject:
    """Connect an existing local CSV as a read-only form-letter source.

    The source remains attached until disconnect; keep it available on
    disk. A current source must be disconnected first. Arbitrary SQL and
    connection strings are not accepted. Export spreadsheet data to CSV
    first; native XLSX selection cannot run unattended.
    """
    path = local_path(params.file_path, exists=True)
    session = get_session("writer")
    document = session.get_document(params.document)
    with native_mail_merge(document) as merge:
        if int(merge.State) in CONNECTED_STATES:
            raise ValueError(
                "A mail-merge source is already connected; disconnect it first"
            )
        with temporary_attributes(session.get_application(), DisplayAlerts=0):
            merge.MainDocumentType = 0
            merge.OpenDataSource(
                Name=str(path),
                ConfirmConversions=False,
                ReadOnly=True,
                LinkToSource=True,
                AddToRecentFiles=False,
                OpenExclusive=False,
            )
        if int(merge.State) not in CONNECTED_STATES:
            raise WpsError(
                "WPS did not attach the requested mail-merge data source"
            )
        source = source_details(merge.DataSource)
        if document_key(source["file_path"]) != document_key(path):
            raise WpsError(
                "WPS attached a different mail-merge source; "
                "inspect before retrying"
            )
        if not source["fields"]:
            raise WpsError(
                "The connected mail-merge source has no named columns"
            )
        return {"document": str(document.FullName), "source": source}


@tool(
    "wps_writer_mail_merge",
    operation="execute",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def execute_merge(params: MergeExecuteInput) -> JsonObject:
    """Merge a record interval into one new, unsaved Writer document.

    Output is restricted to a new document; no email or print destination is
    used.

    Omitted last_record selects all remaining records. Save the returned
    full_name with wps_documents/save. Native failures can partially create
    output; never replay blindly.
    """
    session = get_session("writer")
    document = session.get_document(params.document)
    with native_mail_merge(document) as merge:
        if int(merge.State) not in CONNECTED_STATES:
            raise ValueError(
                "Connect a local mail-merge data source before executing"
            )
        if int(merge.MainDocumentType) not in {0, 3}:
            raise ValueError(
                "Only form-letter and directory documents support "
                "this merge output"
            )
        source = merge.DataSource
        source_path(source)
        count = int(source.RecordCount)
        if count == 0 or (
            count > 0
            and max(
                params.first_record, params.last_record or params.first_record
            )
            > count
        ):
            raise ValueError(
                f"Record interval exceeds the data source's {count} records"
            )
        before = {
            document_key(item.FullName) for item in session.get_documents()
        }
        token = CALL_DOCUMENT.set(str(document.FullName))
        try:
            with (
                session.command_window(),
                temporary_attributes(
                    session.get_application(), DisplayAlerts=0
                ),
                temporary_attributes(
                    merge,
                    Destination=0,
                    SuppressBlankLines=params.suppress_blank_lines,
                ),
                temporary_attributes(
                    source,
                    FirstRecord=params.first_record,
                    LastRecord=params.last_record or -16,
                ),
            ):
                merge.Execute(Pause=False)
                created = [
                    item
                    for item in session.get_documents()
                    if document_key(item.FullName) not in before
                ]
                if len(created) != 1:
                    raise WpsError(
                        "Expected one new merged document; "
                        f"found {len(created)}. "
                        "Inspect open documents before retrying."
                    )
                result = describe_document(created[0])
        finally:
            CALL_DOCUMENT.reset(token)
        return {
            **result,
            "source_document": str(document.FullName),
            "first_record": params.first_record,
            "last_record": params.last_record
            or (count if count > 0 else None),
        }


@tool(
    "wps_writer_mail_merge",
    operation="disconnect",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def disconnect_merge_source(params: DocumentInput) -> JsonObject:
    """Detach the merge data source while retaining template fields."""
    document = get_session("writer").get_document(params.document)
    with native_mail_merge(document) as merge:
        connected = int(merge.State) in CONNECTED_STATES
        if connected:
            merge.DataSource.Close()
            if int(merge.State) in CONNECTED_STATES:
                raise WpsError(
                    "WPS still reports an attached mail-merge source "
                    "after Close"
                )
        return {"document": str(document.FullName), "disconnected": connected}
