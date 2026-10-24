"""Manage Writer contents, notes, and native cross-references."""

import logging
import sys
from typing import Annotated, Literal, Self

from pydantic import Field, model_validator

from src.errors import WpsError
from src.models import DocumentInput, JsonObject, NativeObject
from src.native.context import CALL_DOCUMENT
from src.native.runtime import get_session
from src.native.selection import document_key
from src.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)
from src.writer.access import range_details, resolve_item, resolve_range
from src.writer.models import Index, Position

ReferenceType = Literal["bookmark", "heading", "footnote", "endnote"]


ReferenceKind = Literal[
    "text",
    "page_number",
    "position",
    "number",
    "number_formatted",
    "number_full_context",
    "number_no_context",
    "number_relative_context",
]


REFERENCE_TYPES = {"bookmark": 2, "heading": 1, "footnote": 3, "endnote": 4}


REFERENCE_KINDS = {
    "text": -1,
    "page_number": 7,
    "position": 15,
    "number_full_context": -4,
    "number_no_context": -3,
    "number_relative_context": -2,
}


REFERENCE_FIELDS = {
    3: "ref",
    5: "footnote_ref",
    37: "page_ref",
    72: "note_ref",
}


LOGGER = logging.getLogger(__name__)


class AddTocInput(DocumentInput):
    """Create a contents table from validated heading levels."""

    position: Position
    upper_heading_level: Annotated[int, Field(ge=1, le=9)] = 1
    lower_heading_level: Annotated[int, Field(ge=1, le=9)] = 3
    include_page_numbers: bool = True
    right_align_page_numbers: bool = True
    use_hyperlinks: bool = True
    use_outline_levels: bool = False

    @model_validator(mode="after")
    def validate_heading_levels(self) -> Self:
        """Require ascending bounds for included heading levels."""
        if self.lower_heading_level < self.upper_heading_level:
            raise ValueError(
                "lower_heading_level must be at least upper_heading_level"
            )
        return self


class TocInput(DocumentInput):
    """Select an existing contents table by its one-based index."""

    toc_index: Index


class UpdateTocInput(TocInput):
    """Refresh all contents entries or only their page numbers."""

    mode: Literal["all", "page_numbers"] = "all"


class NoteKindInput(DocumentInput):
    """Select the footnote or endnote collection."""

    kind: Literal["footnote", "endnote"]


class AddNoteInput(NoteKindInput):
    """Insert note text at an explicit main-story position."""

    position: Position
    text: str = Field(min_length=1)
    reference_mark: str | None = Field(default=None, min_length=1)


class NoteInput(NoteKindInput):
    """Select one note by kind and one-based collection index."""

    note_index: Index


class UpdateNoteInput(NoteInput):
    """Replace literal note text while preserving its reference."""

    text: str


class ReferenceTargetsInput(DocumentInput):
    """Select one supported native cross-reference target kind."""

    reference_type: ReferenceType


class InsertReferenceInput(ReferenceTargetsInput):
    """Insert a supported reference to an existing native target."""

    position: Position
    reference_kind: ReferenceKind
    reference_item: Index | Annotated[str, Field(min_length=1)]
    insert_as_hyperlink: bool = True
    include_position: bool = False

    @model_validator(mode="after")
    def validate_reference(self) -> Self:
        """Reject unsupported target, item, and reference-kind combinations."""
        if self.reference_type == "bookmark":
            if not isinstance(self.reference_item, str):
                raise ValueError("A bookmark reference_item must be its name")
        elif not isinstance(self.reference_item, int):
            raise ValueError(
                "reference_item must be a one-based index from targets"
            )
        allowed = {"page_number", "position"}
        if self.reference_type in {"bookmark", "heading"}:
            allowed.add("text")
            if self.reference_type == "heading":
                allowed.update(
                    {
                        "number_full_context",
                        "number_no_context",
                        "number_relative_context",
                    }
                )
        else:
            allowed.update({"number", "number_formatted"})
        if self.reference_kind not in allowed:
            raise ValueError(
                f"reference_kind for {self.reference_type} "
                f"must be one of {sorted(allowed)}"
            )
        return self


class ReferenceFieldInput(DocumentInput):
    """Identify a supported field by its document collection index."""

    field_index: Index = Field(
        description="One-based document field index returned by list"
    )


def toc_details(toc: NativeObject) -> JsonObject:
    """Describe the native TOC without borrowing its proxies across the STA."""
    return {
        **range_details(toc.Range),
        "upper_heading_level": int(toc.UpperHeadingLevel),
        "lower_heading_level": int(toc.LowerHeadingLevel),
        "include_page_numbers": bool(toc.IncludePageNumbers),
        "right_align_page_numbers": bool(toc.RightAlignPageNumbers),
        "use_hyperlinks": bool(toc.UseHyperlinks),
    }


def note_details(note: NativeObject, index: int) -> JsonObject:
    """Describe a note anchor and its separate text story."""
    return {
        "note_index": index,
        "reference": range_details(note.Reference),
        "content": range_details(note.Range),
    }


def reference_targets(
    document: NativeObject, reference_type: ReferenceType
) -> list[str]:
    """Preserve native cross-reference target order and indices.

    Non-bookmark reference_item values follow this native order.
    """
    values = document.GetCrossReferenceItems(REFERENCE_TYPES[reference_type])
    if values is None:
        return (
            _collection_reference_targets(document, reference_type)
            if sys.platform == "linux"
            else []
        )
    if not isinstance(values, (tuple, list)) or any(
        not isinstance(item, str) for item in values
    ):
        raise RuntimeError(
            "WPS returned an invalid cross-reference target array"
        )
    return list(values)


def _collection_reference_targets(
    document: NativeObject, reference_type: ReferenceType
) -> list[str]:
    """List targets when Linux WPS omits its cross-reference item array.

    Native heading reference indices count built-in heading styles in
    document order. An outline level on another style is not a heading
    target. Note indices follow their respective native collections.
    """
    if reference_type == "heading":
        headings = {
            str(document.Styles.Item(identifier).NameLocal)
            for identifier in range(-2, -11, -1)
        }
        paragraphs = document.Paragraphs
        result = []
        for index in range(1, int(paragraphs.Count) + 1):
            paragraph = paragraphs.Item(index)
            text_range = paragraph.Range
            if str(text_range.Style.NameLocal) in headings:
                result.append(str(text_range.Text).rstrip("\r"))
        return result
    name = {
        "bookmark": "Bookmarks",
        "footnote": "Footnotes",
        "endnote": "Endnotes",
    }[reference_type]
    collection = getattr(document, name)
    result = []
    for index in range(1, int(collection.Count) + 1):
        item = collection.Item(index)
        result.append(
            str(item.Name)
            if reference_type == "bookmark"
            else str(item.Range.Text).rstrip("\r")
        )
    return result


def reference_details(field: NativeObject, index: int) -> JsonObject:
    """Describe only supported native cross-reference fields."""
    field_type = int(field.Type)
    if field_type not in REFERENCE_FIELDS:
        raise ValueError(
            f"Field {index} is not a supported cross-reference field"
        )
    return {
        "field_index": index,
        "field_type": REFERENCE_FIELDS[field_type],
        "code": str(field.Code.Text),
        "result": range_details(field.Result),
        "locked": bool(field.Locked),
    }


@tool(
    "wps_writer_toc",
    operation="list",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def list_tocs(params: DocumentInput) -> JsonObject:
    """List contents tables with one-based indices and native positions."""
    document = get_session("writer").get_document(params.document)
    tocs = document.TablesOfContents
    return {
        "tables_of_contents": [
            {"toc_index": index, **toc_details(tocs.Item(index))}
            for index in range(1, int(tocs.Count) + 1)
        ]
    }


@tool(
    "wps_writer_toc",
    operation="add",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def add_toc(params: AddTocInput) -> JsonObject:
    """Insert a contents table from supported native heading levels.

    The insertion range is collapsed and heading levels are limited to 1..9.
    """
    document = get_session("writer").get_document(params.document)
    text_range = resolve_range(document, params.position, params.position)
    toc = document.TablesOfContents.Add(
        Range=text_range,
        UseHeadingStyles=True,
        UpperHeadingLevel=params.upper_heading_level,
        LowerHeadingLevel=params.lower_heading_level,
        UseFields=False,
        RightAlignPageNumbers=params.right_align_page_numbers,
        IncludePageNumbers=params.include_page_numbers,
        UseHyperlinks=params.use_hyperlinks,
        UseOutlineLevels=params.use_outline_levels,
    )
    return toc_details(toc)


@tool(
    "wps_writer_toc",
    operation="update",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def update_toc(params: UpdateTocInput) -> JsonObject:
    """Refresh one TOC's entries and pages, or only its page numbers."""
    document = get_session("writer").get_document(params.document)
    toc = resolve_item(
        document.TablesOfContents, params.toc_index, "Table of contents"
    )
    if params.mode == "page_numbers":
        toc.UpdatePageNumbers()
    else:
        toc.Update()
    return {
        "toc_index": params.toc_index,
        "mode": params.mode,
        **toc_details(toc),
    }


@tool(
    "wps_writer_toc",
    operation="delete",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def delete_toc(params: TocInput) -> JsonObject:
    """Delete only the specified TOC, including its generated content."""
    document = get_session("writer").get_document(params.document)
    toc = resolve_item(
        document.TablesOfContents, params.toc_index, "Table of contents"
    )
    toc.Delete()
    return {"toc_index": params.toc_index, "deleted": True}


@tool(
    "wps_writer_notes",
    operation="list",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def list_notes(params: NoteKindInput) -> JsonObject:
    """List one native note kind with separate anchor and text offsets."""
    document = get_session("writer").get_document(params.document)
    notes = (
        document.Footnotes if params.kind == "footnote" else document.Endnotes
    )
    return {
        "kind": params.kind,
        "notes": [
            note_details(notes.Item(index), index)
            for index in range(1, int(notes.Count) + 1)
        ],
    }


@tool(
    "wps_writer_notes",
    operation="add",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def add_note(params: AddNoteInput) -> JsonObject:
    """Insert a footnote or endnote at a main-story position.

    Native numbering is used unless reference_mark is supplied.
    """
    document = get_session("writer").get_document(params.document)
    text_range = resolve_range(document, params.position, params.position)
    notes = (
        document.Footnotes if params.kind == "footnote" else document.Endnotes
    )
    arguments: JsonObject = {"Range": text_range, "Text": params.text}
    if params.reference_mark is not None:
        arguments["Reference"] = params.reference_mark
    note = notes.Add(**arguments)
    return {"kind": params.kind, **note_details(note, int(note.Index))}


@tool(
    "wps_writer_notes",
    operation="update",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def update_note(params: UpdateNoteInput) -> JsonObject:
    """Replace a note's literal text without replacing its reference mark."""
    document = get_session("writer").get_document(params.document)
    notes = (
        document.Footnotes if params.kind == "footnote" else document.Endnotes
    )
    note = resolve_item(notes, params.note_index, params.kind.capitalize())
    note.Range.Text = params.text
    return {"kind": params.kind, **note_details(note, params.note_index)}


@tool(
    "wps_writer_notes",
    operation="delete",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def delete_note(params: NoteInput) -> JsonObject:
    """Delete one note and its corresponding main-story reference.

    WPS renumbers the remaining notes.
    """
    document = get_session("writer").get_document(params.document)
    notes = (
        document.Footnotes if params.kind == "footnote" else document.Endnotes
    )
    note = resolve_item(notes, params.note_index, params.kind.capitalize())
    note.Delete()
    return {
        "kind": params.kind,
        "note_index": params.note_index,
        "deleted": True,
    }


@tool(
    "wps_writer_cross_references",
    operation="targets",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def list_reference_targets(params: ReferenceTargetsInput) -> JsonObject:
    """List valid bookmark names or native heading and note indices."""
    document = get_session("writer").get_document(params.document)
    targets = reference_targets(document, params.reference_type)
    return {
        "reference_type": params.reference_type,
        "targets": [
            {
                "reference_item": text
                if params.reference_type == "bookmark"
                else index,
                "text": text,
            }
            for index, text in enumerate(targets, start=1)
        ],
    }


@tool(
    "wps_writer_cross_references",
    operation="insert",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def insert_reference(params: InsertReferenceInput) -> JsonObject:
    """Insert and verify a native reference while restoring selection.

    Supported targets are bookmarks, headings, footnotes, and endnotes.

    Use targets immediately before inserting: non-bookmark indices follow WPS's
    cross-reference dialog order. Preserve the target selection and previous
    window; return field_index only after confirming a new reference at the
    requested position.
    """
    session = get_session("writer")
    document = session.get_document(params.document)
    text_range = resolve_range(document, params.position, params.position)
    targets = reference_targets(document, params.reference_type)
    if isinstance(params.reference_item, str):
        if params.reference_item not in targets:
            raise ValueError(
                f"Bookmark target not found: {params.reference_item}"
            )
    elif params.reference_item > len(targets):
        raise ValueError(
            f"reference_item must be between 1 and {len(targets)}"
        )
    if params.reference_kind in {"number", "number_formatted"}:
        reference_kind = 5 if params.reference_type == "footnote" else 6
        if params.reference_kind == "number_formatted":
            reference_kind += 11
    else:
        reference_kind = REFERENCE_KINDS[params.reference_kind]
    if params.reference_kind == "page_number":
        document.Repaginate()
    fields = document.Fields
    previous_count = int(fields.Count)
    token = CALL_DOCUMENT.set(str(document.FullName))
    try:
        with session.command_window():
            selection = session.get_application().Selection
            if document_key(selection.Document.FullName) != document_key(
                document.FullName
            ):
                raise WpsError("Could not activate the target Writer document")
            previous_selection = selection.Range.Duplicate
            failed = False
            try:
                text_range.Select()
                # WPS can silently ignore the equivalent Range method.
                selection.InsertCrossReference(
                    REFERENCE_TYPES[params.reference_type],
                    reference_kind,
                    params.reference_item,
                    params.insert_as_hyperlink,
                    params.include_position,
                    False,
                    " ",
                )
                count = int(fields.Count)
                if count > previous_count:
                    for index in range(1, count + 1):
                        field = fields.Item(index)
                        if (
                            int(field.Type) in REFERENCE_FIELDS
                            and int(field.Code.Start) == params.position + 1
                        ):
                            return {
                                "reference_type": params.reference_type,
                                "reference_kind": params.reference_kind,
                                "reference_item": params.reference_item,
                                "inserted": True,
                                "field_index": index,
                                **range_details(field.Result),
                            }
                raise WpsError(
                    "WPS did not create a reference field at the requested "
                    "position; "
                    "inspect the document before retrying"
                )
            except BaseException:
                failed = True
                raise
            finally:
                try:
                    previous_selection.Select()
                except Exception as error:
                    if failed:
                        LOGGER.error(
                            "Could not restore the Writer selection: %s", error
                        )
                    else:
                        raise WpsError(
                            "Could not restore the Writer selection"
                        ) from error
    finally:
        CALL_DOCUMENT.reset(token)


@tool(
    "wps_writer_cross_references",
    operation="list",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def list_references(params: DocumentInput) -> JsonObject:
    """List supported reference fields in the main document story.

    Include REF, PAGEREF, and note-reference fields.
    """
    document = get_session("writer").get_document(params.document)
    fields = document.Fields
    result: list[JsonObject] = []
    for index in range(1, int(fields.Count) + 1):
        field = fields.Item(index)
        if int(field.Type) in REFERENCE_FIELDS:
            result.append(reference_details(field, index))
    return {"references": result, "story": "main"}


@tool(
    "wps_writer_cross_references",
    operation="update",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def update_reference(params: ReferenceFieldInput) -> JsonObject:
    """Refresh one unlocked and supported native reference field."""
    document = get_session("writer").get_document(params.document)
    field = resolve_item(document.Fields, params.field_index, "Field")
    detail = reference_details(field, params.field_index)
    if detail["locked"]:
        raise ValueError(f"Field {params.field_index} is locked")
    if int(field.Type) == 37:
        document.Repaginate()
    if not field.Update():
        raise RuntimeError(
            f"WPS could not update reference field {params.field_index}; "
            "inspect its target"
        )
    return {"updated": True, **reference_details(field, params.field_index)}


@tool(
    "wps_writer_cross_references",
    operation="delete",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def delete_reference(params: ReferenceFieldInput) -> JsonObject:
    """Delete a reference field without deleting its target.

    The field result text is removed with the field.
    """
    document = get_session("writer").get_document(params.document)
    field = resolve_item(document.Fields, params.field_index, "Field")
    reference_details(field, params.field_index)
    field.Delete()
    return {"field_index": params.field_index, "deleted": True}
