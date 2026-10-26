"""Read, search, edit, and format native Writer text ranges."""

from typing import Annotated, Self

from pydantic import Field, model_validator

from wps_mcp.models import DocumentInput, JsonObject, NativeObject
from wps_mcp.native.runtime import get_session
from wps_mcp.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)
from wps_mcp.writer.access import range_details, resolve_range
from wps_mcp.writer.formatting import apply_character, apply_paragraph
from wps_mcp.writer.models import (
    CharacterFormat,
    ExplicitRangeInput,
    Index,
    ParagraphFormat,
    Position,
    RangeInput,
)


class ParagraphInput(DocumentInput):
    """Select a bounded page of one-based Writer paragraphs."""

    start_index: Index = 1
    count: Annotated[int, Field(ge=1, le=1000, strict=True)] = 100


class InsertTextInput(DocumentInput):
    """Insert literal text at a native UTF-16 position."""

    position: Position
    text: str = Field(min_length=1)


class EditTextInput(ExplicitRangeInput):
    """Replace an explicit native UTF-16 text interval."""

    text: str


class DeleteTextInput(ExplicitRangeInput):
    """Delete an explicit native UTF-16 text interval."""


class FindTextInput(RangeInput):
    """Bound a native text search and its maximum match count."""

    text: str = Field(min_length=1)
    match_case: bool = False
    whole_word: bool = False
    max_results: Annotated[int, Field(ge=1, le=10000, strict=True)] = 100


class ReplaceTextInput(FindTextInput):
    """Replace bounded native matches with literal text."""

    replacement: str
    replace_all: bool = True


class FormatTextInput(RangeInput):
    """Apply character, paragraph, or named-style formatting."""

    character: CharacterFormat | None = None
    paragraph: ParagraphFormat | None = None
    style_name: str | None = Field(default=None, min_length=1)

    @model_validator(mode="after")
    def validate_format(self) -> Self:
        """Require at least one explicit formatting property."""
        character_values = (
            self.character.model_dump(exclude_none=True)
            if self.character
            else {}
        )
        paragraph_values = (
            self.paragraph.model_dump(exclude_none=True)
            if self.paragraph
            else {}
        )
        if (
            not character_values
            and (not paragraph_values)
            and (self.style_name is None)
        ):
            raise ValueError(
                "Specify character, paragraph or style_name formatting"
            )
        return self


@tool(
    "wps_writer_text",
    operation="read",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def read_text(params: RangeInput) -> JsonObject:
    """Read main-story text using native UTF-16 range boundaries.

    Positions are zero-based UTF-16 units and end is exclusive. An omitted end
    reads to the document end. WPS paragraph and control markers are retained.
    """
    document = get_session("writer").get_document(params.document)
    text_range = resolve_range(document, params.start, params.end)
    return {
        "document": str(document.Name),
        **range_details(text_range),
        "position_unit": "UTF-16",
        "end_exclusive": True,
    }


@tool(
    "wps_writer_text",
    operation="paragraphs",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def read_paragraphs(params: ParagraphInput) -> JsonObject:
    """Read paragraphs with native range positions.

    Collection indices start at one.
    """
    document = get_session("writer").get_document(params.document)
    paragraphs = document.Paragraphs
    count = int(paragraphs.Count)
    if params.start_index > count:
        raise ValueError(f"start_index exceeds paragraph count {count}")
    end_index = min(count, params.start_index + params.count - 1)
    result = [
        {"index": index, **range_details(paragraphs.Item(index).Range)}
        for index in range(params.start_index, end_index + 1)
    ]
    return {
        "paragraphs": result,
        "total": count,
        "has_more": end_index < count,
    }


@tool(
    "wps_writer_text",
    operation="insert",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def insert_text(params: InsertTextInput) -> JsonObject:
    """Insert literal text at a collapsed native UTF-16 position.

    Use a carriage return for a paragraph break.
    """
    document = get_session("writer").get_document(params.document)
    text_range = resolve_range(document, params.position, params.position)
    text_range.Text = params.text
    return range_details(text_range)


@tool(
    "wps_writer_text",
    operation="replace",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def edit_text(params: EditTextInput) -> JsonObject:
    """Replace a native text interval while preserving whitespace.

    The selected interval uses an inclusive start and exclusive end.
    """
    document = get_session("writer").get_document(params.document)
    text_range = resolve_range(document, params.start, params.end)
    text_range.Text = params.text
    return range_details(text_range)


@tool(
    "wps_writer_text",
    operation="delete",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def delete_text(params: DeleteTextInput) -> JsonObject:
    """Delete a native [start, end) text range.

    Offsets after the deletion change.
    """
    document = get_session("writer").get_document(params.document)
    text_range = resolve_range(document, params.start, params.end)
    if params.start == params.end:
        return {"start": params.start, "end": params.end, "deleted": False}
    text_range.Delete()
    return {"start": params.start, "end": params.end, "deleted": True}


def find_matches(
    document: NativeObject, params: FindTextInput, limit: int
) -> list[JsonObject]:
    """Collect advancing native matches within the validated interval."""
    search_range = resolve_range(document, params.start, params.end)
    search_end = int(search_range.End)
    cursor = params.start
    matches: list[JsonObject] = []
    while cursor < search_end and len(matches) < limit:
        text_range = document.Range(cursor, search_end)
        finder = text_range.Find
        finder.ClearFormatting()
        found = finder.Execute(
            FindText=params.text,
            MatchCase=params.match_case,
            MatchWholeWord=params.whole_word,
            MatchWildcards=False,
            MatchSoundsLike=False,
            MatchAllWordForms=False,
            Forward=True,
            Wrap=0,
            Format=False,
        )
        if not found:
            break
        match = range_details(text_range)
        if (
            match["end"] <= cursor
            or match["start"] < cursor
            or match["end"] > search_end
        ):
            raise RuntimeError(
                "WPS returned an invalid or non-advancing search range"
            )
        matches.append(match)
        cursor = match["end"]
    return matches


@tool(
    "wps_writer_text",
    operation="find",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def find_text(params: FindTextInput) -> JsonObject:
    """Find bounded native text matches without wrapping or wildcards.

    WPS search semantics, including native ^ special codes, apply.
    """
    document = get_session("writer").get_document(params.document)
    matches = find_matches(document, params, params.max_results + 1)
    return {
        "matches": matches[: params.max_results],
        "has_more": len(matches) > params.max_results,
    }


@tool(
    "wps_writer_text",
    operation="replace_matches",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def replace_text(params: ReplaceTextInput) -> JsonObject:
    """Find with WPS search and replace with literal text.

    All matches are checked before editing; exceeding max_results aborts
    unchanged.
    """
    document = get_session("writer").get_document(params.document)
    limit = params.max_results + 1 if params.replace_all else 1
    matches = find_matches(document, params, limit)
    if params.replace_all and len(matches) > params.max_results:
        raise ValueError(
            "Matches exceed max_results; narrow the range or "
            "raise max_results before replacing"
        )
    text_ranges = [
        resolve_range(document, match["start"], match["end"])
        for match in reversed(matches)
    ]
    for text_range in text_ranges:
        text_range.Text = params.replacement
    return {"replaced_count": len(matches), "original_ranges": matches}


@tool(
    "wps_writer_text",
    operation="format",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def format_text(params: FormatTextInput) -> JsonObject:
    """Apply native character, paragraph, or named-style formatting.

    Dimensions use points; style names use the document locale.
    """
    document = get_session("writer").get_document(params.document)
    text_range = resolve_range(document, params.start, params.end)
    style = (
        document.Styles.Item(params.style_name)
        if params.style_name is not None
        else None
    )
    if style is not None:
        text_range.Style = style
    apply_character(text_range, params.character)
    apply_paragraph(text_range, params.paragraph)
    return {
        "start": int(text_range.Start),
        "end": int(text_range.End),
        "format": params.model_dump(
            exclude={"document", "start", "end"}, exclude_none=True
        ),
    }
