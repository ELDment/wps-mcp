"""Inspect and manage native Writer styles and inheritance."""

from typing import Annotated, Literal, Self

from pydantic import Field, model_validator

from src.models import DocumentInput, JsonObject, NativeObject
from src.native.runtime import get_session
from src.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)
from src.writer.access import resolve_range
from src.writer.formatting import apply_character, apply_paragraph
from src.writer.models import CharacterFormat, ParagraphFormat, RangeInput

STYLE_TYPES = {
    "paragraph": 1,
    "character": 2,
    "table": 3,
    "list": 4,
    "paragraph_only": 5,
    "linked": 6,
}


PARAGRAPH_STYLE_TYPES = {1, 5, 6}


class ListStylesInput(DocumentInput):
    """Filter native document styles by kind, origin, or usage."""

    kind: (
        Literal[
            "paragraph",
            "character",
            "table",
            "list",
            "paragraph_only",
            "linked",
        ]
        | None
    ) = None
    built_in: bool | None = None
    in_use: bool | None = None


class StyleInput(DocumentInput):
    """Identify a document style by its local native name."""

    name: str = Field(min_length=1)


class ReadStyleInput(DocumentInput):
    """Read a style by local name or a negative built-in style identifier."""

    name: Annotated[str, Field(min_length=1)] | Annotated[int, Field(lt=0)]


class StyleSettingsInput(StyleInput):
    """Describe optional style inheritance and formatting changes."""

    base_style: str | None = Field(default=None, min_length=1)
    next_style: str | None = Field(default=None, min_length=1)
    automatically_update: bool | None = None
    character: CharacterFormat | None = None
    paragraph: ParagraphFormat | None = None


class CreateStyleInput(StyleSettingsInput):
    """Create a paragraph or character style with explicit settings."""

    kind: Literal["paragraph", "character"] = "paragraph"


class UpdateStyleInput(StyleSettingsInput):
    """Require at least one change to an existing style."""

    @model_validator(mode="after")
    def validate_changes(self) -> Self:
        """Reject updates that do not change any supported property."""
        changes = self.model_dump(
            exclude={"document", "name"}, exclude_none=True
        )
        if not any(value != {} for value in changes.values()):
            raise ValueError("Specify at least one style setting")
        return self


class ApplyStyleInput(RangeInput):
    """Apply a named style to a native document interval."""

    name: str = Field(min_length=1)


def style_parent(document: NativeObject, style: NativeObject) -> NativeObject:
    """Resolve style inheritance without accessing native root parents.

    Built-in root IDs avoid reading an unavailable BaseStyle property.
    """
    if bool(style.BuiltIn):
        kind = int(style.Type)
        root_id = (
            -66 if kind == 2 else -1 if kind in PARAGRAPH_STYLE_TYPES else None
        )
        if root_id is not None:
            root = document.Styles.Item(root_id)
            if (
                str(style.NameLocal).casefold()
                == str(root.NameLocal).casefold()
            ):
                return None
    return style.BaseStyle


def style_details(
    document: NativeObject, style: NativeObject, *, formatting: bool = False
) -> JsonObject:
    """Describe native styles while preserving mixed-format sentinels."""
    kind = int(style.Type)
    result: JsonObject = {
        "name": str(style.NameLocal),
        "kind": next(
            (name for name, value in STYLE_TYPES.items() if value == kind),
            str(kind),
        ),
        "built_in": bool(style.BuiltIn),
        "in_use": bool(style.InUse),
    }
    if formatting:
        base = style_parent(document, style)
        result["base_style"] = (
            base
            if base is None or isinstance(base, str)
            else str(base.NameLocal)
        )
        font = style.Font
        result["character"] = {
            "font_name": str(font.Name),
            "font_name_far_east": str(font.NameFarEast),
            "font_size": float(font.Size),
            "bold": int(font.Bold),
            "italic": int(font.Italic),
            "underline": int(font.Underline),
            "color": int(font.Color),
        }
        if kind in PARAGRAPH_STYLE_TYPES:
            next_style = style.NextParagraphStyle
            result["next_style"] = (
                next_style
                if next_style is None or isinstance(next_style, str)
                else str(next_style.NameLocal)
            )
            result["automatically_update"] = bool(style.AutomaticallyUpdate)
            paragraph = style.ParagraphFormat
            result["paragraph"] = {
                "alignment": int(paragraph.Alignment),
                "left_indent": float(paragraph.LeftIndent),
                "right_indent": float(paragraph.RightIndent),
                "first_line_indent": float(paragraph.FirstLineIndent),
                "space_before": float(paragraph.SpaceBefore),
                "space_after": float(paragraph.SpaceAfter),
                "line_spacing": float(paragraph.LineSpacing),
                "line_spacing_rule": int(paragraph.LineSpacingRule),
                "keep_with_next": int(paragraph.KeepWithNext),
                "keep_together": int(paragraph.KeepTogether),
                "page_break_before": int(paragraph.PageBreakBefore),
            }
    return result


def prepare_style_settings(
    document: NativeObject,
    params: StyleSettingsInput,
    kind: int,
    target_name: str,
) -> JsonObject:
    """Resolve compatible style inheritance before applying changes."""
    if kind not in PARAGRAPH_STYLE_TYPES | {2}:
        raise ValueError(
            "Only paragraph, linked and character styles "
            "support this operation"
        )
    if kind == 2 and (
        params.paragraph is not None
        or params.next_style is not None
        or params.automatically_update is not None
    ):
        raise ValueError("Character styles cannot define paragraph settings")
    changes: JsonObject = {}
    if params.base_style is not None:
        base = document.Styles.Item(params.base_style)
        base_kind = int(base.Type)
        if (base_kind == 2) != (
            kind == 2
        ) or base_kind not in PARAGRAPH_STYLE_TYPES | {2}:
            raise ValueError("Base style must have a compatible style type")
        changes["BaseStyle"] = str(base.NameLocal)
        seen = {target_name.casefold()}
        while True:
            name = str(base.NameLocal).casefold()
            if name in seen:
                raise ValueError("Style inheritance cannot contain a cycle")
            seen.add(name)
            parent = style_parent(document, base)
            if parent is None or parent == "":
                break
            base = (
                document.Styles.Item(parent)
                if isinstance(parent, str)
                else parent
            )
    if params.next_style is not None:
        if params.next_style.casefold() == target_name.casefold():
            changes["NextParagraphStyle"] = target_name
        else:
            next_style = document.Styles.Item(params.next_style)
            if int(next_style.Type) not in PARAGRAPH_STYLE_TYPES:
                raise ValueError("next_style must be a paragraph style")
            changes["NextParagraphStyle"] = str(next_style.NameLocal)
    if params.automatically_update is not None:
        changes["AutomaticallyUpdate"] = params.automatically_update
    return changes


@tool(
    "wps_writer_styles",
    operation="list",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def list_styles(params: ListStylesInput) -> JsonObject:
    """List native styles with optional kind, usage, and origin filters."""
    document = get_session("writer").get_document(params.document)
    styles = document.Styles
    result = []
    for index in range(1, int(styles.Count) + 1):
        details = style_details(document, styles.Item(index))
        if (
            (params.kind is None or details["kind"] == params.kind)
            and (
                params.built_in is None
                or details["built_in"] == params.built_in
            )
            and (params.in_use is None or details["in_use"] == params.in_use)
        ):
            result.append(details)
    return {"styles": result, "total": int(styles.Count)}


@tool(
    "wps_writer_styles",
    operation="read",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def read_style(params: ReadStyleInput) -> JsonObject:
    """Read formatting by local name or a negative built-in style ID."""
    document = get_session("writer").get_document(params.document)
    return style_details(
        document, document.Styles.Item(params.name), formatting=True
    )


@tool(
    "wps_writer_styles",
    operation="create",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def create_style(params: CreateStyleInput) -> JsonObject:
    """Create a style after validating all inheritance dependencies."""
    document = get_session("writer").get_document(params.document)
    styles = document.Styles
    for index in range(1, int(styles.Count) + 1):
        if (
            str(styles.Item(index).NameLocal).casefold()
            == params.name.casefold()
        ):
            raise ValueError("Style already exists; use update instead")
    changes = prepare_style_settings(
        document, params, STYLE_TYPES[params.kind], params.name
    )
    style = styles.Add(params.name, STYLE_TYPES[params.kind])
    for name, value in changes.items():
        setattr(style, name, value)
    apply_character(style, params.character)
    apply_paragraph(style, params.paragraph)
    return style_details(document, style)


@tool(
    "wps_writer_styles",
    operation="update",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def update_style(params: UpdateStyleInput) -> JsonObject:
    """Update native style inheritance and formatting for existing uses."""
    document = get_session("writer").get_document(params.document)
    style = document.Styles.Item(params.name)
    changes = prepare_style_settings(
        document, params, int(style.Type), str(style.NameLocal)
    )
    for name, value in changes.items():
        setattr(style, name, value)
    apply_character(style, params.character)
    apply_paragraph(style, params.paragraph)
    return style_details(document, style)


@tool(
    "wps_writer_styles",
    operation="apply",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def apply_style(params: ApplyStyleInput) -> JsonObject:
    """Apply a supported native style to a validated UTF-16 interval.

    Only paragraph, linked, and character styles apply to text ranges.
    """
    document = get_session("writer").get_document(params.document)
    text_range = resolve_range(document, params.start, params.end)
    style = document.Styles.Item(params.name)
    if int(style.Type) not in PARAGRAPH_STYLE_TYPES | {2}:
        raise ValueError(
            "Only paragraph, linked and character styles apply to a text range"
        )
    text_range.Style = style
    return {
        "name": str(style.NameLocal),
        "start": int(text_range.Start),
        "end": int(text_range.End),
    }


@tool(
    "wps_writer_styles",
    operation="delete",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def delete_style(params: StyleInput) -> JsonObject:
    """Delete a custom style using native replacement formatting."""
    document = get_session("writer").get_document(params.document)
    style = document.Styles.Item(params.name)
    if bool(style.BuiltIn):
        raise ValueError("Built-in styles cannot be deleted")
    name = str(style.NameLocal)
    style.Delete()
    return {"name": name, "deleted": True}
