"""Presentation selection, metadata and tags."""

from __future__ import annotations

import logging
from datetime import date, datetime
from typing import Literal, Self

from pydantic import Field, model_validator

from wps_mcp.errors import native_errors
from wps_mcp.models import DocumentInput, InputModel, JsonObject, NativeObject
from wps_mcp.native.context import CALL_DOCUMENT
from wps_mcp.native.discovery import template_directory
from wps_mcp.native.runtime import get_session
from wps_mcp.native.selection import document_key
from wps_mcp.presentation.access import get_slide
from wps_mcp.presentation.navigation import goto_slide
from wps_mcp.presentation.shapes.lookup import resolve_shape as _get_shape
from wps_mcp.presentation.shapes.models import ShapeReference
from wps_mcp.protocol.definitions import tool


class GetTagsInput(DocumentInput):
    """Input for getting tags from a shape, slide, or presentation."""

    slide_index: int | None = Field(
        default=None,
        ge=1,
        description="1-based slide index (required for slide/shape targets)",
    )
    shape_name_or_index: ShapeReference | None = Field(
        default=None,
        description="Shape name, group path, or "
        "1-based index; prefer a stable name.",
    )
    target_type: Literal["shape", "slide", "presentation"] = Field(
        default="shape",
        description="Target type: 'shape', 'slide', or 'presentation'",
    )

    @model_validator(mode="after")
    def validate_target(self) -> GetTagsInput:
        """Require the selectors needed by the chosen tag target."""
        if self.target_type != "presentation" and self.slide_index is None:
            raise ValueError(
                "slide_index is required for slide and shape targets"
            )
        if self.target_type == "shape" and self.shape_name_or_index is None:
            raise ValueError(
                "shape_name_or_index is required for shape targets"
            )
        return self


class SetTagInput(GetTagsInput):
    """Set a tag on the same target contract used by tag reads."""

    tag_name: str = Field(..., min_length=1, description="Tag name (key)")
    tag_value: str = Field(..., description="Tag value")


def _resolve_target(params: SetTagInput | GetTagsInput) -> NativeObject:
    """Resolve a tag target within the selected document."""
    if params.target_type == "presentation":
        return get_session("wpp").get_document(params.document)
    if params.slide_index is None:
        raise ValueError("slide_index is required for slide and shape tags")
    slide = get_slide(params.slide_index, params.document)
    if params.target_type == "slide":
        return slide
    if params.shape_name_or_index is None:
        raise ValueError("shape_name_or_index is required for shape tags")
    return _get_shape(slide, params.shape_name_or_index)


@tool(
    "wps_ppt_document",
    component="wpp",
    annotations={"title": "Set Tag", "idempotentHint": True},
    operation="set_tag",
)
def set_tag(params: SetTagInput) -> JsonObject:
    """Set a tag (key-value pair) on a shape, slide, or presentation."""
    target = _resolve_target(params)
    if params.slide_index is not None:
        goto_slide(params.slide_index)
    target.Tags.Add(params.tag_name, params.tag_value)
    result = {
        "success": True,
        "target_type": params.target_type,
        "tag_name": params.tag_name,
        "tag_value": params.tag_value,
    }
    return result


@tool(
    "wps_ppt_document",
    component="wpp",
    annotations={
        "title": "Get Tags",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="tags",
)
def get_tags(params: GetTagsInput) -> JsonObject:
    """Get all tags from a shape, slide, or presentation."""
    target = _resolve_target(params)
    tags = {}
    for i in range(1, target.Tags.Count + 1):
        tags[target.Tags.Name(i)] = target.Tags.Value(i)
    result = {
        "success": True,
        "target_type": params.target_type,
        "tags_count": target.Tags.Count,
        "tags": tags,
    }
    return result


logger = logging.getLogger(__name__)


class PresentationInput(DocumentInput):
    """Select a presentation by document, name or one-based index."""

    presentation_index: int | None = Field(default=None, ge=1)
    presentation_name: str | None = Field(default=None, min_length=1)

    @model_validator(mode="after")
    def validate_selection(self) -> Self:
        """Reject simultaneous presentation name and index selectors."""
        if (
            self.presentation_index is not None
            and self.presentation_name is not None
        ):
            raise ValueError(
                "Specify presentation_index or presentation_name, not both"
            )
        return self


class ListTemplatesInput(InputModel):
    """Choose a directory or the configured template location."""

    templates_dir: str | None = Field(default=None, min_length=1)


def resolve_presentation(
    app: NativeObject,
    presentation_index: int | None = None,
    presentation_name: str | None = None,
    *,
    document: str | None = None,
    enforce_target: bool = True,
) -> NativeObject:
    """Resolve selectors and reject explicit document conflicts."""
    if presentation_index is not None and presentation_name is not None:
        raise ValueError(
            "Specify presentation_index or presentation_name, not both"
        )
    session = get_session("wpp")
    explicit_document = document or (
        CALL_DOCUMENT.get() if enforce_target else None
    )
    if presentation_index is None and presentation_name is None:
        return session.get_document(explicit_document)
    if presentation_index is not None:
        if not 1 <= presentation_index <= app.Presentations.Count:
            raise ValueError(
                f"Presentation index {presentation_index} out of range "
                f"(1-{app.Presentations.Count})"
            )
        candidate = app.Presentations(presentation_index)
    else:
        candidate = session.get_document(presentation_name)
    if explicit_document is not None:
        requested = session.get_document(explicit_document)
        if document_key(candidate.FullName) != document_key(
            requested.FullName
        ):
            raise ValueError(
                "document conflicts with presentation_index/presentation_name"
            )
    return candidate


@tool(
    "wps_ppt_document",
    component="wpp",
    annotations={"idempotentHint": True},
    operation="activate",
)
def activate_presentation(params: PresentationInput) -> JsonObject:
    """Choose the presentation used by calls without a document."""
    if (
        params.document is None
        and params.presentation_index is None
        and params.presentation_name is None
    ):
        raise ValueError(
            "Specify document, presentation_index or presentation_name"
        )
    session = get_session("wpp")
    presentation = resolve_presentation(
        session.get_application(),
        params.presentation_index,
        params.presentation_name,
        document=params.document,
    )
    return session.set_target(presentation.FullName)


@tool(
    "wps_ppt_document",
    component=None,
    annotations={"readOnlyHint": True},
    operation="templates",
)
def list_templates(params: ListTemplatesInput) -> JsonObject:
    """List local WPS and PowerPoint templates without WPS."""
    root, bundled = template_directory(params.templates_dir)
    paths = root.glob("*/templates/**/*") if bundled else root.iterdir()
    templates = [
        {"name": path.stem, "file_name": path.name, "file_path": str(path)}
        for path in sorted(
            paths, key=lambda path: (path.name.casefold(), str(path))
        )
        if path.is_file()
        and path.suffix.lower() in {".dpt", ".potx", ".potm", ".pot"}
    ]
    return {
        "templates_dir": str(root),
        "count": len(templates),
        "templates": templates,
    }


PROPERTY_NAMES = {
    "title": "Title",
    "author": "Author",
    "subject": "Subject",
    "keywords": "Keywords",
    "comments": "Comments",
    "category": "Category",
    "company": "Company",
    "last_author": "Last Author",
    "creation_date": "Creation Date",
    "last_save_time": "Last Save Time",
}


class SetPropertiesInput(DocumentInput):
    """Only supplied non-null metadata fields are changed."""

    title: str | None = None
    author: str | None = None
    subject: str | None = None
    keywords: str | None = None
    comments: str | None = None
    category: str | None = None
    company: str | None = None


@tool(
    "wps_ppt_document",
    component="wpp",
    annotations={"idempotentHint": True},
    operation="set_properties",
)
def set_properties(params: SetPropertiesInput) -> JsonObject:
    """Update built-in document properties, preserving whitespace."""
    properties = (
        get_session("wpp")
        .get_document(params.document)
        .BuiltInDocumentProperties
    )
    updates = params.model_dump(exclude={"document"}, exclude_none=True)
    for name, value in updates.items():
        properties(PROPERTY_NAMES[name]).Value = value
    return {
        "success": True,
        "properties_set": len(updates),
        "set_names": list(updates),
    }


@tool(
    "wps_ppt_document",
    component="wpp",
    annotations={"readOnlyHint": True},
    operation="properties",
)
def get_properties(params: DocumentInput) -> JsonObject:
    """Read built-in properties; unavailable properties are null."""
    properties = (
        get_session("wpp")
        .get_document(params.document)
        .BuiltInDocumentProperties
    )
    result: JsonObject = {}
    for name, native_name in PROPERTY_NAMES.items():
        try:
            value = properties(native_name).Value
            if isinstance(value, (datetime, date)):
                value = value.isoformat()
            elif value is not None and not isinstance(
                value, (str, int, float, bool)
            ):
                value = str(value)
            result[name] = value
        except native_errors() + (AttributeError,) as error:
            logger.debug(
                "Document property %s is unavailable: %s", native_name, error
            )
            result[name] = None
    return {"success": True, "properties": result}
