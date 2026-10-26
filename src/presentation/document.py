"""Presentation lifecycle, selection, metadata and tags."""

from __future__ import annotations

import logging
import os
from datetime import date, datetime
from pathlib import Path
from typing import Literal, Self

from pydantic import Field, model_validator

from src.common.colors import int_to_hex
from src.documents.models import CloseInput, ExportInput, SaveInput
from src.documents.service import close_document, export_pdf, save_document
from src.models import DocumentInput, InputModel, JsonObject, NativeObject
from src.native.context import CALL_DOCUMENT
from src.native.discovery import template_directory
from src.native.runtime import get_session
from src.presentation.constants import (
    MSO_FALSE,
    MSO_TRUE,
)
from src.presentation.navigation import goto_slide
from src.presentation.shape_lookup import resolve_shape as _get_shape
from src.protocol.definitions import tool


class GetTagsInput(DocumentInput):
    """Input for getting tags from a shape, slide, or presentation."""

    slide_index: int | None = Field(
        default=None,
        ge=1,
        description="1-based slide index (required for slide/shape targets)",
    )
    shape_name_or_index: str | int | None = Field(
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
    presentation = get_session("wpp").get_document()
    if params.target_type == "presentation":
        return presentation
    if params.slide_index is None:
        raise ValueError("slide_index is required for slide and shape tags")
    slide = presentation.Slides(params.slide_index)
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
    app = get_session("wpp").get_application()
    target = _resolve_target(params)
    if params.slide_index is not None:
        goto_slide(app, params.slide_index)
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


@tool(
    "wps_ppt_document",
    component="wpp",
    annotations={"readOnlyHint": True},
    operation="list",
)
def list_presentations(params: InputModel) -> JsonObject:
    """List open deck paths, slide counts and save status."""
    presentations = get_session("wpp").get_application().Presentations
    result: list[JsonObject] = []
    for index in range(1, presentations.Count + 1):
        presentation = presentations(index)
        result.append(
            {
                "index": index,
                "name": presentation.Name,
                "full_name": presentation.FullName,
                "path": presentation.Path,
                "slides_count": presentation.Slides.Count,
                "read_only": int(presentation.ReadOnly) == -1,
                "saved": bool(presentation.Saved),
            }
        )
    return {"presentations": result, "count": len(result)}


logger = logging.getLogger(__name__)
SLIDE_SIZE_PRESETS = {"16:9": (960.0, 540.0), "4:3": (720.0, 540.0)}


class CreatePresentationInput(InputModel):
    """Create a blank deck or an untitled copy of a template."""

    template_path: str | None = Field(default=None, min_length=1)
    slide_width: float | None = Field(default=None, gt=0)
    slide_height: float | None = Field(default=None, gt=0)
    preset: Literal["16:9", "4:3"] | None = None
    activate: bool = True

    @model_validator(mode="after")
    def validate_dimensions(self) -> Self:
        """Require paired dimensions for a custom blank deck."""
        if self.template_path is None and self.preset is None:
            if (self.slide_width is None) != (self.slide_height is None):
                raise ValueError(
                    "slide_width and slide_height must be supplied together"
                )
        return self


class OpenPresentationInput(InputModel):
    """Open a local presentation, optionally without a window."""

    file_path: str = Field(min_length=1)
    read_only: bool = False
    with_window: bool = True
    activate: bool = True


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


class SavePresentationAsInput(PresentationInput):
    """Save a deck or export it to PDF or a directory of images."""

    file_path: str = Field(min_length=1)
    format: (
        Literal["dps", "ppt", "pptx", "pdf", "png", "jpg", "default"] | None
    ) = None
    overwrite: bool = False


class ClosePresentationInput(PresentationInput):
    """Explicitly save or discard unsaved edits before closing."""

    save_changes: bool = False
    discard_changes: bool = False

    @model_validator(mode="after")
    def validate_close(self) -> Self:
        """Reject conflicting save and discard instructions."""
        if self.save_changes and self.discard_changes:
            raise ValueError(
                "save_changes and discard_changes cannot both be true"
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
        if os.path.normcase(
            os.path.normpath(candidate.FullName)
        ) != os.path.normcase(os.path.normpath(requested.FullName)):
            raise ValueError(
                "document conflicts with presentation_index/presentation_name"
            )
    return candidate


@tool("wps_ppt_document", component="wpp", operation="create")
def create_presentation(params: CreatePresentationInput) -> JsonObject:
    """Create a blank deck or retain all template content and styles."""
    template = (
        Path(params.template_path).resolve() if params.template_path else None
    )
    if template is not None and not template.is_file():
        raise FileNotFoundError(f"Template not found: {template}")
    session = get_session("wpp")
    app = session.connect(visible=True, allow_launch=True)
    if template is not None:
        presentation = app.Presentations.Open(
            str(template), MSO_FALSE, MSO_TRUE, MSO_TRUE
        )
    else:
        presentation = app.Presentations.Add()
        dimensions = (
            SLIDE_SIZE_PRESETS.get(params.preset) if params.preset else None
        )
        if dimensions is None and params.slide_width is not None:
            assert params.slide_height is not None
            dimensions = (params.slide_width, params.slide_height)
        if dimensions is not None:
            (
                presentation.PageSetup.SlideWidth,
                presentation.PageSetup.SlideHeight,
            ) = dimensions
    if params.activate:
        session.target_name = presentation.FullName
        try:
            presentation.Windows(1).Activate()
        except Exception as error:
            logger.warning(
                "Created presentation; window activation failed: %s", error
            )
    index = next(
        (
            index
            for index in range(1, app.Presentations.Count + 1)
            if app.Presentations(index).FullName == presentation.FullName
        ),
        None,
    )
    try:
        template_name = presentation.TemplateName
    except Exception as error:
        logger.debug("Template name unavailable: %s", error)
        template_name = None
    return {
        "success": True,
        "presentation_index": index,
        "name": presentation.Name,
        "full_name": presentation.FullName,
        "slides_count": presentation.Slides.Count,
        "slide_width": presentation.PageSetup.SlideWidth,
        "slide_height": presentation.PageSetup.SlideHeight,
        "template_name": template_name,
        "activated": params.activate,
    }


@tool(
    "wps_ppt_document",
    component="wpp",
    annotations={"openWorldHint": True},
    operation="open",
)
def open_presentation(params: OpenPresentationInput) -> JsonObject:
    """Open a presentation and optionally make it the session target."""
    file_path = Path(params.file_path).resolve()
    if not file_path.is_file():
        raise FileNotFoundError(f"File not found: {file_path}")
    session = get_session("wpp")
    app = session.connect(
        visible=True if params.with_window else None, allow_launch=True
    )
    presentation = app.Presentations.Open(
        FileName=str(file_path),
        ReadOnly=MSO_TRUE if params.read_only else MSO_FALSE,
        Untitled=MSO_FALSE,
        WithWindow=MSO_TRUE if params.with_window else MSO_FALSE,
    )
    if params.activate:
        session.target_name = presentation.FullName
        if params.with_window:
            try:
                presentation.Windows(1).Activate()
            except Exception as error:
                logger.warning(
                    "Opened presentation; window activation failed: %s", error
                )
    index = next(
        (
            index
            for index in range(1, app.Presentations.Count + 1)
            if app.Presentations(index).FullName == presentation.FullName
        ),
        None,
    )
    return {
        "success": True,
        "presentation_index": index,
        "name": presentation.Name,
        "full_name": presentation.FullName,
        "slides_count": presentation.Slides.Count,
        "read_only": int(presentation.ReadOnly) == MSO_TRUE,
        "activated": params.activate,
    }


@tool(
    "wps_ppt_document",
    component="wpp",
    annotations={"openWorldHint": True},
    operation="save",
)
def save_presentation(params: PresentationInput) -> JsonObject:
    """Save a presentation through the shared document lifecycle."""
    presentation = resolve_presentation(
        get_session("wpp").get_application(),
        params.presentation_index,
        params.presentation_name,
        document=params.document,
    )
    return save_document(
        SaveInput(component="wpp", document=presentation.FullName)
    )


@tool(
    "wps_ppt_document",
    component="wpp",
    annotations={"openWorldHint": True},
    operation="save_as",
)
def save_presentation_as(params: SavePresentationAsInput) -> JsonObject:
    """Save a deck or export a separate PDF/image copy.

    PNG/JPG use file_path as an output directory. Existing images are
    never replaced.
    """
    presentation = resolve_presentation(
        get_session("wpp").get_application(),
        params.presentation_index,
        params.presentation_name,
        document=params.document,
    )
    suffix = Path(params.file_path).suffix.lower().lstrip(".")
    output_format = (
        params.format if params.format not in (None, "default") else suffix
    )
    if output_format in ("png", "jpg"):
        from src.presentation.export import (
            ExportImagesInput,
            export_images,
        )

        if params.overwrite:
            raise ValueError(
                "Image export does not support "
                "overwrite; choose a new directory"
            )
        return export_images(
            ExportImagesInput(
                document=presentation.FullName,
                output_dir=params.file_path,
                format="png" if output_format == "png" else "jpg",
            )
        )
    if params.format not in (None, "default") and output_format != suffix:
        raise ValueError("format must match the output file extension")
    if output_format == "pdf":
        return export_pdf(
            ExportInput(
                component="wpp",
                document=presentation.FullName,
                file_path=params.file_path,
                overwrite=params.overwrite,
            )
        )
    return save_document(
        SaveInput(
            component="wpp",
            document=presentation.FullName,
            file_path=params.file_path,
            overwrite=params.overwrite,
        )
    )


@tool(
    "wps_ppt_document",
    component="wpp",
    annotations={"destructiveHint": True},
    operation="close",
)
def close_presentation(params: ClosePresentationInput) -> JsonObject:
    """Close through the shared lifecycle without implicit discard."""
    presentation = resolve_presentation(
        get_session("wpp").get_application(),
        params.presentation_index,
        params.presentation_name,
        document=params.document,
    )
    return close_document(
        CloseInput(
            component="wpp",
            document=presentation.FullName,
            save_changes=params.save_changes,
            discard_changes=params.discard_changes,
        )
    )


@tool(
    "wps_ppt_document",
    component="wpp",
    annotations={"readOnlyHint": True},
    operation="info",
)
def get_presentation_info(params: PresentationInput) -> JsonObject:
    """Read dimensions, file identity, theme fonts and accent colors."""
    presentation = resolve_presentation(
        get_session("wpp").get_application(),
        params.presentation_index,
        params.presentation_name,
        document=params.document,
    )
    page = presentation.PageSetup
    warnings: list[str] = []
    try:
        template_name = presentation.TemplateName
    except Exception as error:
        template_name = None
        warnings.append(f"Template name unavailable: {error}")
    fonts: dict[str, str | None] = dict.fromkeys(
        ("title_latin", "title_east_asian", "body_latin", "body_east_asian")
    )
    try:
        try:
            font_scheme = presentation.Designs(
                1
            ).SlideMaster.Theme.ThemeFontScheme
        except Exception:
            font_scheme = presentation.SlideMaster.Theme.ThemeFontScheme
        for prefix, native_name in (
            ("body", "MinorFont"),
            ("title", "MajorFont"),
        ):
            for language, index in (("latin", 1), ("east_asian", 3)):
                name = getattr(font_scheme, native_name)(index).Name
                fonts[f"{prefix}_{language}"] = (
                    name if name and not name.startswith("+") else None
                )
    except Exception as error:
        warnings.append(f"Theme fonts unavailable: {error}")
    colors: dict[str, str | None] = {
        f"accent{index}": None for index in range(1, 7)
    }
    try:
        color_scheme = presentation.SlideMaster.Theme.ThemeColorScheme
        for index, name in enumerate(colors, start=5):
            colors[name] = int_to_hex(int(color_scheme(index).RGB))
    except Exception as error:
        warnings.append(f"Theme colors unavailable: {error}")
    full_name = presentation.FullName
    local_path = full_name if Path(full_name).is_absolute() else None
    result: JsonObject = {
        "name": presentation.Name,
        "full_name": full_name,
        "local_path": local_path,
        "local_dir": str(Path(local_path).parent) if local_path else None,
        "path": presentation.Path,
        "slides_count": presentation.Slides.Count,
        "read_only": int(presentation.ReadOnly) == MSO_TRUE,
        "saved": bool(presentation.Saved),
        "slide_width": page.SlideWidth,
        "slide_height": page.SlideHeight,
        "slide_width_inches": round(page.SlideWidth / 72, 3),
        "slide_height_inches": round(page.SlideHeight / 72, 3),
        "first_slide_number": page.FirstSlideNumber,
        "template_name": template_name,
        "fonts": fonts,
        "accent_colors": colors,
    }
    if warnings:
        result["warnings"] = warnings
    return result


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
        except Exception as error:
            logger.debug(
                "Document property %s is unavailable: %s", native_name, error
            )
            result[name] = None
    return {"success": True, "properties": result}
