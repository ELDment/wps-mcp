"""Inspect slide masters, layouts and text placeholders."""

import logging
from typing import Self

from pydantic import Field, model_validator

from wps_mcp.common.colors import int_to_hex
from wps_mcp.common.text import normalize_text
from wps_mcp.models import DocumentInput, JsonObject, NativeObject
from wps_mcp.native.objects import is_optional_member_error, optional_property
from wps_mcp.native.runtime import get_session
from wps_mcp.presentation.constants import (
    PLACEHOLDER_TYPE_NAMES,
    PP_PLACEHOLDER_BODY,
    PP_PLACEHOLDER_CENTER_TITLE,
    PP_PLACEHOLDER_SUBTITLE,
    PP_PLACEHOLDER_TITLE,
)
from wps_mcp.presentation.navigation import goto_slide
from wps_mcp.protocol.definitions import tool

logger = logging.getLogger(__name__)
PLACEHOLDER_TYPE_MAP = {
    "title": PP_PLACEHOLDER_TITLE,
    "body": PP_PLACEHOLDER_BODY,
    "center_title": PP_PLACEHOLDER_CENTER_TITLE,
    "subtitle": PP_PLACEHOLDER_SUBTITLE,
    "object": 7,
    "slide_number": 13,
    "footer": 15,
    "date": 16,
    "picture": 18,
}
CONTAINED_TYPE_NAMES = {
    1: "auto_shape",
    3: "chart",
    11: "linked_picture",
    13: "picture",
    14: "placeholder",
    16: "media",
    19: "table",
    24: "smart_art",
}


class ListPlaceholdersInput(DocumentInput):
    """Select one slide's placeholder collection."""

    slide_index: int = Field(ge=1)


class GetPlaceholderInput(ListPlaceholdersInput):
    """Select one placeholder by collection index or type."""

    placeholder_index: int | None = Field(default=None, ge=1)
    placeholder_type: str | int | None = None

    @model_validator(mode="after")
    def validate_selection(self) -> Self:
        """Require exactly one unambiguous target selection."""
        if (self.placeholder_index is None) == (self.placeholder_type is None):
            raise ValueError(
                "Specify placeholder_index or placeholder_type, not both"
            )
        return self


class SetPlaceholderTextInput(GetPlaceholderInput):
    """Replace text while preserving spaces and paragraph boundaries."""

    text: str


class ListDesignsInput(DocumentInput):
    """Include each design's custom layout names when requested."""

    include_layouts: bool = False


class DesignInput(DocumentInput):
    """Select a slide master by one-based design index."""

    design_index: int = Field(default=1, ge=1)


def resolve_placeholder(
    slide: NativeObject,
    placeholder_index: int | None,
    placeholder_type: str | int | None,
) -> tuple[int, NativeObject]:
    """Resolve a placeholder index or type with title/subtitle fallback."""
    placeholders = slide.Shapes.Placeholders
    if (placeholder_index is None) == (placeholder_type is None):
        raise ValueError(
            "Specify placeholder_index or placeholder_type, not both"
        )
    if placeholder_index is not None:
        if not 1 <= placeholder_index <= placeholders.Count:
            raise ValueError(
                f"Placeholder index out of range (1-{placeholders.Count})"
            )
        return (placeholder_index, placeholders(placeholder_index))
    if isinstance(placeholder_type, str):
        type_id = PLACEHOLDER_TYPE_MAP.get(placeholder_type.lower())
        if type_id is None:
            raise ValueError(f"Unknown placeholder type: {placeholder_type!r}")
    else:
        assert placeholder_type is not None
        type_id = placeholder_type
    fallback_type = {
        PP_PLACEHOLDER_TITLE: PP_PLACEHOLDER_CENTER_TITLE,
        PP_PLACEHOLDER_SUBTITLE: PP_PLACEHOLDER_BODY,
    }.get(type_id)
    fallback = None
    for index in range(1, placeholders.Count + 1):
        placeholder = placeholders(index)
        native_type = placeholder.PlaceholderFormat.Type
        if native_type == type_id:
            return (index, placeholder)
        if fallback is None and native_type == fallback_type:
            fallback = (index, placeholder)
    if fallback is not None:
        return fallback
    raise ValueError(
        f"No placeholder of type {placeholder_type!r} found on slide"
    )


@tool(
    "wps_ppt_placeholders",
    component="wpp",
    annotations={
        "title": "List Placeholders",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="list",
)
def list_placeholders(params: ListPlaceholdersInput) -> JsonObject:
    """List all placeholders on a slide.

    Returns index, type, name, position, size, and text preview for each
    placeholder on the specified slide.
    """
    pres = get_session("wpp").get_document(params.document)
    if params.slide_index > pres.Slides.Count:
        raise ValueError(f"Slide index out of range (1-{pres.Slides.Count})")
    slide = pres.Slides(params.slide_index)
    phs = slide.Shapes.Placeholders
    placeholders = []
    for i in range(1, phs.Count + 1):
        ph = phs(i)
        pf = ph.PlaceholderFormat
        info = {
            "index": i,
            "name": ph.Name,
            "type": pf.Type,
            "type_name": PLACEHOLDER_TYPE_NAMES.get(
                pf.Type, f"Unknown({pf.Type})"
            ),
            "has_text": False,
            "text_preview": None,
            "left": round(ph.Left, 1),
            "top": round(ph.Top, 1),
            "width": round(ph.Width, 1),
            "height": round(ph.Height, 1),
        }
        if ph.HasTextFrame:
            tf = ph.TextFrame
            has_text = bool(tf.HasText)
            info["has_text"] = has_text
            if has_text:
                text = tf.TextRange.Text
                info["text_preview"] = text[:100] + (
                    "..." if len(text) > 100 else ""
                )
        placeholders.append(info)
    return {
        "success": True,
        "slide_index": params.slide_index,
        "placeholder_count": phs.Count,
        "placeholders": placeholders,
    }


@tool(
    "wps_ppt_placeholders",
    component="wpp",
    annotations={
        "title": "Get Placeholder Details",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="get",
)
def get_placeholder(params: GetPlaceholderInput) -> JsonObject:
    """Get detailed information about a specific placeholder.

    Find by placeholder_index (1-based) or placeholder_type (e.g. 'title',
    'body', 'subtitle', or a PpPlaceholderType int). Returns text content,
    formatting, and position info.
    """
    pres = get_session("wpp").get_document(params.document)
    if params.slide_index > pres.Slides.Count:
        raise ValueError(f"Slide index out of range (1-{pres.Slides.Count})")
    slide = pres.Slides(params.slide_index)
    resolved_index, ph = resolve_placeholder(
        slide, params.placeholder_index, params.placeholder_type
    )
    pf = ph.PlaceholderFormat
    info = {
        "index": resolved_index,
        "name": ph.Name,
        "type": pf.Type,
        "type_name": PLACEHOLDER_TYPE_NAMES.get(
            pf.Type, f"Unknown({pf.Type})"
        ),
        "contained_type": pf.ContainedType,
        "contained_type_name": CONTAINED_TYPE_NAMES.get(
            pf.ContainedType, f"Unknown({pf.ContainedType})"
        ),
        "left": round(ph.Left, 1),
        "top": round(ph.Top, 1),
        "width": round(ph.Width, 1),
        "height": round(ph.Height, 1),
        "rotation": ph.Rotation,
        "has_text_frame": bool(ph.HasTextFrame),
    }
    if ph.HasTextFrame:
        tf = ph.TextFrame
        tr = tf.TextRange
        info["text"] = tr.Text if tf.HasText else None
        info["paragraph_count"] = tr.Paragraphs().Count
        if tf.HasText:
            info["font_name"] = tr.Font.Name
            info["font_name_far_east"] = tr.Font.NameFarEast
            try:
                info["font_size"] = tr.Font.Size
            except Exception as error:
                if not is_optional_member_error(error):
                    raise
                info["font_size"] = None
            info["alignment"] = tr.ParagraphFormat.Alignment
    return {"success": True, "placeholder": info}


@tool(
    "wps_ppt_placeholders",
    component="wpp",
    annotations={"title": "Set Placeholder Text", "idempotentHint": True},
    operation="set_text",
)
def set_placeholder_text(params: SetPlaceholderTextInput) -> JsonObject:
    """Set text in a placeholder.

    Find by placeholder_index (1-based) or placeholder_type (e.g. 'title',
    'body', 'subtitle'). Use \\n for paragraph breaks.
    """
    text = params.text
    pres = get_session("wpp").get_document(params.document)
    if params.slide_index > pres.Slides.Count:
        raise ValueError(f"Slide index out of range (1-{pres.Slides.Count})")
    slide = pres.Slides(params.slide_index)
    resolved_index, ph = resolve_placeholder(
        slide, params.placeholder_index, params.placeholder_type
    )
    if not ph.HasTextFrame:
        raise ValueError(
            f"Placeholder '{ph.Name}' does not have a text frame "
            f"(type={ph.PlaceholderFormat.Type})"
        )
    goto_slide(params.slide_index)
    text = normalize_text(text)
    ph.TextFrame.TextRange.Text = text
    return {
        "success": True,
        "slide_index": params.slide_index,
        "placeholder_name": ph.Name,
        "placeholder_index": resolved_index,
        "placeholder_type": ph.PlaceholderFormat.Type,
        "text_length": ph.TextFrame.TextRange.Length,
    }


@tool(
    "wps_ppt_placeholders",
    component="wpp",
    annotations={
        "title": "List Designs",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="designs",
)
def list_designs(params: ListDesignsInput) -> JsonObject:
    """List all designs (slide masters) in the presentation.

    Returns each design's index, name, layout count, and layout names. Use
    the design_index with wps_ppt_slides/add or
    wps_ppt_placeholders/layouts to work with layouts from a specific
    design.
    """
    pres = get_session("wpp").get_document(params.document)
    designs = pres.Designs
    result = []
    for d in range(1, designs.Count + 1):
        design = designs(d)
        master = design.SlideMaster
        entry = {
            "design_index": d,
            "design_name": design.Name,
            "layout_count": master.CustomLayouts.Count,
        }
        if params.include_layouts:
            layout_names = []
            for i in range(1, master.CustomLayouts.Count + 1):
                layout_names.append(master.CustomLayouts(i).Name)
            entry["layout_names"] = layout_names
        result.append(entry)
    return {"success": True, "design_count": designs.Count, "designs": result}


@tool(
    "wps_ppt_placeholders",
    component="wpp",
    annotations={
        "title": "List Slide Layouts",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="layouts",
)
def list_layouts(params: DesignInput) -> JsonObject:
    """List all available slide layouts in the presentation.

    Returns layout name, index, and placeholder info for each CustomLayout
    in the specified design (master).
    """
    pres = get_session("wpp").get_document(params.document)
    if params.design_index > pres.Designs.Count:
        raise ValueError(f"Design index out of range (1-{pres.Designs.Count})")
    design = pres.Designs(params.design_index)
    master = design.SlideMaster
    layouts_col = master.CustomLayouts
    layouts = []
    for i in range(1, layouts_col.Count + 1):
        layout = layouts_col(i)
        phs = layout.Shapes.Placeholders
        ph_list = []
        for j in range(1, phs.Count + 1):
            ph = phs(j)
            pf = ph.PlaceholderFormat
            ph_list.append(
                {
                    "type": pf.Type,
                    "type_name": PLACEHOLDER_TYPE_NAMES.get(
                        pf.Type, f"Unknown({pf.Type})"
                    ),
                    "name": ph.Name,
                }
            )
        layouts.append(
            {
                "index": i,
                "name": layout.Name,
                "placeholder_count": phs.Count,
                "placeholders": ph_list,
            }
        )
    return {
        "success": True,
        "design_name": design.Name,
        "layout_count": layouts_col.Count,
        "layouts": layouts,
    }


@tool(
    "wps_ppt_placeholders",
    component="wpp",
    annotations={
        "title": "Get Slide Master Info",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="master",
)
def get_slide_master_info(params: DesignInput) -> JsonObject:
    """Get slide master information including theme colors.

    Returns the master name, layout count, and theme color scheme for the
    specified design.
    """
    pres = get_session("wpp").get_document(params.document)
    if params.design_index > pres.Designs.Count:
        raise ValueError(f"Design index out of range (1-{pres.Designs.Count})")
    design = pres.Designs(params.design_index)
    master = design.SlideMaster
    colors = []
    try:
        tcs = master.Theme.ThemeColorScheme
        theme_color_names = [
            "Dark1",
            "Light1",
            "Dark2",
            "Light2",
            "Accent1",
            "Accent2",
            "Accent3",
            "Accent4",
            "Accent5",
            "Accent6",
            "Hyperlink",
            "FollowedHyperlink",
        ]
        for i in range(1, min(tcs.Count + 1, 13)):
            color_val = tcs(i).RGB
            hex_color = int_to_hex(color_val)
            colors.append(
                {
                    "index": i,
                    "name": theme_color_names[i - 1]
                    if i <= 12
                    else f"Color{i}",
                    "color_hex": hex_color,
                }
            )
    except Exception as error:
        if not is_optional_member_error(error):
            raise
        logger.debug(
            "Optional native presentation property unavailable: %s", error
        )
    has_title = optional_property(design, "HasTitleMaster")
    return {
        "success": True,
        "design_name": design.Name,
        "master_name": master.Name,
        "layout_count": master.CustomLayouts.Count,
        "has_title_master": None if has_title is None else bool(has_title),
        "theme_colors": colors,
    }
