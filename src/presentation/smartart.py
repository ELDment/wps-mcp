"""SmartArt operations for WPS Presentation COM automation.

Handles creating SmartArt graphics, modifying nodes (set text, add,
delete, format), applying color schemes and quick styles, and listing
available SmartArt layouts, color schemes, and quick styles.
"""

from typing import Literal

from pydantic import Field

from src.common.colors import hex_to_int
from src.common.text import font_size_warning
from src.models import DocumentInput, InputModel, JsonObject, NativeObject
from src.native.objects import require_feature, supports_feature
from src.native.runtime import get_session
from src.presentation.constants import MSO_FALSE, MSO_SMART_ART, MSO_TRUE
from src.presentation.shape_lookup import resolve_shape as _get_shape
from src.protocol.definitions import tool


def _resolve_layout(app: NativeObject, layout_name: str) -> NativeObject:
    """Resolve a layout from the names and IDs supplied by WPS."""
    key = layout_name.casefold()
    layouts = app.SmartArtLayouts
    exact = []
    partial = []
    for index in range(1, layouts.Count + 1):
        layout = layouts(index)
        if key in (str(layout.Name).casefold(), str(layout.Id).casefold()):
            exact.append(layout)
        elif key in str(layout.Name).casefold():
            partial.append(layout)
    matches = exact or partial
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise ValueError(
            "Multiple WPS layouts match; use the "
            "exact layout name or layout_index"
        )
    raise ValueError(
        f"WPS layout '{layout_name}' not found; use wps_ppt_smartart/options."
    )


class AddSmartArtInput(DocumentInput):
    """Input for adding a SmartArt graphic to a slide."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    layout_name: str | None = Field(
        default=None,
        description="Exact or partial layout name, "
        "or an ID from wps_ppt_smartart/options.",
    )
    layout_index: int | None = Field(
        default=None,
        ge=1,
        description="One-based layout index; defaults to the first layout.",
    )
    left: float = Field(default=50.0, description="Left position in points")
    top: float = Field(default=50.0, description="Top position in points")
    width: float = Field(default=600.0, description="Width in points", gt=0)
    height: float = Field(default=400.0, description="Height in points", gt=0)
    node_texts: list[str] | None = Field(
        default=None,
        description="List of text strings to populate SmartArt nodes in order",
    )
    color_index: int | None = Field(
        default=None,
        ge=1,
        description="One-based color scheme "
        "index from wps_ppt_smartart/options.",
    )
    style_index: int | None = Field(
        default=None,
        ge=1,
        description="One-based quick style index; applied before color_index.",
    )
    font_name: str | None = Field(
        default=None, description="Font for both Latin and East Asian text."
    )
    font_size: float | None = Field(
        default=None,
        gt=0,
        description="Font size in points to apply to all nodes.",
    )
    bold: bool | None = Field(
        default=None, description="Bold on/off for all nodes."
    )
    font_color: str | None = Field(
        default=None,
        description="Text color as #RRGGBB.",
        pattern="^#[0-9a-fA-F]{6}$",
    )


class ModifySmartArtInput(DocumentInput):
    """Choose edits for a SmartArt node or the entire diagram."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: str | int = Field(
        ...,
        description="Shape name or one-based index; "
        "names remain stable after insertions.",
    )
    action: Literal[
        "set_text",
        "add_node",
        "delete_node",
        "change_layout",
        "change_color",
        "change_style",
        "format_node",
        "format_all_nodes",
    ] = Field(
        ...,
        description="Node editing, layout "
        "selection, or diagram formatting operation.",
    )
    node_index: int | None = Field(
        default=None,
        ge=1,
        description="One-based node index; omit when appending a node.",
    )
    text: str | None = Field(
        default=None,
        description="Text to set on the node. Required "
        "for 'set_text', optional for 'add_node'.",
    )
    layout_name: str | None = Field(
        default=None,
        description="Exact or partial layout name, "
        "or an ID from wps_ppt_smartart/options.",
    )
    layout_index: int | None = Field(
        default=None,
        ge=1,
        description="One-based layout index; defaults to the first layout.",
    )
    color_index: int | None = Field(
        default=None,
        ge=1,
        description="One-based color scheme "
        "index from wps_ppt_smartart/options.",
    )
    style_index: int | None = Field(
        default=None,
        ge=1,
        description="One-based quick style index; applied before color_index.",
    )
    font_name: str | None = Field(
        default=None,
        description="Font for both Latin and East Asian text.",
    )
    font_size: float | None = Field(
        default=None,
        gt=0,
        description="Font size in points. Used "
        "with 'format_node' and 'format_all_nodes'.",
    )
    bold: bool | None = Field(
        default=None,
        description="Bold on/off. Used with "
        "'format_node' and 'format_all_nodes'.",
    )
    fill_color: str | None = Field(
        default=None,
        description="Node fill color as #RRGGBB.",
        pattern="^#[0-9a-fA-F]{6}$",
    )
    line_color: str | None = Field(
        default=None,
        description="Node border color as #RRGGBB.",
        pattern="^#[0-9a-fA-F]{6}$",
    )
    line_width: float | None = Field(
        default=None,
        gt=0,
        description="Node border width in points. Used "
        "with 'format_node' and 'format_all_nodes'.",
    )
    font_color: str | None = Field(
        default=None,
        description="Text color as #RRGGBB.",
        pattern="^#[0-9a-fA-F]{6}$",
    )


class ListSmartArtInput(InputModel):
    """Select SmartArt layouts, color schemes or quick styles."""

    list_type: Literal["layouts", "colors", "styles", "categories"] = Field(
        default="layouts",
        description="Return layouts, color "
        "schemes, quick styles, or categories.",
    )
    category: str | None = Field(
        default=None,
        description="Case-insensitive substring filter for layout categories.",
    )
    keyword: str | None = Field(
        default=None,
        description="Case-insensitive substring filter for names.",
    )
    include_description: bool = Field(
        default=False,
        description="Include descriptions in returned entries.",
    )


def _apply_node_format(
    node: NativeObject,
    fill_color: str | None,
    line_color: str | None,
    line_width: float | None,
    font_name: str | None,
    font_size: float | None,
    bold: bool | None,
    font_color: str | None,
) -> None:
    """Apply fill/line/font formatting to a single SmartArtNode."""
    if (
        fill_color is not None
        or line_color is not None
        or line_width is not None
    ):
        if node.Shapes.Count == 0:
            raise ValueError("This SmartArt node has no shape to format")
        shape = node.Shapes(1)
        if fill_color is not None:
            shape.Fill.ForeColor.RGB = hex_to_int(fill_color)
        if line_color is not None:
            shape.Line.ForeColor.RGB = hex_to_int(line_color)
        if line_width is not None:
            shape.Line.Weight = line_width
    if (
        font_name is not None
        or font_size is not None
        or bold is not None
        or (font_color is not None)
    ):
        f = node.TextFrame2.TextRange.Font
        if font_size is not None:
            f.Size = font_size
        if font_name is not None:
            f.Name = font_name
            f.NameFarEast = font_name
        if bold is not None:
            f.Bold = MSO_TRUE if bold else MSO_FALSE
        if font_color is not None:
            f.Fill.ForeColor.RGB = hex_to_int(font_color)


@tool(
    "wps_ppt_smartart",
    component="wpp",
    annotations={"title": "Add SmartArt"},
    operation="add",
)
def add_smartart(params: AddSmartArtInput) -> JsonObject:
    """Create SmartArt and populate every requested node."""
    app = get_session("wpp").get_application()
    pres = get_session("wpp").get_document()
    if params.node_texts or any(
        value is not None
        for value in (
            params.font_name,
            params.font_size,
            params.bold,
            params.font_color,
        )
    ):
        require_feature(pres, "smartart_nodes")
    slide = pres.Slides(params.slide_index)
    if params.layout_name:
        layout = _resolve_layout(app, params.layout_name)
    else:
        idx = params.layout_index if params.layout_index else 1
        layout = app.SmartArtLayouts(idx)
    resolved_layout_name = layout.Name
    shape = slide.Shapes.AddSmartArt(
        layout, params.left, params.top, params.width, params.height
    )
    if params.node_texts:
        smart_art = shape.SmartArt
        for i, text in enumerate(params.node_texts):
            if i < smart_art.AllNodes.Count:
                smart_art.AllNodes(i + 1).TextFrame2.TextRange.Text = text
            else:
                node = (
                    smart_art.AllNodes.Add()
                    if smart_art.AllNodes.Count == 0
                    else smart_art.AllNodes(smart_art.AllNodes.Count).AddNode()
                )
                node.TextFrame2.TextRange.Text = text
    smart_art = shape.SmartArt
    if params.style_index is not None:
        smart_art.QuickStyle = app.SmartArtQuickStyles(params.style_index)
    if params.color_index is not None:
        smart_art.Color = app.SmartArtColors(params.color_index)
    if (
        params.font_name is not None
        or params.font_size is not None
        or params.bold is not None
        or (params.font_color is not None)
    ):
        for i in range(1, smart_art.AllNodes.Count + 1):
            _apply_node_format(
                smart_art.AllNodes(i),
                None,
                None,
                None,
                params.font_name,
                params.font_size,
                params.bold,
                params.font_color,
            )
    result = {
        "success": True,
        "shape_name": shape.Name,
        "node_count": (
            None
            if not supports_feature(smart_art, "smartart_nodes")
            else smart_art.AllNodes.Count
        ),
        "layout_name": resolved_layout_name,
    }
    warn = font_size_warning(params.font_size)
    if warn:
        result["warning"] = warn
    return result


@tool(
    "wps_ppt_smartart",
    component="wpp",
    annotations={"title": "Modify SmartArt"},
    operation="modify",
)
def modify_smartart(params: ModifySmartArtInput) -> JsonObject:
    """Edit a SmartArt node or apply a diagram-wide style."""
    app = get_session("wpp").get_application()
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    shape = _get_shape(slide, params.shape_name_or_index)
    if shape.Type != MSO_SMART_ART:
        raise ValueError(f"Shape '{shape.Name}' is not a SmartArt graphic")
    smart_art = shape.SmartArt
    if params.node_index is not None or params.action not in {
        "change_layout",
        "change_color",
        "change_style",
    }:
        require_feature(smart_art, "smartart_nodes")
    if (
        params.node_index is not None
        and params.node_index > smart_art.AllNodes.Count
    ):
        raise ValueError(
            f"node_index must not exceed {smart_art.AllNodes.Count}"
        )
    if params.action == "set_text":
        if params.node_index is None:
            raise ValueError("node_index is required for 'set_text' action")
        if params.text is None:
            raise ValueError("text is required for 'set_text' action")
        smart_art.AllNodes(
            params.node_index
        ).TextFrame2.TextRange.Text = params.text
    elif params.action == "add_node":
        if params.node_index and params.node_index <= smart_art.AllNodes.Count:
            new_node = smart_art.AllNodes(params.node_index).AddNode()
        else:
            new_node = (
                smart_art.AllNodes.Add()
                if smart_art.AllNodes.Count == 0
                else smart_art.AllNodes(smart_art.AllNodes.Count).AddNode()
            )
        if params.text is not None:
            new_node.TextFrame2.TextRange.Text = params.text
    elif params.action == "delete_node":
        if params.node_index is None:
            raise ValueError("node_index is required for 'delete_node' action")
        smart_art.AllNodes(params.node_index).Delete()
    elif params.action == "change_layout":
        if params.layout_name:
            layout = _resolve_layout(app, params.layout_name)
        elif params.layout_index:
            layout = app.SmartArtLayouts(params.layout_index)
        else:
            raise ValueError(
                "layout_name or layout_index is "
                "required for 'change_layout' action"
            )
        smart_art.Layout = layout
    elif params.action == "change_color":
        if params.color_index is None:
            raise ValueError(
                "color_index is required for 'change_color' action"
            )
        smart_art.Color = app.SmartArtColors(params.color_index)
    elif params.action == "change_style":
        if params.style_index is None:
            raise ValueError(
                "style_index is required for 'change_style' action"
            )
        smart_art.QuickStyle = app.SmartArtQuickStyles(params.style_index)
        if params.color_index is not None:
            smart_art.Color = app.SmartArtColors(params.color_index)
    elif params.action == "format_node":
        if params.node_index is None:
            raise ValueError("node_index is required for 'format_node' action")
        node = smart_art.AllNodes(params.node_index)
        _apply_node_format(
            node,
            params.fill_color,
            params.line_color,
            params.line_width,
            params.font_name,
            params.font_size,
            params.bold,
            params.font_color,
        )
    elif params.action == "format_all_nodes":
        for i in range(1, smart_art.AllNodes.Count + 1):
            _apply_node_format(
                smart_art.AllNodes(i),
                params.fill_color,
                params.line_color,
                params.line_width,
                params.font_name,
                params.font_size,
                params.bold,
                params.font_color,
            )
    else:
        raise ValueError(f"Unknown SmartArt action: {params.action}")
    result = {
        "success": True,
        "action": params.action,
        "total_nodes": (
            None
            if not supports_feature(smart_art, "smartart_nodes")
            else smart_art.AllNodes.Count
        ),
    }
    warn = font_size_warning(params.font_size)
    if warn:
        result["warning"] = warn
    return result


@tool(
    "wps_ppt_smartart",
    component="wpp",
    annotations={
        "title": "List SmartArt Layouts / Colors / Styles",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="options",
)
def list_smartart_options(params: ListSmartArtInput) -> JsonObject:
    """List available SmartArt options in the WPS locale."""
    app = get_session("wpp").get_application()
    if params.list_type == "categories":
        collection = app.SmartArtLayouts
        seen: dict[str, int] = {}
        for i in range(1, collection.Count + 1):
            try:
                cat = collection(i).Category or ""
            except Exception:
                cat = ""
            if cat and (
                params.keyword is None
                or params.keyword.casefold() in cat.casefold()
            ):
                seen[cat] = seen.get(cat, 0) + 1
        categories = [{"category": k, "count": v} for k, v in seen.items()]
        return {
            "success": True,
            "list_type": "categories",
            "total_layouts": collection.Count,
            "categories": categories,
        }
    if params.list_type == "layouts":
        collection = app.SmartArtLayouts
        key = "layouts"
    elif params.list_type == "colors":
        collection = app.SmartArtColors
        key = "colors"
    elif params.list_type == "styles":
        collection = app.SmartArtQuickStyles
        key = "styles"
    else:
        raise ValueError(f"Unknown SmartArt list type: {params.list_type}")
    total = collection.Count
    cat_lower = params.category.lower() if params.category else None
    kw_lower = params.keyword.lower() if params.keyword else None
    items = []
    for i in range(1, total + 1):
        item = collection(i)
        item_name = item.Name
        if cat_lower and params.list_type == "layouts":
            try:
                item_cat = item.Category or ""
            except Exception:
                item_cat = ""
            if cat_lower not in item_cat.lower():
                continue
        if kw_lower and kw_lower not in item_name.lower():
            continue
        entry = {"index": i, "name": item_name}
        if params.list_type == "layouts":
            entry["id"] = item.Id
            try:
                entry["category"] = item.Category
            except Exception:
                entry["category"] = None
        if params.include_description:
            entry["description"] = item.Description
        items.append(entry)
    return {
        "success": True,
        "list_type": params.list_type,
        "total_count": total,
        "filtered_count": len(items),
        key: items,
    }
