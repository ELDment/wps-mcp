"""Create, modify, group and arrange shapes on presentation slides."""

from __future__ import annotations

import logging

from src.common.colors import hex_to_int
from src.common.text import font_size_warning, normalize_text
from src.models import JsonObject, NativeObject
from src.native.runtime import get_session
from src.presentation.constants import (
    GRADIENT_STYLE_MAP,
    MSO_FALSE,
    MSO_GROUP,
    MSO_TEXT_ORIENTATION_HORIZONTAL,
    MSO_TRUE,
    SHAPE_TYPE_NAMES,
)
from src.presentation.navigation import goto_slide
from src.presentation.shape_geometry import apply_geometry, select_targets
from src.presentation.shape_lookup import (
    PATH_SEPARATOR,
    require_top_level,
    resolve_shape,
    resolve_with_path,
)
from src.presentation.shape_models import (
    AddLineInput,
    AddShapeInput,
    AddTextboxInput,
    CopyShapeToSlideInput,
    GroupShapesInput,
    LockAspectRatioInput,
    SelectShapesInput,
    SetZOrderInput,
    ShapeIdentifierInput,
    ShapeInput,
    UpdateShapeInput,
)
from src.presentation.shape_types import (
    ADJUSTMENT_LABELS,
    SHAPE_NAME_MAP,
    resolve_shape_type,
)
from src.presentation.text_enums import ALIGNMENT_MAP, VERTICAL_ANCHOR_MAP
from src.presentation.text_formatting import apply_font_format
from src.presentation.text_models import TextFormatSpec
from src.presentation.zorder import (
    BEHIND_TEXT,
    ZORDER_CMD_MAP,
    place_in_zorder,
)
from src.protocol.definitions import tool


@tool(
    "wps_ppt_shapes",
    component="wpp",
    annotations={"title": "Select Shapes", "idempotentHint": True},
    operation="select",
)
def select_shapes(params: SelectShapesInput) -> JsonObject:
    """Resolve all named shapes before changing the current selection."""
    session = get_session("wpp")
    slide = session.get_document().Slides(params.slide_index)
    shapes = [_get_shape(slide, name) for name in params.shape_names]
    session.activate_window()
    goto_slide(session.get_application(), params.slide_index)
    for index, shape in enumerate(shapes):
        shape.Select(MSO_TRUE if index == 0 else MSO_FALSE)
    return {
        "success": True,
        "slide_index": params.slide_index,
        "selected_shapes": params.shape_names,
        "count": len(shapes),
    }


@tool(
    "wps_ppt_shapes",
    component="wpp",
    annotations={"title": "Lock Aspect Ratio", "idempotentHint": True},
    operation="lock_aspect_ratio",
)
def lock_aspect_ratio(params: LockAspectRatioInput) -> JsonObject:
    """Lock or unlock the aspect ratio of a shape."""
    app = get_session("wpp").get_application()
    goto_slide(app, params.slide_index)
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    shape = _get_shape(slide, params.shape_name_or_index)
    shape.LockAspectRatio = MSO_TRUE if params.locked else MSO_FALSE
    result = {
        "success": True,
        "shape_name": shape.Name,
        "locked": params.locked,
    }
    return result


@tool(
    "wps_ppt_shapes",
    component="wpp",
    annotations={"title": "Copy Shape to Slide"},
    operation="copy",
)
def copy_shape_to_slide(params: CopyShapeToSlideInput) -> JsonObject:
    """Copy a shape from one slide to another slide.

    Copies the shape via the clipboard and pastes it onto the destination
    slide. The original shape remains on the source slide. Identify the
    shape by name (string) or 1-based index (int).
    """
    pres = get_session("wpp").get_document()
    src_slide = pres.Slides(params.src_slide_index)
    shape = _get_shape(src_slide, params.shape_name_or_index)
    dst_slide = pres.Slides(params.dst_slide_index)
    shape.Copy()
    pasted = dst_slide.Shapes.Paste()
    new_shape = pasted(1)
    return {
        "success": True,
        "new_shape_name": new_shape.Name,
        "destination_slide": params.dst_slide_index,
    }


@tool("wps_ppt_shapes", component="wpp", operation="group")
def group_shapes(params: GroupShapesInput) -> JsonObject:
    """Replace at least two top-level shapes with a single group."""
    session = get_session("wpp")
    slide = session.get_document().Slides(params.slide_index)
    require_top_level(
        slide, params.shape_names, params.slide_index, "wps_ppt_shapes/group"
    )
    goto_slide(session.get_application(), params.slide_index)
    group = slide.Shapes.Range(tuple(params.shape_names)).Group()
    return {
        "status": "success",
        "group_name": group.Name,
        "shape_index": group.ZOrderPosition,
    }


@tool(
    "wps_ppt_shapes",
    component="wpp",
    annotations={"destructiveHint": True},
    operation="ungroup",
)
def ungroup_shapes(params: ShapeInput) -> JsonObject:
    """Replace a group with its child shapes and return their names."""
    session = get_session("wpp")
    slide = session.get_document().Slides(params.slide_index)
    shape = resolve_shape(slide, params.shape_name_or_index)
    if shape.Type != MSO_GROUP:
        raise ValueError(f"Shape '{shape.Name}' is not a group.")
    goto_slide(session.get_application(), params.slide_index)
    ungrouped = shape.Ungroup()
    return {
        "status": "success",
        "ungrouped_count": ungrouped.Count,
        "shape_names": [
            ungrouped(index).Name for index in range(1, ungrouped.Count + 1)
        ],
    }


@tool(
    "wps_ppt_shapes",
    component="wpp",
    annotations={"readOnlyHint": True},
    operation="group_items",
)
def get_group_items(params: ShapeInput) -> JsonObject:
    """Read child names, reusable group paths, types and geometry in points."""
    slide = get_session("wpp").get_document().Slides(params.slide_index)
    shape, group_path = resolve_with_path(slide, params.shape_name_or_index)
    if shape.Type != MSO_GROUP:
        raise ValueError(f"Shape '{shape.Name}' is not a group.")
    items: list[JsonObject] = []
    for index in range(1, shape.GroupItems.Count + 1):
        item = shape.GroupItems(index)
        items.append(
            {
                "name": item.Name,
                "path": group_path + PATH_SEPARATOR + item.Name,
                "type": item.Type,
                "type_name": SHAPE_TYPE_NAMES.get(
                    item.Type, f"Unknown({item.Type})"
                ),
                "left": round(item.Left, 2),
                "top": round(item.Top, 2),
                "width": round(item.Width, 2),
                "height": round(item.Height, 2),
            }
        )
    return {"status": "success", "group_name": shape.Name, "items": items}


logger = logging.getLogger(__name__)


def _get_shape(
    slide: NativeObject,
    name_or_index: str | int | None,
    shape_name: str | None = None,
    shape_index: int | None = None,
) -> NativeObject:
    """Find a shape on a slide by name or 1-based index."""
    identifier: str | int
    if shape_name is not None:
        identifier = shape_name
    elif shape_index is not None:
        identifier = shape_index
    elif name_or_index is not None:
        identifier = name_or_index
    else:
        raise ValueError("Either shape_name or shape_index must be provided.")
    return resolve_shape(slide, identifier)


def _apply_text_style(
    shape: NativeObject, params: AddShapeInput | AddTextboxInput
) -> None:
    """Apply the same text and font properties to shapes and text boxes."""
    text_range = shape.TextFrame.TextRange
    if params.text is not None:
        text_range.Text = normalize_text(params.text)
    apply_font_format(
        text_range.Font,
        TextFormatSpec(
            font_name=params.font_name,
            font_size=params.font_size,
            bold=params.bold,
            italic=params.italic,
            color=params.font_color,
        ),
    )
    if params.align is not None:
        text_range.ParagraphFormat.Alignment = ALIGNMENT_MAP[params.align]


@tool(
    "wps_ppt_shapes",
    component="wpp",
    annotations={"title": "Add Shape"},
    operation="add",
)
def add_shape(params: AddShapeInput) -> JsonObject:
    """Add an auto shape to a slide (rectangle, oval, arrow, star, etc.)."""
    shape_type_int = resolve_shape_type(params.shape_type)
    app = get_session("wpp").get_application()
    pres = get_session("wpp").get_document()
    goto_slide(app, params.slide_index)
    slide = pres.Slides(params.slide_index)
    shape = slide.Shapes.AddShape(
        Type=shape_type_int,
        Left=params.left,
        Top=params.top,
        Width=params.width,
        Height=params.height,
    )
    _apply_text_style(shape, params)
    if (
        params.fill_color is not None
        or params.fill_type is not None
        or params.fill_transparency is not None
    ):
        effective_type = params.fill_type or (
            "solid" if params.fill_color is not None else None
        )
        fill = shape.Fill
        if effective_type == "none":
            fill.Visible = MSO_FALSE
        elif effective_type == "gradient":
            gstyle = GRADIENT_STYLE_MAP.get(
                params.fill_gradient_style or "horizontal", 1
            )
            fill.TwoColorGradient(Style=gstyle, Variant=1)
            if params.fill_color is not None:
                fill.ForeColor.RGB = hex_to_int(params.fill_color)
            if params.fill_color2 is not None:
                fill.BackColor.RGB = hex_to_int(params.fill_color2)
        elif effective_type == "solid":
            fill.Solid()
            if params.fill_color is not None:
                fill.ForeColor.RGB = hex_to_int(params.fill_color)
        if params.fill_transparency is not None and effective_type != "none":
            fill.Transparency = params.fill_transparency
    if params.line_visible is not None:
        shape.Line.Visible = MSO_TRUE if params.line_visible else MSO_FALSE
    if params.line_color is not None:
        if params.line_visible is None:
            shape.Line.Visible = MSO_TRUE
        shape.Line.ForeColor.RGB = hex_to_int(params.line_color)
    if params.line_weight is not None:
        shape.Line.Weight = params.line_weight
    if params.corner_radius is not None or params.corner_radius_pt is not None:
        try:
            if shape.AutoShapeType == SHAPE_NAME_MAP["rounded_rectangle"]:
                if params.corner_radius_pt is not None:
                    short_side = min(params.width, params.height)
                    adj_value = min(0.5, params.corner_radius_pt / short_side)
                else:
                    adj_value = (params.corner_radius or 0.0) * 0.5
                shape.Adjustments[1] = adj_value
        except Exception as error:
            raise RuntimeError(
                f"Failed to set corner_radius on shape '{shape.Name}': {error}"
            ) from error
    result = {
        "success": True,
        "shape_name": shape.Name,
        "shape_index": shape.ZOrderPosition,
        "shape_type": shape.AutoShapeType,
    }
    placed = place_in_zorder(slide, shape, params.zorder)
    if "z_position" in placed:
        result["shape_index"] = placed["z_position"]
    result.update(placed)
    warn = font_size_warning(params.font_size)
    if warn:
        result["warning"] = warn
    return result


@tool(
    "wps_ppt_shapes",
    component="wpp",
    annotations={"title": "Add Text Box"},
    operation="add_textbox",
)
def add_textbox(params: AddTextboxInput) -> JsonObject:
    """Add a text box to a slide."""
    app = get_session("wpp").get_application()
    goto_slide(app, params.slide_index)
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    textbox = slide.Shapes.AddTextbox(
        Orientation=MSO_TEXT_ORIENTATION_HORIZONTAL,
        Left=params.left,
        Top=params.top,
        Width=params.width,
        Height=params.height,
    )
    _apply_text_style(textbox, params)
    if params.vertical_anchor is not None:
        textbox.TextFrame.VerticalAnchor = VERTICAL_ANCHOR_MAP[
            params.vertical_anchor
        ]
    placed = place_in_zorder(slide, textbox, params.zorder)
    result = {
        "success": True,
        "shape_name": textbox.Name,
        "shape_index": textbox.ZOrderPosition,
        **placed,
    }
    warn = font_size_warning(params.font_size)
    if warn:
        result["warning"] = warn
    return result


@tool(
    "wps_ppt_shapes",
    component="wpp",
    annotations={"title": "Add Line"},
    operation="add_line",
)
def add_line(params: AddLineInput) -> JsonObject:
    """Add a straight line to a slide."""
    app = get_session("wpp").get_application()
    goto_slide(app, params.slide_index)
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    line = slide.Shapes.AddLine(
        BeginX=params.begin_x,
        BeginY=params.begin_y,
        EndX=params.end_x,
        EndY=params.end_y,
    )
    placed = place_in_zorder(slide, line, params.zorder)
    result = {
        "success": True,
        "shape_name": line.Name,
        "shape_index": line.ZOrderPosition,
        **placed,
    }
    return result


@tool(
    "wps_ppt_shapes",
    component="wpp",
    annotations={"title": "Update Shape"},
    operation="update",
)
def update_shape(params: UpdateShapeInput) -> JsonObject:
    """Update each resolved shape once after preflighting all geometry."""
    session = get_session("wpp")
    slide = session.get_document().Slides(params.slide_index)
    is_batch = params.all or params.shape_names is not None
    if is_batch:
        available = [
            slide.Shapes(index) for index in range(1, slide.Shapes.Count + 1)
        ]
        names = [shape.Name for shape in available]
        targets = [
            available[pick]
            if isinstance(pick, int)
            else resolve_shape(slide, pick)
            for pick in select_targets(
                names, params.shape_names, params.all, params.exclude
            )
        ]
        targets = list({shape.Id: shape for shape in targets}.values())
    else:
        targets = [
            _get_shape(slide, None, params.shape_name, params.shape_index)
        ]
    adjustment_count = 0
    if params.adjustments:
        try:
            adjustment_count = targets[0].Adjustments.Count
        except Exception as error:
            raise ValueError(
                "This shape does not support adjustment handles"
            ) from error
        for index in params.adjustments:
            if not 1 <= int(index) <= adjustment_count:
                raise ValueError(
                    f"Adjustment index {index} out of range "
                    f"(1-{adjustment_count})"
                )
    for shape in targets:
        for dimension in ("width", "height"):
            absolute = getattr(params, dimension)
            final = (
                absolute
                if absolute is not None
                else getattr(shape, dimension.title())
            )
            final += getattr(params, "d" + dimension) or 0
            if (
                absolute is not None
                or getattr(params, "d" + dimension) is not None
            ) and final < 0:
                raise ValueError(
                    f"Resulting {dimension} for '{shape.Name}' cannot "
                    f"be negative"
                )
    goto_slide(session.get_application(), params.slide_index)
    updated: list[JsonObject] = []
    for shape in targets:
        apply_geometry(shape, params)
        if params.name is not None:
            shape.Name = params.name
        if params.adjustments:
            for index, value in params.adjustments.items():
                shape.Adjustments[int(index)] = value
        updated.append(
            {
                "shape_name": shape.Name,
                "left": round(shape.Left, 2),
                "top": round(shape.Top, 2),
                "width": round(shape.Width, 2),
                "height": round(shape.Height, 2),
            }
        )
    if is_batch:
        return {"success": True, "count": len(updated), "updated": updated}
    result = {"success": True, **updated[0]}
    if params.adjustments:
        result["adjustments"] = {
            str(index): round(targets[0].Adjustments[index], 4)
            for index in range(1, adjustment_count + 1)
        }
        labels = ADJUSTMENT_LABELS.get(targets[0].AutoShapeType)
        if labels:
            result["adjustment_labels"] = {
                str(index): label for index, label in labels.items()
            }
    return result


@tool(
    "wps_ppt_shapes",
    component="wpp",
    annotations={"title": "Delete Shape", "destructiveHint": True},
    operation="delete",
)
def delete_shape(params: ShapeIdentifierInput) -> JsonObject:
    """Delete a shape from a slide."""
    app = get_session("wpp").get_application()
    goto_slide(app, params.slide_index)
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    shape = _get_shape(
        slide,
        None,
        shape_name=params.shape_name,
        shape_index=params.shape_index,
    )
    deleted_name = shape.Name
    shape.Delete()
    result = {"success": True, "deleted": deleted_name}
    return result


@tool(
    "wps_ppt_shapes",
    component="wpp",
    annotations={"title": "Duplicate Shape"},
    operation="duplicate",
)
def duplicate_shape(params: ShapeIdentifierInput) -> JsonObject:
    """Duplicate a shape on the same slide."""
    app = get_session("wpp").get_application()
    goto_slide(app, params.slide_index)
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    shape = _get_shape(
        slide,
        None,
        shape_name=params.shape_name,
        shape_index=params.shape_index,
    )
    dup = shape.Duplicate()
    new_shape = dup(1)
    new_shape.Left = shape.Left + 20
    new_shape.Top = shape.Top + 20
    result = {
        "success": True,
        "new_shape_name": new_shape.Name,
        "new_shape_index": new_shape.ZOrderPosition,
    }
    return result


@tool(
    "wps_ppt_shapes",
    component="wpp",
    annotations={"title": "Set Shape Z-Order", "idempotentHint": True},
    operation="set_zorder",
)
def set_shape_zorder(params: SetZOrderInput) -> JsonObject:
    """Change stacking order, optionally placing artwork below all text."""
    command = (
        params.command.strip().lower().replace(" ", "_").replace("-", "_")
    )
    if command != BEHIND_TEXT and command not in ZORDER_CMD_MAP:
        raise ValueError(f"Unknown z-order command '{params.command}'")
    session = get_session("wpp")
    slide = session.get_document().Slides(params.slide_index)
    shape = _get_shape(slide, None, params.shape_name, params.shape_index)
    goto_slide(session.get_application(), params.slide_index)
    placement = {}
    if command == BEHIND_TEXT:
        placement = place_in_zorder(slide, shape, "behind_text")
    else:
        shape.ZOrder(ZORDER_CMD_MAP[command])
    result = {
        "success": True,
        "shape_name": shape.Name,
        "new_z_position": shape.ZOrderPosition,
    }
    if "note" in placement:
        result["note"] = placement["note"]
    return result
