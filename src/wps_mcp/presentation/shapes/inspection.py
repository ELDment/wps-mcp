"""Read shape geometry, appearance, text state and native metadata."""

from __future__ import annotations

from wps_mcp.common.colors import int_to_hex
from wps_mcp.models import JsonObject, NativeObject
from wps_mcp.native.objects import is_optional_member_error, optional_members
from wps_mcp.native.runtime import get_session
from wps_mcp.presentation.constants import (
    DASH_STYLE_NAMES,
    MSO_GROUP,
    MSO_TRUE,
    SHAPE_TYPE_NAMES,
)
from wps_mcp.presentation.shapes.lookup import resolve_shape
from wps_mcp.presentation.shapes.models import (
    ListShapesInput,
    ShapeIdentifierInput,
)
from wps_mcp.presentation.shapes.types import ADJUSTMENT_LABELS
from wps_mcp.presentation.text.frame import text_frame_state
from wps_mcp.protocol.definitions import tool

_UNDEFINED = -2147483648


def _glow_of(source: NativeObject) -> JsonObject | None:
    """radius, colour and transparency of one Glow object, or None."""
    try:
        radius = round(source.Radius, 2)
        radius = None if radius <= _UNDEFINED else radius
    except Exception as error:
        if not is_optional_member_error(error):
            raise
        return None
    glow = {"radius": radius, "color_hex": None, "transparency": None}
    if not radius:
        return glow
    with optional_members():
        glow["color_hex"] = int_to_hex(source.Color.RGB)
    with optional_members():
        transparency = round(source.Transparency, 2)
        glow["transparency"] = (
            None if transparency <= _UNDEFINED else transparency
        )
    return glow


def _glows(shape: NativeObject) -> JsonObject:
    """The shape's glow and the glow on its text, side by side."""
    glows: JsonObject = {"glow": None, "text_glow": None}
    with optional_members():
        glows["glow"] = _glow_of(shape.Glow)
    with optional_members():
        if shape.HasTextFrame:
            glows["text_glow"] = _glow_of(shape.TextFrame2.TextRange.Font.Glow)
    return glows


def _fill_info(shape: NativeObject) -> JsonObject | None:
    """Read available fill properties without dropping readable siblings."""
    result = None
    with optional_members():
        fill = shape.Fill
        result = {"type": fill.Type, "visible": bool(fill.Visible)}
        with optional_members():
            result["color_hex"] = int_to_hex(fill.ForeColor.RGB)
        with optional_members():
            result["transparency"] = round(fill.Transparency, 2)
    return result


def _line_info(shape: NativeObject) -> JsonObject | None:
    """Read available outline properties independently."""
    result: JsonObject | None = None
    with optional_members():
        line = shape.Line
        result = {"visible": bool(line.Visible)}
        with optional_members():
            result["weight"] = round(line.Weight, 2)
        with optional_members():
            result["color_hex"] = int_to_hex(line.ForeColor.RGB)
        with optional_members():
            dash = line.DashStyle
            result["dash_style"] = DASH_STYLE_NAMES.get(dash, dash)
        with optional_members():
            result["transparency"] = round(line.Transparency, 2)
    return result


def _connector_info(shape: NativeObject) -> JsonObject:
    """Read connected endpoints without probing non-connector shapes."""
    result: JsonObject = {}
    with optional_members():
        if not shape.Connector:
            return result
        connector = shape.ConnectorFormat
        with optional_members():
            if connector.BeginConnected:
                result["begin_connected_shape"] = (
                    connector.BeginConnectedShape.Name
                )
                result["begin_connection_site"] = connector.BeginConnectionSite
        with optional_members():
            if connector.EndConnected:
                result["end_connected_shape"] = (
                    connector.EndConnectedShape.Name
                )
                result["end_connection_site"] = connector.EndConnectionSite
    return result


def _adjustment_info(shape: NativeObject) -> JsonObject:
    """Read available geometry adjustments and their shape-specific labels."""
    result: JsonObject = {}
    with optional_members():
        count = shape.Adjustments.Count
        if count <= 0:
            return result
        values = {}
        for index in range(1, count + 1):
            with optional_members():
                values[str(index)] = round(shape.Adjustments[index], 4)
        result.update(adjustments=values, adjustments_count=len(values))
        with optional_members():
            labels = ADJUSTMENT_LABELS.get(shape.AutoShapeType)
            if labels:
                result["adjustment_labels"] = {
                    str(index): label for index, label in labels.items()
                }
    return result


@tool(
    "wps_ppt_shapes",
    component="wpp",
    annotations={
        "title": "List Shapes",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="list",
)
def list_shapes(params: ListShapesInput) -> JsonObject:
    """List all shapes on a slide."""
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    shapes = []
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        has_text = False
        text_preview = ""
        with optional_members():
            if shape.HasTextFrame:
                has_text = True
                if shape.TextFrame.HasText:
                    full_text = shape.TextFrame.TextRange.Text
                    text_preview = full_text[:50] + (
                        "..." if len(full_text) > 50 else ""
                    )
        shapes.append(
            {
                "index": i,
                "name": shape.Name,
                "id": shape.Id,
                "type": shape.Type,
                "type_name": SHAPE_TYPE_NAMES.get(
                    shape.Type, f"Unknown({shape.Type})"
                ),
                "left": round(shape.Left, 2),
                "top": round(shape.Top, 2),
                "width": round(shape.Width, 2),
                "height": round(shape.Height, 2),
                "has_text": has_text,
                "text_preview": text_preview,
            }
        )
    result = {
        "slide_index": params.slide_index,
        "shapes_count": slide.Shapes.Count,
        "shapes": shapes,
    }
    return result


@tool(
    "wps_ppt_shapes",
    component="wpp",
    annotations={
        "title": "Get Shape Info",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="get",
)
def get_shape_info(params: ShapeIdentifierInput) -> JsonObject:
    """Get detailed information about a specific shape."""
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    identifier = (
        params.shape_name
        if params.shape_name is not None
        else params.shape_index
    )
    assert identifier is not None
    shape = resolve_shape(slide, identifier)
    info: JsonObject = {
        "name": shape.Name,
        "id": shape.Id,
        "type": shape.Type,
        "type_name": SHAPE_TYPE_NAMES.get(
            shape.Type, f"Unknown({shape.Type})"
        ),
        "left": round(shape.Left, 2),
        "top": round(shape.Top, 2),
        "width": round(shape.Width, 2),
        "height": round(shape.Height, 2),
        "rotation": round(shape.Rotation, 2),
        "z_order": shape.ZOrderPosition,
        "is_group": shape.Type == MSO_GROUP,
        "has_animation": False,
        "aspect_ratio_locked": False,
        "text": None,
        "fill": None,
        "line": None,
        "text_frame": text_frame_state(shape),
        **_glows(shape),
    }
    with optional_members():
        seq = slide.TimeLine.MainSequence
        for i in range(1, seq.Count + 1):
            if seq(i).Shape.Id == shape.Id:
                info["has_animation"] = True
                break
    with optional_members():
        info["aspect_ratio_locked"] = shape.LockAspectRatio == MSO_TRUE
    with optional_members():
        if shape.HasTextFrame and shape.TextFrame.HasText:
            info["text"] = shape.TextFrame.TextRange.Text
    info["fill"] = _fill_info(shape)
    info["line"] = _line_info(shape)
    connector = _connector_info(shape)
    if connector:
        info["connector_format"] = connector
    info.update(_adjustment_info(shape))
    return info
