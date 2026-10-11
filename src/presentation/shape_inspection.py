"""Read shape geometry, appearance, text state and native metadata."""

from __future__ import annotations

from src.common.colors import int_to_hex
from src.models import JsonObject, NativeObject
from src.native.runtime import get_session
from src.presentation.constants import (
    DASH_STYLE_NAMES,
    MSO_GROUP,
    MSO_TRUE,
    SHAPE_TYPE_NAMES,
)
from src.presentation.shape_lookup import resolve_shape
from src.presentation.shape_models import (
    ListShapesInput,
    ShapeIdentifierInput,
)
from src.presentation.shape_types import ADJUSTMENT_LABELS
from src.presentation.text_frame import text_frame_state
from src.protocol.definitions import tool

_UNDEFINED = -2147483648


def _glow_of(source: NativeObject) -> JsonObject | None:
    """radius, colour and transparency of one Glow object, or None."""
    try:
        radius = round(source.Radius, 2)
        radius = None if radius <= _UNDEFINED else radius
    except Exception:
        return None
    glow = {"radius": radius, "color_hex": None, "transparency": None}
    if not radius:
        return glow
    try:
        glow["color_hex"] = int_to_hex(source.Color.RGB)
    except Exception:
        pass
    try:
        transparency = round(source.Transparency, 2)
        glow["transparency"] = (
            None if transparency <= _UNDEFINED else transparency
        )
    except Exception:
        pass
    return glow


def _glows(shape: NativeObject) -> JsonObject:
    """The shape's glow and the glow on its text, side by side."""
    glows: JsonObject = {"glow": None, "text_glow": None}
    try:
        glows["glow"] = _glow_of(shape.Glow)
    except Exception:
        pass
    try:
        if shape.HasTextFrame:
            glows["text_glow"] = _glow_of(shape.TextFrame2.TextRange.Font.Glow)
    except Exception:
        pass
    return glows


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
        try:
            if shape.HasTextFrame:
                has_text = True
                if shape.TextFrame.HasText:
                    full_text = shape.TextFrame.TextRange.Text
                    text_preview = full_text[:50] + (
                        "..." if len(full_text) > 50 else ""
                    )
        except Exception:
            pass
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
    try:
        seq = slide.TimeLine.MainSequence
        for i in range(1, seq.Count + 1):
            if seq(i).Shape.Id == shape.Id:
                info["has_animation"] = True
                break
    except Exception:
        pass
    try:
        info["aspect_ratio_locked"] = shape.LockAspectRatio == MSO_TRUE
    except Exception:
        pass
    try:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            info["text"] = shape.TextFrame.TextRange.Text
    except Exception:
        pass
    try:
        fill = shape.Fill
        info["fill"] = {"type": fill.Type, "visible": bool(fill.Visible)}
        try:
            info["fill"]["color_hex"] = int_to_hex(fill.ForeColor.RGB)
        except Exception:
            pass
        try:
            info["fill"]["transparency"] = round(fill.Transparency, 2)
        except Exception:
            pass
    except Exception:
        pass
    try:
        line = shape.Line
        line_info: JsonObject = {"visible": bool(line.Visible)}
        info["line"] = line_info
        try:
            line_info["weight"] = round(line.Weight, 2)
        except Exception:
            pass
        try:
            line_info["color_hex"] = int_to_hex(line.ForeColor.RGB)
        except Exception:
            pass
        try:
            dash = line.DashStyle
            line_info["dash_style"] = DASH_STYLE_NAMES.get(dash, dash)
        except Exception:
            pass
        try:
            line_info["transparency"] = round(line.Transparency, 2)
        except Exception:
            pass
    except Exception:
        pass
    try:
        cf = shape.ConnectorFormat
        conn_info = {}
        try:
            if cf.BeginConnected:
                conn_info["begin_connected_shape"] = (
                    cf.BeginConnectedShape.Name
                )
                conn_info["begin_connection_site"] = cf.BeginConnectionSite
        except Exception:
            pass
        try:
            if cf.EndConnected:
                conn_info["end_connected_shape"] = cf.EndConnectedShape.Name
                conn_info["end_connection_site"] = cf.EndConnectionSite
        except Exception:
            pass
        if conn_info:
            info["connector_format"] = conn_info
    except Exception:
        pass
    try:
        adj_count = shape.Adjustments.Count
        if adj_count > 0:
            adj_dict = {}
            for i in range(1, adj_count + 1):
                try:
                    adj_dict[str(i)] = round(shape.Adjustments[i], 4)
                except Exception:
                    pass
            info["adjustments"] = adj_dict
            info["adjustments_count"] = len(adj_dict)
            try:
                auto_type = shape.AutoShapeType
                labels = ADJUSTMENT_LABELS.get(auto_type)
                if labels:
                    info["adjustment_labels"] = {
                        str(index): label for index, label in labels.items()
                    }
            except Exception:
                pass
    except Exception:
        pass
    result = info
    return result
