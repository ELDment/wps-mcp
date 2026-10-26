"""Connect to WPS Presentation and inspect its application state."""

from __future__ import annotations

import logging
from typing import Literal

from pydantic import Field

from wps_mcp.errors import native_errors
from wps_mcp.models import DocumentInput, InputModel, JsonObject
from wps_mcp.native.runtime import get_session
from wps_mcp.presentation.constants import (
    PP_SELECTION_NONE,
    PP_SELECTION_SHAPES,
    PP_SELECTION_SLIDES,
    PP_SELECTION_TEXT,
    VIEW_TYPE_MAP,
    VIEW_TYPE_NAMES,
    WINDOW_STATE_NAMES,
)
from wps_mcp.protocol.definitions import tool


class SetViewInput(DocumentInput):
    """Input for setting the WPS Presentation view type and zoom."""

    view_type: str | None = Field(
        default=None,
        description="WPS view name such as "
        "normal, slide_sorter, or slide_master.",
    )
    zoom: int | None = Field(
        default=None, ge=10, le=400, description="Zoom level (10-400)"
    )


@tool(
    "wps_ppt_app",
    component="wpp",
    annotations={
        "title": "Get Selection",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="selection",
)
def get_selection(params: DocumentInput) -> JsonObject:
    """Read the current selection in the target window."""
    get_session("wpp").get_application()
    window = get_session("wpp").get_window()
    if window is None:
        raise RuntimeError(
            "There is no document window to read a selection from (the "
            "presentation was opened with with_window=False?)."
        )
    selection = window.Selection
    sel_type = selection.Type
    result = {"success": True, "type": sel_type}
    if sel_type == PP_SELECTION_NONE:
        result["type_name"] = "none"
    elif sel_type == PP_SELECTION_SLIDES:
        result["type_name"] = "slides"
        slide_indices = []
        for i in range(1, selection.SlideRange.Count + 1):
            slide_indices.append(selection.SlideRange(i).SlideIndex)
        result["slide_indices"] = slide_indices
    elif sel_type == PP_SELECTION_SHAPES:
        result["type_name"] = "shapes"
        shape_names = []
        for i in range(1, selection.ShapeRange.Count + 1):
            shape_names.append(selection.ShapeRange(i).Name)
        result["shape_names"] = shape_names
        result["count"] = selection.ShapeRange.Count
    elif sel_type == PP_SELECTION_TEXT:
        result["type_name"] = "text"
        result["text"] = selection.TextRange.Text
    return result


@tool(
    "wps_ppt_app",
    component="wpp",
    annotations={"title": "Set View", "idempotentHint": True},
    operation="set_view",
)
def set_view(params: SetViewInput) -> JsonObject:
    """Set the WPS Presentation view type and/or zoom level."""
    window = get_session("wpp").get_window()
    if window is None:
        raise RuntimeError("No WPS Presentation window is available.")
    if params.view_type is not None:
        vt_key = (
            params.view_type.strip()
            .lower()
            .replace(" ", "_")
            .replace("-", "_")
        )
        if vt_key not in VIEW_TYPE_MAP:
            raise ValueError(f"Unknown view_type '{params.view_type}'")
        window.ViewType = VIEW_TYPE_MAP[vt_key]
    if params.zoom is not None:
        window.View.Zoom = params.zoom
    current_view_type = window.ViewType
    current_zoom = window.View.Zoom
    result = {
        "success": True,
        "view_type": VIEW_TYPE_NAMES.get(
            current_view_type, f"Unknown({current_view_type})"
        ),
        "view_type_id": current_view_type,
        "zoom": current_zoom,
    }
    return result


logger = logging.getLogger(__name__)
SELECTION_TYPE_NAMES = {
    PP_SELECTION_NONE: "none",
    PP_SELECTION_SLIDES: "slides",
    PP_SELECTION_SHAPES: "shapes",
    PP_SELECTION_TEXT: "text",
}


class SetWindowStateInput(InputModel):
    """Choose the application window state."""

    window_state: Literal["normal", "minimized", "maximized"] = Field(
        default="maximized"
    )


@tool(
    "wps_ppt_app",
    component="wpp",
    annotations={"readOnlyHint": True},
    operation="info",
)
def get_app_info(params: InputModel) -> JsonObject:
    """Read version, visibility, window state and deck count."""
    session = get_session("wpp")
    app = session.get_application()
    result: JsonObject = {
        "name": app.Name,
        "version": app.Version,
        "visible": bool(app.Visible),
        "window_state": WINDOW_STATE_NAMES.get(app.WindowState, "unknown"),
        "presentations_count": app.Presentations.Count,
        "active_presentation": None,
    }
    if app.Presentations.Count:
        try:
            result["active_presentation"] = session.get_document().Name
        except (ValueError, RuntimeError) as error:
            result["selection_warning"] = str(error)
    return result


@tool(
    "wps_ppt_app",
    component="wpp",
    annotations={"readOnlyHint": True},
    operation="window",
)
def get_active_window_info(params: DocumentInput) -> JsonObject:
    """Read the target window, current slide and selection."""
    session = get_session("wpp")
    presentation = session.get_document(params.document)
    if presentation.Windows.Count == 0:
        raise RuntimeError("The selected presentation has no open window")
    window = presentation.Windows(1)
    result: JsonObject = {
        "caption": window.Caption,
        "view_type": window.ViewType,
        "active_slide_index": None,
        "selection_type": "none",
        "selected_shapes": [],
        "selected_text": None,
    }
    try:
        result["active_slide_index"] = window.View.Slide.SlideIndex
    except native_errors() + (AttributeError,) as error:
        logger.debug("Current view does not expose a slide: %s", error)
    selection = window.Selection
    result["selection_type"] = SELECTION_TYPE_NAMES.get(
        selection.Type, "unknown"
    )
    if selection.Type == PP_SELECTION_SHAPES:
        for index in range(1, selection.ShapeRange.Count + 1):
            shape = selection.ShapeRange(index)
            result["selected_shapes"].append(
                {"name": shape.Name, "type": shape.Type, "id": shape.Id}
            )
    elif selection.Type == PP_SELECTION_TEXT:
        result["selected_text"] = selection.TextRange.Text
    return result


@tool(
    "wps_ppt_app",
    component="wpp",
    annotations={"idempotentHint": True},
    operation="set_window",
)
def set_window_state(params: SetWindowStateInput) -> JsonObject:
    """Restore, minimize or maximize WPS Presentation."""
    states = {"normal": 1, "minimized": 2, "maximized": 3}
    get_session("wpp").get_application().WindowState = states[
        params.window_state
    ]
    return {"success": True, "window_state": params.window_state}
