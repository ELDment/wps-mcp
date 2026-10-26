"""Read native text frame state without depending on tool entry points."""

from __future__ import annotations

from wps_mcp.models import JsonObject, NativeObject
from wps_mcp.native.objects import is_optional_member_error, optional_members
from wps_mcp.presentation.constants import MSO_TRI_STATE_MIXED, MSO_TRUE
from wps_mcp.presentation.text.enums import (
    AUTO_SIZE_NAMES,
    ORIENTATION_NAMES,
    VERTICAL_ANCHOR_NAMES,
)


def text_frame_state(shape: NativeObject) -> JsonObject | None:
    """Report the text frame settings that decide how text is drawn."""
    try:
        if not shape.HasTextFrame:
            return None
    except Exception as error:
        if not is_optional_member_error(error):
            raise
        return None
    state: JsonObject = {
        "autofit": None,
        "word_wrap": None,
        "vertical_anchor": None,
        "orientation": None,
        "margins": None,
    }
    with optional_members():
        state["autofit"] = AUTO_SIZE_NAMES.get(shape.TextFrame2.AutoSize)
    try:
        tf = shape.TextFrame
    except Exception as error:
        if not is_optional_member_error(error):
            raise
        return state
    with optional_members():
        wrap = tf.WordWrap
        state["word_wrap"] = (
            None if wrap == MSO_TRI_STATE_MIXED else wrap == MSO_TRUE
        )
    with optional_members():
        state["vertical_anchor"] = VERTICAL_ANCHOR_NAMES.get(tf.VerticalAnchor)
    with optional_members():
        state["orientation"] = ORIENTATION_NAMES.get(tf.Orientation)
    with optional_members():
        state["margins"] = {
            "left": round(tf.MarginLeft, 2),
            "right": round(tf.MarginRight, 2),
            "top": round(tf.MarginTop, 2),
            "bottom": round(tf.MarginBottom, 2),
        }
    return state
