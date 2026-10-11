"""Read native text frame state without depending on tool entry points."""

from __future__ import annotations

from src.models import JsonObject, NativeObject
from src.presentation.constants import MSO_TRI_STATE_MIXED, MSO_TRUE
from src.presentation.text_enums import (
    AUTO_SIZE_NAMES,
    ORIENTATION_NAMES,
    VERTICAL_ANCHOR_NAMES,
)


def text_frame_state(shape: NativeObject) -> JsonObject | None:
    """Report the text frame settings that decide how text is drawn."""
    try:
        if not shape.HasTextFrame:
            return None
    except Exception:
        return None
    state: JsonObject = {
        "autofit": None,
        "word_wrap": None,
        "vertical_anchor": None,
        "orientation": None,
        "margins": None,
    }
    try:
        state["autofit"] = AUTO_SIZE_NAMES.get(shape.TextFrame2.AutoSize)
    except Exception:
        pass
    try:
        tf = shape.TextFrame
    except Exception:
        return state
    try:
        wrap = tf.WordWrap
        state["word_wrap"] = (
            None if wrap == MSO_TRI_STATE_MIXED else wrap == MSO_TRUE
        )
    except Exception:
        pass
    try:
        state["vertical_anchor"] = VERTICAL_ANCHOR_NAMES.get(tf.VerticalAnchor)
    except Exception:
        pass
    try:
        state["orientation"] = ORIENTATION_NAMES.get(tf.Orientation)
    except Exception:
        pass
    try:
        state["margins"] = {
            "left": round(tf.MarginLeft, 2),
            "right": round(tf.MarginRight, 2),
            "top": round(tf.MarginTop, 2),
            "bottom": round(tf.MarginBottom, 2),
        }
    except Exception:
        pass
    return state
