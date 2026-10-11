"""Measure rendered text and compare it with available frame dimensions."""

from __future__ import annotations

from src.models import JsonObject, NativeObject
from src.presentation.text_frame import text_frame_state

FIT_TOLERANCE_PT = 0.5


_VERTICAL_FLOW = frozenset({"vertical", "upward", "downward"})


def build_measurement(
    lines: list[JsonObject],
    text_width: float | None,
    text_height: float | None,
    shape_width: float,
    shape_height: float,
    margins: dict[str, float] | None,
    word_wrap: bool | None,
    autofit: str | None,
    orientation: str | None,
) -> JsonObject:
    """Turn measured bounds into the answer a caller asked the question for."""
    measurement = {
        "line_count": len(lines),
        "lines": lines,
        "text_width_pt": text_width,
        "text_height_pt": text_height,
        "usable_width_pt": None,
        "usable_height_pt": None,
        "overflows": None,
        "autofit": autofit,
    }
    if margins is None:
        return measurement
    usable_width = round(shape_width - margins["left"] - margins["right"], 2)
    usable_height = round(shape_height - margins["top"] - margins["bottom"], 2)
    measurement["usable_width_pt"] = usable_width
    measurement["usable_height_pt"] = usable_height
    if orientation in _VERTICAL_FLOW:
        checks = [(text_width, usable_width)]
        if word_wrap is False:
            checks.append((text_height, usable_height))
    else:
        checks = [(text_height, usable_height)]
        if word_wrap is False:
            checks.append((text_width, usable_width))
    if any((measured is None for measured, _ in checks)):
        return measurement
    measurement["overflows"] = any(
        (
            measured is not None and measured > limit + FIT_TOLERANCE_PT
            for measured, limit in checks
        )
    )
    return measurement


def measure_text(shape: NativeObject, tr: NativeObject) -> JsonObject:
    """Measure a text frame as WPS Presentation currently draws it."""
    state = text_frame_state(shape) or {}
    lines = []
    try:
        line_count = tr.Lines().Count
    except Exception:
        line_count = 0
    for i in range(1, line_count + 1):
        line = tr.Lines(i)
        entry = {"text": None, "width_pt": None, "height_pt": None}
        try:
            entry["text"] = line.Text
        except Exception:
            pass
        try:
            entry["width_pt"] = round(line.BoundWidth, 2)
        except Exception:
            pass
        try:
            entry["height_pt"] = round(line.BoundHeight, 2)
        except Exception:
            pass
        lines.append(entry)
    text_width = text_height = None
    try:
        text_width = round(tr.BoundWidth, 2)
        text_height = round(tr.BoundHeight, 2)
    except Exception:
        pass
    return build_measurement(
        lines,
        text_width,
        text_height,
        round(shape.Width, 2),
        round(shape.Height, 2),
        state.get("margins"),
        state.get("word_wrap"),
        state.get("autofit"),
        state.get("orientation"),
    )
