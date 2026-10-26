"""Apply validated character formatting to native presentation text."""

from __future__ import annotations

from wps_mcp.common.colors import get_theme_color_index, hex_to_int
from wps_mcp.common.text import font_size_warning
from wps_mcp.models import JsonObject, NativeObject
from wps_mcp.presentation.constants import MSO_FALSE, MSO_TRUE
from wps_mcp.presentation.text.models import TextFormatSpec


def apply_font_format(font: NativeObject, params: TextFormatSpec) -> None:
    """Apply shared font properties to an already resolved native font."""
    if params.font_name is not None:
        font.Name = params.font_name
        font.NameFarEast = params.font_name
    if params.font_name_fareast is not None:
        font.NameFarEast = params.font_name_fareast
    if params.font_size is not None:
        font.Size = params.font_size
    for field, attribute in (
        ("bold", "Bold"),
        ("italic", "Italic"),
        ("underline", "Underline"),
    ):
        value = getattr(params, field)
        if value is not None:
            setattr(font, attribute, MSO_TRUE if value else MSO_FALSE)
    if params.color is not None:
        font.Color.RGB = hex_to_int(params.color)
    if params.font_color_theme is not None:
        font.Color.ObjectThemeColor = get_theme_color_index(
            params.font_color_theme
        )


def apply_text_format(
    shape: NativeObject,
    params: TextFormatSpec,
    start: int | None = None,
    length: int | None = None,
) -> JsonObject:
    """Format one native UTF-16 span or the entire text frame."""
    if not shape.HasTextFrame:
        raise ValueError(f"Shape {shape.Name!r} does not have a text frame")
    text_range = shape.TextFrame.TextRange
    target = (
        text_range
        if start is None
        else text_range.Characters(Start=start, Length=length)
    )
    highlight = None
    if params.highlight_color is not None:
        target2 = shape.TextFrame2.TextRange
        if start is not None:
            target2 = target2.Characters(start, length)
        highlight = target2.Font.Highlight
    apply_font_format(target.Font, params)
    if highlight is not None and params.highlight_color is not None:
        highlight.RGB = hex_to_int(params.highlight_color)
    result: JsonObject = {
        "success": True,
        "shape_name": shape.Name,
        "formatted_text": target.Text,
        "start": target.Start,
        "length": target.Length,
        "position_unit": "UTF-16",
    }
    warning = font_size_warning(params.font_size)
    if warning:
        result["warnings"] = [warning]
    return result
