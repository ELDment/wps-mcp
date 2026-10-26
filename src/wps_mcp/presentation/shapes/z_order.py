"""Place native shapes relative to artwork and text in the slide stack."""

from __future__ import annotations

from wps_mcp.models import JsonObject, NativeObject
from wps_mcp.native.objects import is_optional_member_error
from wps_mcp.presentation.constants import (
    MSO_BRING_FORWARD,
    MSO_BRING_TO_FRONT,
    MSO_SEND_BACKWARD,
    MSO_SEND_TO_BACK,
)
from wps_mcp.presentation.shapes.lookup import walk_group_children

ZORDER_CMD_MAP: dict[str, int] = {
    "bring_to_front": MSO_BRING_TO_FRONT,
    "send_to_back": MSO_SEND_TO_BACK,
    "bring_forward": MSO_BRING_FORWARD,
    "send_backward": MSO_SEND_BACKWARD,
}


ZORDER_FIELD_DESCRIPTION = (
    "Where the new shape lands in the stack. 'front' (default) is"
    " what WPS Presentation does. 'behind_text' puts it directly "
    "below the lowest shape carrying text, which is what art "
    "under a caption wants and what 'back' gets wrong on a deck "
    "with a full bleed background. 'back' is the very bottom."
)


ZORDER_PLACEMENTS = ("front", "back", "behind_text")


BEHIND_TEXT = "send_behind_text"


def _carries_text(shape: NativeObject) -> bool:
    """True when this shape, or anything inside it, has text on it."""
    try:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            return True
    except Exception as error:
        if not is_optional_member_error(error):
            raise
        pass
    for child, _ in walk_group_children(shape):
        try:
            if child.HasTextFrame and child.TextFrame.HasText:
                return True
        except Exception as error:
            if not is_optional_member_error(error):
                raise
            continue
    return False


def _text_positions(slide: NativeObject, shape: NativeObject) -> list[int]:
    positions = []
    for i in range(1, slide.Shapes.Count + 1):
        other = slide.Shapes(i)
        if other.Name == shape.Name:
            continue
        if _carries_text(other):
            positions.append(other.ZOrderPosition)
    return positions


def place_in_zorder(
    slide: NativeObject, shape: NativeObject, where: str | None
) -> JsonObject:
    """Put a freshly added shape where the caller asked for it."""
    if where in (None, "front"):
        return {}
    if where == "back":
        shape.ZOrder(MSO_SEND_TO_BACK)
        return {"zorder": "back", "z_position": shape.ZOrderPosition}
    if where != "behind_text":
        raise ValueError(
            f"Unknown zorder '{where}'. Use one of: "
            f"{', '.join(ZORDER_PLACEMENTS)}"
        )
    if not _text_positions(slide, shape):
        return {
            "zorder": "front",
            "z_position": shape.ZOrderPosition,
            "note": (
                "zorder was behind_text and nothing on this slide has text, "
                "so the shape was left at the front rather than hidden under "
                "the background."
            ),
        }
    shape.ZOrder(MSO_SEND_TO_BACK)
    for _ in range(slide.Shapes.Count):
        lowest = min(_text_positions(slide, shape))
        if shape.ZOrderPosition + 1 >= lowest:
            break
        shape.ZOrder(MSO_BRING_FORWARD)
    return {"zorder": "behind_text", "z_position": shape.ZOrderPosition}
