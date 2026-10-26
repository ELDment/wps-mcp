"""Resolve presentation targets before navigation or native mutation."""

from wps_mcp.models import NativeObject
from wps_mcp.native.runtime import get_session
from wps_mcp.presentation.navigation import goto_slide
from wps_mcp.presentation.shapes.lookup import resolve_shape
from wps_mcp.presentation.shapes.models import ShapeInput


def get_slide(
    slide_index: int, document: str | None = None, *, navigate: bool = False
) -> NativeObject:
    """Resolve an in-range slide in the explicitly selected document."""
    presentation = get_session("wpp").get_document(document)
    total = int(presentation.Slides.Count)
    if not 1 <= slide_index <= total:
        raise ValueError(f"Slide index {slide_index} out of range (1-{total})")
    slide = presentation.Slides(slide_index)
    if navigate:
        goto_slide(slide_index)
    return slide


def get_shape(params: ShapeInput, *, navigate: bool = True) -> NativeObject:
    """Resolve an unambiguous shape before changing the editor view."""
    slide = get_slide(params.slide_index, params.document)
    shape = resolve_shape(slide, params.shape_name_or_index)
    if navigate:
        goto_slide(params.slide_index)
    return shape
