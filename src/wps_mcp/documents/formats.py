"""Check presentation format capabilities before native saves."""

from pathlib import Path

from wps_mcp.models import NativeObject


def check_presentation_format(
    document: NativeObject, file_path: str | Path
) -> None:
    """Reject DPS saves that would convert editable chart shapes."""
    if Path(file_path).suffix.lower() != ".dps":
        return
    for slide_index in range(1, document.Slides.Count + 1):
        shapes = document.Slides.Item(slide_index).Shapes
        pending = [shapes.Item(index) for index in range(1, shapes.Count + 1)]
        while pending:
            shape = pending.pop()
            if shape.HasChart:
                raise ValueError(
                    "Save as .pptx to preserve editable "
                    "charts; WPS converts them in .dps"
                )
            if shape.Type == 6:
                pending.extend(
                    shape.GroupItems.Item(index)
                    for index in range(1, shape.GroupItems.Count + 1)
                )
