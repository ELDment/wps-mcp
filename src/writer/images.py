"""Embed local images into native Writer document ranges."""

from pydantic import Field

from src.common.paths import local_path
from src.models import DocumentInput, JsonObject
from src.native.runtime import get_session
from src.protocol.definitions import WRITE_ANNOTATIONS, tool
from src.writer.access import resolve_range
from src.writer.models import Position, PositivePoints


class ImageInput(DocumentInput):
    """Embed a local image at a native position with optional sizing."""

    position: Position
    path: str = Field(min_length=1)
    width: PositivePoints | None = None
    height: PositivePoints | None = None
    alternative_text: str | None = None


@tool(
    "wps_writer_images",
    operation="add",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def add_image(params: ImageInput) -> JsonObject:
    """Embed a local image as an inline shape.

    Sizes use points; one dimension preserves aspect ratio, two set an
    explicit size.
    """
    image_path = local_path(params.path, exists=True)
    if image_path.suffix.lower() not in {
        ".png",
        ".jpg",
        ".jpeg",
        ".bmp",
        ".gif",
        ".tif",
        ".tiff",
        ".svg",
        ".emf",
        ".wmf",
    }:
        raise ValueError("Unsupported image extension")
    document = get_session("writer").get_document(params.document)
    text_range = resolve_range(document, params.position, params.position)
    shape = document.InlineShapes.AddPicture(
        str(image_path), False, True, text_range
    )
    if params.width is not None and params.height is not None:
        shape.LockAspectRatio = 0
    elif params.width is not None or params.height is not None:
        shape.LockAspectRatio = -1
    if params.width is not None:
        shape.Width = params.width
    if params.height is not None:
        shape.Height = params.height
    if params.alternative_text is not None:
        shape.AlternativeText = params.alternative_text
    return {
        "start": int(shape.Range.Start),
        "end": int(shape.Range.End),
        "width": float(shape.Width),
        "height": float(shape.Height),
    }
