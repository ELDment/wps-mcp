"""Render bounded PDF previews with deterministic resource cleanup."""

from contextlib import ExitStack, closing
from io import BytesIO
from pathlib import Path
from threading import Lock

RENDER_LOCK = Lock()


def render_pdf_page(
    file_path: str | Path, page_index: int = 1, width: int = 1280
) -> bytes:
    """Release rendering resources in reverse ownership order."""
    import pypdfium2

    if not 320 <= width <= 4096:
        raise ValueError("Preview width must be between 320 and 4096 pixels")
    with RENDER_LOCK, ExitStack() as resources:
        pdf = resources.enter_context(pypdfium2.PdfDocument(str(file_path)))
        if not 1 <= page_index <= len(pdf):
            raise ValueError(f"page_index must be between 1 and {len(pdf)}")
        page = resources.enter_context(closing(pdf[page_index - 1]))
        scale = width / page.get_width()
        if page.get_height() * scale * width > 24_000_000:
            raise ValueError(
                "Preview exceeds 24 million pixels; choose a smaller width"
            )
        bitmap = resources.enter_context(closing(page.render(scale=scale)))
        image = resources.enter_context(closing(bitmap.to_pil()))
        output = resources.enter_context(BytesIO())
        image.save(output, format="PNG")
        return output.getvalue()
