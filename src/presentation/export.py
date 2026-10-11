"""Validate exports and publish files with scoped native resources."""

from __future__ import annotations

import logging
import struct
from collections.abc import Iterator
from contextlib import ExitStack, closing, contextmanager
from io import BytesIO
from pathlib import Path, PureWindowsPath
from tempfile import TemporaryDirectory
from typing import Literal, Self

import pypdfium2 as pdfium
import win32clipboard
from PIL import Image
from pydantic import Field, model_validator

from src.common.files import publish_file
from src.common.rendering import RENDER_LOCK
from src.models import DocumentInput, JsonObject, NativeObject
from src.native.runtime import get_runtime, get_session
from src.presentation.constants import (
    MSO_TRUE,
    PP_FIXED_FORMAT_TYPE_PDF,
    SHAPE_FORMAT_MAP,
)
from src.presentation.navigation import goto_slide
from src.presentation.shape_lookup import resolve_shape as _get_shape
from src.protocol.definitions import tool


class ExportShapeInput(DocumentInput):
    """Input for exporting a shape as an image file."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: str | int = Field(
        ...,
        description="Shape name, group path, or "
        "1-based index; prefer a stable name.",
    )
    file_path: str = Field(..., description="Output file path")
    format: str = Field(
        default="png",
        description="Image format: 'png', "
        "'jpg', 'gif', 'bmp', 'wmf', or 'emf'",
    )
    width: int | None = Field(
        default=None, description="Export width in pixels", gt=0
    )
    height: int | None = Field(
        default=None, description="Export height in pixels", gt=0
    )

    @model_validator(mode="after")
    def validate_size(self) -> ExportShapeInput:
        """Require a width when an export height is supplied."""
        if self.height is not None and self.width is None:
            raise ValueError("Export height requires width")
        return self


@tool(
    "wps_ppt_export",
    component="wpp",
    annotations={"title": "Export Shape", "openWorldHint": True},
    operation="shape",
)
def export_shape(params: ExportShapeInput) -> JsonObject:
    """Export a verified shape image to an unused pathname."""
    fmt_key = params.format.strip().lower()
    if fmt_key not in SHAPE_FORMAT_MAP:
        raise ValueError(
            f"Unknown format '{params.format}'. Use one "
            f"of: {', '.join(SHAPE_FORMAT_MAP.keys())}"
        )
    format_type = SHAPE_FORMAT_MAP[fmt_key]
    target = Path(params.file_path).resolve()
    if target.exists():
        raise FileExistsError(f"Refusing to replace existing file: {target}")
    session = get_session("wpp")
    slide = session.get_document(params.document).Slides(params.slide_index)
    shape = _get_shape(slide, params.shape_name_or_index)
    dimensions = (
        ()
        if params.width is None
        else (params.width,)
        if params.height is None
        else (params.width, params.height)
    )
    target.parent.mkdir(parents=True, exist_ok=True)
    goto_slide(session.get_application(), params.slide_index)
    with TemporaryDirectory(
        prefix="wps_export_", dir=target.parent
    ) as directory:
        staged = Path(directory) / ("shape." + fmt_key)
        shape.Export(str(staged), format_type, *dimensions)
        if not staged.is_file() or staged.stat().st_size == 0:
            raise RuntimeError("WPS did not create a non-empty shape image")
        if fmt_key in {"png", "jpg", "jpeg", "gif", "bmp"}:
            with Image.open(staged) as exported:
                expected = (
                    "JPEG" if fmt_key in {"jpg", "jpeg"} else fmt_key.upper()
                )
                if exported.format != expected:
                    raise RuntimeError(
                        f"WPS produced {exported.format}, expected {expected}"
                    )
                exported.verify()
        publish_file(staged, target)
    return {
        "success": True,
        "shape_name": shape.Name,
        "file_path": str(target),
    }


logger = logging.getLogger(__name__)


class ExportPDFInput(DocumentInput):
    """Export all slides or an inclusive range to a new PDF."""

    file_path: str = Field(min_length=1)
    slide_range_start: int | None = Field(default=None, ge=1)
    slide_range_end: int | None = Field(default=None, ge=1)

    @model_validator(mode="after")
    def validate_range(self) -> Self:
        """Require paired, ascending slide range boundaries."""
        if (self.slide_range_start is None) != (self.slide_range_end is None):
            raise ValueError(
                "Provide both slide_range_start and slide_range_end"
            )
        if (
            self.slide_range_start is not None
            and self.slide_range_end is not None
            and self.slide_range_start > self.slide_range_end
        ):
            raise ValueError(
                "slide_range_start must not exceed slide_range_end"
            )
        return self


class ExportImagesInput(DocumentInput):
    """Choose a single slide, a list, a range, or every slide."""

    output_dir: str = Field(min_length=1)
    format: Literal["png", "jpg"] = "png"
    slide_index: int | None = Field(default=None, ge=1)
    slide_indices: list[int] | None = Field(default=None, min_length=1)
    from_index: int | None = Field(default=None, ge=1)
    to_index: int | None = Field(default=None, ge=1)
    width: int | None = Field(default=None, gt=0)
    height: int | None = Field(default=None, gt=0)
    file_name: str | None = Field(default=None, min_length=1)

    @model_validator(mode="after")
    def validate_selection(self) -> Self:
        """Validate slide selection, image size and output filename."""
        ranged = self.from_index is not None or self.to_index is not None
        if (
            sum(
                (
                    self.slide_index is not None,
                    self.slide_indices is not None,
                    ranged,
                )
            )
            > 1
        ):
            raise ValueError("Choose one slide selection mode")
        if self.slide_indices is not None:
            if min(self.slide_indices) < 1 or len(
                set(self.slide_indices)
            ) != len(self.slide_indices):
                raise ValueError(
                    "slide_indices must contain unique positive indices"
                )
        if ranged:
            if self.from_index is None or self.to_index is None:
                raise ValueError("Provide both from_index and to_index")
            if self.from_index > self.to_index:
                raise ValueError("from_index must not exceed to_index")
        if self.height is not None and self.width is None:
            raise ValueError("height requires width")
        if self.file_name is not None:
            path = PureWindowsPath(self.file_name)
            if self.slide_index is None:
                raise ValueError("file_name requires slide_index")
            if (
                path.name != self.file_name
                or path.drive
                or path.root
                or self.file_name in (".", "..")
                or ":" in self.file_name
            ):
                raise ValueError(
                    "file_name must be a filename "
                    "without directories or streams"
                )
        return self


class CopyToClipboardInput(DocumentInput):
    """Copy the current slide or explicitly selected slides."""

    slide_indices: list[int] | None = Field(default=None, min_length=1)
    width: int | None = Field(default=None, gt=0)
    height: int | None = Field(default=None, gt=0)

    @model_validator(mode="after")
    def validate_selection(self) -> Self:
        """Validate unique indices and width-dependent height."""
        if self.slide_indices is not None:
            if min(self.slide_indices) < 1 or len(
                set(self.slide_indices)
            ) != len(self.slide_indices):
                raise ValueError(
                    "slide_indices must contain unique positive indices"
                )
        if self.height is not None and self.width is None:
            raise ValueError("height requires width")
        return self


def _export_slide_image(
    slide: NativeObject,
    path: Path,
    image_format: str,
    width: int | None,
    height: int | None,
) -> None:
    """Render a slide and verify that WPS produced a readable image."""
    dimensions = (
        ()
        if width is None
        else ((width,) if height is None else (width, height))
    )
    slide.Export(str(path), image_format.upper(), *dimensions)
    if not path.is_file() or path.stat().st_size == 0:
        raise RuntimeError(f"WPS did not create a slide image at {path}")
    with Image.open(path) as exported:
        expected = "JPEG" if image_format.lower() == "jpg" else "PNG"
        if exported.format != expected:
            raise RuntimeError(
                f"WPS produced {exported.format}, expected {expected}"
            )
        if width is not None and exported.width != width:
            raise RuntimeError("WPS did not honor the requested image width")
        if height is not None and exported.height != height:
            raise RuntimeError("WPS did not honor the requested image height")
        exported.verify()


@contextmanager
def _clipboard() -> Iterator[None]:
    """Pair clipboard access while preserving the primary failure."""
    win32clipboard.OpenClipboard()
    failed = False
    try:
        yield
    except BaseException:
        failed = True
        raise
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            if not failed:
                raise
            logger.exception(
                "Could not close clipboard after an operation failed"
            )


@tool(
    "wps_ppt_export",
    component="wpp",
    annotations={"title": "Export to PDF", "openWorldHint": True},
    operation="pdf",
)
def export_pdf(params: ExportPDFInput) -> JsonObject:
    """Export to a new PDF without changing the deck's print ranges."""
    presentation = get_session("wpp").get_document(params.document)
    total = int(presentation.Slides.Count)
    start, end = params.slide_range_start or 1, params.slide_range_end or total
    if not 1 <= start <= end <= total:
        raise ValueError(
            f"Slide range must satisfy 1 <= start <= end <= {total}"
        )
    target = Path(params.file_path).resolve()
    if target.exists():
        raise FileExistsError(f"Refusing to replace existing file: {target}")
    target.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(
        prefix="wps_export_", dir=target.parent
    ) as directory:
        rendered = Path(directory) / "presentation.pdf"
        presentation.ExportAsFixedFormat(
            str(rendered),
            PP_FIXED_FORMAT_TYPE_PDF,
            PrintHiddenSlides=MSO_TRUE,
            PrintRange=None,
            RangeType=1,
        )
        with RENDER_LOCK:
            with closing(pdfium.PdfDocument(str(rendered))) as source:
                if len(source) != total:
                    raise RuntimeError(
                        f"WPS exported {len(source)} pages; expected {total}"
                    )
                if start != 1 or end != total:
                    selected = Path(directory) / "selected.pdf"
                    with closing(pdfium.PdfDocument.new()) as output:
                        output.import_pages(
                            source, pages=list(range(start - 1, end))
                        )
                        output.save(str(selected))
                    rendered = selected
            with closing(pdfium.PdfDocument(str(rendered))) as verified:
                if len(verified) != end - start + 1:
                    raise RuntimeError(
                        "The exported PDF has an unexpected page count"
                    )
        publish_file(rendered, target)
    return {
        "success": True,
        "file_path": str(target),
        "slide_range_start": params.slide_range_start,
        "slide_range_end": params.slide_range_end,
        "total_slides": total,
    }


@tool(
    "wps_ppt_export",
    component="wpp",
    annotations={"title": "Export as Images", "openWorldHint": True},
    operation="images",
)
def export_images(params: ExportImagesInput) -> JsonObject:
    """Export verified slide images without replacing existing files."""
    presentation = get_session("wpp").get_document(params.document)
    total = int(presentation.Slides.Count)
    if params.slide_index is not None:
        targets = [params.slide_index]
    elif params.slide_indices is not None:
        targets = params.slide_indices
    elif params.from_index is not None and params.to_index is not None:
        targets = list(range(params.from_index, params.to_index + 1))
    else:
        targets = list(range(1, total + 1))
    if not targets or any(index > total for index in targets):
        raise ValueError(f"Slide indices must be within 1-{total}")
    destination = Path(params.output_dir).resolve()
    paths = [
        destination
        / (
            (
                Path(params.file_name).stem
                if params.file_name
                else f"Slide{index}"
            )
            + "."
            + params.format
        )
        for index in targets
    ]
    for path in paths:
        if path.exists():
            raise FileExistsError(f"Refusing to replace existing file: {path}")
    destination.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(
        prefix="wps_export_", dir=destination
    ) as directory:
        staged = [Path(directory) / path.name for path in paths]
        for index, path in zip(targets, staged, strict=True):
            _export_slide_image(
                presentation.Slides(index),
                path,
                params.format,
                params.width,
                params.height,
            )
        published: list[str] = []
        try:
            for source, target in zip(staged, paths, strict=True):
                publish_file(source, target)
                published.append(str(target))
        except OSError as error:
            raise RuntimeError(
                f"Export publication stopped; "
                f"completed files: {published}. {error}"
            ) from error
    result: JsonObject = {
        "success": True,
        "output_dir": str(destination),
        "format": params.format,
        "slide_indices": targets,
        "total_slides": total,
        "files_count": len(paths),
        "files": [str(path) for path in paths],
    }
    if params.slide_index is not None:
        result["slide_index"] = params.slide_index
    return result


@tool(
    "wps_ppt_export",
    component="wpp",
    annotations={"title": "Copy Slides to Clipboard", "idempotentHint": True},
    operation="clipboard",
)
def copy_to_clipboard(params: CopyToClipboardInput) -> JsonObject:
    """Copy a bitmap or a file drop list for multiple images."""
    session = get_session("wpp")
    presentation = session.get_document(params.document)
    indices = params.slide_indices
    if indices is None:
        window = session.get_window()
        if window is None:
            raise ValueError(
                "No editor window is available; specify slide_indices"
            )
        indices = [int(window.View.Slide.SlideIndex)]
    if any(index > presentation.Slides.Count for index in indices):
        raise ValueError(
            f"Slide indices must be within 1-{presentation.Slides.Count}"
        )
    with ExitStack() as resources:
        directory = resources.enter_context(
            TemporaryDirectory(prefix="wps_clipboard_")
        )
        paths = [Path(directory) / f"Slide{index}.png" for index in indices]
        for index, path in zip(indices, paths, strict=True):
            _export_slide_image(
                presentation.Slides(index),
                path,
                "png",
                params.width,
                params.height,
            )
        if len(paths) == 1:
            with (
                Image.open(paths[0]) as source,
                source.convert("RGBA") as rgba,
            ):
                with (
                    Image.new("RGB", rgba.size, "white") as canvas,
                    rgba.getchannel("A") as alpha,
                ):
                    canvas.paste(rgba, mask=alpha)
                    with BytesIO() as buffer:
                        canvas.save(buffer, format="BMP")
                        payload = buffer.getvalue()[14:]
            clipboard_format, method = 8, "CF_DIB"
        else:
            payload = struct.pack("<IiiII", 20, 0, 0, 0, 1)
            payload += (
                "\0".join(str(path) for path in paths) + "\0\0"
            ).encode("utf-16-le")
            clipboard_format, method = 15, "CF_HDROP"
        with _clipboard():
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(clipboard_format, payload)
        if len(paths) > 1:
            get_runtime().resources.enter_context(resources.pop_all())
    return {
        "success": True,
        "slide_indices": indices,
        "count": len(indices),
        "method": method,
        "temp_dir": directory if len(indices) > 1 else None,
    }


class GetSlidePreviewInput(DocumentInput):
    """Render one slide by its one-based index."""

    slide_index: int = Field(default=1, ge=1)


@tool(
    "wps_ppt_export",
    component="wpp",
    annotations={"readOnlyHint": True},
    operation="preview",
)
def get_slide_preview(params: GetSlidePreviewInput) -> bytes:
    """Return PNG bytes for serialization by the MCP adapter."""
    presentation = get_session("wpp").get_document(params.document)
    if params.slide_index > presentation.Slides.Count:
        raise ValueError(
            f"Slide index {params.slide_index} out "
            f"of range (1-{presentation.Slides.Count})"
        )
    with TemporaryDirectory(prefix="wpp-preview-") as temporary_directory:
        output_path = Path(temporary_directory) / "slide.png"
        presentation.Slides(params.slide_index).Export(str(output_path), "PNG")
        with output_path.open("rb") as image_file:
            image_data = image_file.read()
        if not image_data.startswith(b"\x89PNG\r\n\x1a\n"):
            raise RuntimeError("WPS did not produce a valid PNG preview")
        return image_data
