"""Validate exports and publish files with scoped native resources."""

from __future__ import annotations

from contextlib import ExitStack
from pathlib import Path, PureWindowsPath
from typing import Literal, Self

from PIL import Image
from pydantic import Field, model_validator

from wps_mcp.common.files import publish_file, verify_output
from wps_mcp.common.paths import destination, local_path
from wps_mcp.models import DocumentInput, JsonObject, NativeObject
from wps_mcp.native.clipboard import copy_images, ensure_clipboard_available
from wps_mcp.native.files import temporary_directory
from wps_mcp.native.runtime import get_runtime, get_session
from wps_mcp.presentation.access import get_shape, get_slide
from wps_mcp.presentation.constants import SHAPE_FORMAT_MAP
from wps_mcp.presentation.shapes.models import ShapeInput
from wps_mcp.protocol.definitions import tool


class ExportShapeInput(ShapeInput):
    """Input for exporting a shape as an image file."""

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
    extensions = (
        {".jpg", ".jpeg"} if fmt_key in {"jpg", "jpeg"} else {"." + fmt_key}
    )
    target = destination(params.file_path, extensions)
    shape = get_shape(params)
    dimensions = (
        ()
        if params.width is None
        else (params.width,)
        if params.height is None
        else (params.width, params.height)
    )
    with temporary_directory(
        prefix="wps_export_", dir=target.parent
    ) as directory:
        staged = Path(directory) / ("shape." + fmt_key)
        shape.Export(str(staged), format_type, *dimensions)
        verify_output(staged)
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
    verify_output(path)
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


@tool(
    "wps_ppt_export",
    component="wpp",
    annotations={"title": "Export as Images", "openWorldHint": True},
    operation="images",
)
def export_images(params: ExportImagesInput) -> JsonObject:
    """Export verified slide images without replacing existing files."""
    output_dir = local_path(params.output_dir)
    if output_dir.exists() and not output_dir.is_dir():
        raise ValueError("output_dir must be a directory")
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
    paths = [
        output_dir
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
    output_dir.mkdir(parents=True, exist_ok=True)
    with temporary_directory(
        prefix="wps_export_", dir=output_dir
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
        "output_dir": str(output_dir),
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
    """Copy one image or a desktop file list for multiple images."""
    ensure_clipboard_available()
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
            temporary_directory(prefix="wps_clipboard_")
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
        method = copy_images(paths)
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
    slide = get_slide(params.slide_index, params.document)
    with temporary_directory(prefix="wpp-preview-") as directory:
        output_path = Path(directory) / "slide.png"
        _export_slide_image(
            slide,
            output_path,
            "png",
            None,
            None,
        )
        return output_path.read_bytes()
