"""Picture editing, bounded downloads and Material icons."""

from __future__ import annotations

import logging
import urllib.error
from pathlib import Path
from typing import Literal
from urllib.parse import urlsplit

from pydantic import Field, model_validator

from wps_mcp.common.colors import RgbColor, hex_to_int, int_to_hex
from wps_mcp.common.paths import local_path
from wps_mcp.errors import (
    OperationTimeoutError,
    UnknownOutcomeError,
    WpsUnavailableError,
)
from wps_mcp.models import DocumentInput, JsonObject, NativeObject
from wps_mcp.native.files import temporary_directory
from wps_mcp.native.objects import is_optional_member_error, optional_property
from wps_mcp.native.runtime import get_session
from wps_mcp.native.state import temporary_attributes
from wps_mcp.presentation import assets
from wps_mcp.presentation.constants import (
    ICON_PACKAGE_BASE,
    MSO_FALSE,
    MSO_LINKED_PICTURE,
    MSO_PICTURE,
    MSO_TRUE,
    PICTURE_COLOR_TYPE_MAP,
    PICTURE_COLOR_TYPE_NAMES,
)
from wps_mcp.presentation.navigation import goto_slide
from wps_mcp.presentation.shapes.formatting import resolve_color
from wps_mcp.presentation.shapes.lookup import resolve_shape as _get_shape
from wps_mcp.presentation.shapes.models import ShapeReference
from wps_mcp.presentation.shapes.types import SHAPE_NAME_MAP
from wps_mcp.presentation.shapes.z_order import (
    ZORDER_FIELD_DESCRIPTION,
    place_in_zorder,
)
from wps_mcp.protocol.definitions import tool

logger = logging.getLogger(__name__)


def _size_picture(
    picture: NativeObject,
    left: float,
    top: float,
    width: float | None,
    height: float | None,
    fit: bool = False,
) -> None:
    """Apply dimensions, preserving aspect ratio for a single dimension."""
    if fit and width is not None and height is not None:
        if picture.Width <= 0 or picture.Height <= 0:
            raise ValueError("The image has no usable dimensions")
        picture.LockAspectRatio = MSO_TRUE
        scale = min(width / picture.Width, height / picture.Height)
        new_width = picture.Width * scale
        new_height = picture.Height * scale
        picture.Width = new_width
        picture.Left = left + (width - new_width) / 2
        picture.Top = top + (height - new_height) / 2
    elif width is not None and height is not None:
        picture.LockAspectRatio = MSO_FALSE
        picture.Width, picture.Height = width, height
    elif width is not None:
        picture.LockAspectRatio = MSO_TRUE
        picture.Width = width
    elif height is not None:
        picture.LockAspectRatio = MSO_TRUE
        picture.Height = height


def _picture_result(
    slide: NativeObject, picture: NativeObject, zorder: str
) -> JsonObject:
    """Apply stacking order and report the inserted picture dimensions."""
    placement = place_in_zorder(slide, picture, zorder)
    return {
        "success": True,
        "shape_name": picture.Name,
        "shape_index": picture.ZOrderPosition,
        "width": round(picture.Width, 2),
        "height": round(picture.Height, 2),
        **placement,
    }


def _insert_picture(
    slide: NativeObject,
    data: bytes,
    suffix: str,
    left: float,
    top: float,
    width: float | None,
    height: float | None,
    fit: bool,
    zorder: str,
) -> JsonObject:
    """Embed downloaded image bytes while the temporary file is alive."""
    with temporary_directory(prefix="wps-image-") as directory:
        path = Path(directory) / ("image" + suffix)
        with path.open("wb") as stream:
            stream.write(data)
        picture = slide.Shapes.AddPicture(
            str(path), MSO_FALSE, MSO_TRUE, left, top, -1, -1
        )
        _size_picture(picture, left, top, width, height, fit)
        return _picture_result(slide, picture, zorder)


class CropPictureInput(DocumentInput):
    """Input for cropping a picture shape."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: ShapeReference = Field(
        ...,
        description=(
            "Shape name, group path, or 1-based index; prefer a stable name."
        ),
    )
    crop_left: float | None = Field(
        default=None,
        description="Crop from left in points (original image coordinates)",
    )
    crop_right: float | None = Field(
        default=None,
        description="Crop from right in points (original image coordinates)",
    )
    crop_top: float | None = Field(
        default=None,
        description="Crop from top in points (original image coordinates)",
    )
    crop_bottom: float | None = Field(
        default=None,
        description="Crop from bottom in points (original image coordinates)",
    )
    crop_shape: str | int | None = Field(
        default=None,
        description=(
            "Crop mask: friendly shape name or MsoAutoShapeType integer."
        ),
    )
    crop_fit: Literal["square", "1:1"] | None = Field(
        default=None,
        description="Fit a square crop; exclusive with manual crop edges.",
    )
    crop_anchor: float | None = Field(
        default=0.5,
        ge=0.0,
        le=1.0,
        description=(
            "Square crop anchor: 0 keeps the start, 0.5 centers, 1 "
            "keeps the end."
        ),
    )
    corner_radius_pt: float | None = Field(
        default=None,
        ge=0.0,
        description=(
            "Rounded-rectangle radius in points; at most half the short side."
        ),
    )

    @model_validator(mode="after")
    def validate_crop_fit_exclusivity(self) -> CropPictureInput:
        """Reject automatic cropping combined with manual crop edges."""
        if self.crop_fit is not None:
            manual = [
                self.crop_left,
                self.crop_right,
                self.crop_top,
                self.crop_bottom,
            ]
            if any(v is not None for v in manual):
                raise ValueError(
                    "crop_fit cannot be combined with "
                    "crop_left/crop_right/crop_top/crop_bottom. Use crop_fit "
                    "alone for automatic cropping, or manual crop "
                    "values without "
                    "crop_fit."
                )
        return self


class SetPictureFormatInput(DocumentInput):
    """Adjust brightness, contrast, color and picture transparency."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: ShapeReference = Field(
        ...,
        description=(
            "Shape name, group path, or 1-based index; prefer a stable name."
        ),
    )
    brightness: float | None = Field(
        default=None,
        ge=0.0,
        le=1.0,
        description="Picture brightness (0.0 = darkest, 1.0 = brightest)",
    )
    contrast: float | None = Field(
        default=None,
        ge=0.0,
        le=1.0,
        description="Picture contrast (0.0 = lowest, 1.0 = highest)",
    )
    color_type: str | None = Field(
        default=None,
        description=(
            "Picture color type: 'automatic', 'grayscale', "
            "'black_and_white', 'watermark'"
        ),
    )
    transparent_color: RgbColor | None = Field(
        default=None,
        description=(
            "Color key #RRGGBB; enables transparency unless explicitly "
            "disabled."
        ),
    )
    transparent_background: bool | None = Field(
        default=None,
        description="Explicitly enable/disable color-key transparency",
    )

    @model_validator(mode="after")
    def validate_at_least_one_param(self) -> SetPictureFormatInput:
        """Require a picture adjustment and validate its color mode."""
        params = [
            self.brightness,
            self.contrast,
            self.color_type,
            self.transparent_color,
            self.transparent_background,
        ]
        if all(v is None for v in params):
            raise ValueError(
                "At least one adjustment parameter must be provided: "
                "brightness, contrast, color_type, transparent_color, or "
                "transparent_background."
            )
        if (
            self.color_type is not None
            and self.color_type not in PICTURE_COLOR_TYPE_MAP
        ):
            raise ValueError(f"Unknown color_type '{self.color_type}'")
        return self


class AddPictureFromUrlInput(DocumentInput):
    """Input for adding a picture from a URL."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    url: str = Field(..., description="URL of the image to download")
    left: float = Field(default=100, description="Left position in points")
    top: float = Field(default=100, description="Top position in points")
    width: float | None = Field(
        default=None, description="Width in points (auto if not set)", gt=0
    )
    height: float | None = Field(
        default=None, description="Height in points (auto if not set)", gt=0
    )
    svg_color: RgbColor | None = Field(
        default=None,
        description=(
            "Replace 'currentColor' in SVG files with this color (e.g. "
            "'#1A73E8'). Only applies to SVG files."
        ),
    )
    fit: bool = Field(
        default=False,
        description=(
            "Preserve aspect ratio and center within the required width "
            "and height."
        ),
    )
    zorder: Literal["front", "back", "behind_text"] = Field(
        default="front", description=ZORDER_FIELD_DESCRIPTION
    )

    @model_validator(mode="after")
    def validate_fit(self) -> AddPictureFromUrlInput:
        """Require a valid HTTP URL and dimensions for proportional fit."""
        if self.fit and (self.width is None or self.height is None):
            raise ValueError("fit requires both width and height")
        parsed = urlsplit(self.url)
        if parsed.scheme not in {"http", "https"} or not parsed.hostname:
            raise ValueError("Only absolute HTTP or HTTPS URLs are supported")
        return self


class AddSvgIconInput(DocumentInput):
    """Input for adding a Material Symbols icon as SVG image."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    icon_name: str = Field(
        ...,
        description="Material icon name, such as home or arrow_forward.",
        min_length=1,
        pattern="^[a-z0-9_]+$",
    )
    left: float = Field(default=100, description="Left position in points")
    top: float = Field(default=100, description="Top position in points")
    width: float = Field(
        default=72, description="Width of the area in points", gt=0
    )
    height: float = Field(
        default=72, description="Height of the area in points", gt=0
    )
    color: str = Field(
        default="accent1",
        description="Icon color: #RRGGBB or a theme color such as accent1.",
    )
    style: Literal["outlined", "rounded", "sharp"] = Field(
        default="outlined",
        description="Icon style: 'outlined', 'rounded', or 'sharp'",
    )
    filled: bool = Field(
        default=False,
        description=(
            "If true, use the filled variant of the icon instead of outline."
        ),
    )
    zorder: Literal["front", "back", "behind_text"] = Field(
        default="front", description=ZORDER_FIELD_DESCRIPTION
    )


@tool(
    "wps_ppt_pictures",
    component="wpp",
    annotations={"title": "Crop Picture", "idempotentHint": True},
    operation="crop",
)
def crop_picture(params: CropPictureInput) -> JsonObject:
    """Crop a picture shape — rectangular trim and/or crop-to-shape."""
    goto_slide(params.slide_index)
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    shape = _get_shape(slide, params.shape_name_or_index)
    if shape.Type not in (MSO_LINKED_PICTURE, MSO_PICTURE):
        raise ValueError(
            f"Shape '{shape.Name}' is not a picture (type={shape.Type})"
        )
    auto_shape_int = None
    if params.crop_shape is not None:
        if isinstance(params.crop_shape, str):
            key = params.crop_shape.strip().lower()
            if key.isdigit():
                auto_shape_int = int(key)
            elif key not in SHAPE_NAME_MAP:
                raise ValueError(f"Unknown crop_shape '{params.crop_shape}'")
            else:
                auto_shape_int = SHAPE_NAME_MAP[key]
        else:
            auto_shape_int = int(params.crop_shape)
    if params.crop_fit is not None:
        fit_key = params.crop_fit.strip().lower()
        if fit_key not in ("square", "1:1"):
            raise ValueError(
                f"Unknown crop_fit '{params.crop_fit}'. Supported "
                f"values: 'square', '1:1'"
            )
    pic_fmt = shape.PictureFormat
    if params.crop_fit is not None:
        anchor = params.crop_anchor if params.crop_anchor is not None else 0.5
        with temporary_attributes(shape, LockAspectRatio=MSO_FALSE):
            cur_w, cur_h = shape.Width, shape.Height
            with (
                temporary_attributes(
                    pic_fmt, CropLeft=0, CropRight=0, CropTop=0, CropBottom=0
                ),
                temporary_attributes(shape, Width=cur_w, Height=cur_h),
            ):
                shape.ScaleWidth(1.0, MSO_TRUE)
                shape.ScaleHeight(1.0, MSO_TRUE)
                orig_w, orig_h = shape.Width, shape.Height
            horizontal = max(0, orig_w - orig_h)
            vertical = max(0, orig_h - orig_w)
            pic_fmt.CropLeft = horizontal * anchor
            pic_fmt.CropRight = horizontal * (1.0 - anchor)
            pic_fmt.CropTop = vertical * anchor
            pic_fmt.CropBottom = vertical * (1.0 - anchor)
            min_dim = min(cur_w, cur_h)
            shape.Width = min_dim
            shape.Height = min_dim
    else:
        if params.crop_left is not None:
            pic_fmt.CropLeft = params.crop_left
        if params.crop_right is not None:
            pic_fmt.CropRight = params.crop_right
        if params.crop_top is not None:
            pic_fmt.CropTop = params.crop_top
        if params.crop_bottom is not None:
            pic_fmt.CropBottom = params.crop_bottom
    if auto_shape_int is not None:
        try:
            shape.AutoShapeType = auto_shape_int
        except (
            UnknownOutcomeError,
            OperationTimeoutError,
            WpsUnavailableError,
        ):
            raise
        except Exception as e:
            raise ValueError(
                f"Could not apply crop_shape {auto_shape_int}: {e}"
            ) from e
    if params.corner_radius_pt is not None:
        try:
            current_type = shape.AutoShapeType
        except Exception as error:
            if not is_optional_member_error(error):
                raise
            current_type = None
        if current_type == 5:
            short_side = min(shape.Width, shape.Height)
            adj_value = (
                min(0.5, params.corner_radius_pt / short_side)
                if short_side > 0
                else 0
            )
            shape.Adjustments[1] = adj_value
    try:
        auto_shape_val = shape.AutoShapeType
    except Exception as error:
        if not is_optional_member_error(error):
            raise
        auto_shape_val = None
    result = {
        "success": True,
        "shape_name": shape.Name,
        "width": round(shape.Width, 2),
        "height": round(shape.Height, 2),
        "crop_left": round(pic_fmt.CropLeft, 2),
        "crop_right": round(pic_fmt.CropRight, 2),
        "crop_top": round(pic_fmt.CropTop, 2),
        "crop_bottom": round(pic_fmt.CropBottom, 2),
        "crop_shape": auto_shape_val,
    }
    return result


@tool(
    "wps_ppt_pictures",
    component="wpp",
    annotations={"title": "Set Picture Format", "idempotentHint": True},
    operation="format",
)
def set_picture_format(params: SetPictureFormatInput) -> JsonObject:
    """Apply picture brightness, contrast, color and transparency."""
    goto_slide(params.slide_index)
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    shape = _get_shape(slide, params.shape_name_or_index)
    if shape.Type not in (MSO_LINKED_PICTURE, MSO_PICTURE):
        raise ValueError(
            f"Shape '{shape.Name}' is not a picture (type={shape.Type})"
        )
    pf = shape.PictureFormat
    if params.brightness is not None:
        pf.Brightness = params.brightness
    if params.contrast is not None:
        pf.Contrast = params.contrast
    if params.color_type is not None:
        pf.ColorType = PICTURE_COLOR_TYPE_MAP[params.color_type]
    if params.transparent_color is not None:
        pf.TransparencyColor = hex_to_int(params.transparent_color)
        if params.transparent_background is not False:
            pf.TransparentBackground = MSO_TRUE
    if params.transparent_background is not None:
        pf.TransparentBackground = (
            MSO_TRUE if params.transparent_background else MSO_FALSE
        )
    transparency = optional_property(pf, "TransparencyColor")
    cur_color_type = pf.ColorType
    color_type_name = PICTURE_COLOR_TYPE_NAMES.get(cur_color_type, "unknown")
    result = {
        "success": True,
        "shape_name": shape.Name,
        "brightness": pf.Brightness,
        "contrast": pf.Contrast,
        "color_type": cur_color_type,
        "color_type_name": color_type_name,
        "transparent_background": bool(pf.TransparentBackground),
        "transparent_color_hex": (
            None if transparency is None else int_to_hex(transparency)
        ),
    }
    return result


@tool(
    "wps_ppt_pictures",
    component="wpp",
    annotations={"title": "Add Picture from URL", "openWorldHint": True},
    operation="add_url",
)
def add_picture_from_url(params: AddPictureFromUrlInput) -> JsonObject:
    """Download and embed a picture; fit preserves its aspect ratio."""
    session = get_session("wpp")
    slide = session.get_document().Slides(params.slide_index)
    data, content_type = assets.download_resource(params.url)
    suffix = Path(urlsplit(params.url).path).suffix.lower()
    if suffix == ".svg" or "svg" in content_type.lower():
        suffix = ".svg"
        if params.svg_color:
            data = (
                data.decode("utf-8")
                .replace("currentColor", params.svg_color)
                .encode("utf-8")
            )
    elif suffix not in {
        ".png",
        ".jpg",
        ".jpeg",
        ".gif",
        ".bmp",
        ".tif",
        ".tiff",
        ".webp",
        ".emf",
        ".wmf",
    }:
        suffix = ".png"
    goto_slide(params.slide_index)
    result = _insert_picture(
        slide,
        data,
        suffix,
        params.left,
        params.top,
        params.width,
        params.height,
        params.fit,
        params.zorder,
    )
    result["source_url"] = params.url
    return result


@tool(
    "wps_ppt_pictures",
    component="wpp",
    annotations={"title": "Add SVG Icon", "openWorldHint": True},
    operation="add_icon",
)
def add_svg_icon(params: AddSvgIconInput) -> JsonObject:
    """Download a public Material icon and embed it using the chosen color."""
    session = get_session("wpp")
    presentation = session.get_document()
    slide = presentation.Slides(params.slide_index)
    color = resolve_color(presentation, params.color)
    filename = params.icon_name + ("-fill" if params.filled else "")
    url = f"{ICON_PACKAGE_BASE}/{params.style}/{filename}.svg"
    try:
        data, _ = assets.download_resource(url)
    except urllib.error.HTTPError as error:
        if error.code == 404:
            raise ValueError(
                f"Icon '{params.icon_name}' is unavailable in this style"
            ) from error
        raise
    svg = data.decode("utf-8").replace("currentColor", color)
    if f'fill="{color}"' not in svg:
        svg = svg.replace("<svg ", f'<svg fill="{color}" ', 1)
    goto_slide(params.slide_index)
    result = _insert_picture(
        slide,
        svg.encode("utf-8"),
        ".svg",
        params.left,
        params.top,
        params.width,
        params.height,
        True,
        params.zorder,
    )
    return {**result, "icon_name": params.icon_name, "source_url": url}


class AddPictureInput(DocumentInput):
    """Input for adding an image to a slide."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    file_path: str = Field(..., description="Path to image file")
    left: float = Field(..., description="Left position in points")
    top: float = Field(..., description="Top position in points")
    width: float | None = Field(
        default=None,
        description="Width in points (auto-scale if not provided)",
        gt=0,
    )
    height: float | None = Field(
        default=None,
        description="Height in points (auto-scale if not provided)",
        gt=0,
    )
    zorder: Literal["front", "back", "behind_text"] = Field(
        default="front", description=ZORDER_FIELD_DESCRIPTION
    )


@tool(
    "wps_ppt_pictures",
    component="wpp",
    annotations={"title": "Add Picture"},
    operation="add",
)
def add_picture(params: AddPictureInput) -> JsonObject:
    """Add an image from a file path to a slide."""
    path = local_path(params.file_path, exists=True)
    goto_slide(params.slide_index)
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    pic = slide.Shapes.AddPicture(
        FileName=str(path),
        LinkToFile=MSO_FALSE,
        SaveWithDocument=MSO_TRUE,
        Left=params.left,
        Top=params.top,
        Width=-1,
        Height=-1,
    )
    _size_picture(pic, params.left, params.top, params.width, params.height)
    return _picture_result(slide, pic, params.zorder)
