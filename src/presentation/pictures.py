"""Picture editing, bounded downloads and Material icons."""

from __future__ import annotations

import json
import logging
import threading
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from contextlib import closing
from pathlib import Path
from typing import Literal
from urllib.parse import urlsplit

from pydantic import Field, model_validator

from src.common.colors import hex_to_int, int_to_hex
from src.models import DocumentInput, InputModel, JsonObject, NativeObject
from src.native.files import temporary_directory
from src.native.objects import optional_property
from src.native.runtime import get_session
from src.native.state import temporary_attributes
from src.presentation.constants import (
    ICON_PACKAGE_BASE,
    MSO_FALSE,
    MSO_LINKED_PICTURE,
    MSO_PICTURE,
    MSO_TRUE,
    PICTURE_COLOR_TYPE_MAP,
    PICTURE_COLOR_TYPE_NAMES,
)
from src.presentation.navigation import goto_slide
from src.presentation.shape_formatting import resolve_color
from src.presentation.shape_lookup import resolve_shape as _get_shape
from src.presentation.shape_types import SHAPE_NAME_MAP
from src.presentation.zorder import (
    ZORDER_FIELD_DESCRIPTION,
    place_in_zorder,
)
from src.protocol.definitions import tool

logger = logging.getLogger(__name__)
HTTP_TIMEOUT_SECONDS = 15
MAX_DOWNLOAD_BYTES = 20 * 1024 * 1024
MAX_METADATA_BYTES = 8 * 1024 * 1024


def _download_resource(
    url: str, *, max_bytes: int = MAX_DOWNLOAD_BYTES, method: str = "GET"
) -> tuple[bytes, str]:
    """Download a bounded HTTP response and close it on every exit path."""
    parsed = urlsplit(url)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        raise ValueError("Only absolute HTTP or HTTPS URLs are supported")
    request = urllib.request.Request(url, method=method)
    try:
        with urllib.request.urlopen(
            request, timeout=HTTP_TIMEOUT_SECONDS
        ) as response:
            if urlsplit(response.url).scheme not in {"http", "https"}:
                raise ValueError("The response redirected to a non-HTTP URL")
            content_type = response.headers.get("Content-Type", "")
            if method == "HEAD":
                return (b"", content_type)
            length = response.headers.get("Content-Length")
            if length is not None and int(length) > max_bytes:
                raise ValueError(
                    f"Response exceeds the {max_bytes}-byte limit"
                )
            data = response.read(max_bytes + 1)
            if len(data) > max_bytes:
                raise ValueError(
                    f"Response exceeds the {max_bytes}-byte limit"
                )
            return (data, content_type)
    except urllib.error.HTTPError as error:
        with closing(error):
            raise


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


_icon_cache: list[JsonObject] | None = None
_icon_cache_time = 0.0
_ICON_CACHE_TTL = 86400
_icon_cache_lock = threading.Lock()
_ICON_METADATA_URL = "https://fonts.google.com/metadata/icons"


def _fetch_icon_metadata() -> list[JsonObject]:
    """Fetch public icon metadata only when a caller explicitly searches."""
    global _icon_cache, _icon_cache_time
    with _icon_cache_lock:
        now = time.monotonic()
        if (
            _icon_cache is not None
            and now - _icon_cache_time < _ICON_CACHE_TTL
        ):
            return _icon_cache
        payload, _ = _download_resource(
            _ICON_METADATA_URL, max_bytes=MAX_METADATA_BYTES
        )
        raw = payload.decode("utf-8")
        if raw.startswith(")]}'"):
            raw = raw.partition("\n")[2]
        metadata = json.loads(raw)
        if not isinstance(metadata, dict) or not isinstance(
            metadata.get("icons"), list
        ):
            raise ValueError("Invalid icon metadata response")
        icons = metadata["icons"]
        if any(
            not isinstance(icon, dict) or not isinstance(icon.get("name"), str)
            for icon in icons
        ):
            raise ValueError("Invalid icon entry in metadata response")
        _icon_cache = icons
        _icon_cache_time = time.monotonic()
        return icons


_icon_on_cdn_cache: dict[str, bool] = {}


def _icon_is_on_cdn(name: str) -> bool:
    """Exclude known missing icons; leave transient availability unresolved."""
    with _icon_cache_lock:
        cached = _icon_on_cdn_cache.get(name)
    if cached is not None:
        return cached
    try:
        _download_resource(
            f"{ICON_PACKAGE_BASE}/outlined/{name}.svg", method="HEAD"
        )
        available = True
    except urllib.error.HTTPError as error:
        if error.code != 404:
            return True
        available = False
    except (urllib.error.URLError, TimeoutError, OSError):
        return True
    with _icon_cache_lock:
        _icon_on_cdn_cache[name] = available
    return available


class CropPictureInput(DocumentInput):
    """Input for cropping a picture shape."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: str | int = Field(
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
    shape_name_or_index: str | int = Field(
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
    transparent_color: str | None = Field(
        default=None,
        pattern="^#[0-9A-Fa-f]{6}$",
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
    svg_color: str | None = Field(
        default=None,
        description=(
            "Replace 'currentColor' in SVG files with this color (e.g. "
            "'#1A73E8'). Only applies to SVG files."
        ),
        pattern="^#[0-9A-Fa-f]{6}$",
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


class SearchIconsInput(InputModel):
    """Input for searching Material Symbols icons by keyword."""

    query: str = Field(
        ...,
        description="Icon search keywords; multiple words narrow the search.",
        min_length=1,
    )
    max_results: int = Field(
        default=20,
        ge=1,
        le=100,
        description="Maximum number of results to return (default: 20)",
    )


@tool(
    "wps_ppt_pictures",
    component="wpp",
    annotations={"title": "Crop Picture", "idempotentHint": True},
    operation="crop",
)
def crop_picture(params: CropPictureInput) -> JsonObject:
    """Crop a picture shape — rectangular trim and/or crop-to-shape."""
    app = get_session("wpp").get_application()
    goto_slide(app, params.slide_index)
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
        except Exception as e:
            raise ValueError(
                f"Could not apply crop_shape {auto_shape_int}: {e}"
            ) from e
    if params.corner_radius_pt is not None:
        try:
            current_type = shape.AutoShapeType
        except Exception:
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
    except Exception:
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
    app = get_session("wpp").get_application()
    goto_slide(app, params.slide_index)
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
    data, content_type = _download_resource(params.url)
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
    goto_slide(session.get_application(), params.slide_index)
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
        data, _ = _download_resource(url)
    except urllib.error.HTTPError as error:
        if error.code == 404:
            raise ValueError(
                f"Icon '{params.icon_name}' is unavailable in this style"
            ) from error
        raise
    svg = data.decode("utf-8").replace("currentColor", color)
    if f'fill="{color}"' not in svg:
        svg = svg.replace("<svg ", f'<svg fill="{color}" ', 1)
    goto_slide(session.get_application(), params.slide_index)
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


@tool(
    "wps_ppt_pictures",
    component=None,
    annotations={
        "title": "Search Material Icons",
        "readOnlyHint": True,
        "idempotentHint": True,
        "openWorldHint": True,
    },
    operation="search_icons",
)
def search_icons(params: SearchIconsInput) -> JsonObject:
    """Search Google's Material Symbols icon library by keyword."""
    icons = _fetch_icon_metadata()
    query_lower = params.query.lower().strip()
    query_words = query_lower.split()
    results = []
    for icon in icons:
        name = icon.get("name", "")
        tags = [t.lower() for t in icon.get("tags", [])]
        categories = [c.lower() for c in icon.get("categories", [])]
        popularity = icon.get("popularity", 0)
        score = 0
        if name == query_lower:
            score += 100
        elif query_lower.replace(" ", "_") == name:
            score += 90
        elif query_lower in name:
            score += 50
        if len(query_words) > 1 and all(w in name for w in query_words):
            score += 40
        for word in query_words:
            if word in name:
                score += 30
            for tag in tags:
                if word == tag:
                    score += 20
                elif word in tag:
                    score += 10
            for cat in categories:
                if word in cat:
                    score += 5
        if score > 0:
            score += min(popularity / 1000, 10)
            results.append(
                {
                    "name": name,
                    "categories": icon.get("categories", []),
                    "tags": icon.get("tags", [])[:8],
                    "popularity": popularity,
                    "score": round(score, 2),
                }
            )
    results.sort(key=lambda x: x["score"], reverse=True)
    kept: list[JsonObject] = []
    with ThreadPoolExecutor(max_workers=8) as pool:
        for start in range(0, len(results), params.max_results + 4):
            batch = results[start : start + params.max_results + 4]
            for icon, available in zip(
                batch,
                pool.map(_icon_is_on_cdn, [item["name"] for item in batch]),
                strict=True,
            ):
                if available:
                    kept.append(icon)
            if len(kept) >= params.max_results:
                break
    kept = kept[: params.max_results]
    return {
        "success": True,
        "query": params.query,
        "count": len(kept),
        "icons": kept,
    }


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
    app = get_session("wpp").get_application()
    goto_slide(app, params.slide_index)
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    pic = slide.Shapes.AddPicture(
        FileName=params.file_path,
        LinkToFile=MSO_FALSE,
        SaveWithDocument=MSO_TRUE,
        Left=params.left,
        Top=params.top,
        Width=-1,
        Height=-1,
    )
    _size_picture(pic, params.left, params.top, params.width, params.height)
    return _picture_result(slide, pic, params.zorder)
