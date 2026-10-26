"""Presentation themes, accessible palettes, and footer settings."""

from __future__ import annotations

import colorsys
from pathlib import Path
from xml.etree.ElementTree import Element, SubElement, tostring

from pydantic import Field

from src.common.colors import THEME_COLOR_MAP, hex_to_int, int_to_hex
from src.models import DocumentInput, JsonObject, NativeObject
from src.native.files import temporary_directory
from src.native.objects import supports_feature
from src.native.runtime import get_session
from src.presentation.constants import (
    MSO_FALSE,
    MSO_GROUP,
    MSO_TRUE,
)
from src.protocol.definitions import tool


class ReplaceFontInput(DocumentInput):
    """Input for replacing a font throughout the presentation."""

    original_font: str = Field(
        ..., description="Font name to replace", min_length=1
    )
    replacement_font: str = Field(
        ..., description="New font name", min_length=1
    )


class SetDefaultFontsInput(DocumentInput):
    """Input for setting default fonts for the entire presentation."""

    latin: str | None = Field(
        default=None, description="Default Latin font name."
    )
    east_asian: str | None = Field(
        default=None, description="Default East Asian font name."
    )
    apply_to_existing: bool = Field(
        default=True,
        description="Apply the fonts to existing text as well as the theme.",
    )


@tool(
    "wps_ppt_themes",
    component="wpp",
    annotations={"title": "Replace Font", "idempotentHint": True},
    operation="replace_font",
)
def replace_font(params: ReplaceFontInput) -> JsonObject:
    """Replace a font throughout the selected presentation."""
    pres = get_session("wpp").get_document()
    pres.Fonts.Replace(params.original_font, params.replacement_font)
    result = {
        "success": True,
        "original_font": params.original_font,
        "replacement_font": params.replacement_font,
    }
    return result


@tool(
    "wps_ppt_themes",
    component="wpp",
    annotations={
        "title": "List Fonts",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="fonts",
)
def list_fonts(params: DocumentInput) -> JsonObject:
    """List all fonts used in the active presentation."""
    pres = get_session("wpp").get_document()
    fonts = []
    for i in range(1, pres.Fonts.Count + 1):
        fonts.append(pres.Fonts(i).Name)
    result = {"success": True, "fonts_count": len(fonts), "fonts": fonts}
    return result


@tool(
    "wps_ppt_themes",
    component="wpp",
    annotations={"title": "Set Default Fonts", "idempotentHint": True},
    operation="set_fonts",
)
def set_default_fonts(params: SetDefaultFontsInput) -> JsonObject:
    """Set separate Latin and East Asian fonts across the deck."""
    if not params.latin and (not params.east_asian):
        raise ValueError(
            "At least one of 'latin' or 'east_asian' must be provided"
        )
    pres = get_session("wpp").get_document()
    for index in range(1, pres.Designs.Count + 1):
        scheme = pres.Designs(index).SlideMaster.Theme.ThemeFontScheme
        if params.latin:
            scheme.MajorFont(1).Name = params.latin
            scheme.MinorFont(1).Name = params.latin
        if params.east_asian:
            scheme.MajorFont(3).Name = params.east_asian
            scheme.MinorFont(3).Name = params.east_asian

    def apply_to_shape(shape: NativeObject) -> int:
        if shape.Type == MSO_GROUP:
            return sum(
                apply_to_shape(shape.GroupItems(index))
                for index in range(1, shape.GroupItems.Count + 1)
            )
        if not shape.HasTextFrame:
            return 0
        font = shape.TextFrame.TextRange.Font
        if params.latin:
            font.Name = params.latin
        if params.east_asian:
            font.NameFarEast = params.east_asian
        return 1

    shapes_updated = 0
    if params.apply_to_existing:
        for index in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(index)
            for shape_index in range(1, slide.Shapes.Count + 1):
                shapes_updated += apply_to_shape(slide.Shapes(shape_index))
    result = {"success": True, "theme_updated": pres.Designs.Count > 0}
    if params.latin:
        result["latin"] = params.latin
    if params.east_asian:
        result["east_asian"] = params.east_asian
    if params.apply_to_existing:
        result["slides_processed"] = pres.Slides.Count
        result["shapes_updated"] = shapes_updated
    return result


THEME_COLOR_NAMES = {index: name for name, index in THEME_COLOR_MAP.items()}
PRESET_PALETTES: dict[str, dict[str, str]] = {
    "corporate_blue": {
        "dark1": "#1B2A4A",
        "light1": "#FFFFFF",
        "dark2": "#44546A",
        "light2": "#F2F2F2",
        "accent1": "#2B579A",
        "accent2": "#BF4B28",
        "accent3": "#2E7D32",
        "accent4": "#7B3FA0",
        "accent5": "#C4652A",
        "accent6": "#1A7A8A",
    },
    "executive": {
        "dark1": "#1C1C1C",
        "light1": "#FFFFFF",
        "dark2": "#3D3D3D",
        "light2": "#F5F5F0",
        "accent1": "#2F4858",
        "accent2": "#6B3A3A",
        "accent3": "#4A6741",
        "accent4": "#5B4A6E",
        "accent5": "#7A5C3E",
        "accent6": "#3A5F6F",
    },
    "consulting": {
        "dark1": "#0A1628",
        "light1": "#FFFFFF",
        "dark2": "#2C3E50",
        "light2": "#F0F3F7",
        "accent1": "#003A70",
        "accent2": "#00796B",
        "accent3": "#6A1B9A",
        "accent4": "#B71C1C",
        "accent5": "#2E5090",
        "accent6": "#455A64",
    },
    "tailwind": {
        "dark1": "#1E293B",
        "light1": "#FFFFFF",
        "dark2": "#334155",
        "light2": "#F1F5F9",
        "accent1": "#1D4ED8",
        "accent2": "#B91C1C",
        "accent3": "#15803D",
        "accent4": "#7E22CE",
        "accent5": "#B45309",
        "accent6": "#0F766E",
    },
    "chakra": {
        "dark1": "#1A202C",
        "light1": "#FFFFFF",
        "dark2": "#2D3748",
        "light2": "#EDF2F7",
        "accent1": "#2C5282",
        "accent2": "#9B2C2C",
        "accent3": "#276749",
        "accent4": "#553C9A",
        "accent5": "#9C4221",
        "accent6": "#285E61",
    },
    "open_color": {
        "dark1": "#212529",
        "light1": "#FFFFFF",
        "dark2": "#343A40",
        "light2": "#F1F3F5",
        "accent1": "#1864AB",
        "accent2": "#C92A2A",
        "accent3": "#2F9E44",
        "accent4": "#6741D9",
        "accent5": "#E8590C",
        "accent6": "#099268",
    },
    "radix": {
        "dark1": "#11181C",
        "light1": "#FFFFFF",
        "dark2": "#1C2024",
        "light2": "#F0F2F4",
        "accent1": "#006ADC",
        "accent2": "#D31E66",
        "accent3": "#18794E",
        "accent4": "#793AAF",
        "accent5": "#BD4B00",
        "accent6": "#067A6F",
    },
    "ocean": {
        "dark1": "#0A1929",
        "light1": "#FFFFFF",
        "dark2": "#132F4C",
        "light2": "#E3F2FD",
        "accent1": "#0D47A1",
        "accent2": "#006064",
        "accent3": "#01579B",
        "accent4": "#0277BD",
        "accent5": "#00695C",
        "accent6": "#1A237E",
    },
    "forest": {
        "dark1": "#1B2418",
        "light1": "#FFFFFF",
        "dark2": "#33402E",
        "light2": "#F1F4EC",
        "accent1": "#2E7D32",
        "accent2": "#5D4037",
        "accent3": "#1B5E20",
        "accent4": "#795548",
        "accent5": "#827717",
        "accent6": "#4E342E",
    },
    "sunset": {
        "dark1": "#2C1810",
        "light1": "#FFFFFF",
        "dark2": "#4E342E",
        "light2": "#FFF3E0",
        "accent1": "#C62828",
        "accent2": "#BF360C",
        "accent3": "#AD1457",
        "accent4": "#880E4F",
        "accent5": "#6A1B9A",
        "accent6": "#B71C1C",
    },
    "sage": {
        "dark1": "#263238",
        "light1": "#FFFFFF",
        "dark2": "#37474F",
        "light2": "#F1F4EC",
        "accent1": "#558B2F",
        "accent2": "#6D4C41",
        "accent3": "#33691E",
        "accent4": "#4E6B45",
        "accent5": "#8D6E63",
        "accent6": "#5D6B3C",
    },
    "nord_light": {
        "dark1": "#2E3440",
        "light1": "#FFFFFF",
        "dark2": "#3B4252",
        "light2": "#ECEFF4",
        "accent1": "#3868A6",
        "accent2": "#A3394B",
        "accent3": "#6B5DAD",
        "accent4": "#2C7A5D",
        "accent5": "#B06D2F",
        "accent6": "#2E6C8F",
    },
    "pastel_deep": {
        "dark1": "#2D2D2D",
        "light1": "#FFFFFF",
        "dark2": "#4A4A4A",
        "light2": "#FAF8F5",
        "accent1": "#5B72A8",
        "accent2": "#A85B6E",
        "accent3": "#6B8F5E",
        "accent4": "#8B6BA8",
        "accent5": "#B07D4F",
        "accent6": "#4F8F8B",
    },
    "swiss": {
        "dark1": "#000000",
        "light1": "#FFFFFF",
        "dark2": "#1A1A1A",
        "light2": "#F5F5F5",
        "accent1": "#CC0000",
        "accent2": "#000000",
        "accent3": "#004F9E",
        "accent4": "#6B6B6B",
        "accent5": "#8B0000",
        "accent6": "#2B5B2B",
    },
    "vivid": {
        "dark1": "#0D0D0D",
        "light1": "#FFFFFF",
        "dark2": "#262626",
        "light2": "#F5F5F5",
        "accent1": "#0050D0",
        "accent2": "#C60040",
        "accent3": "#007A33",
        "accent4": "#7B00B5",
        "accent5": "#B85000",
        "accent6": "#00787C",
    },
    "rainbow": {
        "dark1": "#1A1A1A",
        "light1": "#FFFFFF",
        "dark2": "#333333",
        "light2": "#F7F7F7",
        "accent1": "#C62828",
        "accent2": "#C45000",
        "accent3": "#2E7D32",
        "accent4": "#1565C0",
        "accent5": "#6A1B9A",
        "accent6": "#AD1457",
    },
    "neon_safe": {
        "dark1": "#0F0F1A",
        "light1": "#FFFFFF",
        "dark2": "#1A1A2E",
        "light2": "#F0F0FA",
        "accent1": "#0055CC",
        "accent2": "#CC0066",
        "accent3": "#008844",
        "accent4": "#7700BB",
        "accent5": "#CC5500",
        "accent6": "#007788",
    },
}


def contrast_ratio(first: str, second: str) -> float:
    """Return WCAG contrast after validating both RGB colors."""
    luminances = []
    for value in (first, second):
        rgb = int_to_hex(hex_to_int(value)).removeprefix("#")
        channels = [
            int(rgb[offset : offset + 2], 16) / 255 for offset in (0, 2, 4)
        ]
        linear = [
            channel / 12.92
            if channel <= 0.04045
            else ((channel + 0.055) / 1.055) ** 2.4
            for channel in channels
        ]
        luminances.append(
            sum(
                channel * weight
                for channel, weight in zip(
                    linear, (0.2126, 0.7152, 0.0722), strict=True
                )
            )
        )
    return (max(luminances) + 0.05) / (min(luminances) + 0.05)


def _hsl_to_hex(hue: float, saturation: float, lightness: float) -> str:
    channels = colorsys.hls_to_rgb(hue, lightness, saturation)
    return "#" + "".join(
        f"{int(channel * 255 + 0.5):02X}" for channel in channels
    )


def generate_palette_from_primary(primary: str) -> dict[str, str]:
    """Generate a palette whose accents reach 3:1 contrast on white."""
    rgb = int_to_hex(hex_to_int(primary)).removeprefix("#")
    channels = [
        int(rgb[offset : offset + 2], 16) / 255 for offset in (0, 2, 4)
    ]
    hue, lightness, saturation = colorsys.rgb_to_hls(*channels)
    palette = {
        "dark1": _hsl_to_hex(hue, max(0.3, saturation * 0.6), 0.12),
        "light1": "#FFFFFF",
        "dark2": _hsl_to_hex(hue, max(0.2, saturation * 0.4), 0.22),
        "light2": _hsl_to_hex(hue, max(0.1, saturation * 0.3), 0.95),
    }
    for index, offset in enumerate((0, 30, 120, 180, 210, 330), 1):
        accent_hue = (hue + offset / 360) % 1
        accent_lightness = max(0.25, min(0.45, lightness))
        accent_saturation = max(0.5, saturation)
        color = _hsl_to_hex(accent_hue, accent_saturation, accent_lightness)
        while contrast_ratio(color, "#FFFFFF") < 3.0 and accent_lightness > 0:
            accent_lightness = max(0.0, accent_lightness - 0.01)
            color = _hsl_to_hex(
                accent_hue, accent_saturation, accent_lightness
            )
        palette[f"accent{index}"] = color
    return palette


class ApplyThemeInput(DocumentInput):
    """Apply an existing .thmx theme or themed presentation file."""

    theme_path: str = Field(min_length=1)


class GetThemeColorsInput(DocumentInput):
    """Read the twelve slots in the first design's theme."""


class SetThemeColorsInput(DocumentInput):
    """Choose explicit colors, a preset or a primary palette.

    Explicit slots override a preset, which overrides generated colors.
    """

    preset: str | None = None
    primary: str | None = None
    dark1: str | None = None
    light1: str | None = None
    dark2: str | None = None
    light2: str | None = None
    accent1: str | None = None
    accent2: str | None = None
    accent3: str | None = None
    accent4: str | None = None
    accent5: str | None = None
    accent6: str | None = None
    hyperlink: str | None = None
    followed_hyperlink: str | None = None


class SetHeadersFootersInput(DocumentInput):
    """Set footers, slide numbers and dates across all slides."""

    footer_text: str | None = None
    footer_visible: bool | None = None
    slide_number_visible: bool | None = None
    date_visible: bool | None = None
    date_format: int | None = Field(default=None, ge=1, le=13)
    date_fixed_text: str | None = None


@tool("wps_ppt_themes", component="wpp", operation="apply")
def apply_theme(params: ApplyThemeInput) -> JsonObject:
    """Apply a theme file to the selected presentation."""
    path = Path(params.theme_path).resolve()
    if not path.is_file():
        raise ValueError(f"Theme file not found: {path}")
    get_session("wpp").get_document().ApplyTheme(str(path))
    return {"success": True, "theme_path": str(path)}


@tool(
    "wps_ppt_themes",
    component="wpp",
    annotations={"readOnlyHint": True},
    operation="colors",
)
def get_theme_colors(params: GetThemeColorsInput) -> JsonObject:
    """Read all twelve colors from the first slide master."""
    presentation = get_session("wpp").get_document()
    if not presentation.Designs.Count:
        raise ValueError("The presentation has no slide designs")
    scheme = presentation.Designs(1).SlideMaster.Theme.ThemeColorScheme
    colors = [
        {
            "index": index,
            "name": name,
            "color_hex": int_to_hex(scheme(index).RGB),
        }
        for index, name in sorted(THEME_COLOR_NAMES.items())
    ]
    return {
        "success": True,
        "colors": colors,
        "design_count": presentation.Designs.Count,
    }


def _load_theme_colors(scheme: NativeObject, colors: dict[int, int]) -> None:
    """Import a native color scheme where the RGB setter is unavailable."""
    tags = (
        "dk1",
        "lt1",
        "dk2",
        "lt2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hlink",
        "folHlink",
    )
    root = Element(
        "a:clrScheme",
        {
            "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "name": "WPS MCP",
        },
    )
    for slot, tag in enumerate(tags, 1):
        color = colors[slot] if slot in colors else scheme(slot).RGB
        node = SubElement(root, f"a:{tag}")
        SubElement(node, "a:srgbClr", val=int_to_hex(color).lstrip("#"))
    with temporary_directory(prefix="wps-theme-") as directory:
        path = Path(directory) / "colors.xml"
        path.write_bytes(
            tostring(root, encoding="utf-8", xml_declaration=True)
        )
        scheme.Load(str(path))
    if any(scheme(slot).RGB != value for slot, value in colors.items()):
        raise RuntimeError("WPS theme colors differ after native Load")


@tool("wps_ppt_themes", component="wpp", operation="set_colors")
def set_theme_colors(params: SetThemeColorsInput) -> JsonObject:
    """Apply validated colors to every slide master.

    Explicit slots override a preset, which overrides generated colors.
    """
    merged = (
        generate_palette_from_primary(params.primary)
        if params.primary is not None
        else {}
    )
    preset = params.preset.lower() if params.preset is not None else None
    if preset is not None:
        if preset not in PRESET_PALETTES:
            raise ValueError(
                f"Unknown preset {params.preset!r}; "
                f"available: {sorted(PRESET_PALETTES)}"
            )
        merged.update(PRESET_PALETTES[preset])
    for name in THEME_COLOR_MAP:
        value = getattr(params, name)
        if value is not None:
            merged[name] = value
    if not merged:
        raise ValueError("No colors specified")
    colors = {
        THEME_COLOR_MAP[name]: hex_to_int(value)
        for name, value in merged.items()
    }
    presentation = get_session("wpp").get_document()
    design_count = presentation.Designs.Count
    if not design_count:
        raise ValueError("The presentation has no slide designs")
    for index in range(1, design_count + 1):
        scheme = presentation.Designs(index).SlideMaster.Theme.ThemeColorScheme
        if not supports_feature(scheme, "theme_color_assignment"):
            _load_theme_colors(scheme, colors)
        else:
            for slot, value in colors.items():
                scheme(slot).RGB = value
    changed = [
        {"name": THEME_COLOR_NAMES[slot], "color_hex": int_to_hex(value)}
        for slot, value in colors.items()
    ]
    result: JsonObject = {
        "success": True,
        "changed": changed,
        "changed_count": len(changed),
        "designs_updated": design_count,
    }
    if preset is not None:
        result["preset"] = preset
    if params.primary is not None:
        result["primary"] = params.primary
    return result


@tool("wps_ppt_themes", component="wpp", operation="headers_footers")
def set_headers_footers(params: SetHeadersFootersInput) -> JsonObject:
    """Set slide footers; fixed date text overrides its format."""
    presentation = get_session("wpp").get_document()
    for index in range(1, presentation.Slides.Count + 1):
        settings = presentation.Slides(index).HeadersFooters
        if params.footer_visible is not None:
            settings.Footer.Visible = (
                MSO_TRUE if params.footer_visible else MSO_FALSE
            )
        if params.footer_text is not None:
            settings.Footer.Visible = MSO_TRUE
            settings.Footer.Text = params.footer_text
        if params.slide_number_visible is not None:
            settings.SlideNumber.Visible = (
                MSO_TRUE if params.slide_number_visible else MSO_FALSE
            )
        if params.date_visible is not None:
            settings.DateAndTime.Visible = (
                MSO_TRUE if params.date_visible else MSO_FALSE
            )
        if params.date_format is not None:
            settings.DateAndTime.UseFormat = MSO_TRUE
            settings.DateAndTime.Format = params.date_format
        if params.date_fixed_text is not None:
            settings.DateAndTime.UseFormat = MSO_FALSE
            settings.DateAndTime.Text = params.date_fixed_text
    return {"success": True, "slides_updated": presentation.Slides.Count}
