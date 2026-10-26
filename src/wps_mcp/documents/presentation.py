"""Describe presentation documents without registering tool operations."""

from pathlib import Path

from wps_mcp.common.colors import int_to_hex
from wps_mcp.errors import native_errors
from wps_mcp.models import JsonObject, NativeObject

SLIDE_SIZE_PRESETS = {"16:9": (960.0, 540.0), "4:3": (720.0, 540.0)}


def presentation_summary(document: NativeObject) -> JsonObject:
    """Read the presentation fields shared by lifecycle responses."""
    return {
        "path": str(document.Path),
        "slides_count": int(document.Slides.Count),
        "read_only": int(document.ReadOnly) == -1,
        "slide_width": float(document.PageSetup.SlideWidth),
        "slide_height": float(document.PageSetup.SlideHeight),
    }


def presentation_details(document: NativeObject) -> JsonObject:
    """Read theme metadata while preserving operational failures."""
    page = document.PageSetup
    result = presentation_summary(document)
    warnings: list[str] = []
    try:
        template_name = document.TemplateName
    except native_errors() + (AttributeError,) as error:
        template_name = None
        warnings.append(f"Template name unavailable: {error}")
    fonts: dict[str, str | None] = dict.fromkeys(
        ("title_latin", "title_east_asian", "body_latin", "body_east_asian")
    )
    try:
        try:
            font_scheme = document.Designs(1).SlideMaster.Theme.ThemeFontScheme
        except native_errors() + (AttributeError,):
            font_scheme = document.SlideMaster.Theme.ThemeFontScheme
        for prefix, native_name in (
            ("body", "MinorFont"),
            ("title", "MajorFont"),
        ):
            for language, index in (("latin", 1), ("east_asian", 3)):
                name = getattr(font_scheme, native_name)(index).Name
                fonts[f"{prefix}_{language}"] = (
                    name if name and not name.startswith("+") else None
                )
    except native_errors() + (AttributeError,) as error:
        warnings.append(f"Theme fonts unavailable: {error}")
    colors: dict[str, str | None] = {
        f"accent{index}": None for index in range(1, 7)
    }
    try:
        color_scheme = document.SlideMaster.Theme.ThemeColorScheme
        for index, name in enumerate(colors, start=5):
            colors[name] = int_to_hex(int(color_scheme(index).RGB))
    except native_errors() + (AttributeError,) as error:
        warnings.append(f"Theme colors unavailable: {error}")
    full_name = str(document.FullName)
    file_path = full_name if Path(full_name).is_absolute() else None
    result.update(
        local_path=file_path,
        local_dir=str(Path(file_path).parent) if file_path else None,
        slide_width_inches=round(page.SlideWidth / 72, 3),
        slide_height_inches=round(page.SlideHeight / 72, 3),
        first_slide_number=page.FirstSlideNumber,
        template_name=template_name,
        fonts=fonts,
        accent_colors=colors,
    )
    if warnings:
        result["warnings"] = warnings
    return result
