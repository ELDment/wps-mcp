"""Convert RGB colors and resolve Office theme color indices."""

import re
from typing import Annotated

from pydantic import AfterValidator, Field

THEME_COLOR_MAP = {
    "dark1": 1,
    "light1": 2,
    "dark2": 3,
    "light2": 4,
    "accent1": 5,
    "accent2": 6,
    "accent3": 7,
    "accent4": 8,
    "accent5": 9,
    "accent6": 10,
    "hyperlink": 11,
    "followed_hyperlink": 12,
}


def rgb_to_int(red: int, green: int, blue: int) -> int:
    """Pack three byte-sized RGB channels into an Office color integer."""
    if any(
        type(channel) is not int or not 0 <= channel <= 255
        for channel in (red, green, blue)
    ):
        raise ValueError("RGB channels must be integers between 0 and 255")
    return red | green << 8 | blue << 16


def int_to_rgb(value: int) -> tuple[int, int, int]:
    """Unpack an Office color integer into red, green and blue channels."""
    return value & 255, value >> 8 & 255, value >> 16 & 255


def hex_to_rgb(value: str) -> tuple[int, int, int]:
    """Parse a three- or six-digit RGB color, with an optional hash."""
    if not re.fullmatch(r"#?(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})", value):
        raise ValueError(
            "Color must be a hexadecimal RGB value, such as #156082"
        )
    raw = value.removeprefix("#")
    if len(raw) == 3:
        raw = "".join(channel * 2 for channel in raw)
    return int(raw[:2], 16), int(raw[2:4], 16), int(raw[4:], 16)


def hex_to_int(value: str) -> int:
    """Convert a hexadecimal RGB color to an Office color integer."""
    return rgb_to_int(*hex_to_rgb(value))


def int_to_hex(value: int) -> str:
    """Format an Office color integer as an uppercase RGB hex string."""
    red, green, blue = int_to_rgb(value)
    return f"#{red:02X}{green:02X}{blue:02X}"


def get_theme_color_index(name: str) -> int:
    """Resolve a named Office theme color to its one-based index."""
    key = name.lower().replace(" ", "_").replace("-", "_")
    if key not in THEME_COLOR_MAP:
        raise ValueError(
            f"Unknown theme color: {name}; "
            f"choose from {', '.join(THEME_COLOR_MAP)}"
        )
    return THEME_COLOR_MAP[key]


def normalize_rgb(value: str) -> str:
    """Normalize a validated RGB input to uppercase six-digit notation."""
    return int_to_hex(hex_to_int(value))


RgbColor = Annotated[
    str,
    Field(pattern=r"^#?(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})$"),
    AfterValidator(normalize_rgb),
]
