"""Normalize text line endings and provide shared formatting guidance."""

from typing import Literal


def normalize_text(
    value: str, *, newline: Literal["\r", "\n", "\r\n"] = "\r"
) -> str:
    """Convert each logical line break to the requested native separator.

    Normalize rich-text runs independently so a trailing CR and a leading
    LF in different runs remain two intentional paragraph breaks.
    """
    normalized = value.replace("\r\n", "\n").replace("\r", "\n")
    return normalized.replace("\n", newline)


def font_size_warning(size: float | None) -> str | None:
    """Return presentation readability guidance for small font sizes."""
    return (
        f"Font size {size}pt is small; "
        "use at least 16pt, or 20pt for body text."
        if size is not None and size < 16
        else None
    )
