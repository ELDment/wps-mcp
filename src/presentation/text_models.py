"""Validated inputs for presentation text editing and layout."""

from __future__ import annotations

from typing import Annotated, Literal

from pydantic import Field, field_validator, model_validator

from src.common.colors import get_theme_color_index, hex_to_int
from src.models import DocumentInput, InputModel
from src.presentation.shape_models import ShapeInput
from src.presentation.text_enums import NUMBERED_STYLE_MAP

PositiveIndex = Annotated[int, Field(ge=1)]


SlideIndices = Annotated[list[PositiveIndex], Field(min_length=1)]


class ShapeTextInput(ShapeInput):
    """Select a slide and a text-bearing shape by name, path, or index."""


class TextFormatSpec(InputModel):
    """Font properties shared by whole-frame, range, and run formatting."""

    font_name: str | None = Field(default=None, min_length=1)
    font_name_fareast: str | None = Field(default=None, min_length=1)
    font_size: float | None = Field(default=None, gt=0)
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    color: str | None = None
    font_color_theme: str | None = None
    highlight_color: str | None = None

    @field_validator("color", "highlight_color")
    @classmethod
    def validate_color(cls, value: str | None) -> str | None:
        """Validate RGB colors and reject unsupported highlight removal."""
        if value is not None:
            if value.lower() == "clear":
                raise ValueError(
                    "WPS does not support verified highlight removal"
                )
            hex_to_int(value)
        return value

    @field_validator("font_color_theme")
    @classmethod
    def validate_theme_color(cls, value: str | None) -> str | None:
        """Require a supported theme color name."""
        if value is not None:
            get_theme_color_index(value)
        return value


class SpanSpec(InputModel):
    """Select a search match or UTF-16 span; zero length inserts text."""

    search_text: str | None = Field(default=None, min_length=1)
    occurrence: PositiveIndex = 1
    start: PositiveIndex | None = Field(
        default=None,
        description=(
            "One-based UTF-16 code-unit offset; emoji usually occupies two "
            "units."
        ),
    )
    length: int | None = Field(
        default=None,
        ge=0,
        description=(
            "UTF-16 code-unit count. Zero inserts without removing text."
        ),
    )

    def check_span(self) -> None:
        """Require one search selector or a complete native UTF-16 span."""
        if self.search_text is not None:
            if self.start is not None or self.length is not None:
                raise ValueError(
                    "search_text and start/length are mutually exclusive"
                )
        elif self.start is None or self.length is None:
            raise ValueError("Provide search_text or both start and length")
        elif self.occurrence != 1:
            raise ValueError("occurrence requires search_text")


class RunSpec(TextFormatSpec):
    """One literal text fragment and its formatting."""

    text: str = Field(min_length=1)


class TextRangeSpec(TextFormatSpec, SpanSpec):
    """One span with formatting, used inside a range batch."""

    @model_validator(mode="after")
    def validate_span(self) -> TextRangeSpec:
        """Validate the selected search occurrence or native text span."""
        self.check_span()
        return self


class SetTextInput(ShapeTextInput, SpanSpec):
    """Replace a frame or span, or write a sequence of formatted runs."""

    text: str | None = None
    runs: list[RunSpec] | None = Field(default=None, min_length=1)

    @model_validator(mode="after")
    def validate_content(self) -> SetTextInput:
        """Require text or runs and validate any replacement span."""
        if (self.text is None) == (self.runs is None):
            raise ValueError("Provide exactly one of text or runs")
        if any(
            value is not None
            for value in (self.start, self.length, self.search_text)
        ):
            if self.runs is not None:
                raise ValueError(
                    "runs replaces the whole frame and cannot target a span"
                )
            self.check_span()
        elif self.occurrence != 1:
            raise ValueError("occurrence requires search_text")
        return self


class GetTextInput(ShapeTextInput):
    """Read text and optionally its rendered dimensions."""

    measure: bool = False


class FormatTextInput(ShapeTextInput, TextFormatSpec):
    """Apply font properties to every character in the shape."""


class FormatTextRangeInput(ShapeTextInput, TextFormatSpec, SpanSpec):
    """Format one span, or ordered spans with an optional frame-wide base."""

    base: TextFormatSpec | None = None
    ranges: list[TextRangeSpec] | None = Field(default=None, min_length=1)

    @model_validator(mode="after")
    def validate_ranges(self) -> FormatTextRangeInput:
        """Keep batch formatting separate from single-span options."""
        if self.ranges is None:
            if self.base is not None:
                raise ValueError("base requires ranges")
            self.check_span()
        else:
            fields = (
                "start",
                "length",
                "search_text",
                *TextFormatSpec.model_fields,
            )
            if any(getattr(self, field) is not None for field in fields):
                raise ValueError(
                    "With ranges, put formatting and spans inside each entry"
                )
            if self.occurrence != 1:
                raise ValueError("occurrence belongs inside a ranges entry")
        return self


class SetParagraphFormatInput(ShapeTextInput):
    """Set paragraph spacing in points and line spacing as a multiplier."""

    paragraph_index: PositiveIndex | None = None
    alignment: (
        Literal["left", "center", "right", "justify", "distribute"] | None
    ) = None
    line_spacing: float | None = Field(default=None, gt=0)
    space_before: float | None = Field(default=None, ge=0)
    space_after: float | None = Field(default=None, ge=0)
    indent_level: int | None = Field(default=None, ge=1, le=9)
    first_line_indent: float | None = None


class SetBulletInput(ShapeTextInput):
    """Set bullet numbering, nesting, and appearance on selected paragraphs."""

    paragraph_index: PositiveIndex | None = None
    bullet_type: Literal["none", "unnumbered", "numbered"]
    bullet_char: str | None = Field(default=None, min_length=1, max_length=1)
    bullet_start_value: PositiveIndex | None = None
    indent_level: int | None = Field(default=None, ge=1, le=9)
    color: str | None = None
    size: float | None = Field(default=None, ge=0.25, le=4.0)
    font_name: str | None = Field(default=None, min_length=1)
    numbered_style: str | None = None
    use_text_color: bool | None = None
    use_text_font: bool | None = None

    @model_validator(mode="after")
    def validate_bullet(self) -> SetBulletInput:
        """Validate colors and resolve the requested numbering style."""
        if self.color is not None:
            hex_to_int(self.color)
        if self.numbered_style is not None:
            if self.numbered_style not in NUMBERED_STYLE_MAP:
                raise ValueError(
                    f"Unknown numbered_style: {self.numbered_style}"
                )
            self.bullet_type = "numbered"
        return self


class FindReplaceTextInput(DocumentInput):
    """Search text frames; group children are included only when requested."""

    find_text: str = Field(min_length=1)
    replace_text: str | None = None
    dry_run: bool = False
    match_case: bool = False
    whole_words: bool = False
    slide_indices: SlideIndices | None = None
    shape_name: str | None = Field(default=None, min_length=1)
    context_chars: int = Field(
        default=0,
        ge=0,
        le=500,
        description=(
            "Unicode characters before and after a match; never splits an "
            "emoji."
        ),
    )
    include_groups: bool = False


class SetTextframeInput(ShapeTextInput):
    """Set margins in points, wrapping, flow, and auto-fit behavior."""

    auto_size: Literal["none", "shape_to_fit", "shrink_to_fit"] | None = None
    word_wrap: bool | None = None
    margin_left: float | None = Field(default=None, ge=0)
    margin_right: float | None = Field(default=None, ge=0)
    margin_top: float | None = Field(default=None, ge=0)
    margin_bottom: float | None = Field(default=None, ge=0)
    orientation: (
        Literal["horizontal", "vertical", "upward", "downward"] | None
    ) = None
    vertical_anchor: Literal["top", "middle", "bottom"] | None = None


class GetAllTextInput(DocumentInput):
    """Extract a reading-order Markdown view, optionally to a UTF-8 file."""

    slide_indices: SlideIndices | None = None
    output_path: str | None = Field(default=None, min_length=1)


class CheckTypographyInput(DocumentInput):
    """Detect layout issues; optionally widen shapes or insert soft returns."""

    slide_index: PositiveIndex | None = None
    max_chars: int = Field(default=3, ge=1, le=10)
    max_words: int = Field(default=2, ge=1, le=5)
    fix: bool = False
    max_expand_pt: float = Field(default=20.0, ge=1, le=50)
