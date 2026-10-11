"""Define shared Writer positions, ranges, and formatting inputs."""

from typing import Annotated, Literal, Self

from pydantic import Field, model_validator

from src.models import (
    Color,
    DocumentInput,
    InputModel,
)

Position = Annotated[int, Field(ge=0, strict=True)]


Index = Annotated[int, Field(ge=1, strict=True)]


PositivePoints = Annotated[float, Field(gt=0, allow_inf_nan=False)]


class RangeInput(DocumentInput):
    """Select a range, defaulting an omitted end to the document end."""

    start: Position = 0
    end: Position | None = None

    @model_validator(mode="after")
    def validate_range(self) -> Self:
        """Reject reversed native text intervals."""
        if self.end is not None and self.end < self.start:
            raise ValueError("end must be greater than or equal to start")
        return self


class ExplicitRangeInput(DocumentInput):
    """Require an explicit end before replacing, deleting or anchoring text."""

    start: Position = 0
    end: Position

    @model_validator(mode="after")
    def validate_range(self) -> Self:
        """Reject reversed native text intervals."""
        if self.end < self.start:
            raise ValueError("end must be greater than or equal to start")
        return self


class CharacterFormat(InputModel):
    """Specify optional Writer font properties in points and RGB."""

    font_name: str | None = Field(default=None, min_length=1)
    font_name_far_east: str | None = Field(default=None, min_length=1)
    font_size: PositivePoints | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: Literal["none", "single", "double"] | None = None
    color: Color | None = None


class ParagraphFormat(InputModel):
    """Specify Writer paragraph alignment, spacing, and pagination."""

    alignment: (
        Literal["left", "center", "right", "justify", "distribute"] | None
    ) = None
    left_indent: float | None = None
    right_indent: float | None = None
    first_line_indent: float | None = None
    space_before: Annotated[float, Field(ge=0)] | None = None
    space_after: Annotated[float, Field(ge=0)] | None = None
    line_spacing: PositivePoints | None = Field(
        default=None, description="Exact line spacing in points"
    )
    keep_with_next: bool | None = None
    keep_together: bool | None = None
    page_break_before: bool | None = None
