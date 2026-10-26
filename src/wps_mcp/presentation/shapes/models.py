"""Validated shape selectors, creation options and geometry updates."""

from __future__ import annotations

from typing import Annotated, Literal

from pydantic import Field, field_validator, model_validator

from wps_mcp.common.colors import RgbColor
from wps_mcp.models import DocumentInput
from wps_mcp.presentation.shapes.z_order import ZORDER_FIELD_DESCRIPTION

ShapeReference = (
    Annotated[str, Field(min_length=1)] | Annotated[int, Field(ge=1)]
)


class ShapeInput(DocumentInput):
    """Identify one shape, including group paths."""

    slide_index: int = Field(ge=1)
    shape_name_or_index: ShapeReference


class SelectShapesInput(DocumentInput):
    """Input for selecting multiple shapes on a slide."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_names: list[str] = Field(
        ...,
        description=(
            "Shape names or group paths; all names resolve before editing."
        ),
        min_length=1,
    )


class LockAspectRatioInput(DocumentInput):
    """Input for locking/unlocking shape aspect ratio."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: ShapeReference = Field(
        ...,
        description=(
            "Shape name, group path, or 1-based index; prefer a stable name."
        ),
    )
    locked: bool = Field(..., description="True to lock, False to unlock")


class CopyShapeToSlideInput(DocumentInput):
    """Input for copying a shape to another slide."""

    src_slide_index: int = Field(
        ..., ge=1, description="1-based source slide index"
    )
    shape_name_or_index: ShapeReference = Field(
        ...,
        description=(
            "Shape name or one-based index; names remain stable after "
            "insertions."
        ),
    )
    dst_slide_index: int = Field(
        ..., ge=1, description="1-based destination slide index"
    )


class GroupShapesInput(DocumentInput):
    """Identify distinct top-level shapes to replace with one group."""

    slide_index: int = Field(ge=1)
    shape_names: list[str] = Field(min_length=2)

    @field_validator("shape_names")
    @classmethod
    def validate_names(cls, values: list[str]) -> list[str]:
        """Reject empty or duplicate shape names before grouping."""
        if any(not name for name in values) or len(set(values)) != len(values):
            raise ValueError(
                "shape_names must contain distinct, nonempty names."
            )
        return values


class NewTextShapeInput(DocumentInput):
    """Share position, text, font and stacking options for new text shapes."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    left: float = Field(..., description="Left position in points")
    top: float = Field(..., description="Top position in points")
    width: float = Field(..., description="Width in points", gt=0)
    height: float = Field(..., description="Height in points", gt=0)
    text: str | None = Field(
        default=None,
        description=(
            "Text; newline starts a paragraph, vertical tab inserts a "
            "soft break."
        ),
    )
    font_name: str | None = Field(
        default=None,
        min_length=1,
        description="Font name applied to Latin and East Asian text.",
    )
    font_size: float | None = Field(
        default=None, description="Font size in points.", gt=0
    )
    bold: bool | None = Field(default=None, description="Bold on/off.")
    italic: bool | None = Field(default=None, description="Italic on/off.")
    font_color: RgbColor | None = Field(
        default=None,
        description="Text color '#RRGGBB'.",
    )
    align: Literal["left", "center", "right", "justify"] | None = Field(
        default=None,
        description=(
            "Paragraph alignment for shape text: 'left', 'center', "
            "'right', or 'justify'."
        ),
    )
    zorder: Literal["front", "back", "behind_text"] = Field(
        default="front", description=ZORDER_FIELD_DESCRIPTION
    )


class AddShapeInput(NewTextShapeInput):
    """Input for adding an auto shape to a slide."""

    shape_type: int | str = Field(
        ..., description="Friendly shape name or MsoAutoShapeType integer."
    )
    fill_color: RgbColor | None = Field(
        default=None,
        description=(
            "Fill color #RRGGBB; implies solid unless fill_type is given."
        ),
    )
    fill_type: Literal["solid", "none", "gradient"] | None = Field(
        default=None,
        description=(
            "Fill type: 'solid', 'none', or 'gradient'. Defaults to "
            "'solid' when fill_color is given."
        ),
    )
    fill_color2: RgbColor | None = Field(
        default=None,
        description="Gradient end color #RRGGBB.",
    )
    fill_gradient_style: (
        Literal[
            "horizontal",
            "vertical",
            "diagonal_up",
            "diagonal_down",
            "from_corner",
            "from_center",
        ]
        | None
    ) = Field(
        default=None, description="Gradient direction from GRADIENT_STYLE_MAP."
    )
    fill_transparency: float | None = Field(
        default=None,
        description=(
            "Fill transparency: 0.0 = opaque, 1.0 = fully transparent."
        ),
        ge=0,
        le=1,
    )
    line_visible: bool | None = Field(
        default=None,
        description=(
            "Border visibility. Set to false to remove the default border"
            " (recommended for most card/box shapes)."
        ),
    )
    line_color: RgbColor | None = Field(
        default=None,
        description=(
            "Border color #RRGGBB; enables the line unless explicitly "
            "disabled."
        ),
    )
    line_weight: float | None = Field(
        default=None, description="Border weight in points.", ge=0
    )
    corner_radius: float | None = Field(
        default=None,
        ge=0.0,
        le=1.0,
        description=(
            "Rounded-rectangle radius ratio 0-1; exclusive with "
            "corner_radius_pt."
        ),
    )
    corner_radius_pt: float | None = Field(
        default=None,
        gt=0.0,
        description=(
            "Rounded-rectangle radius in points; at most half the short side."
        ),
    )

    @model_validator(mode="after")
    def check_corner_radius_exclusivity(self) -> AddShapeInput:
        """Ensure corner_radius and corner_radius_pt are mutually exclusive."""
        if (
            self.corner_radius is not None
            and self.corner_radius_pt is not None
        ):
            raise ValueError(
                "corner_radius and corner_radius_pt are mutually exclusive — "
                "set one or the other, not both"
            )
        return self


class AddTextboxInput(NewTextShapeInput):
    """Input for adding a text box to a slide."""

    vertical_anchor: Literal["top", "middle", "bottom"] | None = Field(
        default=None,
        description="Vertical text anchor: 'top', 'middle', or 'bottom'.",
    )


class AddLineInput(DocumentInput):
    """Input for adding a line to a slide."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    begin_x: float = Field(..., description="Start X position in points")
    begin_y: float = Field(..., description="Start Y position in points")
    end_x: float = Field(..., description="End X position in points")
    end_y: float = Field(..., description="End Y position in points")
    zorder: Literal["front", "back", "behind_text"] = Field(
        default="front", description=ZORDER_FIELD_DESCRIPTION
    )


class ListShapesInput(DocumentInput):
    """Input for listing shapes on a slide."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")


class ShapeIdentifierInput(DocumentInput):
    """Input for identifying a shape by name or index."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name: str | None = Field(
        default=None,
        description=(
            "Shape name (preferred — indices shift when shapes are "
            "added/removed)"
        ),
    )
    shape_index: int | None = Field(
        default=None,
        ge=1,
        description="1-based shape index (unstable — prefer shape_name)",
    )

    @model_validator(mode="after")
    def validate_identifier(self) -> ShapeIdentifierInput:
        """Require exactly one shape name or one-based index."""
        if (self.shape_name is None) == (self.shape_index is None):
            raise ValueError(
                "Provide exactly one of shape_name or shape_index"
            )
        return self


class UpdateShapeInput(DocumentInput):
    """Input for updating shape properties."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name: str | None = Field(
        default=None,
        description=(
            "Shape name (preferred — indices shift when shapes are "
            "added/removed)"
        ),
    )
    shape_index: int | None = Field(
        default=None,
        ge=1,
        description="1-based shape index (unstable — prefer shape_name)",
    )
    shape_names: list[str] | None = Field(
        default=None,
        description=(
            "Shape names or group paths; all names resolve before editing."
        ),
    )
    all: bool = Field(
        default=False,
        description=(
            "Update every top-level shape except names listed in exclude."
        ),
    )
    exclude: list[str] | None = Field(
        default=None,
        description="Names to leave alone. Only with all or shape_names.",
    )
    left: float | None = Field(
        default=None, description="New left position in points"
    )
    top: float | None = Field(
        default=None, description="New top position in points"
    )
    width: float | None = Field(
        default=None, description="New width in points", ge=0
    )
    height: float | None = Field(
        default=None, description="New height in points", ge=0
    )
    rotation: float | None = Field(
        default=None, description="Rotation in degrees (0-360)"
    )
    dleft: float | None = Field(
        default=None,
        description="Horizontal offset in points; exclusive with left.",
    )
    dtop: float | None = Field(
        default=None,
        description="Vertical offset in points; exclusive with top.",
    )
    dwidth: float | None = Field(
        default=None,
        description=(
            "Widen by this many points, relative to the current width."
        ),
    )
    dheight: float | None = Field(
        default=None,
        description=(
            "Heighten by this many points, relative to the current height."
        ),
    )
    name: str | None = Field(
        default=None, description="New name for the shape"
    )
    adjustments: dict[str, float] | None = Field(
        default=None,
        description=(
            "1-based adjustment handles; inspect them with wps_ppt_shapes/get."
        ),
    )

    @model_validator(mode="after")
    def validate_adjustment_keys(self) -> UpdateShapeInput:
        """Require positive one-based adjustment handle indices."""
        if self.adjustments:
            for k in self.adjustments:
                if not k.isdecimal() or int(k) < 1:
                    raise ValueError(
                        f"Adjustment index {k} must be >= 1 (1-based indexing)"
                    )
        return self

    @model_validator(mode="after")
    def validate_selection(self) -> UpdateShapeInput:
        """Require exactly one unambiguous target selection."""
        chosen = [
            name
            for name, value in (
                ("shape_name", self.shape_name),
                ("shape_index", self.shape_index),
                ("shape_names", self.shape_names),
                ("all", self.all or None),
            )
            if value is not None
        ]
        if not chosen:
            raise ValueError(
                "Say which shapes to update: shape_name, shape_index, "
                "shape_names, or all=true"
            )
        if len(chosen) > 1:
            raise ValueError(
                f"Use one way of choosing shapes, not {' and '.join(chosen)}"
            )
        if self.shape_names is not None and (not self.shape_names):
            raise ValueError("shape_names must not be empty if provided")
        if self.exclude is not None and (not (self.all or self.shape_names)):
            raise ValueError(
                "exclude only means something with all or shape_names"
            )
        return self

    @model_validator(mode="after")
    def validate_offsets(self) -> UpdateShapeInput:
        """Reject simultaneous absolute and relative coordinates."""
        for absolute, relative in (
            ("left", "dleft"),
            ("top", "dtop"),
            ("width", "dwidth"),
            ("height", "dheight"),
        ):
            if (
                getattr(self, absolute) is not None
                and getattr(self, relative) is not None
            ):
                raise ValueError(
                    f"{absolute} and {relative} are mutually exclusive"
                )
        return self

    @model_validator(mode="after")
    def validate_single_shape_only_fields(self) -> UpdateShapeInput:
        """Limit names and adjustment handles to a single target."""
        if self.all or self.shape_names:
            for field in ("name", "adjustments"):
                if getattr(self, field) is not None:
                    raise ValueError(
                        f"{field} requires one shape; exclude all and "
                        f"shape_names"
                    )
        return self


class SetZOrderInput(ShapeIdentifierInput):
    """Input for changing shape z-order."""

    command: str = Field(
        ...,
        description=(
            "bring_to_front, send_to_back, bring_forward, "
            "send_backward, send_behind_text."
        ),
    )
