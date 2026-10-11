"""Validated visual style options and native formatting operations."""

from __future__ import annotations

from typing import Annotated, Literal

from pydantic import Field, field_validator

from src.common.colors import get_theme_color_index, hex_to_int, int_to_hex
from src.models import InputModel, JsonObject, NativeObject
from src.presentation.constants import (
    DASH_STYLE_DESCRIPTION,
    GRADIENT_STYLE_MAP,
    MSO_FALSE,
    MSO_TRUE,
    check_dash_style,
    dash_style_value,
)

HexColor = Annotated[
    str, Field(pattern=r"^#?(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})$")
]


EffectTarget = Literal["shape", "text"]


TARGET_FIELD_DESCRIPTION = (
    "Apply to the shape outline or to its text glyphs. "
    "A text box with no fill "
    "and no line needs target='text' for a visible effect."
)


class GlowOptions(InputModel):
    """Set the glow radius in points; zero removes it."""

    radius: float = Field(ge=0)
    color: HexColor | None = None
    transparency: float | None = Field(default=None, ge=0, le=1)
    target: EffectTarget = Field(
        default="shape", description=TARGET_FIELD_DESCRIPTION
    )


class ReflectionOptions(InputModel):
    """Set a reflection preset or individual reflection properties."""

    reflection_type: int | None = Field(default=None, ge=0, le=9)
    blur: float | None = Field(default=None, ge=0)
    offset: float | None = Field(default=None, ge=0)
    size: float | None = Field(default=None, ge=0, le=100)
    transparency: float | None = Field(default=None, ge=0, le=1)
    target: EffectTarget = Field(
        default="shape", description=TARGET_FIELD_DESCRIPTION
    )


class SoftEdgeOptions(InputModel):
    """Set the soft edge radius in points; zero removes it."""

    radius: float = Field(ge=0)


def effect_of(
    shape: NativeObject, target: EffectTarget, name: str
) -> NativeObject:
    """Resolve the requested shape or text effect before native writes."""
    if target == "shape":
        return getattr(shape, name)
    if not shape.HasTextFrame:
        raise ValueError(
            f"Shape '{shape.Name}' has no text frame for {name.lower()}."
        )
    return getattr(shape.TextFrame2.TextRange.Font, name)


def will_not_draw(shape: NativeObject) -> bool:
    """Detect text frames with no visible fill or outline."""
    return bool(
        shape.HasTextFrame
        and not shape.Fill.Visible
        and not shape.Line.Visible
    )


def nothing_drawn_warning(shape: NativeObject, name: str) -> str:
    """Explain a successfully applied effect that has no visible outline."""
    return (
        f"The {name} on '{shape.Name}' has no fill or line to surround. "
        "Use target='text' to apply the effect to its glyphs."
    )


def apply_effect(
    shape: NativeObject,
    params: GlowOptions | ReflectionOptions | SoftEdgeOptions,
) -> JsonObject:
    """Apply validated effects to an already resolved shape on the STA."""
    result: JsonObject = {"status": "success", "shape_name": shape.Name}
    if isinstance(params, SoftEdgeOptions):
        shape.SoftEdge.Radius = params.radius
        result["soft_edge_radius"] = params.radius
        return result
    result["target"] = params.target
    if isinstance(params, GlowOptions):
        glow = effect_of(shape, params.target, "Glow")
        glow.Radius = params.radius
        if params.color is not None:
            glow.Color.RGB = hex_to_int(params.color)
        if params.transparency is not None:
            glow.Transparency = params.transparency
        result["glow_radius"] = params.radius
        effect_name = "glow"
        enabled = params.radius > 0
    else:
        reflection = effect_of(shape, params.target, "Reflection")
        for attribute, value in (
            ("Type", params.reflection_type),
            ("Blur", params.blur),
            ("Offset", params.offset),
            ("Size", params.size),
            ("Transparency", params.transparency),
        ):
            if value is not None:
                setattr(reflection, attribute, value)
        effect_name = "reflection"
        enabled = bool(reflection.Type)
    if params.target == "shape" and enabled and will_not_draw(shape):
        result["warnings"] = [nothing_drawn_warning(shape, effect_name)]
    return result


GradientStyle = Literal[
    "horizontal",
    "vertical",
    "diagonal_up",
    "diagonal_down",
    "from_corner",
    "from_title",
    "from_center",
]


class FillOptions(InputModel):
    """Set solid, gradient or invisible shape fill."""

    fill_type: Literal["solid", "gradient", "none"]
    color: HexColor | None = None
    gradient_color1: HexColor | None = None
    gradient_color2: HexColor | None = None
    gradient_style: GradientStyle | None = None
    transparency: float | None = Field(default=None, ge=0, le=1)


class LineOptions(InputModel):
    """Set line properties; unspecified properties remain unchanged."""

    color: HexColor | None = None
    weight: float | None = Field(default=None, ge=0)
    dash_style: str | None = Field(
        default=None, description=DASH_STYLE_DESCRIPTION
    )
    visible: bool | None = None
    transparency: float | None = Field(default=None, ge=0, le=1)

    @field_validator("dash_style")
    @classmethod
    def validate_dash_style(cls, value: str | None) -> str | None:
        """Resolve a supported native dash style before formatting."""
        return check_dash_style(value)


class ShadowOptions(InputModel):
    """Set the shadow on a shape outline or text glyphs."""

    visible: bool
    blur: float | None = Field(default=None, ge=0)
    offset_x: float | None = None
    offset_y: float | None = None
    color: HexColor | None = None
    transparency: float | None = Field(default=None, ge=0, le=1)
    target: EffectTarget = Field(
        default="shape", description=TARGET_FIELD_DESCRIPTION
    )


def apply_formatting(
    shape: NativeObject,
    params: FillOptions | LineOptions | ShadowOptions,
) -> JsonObject:
    """Apply validated formatting to an already resolved shape on the STA."""
    result: JsonObject = {"status": "success", "shape_name": shape.Name}
    if isinstance(params, FillOptions):
        fill = shape.Fill
        if params.fill_type == "none":
            fill.Visible = MSO_FALSE
        else:
            fill.Visible = MSO_TRUE
            if params.fill_type == "solid":
                fill.Solid()
                if params.color is not None:
                    fill.ForeColor.RGB = hex_to_int(params.color)
            else:
                fill.TwoColorGradient(
                    GRADIENT_STYLE_MAP[params.gradient_style or "horizontal"],
                    1,
                )
                if params.gradient_color1 is not None:
                    fill.ForeColor.RGB = hex_to_int(params.gradient_color1)
                if params.gradient_color2 is not None:
                    fill.BackColor.RGB = hex_to_int(params.gradient_color2)
            if params.transparency is not None:
                fill.Transparency = params.transparency
        result["fill_type"] = params.fill_type
    elif isinstance(params, LineOptions):
        line = shape.Line
        if params.visible is not None:
            line.Visible = MSO_TRUE if params.visible else MSO_FALSE
        if params.color is not None:
            line.ForeColor.RGB = hex_to_int(params.color)
        if params.weight is not None:
            line.Weight = params.weight
        if params.dash_style is not None:
            line.DashStyle = dash_style_value(params.dash_style)
        if params.transparency is not None:
            line.Transparency = params.transparency
    else:
        shadow = effect_of(shape, params.target, "Shadow")
        shadow.Visible = MSO_TRUE if params.visible else MSO_FALSE
        if params.visible:
            for attribute, value in (
                ("Blur", params.blur),
                ("OffsetX", params.offset_x),
                ("OffsetY", params.offset_y),
                ("Transparency", params.transparency),
            ):
                if value is not None:
                    setattr(shadow, attribute, value)
            if params.color is not None:
                shadow.ForeColor.RGB = hex_to_int(params.color)
        result.update(target=params.target, shadow_visible=params.visible)
        if (
            params.target == "shape"
            and params.visible
            and will_not_draw(shape)
        ):
            result["warnings"] = [nothing_drawn_warning(shape, "shadow")]
    return result


def resolve_color(pres: NativeObject, color_str: str) -> str:
    """Resolve a color string to a hex '#RRGGBB' value."""
    if color_str.startswith("#"):
        return int_to_hex(hex_to_int(color_str))
    index = get_theme_color_index(color_str)
    return int_to_hex(pres.SlideMaster.Theme.ThemeColorScheme(index).RGB)
