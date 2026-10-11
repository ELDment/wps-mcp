"""Apply validated shape styles through one preflighted pipeline."""

from __future__ import annotations

import logging
from collections.abc import Iterator
from contextlib import contextmanager
from typing import Annotated, Generic, Literal, TypeAlias, TypeVar

from pydantic import Field, field_validator, model_validator

from src.common.colors import hex_to_int
from src.models import DocumentInput, InputModel, JsonObject, NativeObject
from src.native.runtime import get_session
from src.presentation.constants import (
    MSO_FALSE,
    MSO_SHAPE_RECTANGLE,
    MSO_TRUE,
)
from src.presentation.navigation import goto_slide
from src.presentation.shape_formatting import (
    FillOptions,
    GlowOptions,
    LineOptions,
    ReflectionOptions,
    ShadowOptions,
    SoftEdgeOptions,
    apply_effect,
    apply_formatting,
    effect_of,
    resolve_color,
)
from src.presentation.shape_lookup import resolve_shape
from src.presentation.shape_models import ShapeReference
from src.presentation.text_formatting import (
    apply_font_format,
    apply_text_format,
)
from src.presentation.text_models import TextFormatSpec
from src.protocol.definitions import tool

logger = logging.getLogger(__name__)


@contextmanager
def _temporary_shape(slide: NativeObject) -> Iterator[NativeObject]:
    """Delete the style scratch shape without masking an operation failure."""
    shape = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, -10000, -10000, 1, 1)
    try:
        yield shape
    except BaseException:
        try:
            shape.Delete()
        except Exception:
            logger.exception(
                "Failed to delete the temporary default-style shape"
            )
        raise
    else:
        shape.Delete()


class SetDefaultShapeStyleInput(DocumentInput):
    """Input for setting the default shape style for new shapes."""

    slide_index: int | None = Field(
        default=None,
        ge=1,
        description=(
            "1-based slide index of the template shape. When provided "
            "together with shape_name_or_index, that shape's full style "
            "(fill, border, effects, font, etc.) is captured as the "
            "default. Omit to use property-based mode instead."
        ),
    )
    shape_name_or_index: str | int | None = Field(
        default=None,
        description=(
            "Shape name, group path, or 1-based index; prefer a stable name."
        ),
    )
    fill_color: str | None = Field(
        default=None,
        description=(
            "Fill color #RRGGBB; implies solid unless fill_type is given."
        ),
    )
    fill_type: str | None = Field(
        default=None,
        description=(
            "'solid' to apply fill_color, 'none' for no fill. Omit to "
            "leave fill unchanged."
        ),
    )
    line_visible: bool | None = Field(
        default=None,
        description="Show (true) or hide (false) the shape border.",
    )
    line_color: str | None = Field(
        default=None,
        description=(
            "Border color #RRGGBB; enables the line unless explicitly "
            "disabled."
        ),
    )
    line_weight: float | None = Field(
        default=None, ge=0, description="Border weight in points."
    )
    font_name: str | None = Field(
        default=None,
        description="Font name applied to Latin and East Asian text.",
    )
    font_size: float | None = Field(
        default=None, gt=0, description="Default font size in points."
    )
    font_bold: bool | None = Field(
        default=None, description="Default bold setting."
    )
    font_italic: bool | None = Field(
        default=None, description="Default italic setting."
    )
    font_color: str | None = Field(
        default=None,
        description="Default font color as '#RRGGBB' hex or theme name.",
    )

    @model_validator(mode="after")
    def validate_mode(self) -> SetDefaultShapeStyleInput:
        """Keep template selection separate from explicit style options."""
        shape_mode = (
            self.slide_index is not None
            or self.shape_name_or_index is not None
        )
        prop_mode = any(
            v is not None
            for v in [
                self.fill_color,
                self.fill_type,
                self.line_visible,
                self.line_color,
                self.line_weight,
                self.font_name,
                self.font_size,
                self.font_bold,
                self.font_italic,
                self.font_color,
            ]
        )
        if shape_mode and prop_mode:
            raise ValueError(
                "Cannot mix shape-based and property-based modes. Either "
                "provide slide_index + shape_name_or_index, or fill/line/font"
                " properties."
            )
        if shape_mode:
            if self.slide_index is None or self.shape_name_or_index is None:
                raise ValueError(
                    "slide_index and shape_name_or_index must both be "
                    "provided."
                )
        else:
            if self.fill_type is not None and self.fill_type not in (
                "solid",
                "none",
            ):
                raise ValueError(
                    f"fill_type must be 'solid' or 'none', got "
                    f"'{self.fill_type}'"
                )
            if self.fill_type == "solid" and self.fill_color is None:
                raise ValueError(
                    "fill_color is required when fill_type='solid'"
                )
        return self


@tool(
    "wps_ppt_style",
    component="wpp",
    annotations={"title": "Set Default Shape Style", "idempotentHint": True},
    operation="set_default",
)
def set_default_shape_style(params: SetDefaultShapeStyleInput) -> JsonObject:
    """Set future shape defaults from a template or explicit properties."""
    session = get_session("wpp")
    presentation = session.get_document()
    if params.slide_index is not None:
        slide = presentation.Slides(params.slide_index)
        if params.shape_name_or_index is None:
            raise ValueError("A template shape is required")
        shape = resolve_shape(slide, params.shape_name_or_index)
        goto_slide(session.get_application(), params.slide_index)
        shape.SetShapesDefaultProperties()
        return {"success": True, "source_shape": shape.Name}
    if presentation.Slides.Count == 0:
        raise ValueError("Presentation has no slides. Add a slide first.")
    colors = {
        field: resolve_color(presentation, value)
        for field in ("fill_color", "line_color", "font_color")
        if (value := getattr(params, field)) is not None
    }
    with _temporary_shape(presentation.Slides(1)) as shape:
        if params.fill_type == "none":
            shape.Fill.Visible = MSO_FALSE
        elif "fill_color" in colors:
            shape.Fill.Visible = MSO_TRUE
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = hex_to_int(colors["fill_color"])
        if params.line_visible is not None:
            shape.Line.Visible = MSO_TRUE if params.line_visible else MSO_FALSE
        if "line_color" in colors:
            shape.Line.ForeColor.RGB = hex_to_int(colors["line_color"])
        if params.line_weight is not None:
            shape.Line.Weight = params.line_weight
        apply_font_format(
            shape.TextFrame.TextRange.Font,
            TextFormatSpec(
                font_name=params.font_name,
                font_size=params.font_size,
                bold=params.font_bold,
                italic=params.font_italic,
                color=colors.get("font_color"),
            ),
        )
        shape.SetShapesDefaultProperties()
    return {"success": True}


class CopyFormattingInput(DocumentInput):
    """Input for copying formatting from one shape to others."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    source_shape: str | int = Field(
        ..., description="Source shape name or one-based index."
    )
    target_shapes: list[str | int] = Field(
        ...,
        min_length=1,
        description=(
            "List of target shape names (string) or 1-based indices (int)"
        ),
    )


@tool(
    "wps_ppt_style",
    component="wpp",
    annotations={"title": "Copy Formatting", "idempotentHint": True},
    operation="copy",
)
def copy_formatting(params: CopyFormattingInput) -> JsonObject:
    """Copy formatting from a source shape to one or more target shapes.

    Uses PickUp/Apply to transfer fill, line, shadow, and other visual
    formatting from the source shape to each target shape on the same
    slide. Identify shapes by name (string) or 1-based index (int).
    """
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    src = resolve_shape(slide, params.source_shape)
    targets = [
        resolve_shape(slide, target_id) for target_id in params.target_shapes
    ]
    src.PickUp()
    applied_to = []
    for target in targets:
        target.Apply()
        applied_to.append(target.Name)
    return {"success": True, "source": src.Name, "applied_to": applied_to}


StyleKind = TypeVar("StyleKind", bound=str)
StyleOptions = TypeVar("StyleOptions", bound=InputModel)


class StyleOperation(InputModel, Generic[StyleKind, StyleOptions]):
    """One typed formatting change, without embedded target selectors."""

    kind: StyleKind
    options: StyleOptions


StyleChange: TypeAlias = Annotated[
    StyleOperation[Literal["fill"], FillOptions]
    | StyleOperation[Literal["line"], LineOptions]
    | StyleOperation[Literal["shadow"], ShadowOptions]
    | StyleOperation[Literal["glow"], GlowOptions]
    | StyleOperation[Literal["reflection"], ReflectionOptions]
    | StyleOperation[Literal["soft_edge"], SoftEdgeOptions]
    | StyleOperation[Literal["text"], TextFormatSpec],
    Field(discriminator="kind"),
]


class ApplyStyleInput(DocumentInput):
    """Apply ordered formatting changes to one or more distinct shapes."""

    slide_index: int = Field(ge=1)
    shapes: list[ShapeReference] = Field(min_length=1)
    operations: list[StyleChange] = Field(min_length=1)

    @field_validator("shapes")
    @classmethod
    def validate_shapes(
        cls, values: list[ShapeReference]
    ) -> list[ShapeReference]:
        """Require distinct shape selectors before resolving targets."""
        if len(set(values)) != len(values):
            raise ValueError("shapes must contain distinct selectors")
        return values


@tool("wps_ppt_style", operation="apply", component="wpp")
def apply_style(params: ApplyStyleInput) -> JsonObject:
    """Preflight all options and shapes, then apply each change in order.

    Native failures are reported per shape and operation. Successful
    changes remain applied; a failed native operation may have partially
    changed its target and is never automatically replayed.
    """
    session = get_session("wpp")
    slide = session.get_document(params.document).Slides(params.slide_index)
    shapes = [resolve_shape(slide, reference) for reference in params.shapes]
    for shape in shapes:
        for operation in params.operations:
            model = operation.options
            if isinstance(model, FillOptions):
                _ = shape.Fill
            elif isinstance(model, LineOptions):
                _ = shape.Line
            elif isinstance(
                model, (ShadowOptions, GlowOptions, ReflectionOptions)
            ):
                effect_name = {
                    "shadow": "Shadow",
                    "glow": "Glow",
                    "reflection": "Reflection",
                }[operation.kind]
                effect_of(shape, model.target, effect_name)
            elif isinstance(model, SoftEdgeOptions):
                _ = shape.SoftEdge
            else:
                if not shape.HasTextFrame:
                    raise ValueError(
                        f"Shape '{shape.Name}' has no text frame."
                    )
                _ = shape.TextFrame.TextRange.Font
                if model.highlight_color is not None:
                    _ = shape.TextFrame2.TextRange.Font.Highlight
    goto_slide(session.get_application(), params.slide_index)
    results: list[JsonObject] = []
    failed = 0
    for reference, shape in zip(params.shapes, shapes, strict=True):
        outcomes: list[JsonObject] = []
        for operation in params.operations:
            model = operation.options
            try:
                if isinstance(
                    model, (FillOptions, LineOptions, ShadowOptions)
                ):
                    outcome = apply_formatting(shape, model)
                elif isinstance(
                    model, (GlowOptions, ReflectionOptions, SoftEdgeOptions)
                ):
                    outcome = apply_effect(shape, model)
                else:
                    outcome = apply_text_format(shape, model)
                outcomes.append({**outcome, "kind": operation.kind})
            except Exception as error:
                failed += 1
                outcomes.append(
                    {
                        "kind": operation.kind,
                        "status": "failed",
                        "error": str(error),
                        "may_have_changes": True,
                    }
                )
        results.append({"shape": reference, "operations": outcomes})
    attempted = len(shapes) * len(params.operations)
    return {
        "status": "success"
        if not failed
        else ("failed" if failed == attempted else "partial_success"),
        "completed_operations": attempted - failed,
        "failed_operations": failed,
        "results": results,
    }
