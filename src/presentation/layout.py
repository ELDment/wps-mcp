"""Shape layout and presentation geometry with preflighted inputs."""

from pathlib import Path
from typing import Annotated, Literal

from pydantic import Field, model_validator

from src.common.colors import hex_to_int
from src.models import DocumentInput, JsonObject
from src.native.runtime import get_session
from src.presentation.constants import (
    ALIGN_CMD_MAP,
    DISTRIBUTE_CMD_MAP,
    FLIP_CMD_MAP,
    GRADIENT_STYLE_MAP,
    MERGE_CMD_MAP,
    MSO_FALSE,
    MSO_TRUE,
    SLIDE_SIZE_MAP,
)
from src.presentation.navigation import goto_slide
from src.presentation.shape_formatting import GradientStyle, HexColor
from src.presentation.shape_lookup import require_top_level, resolve_shape
from src.presentation.shape_models import GroupShapesInput, ShapeInput
from src.protocol.definitions import tool

PRESET_DIMENSIONS: dict[str, tuple[float, float]] = {
    "16:9": (960, 540),
    "widescreen": (960, 540),
    "4:3": (720, 540),
    "16:10": (960, 600),
    "a4": (842, 595),
    "a3": (1191, 842),
    "letter": (792, 612),
    "35mm": (792, 528),
    "overhead": (720, 540),
    "banner": (576, 72),
}


class AlignShapesInput(GroupShapesInput):
    """Align distinct top-level shapes to an edge or center."""

    align_to: Literal["left", "center", "right", "top", "middle", "bottom"]
    relative_to_slide: bool = False


class DistributeShapesInput(GroupShapesInput):
    """Distribute at least three distinct top-level shapes."""

    shape_names: list[str] = Field(min_length=3)
    direction: Literal["horizontal", "vertical"]
    relative_to_slide: bool = False


class GetSlideSizeInput(DocumentInput):
    """Select a presentation whose dimensions should be read."""


class SetSlideSizeInput(DocumentInput):
    """Apply a preset, then dimension overrides, then orientation."""

    width: float | None = Field(default=None, gt=0)
    height: float | None = Field(default=None, gt=0)
    preset: (
        Literal[
            "16:9",
            "widescreen",
            "4:3",
            "16:10",
            "a4",
            "a3",
            "letter",
            "35mm",
            "overhead",
            "banner",
        ]
        | None
    ) = None
    orientation: Literal["landscape", "portrait"] | None = None


class SetSlideBackgroundInput(DocumentInput):
    """Apply one background to a slide or a distinct slide collection."""

    slide_index: int | None = Field(default=None, ge=1)
    slide_indices: list[Annotated[int, Field(ge=1)]] | None = Field(
        default=None, min_length=1
    )
    fill_type: Literal["solid", "gradient", "picture", "none", "master"]
    color: HexColor | None = None
    gradient_color1: HexColor | None = None
    gradient_color2: HexColor | None = None
    gradient_style: GradientStyle | None = None
    image_path: str | None = Field(default=None, min_length=1)
    transparency: float | None = Field(default=None, ge=0, le=1)

    @model_validator(mode="after")
    def validate_background(self) -> "SetSlideBackgroundInput":
        """Require one slide selector and the requested fill's parameters."""
        if (self.slide_index is None) == (self.slide_indices is None):
            raise ValueError(
                "Provide exactly one of slide_index or slide_indices."
            )
        if self.slide_indices is not None and len(
            set(self.slide_indices)
        ) != len(self.slide_indices):
            raise ValueError("slide_indices must contain distinct indices.")
        if self.fill_type == "solid" and self.color is None:
            raise ValueError("Solid backgrounds require color.")
        if self.fill_type == "gradient" and (
            self.gradient_color1 is None or self.gradient_color2 is None
        ):
            raise ValueError(
                "Gradient backgrounds require both gradient colors."
            )
        if self.fill_type == "picture" and self.image_path is None:
            raise ValueError("Picture backgrounds require image_path.")
        return self


class FlipShapeInput(ShapeInput):
    """Mirror one shape horizontally or vertically."""

    direction: Literal["horizontal", "vertical"]


class MergeShapesInput(GroupShapesInput):
    """Merge shapes using a Boolean operation and optional primary shape."""

    merge_type: Literal[
        "union", "combine", "intersect", "subtract", "fragment"
    ]
    primary_shape: str | None = Field(default=None, min_length=1)

    @model_validator(mode="after")
    def validate_primary(self) -> "MergeShapesInput":
        """Require the primary shape to belong to the merge selection."""
        if (
            self.primary_shape is not None
            and self.primary_shape not in self.shape_names
        ):
            raise ValueError("primary_shape must be one of shape_names.")
        return self


@tool(
    "wps_ppt_layout",
    component="wpp",
    annotations={"idempotentHint": True},
    operation="align",
)
def align_shapes(params: AlignShapesInput) -> JsonObject:
    """Align shapes to each other or relative to slide edges."""
    session = get_session("wpp")
    slide = session.get_document().Slides(params.slide_index)
    require_top_level(
        slide, params.shape_names, params.slide_index, "wps_ppt_layout/align"
    )
    goto_slide(session.get_application(), params.slide_index)
    slide.Shapes.Range(tuple(params.shape_names)).Align(
        ALIGN_CMD_MAP[params.align_to],
        MSO_TRUE if params.relative_to_slide else MSO_FALSE,
    )
    return {
        "status": "success",
        "aligned_count": len(params.shape_names),
        "align_to": params.align_to,
        "relative_to_slide": params.relative_to_slide,
    }


@tool(
    "wps_ppt_layout",
    component="wpp",
    annotations={"idempotentHint": True},
    operation="distribute",
)
def distribute_shapes(params: DistributeShapesInput) -> JsonObject:
    """Distribute shapes evenly between the outermost shapes or slide edges."""
    session = get_session("wpp")
    slide = session.get_document().Slides(params.slide_index)
    require_top_level(
        slide,
        params.shape_names,
        params.slide_index,
        "wps_ppt_layout/distribute",
    )
    goto_slide(session.get_application(), params.slide_index)
    slide.Shapes.Range(tuple(params.shape_names)).Distribute(
        DISTRIBUTE_CMD_MAP[params.direction],
        MSO_TRUE if params.relative_to_slide else MSO_FALSE,
    )
    return {
        "status": "success",
        "distributed_count": len(params.shape_names),
        "direction": params.direction,
        "relative_to_slide": params.relative_to_slide,
    }


@tool(
    "wps_ppt_layout",
    component="wpp",
    annotations={"readOnlyHint": True},
    operation="get_size",
)
def get_slide_size(params: GetSlideSizeInput) -> JsonObject:
    """Read slide dimensions in points and inches, preset and orientation."""
    setup = get_session("wpp").get_document().PageSetup
    width, height = setup.SlideWidth, setup.SlideHeight
    size = setup.SlideSize
    return {
        "status": "success",
        "width_points": round(width, 2),
        "height_points": round(height, 2),
        "width_inches": round(width / 72, 4),
        "height_inches": round(height / 72, 4),
        "slide_size_type": size,
        "slide_size_name": next(
            (name for name, value in SLIDE_SIZE_MAP.items() if value == size),
            None,
        ),
        "orientation": "landscape"
        if setup.SlideOrientation == 1
        else "portrait",
    }


@tool(
    "wps_ppt_layout",
    component="wpp",
    annotations={"idempotentHint": True},
    operation="set_size",
)
def set_slide_size(params: SetSlideSizeInput) -> JsonObject:
    """Apply a size preset, explicit point dimensions, then orientation."""
    setup = get_session("wpp").get_document().PageSetup
    width, height = (
        PRESET_DIMENSIONS[params.preset]
        if params.preset is not None
        else (None, None)
    )
    if params.width is not None:
        width = params.width
    if params.height is not None:
        height = params.height
    if width is not None:
        setup.SlideWidth = width
    if height is not None:
        setup.SlideHeight = height
    if params.orientation is not None:
        setup.SlideOrientation = 1 if params.orientation == "landscape" else 2
    return {
        "status": "success",
        "width_points": round(setup.SlideWidth, 2),
        "height_points": round(setup.SlideHeight, 2),
        "width_inches": round(setup.SlideWidth / 72, 4),
        "height_inches": round(setup.SlideHeight / 72, 4),
    }


@tool(
    "wps_ppt_layout",
    component="wpp",
    annotations={"idempotentHint": True},
    operation="background",
)
def set_slide_background(params: SetSlideBackgroundInput) -> JsonObject:
    """Set slide backgrounds after validating every slide and any image file.

    Runtime failures are recorded per slide; already completed backgrounds
    are retained.
    """
    image_path = None
    if params.fill_type == "picture":
        assert params.image_path is not None
        image_path = Path(params.image_path).resolve()
        if not image_path.is_file():
            raise ValueError(f"Image file not found: {image_path}")
    color = hex_to_int(params.color) if params.color is not None else None
    color1 = (
        hex_to_int(params.gradient_color1) if params.gradient_color1 else None
    )
    color2 = (
        hex_to_int(params.gradient_color2) if params.gradient_color2 else None
    )
    session = get_session("wpp")
    presentation = session.get_document()
    if params.slide_indices is not None:
        indices = params.slide_indices
    else:
        assert params.slide_index is not None
        indices = [params.slide_index]
    slides = [presentation.Slides(index) for index in indices]
    outcomes: list[JsonObject] = []
    applied: list[int] = []
    for index, slide in zip(indices, slides, strict=True):
        try:
            goto_slide(session.get_application(), index)
            if params.fill_type == "master":
                slide.FollowMasterBackground = MSO_TRUE
            else:
                slide.FollowMasterBackground = MSO_FALSE
                fill = slide.Background.Fill
                if params.fill_type == "solid":
                    fill.Solid()
                    fill.ForeColor.RGB = color
                elif params.fill_type == "gradient":
                    fill.TwoColorGradient(
                        GRADIENT_STYLE_MAP[
                            params.gradient_style or "horizontal"
                        ],
                        1,
                    )
                    fill.ForeColor.RGB = color1
                    fill.BackColor.RGB = color2
                elif params.fill_type == "picture":
                    fill.UserPicture(str(image_path))
                else:
                    fill.Background()
                if (
                    params.transparency is not None
                    and params.fill_type != "none"
                ):
                    fill.Transparency = params.transparency
            applied.append(index)
            outcomes.append({"slide_index": index, "status": "success"})
        except Exception as error:
            outcomes.append(
                {
                    "slide_index": index,
                    "status": "failed",
                    "error": str(error),
                    "may_have_changes": True,
                }
            )
    result: JsonObject = {
        "status": "success"
        if len(applied) == len(indices)
        else ("partial_success" if applied else "failed"),
        "slide_indices": applied,
        "fill_type": params.fill_type,
        "results": outcomes,
    }
    if params.slide_indices is None:
        result["slide_index"] = params.slide_index
    return result


@tool("wps_ppt_layout", component="wpp", operation="flip")
def flip_shape(params: FlipShapeInput) -> JsonObject:
    """Mirror a shape and report its horizontal and vertical flip state."""
    session = get_session("wpp")
    slide = session.get_document().Slides(params.slide_index)
    shape = resolve_shape(slide, params.shape_name_or_index)
    goto_slide(session.get_application(), params.slide_index)
    shape.Flip(FLIP_CMD_MAP[params.direction])
    return {
        "status": "success",
        "shape_name": shape.Name,
        "horizontal_flip": bool(shape.HorizontalFlip),
        "vertical_flip": bool(shape.VerticalFlip),
    }


@tool(
    "wps_ppt_layout",
    component="wpp",
    annotations={"destructiveHint": True},
    operation="merge",
)
def merge_shapes(params: MergeShapesInput) -> JsonObject:
    """Merge shapes using union, combine, intersect, subtract or fragment."""
    session = get_session("wpp")
    slide = session.get_document().Slides(params.slide_index)
    require_top_level(
        slide, params.shape_names, params.slide_index, "wps_ppt_layout/merge"
    )
    primary = resolve_shape(
        slide, params.primary_shape or params.shape_names[0]
    )
    goto_slide(session.get_application(), params.slide_index)
    slide.Shapes.Range(tuple(params.shape_names)).MergeShapes(
        MERGE_CMD_MAP[params.merge_type], primary
    )
    return {
        "status": "success",
        "merge_type": params.merge_type,
        "merged_count": len(params.shape_names),
    }
