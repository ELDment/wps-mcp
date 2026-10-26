"""Create and edit freeform paths with validated node specifications."""

from typing import Literal

from pydantic import Field, model_validator

from wps_mcp.models import DocumentInput, InputModel, JsonObject, NativeObject
from wps_mcp.native.objects import capabilities_for, is_optional_member_error
from wps_mcp.native.runtime import get_session
from wps_mcp.presentation.constants import (
    EDITING_TYPE_MAP,
    EDITING_TYPE_NAMES,
    MSO_EDITING_AUTO,
    MSO_FREEFORM,
    MSO_SEGMENT_LINE,
    SEGMENT_TYPE_MAP,
    SEGMENT_TYPE_NAMES,
)
from wps_mcp.presentation.navigation import goto_slide
from wps_mcp.presentation.shapes.lookup import resolve_shape
from wps_mcp.protocol.definitions import tool

EditingType = Literal["auto", "corner", "smooth", "symmetric"]
SegmentType = Literal["line", "curve"]


class NodeSpec(InputModel):
    """A line or Bézier endpoint with optional control points."""

    segment_type: SegmentType
    editing_type: EditingType = "auto"
    x1: float
    y1: float
    x2: float | None = None
    y2: float | None = None
    x3: float | None = None
    y3: float | None = None

    @model_validator(mode="after")
    def validate_node(self) -> "NodeSpec":
        """Normalize lines and require corner curve control points."""
        if self.segment_type == "line":
            self.editing_type = "auto"
        elif self.editing_type not in ("auto", "corner"):
            raise ValueError("New curves require auto or corner editing_type.")
        elif self.editing_type == "corner" and any(
            value is None for value in (self.x2, self.y2, self.x3, self.y3)
        ):
            raise ValueError("Corner curves require x2, y2, x3 and y3.")
        return self

    def native_arguments(self) -> tuple[int | float, ...]:
        """Build shared COM arguments for creation and insertion."""
        arguments: tuple[int | float, ...] = (
            SEGMENT_TYPE_MAP[self.segment_type],
            EDITING_TYPE_MAP[self.editing_type],
            self.x1,
            self.y1,
        )
        if self.segment_type == "curve" and self.editing_type == "corner":
            assert self.x2 is not None and self.y2 is not None
            assert self.x3 is not None and self.y3 is not None
            arguments += (self.x2, self.y2, self.x3, self.y3)
        return arguments


class BuildFreeformInput(DocumentInput):
    """Build a path from a starting point and measured segments."""

    slide_index: int = Field(ge=1)
    start_x: float
    start_y: float
    start_editing_type: Literal["auto", "corner"] = "corner"
    nodes: list[NodeSpec] = Field(min_length=1)
    close_path: bool = False
    shape_name: str | None = Field(default=None, min_length=1)


class GetShapeNodesInput(DocumentInput):
    """Select an existing freeform by name or one-based shape index."""

    slide_index: int = Field(ge=1)
    shape_name: str | None = Field(default=None, min_length=1)
    shape_index: int | None = Field(default=None, ge=1)

    @model_validator(mode="after")
    def validate_shape(self) -> "GetShapeNodesInput":
        """Require exactly one shape name or index."""
        if (self.shape_name is None) == (self.shape_index is None):
            raise ValueError(
                "Provide exactly one of shape_name and shape_index."
            )
        return self


class ShapeNodeInput(GetShapeNodesInput):
    """Select a node by its current one-based index."""

    node_index: int = Field(ge=1)


class SetNodePositionInput(ShapeNodeInput):
    """Move a node to slide coordinates in points."""

    x: float
    y: float


class InsertNodeInput(NodeSpec, GetShapeNodesInput):
    """Insert one segment after an existing node."""

    after_index: int = Field(ge=1)
    segment_type: SegmentType = "line"


class DeleteNodeInput(ShapeNodeInput):
    """Delete a node and its following segment."""


class SetNodeEditingTypeInput(ShapeNodeInput):
    """Choose node editing, including smooth and symmetric modes."""

    editing_type: EditingType


class SetSegmentTypeInput(ShapeNodeInput):
    """Change the segment after a node to line or curve."""

    segment_type: SegmentType


def resolve_freeform(
    params: GetShapeNodesInput,
    node_index: int | None = None,
) -> NativeObject:
    """Resolve a freeform and validate its node before editing."""
    slide = get_session("wpp").get_document().Slides(params.slide_index)
    reference = (
        params.shape_name
        if params.shape_name is not None
        else params.shape_index
    )
    assert reference is not None
    shape = resolve_shape(slide, reference)
    if shape.Type != MSO_FREEFORM:
        raise ValueError(f"Shape '{shape.Name}' is not a freeform.")
    capabilities_for(shape).require("freeform_nodes")
    if node_index is not None and node_index > shape.Nodes.Count:
        raise ValueError(
            f"Node {node_index} exceeds {shape.Nodes.Count} nodes."
        )
    return shape


@tool("wps_ppt_freeform", component="wpp", operation="build")
def build_freeform(params: BuildFreeformInput) -> JsonObject:
    """Create a freeform path using lines or Bézier segments.

    For auto curves x1/y1 are the endpoint. Corner curves use x1/y1 and
    x2/y2 as control points and x3/y3 as the endpoint. close_path adds a
    straight segment back to the starting point.
    """
    session = get_session("wpp")
    slide = session.get_document().Slides(params.slide_index)
    if not capabilities_for(slide).allows("freeform_builder"):
        if any(node.segment_type != "line" for node in params.nodes):
            capabilities_for(slide).require("freeform_builder")
        goto_slide(params.slide_index)
        points = [[params.start_x, params.start_y]]
        points.extend([node.x1, node.y1] for node in params.nodes)
        if params.close_path and points[-1] != points[0]:
            points.append(points[0])
        shape = slide.Shapes.AddPolyline(points)
    else:
        goto_slide(params.slide_index)
        builder = slide.Shapes.BuildFreeform(
            EDITING_TYPE_MAP[params.start_editing_type],
            params.start_x,
            params.start_y,
        )
        for node in params.nodes:
            builder.AddNodes(*node.native_arguments())
        if params.close_path:
            builder.AddNodes(
                MSO_SEGMENT_LINE,
                MSO_EDITING_AUTO,
                params.start_x,
                params.start_y,
            )
        shape = builder.ConvertToShape()
    if params.shape_name is not None:
        shape.Name = params.shape_name
    return {
        "success": True,
        "shape_name": shape.Name,
        "shape_index": shape.ZOrderPosition,
        "left": round(shape.Left, 2),
        "top": round(shape.Top, 2),
        "width": round(shape.Width, 2),
        "height": round(shape.Height, 2),
    }


@tool(
    "wps_ppt_freeform",
    component="wpp",
    annotations={"readOnlyHint": True},
    operation="nodes",
)
def get_shape_nodes(params: GetShapeNodesInput) -> JsonObject:
    """Read freeform positions, control points and closing nodes."""
    shape = resolve_freeform(params)
    nodes = shape.Nodes
    vertices = shape.Vertices
    result: list[JsonObject] = []
    for index in range(1, nodes.Count + 1):
        item: JsonObject = {
            "index": index,
            "x": round(float(vertices[index - 1][0]), 2),
            "y": round(float(vertices[index - 1][1]), 2),
        }
        try:
            node = nodes.Item(index)
            editing_type, segment_type = node.EditingType, node.SegmentType
            item.update(
                editing_type=EDITING_TYPE_NAMES.get(
                    editing_type, str(editing_type)
                ),
                segment_type=SEGMENT_TYPE_NAMES.get(
                    segment_type, str(segment_type)
                ),
            )
        except Exception as error:
            if not is_optional_member_error(error):
                raise
            item.update(
                editing_type="inaccessible",
                segment_type="inaccessible",
                warnings=[
                    f"Control or closing node metadata unavailable: {error}"
                ],
            )
        result.append(item)
    return {
        "shape_name": shape.Name,
        "node_count": len(result),
        "nodes": result,
    }


@tool("wps_ppt_freeform", component="wpp", operation="move_node")
def set_node_position(params: SetNodePositionInput) -> JsonObject:
    """Move a node and any linked smooth-mode control points."""
    shape = resolve_freeform(params, params.node_index)
    goto_slide(params.slide_index)
    shape.Nodes.SetPosition(params.node_index, params.x, params.y)
    vertex = shape.Vertices[params.node_index - 1]
    return {
        "success": True,
        "shape_name": shape.Name,
        "node_index": params.node_index,
        "x": round(float(vertex[0]), 2),
        "y": round(float(vertex[1]), 2),
    }


@tool("wps_ppt_freeform", component="wpp", operation="insert_node")
def insert_node(params: InsertNodeInput) -> JsonObject:
    """Insert a segment and shift subsequent node indices."""
    shape = resolve_freeform(params, params.after_index)
    goto_slide(params.slide_index)
    shape.Nodes.Insert(params.after_index, *params.native_arguments())
    return {
        "success": True,
        "shape_name": shape.Name,
        "new_node_count": shape.Nodes.Count,
    }


@tool(
    "wps_ppt_freeform",
    component="wpp",
    annotations={"destructiveHint": True},
    operation="delete_node",
)
def delete_node(params: DeleteNodeInput) -> JsonObject:
    """Delete a node's segment or a control point's curve."""
    shape = resolve_freeform(params, params.node_index)
    goto_slide(params.slide_index)
    shape.Nodes.Delete(params.node_index)
    return {
        "success": True,
        "shape_name": shape.Name,
        "remaining_node_count": shape.Nodes.Count,
    }


@tool("wps_ppt_freeform", component="wpp", operation="set_editing")
def set_node_editing_type(params: SetNodeEditingTypeInput) -> JsonObject:
    """Set node editing mode, allowing control points to reposition."""
    shape = resolve_freeform(params, params.node_index)
    goto_slide(params.slide_index)
    shape.Nodes.SetEditingType(
        params.node_index, EDITING_TYPE_MAP[params.editing_type]
    )
    result: JsonObject = {
        "success": True,
        "shape_name": shape.Name,
        "node_index": params.node_index,
    }
    try:
        editing_type = shape.Nodes.Item(params.node_index).EditingType
        result["editing_type"] = EDITING_TYPE_NAMES.get(
            editing_type, str(editing_type)
        )
    except Exception as error:
        if not is_optional_member_error(error):
            raise
        result.update(
            editing_type="inaccessible",
            warnings=[f"Applied edit; node metadata unavailable: {error}"],
        )
    return result


@tool("wps_ppt_freeform", component="wpp", operation="set_segment")
def set_segment_type(params: SetSegmentTypeInput) -> JsonObject:
    """Set the next segment type, allowing control points to change."""
    shape = resolve_freeform(params, params.node_index)
    goto_slide(params.slide_index)
    old_count = shape.Nodes.Count
    shape.Nodes.SetSegmentType(
        params.node_index, SEGMENT_TYPE_MAP[params.segment_type]
    )
    new_count = shape.Nodes.Count
    result: JsonObject = {
        "success": True,
        "shape_name": shape.Name,
        "node_index": params.node_index,
        "segment_type": params.segment_type,
        "old_node_count": old_count,
        "new_node_count": new_count,
    }
    if old_count != new_count:
        result["warnings"] = [
            "Control point count changed; call "
            "wps_ppt_freeform/nodes for current indices."
        ]
    return result
