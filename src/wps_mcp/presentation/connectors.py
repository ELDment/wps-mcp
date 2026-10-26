"""Connector creation and formatting with preflighted endpoints."""

from typing import Annotated, Literal

from pydantic import Field, model_validator

from wps_mcp.models import DocumentInput, JsonObject, NativeObject
from wps_mcp.native.objects import capabilities_for
from wps_mcp.native.runtime import get_session
from wps_mcp.presentation.constants import (
    MSO_ARROWHEAD_DIAMOND,
    MSO_ARROWHEAD_NONE,
    MSO_ARROWHEAD_OPEN,
    MSO_ARROWHEAD_OVAL,
    MSO_ARROWHEAD_STEALTH,
    MSO_ARROWHEAD_TRIANGLE,
)
from wps_mcp.presentation.navigation import goto_slide
from wps_mcp.presentation.shapes.formatting import (
    LineOptions,
    apply_formatting,
)
from wps_mcp.presentation.shapes.lookup import resolve_shape
from wps_mcp.presentation.shapes.models import ShapeInput
from wps_mcp.protocol.definitions import tool

CONNECTOR_TYPE_MAP = {"straight": 1, "elbow": 2, "curve": 3}
ARROW_STYLE_MAP = {
    "none": MSO_ARROWHEAD_NONE,
    "triangle": MSO_ARROWHEAD_TRIANGLE,
    "open": MSO_ARROWHEAD_OPEN,
    "stealth": MSO_ARROWHEAD_STEALTH,
    "diamond": MSO_ARROWHEAD_DIAMOND,
    "oval": MSO_ARROWHEAD_OVAL,
}
ARROW_LENGTH_MAP = {"short": 1, "medium": 2, "long": 3}
ARROW_WIDTH_MAP = {"narrow": 1, "medium": 2, "wide": 3}
SITE_DIRECTION_VECTORS = {
    "top": (0, -1),
    "bottom": (0, 1),
    "left": (-1, 0),
    "right": (1, 0),
}
ConnectionSite = (
    Annotated[int, Field(ge=1)] | Literal["top", "bottom", "left", "right"]
)
ArrowStyle = Literal["none", "triangle", "open", "stealth", "diamond", "oval"]
ArrowLength = Literal["short", "medium", "long"]
ArrowWidth = Literal["narrow", "medium", "wide"]


class AddConnectorInput(DocumentInput):
    """Attach a new connector to two named shapes."""

    slide_index: int = Field(ge=1)
    connector_type: Literal["straight", "elbow", "curve"] = "elbow"
    begin_shape: str = Field(min_length=1)
    begin_site: ConnectionSite = 1
    end_shape: str = Field(min_length=1)
    end_site: ConnectionSite = 1


class FormatConnectorInput(ShapeInput, LineOptions):
    """Format a connector and optionally reconnect either endpoint."""

    begin_arrow: ArrowStyle | None = None
    begin_arrow_length: ArrowLength | None = None
    begin_arrow_width: ArrowWidth | None = None
    end_arrow: ArrowStyle | None = None
    end_arrow_length: ArrowLength | None = None
    end_arrow_width: ArrowWidth | None = None
    begin_shape: str | None = Field(default=None, min_length=1)
    begin_site: ConnectionSite | None = None
    end_shape: str | None = Field(default=None, min_length=1)
    end_site: ConnectionSite | None = None

    @model_validator(mode="after")
    def validate_endpoints(self) -> "FormatConnectorInput":
        """Require a shape whenever a connection site is specified."""
        if self.begin_site is not None and self.begin_shape is None:
            raise ValueError("begin_site requires begin_shape.")
        if self.end_site is not None and self.end_shape is None:
            raise ValueError("end_site requires end_shape.")
        return self


def resolve_site(shape: NativeObject, site: ConnectionSite) -> int:
    """Choose a native connection site, rejecting invalid integer indices."""
    if isinstance(site, int):
        count = int(shape.ConnectionSiteCount)
        if site > count:
            raise ValueError(
                f"Shape '{shape.Name}' has {count} connection sites."
            )
        return site
    capabilities_for(shape).require("connection_sites")
    sites = shape.ConnectionSites
    count = int(sites.Count)
    if not count:
        raise ValueError(f"Shape '{shape.Name}' has no connection sites.")
    center_x = shape.Left + shape.Width / 2
    center_y = shape.Top + shape.Height / 2
    vector = SITE_DIRECTION_VECTORS[site]
    best_index, best_dot = 1, float("-inf")
    for index in range(1, count + 1):
        point = sites(index)
        delta_x, delta_y = point[0] - center_x, point[1] - center_y
        distance = (delta_x * delta_x + delta_y * delta_y) ** 0.5
        if distance:
            dot = (delta_x * vector[0] + delta_y * vector[1]) / distance
            if dot > best_dot:
                best_index, best_dot = index, dot
    return best_index


@tool("wps_ppt_connectors", component="wpp", operation="add")
def add_connector(params: AddConnectorInput) -> JsonObject:
    """Connect named shapes with a straight, elbow or curved connector.

    Sites are one-based indices or top/bottom/left/right directions. Both
    endpoints are resolved before the new connector is created.
    """
    session = get_session("wpp")
    slide = session.get_document().Slides(params.slide_index)
    begin = resolve_shape(slide, params.begin_shape)
    end = resolve_shape(slide, params.end_shape)
    begin_site = resolve_site(begin, params.begin_site)
    end_site = resolve_site(end, params.end_site)
    goto_slide(params.slide_index)
    connector = slide.Shapes.AddConnector(
        CONNECTOR_TYPE_MAP[params.connector_type], 0, 0, 100, 100
    )
    connector.ConnectorFormat.BeginConnect(begin, begin_site)
    connector.ConnectorFormat.EndConnect(end, end_site)
    connector.RerouteConnections()
    return {
        "success": True,
        "shape_name": connector.Name,
        "connector_type": params.connector_type,
    }


@tool(
    "wps_ppt_connectors",
    component="wpp",
    annotations={"idempotentHint": True},
    operation="format",
)
def format_connector(params: FormatConnectorInput) -> JsonObject:
    """Set line and arrowhead properties and optionally reconnect endpoints."""
    session = get_session("wpp")
    slide = session.get_document().Slides(params.slide_index)
    shape = resolve_shape(slide, params.shape_name_or_index)
    if not shape.Connector:
        raise ValueError(f"Shape '{shape.Name}' is not a connector.")
    connections: list[tuple[str, NativeObject, int]] = []
    endpoints: tuple[tuple[str, str | None, ConnectionSite | None], ...] = (
        ("BeginConnect", params.begin_shape, params.begin_site),
        ("EndConnect", params.end_shape, params.end_site),
    )
    for name, reference, site in endpoints:
        if reference is not None:
            target = resolve_shape(slide, reference)
            connections.append((name, target, resolve_site(target, site or 1)))
    goto_slide(params.slide_index)
    result = apply_formatting(shape, params)
    for attribute, value, mapping in (
        ("BeginArrowheadStyle", params.begin_arrow, ARROW_STYLE_MAP),
        ("BeginArrowheadLength", params.begin_arrow_length, ARROW_LENGTH_MAP),
        ("BeginArrowheadWidth", params.begin_arrow_width, ARROW_WIDTH_MAP),
        ("EndArrowheadStyle", params.end_arrow, ARROW_STYLE_MAP),
        ("EndArrowheadLength", params.end_arrow_length, ARROW_LENGTH_MAP),
        ("EndArrowheadWidth", params.end_arrow_width, ARROW_WIDTH_MAP),
    ):
        if value is not None:
            setattr(shape.Line, attribute, mapping[value])
    for name, target, site in connections:
        getattr(shape.ConnectorFormat, name)(target, site)
    if connections:
        shape.RerouteConnections()
    return result
