"""Load built-in operation modules and publish their MCP schemas."""

from importlib import import_module
from typing import Any

from mcp.server.mcpserver import MCPServer
from mcp.types import ToolAnnotations

from wps_mcp.models import JsonObject
from wps_mcp.protocol.adapters import make_adapter
from wps_mcp.protocol.definitions import DEFINITIONS, ToolDefinition

MODULES = (
    "documents.tools",
    "writer.text",
    "writer.tables",
    "writer.pages",
    "writer.bookmarks",
    "writer.images",
    "writer.statistics",
    "writer.styles",
    "writer.review",
    "writer.references",
    "writer.merge",
    "spreadsheet.ranges",
    "spreadsheet.sheets",
    "spreadsheet.charts",
    "spreadsheet.pivots",
    "spreadsheet.tables",
    "spreadsheet.conditional",
    "spreadsheet.validation",
    "spreadsheet.names",
    *(
        f"presentation.{name}"
        for name in (
            "animations",
            "annotations",
            "application",
            "assets",
            "charts",
            "connectors",
            "documents",
            "exports",
            "freeform",
            "history",
            "layouts",
            "media",
            "pictures",
            "placeholders",
            "sections",
            "shapes.tools",
            "shapes.inspection",
            "slides",
            "slideshow",
            "smart_art",
            "styles",
            "tables",
            "text.tools",
            "themes",
            "text.typography",
        )
    ),
)


def load_definitions() -> None:
    """Import operation definitions without constructing an MCP server."""
    for module in MODULES:
        import_module(f"wps_mcp.{module}")


def register_tools(server: MCPServer[Any]) -> None:
    """Load built-in operations and publish deterministic tool groups."""
    load_definitions()
    groups: dict[str, list[ToolDefinition]] = {}
    for definition in DEFINITIONS.values():
        groups.setdefault(definition.name, []).append(definition)
    for name, definitions in sorted(groups.items()):
        annotations = {
            "readOnlyHint": all(
                item.annotations["readOnlyHint"] for item in definitions
            ),
            "destructiveHint": any(
                item.annotations["destructiveHint"] for item in definitions
            ),
            "idempotentHint": all(
                item.annotations["idempotentHint"] for item in definitions
            ),
            "openWorldHint": any(
                item.annotations["openWorldHint"] for item in definitions
            ),
        }
        server.add_tool(
            make_adapter(definitions),
            name=name,
            annotations=ToolAnnotations.model_validate(annotations),
            structured_output=all(
                item.output_type == JsonObject for item in definitions
            ),
        )
        registered = server._tool_manager.get_tool(name)
        if registered is not None:
            registered.parameters["additionalProperties"] = False
