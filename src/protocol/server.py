"""Build the stdio service without starting WPS or creating native workers."""

from __future__ import annotations

import logging
import sys
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from typing import Any

from mcp.server.mcpserver import Context, MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import CallToolResult, InputRequiredResult

from src.models import JsonObject
from src.native.context import CURRENT_RUNTIME
from src.native.runtime import Runtime
from src.protocol.catalog import register_tools


class WpsServer(MCPServer[Any]):
    """Validate outer arguments and isolate each service runtime."""

    runtime: Runtime | None = None

    async def call_tool(
        self,
        name: str,
        arguments: JsonObject,
        context: Context[Any, Any] | None = None,
    ) -> CallToolResult | InputRequiredResult:
        """Validate outer arguments and bind this server's isolated runtime."""
        if (
            not isinstance(arguments, dict)
            or set(arguments) != {"params"}
            or not isinstance(arguments["params"], dict)
        ):
            raise ToolError("Arguments must contain exactly one params object")
        if self.runtime is None:
            raise ToolError("The WPS server lifespan is not running")
        token = CURRENT_RUNTIME.set(self.runtime)
        try:
            return await super().call_tool(name, arguments, context)
        finally:
            CURRENT_RUNTIME.reset(token)


@asynccontextmanager
async def lifespan(server: MCPServer[Any]) -> AsyncIterator[JsonObject]:
    """Create and release one runtime for each active server lifespan."""
    if not isinstance(server, WpsServer):
        raise TypeError("WPS lifespan requires WpsServer")
    if server.runtime is not None:
        raise RuntimeError("Server lifespan is already running")
    with Runtime() as runtime:
        server.runtime = runtime
        try:
            yield {}
        finally:
            server.runtime = None


def build_server() -> WpsServer:
    """Build a complete tool catalog without connecting to native WPS."""
    server = WpsServer(
        "wps-mcp",
        lifespan=lifespan,
        instructions=(
            "Control WPS Writer, ET and Presentation on Windows and Linux. "
            "Linux uses MCP-managed WPS instances and requires CPython 3.11. "
            "Existing Linux desktop windows are not adopted. "
            "Start with wps_components using operation=detect. "
            "All tools take a params object with a required "
            "operation discriminator. "
            "Each operation has its own strict schema. "
            "All components default to visible=true. Do not attempt to hide "
            "WPP: visible=false is unsupported. In wps_ppt_document/open, "
            "with_window=true shows the application and document window; "
            "with_window=false opens without a document window and leaves "
            "application visibility unchanged. "
            "Use full_name returned by create/open as params.document "
            "on every subsequent call. "
            "Indices start at 1; Writer text positions are UTF-16 offsets "
            "starting at 0. "
            "Geometry uses points. Save new documents early. "
            "Inspect previews after editing. Native failures may leave "
            "partial changes; inspect before repeating writes. "
            "Advanced interfaces depend on the WPS build."
        ),
    )
    register_tools(server)
    return server


def main() -> None:
    """Run the stdio service, reserving stdout for protocol messages."""
    logging.basicConfig(level=logging.INFO, stream=sys.stderr)
    build_server().run(transport="stdio")


if __name__ == "__main__":
    main()
