"""Verify the installed entry path through the actual MCP stdio transport."""

import os
import sys
from pathlib import Path

import anyio
import pytest
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from mcp.types import TextContent

from wps_mcp import __version__


@pytest.mark.parametrize("entry_point", ["module", "console"])
def test_stdio_catalog_validation_and_shutdown(entry_point: str) -> None:
    async def scenario() -> None:
        root = Path(__file__).resolve().parents[2]
        executable = Path(sys.executable)
        command = str(executable)
        arguments = ["-m", "wps_mcp"]
        if entry_point == "console":
            command = str(
                executable.with_name(
                    "wps-mcp.exe" if sys.platform == "win32" else "wps-mcp"
                )
            )
            arguments = []
        parameters = StdioServerParameters(
            command=command,
            args=arguments,
            cwd=root,
            env={**os.environ, "PYTHONPATH": str(root / "src")},
        )
        with anyio.fail_after(20):
            async with stdio_client(parameters) as (read, write):
                async with ClientSession(read, write) as client:
                    initialized = await client.initialize()
                    assert initialized.server_info.version == __version__
                    assert initialized.instructions is not None
                    assert "Do not attempt to hide WPP" in (
                        initialized.instructions
                    )
                    catalog = await client.list_tools()
                    assert len(catalog.tools) == 45
                    detected = await client.call_tool(
                        "wps_components", {"params": {"operation": "detect"}}
                    )
                    assert not detected.is_error
                    assert detected.structured_content is not None
                    rejected = await client.call_tool(
                        "wps_documents",
                        {
                            "params": {
                                "operation": "create",
                                "component": "writer",
                                "typo": True,
                            }
                        },
                    )
                    assert rejected.is_error
                    hidden = await client.call_tool(
                        "wps_components",
                        {
                            "params": {
                                "operation": "connect",
                                "component": "wpp",
                                "visible": False,
                            }
                        },
                    )
                    assert hidden.is_error
                    assert any(
                        isinstance(content, TextContent)
                        and "WPP cannot be hidden; use visible=true"
                        in content.text
                        for content in hidden.content
                    )

    anyio.run(scenario)
