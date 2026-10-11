"""Track and close only the documents created by a native test."""

import json
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from dataclasses import dataclass, field
from typing import Any

import anyio
from mcp.types import CallToolResult, TextContent

from src.models import Component, JsonObject
from src.protocol.server import WpsServer, build_server, lifespan


@dataclass
class NativeClient:
    server: WpsServer
    documents: dict[tuple[Component, str], None] = field(default_factory=dict)

    async def call(
        self, tool_name: str, operation: str, **params: Any
    ) -> JsonObject:
        print(f"{tool_name}/{operation}", flush=True)
        result = await self.server.call_tool(
            tool_name, {"params": {"operation": operation, **params}}
        )
        assert isinstance(result, CallToolResult)
        assert not result.is_error, result
        content = result.content[0]
        assert isinstance(content, TextContent)
        payload = json.loads(content.text)
        assert isinstance(payload, dict)
        if tool_name == "wps_documents":
            component = params["component"]
            if operation in {"create", "open", "save"}:
                self.documents.pop(
                    (component, params.get("document", "")), None
                )
                self.documents[component, payload["full_name"]] = None
            elif operation == "close":
                self.documents.pop((component, params["document"]), None)
        elif tool_name == "wps_writer_mail_merge" and operation == "execute":
            self.documents["writer", payload["full_name"]] = None
        return payload


@asynccontextmanager
async def native_client() -> AsyncIterator[NativeClient]:
    server = build_server()
    async with lifespan(server):
        client = NativeClient(server)
        try:
            yield client
        finally:
            failures = []
            with anyio.move_on_after(60, shield=True) as cleanup:
                for component, identity in reversed(list(client.documents)):
                    try:
                        await client.call(
                            "wps_documents",
                            "close",
                            component=component,
                            document=identity,
                            discard_changes=True,
                        )
                    except Exception as error:
                        failures.append(
                            RuntimeError(
                                f"Test document still open: "
                                f"{component}/{identity}: {error}"
                            )
                        )
            if cleanup.cancel_called:
                failures.append(
                    TimeoutError(
                        f"Cleanup timed out: {list(client.documents)}"
                    )
                )
            if failures:
                raise ExceptionGroup(
                    "Could not close native test documents", failures
                )
