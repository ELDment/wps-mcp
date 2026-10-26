"""Track and close only the documents created by a native test."""

import base64
import json
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import anyio
from mcp.types import CallToolResult, ImageContent, TextContent

from wps_mcp.models import Component, JsonObject
from wps_mcp.protocol.server import WpsServer, build_server, lifespan


@dataclass
class NativeClient:
    server: WpsServer
    documents: dict[tuple[Component, str], None] = field(default_factory=dict)

    async def _invoke(
        self, tool_name: str, operation: str, **params: Any
    ) -> CallToolResult:
        print(f"{tool_name}/{operation}", flush=True)
        result = await self.server.call_tool(
            tool_name, {"params": {"operation": operation, **params}}
        )
        assert isinstance(result, CallToolResult)
        assert not result.is_error, result
        return result

    async def call(
        self, tool_name: str, operation: str, **params: Any
    ) -> JsonObject:
        result = await self._invoke(tool_name, operation, **params)
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
                self.documents.pop((component, payload["closed"]), None)
        elif tool_name == "wps_writer_mail_merge" and operation == "execute":
            self.documents["writer", payload["full_name"]] = None
        return payload

    async def image(
        self, tool_name: str, operation: str, **params: Any
    ) -> bytes:
        result = await self._invoke(tool_name, operation, **params)
        content = result.content[0]
        assert isinstance(content, ImageContent)
        assert content.mime_type == "image/png"
        data = base64.b64decode(content.data, validate=True)
        assert data.startswith(b"\x89PNG\r\n\x1a\n")
        return data

    async def save_and_reopen(
        self, component: Component, identity: str, target: Path
    ) -> str:
        saved = await self.call(
            "wps_documents",
            "save",
            component=component,
            document=identity,
            file_path=str(target),
        )
        assert (
            saved["saved"] and target.is_file() and target.stat().st_size > 0
        )
        await self.call(
            "wps_documents",
            "close",
            component=component,
            document=saved["full_name"],
        )
        opened = await self.call(
            "wps_documents", "open", component=component, file_path=str(target)
        )
        identity = opened["full_name"]
        assert isinstance(identity, str)
        return identity


@asynccontextmanager
async def native_client() -> AsyncIterator[NativeClient]:
    server = build_server()
    async with lifespan(server):
        client = NativeClient(server)
        primary_error: BaseException | None = None
        try:
            yield client
        except BaseException as error:
            primary_error = error
            raise
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
                raise BaseExceptionGroup(
                    "Could not close native test documents",
                    ([primary_error] if primary_error is not None else [])
                    + failures,
                )
