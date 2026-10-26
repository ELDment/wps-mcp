"""Keep test document ownership and image decoding explicit."""

import base64
import json
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from pathlib import Path
from typing import cast
from unittest.mock import AsyncMock, Mock

import anyio
import pytest
from mcp.types import CallToolResult, ImageContent, TextContent

from tests.native import client as client_module
from tests.native.client import NativeClient, native_client
from wps_mcp.models import JsonObject
from wps_mcp.protocol.server import WpsServer


def response(payload: JsonObject) -> CallToolResult:
    return CallToolResult(
        content=[TextContent(type="text", text=json.dumps(payload))]
    )


def test_document_tracking_follows_save_as_and_resolved_close() -> None:
    server = Mock(spec=WpsServer)
    server.call_tool = AsyncMock(
        side_effect=[
            response({"full_name": "New document"}),
            response({"full_name": "saved.docx", "saved": True}),
            response({"closed": "saved.docx"}),
            response({"full_name": "merged.docx"}),
        ]
    )
    client = NativeClient(cast(WpsServer, server))

    async def scenario() -> None:
        await client.call("wps_documents", "create", component="writer")
        assert set(client.documents) == {("writer", "New document")}
        await client.call(
            "wps_documents",
            "save",
            component="writer",
            document="New document",
            file_path="saved.docx",
        )
        assert set(client.documents) == {("writer", "saved.docx")}
        await client.call(
            "wps_documents", "close", component="writer", document="saved"
        )
        assert not client.documents
        await client.call(
            "wps_writer_mail_merge", "execute", document="template.docx"
        )
        assert set(client.documents) == {("writer", "merged.docx")}

    anyio.run(scenario)


def test_save_and_reopen_transfers_cleanup_ownership(tmp_path: Path) -> None:
    target = tmp_path / "roundtrip.xlsx"
    target.write_bytes(b"saved workbook")
    server = Mock(spec=WpsServer)
    server.call_tool = AsyncMock(
        side_effect=[
            response({"full_name": str(target), "saved": True}),
            response({"closed": str(target)}),
            response({"full_name": str(target)}),
        ]
    )
    client = NativeClient(cast(WpsServer, server))
    client.documents["et", "Workbook1"] = None

    async def scenario() -> None:
        assert await client.save_and_reopen("et", "Workbook1", target) == str(
            target
        )

    anyio.run(scenario)
    assert set(client.documents) == {("et", str(target))}
    assert [
        call.args[1]["params"]["operation"]
        for call in server.call_tool.await_args_list
    ] == ["save", "close", "open"]


@pytest.mark.parametrize("failed", [False, True])
def test_image_response_checks_errors_before_decoding(failed: bool) -> None:
    image = b"\x89PNG\r\n\x1a\nimage payload"
    server = Mock(spec=WpsServer)
    server.call_tool = AsyncMock(
        return_value=CallToolResult(
            is_error=failed,
            content=[
                ImageContent(
                    type="image",
                    mime_type="image/png",
                    data=base64.b64encode(image).decode("ascii"),
                )
            ],
        )
    )
    client = NativeClient(cast(WpsServer, server))

    async def scenario() -> None:
        if failed:
            with pytest.raises(AssertionError):
                await client.image("wps_ppt_export", "preview")
        else:
            assert await client.image("wps_ppt_export", "preview") == image

    anyio.run(scenario)


def test_cleanup_closes_only_tracked_documents_after_failure(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    server = Mock(spec=WpsServer)
    server.call_tool = AsyncMock(
        side_effect=[
            response({"full_name": "first.docx"}),
            response({"full_name": "second.xlsx"}),
            response({"closed": "second.xlsx"}),
            response({"closed": "first.docx"}),
        ]
    )

    @asynccontextmanager
    async def lifespan(unused: WpsServer) -> AsyncIterator[JsonObject]:
        yield {}

    monkeypatch.setattr(client_module, "build_server", lambda: server)
    monkeypatch.setattr(client_module, "lifespan", lifespan)

    async def scenario() -> None:
        with pytest.raises(ValueError, match="scenario failed"):
            async with native_client() as client:
                await client.call(
                    "wps_documents", "create", component="writer"
                )
                await client.call("wps_documents", "create", component="et")
                raise ValueError("scenario failed")

    anyio.run(scenario)
    calls = server.call_tool.await_args_list
    assert [call.args[1]["params"]["document"] for call in calls[2:]] == [
        "second.xlsx",
        "first.docx",
    ]
    assert all(call.args[1]["params"]["discard_changes"] for call in calls[2:])


def test_cleanup_preserves_operation_and_cleanup_failures(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    server = Mock(spec=WpsServer)
    server.call_tool = AsyncMock(
        side_effect=[response({"full_name": "first.docx"}), OSError("locked")]
    )

    @asynccontextmanager
    async def lifespan(unused: WpsServer) -> AsyncIterator[JsonObject]:
        yield {}

    monkeypatch.setattr(client_module, "build_server", lambda: server)
    monkeypatch.setattr(client_module, "lifespan", lifespan)

    async def scenario() -> None:
        with pytest.raises(ExceptionGroup) as caught:
            async with native_client() as client:
                await client.call(
                    "wps_documents", "create", component="writer"
                )
                raise ValueError("operation failed")
        original, cleanup = caught.value.exceptions
        assert isinstance(original, ValueError)
        assert str(original) == "operation failed"
        assert "locked" in str(cleanup)

    anyio.run(scenario)
