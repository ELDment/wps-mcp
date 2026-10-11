"""Grouped tools validate each operation before dispatching native work."""

from contextlib import nullcontext
from typing import get_type_hints

import anyio
import pytest
from mcp.server.mcpserver import Image, MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from pydantic import Field, TypeAdapter

from src.models import Component, DocumentInput, InputModel, JsonObject
from src.native.context import CALL_DOCUMENT
from src.native.runtime import Runtime
from src.native.session import WpsSession
from src.protocol.adapters import make_adapter
from src.protocol.catalog import register_tools
from src.protocol.definitions import ToolDefinition


class ReadInput(DocumentInput):
    limit: int = Field(default=10, ge=1)


class WriteInput(DocumentInput):
    text: str
    bold: bool | None = None


def test_group_dispatch_preserves_supplied_fields_and_document_context() -> (
    None
):
    def read(params: ReadInput) -> JsonObject:
        return {"limit": params.limit, "document": CALL_DOCUMENT.get()}

    def write(params: WriteInput) -> JsonObject:
        assert type(params) is WriteInput
        return {
            "text": params.text,
            "supplied": sorted(params.model_fields_set),
        }

    adapter = make_adapter(
        [
            ToolDefinition("test_document", read, None, {}, "read"),
            ToolDefinition("test_document", write, None, {}, "write"),
        ]
    )
    schema = TypeAdapter(
        get_type_hints(adapter, include_extras=True)["params"]
    )
    assert schema.json_schema()["discriminator"]["propertyName"] == "operation"

    async def scenario() -> None:
        assert await adapter(
            schema.validate_python({"operation": "read", "document": "one"})
        ) == {
            "limit": 10,
            "document": "one",
        }
        assert await adapter(
            schema.validate_python({"operation": "write", "text": "  text  "})
        ) == {"text": "  text  ", "supplied": ["text"]}
        assert await adapter(
            schema.validate_python(
                {"operation": "write", "text": "x", "bold": None}
            )
        ) == {"text": "x", "supplied": ["bold", "text"]}
        assert CALL_DOCUMENT.get() is None

    anyio.run(scenario)


def test_group_rejects_wrong_operation_fields_before_execution() -> None:
    called = []

    def read(params: ReadInput) -> JsonObject:
        called.append(params)
        return {}

    def write(params: WriteInput) -> JsonObject:
        called.append(params)
        return {}

    server = MCPServer("group-test")
    server.add_tool(
        make_adapter(
            [
                ToolDefinition("test_document", read, None, {}, "read"),
                ToolDefinition("test_document", write, None, {}, "write"),
            ]
        )
    )

    async def scenario() -> None:
        for params in (
            {},
            {"operation": "missing"},
            {"operation": "read", "text": "wrong model"},
            {"operation": "write", "limit": 2},
            {"operation": "read", "limit": "2"},
            {"operation": "write", "text": "x", "bold": "false"},
        ):
            with pytest.raises(ToolError):
                await server.call_tool("test_document", {"params": params})
        assert called == []

    anyio.run(scenario)


@pytest.mark.parametrize("component", [None, "wpp"])
def test_group_can_return_both_data_and_an_image(
    component: Component | None,
) -> None:
    def info(params: InputModel) -> JsonObject:
        return {"count": 1}

    def preview(params: InputModel) -> bytes:
        return b"png-data"

    adapter = make_adapter(
        [
            ToolDefinition("test_output", info, component, {}, "info"),
            ToolDefinition("test_output", preview, component, {}, "preview"),
        ]
    )
    schema = TypeAdapter(
        get_type_hints(adapter, include_extras=True)["params"]
    )

    async def scenario() -> None:
        assert await adapter(
            schema.validate_python({"operation": "info"})
        ) == {"count": 1}
        result = await adapter(
            schema.validate_python({"operation": "preview"})
        )
        assert isinstance(result, Image)
        assert result.data == b"png-data"

    with Runtime(lambda name: WpsSession(name, scope=nullcontext)):
        anyio.run(scenario)


def test_group_annotations_cover_every_operation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from src.protocol import catalog as registry

    def read(params: InputModel) -> JsonObject:
        return {}

    def delete(params: InputModel) -> JsonObject:
        return {}

    common = {"openWorldHint": False}
    definitions = {
        "test_group/read": ToolDefinition(
            "test_group",
            read,
            None,
            common
            | {
                "readOnlyHint": True,
                "destructiveHint": False,
                "idempotentHint": True,
            },
            "read",
        ),
        "test_group/delete": ToolDefinition(
            "test_group",
            delete,
            None,
            common
            | {
                "readOnlyHint": False,
                "destructiveHint": True,
                "idempotentHint": False,
            },
            "delete",
        ),
    }
    monkeypatch.setattr(registry, "MODULES", ())
    monkeypatch.setattr(registry, "DEFINITIONS", definitions)
    server = MCPServer("hints-test")
    register_tools(server)
    registered = server._tool_manager.get_tool("test_group")
    assert registered is not None
    assert registered.annotations is not None
    assert registered.annotations.read_only_hint is False
    assert registered.annotations.destructive_hint is True
    assert registered.annotations.idempotent_hint is False
    assert registered.parameters["additionalProperties"] is False
