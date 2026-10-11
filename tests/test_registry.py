"""All tools share one strict contract and isolated per-call targeting."""

import inspect
import re
from contextlib import nullcontext
from types import SimpleNamespace
from typing import get_type_hints
from unittest.mock import Mock

import anyio
import pytest
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import CallToolResult
from pydantic import ValidationError

from src.models import Component, DocumentInput, InputModel, JsonObject
from src.native.context import CALL_DOCUMENT
from src.native.runtime import Runtime
from src.native.session import WpsSession
from src.protocol.adapters import make_adapter
from src.protocol.definitions import DEFINITIONS, ToolDefinition
from src.protocol.server import build_server, lifespan


def test_tool_catalog_is_complete_and_has_one_model_contract() -> None:
    server = build_server()
    assert server.runtime is None
    catalog = server._tool_manager.list_tools()
    assert len(catalog) == 45
    assert len(DEFINITIONS) == 261
    assert {registered.name for registered in catalog} == {
        definition.name for definition in DEFINITIONS.values()
    }
    operations = {(item.name, item.operation) for item in DEFINITIONS.values()}
    assert len(operations) == len(DEFINITIONS)
    assert all(operation is not None for _, operation in operations)
    for definition in DEFINITIONS.values():
        assert re.fullmatch(r"[a-z][a-z0-9_]*", definition.name)
        assert list(inspect.signature(definition.function).parameters) == [
            "params"
        ]
        model = get_type_hints(definition.function)["params"]
        assert issubclass(model, InputModel)
        assert model.model_config.get("strict") is True
        assert model.model_config.get("extra") == "forbid"
        assert model.model_config.get("allow_inf_nan") is False
        assert not model.model_config.get("str_strip_whitespace", False)
    for registered in catalog:
        assert list(registered.parameters["properties"]) == ["params"]


def test_ppt_connect_schema_requires_true_visibility() -> None:
    server = build_server()
    registered = next(
        tool
        for tool in server._tool_manager.list_tools()
        if tool.name == "wps_ppt_app"
    )
    connect = registered.parameters["$defs"]["wps_ppt_app_connect"]
    visible = connect["properties"]["visible"]
    assert visible["type"] == "boolean"
    assert visible["const"] is True
    assert visible["default"] is True
    assert "visible" not in connect["required"]


def test_lifespan_restarts_without_reusing_closed_sessions() -> None:
    async def scenario() -> None:
        server = build_server()
        async with lifespan(server):
            first = server.runtime
            assert first is not None
            result = await server.call_tool(
                "wps_components", {"params": {"operation": "detect"}}
            )
            assert isinstance(result, CallToolResult)
            assert not result.is_error
            assert result.structured_content is not None
            assert "components" in result.structured_content
            assert all(
                session.worker.thread is None
                for session in first.sessions.values()
            )
        assert server.runtime is None
        async with lifespan(server):
            assert server.runtime is not first
            result = await server.call_tool(
                "wps_components", {"params": {"operation": "detect"}}
            )
            assert isinstance(result, CallToolResult)
            assert not result.is_error

    anyio.run(scenario)


def test_invalid_calls_fail_before_native_launch() -> None:
    async def scenario() -> None:
        server = build_server()
        async with lifespan(server):
            for arguments in (
                {"args": {}},
                {"params": "{}"},
                {"params": None},
                {"params": []},
                {"params": {"operation": "missing"}},
                {"params": {"component": "writer", "extra": True}},
                {
                    "params": {
                        "operation": "connect",
                        "component": "writer",
                        "extra": True,
                    }
                },
                {
                    "params": {
                        "operation": "connect",
                        "component": "writer",
                        "visible": "false",
                    }
                },
            ):
                with pytest.raises(ToolError):
                    await server.call_tool("wps_components", arguments)
            assert server.runtime is not None
            assert all(
                session.worker.thread is None
                for session in server.runtime.sessions.values()
            )

    anyio.run(scenario)


@pytest.mark.parametrize(
    ("name", "params"),
    [
        ("wps_ppt_app", {"operation": "connect"}),
        ("wps_components", {"operation": "connect", "component": "wpp"}),
        ("wps_documents", {"operation": "create", "component": "wpp"}),
        (
            "wps_documents",
            {
                "operation": "open",
                "component": "wpp",
                "file_path": "C:/missing.pptx",
            },
        ),
    ],
)
@pytest.mark.parametrize("visible", [False, None, 0, 1, 1.0, "true", "false"])
def test_invalid_wpp_visibility_fails_before_native_launch(
    name: str, params: JsonObject, visible: object
) -> None:
    async def scenario() -> None:
        server = build_server()
        async with lifespan(server):
            message = (
                "WPP cannot be hidden; use visible=true"
                if visible is False
                else "validation error"
            )
            with pytest.raises(ToolError, match=message):
                await server.call_tool(
                    name, {"params": {**params, "visible": visible}}
                )
            assert server.runtime is not None
            assert all(
                session.worker.thread is None
                for session in server.runtime.sessions.values()
            )

    anyio.run(scenario)


@pytest.mark.parametrize(
    ("name", "component"),
    [
        ("wps_components", "writer"),
        ("wps_components", "et"),
        ("wps_components", "wpp"),
        ("wps_ppt_app", "wpp"),
    ],
)
def test_connect_defaults_to_visible_application(
    name: str, component: Component
) -> None:
    collection = SimpleNamespace(Count=0)
    application = SimpleNamespace(
        Name="Fake WPS",
        Version="test",
        Visible=False,
        Documents=collection,
        Workbooks=collection,
        Presentations=collection,
    )
    connector = Mock(return_value=application)

    async def scenario() -> None:
        with Runtime(
            lambda kind: WpsSession(
                kind, connector=connector, scope=nullcontext
            )
        ) as runtime:
            server = build_server()
            server.runtime = runtime
            params: JsonObject = {"operation": "connect"}
            if name == "wps_components":
                params["component"] = component
            result = await server.call_tool(name, {"params": params})
            assert isinstance(result, CallToolResult)
            assert not result.is_error
            assert application.Visible is True
            if component in {"writer", "et"}:
                result = await server.call_tool(
                    name, {"params": {**params, "visible": False}}
                )
                assert isinstance(result, CallToolResult)
                assert not result.is_error
                assert application.Visible is False
            connector.assert_called_once_with(
                runtime.sessions[component].spec.prog_id, True
            )

    anyio.run(scenario)


def test_call_target_is_restored_on_success_and_error() -> None:
    def read_context(params: DocumentInput) -> JsonObject:
        return {"document": CALL_DOCUMENT.get()}

    def fail(params: DocumentInput) -> JsonObject:
        raise ValueError("partial native failure")

    adapter = make_adapter(
        ToolDefinition("test_context", read_context, None, {})
    )
    failure = make_adapter(ToolDefinition("test_failure", fail, None, {}))

    async def scenario() -> None:
        results: dict[str, object] = {}
        async with anyio.create_task_group() as group:

            async def invoke(name: str) -> None:
                results[name] = await adapter(DocumentInput(document=name))

            group.start_soon(invoke, "first")
            group.start_soon(invoke, "second")
        assert results == {
            "first": {"document": "first"},
            "second": {"document": "second"},
        }
        with pytest.raises(ToolError, match="partial native failure"):
            await failure(DocumentInput(document="failed"))
        assert CALL_DOCUMENT.get() is None

    anyio.run(scenario)


def test_models_preserve_text_and_reject_coercion_and_nonfinite_numbers() -> (
    None
):
    class Input(InputModel):
        text: str
        count: int
        width: float

    assert Input(text="  内容  ", count=1, width=2).text == "  内容  "
    for change in (
        {"count": True},
        {"width": float("nan")},
        {"width": float("inf")},
    ):
        with pytest.raises(ValidationError):
            Input.model_validate(
                {"text": "ok", "count": 1, "width": 2, **change}
            )
