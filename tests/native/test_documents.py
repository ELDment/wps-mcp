"""Opt-in integration tests; every edited document is created by the test."""

import base64
import json
import os
from pathlib import Path
from typing import Any

import anyio
import pytest
from mcp.types import CallToolResult, ImageContent, TextContent

from src.models import Component, JsonObject
from src.protocol.server import build_server, lifespan


@pytest.mark.parametrize("component", ["writer", "et", "wpp"])
def test_native_edit_save_reopen_and_preview(
    component: Component, tmp_path: Path
) -> None:
    async def scenario() -> None:
        server = build_server()
        identity: str | None = None
        extension = {"writer": ".docx", "et": ".xlsx", "wpp": ".pptx"}[
            component
        ]
        target = tmp_path / ("native-smoke" + extension)
        async with lifespan(server):

            async def invoke(name: str, **params: Any) -> JsonObject:
                print(f"{component}: {name}", flush=True)
                result = await server.call_tool(name, {"params": params})
                print(f"{component}: {name} completed", flush=True)
                assert isinstance(result, CallToolResult)
                assert not result.is_error
                content = result.content[0]
                assert isinstance(content, TextContent)
                payload = json.loads(content.text)
                assert isinstance(payload, dict)
                return payload

            created = await invoke(
                "wps_documents", operation="create", component=component
            )
            identity = created["full_name"]
            try:
                if component == "writer":
                    await invoke(
                        "wps_writer_text",
                        operation="insert",
                        document=identity,
                        position=0,
                        text="WPS native smoke\r中文保留  空白\r",
                    )
                    text = await invoke(
                        "wps_writer_text", operation="read", document=identity
                    )
                    assert "中文保留  空白" in text["text"]
                elif component == "et":
                    await invoke(
                        "wps_et_ranges",
                        operation="write_values",
                        document=identity,
                        address="A1:B3",
                        values=[
                            ["名称", "数值"],
                            ["普通文字", 12],
                            ["=1+1", 34],
                        ],
                    )
                    data = await invoke(
                        "wps_et_ranges",
                        operation="read",
                        document=identity,
                        address="A1:B3",
                        include_formulas=True,
                    )
                    assert data["values"][2][0] == "=1+1"
                    await invoke(
                        "wps_et_ranges",
                        operation="write_formulas",
                        document=identity,
                        address="C2",
                        formulas=[["=B2+B3"]],
                    )
                    await invoke(
                        "wps_et_sheets",
                        operation="recalculate",
                        document=identity,
                    )
                    computed = await invoke(
                        "wps_et_ranges",
                        operation="read",
                        document=identity,
                        address="C2",
                    )
                    assert computed["values"][0][0] == 46
                else:
                    await invoke(
                        "wps_ppt_slides",
                        operation="add",
                        document=identity,
                        layout="blank",
                    )
                    shape = await invoke(
                        "wps_ppt_shapes",
                        operation="add_textbox",
                        document=identity,
                        slide_index=1,
                        left=40,
                        top=40,
                        width=500,
                        height=100,
                        text="WPS native smoke\r中文保留  空白",
                        font_size=28,
                    )
                    text = await invoke(
                        "wps_ppt_text",
                        operation="get",
                        document=identity,
                        slide_index=1,
                        shape_name_or_index=shape["shape_name"],
                    )
                    assert "中文保留  空白" in text["text"]
                    runs = await invoke(
                        "wps_ppt_text",
                        operation="set",
                        document=identity,
                        slide_index=1,
                        shape_name_or_index=shape["shape_name"],
                        runs=[
                            {"text": "A😀", "bold": True},
                            {"text": "B中文"},
                        ],
                    )
                    assert [
                        (run["start"], run["length"]) for run in runs["runs"]
                    ] == [
                        (1, 3),
                        (4, 3),
                    ]
                    span = await invoke(
                        "wps_ppt_text",
                        operation="format_range",
                        document=identity,
                        slide_index=1,
                        shape_name_or_index=shape["shape_name"],
                        search_text="B中",
                        bold=True,
                    )
                    assert (
                        span["start"],
                        span["length"],
                        span["formatted_text"],
                    ) == (4, 2, "B中")
                    await invoke(
                        "wps_ppt_text",
                        operation="set",
                        document=identity,
                        slide_index=1,
                        shape_name_or_index=shape["shape_name"],
                        text="😀a😀a",
                    )
                    await invoke(
                        "wps_ppt_text",
                        operation="replace",
                        document=identity,
                        find_text="a",
                        replace_text="🚀",
                        context_chars=1,
                    )
                    updated = await invoke(
                        "wps_ppt_text",
                        operation="get",
                        document=identity,
                        slide_index=1,
                        shape_name_or_index=shape["shape_name"],
                    )
                    assert updated["text"] == "😀🚀😀🚀"
                    await invoke(
                        "wps_ppt_text",
                        operation="set",
                        document=identity,
                        slide_index=1,
                        shape_name_or_index=shape["shape_name"],
                        text="WPS native smoke\r中文保留  空白",
                    )
                    await invoke(
                        "wps_ppt_tables",
                        operation="add",
                        document=identity,
                        slide_index=1,
                        rows=2,
                        cols=2,
                        left=40,
                        top=180,
                        width=400,
                        height=100,
                    )
                    animation = await invoke(
                        "wps_ppt_animation",
                        operation="add",
                        document=identity,
                        slide_index=1,
                        shape_name_or_index=shape["shape_name"],
                        effect="appear",
                    )
                    assert animation["success"]
                    await invoke(
                        "wps_ppt_slides",
                        operation="add",
                        document=identity,
                        layout="blank",
                    )
                    chart = await invoke(
                        "wps_ppt_charts",
                        operation="add",
                        document=identity,
                        slide_index=2,
                        chart_type="column",
                        left=40,
                        top=40,
                        width=500,
                        height=300,
                    )
                    plotted = await invoke(
                        "wps_ppt_charts",
                        operation="get_data",
                        document=identity,
                        slide_index=2,
                        shape_name_or_index=chart["shape_name"],
                    )
                    assert plotted["categories"] and plotted["series"]
                    await invoke(
                        "wps_ppt_export",
                        operation="pdf",
                        document=identity,
                        file_path=str(tmp_path / "selected.pdf"),
                        slide_range_start=1,
                        slide_range_end=1,
                    )
                    exported = await invoke(
                        "wps_ppt_export",
                        operation="images",
                        document=identity,
                        output_dir=str(tmp_path / "slides"),
                        width=640,
                    )
                    assert exported["files_count"] == 2
                saved = await invoke(
                    "wps_documents",
                    operation="save",
                    component=component,
                    document=identity,
                    file_path=str(target),
                )
                identity = saved["full_name"]
                assert target.is_file() and saved["saved"]
                preview = await server.call_tool(
                    "wps_documents",
                    {
                        "params": {
                            "operation": "preview",
                            "component": component,
                            "document": identity,
                            "width": 640,
                        }
                    },
                )
                assert isinstance(preview, CallToolResult)
                image = preview.content[0]
                assert isinstance(image, ImageContent)
                raw = base64.b64decode(image.data)
                assert raw.startswith(b"\x89PNG\r\n\x1a\n")
                output = Path(
                    os.environ.get("WPS_NATIVE_OUTPUT", str(tmp_path))
                )
                output.mkdir(parents=True, exist_ok=True)
                (output / f"{component}.png").write_bytes(raw)
                await invoke(
                    "wps_documents",
                    operation="close",
                    component=component,
                    document=identity,
                )
                identity = None
                reopened = await invoke(
                    "wps_documents",
                    operation="open",
                    component=component,
                    file_path=str(target),
                )
                identity = reopened["full_name"]
                info = await invoke(
                    "wps_documents",
                    operation="info",
                    component=component,
                    document=identity,
                )
                assert info["saved"]
            finally:
                if identity is not None:
                    await invoke(
                        "wps_documents",
                        operation="close",
                        component=component,
                        document=identity,
                        discard_changes=True,
                    )

    anyio.run(scenario)
