"""Opt-in integration tests; every edited document is created by the test."""

import os
from pathlib import Path
from uuid import uuid4
from zipfile import ZipFile

import anyio
import pytest

from tests.native.client import native_client
from wps_mcp.models import Component


@pytest.mark.parametrize("component", ["writer", "et", "wpp"])
def test_native_edit_save_reopen_and_preview(
    component: Component, tmp_path: Path
) -> None:
    async def scenario() -> None:
        extension = {"writer": ".docx", "et": ".xlsx", "wpp": ".pptx"}[
            component
        ]
        target = tmp_path / ("native-smoke" + extension)
        async with native_client() as client:
            detected = await client.call("wps_components", operation="detect")
            assert any(
                item["component"] == component
                for item in detected["components"]
            )
            connected = await client.call(
                "wps_components", operation="connect", component=component
            )
            assert connected["component"] == component
            listed = await client.call(
                "wps_documents", operation="list", component=component
            )
            assert isinstance(listed["documents"], list)
            created = await client.call(
                "wps_documents", operation="create", component=component
            )
            identity = created["full_name"]
            if component == "writer":
                await client.call(
                    "wps_writer_text",
                    operation="insert",
                    document=identity,
                    position=0,
                    text="WPS native smoke\r中文保留  空白\r",
                )
                text = await client.call(
                    "wps_writer_text", operation="read", document=identity
                )
                assert "中文保留  空白" in text["text"]
            elif component == "et":
                await client.call(
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
                data = await client.call(
                    "wps_et_ranges",
                    operation="read",
                    document=identity,
                    address="A1:B3",
                    include_formulas=True,
                )
                assert data["values"][2][0] == "=1+1"
                await client.call(
                    "wps_et_ranges",
                    operation="write_formulas",
                    document=identity,
                    address="C2",
                    formulas=[["=B2+B3"]],
                )
                await client.call(
                    "wps_et_sheets",
                    operation="recalculate",
                    document=identity,
                )
                computed = await client.call(
                    "wps_et_ranges",
                    operation="read",
                    document=identity,
                    address="C2",
                )
                assert computed["values"][0][0] == 46
            else:
                await client.call(
                    "wps_ppt_slides",
                    operation="add",
                    document=identity,
                    layout="blank",
                )
                shape = await client.call(
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
                text = await client.call(
                    "wps_ppt_text",
                    operation="get",
                    document=identity,
                    slide_index=1,
                    shape_name_or_index=shape["shape_name"],
                )
                assert "中文保留  空白" in text["text"]
                runs = await client.call(
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
                span = await client.call(
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
                await client.call(
                    "wps_ppt_text",
                    operation="set",
                    document=identity,
                    slide_index=1,
                    shape_name_or_index=shape["shape_name"],
                    text="😀a😀a",
                )
                await client.call(
                    "wps_ppt_text",
                    operation="replace",
                    document=identity,
                    find_text="a",
                    replace_text="🚀",
                    context_chars=1,
                )
                updated = await client.call(
                    "wps_ppt_text",
                    operation="get",
                    document=identity,
                    slide_index=1,
                    shape_name_or_index=shape["shape_name"],
                )
                assert updated["text"] == "😀🚀😀🚀"
                await client.call(
                    "wps_ppt_text",
                    operation="set",
                    document=identity,
                    slide_index=1,
                    shape_name_or_index=shape["shape_name"],
                    text="WPS native smoke\r中文保留  空白",
                )
                await client.call(
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
                animation = await client.call(
                    "wps_ppt_animation",
                    operation="add",
                    document=identity,
                    slide_index=1,
                    shape_name_or_index=shape["shape_name"],
                    effect="appear",
                )
                assert animation["success"]
                await client.call(
                    "wps_ppt_slides",
                    operation="add",
                    document=identity,
                    layout="blank",
                )
                chart = await client.call(
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
                plotted = await client.call(
                    "wps_ppt_charts",
                    operation="get_data",
                    document=identity,
                    slide_index=2,
                    shape_name_or_index=chart["shape_name"],
                )
                assert plotted["categories"] and plotted["series"]
                await client.call(
                    "wps_documents",
                    operation="export_pdf",
                    component="wpp",
                    document=identity,
                    file_path=str(tmp_path / "selected.pdf"),
                    slide_range_start=1,
                    slide_range_end=1,
                )
                exported = await client.call(
                    "wps_ppt_export",
                    operation="images",
                    document=identity,
                    output_dir=str(tmp_path / "slides"),
                    width=640,
                )
                assert exported["files_count"] == 2
            saved = await client.call(
                "wps_documents",
                operation="save",
                component=component,
                document=identity,
                file_path=str(target),
            )
            identity = saved["full_name"]
            assert target.is_file() and saved["saved"]
            if component == "writer":
                marker = "Unsaved native marker " + uuid4().hex
                await client.call(
                    "wps_writer_text",
                    operation="insert",
                    document=identity,
                    position=0,
                    text=marker,
                )
                text = await client.call(
                    "wps_writer_text", operation="read", document=identity
                )
                state = await client.call(
                    "wps_documents",
                    operation="info",
                    component=component,
                    document=identity,
                )
                assert marker in text["text"]
                assert state["saved"] is False
                with ZipFile(target) as archive:
                    assert marker.encode() not in archive.read(
                        "word/document.xml"
                    )
                saved = await client.call(
                    "wps_documents",
                    operation="save",
                    component=component,
                    document=identity,
                )
                assert saved["saved"]
            pdf = tmp_path / "document.pdf"
            await client.call(
                "wps_documents",
                operation="export_pdf",
                component=component,
                document=identity,
                file_path=str(pdf),
            )
            assert pdf.read_bytes().startswith(b"%PDF")
            raw = await client.image(
                "wps_documents",
                "preview",
                component=component,
                document=identity,
                width=640,
            )
            output = Path(os.environ.get("WPS_NATIVE_OUTPUT", str(tmp_path)))
            output.mkdir(parents=True, exist_ok=True)
            (output / f"{component}.png").write_bytes(raw)
            await client.call(
                "wps_documents",
                operation="close",
                component=component,
                document=identity,
                save_changes=True,
            )
            reopened = await client.call(
                "wps_documents",
                operation="open",
                component=component,
                file_path=str(target),
            )
            identity = reopened["full_name"]
            info = await client.call(
                "wps_documents",
                operation="info",
                component=component,
                document=identity,
            )
            assert info["saved"]

    anyio.run(scenario)
