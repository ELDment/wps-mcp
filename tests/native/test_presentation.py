"""Exercise presentation tool groups using only test-owned native decks."""

from __future__ import annotations

import base64
import json
import math
import os
import struct
import sys
import wave
from collections.abc import Awaitable, Callable
from dataclasses import dataclass
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from threading import Thread
from typing import Any
from xml.etree import ElementTree
from zipfile import ZipFile

import anyio
import pytest
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import CallToolResult, ImageContent
from PIL import Image

from src.common.colors import hex_to_int
from src.models import JsonObject
from src.native import clipboard
from src.presentation import pictures
from tests.native.client import NativeClient, native_client


async def read_shape_styles(
    client: NativeClient, identity: str, names: list[str]
) -> list[JsonObject]:
    """Read every shape and text run through the public tool contract."""
    values: list[JsonObject] = []
    for name in names:
        shape = await client.call(
            "wps_ppt_shapes",
            "get",
            document=identity,
            slide_index=1,
            shape_name=name,
        )
        text = await client.call(
            "wps_ppt_text",
            "get",
            document=identity,
            slide_index=1,
            shape_name_or_index=name,
        )
        fill, line = shape["fill"], shape["line"]
        assert fill is not None and line is not None
        runs = text["runs"]
        assert runs
        font = runs[0]
        assert all(
            run[field] == font[field]
            for run in runs
            for field in ("font_size", "bold", "italic", "color_hex")
        )
        values.append(
            {
                "name": str(shape["name"]),
                "fill_visible": bool(fill["visible"]),
                "fill_color": hex_to_int(fill["color_hex"]),
                "line_visible": bool(line["visible"]),
                "line_color": hex_to_int(line["color_hex"]),
                "line_weight": float(line["weight"]),
                "text": str(text["text"]),
                "font_size": float(font["font_size"]),
                "font_bold": bool(font["bold"]),
                "font_italic": bool(font["italic"]),
                "font_color": hex_to_int(font["color_hex"]),
            }
        )
    return values


def test_native_ppt_grouped_styles_save_reopen_and_shape_export(
    tmp_path: Path,
) -> None:
    async def scenario() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="wpp"
            )
            identity = created["full_name"]
            await client.call(
                "wps_ppt_slides", "add", document=identity, layout="blank"
            )
            names: list[str] = []
            for index in range(2):
                shape = await client.call(
                    "wps_ppt_shapes",
                    "add",
                    document=identity,
                    slide_index=1,
                    shape_type="rectangle",
                    left=50.0 + index * 280.0,
                    top=60.0,
                    width=240.0,
                    height=120.0,
                    text=f"Shape {index + 1} 中文",
                    fill_type="none",
                    line_visible=False,
                    bold=False,
                    italic=False,
                )
                names.append(shape["shape_name"])
            applied = await client.call(
                "wps_ppt_style",
                "apply",
                document=identity,
                slide_index=1,
                shapes=names,
                operations=[
                    {
                        "kind": "fill",
                        "options": {"fill_type": "solid", "color": "#123456"},
                    },
                    {
                        "kind": "line",
                        "options": {
                            "visible": True,
                            "color": "#ABCDEF",
                            "weight": 2.5,
                        },
                    },
                    {
                        "kind": "text",
                        "options": {
                            "font_size": 24.0,
                            "bold": True,
                            "italic": True,
                            "color": "#112233",
                        },
                    },
                ],
            )
            assert applied["status"] == "success"
            assert applied["completed_operations"] == 6
            assert applied["failed_operations"] == 0
            assert [result["shape"] for result in applied["results"]] == names
            for result in applied["results"]:
                assert [
                    operation["kind"] for operation in result["operations"]
                ] == [
                    "fill",
                    "line",
                    "text",
                ]
                assert all(
                    operation["status"] == "success"
                    for operation in result["operations"]
                )

            expected = [
                {
                    "name": name,
                    "fill_visible": True,
                    "fill_color": 0x563412,
                    "line_visible": True,
                    "line_color": 0xEFCDAB,
                    "line_weight": 2.5,
                    "text": f"Shape {index + 1} 中文",
                    "font_size": 24.0,
                    "font_bold": True,
                    "font_italic": True,
                    "font_color": 0x332211,
                }
                for index, name in enumerate(names)
            ]
            assert await read_shape_styles(client, identity, names) == expected

            target = tmp_path / "ppt-grouped-style.pptx"
            saved = await client.call(
                "wps_documents",
                "save",
                component="wpp",
                document=identity,
                file_path=str(target),
            )
            assert (
                saved["saved"]
                and target.is_file()
                and target.stat().st_size > 0
            )
            await client.call(
                "wps_documents",
                "close",
                component="wpp",
                document=saved["full_name"],
            )
            reopened = await client.call(
                "wps_documents", "open", component="wpp", file_path=str(target)
            )
            identity = reopened["full_name"]
            assert await read_shape_styles(client, identity, names) == expected

            image_path = tmp_path / "export" / "styled-shape.png"
            exported = await client.call(
                "wps_ppt_export",
                "shape",
                document=identity,
                slide_index=1,
                shape_name_or_index=names[0],
                file_path=str(image_path),
                format="png",
            )
            assert exported["success"] and exported["file_path"] == str(
                image_path.resolve()
            )
            with Image.open(image_path) as image:
                assert (
                    image.format == "PNG"
                    and image.width > 0
                    and image.height > 0
                )
                image.verify()
            assert list(image_path.parent.iterdir()) == [image_path]

    anyio.run(scenario)


def test_native_picture_proportions_and_shape_line_breaks(
    tmp_path: Path,
) -> None:
    image_path = tmp_path / "wide.png"
    with Image.new("RGB", (200, 100), "blue") as image:
        image.save(image_path)

    async def scenario() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="wpp"
            )
            identity = created["full_name"]
            await client.call(
                "wps_ppt_slides", "add", document=identity, layout="blank"
            )
            for dimensions, expected in (
                ({"width": 140.0}, (140.0, 70.0)),
                ({"height": 40.0}, (80.0, 40.0)),
                ({"width": 60.0, "height": 90.0}, (60.0, 90.0)),
            ):
                picture = await client.call(
                    "wps_ppt_pictures",
                    "add",
                    document=identity,
                    slide_index=1,
                    file_path=str(image_path),
                    left=10.0,
                    top=10.0,
                    **dimensions,
                )
                assert (picture["width"], picture["height"]) == pytest.approx(
                    expected, abs=0.1
                )
            shape = await client.call(
                "wps_ppt_shapes",
                "add_textbox",
                document=identity,
                slide_index=1,
                left=10.0,
                top=120.0,
                width=300.0,
                height=100.0,
                text="first\r\nsecond\nthird",
            )
            content = await client.call(
                "wps_ppt_text",
                "get",
                document=identity,
                slide_index=1,
                shape_name_or_index=shape["shape_name"],
            )
            assert content["text"] == "first\rsecond\rthird"

    anyio.run(scenario)


@dataclass
class Deck:
    client: NativeClient
    identity: str
    monkeypatch: pytest.MonkeyPatch

    async def call(
        self, group: str, operation: str, **params: Any
    ) -> JsonObject:
        name = f"wps_ppt_{group}"
        prior = self.identity
        result = await self.client.call(name, operation, **params)
        if group == "document":
            if (
                operation in {"create", "open", "save_as"}
                and "full_name" in result
            ):
                self.identity = str(result["full_name"])
                self.client.documents["wpp", self.identity] = None
                if operation == "save_as":
                    self.client.documents.pop(("wpp", prior), None)
            elif operation == "close":
                self.client.documents.pop(("wpp", prior), None)
        return result

    async def shape(self, text: str = "原生测试", **params: Any) -> str:
        options = {
            "shape_type": "rectangle",
            "left": 40.0,
            "top": 50.0,
            "width": 180.0,
            "height": 90.0,
            "text": text,
        }
        result = await self.call(
            "shapes",
            "add",
            document=self.identity,
            slide_index=1,
            **options | params,
        )
        return str(result["shape_name"])

    async def text(self, name: str) -> str:
        return str(
            (
                await self.call(
                    "text",
                    "get",
                    document=self.identity,
                    slide_index=1,
                    shape_name_or_index=name,
                )
            )["text"]
        )


async def document(deck: Deck, directory: Path) -> None:
    created = await deck.call("document", "create", preset="4:3")
    assert created["slide_width"] == pytest.approx(720.0)
    await deck.call("slides", "add", document=deck.identity, layout="blank")
    await deck.call("document", "activate", document=deck.identity)
    await deck.call("document", "info", document=deck.identity)
    listed = await deck.call("document", "list")
    assert any(
        item["full_name"] == deck.identity for item in listed["presentations"]
    )
    await deck.call(
        "document",
        "set_properties",
        document=deck.identity,
        title="原生属性",
        subject="测试",
    )
    properties = await deck.call(
        "document", "properties", document=deck.identity
    )
    assert properties["properties"]["title"] == "原生属性"
    await deck.call(
        "document",
        "set_tag",
        document=deck.identity,
        target_type="presentation",
        tag_name="case",
        tag_value="native",
    )
    assert "native" in json.dumps(
        await deck.call(
            "document",
            "tags",
            document=deck.identity,
            target_type="presentation",
        )
    )
    path = directory / "document.pptx"
    await deck.call(
        "document", "save_as", document=deck.identity, file_path=str(path)
    )
    await deck.call("document", "save", document=deck.identity)
    assert path.stat().st_size > 0
    await deck.call("document", "close", document=deck.identity)
    await deck.call("document", "open", file_path=str(path))
    assert (await deck.call("document", "properties", document=deck.identity))[
        "properties"
    ]["title"] == "原生属性"
    templates = directory / "templates"
    templates.mkdir()
    # Discovery checks the template filename without opening its contents.
    template = templates / "fixture.potx"
    template.write_bytes(path.read_bytes())
    result = await deck.call(
        "document", "templates", templates_dir=str(templates)
    )
    assert result["count"] == 1


async def app(deck: Deck, directory: Path) -> None:
    del directory
    await deck.call("app", "connect")
    original = await deck.call("app", "info")
    assert original["presentations_count"] >= 1
    await deck.call("app", "window", document=deck.identity)
    shape = await deck.shape()
    try:
        await deck.call(
            "app",
            "set_view",
            document=deck.identity,
            view_type="normal",
            zoom=90,
        )
        await deck.call("app", "set_window", window_state="normal")
        await deck.call(
            "shapes",
            "select",
            document=deck.identity,
            slide_index=1,
            shape_names=[shape],
        )
        assert shape in json.dumps(
            await deck.call("app", "selection", document=deck.identity),
            ensure_ascii=False,
        )
    finally:
        state = original.get("window_state")
        if state in {"normal", "minimized", "maximized"}:
            await deck.call("app", "set_window", window_state=state)


async def slides(deck: Deck, directory: Path) -> None:
    del directory
    await deck.call(
        "slides", "add", document=deck.identity, layout="blank", count=2
    )
    await deck.call(
        "slides",
        "set_hidden",
        document=deck.identity,
        slide_index=2,
        hidden=True,
    )
    await deck.call(
        "slides",
        "set_notes",
        document=deck.identity,
        slide_index=1,
        notes_text="讲稿原生验证",
        font_size=18.0,
    )
    assert "讲稿原生验证" in json.dumps(
        await deck.call(
            "slides", "notes", document=deck.identity, slide_index=1
        ),
        ensure_ascii=False,
    )
    await deck.call("slides", "get", document=deck.identity, slide_index=1)
    before = await deck.call("slides", "list", document=deck.identity)
    await deck.call(
        "slides",
        "duplicate",
        document=deck.identity,
        slide_index=1,
        insert_at=-1,
    )
    await deck.call(
        "slides", "move", document=deck.identity, slide_index=4, new_position=2
    )
    await deck.call(
        "slides", "copy", document=deck.identity, slide_index=1, insert_at=-1
    )
    await deck.call("slides", "goto", document=deck.identity, slide_index=2)
    await deck.call(
        "slides", "delete", document=deck.identity, slide_indices=[4, 5]
    )
    after = await deck.call("slides", "list", document=deck.identity)
    assert len(after["slides"]) == len(before["slides"])


async def shapes(deck: Deck, directory: Path) -> None:
    del directory
    first = await deck.shape(
        "第一形状", shape_type="rounded_rectangle", corner_radius=0.15
    )
    second = await deck.shape("第二形状", left=260.0)
    await deck.call(
        "shapes",
        "add_textbox",
        document=deck.identity,
        slide_index=1,
        left=40.0,
        top=200.0,
        width=200.0,
        height=70.0,
        text="文本框",
    )
    line = await deck.call(
        "shapes",
        "add_line",
        document=deck.identity,
        slide_index=1,
        begin_x=40.0,
        begin_y=300.0,
        end_x=250.0,
        end_y=300.0,
    )
    await deck.call(
        "shapes",
        "lock_aspect_ratio",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=first,
        locked=True,
    )
    await deck.call(
        "shapes",
        "update",
        document=deck.identity,
        slide_index=1,
        shape_name=first,
        dleft=12.0,
        rotation=15.0,
        adjustments={"1": 0.2},
    )
    state = await deck.call(
        "shapes",
        "get",
        document=deck.identity,
        slide_index=1,
        shape_name=first,
    )
    assert state["left"] == pytest.approx(52.0)
    await deck.call(
        "shapes",
        "set_zorder",
        document=deck.identity,
        slide_index=1,
        shape_name=first,
        command="send_to_back",
    )
    copied = await deck.call(
        "shapes",
        "duplicate",
        document=deck.identity,
        slide_index=1,
        shape_name=first,
    )
    assert copied
    await deck.call("slides", "add", document=deck.identity, layout="blank")
    await deck.call(
        "shapes",
        "copy",
        document=deck.identity,
        shape_name_or_index=first,
        src_slide_index=1,
        dst_slide_index=2,
    )
    grouped = await deck.call(
        "shapes",
        "group",
        document=deck.identity,
        slide_index=1,
        shape_names=[first, second],
    )
    group = str(grouped["group_name"])
    assert (
        len(
            (
                await deck.call(
                    "shapes",
                    "group_items",
                    document=deck.identity,
                    slide_index=1,
                    shape_name_or_index=group,
                )
            )["items"]
        )
        == 2
    )
    ungrouped = await deck.call(
        "shapes",
        "ungroup",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=group,
    )
    assert ungrouped["ungrouped_count"] == 2
    await deck.call(
        "shapes",
        "delete",
        document=deck.identity,
        slide_index=1,
        shape_name=str(line["shape_name"]),
    )
    assert (
        await deck.call(
            "shapes", "list", document=deck.identity, slide_index=1
        )
    )["shapes"]


async def text(deck: Deck, directory: Path) -> None:
    shape = await deck.shape("第一段\n第二段", width=360.0, height=180.0)
    await deck.call(
        "text",
        "set",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=shape,
        text="中文😀原生\n第二段",
    )
    assert "中文😀原生" in await deck.text(shape)
    await deck.call(
        "text",
        "format",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=shape,
        font_size=22.0,
        bold=True,
    )
    await deck.call(
        "text",
        "format_range",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=shape,
        start=1,
        length=2,
        color="#112233",
    )
    await deck.call(
        "text",
        "paragraph",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=shape,
        paragraph_index=1,
        alignment="center",
        space_after=8.0,
    )
    await deck.call(
        "text",
        "bullets",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=shape,
        paragraph_index=2,
        bullet_type="unnumbered",
    )
    await deck.call(
        "text",
        "frame",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=shape,
        word_wrap=True,
        margin_left=8.0,
        auto_size="none",
    )
    await deck.call(
        "text",
        "replace",
        document=deck.identity,
        find_text="第二段",
        replace_text="替换成功",
        shape_name=shape,
    )
    assert "替换成功" in await deck.text(shape)
    await deck.call(
        "text", "check", document=deck.identity, slide_index=1, fix=False
    )
    path = directory / "text.md"
    await deck.call(
        "text", "all", document=deck.identity, output_path=str(path)
    )
    assert "替换成功" in path.read_text(encoding="utf-8")


async def annotations_case(deck: Deck, directory: Path) -> None:
    del directory
    shape = await deck.shape()
    await deck.call(
        "annotations",
        "add_comment",
        document=deck.identity,
        slide_index=1,
        text="原生批注",
        author="Native test",
        author_initials="NT",
    )
    comments = await deck.call(
        "annotations", "comments", document=deck.identity, slide_index=1
    )
    assert comments["comments"][0]["text"] == "原生批注"
    await deck.call(
        "annotations",
        "delete_comment",
        document=deck.identity,
        slide_index=1,
        comment_index=1,
    )
    assert (
        await deck.call(
            "annotations", "comments", document=deck.identity, slide_index=1
        )
    )["comments_count"] == 0
    await deck.call(
        "annotations",
        "add_link",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=shape,
        address="https://example.com/native-test",
        screen_tip="测试链接",
    )
    assert "example.com/native-test" in json.dumps(
        await deck.call(
            "annotations", "links", document=deck.identity, slide_index=1
        )
    )
    await deck.call(
        "annotations",
        "remove_link",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=shape,
    )
    assert "example.com/native-test" not in json.dumps(
        await deck.call(
            "annotations", "links", document=deck.identity, slide_index=1
        )
    )


async def tables(deck: Deck, directory: Path) -> None:
    del directory
    table = await deck.call(
        "tables", "add", document=deck.identity, slide_index=1, rows=3, cols=3
    )
    name = str(table["shape_name"])
    await deck.call(
        "tables",
        "set_data",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        data=[["姓名", "数值"], ["测试", "42"]],
        bold_first_row=True,
    )
    await deck.call(
        "tables",
        "set_cell",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        row=3,
        col=3,
        text="尾格",
        fill_color="#ABCDEF",
        font_size=18.0,
    )
    data = await deck.call(
        "tables",
        "get_data",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        include_format=True,
    )
    assert data["data"][1][1] == "42" and data["data"][2][2] == "尾格"
    await deck.call(
        "tables",
        "add_row",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        position=2,
        height=45.0,
    )
    await deck.call(
        "tables",
        "delete_row",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        position=2,
    )
    await deck.call(
        "tables",
        "add_column",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        position=2,
        width=150.0,
    )
    await deck.call(
        "tables",
        "delete_column",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        position=2,
    )
    await deck.call(
        "tables",
        "set_layout",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        row_heights=[60.0, 60.0, 60.0],
        col_widths=[150.0, 150.0, 150.0],
    )
    await deck.call(
        "tables",
        "set_style",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        first_row=True,
        banding_rows=True,
    )
    await deck.call(
        "tables",
        "set_borders",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        sides=["top", "bottom"],
        color="#112233",
        weight=1.5,
        visible=True,
    )
    await deck.call(
        "tables",
        "merge",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        start_row=3,
        start_col=1,
        end_row=3,
        end_col=2,
    )
    await deck.call(
        "tables",
        "split",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        row=3,
        col=1,
        num_cols=2,
    )
    assert (
        await deck.call(
            "tables",
            "get_data",
            document=deck.identity,
            slide_index=1,
            shape_name_or_index=name,
        )
    )["data"][1][1] == "42"


async def layout(deck: Deck, directory: Path) -> None:
    del directory
    names = [
        await deck.shape(
            str(index), left=40.0 + index * 220.0, top=50.0 + index * 20.0
        )
        for index in range(3)
    ]
    await deck.call(
        "layout",
        "align",
        document=deck.identity,
        slide_index=1,
        shape_names=names,
        align_to="top",
    )
    positions = [
        (
            await deck.call(
                "shapes",
                "get",
                document=deck.identity,
                slide_index=1,
                shape_name=name,
            )
        )["top"]
        for name in names
    ]
    assert len(set(positions)) == 1
    await deck.call(
        "layout",
        "distribute",
        document=deck.identity,
        slide_index=1,
        shape_names=names,
        direction="horizontal",
    )
    await deck.call(
        "layout", "set_size", document=deck.identity, preset="16:9"
    )
    size = await deck.call("layout", "get_size", document=deck.identity)
    assert size["width_points"] == pytest.approx(960.0)
    await deck.call(
        "layout",
        "background",
        document=deck.identity,
        slide_index=1,
        fill_type="solid",
        color="#F0F0F0",
    )
    await deck.call(
        "layout",
        "flip",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=names[0],
        direction="horizontal",
    )
    pair = [
        await deck.shape("合并", left=80.0),
        await deck.shape("合并", left=100.0),
    ]
    await deck.call(
        "layout",
        "merge",
        document=deck.identity,
        slide_index=1,
        shape_names=pair,
        merge_type="union",
    )


async def connectors(deck: Deck, directory: Path) -> None:
    del directory
    start = await deck.shape("起点")
    end = await deck.shape("终点", left=350.0)
    if sys.platform == "darwin":
        before = await deck.call(
            "shapes", "list", document=deck.identity, slide_index=1
        )
        with pytest.raises(ToolError, match="ConnectionSites"):
            await deck.call(
                "connectors",
                "add",
                document=deck.identity,
                slide_index=1,
                begin_shape=start,
                begin_site="right",
                end_shape=end,
                end_site="left",
            )
        assert (
            await deck.call(
                "shapes", "list", document=deck.identity, slide_index=1
            )
            == before
        )
    result = await deck.call(
        "connectors",
        "add",
        document=deck.identity,
        slide_index=1,
        begin_shape=start,
        begin_site=1 if sys.platform == "darwin" else "right",
        end_shape=end,
        end_site=1 if sys.platform == "darwin" else "left",
    )
    name = str(result["shape_name"])
    await deck.call(
        "connectors",
        "format",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        color="#123456",
        weight=2.0,
        end_arrow="triangle",
    )
    assert (
        await deck.call(
            "shapes",
            "get",
            document=deck.identity,
            slide_index=1,
            shape_name=name,
        )
    )["name"] == name


async def freeform(deck: Deck, directory: Path) -> None:
    result = await deck.call(
        "freeform",
        "build",
        document=deck.identity,
        slide_index=1,
        start_x=40.0,
        start_y=40.0,
        nodes=[
            {"segment_type": "line", "x1": 220.0, "y1": 40.0},
            {"segment_type": "line", "x1": 220.0, "y1": 180.0},
            {"segment_type": "line", "x1": 40.0, "y1": 180.0},
        ],
        close_path=True,
        shape_name="native-freeform",
    )
    name = str(result["shape_name"])
    if sys.platform == "darwin":
        assert name == "native-freeform"

        async def saved_paths(filename: str) -> list[bytes]:
            path = directory / filename
            await deck.call(
                "document",
                "save_as",
                document=deck.identity,
                file_path=str(path),
            )
            with ZipFile(path) as archive:
                slide = ElementTree.fromstring(
                    archive.read("ppt/slides/slide1.xml")
                )
            return [
                ElementTree.tostring(item)
                for item in slide.findall(
                    ".//{http://schemas.openxmlformats.org/"
                    "drawingml/2006/main}pathLst"
                )
            ]

        before = await saved_paths("freeform-before.pptx")
        assert before and b"lnTo" in b"".join(before)
        shapes_before = await deck.call(
            "shapes", "list", document=deck.identity, slide_index=1
        )
        options: dict[str, Any]
        for operation, options in (
            ("nodes", {}),
            ("move_node", {"node_index": 2, "x": 240.0, "y": 40.0}),
            ("insert_node", {"after_index": 2, "x1": 240.0, "y1": 100.0}),
            ("delete_node", {"node_index": 3}),
            ("set_editing", {"node_index": 2, "editing_type": "corner"}),
            ("set_segment", {"node_index": 3, "segment_type": "curve"}),
        ):
            with pytest.raises(ToolError, match="macOS WPS.*freeform"):
                await deck.call(
                    "freeform",
                    operation,
                    document=deck.identity,
                    slide_index=1,
                    shape_name=name,
                    **options,
                )
        with pytest.raises(ToolError, match="macOS WPS.*freeform"):
            await deck.call(
                "freeform",
                "build",
                document=deck.identity,
                slide_index=1,
                start_x=40.0,
                start_y=40.0,
                nodes=[{"segment_type": "curve", "x1": 220.0, "y1": 180.0}],
            )
        assert (
            await deck.call(
                "shapes", "list", document=deck.identity, slide_index=1
            )
            == shapes_before
        )
        assert await saved_paths("freeform-after.pptx") == before
        return
    original = await deck.call(
        "freeform",
        "nodes",
        document=deck.identity,
        slide_index=1,
        shape_name=name,
    )
    assert original["node_count"] >= 4
    moved = await deck.call(
        "freeform",
        "move_node",
        document=deck.identity,
        slide_index=1,
        shape_name=name,
        node_index=2,
        x=240.0,
        y=40.0,
    )
    assert moved["x"] == pytest.approx(240.0)
    await deck.call(
        "freeform",
        "insert_node",
        document=deck.identity,
        slide_index=1,
        shape_name=name,
        after_index=2,
        x1=240.0,
        y1=100.0,
    )
    await deck.call(
        "freeform",
        "set_editing",
        document=deck.identity,
        slide_index=1,
        shape_name=name,
        node_index=2,
        editing_type="corner",
    )
    await deck.call(
        "freeform",
        "set_segment",
        document=deck.identity,
        slide_index=1,
        shape_name=name,
        node_index=3,
        segment_type="line",
    )
    await deck.call(
        "freeform",
        "delete_node",
        document=deck.identity,
        slide_index=1,
        shape_name=name,
        node_index=3,
    )
    assert (
        await deck.call(
            "freeform",
            "nodes",
            document=deck.identity,
            slide_index=1,
            shape_name=name,
        )
    )["node_count"] == original["node_count"]


async def sections(deck: Deck, directory: Path) -> None:
    del directory
    await deck.call(
        "slides", "add", document=deck.identity, count=2, layout="blank"
    )
    first = await deck.call(
        "sections", "add", document=deck.identity, name="第一节", slide_index=1
    )
    await deck.call(
        "sections", "add", document=deck.identity, name="第二节", slide_index=3
    )
    await deck.call(
        "sections",
        "manage",
        document=deck.identity,
        section_index=first["section_index"],
        action="rename",
        new_name="原生节",
    )
    state = await deck.call("sections", "list", document=deck.identity)
    assert any(item["name"] == "原生节" for item in state["sections"])
    index = next(
        item["index"] for item in state["sections"] if item["name"] == "第二节"
    )
    await deck.call(
        "sections",
        "manage",
        document=deck.identity,
        section_index=index,
        action="move",
        move_to_index=1,
    )
    before = await deck.call("slides", "list", document=deck.identity)
    await deck.call(
        "sections",
        "manage",
        document=deck.identity,
        section_index=1,
        action="delete",
    )
    assert await deck.call("slides", "list", document=deck.identity) == before
    assert (await deck.call("sections", "list", document=deck.identity))[
        "slides_total"
    ] == 3


async def placeholders(deck: Deck, directory: Path) -> None:
    del directory
    await deck.call(
        "slides", "add", document=deck.identity, layout="title", position=1
    )
    listed = await deck.call(
        "placeholders", "list", document=deck.identity, slide_index=1
    )
    assert listed["placeholders"]
    index = listed["placeholders"][0]["index"]
    await deck.call(
        "placeholders",
        "set_text",
        document=deck.identity,
        slide_index=1,
        placeholder_index=index,
        text="占位符原生验证",
    )
    assert "占位符原生验证" in json.dumps(
        await deck.call(
            "placeholders",
            "get",
            document=deck.identity,
            slide_index=1,
            placeholder_index=index,
        ),
        ensure_ascii=False,
    )
    await deck.call(
        "placeholders", "designs", document=deck.identity, include_layouts=True
    )
    layouts = await deck.call(
        "placeholders", "layouts", document=deck.identity, design_index=1
    )
    assert layouts["layouts"]
    await deck.call(
        "placeholders", "master", document=deck.identity, design_index=1
    )


async def style(deck: Deck, directory: Path) -> None:
    del directory
    first = await deck.shape("样式来源", fill_color="#123456", font_size=24.0)
    second = await deck.shape("样式目标", left=300.0)
    await deck.call(
        "style",
        "copy",
        document=deck.identity,
        slide_index=1,
        source_shape=first,
        target_shapes=[second],
    )
    assert (
        await deck.call(
            "shapes",
            "get",
            document=deck.identity,
            slide_index=1,
            shape_name=second,
        )
    )["fill"]["color_hex"].upper() == "#123456"
    await deck.call(
        "style",
        "set_default",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=first,
    )
    outcome = await deck.call(
        "style",
        "apply",
        document=deck.identity,
        slide_index=1,
        shapes=[first, second],
        operations=[
            {"kind": "glow", "options": {"radius": 3.0, "color": "#2266AA"}},
            {"kind": "soft_edge", "options": {"radius": 1.0}},
        ],
    )
    assert outcome["failed_operations"] == 0


async def themes(deck: Deck, directory: Path) -> None:
    del directory
    name = await deck.shape("字体原生验证", font_name="Arial")
    await deck.call(
        "themes",
        "replace_font",
        document=deck.identity,
        original_font="Arial",
        replacement_font="Times New Roman",
    )
    assert (
        "Times New Roman"
        in (await deck.call("themes", "fonts", document=deck.identity))[
            "fonts"
        ]
    )
    await deck.call(
        "themes",
        "set_fonts",
        document=deck.identity,
        latin="Arial",
        east_asian="Arial",
        apply_to_existing=True,
    )
    before = await deck.call("themes", "colors", document=deck.identity)
    await deck.call(
        "themes", "set_colors", document=deck.identity, accent1="#2266AA"
    )
    colors = await deck.call("themes", "colors", document=deck.identity)
    assert (
        next(
            item["color_hex"]
            for item in colors["colors"]
            if item["name"] == "accent1"
        ).upper()
        == "#2266AA"
    )
    assert [
        item for item in colors["colors"] if item["name"] != "accent1"
    ] == [item for item in before["colors"] if item["name"] != "accent1"]
    await deck.call(
        "themes",
        "headers_footers",
        document=deck.identity,
        footer_visible=True,
        footer_text="原生页脚",
        slide_number_visible=True,
    )
    assert "字体原生验证" in await deck.text(name)


async def animation(deck: Deck, directory: Path) -> None:
    del directory
    first = await deck.shape("动画来源")
    second = await deck.shape("动画目标", left=300.0)
    added = await deck.call(
        "animation",
        "add",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=first,
        effect="fade",
        duration=0.5,
    )
    await deck.call(
        "animation",
        "transition",
        document=deck.identity,
        slide_index=1,
        effect="fade",
        duration=0.5,
        advance_on_click=True,
    )
    await deck.call(
        "animation",
        "update",
        document=deck.identity,
        slide_index=1,
        animation_index=added["animation_index"],
        delay=0.25,
    )
    listed = await deck.call(
        "animation", "list", document=deck.identity, slide_index=1
    )
    assert listed["main_sequence_count"] == 1
    await deck.call(
        "animation",
        "copy",
        document=deck.identity,
        slide_index=1,
        source_shape=first,
        target_shape=second,
    )
    assert (
        await deck.call(
            "animation", "list", document=deck.identity, slide_index=1
        )
    )["main_sequence_count"] == 2
    await deck.call(
        "animation",
        "remove",
        document=deck.identity,
        slide_index=1,
        animation_index=1,
    )
    await deck.call(
        "animation",
        "clear",
        document=deck.identity,
        slide_index=1,
        clear_transitions=True,
    )
    assert (
        await deck.call(
            "animation", "list", document=deck.identity, slide_index=1
        )
    )["main_sequence_count"] == 0


async def history(deck: Deck, directory: Path) -> None:
    del directory
    name = await deck.shape("撤销前")
    await deck.call("history", "begin", document=deck.identity)
    await deck.call(
        "text",
        "set",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        text="撤销后",
    )
    assert await deck.text(name) == "撤销后"
    assert (await deck.call("history", "undo", document=deck.identity))[
        "actions_undone"
    ] == 1
    assert await deck.text(name) == "撤销前"
    assert (await deck.call("history", "redo", document=deck.identity))[
        "actions_redone"
    ] == 1
    assert await deck.text(name) == "撤销后"


async def slideshow(deck: Deck, directory: Path) -> None:
    del directory
    if sys.platform == "darwin":
        before = await deck.call("slides", "list", document=deck.identity)
        options: dict[str, Any]
        for operation, options in (
            (
                "start",
                {"show_type": "window", "start_slide": 1, "end_slide": 1},
            ),
            ("status", {}),
            ("next", {}),
            ("previous", {}),
            ("goto", {"slide_index": 1}),
            ("stop", {}),
        ):
            with pytest.raises(
                ToolError, match="macOS WPS slide show.*unsafe"
            ):
                await deck.call(
                    "slideshow", operation, document=deck.identity, **options
                )
            assert (
                await deck.call("slides", "list", document=deck.identity)
                == before
            )
        return
    await deck.call(
        "slides", "add", document=deck.identity, layout="blank", count=2
    )
    try:
        await deck.call(
            "slideshow",
            "start",
            document=deck.identity,
            show_type="window",
            start_slide=1,
            end_slide=3,
        )
        assert (
            await deck.call("slideshow", "status", document=deck.identity)
        )["running"]
        assert (await deck.call("slideshow", "next", document=deck.identity))[
            "current_slide"
        ] == 2
        assert (
            await deck.call("slideshow", "previous", document=deck.identity)
        )["current_slide"] == 1
        assert (
            await deck.call(
                "slideshow", "goto", document=deck.identity, slide_index=3
            )
        )["current_slide"] == 3
    finally:
        await deck.call("slideshow", "stop", document=deck.identity)
    assert not (
        await deck.call("slideshow", "status", document=deck.identity)
    )["running"]


async def charts(deck: Deck, directory: Path) -> None:
    del directory
    added = await deck.call(
        "charts",
        "add",
        document=deck.identity,
        slide_index=1,
        chart_type="column",
    )
    name = str(added["shape_name"])
    initial = await deck.call(
        "charts",
        "get_data",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
    )
    assert initial["categories"] and initial["series"]
    categories = [
        f"分类{index + 1}" for index in range(len(initial["categories"]))
    ]
    series = [
        {
            "name": f"系列{index + 1}",
            "values": [
                float(10 * index + item + 1) for item in range(len(categories))
            ],
        }
        for index in range(len(initial["series"]))
    ]
    if sys.platform in {"darwin", "linux"}:
        assert all(
            item["name"] is None or isinstance(item["name"], str)
            for item in initial["series"]
        )
        assert all(
            len(item["values"]) == len(initial["categories"])
            and all(
                isinstance(value, (int, float)) for value in item["values"]
            )
            for item in initial["series"]
        )
        message = (
            r"ChartData\.Workbook"
            if sys.platform == "darwin"
            else "Linux WPS RPC cannot safely edit embedded chart workbooks"
        )
        with pytest.raises(ToolError, match=message):
            await deck.call(
                "charts",
                "set_data",
                document=deck.identity,
                slide_index=1,
                shape_name_or_index=name,
                categories=categories,
                series=series,
            )
        assert (
            await deck.call(
                "charts",
                "get_data",
                document=deck.identity,
                slide_index=1,
                shape_name_or_index=name,
            )
            == initial
        )
        series = initial["series"]
        categories = initial["categories"]
    else:
        await deck.call(
            "charts",
            "set_data",
            document=deck.identity,
            slide_index=1,
            shape_name_or_index=name,
            categories=categories,
            series=series,
        )
        actual = await deck.call(
            "charts",
            "get_data",
            document=deck.identity,
            slide_index=1,
            shape_name_or_index=name,
        )
        assert (
            actual["categories"] == categories and actual["series"] == series
        )
    await deck.call(
        "charts",
        "format",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        title="原生图表",
        has_legend=True,
        legend_position="bottom",
    )
    await deck.call(
        "charts",
        "format_axis",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        axis="value",
        min_scale=0.0,
        max_scale=100.0,
        major_unit=20.0,
    )
    await deck.call(
        "charts",
        "set_series",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        series_index=1,
        color="#2266AA",
        show_data_labels=True,
    )
    await deck.call(
        "charts",
        "set_type",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        chart_type="line",
    )
    final = await deck.call(
        "charts",
        "get_data",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
    )
    assert [item["values"] for item in final["series"]] == [
        item["values"] for item in series
    ]
    if sys.platform == "darwin":
        # WPS can initialize default labels when its data editor becomes ready.
        assert len(final["categories"]) == len(initial["categories"])
        assert all(isinstance(value, str) for value in final["categories"])
    else:
        assert final["categories"] == categories
    for before, after in zip(series, final["series"], strict=True):
        if before["name"] is not None:
            assert after["name"] == before["name"]


async def smartart(deck: Deck, directory: Path) -> None:
    del directory
    choices = await deck.call("smartart", "options", list_type="layouts")
    assert choices["layouts"]
    index = choices["layouts"][0]["index"]
    if sys.platform == "darwin":
        before = await deck.call(
            "shapes", "list", document=deck.identity, slide_index=1
        )
        options: dict[str, Any]
        for options in (
            {"node_texts": ["起点", "终点"]},
            {"font_name": "Arial"},
            {"font_size": 24.0},
            {"bold": False},
            {"font_color": "#2266AA"},
        ):
            with pytest.raises(ToolError, match="SmartArt nodes"):
                await deck.call(
                    "smartart",
                    "add",
                    document=deck.identity,
                    slide_index=1,
                    layout_index=index,
                    **options,
                )
            assert (
                await deck.call(
                    "shapes", "list", document=deck.identity, slide_index=1
                )
                == before
            )
    result = await deck.call(
        "smartart",
        "add",
        document=deck.identity,
        slide_index=1,
        layout_index=index,
        node_texts=None if sys.platform == "darwin" else ["起点", "终点"],
    )
    name = str(result["shape_name"])
    if sys.platform == "darwin":
        assert result["node_count"] is None
        before = await deck.call(
            "shapes", "list", document=deck.identity, slide_index=1
        )
        for action, options in (
            ("set_text", {"node_index": 1, "text": "原生节点"}),
            ("add_node", {"text": "新节点"}),
            ("delete_node", {"node_index": 1}),
            ("format_node", {"node_index": 1, "font_size": 24.0}),
            ("format_all_nodes", {"font_color": "#2266AA"}),
        ):
            with pytest.raises(ToolError, match="SmartArt nodes"):
                await deck.call(
                    "smartart",
                    "modify",
                    document=deck.identity,
                    slide_index=1,
                    shape_name_or_index=name,
                    action=action,
                    **options,
                )
            assert (
                await deck.call(
                    "shapes", "list", document=deck.identity, slide_index=1
                )
                == before
            )
    else:
        await deck.call(
            "smartart",
            "modify",
            document=deck.identity,
            slide_index=1,
            shape_name_or_index=name,
            action="set_text",
            node_index=1,
            text="原生节点",
        )
        modified = await deck.call(
            "smartart",
            "modify",
            document=deck.identity,
            slide_index=1,
            shape_name_or_index=name,
            action="add_node",
            text="新节点",
        )
        count = modified["total_nodes"]
        deleted = await deck.call(
            "smartart",
            "modify",
            document=deck.identity,
            slide_index=1,
            shape_name_or_index=name,
            action="delete_node",
            node_index=count,
        )
        assert deleted["total_nodes"] == count - 1
    for action, field, list_type in (
        ("change_layout", "layout_index", "layouts"),
        ("change_color", "color_index", "colors"),
        ("change_style", "style_index", "styles"),
    ):
        options = await deck.call("smartart", "options", list_type=list_type)
        assert options[list_type]
        choice = options[list_type][-1]["index"]
        changed = await deck.call(
            "smartart",
            "modify",
            document=deck.identity,
            slide_index=1,
            shape_name_or_index=name,
            action=action,
            **{field: choice},
        )
        assert changed["action"] == action
        if sys.platform == "darwin":
            assert changed["total_nodes"] is None


async def media(deck: Deck, directory: Path) -> None:
    audio = directory / "tone.wav"
    with wave.open(str(audio), "wb") as stream:
        stream.setnchannels(1)
        stream.setsampwidth(2)
        stream.setframerate(8000)
        stream.writeframes(
            b"".join(
                struct.pack(
                    "<h", int(1500 * math.sin(index * math.tau * 440 / 8000))
                )
                for index in range(8000)
            )
        )
    result = await deck.call(
        "media",
        "add",
        document=deck.identity,
        slide_index=1,
        kind="audio",
        file_path=str(audio),
    )
    name = str(result["shape_name"])
    await deck.call(
        "media",
        "configure",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        volume=0.5,
        muted=True,
        loop=False,
    )
    assert (
        await deck.call(
            "shapes",
            "get",
            document=deck.identity,
            slide_index=1,
            shape_name=name,
        )
    )["name"] == name


async def pictures_case(deck: Deck, directory: Path) -> None:
    image = directory / "fixture.png"
    with Image.new("RGB", (200, 100), "blue") as generated:
        generated.save(image)
    result = await deck.call(
        "pictures",
        "add",
        document=deck.identity,
        slide_index=1,
        file_path=str(image),
        left=30.0,
        top=30.0,
        width=200.0,
    )
    name = str(result["shape_name"])
    await deck.call(
        "pictures",
        "crop",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        crop_left=10.0,
    )
    formatted = await deck.call(
        "pictures",
        "format",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=name,
        brightness=0.4,
        contrast=0.6,
    )
    assert formatted["brightness"] == pytest.approx(0.4, abs=0.01)
    icons = directory / "outlined"
    icons.mkdir()
    (icons / "native.svg").write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">'
        '<rect width="24" height="24" fill="currentColor"/></svg>',
        encoding="utf-8",
    )
    (directory / "metadata.json").write_text(
        json.dumps(
            {
                "icons": [
                    {
                        "name": "native",
                        "tags": ["native"],
                        "categories": ["test"],
                        "popularity": 1,
                    }
                ]
            }
        ),
        encoding="utf-8",
    )
    handler = partial(SimpleHTTPRequestHandler, directory=str(directory))
    host = ThreadingHTTPServer(("127.0.0.1", 0), handler)
    worker = Thread(target=host.serve_forever, daemon=True)
    worker.start()
    base = f"http://127.0.0.1:{host.server_port}"
    with deck.monkeypatch.context() as patch:
        patch.setenv("no_proxy", "127.0.0.1,localhost")
        patch.setattr(pictures, "ICON_PACKAGE_BASE", base)
        patch.setattr(pictures, "_ICON_METADATA_URL", base + "/metadata.json")
        patch.setattr(pictures, "_icon_cache", None)
        patch.setattr(pictures, "_icon_on_cdn_cache", {})
        try:
            await deck.call(
                "pictures",
                "add_url",
                document=deck.identity,
                slide_index=1,
                url=base + "/fixture.png",
                width=100.0,
            )
            found = await deck.call("pictures", "search_icons", query="native")
            assert found["icons"][0]["name"] == "native"
            await deck.call(
                "pictures",
                "add_icon",
                document=deck.identity,
                slide_index=1,
                icon_name="native",
                color="#2266AA",
            )
        finally:
            host.shutdown()
            host.server_close()
            worker.join(timeout=2)


async def export(deck: Deck, directory: Path) -> None:
    shape = await deck.shape("导出验证")
    png = directory / "shape.png"
    await deck.call(
        "export",
        "shape",
        document=deck.identity,
        slide_index=1,
        shape_name_or_index=shape,
        file_path=str(png),
    )
    pdf = directory / "deck.pdf"
    await deck.call(
        "export", "pdf", document=deck.identity, file_path=str(pdf)
    )
    assert pdf.read_bytes().startswith(b"%PDF-")
    images = await deck.call(
        "export",
        "images",
        document=deck.identity,
        output_dir=str(directory / "slides"),
        width=640,
    )
    assert len(images["files"]) == 1
    with Image.open(images["files"][0]) as rendered:
        assert rendered.width == 640
    response = await deck.client.server.call_tool(
        "wps_ppt_export",
        {
            "params": {
                "operation": "preview",
                "document": deck.identity,
                "slide_index": 1,
            }
        },
    )
    assert isinstance(response, CallToolResult) and not response.is_error
    content = response.content[0]
    assert isinstance(content, ImageContent)
    assert base64.b64decode(content.data).startswith(b"\x89PNG\r\n\x1a\n")


Scenario = Callable[[Deck, Path], Awaitable[None]]


@pytest.mark.parametrize(
    "scenario",
    [
        document,
        app,
        slides,
        shapes,
        text,
        annotations_case,
        tables,
        layout,
        connectors,
        freeform,
        sections,
        placeholders,
        style,
        themes,
        animation,
        history,
        slideshow,
        charts,
        smartart,
        media,
        pictures_case,
        export,
    ],
    ids=lambda scenario: scenario.__name__,
)
def test_native_presentation_operations(
    scenario: Scenario,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    async def run() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="wpp"
            )
            deck = Deck(client, str(created["full_name"]), monkeypatch)
            await deck.call(
                "slides", "add", document=deck.identity, layout="blank"
            )
            await scenario(deck, tmp_path)

    anyio.run(run)


@pytest.mark.skipif(
    sys.platform != "darwin" and os.environ.get("WPS_NATIVE_CLIPBOARD") != "1",
    reason="Set WPS_NATIVE_CLIPBOARD=1 to replace the desktop clipboard",
)
@pytest.mark.parametrize("count", [1, 2])
def test_native_presentation_clipboard(
    monkeypatch: pytest.MonkeyPatch, count: int
) -> None:
    if sys.platform == "darwin":
        script = clipboard.MACOS_SCRIPT.replace(
            "$.NSPasteboard.generalPasteboard",
            "$.NSPasteboard.pasteboardWithUniqueName",
        )
        # Validate the real AppKit payload on an isolated, disposable board.
        script = (
            script.rsplit("}", 1)[0]
            + """
    try {
        var saved = pasteboard.pasteboardItems;
        if (Number(saved.count) !== paths.length)
            throw new Error('Unexpected pasteboard item count');
        for (var i = 0; i < paths.length; i++) {
            var item = saved.objectAtIndex(i), actual, expected;
            if (paths.length === 1) {
                actual = ObjC.unwrap(item.dataForType($.NSPasteboardTypePNG)
                    .base64EncodedStringWithOptions(0));
                var data = $.NSData.dataWithContentsOfFile($(paths[i]));
                expected = ObjC.unwrap(data.base64EncodedStringWithOptions(0));
            } else {
                actual = ObjC.unwrap(
                    item.stringForType($.NSPasteboardTypeFileURL));
                var url = $.NSURL.fileURLWithPath($(paths[i]));
                expected = ObjC.unwrap(url.absoluteString);
            }
            if (actual !== expected)
                throw new Error('Pasteboard payload differs');
        }
    } finally { pasteboard.releaseGlobally; }
}
"""
        )
        assert "generalPasteboard" not in script
        monkeypatch.setattr(clipboard, "MACOS_SCRIPT", script)

    async def run() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="wpp"
            )
            deck = Deck(client, str(created["full_name"]), monkeypatch)
            await deck.call(
                "slides",
                "add",
                document=deck.identity,
                layout="blank",
                count=count,
            )
            await deck.shape("剪贴板验证")
            result = await deck.call(
                "export",
                "clipboard",
                document=deck.identity,
                slide_indices=list(range(1, count + 1)),
            )
            assert result["count"] == count and result["method"]
            if sys.platform == "darwin":
                assert result["method"] == (
                    "public.png" if count == 1 else "public.file-url"
                )

    anyio.run(run)


def test_native_presentation_theme_file(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    async def run() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="wpp"
            )
            deck = Deck(client, str(created["full_name"]), monkeypatch)
            await deck.call(
                "slides", "add", document=deck.identity, layout="blank"
            )
            path = tmp_path / "theme-source.pptx"
            await deck.call(
                "document",
                "save_as",
                document=deck.identity,
                file_path=str(path),
            )
            # ApplyTheme also accepts a presentation containing a theme.
            source = Path(os.environ.get("WPS_NATIVE_THEME", str(path)))
            await deck.call("document", "create")
            await deck.call(
                "slides", "add", document=deck.identity, layout="blank"
            )
            await deck.call(
                "themes",
                "apply",
                document=deck.identity,
                theme_path=str(source.resolve(strict=True)),
            )
            assert (
                len(
                    (
                        await deck.call(
                            "themes", "colors", document=deck.identity
                        )
                    )["colors"]
                )
                == 12
            )

    anyio.run(run)
