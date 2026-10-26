"""Exercise presentation tool groups using only test-owned native decks."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any
from xml.etree import ElementTree
from zipfile import ZipFile

import pytest
from mcp.server.mcpserver.exceptions import ToolError

from tests.native.support import Deck, Scenario, run_presentation


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
            await deck.document(
                "save", document=deck.identity, file_path=str(path)
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


@pytest.mark.parametrize(
    "scenario",
    [shapes, connectors, freeform],
    ids=lambda scenario: scenario.__name__,
)
def test_native_presentation_shapes(
    scenario: Scenario, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    run_presentation(scenario, tmp_path, monkeypatch)
