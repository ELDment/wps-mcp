"""Exercise presentation tool groups using only test-owned native decks."""

from __future__ import annotations

import json
import os
from pathlib import Path

import anyio
import pytest
from PIL import Image

from tests.native.client import NativeClient, native_client
from tests.native.support import Deck, Scenario, run_presentation
from wps_mcp.common.colors import hex_to_int
from wps_mcp.models import JsonObject


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
                    operation["success"] is True
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
            identity = await client.save_and_reopen("wpp", identity, target)
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
            await deck.document(
                "save", document=deck.identity, file_path=str(path)
            )
            # ApplyTheme also accepts a presentation containing a theme.
            source = Path(os.environ.get("WPS_NATIVE_THEME", str(path)))
            await deck.document("create")
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


@pytest.mark.parametrize(
    "scenario",
    [layout, placeholders, style, themes],
    ids=lambda scenario: scenario.__name__,
)
def test_native_presentation_design(
    scenario: Scenario, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    run_presentation(scenario, tmp_path, monkeypatch)
