"""Opt-in presentation styling and persistence on a test-owned deck only."""

from pathlib import Path

import anyio
import pytest
from PIL import Image

from src.models import JsonObject
from tests.native.client import NativeClient, native_client


async def read_shape_styles(
    client: NativeClient, identity: str, names: list[str]
) -> list[JsonObject]:
    """Read native values on the WPP STA and return only plain data."""
    assert client.server.runtime is not None
    session = client.server.runtime.sessions["wpp"]

    def read_shapes() -> list[JsonObject]:
        shapes = session.get_document(identity).Slides(1).Shapes
        values: list[JsonObject] = []
        for name in names:
            shape = shapes(name)
            text = shape.TextFrame.TextRange
            font = text.Font
            values.append(
                {
                    "name": str(shape.Name),
                    "fill_visible": bool(shape.Fill.Visible),
                    "fill_color": int(shape.Fill.ForeColor.RGB),
                    "line_visible": bool(shape.Line.Visible),
                    "line_color": int(shape.Line.ForeColor.RGB),
                    "line_weight": float(shape.Line.Weight),
                    "text": str(text.Text),
                    "font_size": float(font.Size),
                    "font_bold": bool(font.Bold),
                    "font_italic": bool(font.Italic),
                    "font_color": int(font.Color.RGB),
                }
            )
        return values

    return await anyio.to_thread.run_sync(session.execute, read_shapes)


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
