"""Exercise presentation tool groups using only test-owned native decks."""

from __future__ import annotations

import json
import math
import struct
import wave
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from threading import Thread

import anyio
import pytest
from PIL import Image

from tests.native.client import native_client
from tests.native.support import Deck, Scenario, run_presentation
from wps_mcp.presentation import assets, pictures


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
        patch.setattr(assets, "ICON_PACKAGE_BASE", base)
        patch.setattr(assets, "_ICON_METADATA_URL", base + "/metadata.json")
        patch.setattr(assets, "_icon_cache", None)
        patch.setattr(assets, "_icon_on_cdn_cache", {})
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


@pytest.mark.parametrize(
    "scenario",
    [media, pictures_case],
    ids=lambda scenario: scenario.__name__,
)
def test_native_presentation_assets(
    scenario: Scenario, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    run_presentation(scenario, tmp_path, monkeypatch)
