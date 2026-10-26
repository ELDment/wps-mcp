"""Exercise presentation tool groups using only test-owned native decks."""

from __future__ import annotations

import os
import sys
from pathlib import Path

import anyio
import pytest
from PIL import Image

from tests.native.client import native_client
from tests.native.support import Deck, Scenario, run_presentation
from wps_mcp.native import clipboard


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
    await deck.document(
        "export_pdf", document=deck.identity, file_path=str(pdf)
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
    preview = await deck.client.image(
        "wps_ppt_export",
        "preview",
        document=deck.identity,
        slide_index=1,
    )
    assert preview.startswith(b"\x89PNG\r\n\x1a\n")


@pytest.mark.parametrize(
    "scenario",
    [export],
    ids=lambda scenario: scenario.__name__,
)
def test_native_presentation_export(
    scenario: Scenario, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    run_presentation(scenario, tmp_path, monkeypatch)
