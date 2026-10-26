"""Exercise presentation tool groups using only test-owned native decks."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import pytest
from mcp.server.mcpserver.exceptions import ToolError

from tests.native.support import Deck, Scenario, run_presentation


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


@pytest.mark.parametrize(
    "scenario",
    [animation, history, slideshow],
    ids=lambda scenario: scenario.__name__,
)
def test_native_presentation_playback(
    scenario: Scenario, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    run_presentation(scenario, tmp_path, monkeypatch)
