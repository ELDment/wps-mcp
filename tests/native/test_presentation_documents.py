"""Exercise presentation tool groups using only test-owned native decks."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from tests.native.support import Deck, Scenario, run_presentation


async def document(deck: Deck, directory: Path) -> None:
    created = await deck.document("create", preset="4:3")
    assert created["slide_width"] == pytest.approx(720.0)
    await deck.call("slides", "add", document=deck.identity, layout="blank")
    await deck.call("document", "activate", document=deck.identity)
    await deck.document("info", document=deck.identity)
    listed = await deck.document("list")
    assert any(
        item["full_name"] == deck.identity for item in listed["documents"]
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
    await deck.document("save", document=deck.identity, file_path=str(path))
    await deck.document("save", document=deck.identity)
    assert path.stat().st_size > 0
    await deck.document("close", document=deck.identity)
    await deck.document("open", file_path=str(path))
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
    await deck.client.call("wps_components", "connect", component="wpp")
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


@pytest.mark.parametrize(
    "scenario",
    [document, app, slides, sections],
    ids=lambda scenario: scenario.__name__,
)
def test_native_presentation_documents(
    scenario: Scenario, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    run_presentation(scenario, tmp_path, monkeypatch)
