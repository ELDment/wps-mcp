"""Exercise presentation tool groups using only test-owned native decks."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from tests.native.support import Deck, Scenario, run_presentation


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


@pytest.mark.parametrize(
    "scenario",
    [text, annotations_case, tables],
    ids=lambda scenario: scenario.__name__,
)
def test_native_presentation_text(
    scenario: Scenario, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    run_presentation(scenario, tmp_path, monkeypatch)
