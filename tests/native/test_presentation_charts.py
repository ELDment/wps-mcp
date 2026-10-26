"""Exercise presentation tool groups using only test-owned native decks."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import pytest
from mcp.server.mcpserver.exceptions import ToolError

from tests.native.support import Deck, Scenario, run_presentation


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


@pytest.mark.parametrize(
    "scenario",
    [charts, smartart],
    ids=lambda scenario: scenario.__name__,
)
def test_native_presentation_charts(
    scenario: Scenario, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    run_presentation(scenario, tmp_path, monkeypatch)
