"""Verify grouped presentation contracts and resource publication."""

from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import anyio
import pytest
from mcp.server.mcpserver.exceptions import ToolError
from PIL import Image
from pydantic import ValidationError

from src.models import JsonObject
from src.presentation import export, media, style
from src.protocol.definitions import DEFINITIONS
from src.protocol.server import build_server, lifespan


def test_presentation_catalog_contains_only_grouped_operations() -> None:
    server = build_server()
    definitions = {
        (definition.name, definition.operation)
        for definition in DEFINITIONS.values()
        if definition.name.startswith("wps_ppt_")
    }
    assert len(definitions) == 148
    assert all(operation is not None for _, operation in definitions)
    catalog = {tool.name for tool in server._tool_manager.list_tools()}
    groups = {name for name, _ in definitions}
    assert len(groups) == 22
    assert groups == {name for name in catalog if name.startswith("wps_ppt_")}
    assert not any(name.startswith("wpp_") for name in catalog)


@pytest.mark.parametrize(
    ("name", "params"),
    [
        ("wps_ppt_slides", {"slide_index": 1}),
        (
            "wps_ppt_slides",
            {"operation": "get", "slide_index": 1, "hidden": True},
        ),
        (
            "wps_ppt_media",
            {"operation": "add", "slide_index": 1, "file_path": "clip.mp4"},
        ),
        (
            "wps_ppt_style",
            {
                "operation": "apply",
                "slide_index": 1,
                "shapes": [1],
                "operations": [
                    {
                        "kind": "fill",
                        "options": {"fill_type": "solid", "weight": 2},
                    }
                ],
            },
        ),
        (
            "wps_ppt_style",
            {
                "operation": "apply",
                "slide_index": 1,
                "shapes": [1],
                "operations": [{"kind": "glow", "options": {"radius": -1}}],
            },
        ),
        (
            "wps_ppt_style",
            {
                "operation": "apply",
                "slide_index": 1,
                "shapes": [1],
                "operations": [
                    {"kind": "line", "options": {"document": "another.pptx"}}
                ],
            },
        ),
    ],
)
def test_grouped_inputs_are_rejected_before_native_launch(
    name: str, params: JsonObject
) -> None:
    async def scenario() -> None:
        server = build_server()
        async with lifespan(server):
            with pytest.raises(ToolError):
                await server.call_tool(name, {"params": params})
            assert server.runtime is not None
            assert all(
                not session.started
                for session in server.runtime.sessions.values()
            )

    anyio.run(scenario)


def test_style_rejects_duplicate_selectors_before_native_access() -> None:
    with pytest.raises(ValidationError, match="distinct"):
        style.ApplyStyleInput.model_validate(
            {
                "slide_index": 1,
                "shapes": ["title", "title"],
                "operations": [
                    {"kind": "fill", "options": {"fill_type": "none"}}
                ],
            }
        )


@pytest.mark.parametrize("kind", ["audio", "video"])
@pytest.mark.parametrize("linked", [False, True])
def test_media_insert_uses_one_native_contract(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, kind: str, linked: bool
) -> None:
    source = tmp_path / "media.bin"
    source.write_bytes(b"native media")
    insert = Mock(return_value=SimpleNamespace(Name="Media"))
    session = Mock()
    session.get_document().Slides().Shapes.AddMediaObject2 = insert
    monkeypatch.setattr(media, "get_session", Mock(return_value=session))
    params = media.AddMediaInput.model_validate(
        {
            "slide_index": 1,
            "file_path": str(source),
            "kind": kind,
            "link_to_file": linked,
            "width": 320.0,
            "height": 180.0,
        }
    )
    result = media.add_media(params)
    insert.assert_called_once_with(
        str(source),
        -1 if linked else 0,
        0 if linked else -1,
        100.0,
        100.0,
        320.0,
        180.0,
    )
    assert result["kind"] == kind
    assert source.read_bytes() == b"native media"


def test_shape_export_refuses_existing_files_before_connecting(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    target = tmp_path / "shape.png"
    target.write_bytes(b"existing")
    session = Mock()
    monkeypatch.setattr(export, "get_session", session)
    with pytest.raises(FileExistsError):
        export.export_shape(
            export.ExportShapeInput(
                slide_index=1, shape_name_or_index="box", file_path=str(target)
            )
        )
    session.assert_not_called()
    assert target.read_bytes() == b"existing"


@pytest.mark.parametrize("concurrent", [False, True])
def test_shape_export_publishes_complete_images_and_cleans_temporary_files(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, concurrent: bool
) -> None:
    target = tmp_path / "images" / "shape.png"
    temporary: list[Path] = []

    def render(file_path: str, format_type: int, *dimensions: int) -> None:
        staged = Path(file_path)
        temporary.append(staged)
        with Image.new("RGB", (8, 8), "blue") as image:
            image.save(staged, format="PNG")
        if concurrent:
            target.write_bytes(b"concurrent file")

    shape = SimpleNamespace(Name="box", Export=Mock(side_effect=render))
    monkeypatch.setattr(export, "get_session", Mock(return_value=Mock()))
    monkeypatch.setattr(export, "_get_shape", Mock(return_value=shape))
    monkeypatch.setattr(export, "goto_slide", Mock())
    params = export.ExportShapeInput(
        slide_index=1, shape_name_or_index="box", file_path=str(target)
    )
    if concurrent:
        with pytest.raises(FileExistsError):
            export.export_shape(params)
        assert target.read_bytes() == b"concurrent file"
    else:
        result = export.export_shape(params)
        assert result["file_path"] == str(target)
        with Image.open(target) as image:
            image.verify()
    assert len(temporary) == 1
    assert not temporary[0].parent.exists()


def test_shape_export_does_not_publish_invalid_native_output(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    target = tmp_path / "shape.png"
    temporary: list[Path] = []

    def render(file_path: str, format_type: int, *dimensions: int) -> None:
        staged = Path(file_path)
        temporary.append(staged)
        staged.write_bytes(b"not an image")

    shape = SimpleNamespace(Name="box", Export=Mock(side_effect=render))
    monkeypatch.setattr(export, "get_session", Mock(return_value=Mock()))
    monkeypatch.setattr(export, "_get_shape", Mock(return_value=shape))
    monkeypatch.setattr(export, "goto_slide", Mock())
    with pytest.raises(OSError):
        export.export_shape(
            export.ExportShapeInput(
                slide_index=1, shape_name_or_index="box", file_path=str(target)
            )
        )
    assert not target.exists()
    assert not temporary[0].parent.exists()
