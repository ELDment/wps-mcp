"""Shape lookup, validation and atomic preflight behavior."""

from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from pydantic import ValidationError

from wps_mcp.presentation import pictures
from wps_mcp.presentation import styles as style
from wps_mcp.presentation.constants import MSO_GROUP, MSO_PICTURE
from wps_mcp.presentation.shapes import lookup as shape_lookup
from wps_mcp.presentation.shapes import models as shape_models
from wps_mcp.presentation.shapes import tools as shapes


def make_shape(
    name: str, identity: int = 1, children: list[SimpleNamespace] | None = None
) -> SimpleNamespace:
    return SimpleNamespace(
        Name=name,
        Id=identity,
        Type=MSO_GROUP if children else 1,
        GroupItems=Mock(
            Count=len(children or []),
            side_effect=lambda index: (children or [])[index - 1],
        ),
        Left=10.0,
        Top=20.0,
        Width=100.0,
        Height=60.0,
        Rotation=0.0,
        HasTextFrame=False,
    )


def make_slide(items: list[SimpleNamespace]) -> SimpleNamespace:
    return SimpleNamespace(
        Shapes=Mock(
            Count=len(items), side_effect=lambda index: items[index - 1]
        )
    )


@pytest.fixture
def environment(monkeypatch: pytest.MonkeyPatch) -> SimpleNamespace:
    first, second = make_shape("First", 1), make_shape("Second", 2)
    slide = make_slide([first, second])
    presentation = SimpleNamespace(Slides=Mock(Count=1, return_value=slide))
    session = SimpleNamespace(
        get_document=Mock(return_value=presentation),
        get_application=Mock(),
        activate_window=Mock(),
    )
    for module in (shapes, pictures, style):
        monkeypatch.setattr(module, "get_session", lambda component: session)
        monkeypatch.setattr(module, "goto_slide", Mock())
    return SimpleNamespace(
        first=first, second=second, slide=slide, session=session
    )


def test_nested_shape_paths_and_separator_names() -> None:
    child = make_shape("A/B")
    inner = make_shape("Inner", children=[child])
    outer = make_shape("Outer", children=[inner])
    slide = make_slide([outer])
    assert shape_lookup.resolve_with_path(slide, "A/B") == (
        child,
        "Outer/Inner/A/B",
    )
    assert shape_lookup.resolve_shape(slide, "Outer/Inner/A/B") is child
    with pytest.raises(shape_lookup.ShapeNotFoundError):
        shape_lookup.resolve_shape(slide, "Outer/A/B")


def test_duplicate_child_names_require_group_path() -> None:
    first, second = make_shape("Child"), make_shape("Child")
    slide = make_slide(
        [
            make_shape("Left", children=[first]),
            make_shape("Right", children=[second]),
        ]
    )
    with pytest.raises(shape_lookup.AmbiguousShapeError, match="Left/Child"):
        shape_lookup.resolve_shape(slide, "Child")
    assert shape_lookup.resolve_shape(slide, "Right/Child") is second
    with pytest.raises(ValueError, match="top-level"):
        shape_lookup.require_top_level(slide, ["Right/Child"], 1, "align")


def test_exact_top_level_name_takes_precedence() -> None:
    top, child = make_shape("Child"), make_shape("Child")
    slide = make_slide([make_shape("Group", children=[child]), top])
    assert shape_lookup.resolve_shape(slide, "Child") is top
    with pytest.raises(shape_lookup.ShapeNotFoundError):
        shape_lookup.resolve_shape(slide, True)


def test_group_read_errors_are_not_reported_as_missing() -> None:
    group = make_shape("Group", children=[make_shape("Child")])
    group.GroupItems.side_effect = RuntimeError("disconnected")
    with pytest.raises(RuntimeError, match="disconnected"):
        shape_lookup.resolve_shape(make_slide([group]), "Child")


def test_batch_resolves_every_target_before_writing(
    environment: SimpleNamespace,
) -> None:
    params = shape_models.UpdateShapeInput(
        slide_index=1, shape_names=["First", "Missing"], dleft=12.0
    )
    with pytest.raises(shape_lookup.ShapeNotFoundError):
        shapes.update_shape(params)
    assert environment.first.Left == 10.0


def test_batch_validates_all_dimensions_before_writing(
    environment: SimpleNamespace,
) -> None:
    environment.second.Width = 5.0
    params = shape_models.UpdateShapeInput(
        slide_index=1, all=True, dwidth=-10.0
    )
    with pytest.raises(ValueError, match="negative"):
        shapes.update_shape(params)
    assert environment.first.Width == 100.0
    assert environment.second.Width == 5.0


def test_batch_offsets_each_shape_only_once(
    environment: SimpleNamespace,
) -> None:
    params = shape_models.UpdateShapeInput(
        slide_index=1, shape_names=["First", "First"], dleft=5.0
    )
    result = shapes.update_shape(params)
    assert result["count"] == 1
    assert environment.first.Left == 15.0
    assert environment.second.Left == 10.0


def test_all_shapes_preserves_duplicate_names(
    environment: SimpleNamespace,
) -> None:
    environment.second.Name = "First"
    result = shapes.update_shape(
        shape_models.UpdateShapeInput(slide_index=1, all=True, dtop=5.0)
    )
    assert result["count"] == 2
    assert environment.first.Top == environment.second.Top == 25.0


def test_adjustments_validate_before_geometry(
    environment: SimpleNamespace,
) -> None:
    environment.first.Adjustments = Mock(Count=1)
    params = shape_models.UpdateShapeInput(
        slide_index=1, shape_name="First", left=30.0, adjustments={"2": 0.5}
    )
    with pytest.raises(ValueError, match="out of range"):
        shapes.update_shape(params)
    assert environment.first.Left == 10.0


def test_model_keeps_text_whitespace_and_document() -> None:
    params = shape_models.AddTextboxInput(
        document="deck.pptx",
        slide_index=1,
        left=0.0,
        top=0.0,
        width=100.0,
        height=20.0,
        text="  exact text  ",
    )
    assert params.document == "deck.pptx"
    assert params.text == "  exact text  "
    with pytest.raises(ValidationError):
        shape_models.AddTextboxInput.model_validate(
            {**params.model_dump(), "width": "100"}
        )
    with pytest.raises(ValidationError):
        shape_models.AddTextboxInput.model_validate(
            {**params.model_dump(), "width": -1}
        )
    with pytest.raises(ValidationError):
        shape_models.AddTextboxInput.model_validate(
            {**params.model_dump(), "font_name": ""}
        )


def test_shape_identifier_is_exclusive() -> None:
    with pytest.raises(ValidationError, match="exactly one"):
        shape_models.ShapeIdentifierInput(slide_index=1)
    with pytest.raises(ValidationError, match="exactly one"):
        shape_models.ShapeIdentifierInput(
            slide_index=1, shape_name="First", shape_index=1
        )


def test_invalid_zorder_raises_without_native_write(
    environment: SimpleNamespace,
) -> None:
    with pytest.raises(ValueError, match="Unknown"):
        shapes.set_shape_zorder(
            shape_models.SetZOrderInput(
                slide_index=1, shape_name="First", command="oops"
            )
        )
    environment.session.get_document.assert_not_called()


def test_selection_resolves_every_name_before_selecting(
    environment: SimpleNamespace,
) -> None:
    environment.first.Select = Mock()
    with pytest.raises(shape_lookup.ShapeNotFoundError):
        shapes.select_shapes(
            shape_models.SelectShapesInput(
                slide_index=1, shape_names=["First", "Missing"]
            )
        )
    environment.first.Select.assert_not_called()
    environment.session.activate_window.assert_not_called()


def test_crop_probe_restores_temporary_native_state(
    environment: SimpleNamespace,
) -> None:
    shape = environment.first
    shape.Type = MSO_PICTURE
    shape.LockAspectRatio = -1
    shape.PictureFormat = SimpleNamespace(
        CropLeft=1, CropRight=2, CropTop=3, CropBottom=4
    )

    def scale_width(scale: float, original: int) -> None:
        shape.Width = 500.0

    shape.ScaleWidth = scale_width
    shape.ScaleHeight = Mock(side_effect=RuntimeError("native probe failed"))
    with pytest.raises(RuntimeError, match="native probe"):
        pictures.crop_picture(
            pictures.CropPictureInput(
                slide_index=1, shape_name_or_index="First", crop_fit="square"
            )
        )
    assert shape.LockAspectRatio == -1
    assert shape.Width == 100.0
    assert shape.Height == 60.0
    assert vars(shape.PictureFormat) == dict(
        CropLeft=1, CropRight=2, CropTop=3, CropBottom=4
    )


def test_temporary_style_shape_is_deleted_when_operation_fails() -> None:
    scratch = SimpleNamespace(Delete=Mock())
    slide = SimpleNamespace(
        Shapes=SimpleNamespace(AddShape=Mock(return_value=scratch))
    )
    with pytest.raises(RuntimeError, match="style failed"):
        with style._temporary_shape(slide):
            raise RuntimeError("style failed")
    scratch.Delete.assert_called_once()


def test_style_cleanup_does_not_mask_original_failure(
    caplog: pytest.LogCaptureFixture,
) -> None:
    scratch = SimpleNamespace(
        Delete=Mock(side_effect=RuntimeError("cleanup failed"))
    )
    slide = SimpleNamespace(
        Shapes=SimpleNamespace(AddShape=Mock(return_value=scratch))
    )
    with pytest.raises(ValueError, match="original"):
        with style._temporary_shape(slide):
            raise ValueError("original")
    assert "temporary default-style shape" in caplog.text


def test_moving_a_horizontal_line_preserves_zero_height(
    environment: SimpleNamespace,
) -> None:
    environment.first.Height = 0.0
    result = shapes.update_shape(
        shape_models.UpdateShapeInput(
            slide_index=1, shape_name="First", dleft=5.0
        )
    )
    assert result["left"] == 15.0
    assert environment.first.Height == 0.0


@pytest.mark.parametrize("kind", ["shape", "textbox"])
def test_new_shape_text_normalizes_crlf_once(
    environment: SimpleNamespace, kind: str
) -> None:
    native = Mock(Name="Body", ZOrderPosition=1, AutoShapeType=1)
    environment.slide.Shapes.AddShape.return_value = native
    environment.slide.Shapes.AddTextbox.return_value = native
    options = {
        "slide_index": 1,
        "left": 0.0,
        "top": 0.0,
        "width": 100.0,
        "height": 50.0,
        "text": " first\r\nsecond\nthird\rfourth\v ",
    }
    if kind == "shape":
        shapes.add_shape(
            shape_models.AddShapeInput(**options, shape_type="rectangle")
        )
    else:
        shapes.add_textbox(shape_models.AddTextboxInput(**options))
    assert native.TextFrame.TextRange.Text == (
        " first\rsecond\rthird\rfourth\v "
    )


@pytest.mark.parametrize(
    ("width", "height", "locked"),
    [(80.0, 40.0, 0), (80.0, None, -1), (None, 40.0, -1)],
)
def test_local_and_downloaded_pictures_share_dimension_rules(
    environment: SimpleNamespace,
    tmp_path: Path,
    width: float | None,
    height: float | None,
    locked: int,
) -> None:
    local = SimpleNamespace(
        Name="Local", ZOrderPosition=1, Width=100.0, Height=50.0
    )
    downloaded = SimpleNamespace(
        Name="Downloaded", ZOrderPosition=2, Width=100.0, Height=50.0
    )
    path = tmp_path / "picture.png"
    path.write_bytes(b"local image")
    environment.slide.Shapes.AddPicture.side_effect = [local, downloaded]
    pictures.add_picture(
        pictures.AddPictureInput(
            slide_index=1,
            file_path=str(path),
            left=0.0,
            top=0.0,
            width=width,
            height=height,
        )
    )
    pictures._insert_picture(
        environment.slide,
        b"image",
        ".png",
        0.0,
        0.0,
        width,
        height,
        False,
        "front",
    )
    assert local.LockAspectRatio == downloaded.LockAspectRatio == locked
    assert (local.Width, local.Height) == (
        downloaded.Width,
        downloaded.Height,
    )


def test_access_validates_slide_and_shape_before_navigation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from wps_mcp.presentation import access

    slide = make_slide([make_shape("Body")])
    presentation = SimpleNamespace(Slides=Mock(Count=1, return_value=slide))
    session = SimpleNamespace(get_document=Mock(return_value=presentation))
    navigate = Mock()
    monkeypatch.setattr(access, "get_session", lambda component: session)
    monkeypatch.setattr(access, "goto_slide", navigate)
    with pytest.raises(ValueError, match="out of range"):
        access.get_slide(2, "chosen.pptx", navigate=True)
    presentation.Slides.assert_not_called()
    navigate.assert_not_called()
    with pytest.raises(shape_lookup.ShapeNotFoundError):
        access.get_shape(
            shape_models.ShapeInput(
                document="chosen.pptx",
                slide_index=1,
                shape_name_or_index="missing",
            )
        )
    navigate.assert_not_called()
    shape = access.get_shape(
        shape_models.ShapeInput(
            document="chosen.pptx",
            slide_index=1,
            shape_name_or_index="Body",
        )
    )
    assert shape.Name == "Body"
    session.get_document.assert_called_with("chosen.pptx")
    navigate.assert_called_once_with(1)
