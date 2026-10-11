"""Behavior checks for drawing preflight and shared native style."""

from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from pydantic import ValidationError

from src.models import JsonObject
from src.presentation import (
    connectors,
    freeform,
    layout,
    shape_formatting,
    style,
)


@pytest.fixture
def drawing_session(monkeypatch: pytest.MonkeyPatch) -> Mock:
    session = Mock()
    for module in (connectors, freeform, layout, style):
        monkeypatch.setattr(module, "get_session", Mock(return_value=session))
        monkeypatch.setattr(module, "goto_slide", Mock())
    return session


def test_batch_validates_all_parameters_before_resolving_targets(
    drawing_session: Mock,
) -> None:
    with pytest.raises(ValidationError, match="font_color"):
        style.ApplyStyleInput.model_validate(
            {
                "slide_index": 1,
                "shapes": ["first", "second"],
                "operations": [
                    {"kind": "fill", "options": {"fill_type": "solid"}},
                    {"kind": "line", "options": {"font_color": "#fff"}},
                ],
            }
        )
    drawing_session.get_document.assert_not_called()


@pytest.mark.parametrize(
    "key", ["document", "slide_index", "shape_name_or_index"]
)
def test_batch_rejects_embedded_selectors(
    drawing_session: Mock, key: str
) -> None:
    with pytest.raises(ValidationError, match=key):
        style.ApplyStyleInput.model_validate(
            {
                "slide_index": 1,
                "shapes": ["first"],
                "operations": [{"kind": "line", "options": {key: 1}}],
            }
        )
    drawing_session.get_document.assert_not_called()


def test_batch_resolves_all_shapes_before_first_write(
    drawing_session: Mock,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    resolver = Mock(side_effect=[Mock(), ValueError("Missing second shape")])
    apply = Mock()
    monkeypatch.setattr(style, "resolve_shape", resolver)
    monkeypatch.setattr(style, "apply_formatting", apply)
    params = style.ApplyStyleInput.model_validate(
        {
            "slide_index": 1,
            "shapes": ["first", "missing"],
            "operations": [{"kind": "line", "options": {"weight": 2.0}}],
        }
    )
    with pytest.raises(ValueError, match="Missing second"):
        style.apply_style(params)
    apply.assert_not_called()
    assert resolver.call_count == 2


def test_batch_checks_text_effect_targets_before_writing(
    drawing_session: Mock,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    first = Mock(Name="first", HasTextFrame=True)
    second = Mock(Name="second", HasTextFrame=False)
    monkeypatch.setattr(
        style, "resolve_shape", Mock(side_effect=[first, second])
    )
    apply = Mock()
    monkeypatch.setattr(style, "apply_effect", apply)
    params = style.ApplyStyleInput.model_validate(
        {
            "slide_index": 1,
            "shapes": ["first", "second"],
            "operations": [
                {"kind": "glow", "options": {"target": "text", "radius": 3.0}}
            ],
        }
    )
    with pytest.raises(ValueError, match="second.*text frame"):
        style.apply_style(params)
    apply.assert_not_called()


def test_batch_validates_each_operation_once_and_keeps_warnings(
    drawing_session: Mock,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    shapes = [Mock(Name="first"), Mock(Name="second")]
    monkeypatch.setattr(style, "resolve_shape", Mock(side_effect=shapes))
    apply = Mock(
        return_value={"status": "success", "warnings": ["Visible caution"]}
    )
    monkeypatch.setattr(style, "apply_formatting", apply)
    params = style.ApplyStyleInput.model_validate(
        {
            "slide_index": 1,
            "shapes": ["first", "second"],
            "operations": [{"kind": "fill", "options": {"fill_type": "none"}}],
        }
    )
    result = style.apply_style(params)
    assert result["status"] == "success"
    assert result["completed_operations"] == 2
    assert apply.call_args_list[0].args[1] is apply.call_args_list[1].args[1]
    assert apply.call_count == 2
    assert result["results"][1]["operations"][0]["warnings"] == [
        "Visible caution"
    ]


def test_batch_reports_runtime_partial_changes_and_continues(
    drawing_session: Mock,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(style, "resolve_shape", Mock(return_value=Mock()))
    apply = Mock(
        side_effect=[
            RuntimeError("Native write failed"),
            {"status": "success"},
        ]
    )
    monkeypatch.setattr(style, "apply_formatting", apply)
    params = style.ApplyStyleInput.model_validate(
        {
            "slide_index": 1,
            "shapes": ["first"],
            "operations": [
                {"kind": "line", "options": {"weight": 2.0}},
                {"kind": "fill", "options": {"fill_type": "none"}},
            ],
        }
    )
    result = style.apply_style(params)
    assert result["status"] == "partial_success"
    assert result["completed_operations"] == result["failed_operations"] == 1
    assert result["results"][0]["operations"][0]["may_have_changes"] is True
    assert result["results"][0]["operations"][1]["status"] == "success"


def test_glow_preserves_warning_for_invisible_shape_outline() -> None:
    shape = SimpleNamespace(
        Name="caption",
        HasTextFrame=True,
        Fill=SimpleNamespace(Visible=0),
        Line=SimpleNamespace(Visible=0),
        Glow=SimpleNamespace(Radius=0, Color=SimpleNamespace(RGB=0)),
    )
    params = shape_formatting.GlowOptions(radius=5.0, color="#123456")
    result = shape_formatting.apply_effect(shape, params)
    assert shape.Glow.Radius == 5.0
    assert shape.Glow.Color.RGB == 0x563412
    assert "target='text'" in result["warnings"][0]


def test_solid_fill_restores_visibility_after_no_fill() -> None:
    fill = SimpleNamespace(
        Visible=0, Solid=Mock(), ForeColor=SimpleNamespace(RGB=0)
    )
    shape = SimpleNamespace(Name="box", Fill=fill)
    shape_formatting.apply_formatting(
        shape,
        shape_formatting.FillOptions(fill_type="solid", color="#f00"),
    )
    assert fill.Visible == -1
    assert fill.ForeColor.RGB == 255
    fill.Solid.assert_called_once_with()


@pytest.mark.parametrize(
    "values",
    [
        {"fill_type": "gradient", "gradient_style": "unknown"},
        {"fill_type": "solid", "color": "not-a-color"},
        {"fill_type": "none", "transparency": 1.5},
    ],
)
def test_fill_rejects_invalid_parameters(values: JsonObject) -> None:
    with pytest.raises(ValidationError):
        shape_formatting.FillOptions(**values)


def test_connector_resolves_both_endpoints_before_creating(
    drawing_session: Mock,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        connectors,
        "resolve_shape",
        Mock(side_effect=[Mock(), ValueError("Missing target")]),
    )
    with pytest.raises(ValueError, match="Missing target"):
        connectors.add_connector(
            connectors.AddConnectorInput(
                slide_index=1, begin_shape="first", end_shape="missing"
            )
        )
    shapes = drawing_session.get_document().Slides().Shapes
    shapes.AddConnector.assert_not_called()


def test_connector_reconnect_failure_precedes_line_edits(
    drawing_session: Mock,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        connectors,
        "resolve_shape",
        Mock(
            side_effect=[Mock(Connector=True), ValueError("Missing endpoint")]
        ),
    )
    apply = Mock()
    monkeypatch.setattr(connectors, "apply_formatting", apply)
    with pytest.raises(ValueError, match="Missing endpoint"):
        connectors.format_connector(
            connectors.FormatConnectorInput(
                slide_index=1,
                shape_name_or_index="line",
                weight=2.0,
                end_shape="missing",
            )
        )
    apply.assert_not_called()


def test_connector_resolves_direction_without_distance_bias() -> None:
    sites = Mock(
        Count=3, side_effect=[(100.0, 50.0), (50.0, 150.0), (0.0, 50.0)]
    )
    shape = SimpleNamespace(
        Name="box", Left=0, Top=0, Width=100, Height=100, ConnectionSites=sites
    )
    assert connectors.resolve_site(shape, "right") == 1


def test_connector_rejects_out_of_range_numeric_site() -> None:
    shape = SimpleNamespace(Name="box", ConnectionSiteCount=4)
    with pytest.raises(ValueError, match="4 connection sites"):
        connectors.resolve_site(shape, 5)


def test_freeform_line_forces_auto_and_corner_requires_controls() -> None:
    line = freeform.NodeSpec(
        segment_type="line", editing_type="corner", x1=1, y1=2
    )
    assert line.editing_type == "auto"
    assert line.native_arguments() == (0, 0, 1.0, 2.0)
    with pytest.raises(ValidationError, match="require x2"):
        freeform.NodeSpec(
            segment_type="curve", editing_type="corner", x1=1, y1=2
        )


def test_freeform_insert_reuses_corner_curve_arguments() -> None:
    params = freeform.InsertNodeInput(
        slide_index=1,
        shape_name="path",
        after_index=1,
        segment_type="curve",
        editing_type="corner",
        x1=1,
        y1=2,
        x2=3,
        y2=4,
        x3=5,
        y3=6,
    )
    assert params.native_arguments() == (1, 1, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0)


def test_background_resolves_every_slide_before_changing_first(
    drawing_session: Mock,
) -> None:
    first = Mock()
    drawing_session.get_document().Slides.side_effect = [
        first,
        ValueError("Slide missing"),
    ]
    with pytest.raises(ValueError, match="Slide missing"):
        layout.set_slide_background(
            layout.SetSlideBackgroundInput(
                slide_indices=[1, 99], fill_type="solid", color="#fff"
            )
        )
    first.Background.Fill.Solid.assert_not_called()


def test_four_three_size_matches_creation_preset(
    drawing_session: Mock,
) -> None:
    result = layout.set_slide_size(layout.SetSlideSizeInput(preset="4:3"))
    assert result["width_points"] == 720
    assert result["height_points"] == 540


def test_background_rejects_ambiguous_slide_selection(
    drawing_session: Mock,
) -> None:
    with pytest.raises(ValidationError, match="exactly one"):
        layout.SetSlideBackgroundInput(
            slide_index=1, slide_indices=[2], fill_type="master"
        )
    drawing_session.get_document.assert_not_called()
