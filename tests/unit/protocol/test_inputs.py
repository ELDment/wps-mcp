"""Keep creation requirements distinct from optional read and update fields."""

import pytest
from pydantic import ValidationError

from wps_mcp.presentation.animations import (
    AddAnimationInput,
    UpdateAnimationInput,
)
from wps_mcp.writer.models import ExplicitRangeInput, RangeInput
from wps_mcp.writer.review import AddCommentInput
from wps_mcp.writer.text import DeleteTextInput, EditTextInput


@pytest.mark.parametrize(
    "model", [EditTextInput, DeleteTextInput, AddCommentInput]
)
@pytest.mark.parametrize("end", [{}, {"end": None}, {"end": 1}])
def test_writer_mutations_require_an_ordered_explicit_range(
    model: type[ExplicitRangeInput], end: dict[str, object]
) -> None:
    payload: dict[str, object] = {"start": 2, **end}
    if model is not DeleteTextInput:
        payload["text"] = "content"
    with pytest.raises(ValidationError):
        model.model_validate(payload)


def test_writer_read_allows_omitted_end_and_mutations_allow_empty_ranges() -> (
    None
):
    assert RangeInput(start=2).end is None
    assert EditTextInput(start=2, end=2, text="inserted").end == 2
    assert DeleteTextInput(end=0).start == 0
    assert AddCommentInput(end=0, text="comment").start == 0


def test_animation_creation_and_partial_update_have_distinct_defaults() -> (
    None
):
    created = AddAnimationInput(slide_index=1, shape_name_or_index="Title")
    assert (created.effect, created.trigger, created.exit) == (
        "appear",
        "on_click",
        False,
    )
    changed = UpdateAnimationInput(
        slide_index=1, animation_index=1, duration=0.5
    )
    assert (changed.effect, changed.trigger, changed.exit) == (
        None,
        None,
        None,
    )
    assert changed.model_fields_set == {
        "slide_index",
        "animation_index",
        "duration",
    }


@pytest.mark.parametrize("field", ["effect", "trigger", "exit"])
def test_animation_creation_rejects_null_settings(field: str) -> None:
    with pytest.raises(ValidationError):
        AddAnimationInput.model_validate(
            {"slide_index": 1, "shape_name_or_index": "Title", field: None}
        )


@pytest.mark.parametrize(
    "options",
    [
        {"effect": "missing"},
        {"trigger": "missing"},
        {"effect": "spin", "exit": True},
    ],
)
@pytest.mark.parametrize("create", [True, False])
def test_animation_selectors_are_validated_for_create_and_update(
    options: dict[str, object], create: bool
) -> None:
    payload = {"slide_index": 1, **options}
    with pytest.raises(ValidationError):
        if create:
            AddAnimationInput.model_validate(
                {"shape_name_or_index": "Title", **payload}
            )
        else:
            UpdateAnimationInput.model_validate(
                {"animation_index": 1, **payload}
            )


@pytest.mark.parametrize("value", ["abc", "#abc", "a1b2c3", "#A1b2C3"])
def test_rgb_contract_is_shared_and_normalized(value: str) -> None:
    from pydantic import TypeAdapter

    from wps_mcp.common.colors import RgbColor
    from wps_mcp.presentation.shapes.formatting import GlowOptions
    from wps_mcp.presentation.text.models import TextFormatSpec
    from wps_mcp.spreadsheet.ranges import FormatRangeInput

    expected = "#AABBCC" if len(value.removeprefix("#")) == 3 else "#A1B2C3"
    assert TypeAdapter(RgbColor).validate_python(value) == expected
    assert GlowOptions(radius=1, color=value).color == expected
    assert TextFormatSpec(color=value).color == expected
    assert (
        FormatRangeInput(address="A1", font_color=value).font_color == expected
    )
    assert TypeAdapter(RgbColor).json_schema()["pattern"]


@pytest.mark.parametrize("value", ["", "#12", "#12345G", "red", " #fff", 42])
def test_rgb_contract_rejects_invalid_values_before_native_work(
    value: object,
) -> None:
    from pydantic import TypeAdapter

    from wps_mcp.common.colors import RgbColor

    with pytest.raises(ValidationError):
        TypeAdapter(RgbColor).validate_python(value)


def test_removed_dash_alias_is_not_accepted() -> None:
    from wps_mcp.presentation.constants import dash_style_value

    with pytest.raises(ValueError, match="Unknown dash_style"):
        dash_style_value("dot")
    assert dash_style_value("round_dot") == 3
