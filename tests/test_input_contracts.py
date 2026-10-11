"""Keep creation requirements distinct from optional read and update fields."""

import pytest
from pydantic import ValidationError

from src.presentation.animation import (
    AddAnimationInput,
    UpdateAnimationInput,
)
from src.writer.models import ExplicitRangeInput, RangeInput
from src.writer.review import AddCommentInput
from src.writer.text import DeleteTextInput, EditTextInput


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
