"""Automate WPS Presentation animations and transitions.

Handles slide transitions, adding/listing/removing shape animations, and
clearing all animations from a slide.
"""

from __future__ import annotations

from typing import Self

from pydantic import Field, model_validator

from src.common.colors import hex_to_int
from src.models import DocumentInput, JsonObject, NativeObject
from src.native.runtime import get_session
from src.presentation.constants import (
    AFTER_EFFECT_MAP,
    AFTER_EFFECT_NAMES,
    ANIM_DIRECTION_MAP,
    ANIM_DIRECTION_NAMES,
    ANIMATION_EFFECT_NAMES,
    ANIMATION_TRIGGER_NAMES,
    BUILD_LEVEL_MAP,
    BUILD_LEVEL_NAMES,
    MSO_FALSE,
    MSO_TRUE,
    TEXT_UNIT_EFFECT_MAP,
    TEXT_UNIT_EFFECT_NAMES,
)
from src.presentation.navigation import goto_slide
from src.presentation.shape_lookup import resolve_shape as _get_shape
from src.protocol.definitions import tool


class CopyAnimationInput(DocumentInput):
    """Input for copying animation from one shape to another."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    source_shape: str | int = Field(
        ...,
        description=(
            "Source shape name (str) or 1-based index (int). Prefer name "
            "— indices shift when shapes are added/removed"
        ),
    )
    target_shape: str | int = Field(
        ...,
        description=(
            "Target shape name (str) or 1-based index (int). Prefer name "
            "— indices shift when shapes are added/removed"
        ),
    )


@tool(
    "wps_ppt_animation",
    component="wpp",
    annotations={"title": "Copy Animation", "idempotentHint": True},
    operation="copy",
)
def copy_animation(params: CopyAnimationInput) -> JsonObject:
    """Copy effects between shapes on the same slide."""
    app = get_session("wpp").get_application()
    goto_slide(app, params.slide_index)
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    src = _get_shape(slide, params.source_shape)
    tgt = _get_shape(slide, params.target_shape)
    src.PickupAnimation()
    tgt.ApplyAnimation()
    result = {
        "success": True,
        "source_shape": src.Name,
        "target_shape": tgt.Name,
    }
    return result


ANIMATION_EFFECT_MAP: dict[str, int] = {
    "appear": 1,
    "fly": 2,
    "blinds": 3,
    "box": 4,
    "checkerboard": 5,
    "circle": 6,
    "diamond": 8,
    "dissolve": 9,
    "fade": 10,
    "split": 16,
    "wipe": 22,
    "zoom": 23,
    "bounce": 26,
    "float": 30,
    "grow_and_turn": 31,
    "change_fill_color": 54,
    "change_font": 55,
    "change_font_color": 56,
    "change_font_size": 57,
    "grow_shrink": 59,
    "spin": 61,
    "transparency": 62,
    "bold_flash": 63,
    "color_wave": 69,
    "darken": 73,
    "desaturate": 74,
    "flash_bulb": 75,
    "lighten": 78,
    "teeter": 80,
    "wave": 82,
    "path_circle": 86,
    "path_diamond": 88,
    "path_star": 90,
    "path_square": 92,
    "path_heart": 94,
    "path_loop": 109,
    "path_left": 120,
    "path_arc_down": 122,
    "path_zigzag": 123,
    "path_sine_wave": 125,
    "path_bounce_left": 126,
    "path_down": 127,
    "path_arc_up": 129,
    "path_spiral_right": 131,
    "path_wave": 132,
    "path_diagonal_down_right": 134,
    "path_arc_left": 136,
    "path_funnel": 137,
    "path_spring": 138,
    "path_bounce_right": 139,
    "path_diagonal_up_right": 141,
    "path_arc_right": 143,
    "path_up": 148,
    "path_right": 149,
}
TRIGGER_MAP: dict[str, int] = {
    "on_click": 1,
    "with_previous": 2,
    "after_previous": 3,
    "on_shape_click": 4,
}
TRANSITION_EFFECT_MAP: dict[str, int] = {
    "none": 0,
    "cut": 257,
    "fade": 3844,
    "push": 3845,
    "wipe": 3846,
    "split": 3847,
    "reveal": 3848,
    "random": 513,
    "blinds_horizontal": 769,
    "blinds_vertical": 770,
    "dissolve": 1537,
}


class SetSlideTransitionInput(DocumentInput):
    """Transition effect and timing for a single slide."""

    slide_index: int = Field(ge=1)
    effect: int | str = "fade"
    duration: float | None = Field(default=None, ge=0)
    advance_on_click: bool | None = None
    advance_on_time: bool | None = None
    advance_time: float | None = Field(default=None, ge=0)

    @model_validator(mode="after")
    def validate_effect(self) -> Self:
        """Reject unknown transition effect names."""
        if (
            isinstance(self.effect, str)
            and self.effect not in TRANSITION_EFFECT_MAP
        ):
            raise ValueError(f"Unknown transition effect: {self.effect}")
        return self


class AnimationOptions(DocumentInput):
    """Shared animation settings; timing uses seconds."""

    slide_index: int = Field(ge=1)
    duration: float | None = Field(default=None, ge=0)
    delay: float | None = Field(default=None, ge=0)
    direction: str | int | None = None
    repeat_count: int | None = Field(default=None, ge=0)
    auto_reverse: bool | None = None
    rewind: bool | None = None
    smooth_start: bool | None = None
    smooth_end: bool | None = None
    after_effect: str | None = None
    dim_color: str | None = Field(default=None, pattern="^#[0-9a-fA-F]{6}$")
    build_level: str | None = None
    text_unit_effect: str | None = None
    animate_in_reverse: bool | None = None
    animate_background: bool | None = None

    @model_validator(mode="after")
    def validate_options(self) -> Self:
        """Validate effect names and dependent animation settings."""
        for field, choices in (
            ("direction", ANIM_DIRECTION_MAP),
            ("after_effect", AFTER_EFFECT_MAP),
            ("build_level", BUILD_LEVEL_MAP),
            ("text_unit_effect", TEXT_UNIT_EFFECT_MAP),
        ):
            value = getattr(self, field)
            if isinstance(value, str) and value not in choices:
                raise ValueError(
                    f"Unknown {field}: {value}; "
                    f"choose from {', '.join(choices)}"
                )
        if self.dim_color is not None and self.after_effect != "dim":
            raise ValueError("dim_color requires after_effect='dim'")
        return self


def _validate_effect_options(
    effect: int | str | None, trigger: str | None, exit_effect: bool | None
) -> None:
    """Validate effect selectors shared by creation and partial updates."""
    for field, value, choices in (
        ("effect", effect, ANIMATION_EFFECT_MAP),
        ("trigger", trigger, TRIGGER_MAP),
    ):
        if isinstance(value, str) and value not in choices:
            raise ValueError(
                f"Unknown {field}: {value}; choose from {', '.join(choices)}"
            )
    effect_id = (
        ANIMATION_EFFECT_MAP[effect] if isinstance(effect, str) else effect
    )
    if exit_effect and effect_id is not None and not 1 <= effect_id <= 53:
        raise ValueError(
            "Exit animations require an entrance/exit effect (1-53)"
        )


class AddAnimationInput(AnimationOptions):
    """Add an effect by shape name or one-based index."""

    shape_name_or_index: str | int
    effect: int | str = "appear"
    trigger: str = "on_click"
    exit: bool = False
    trigger_shape: str | int | None = None

    @model_validator(mode="after")
    def validate_trigger(self) -> Self:
        """Require a trigger shape exactly for shape-click effects."""
        _validate_effect_options(self.effect, self.trigger, self.exit)
        if (self.trigger == "on_shape_click") != (
            self.trigger_shape is not None
        ):
            raise ValueError(
                "trigger_shape is required exactly "
                "when trigger='on_shape_click'"
            )
        return self


class ListAnimationsInput(DocumentInput):
    """Read every animation on the selected slide."""

    slide_index: int = Field(ge=1)


class RemoveAnimationInput(DocumentInput):
    """An effect index in the main or specified interactive sequence."""

    slide_index: int = Field(ge=1)
    animation_index: int = Field(ge=1)
    sequence_index: int | None = Field(default=None, ge=1)


class ClearAnimationsInput(DocumentInput):
    """Remove all effects from the selected slide."""

    slide_index: int = Field(ge=1)
    clear_transitions: bool = False


class UpdateAnimationInput(AnimationOptions):
    """Change one explicitly selected effect."""

    animation_index: int = Field(ge=1)
    effect: int | str | None = None
    trigger: str | None = None
    exit: bool | None = None
    sequence_index: int | None = Field(default=None, ge=1)
    move_to: int | None = Field(default=None, ge=1)

    @model_validator(mode="after")
    def validate_changes(self) -> Self:
        """Require at least one explicitly supplied animation change."""
        _validate_effect_options(self.effect, self.trigger, self.exit)
        changes = self.model_fields_set - {
            "document",
            "slide_index",
            "animation_index",
            "sequence_index",
        }
        if not any(getattr(self, field) is not None for field in changes):
            raise ValueError(
                "Provide at least one animation setting to change"
            )
        return self


def _animation_sequence(
    slide: NativeObject, sequence_index: int | None
) -> NativeObject:
    """Resolve an explicit sequence before any animation mutation."""
    if sequence_index is None:
        return slide.TimeLine.MainSequence
    sequences = slide.TimeLine.InteractiveSequences
    if sequence_index > sequences.Count:
        raise ValueError(f"sequence_index must not exceed {sequences.Count}")
    return sequences(sequence_index)


def _configure_animation(
    sequence: NativeObject,
    effect: NativeObject,
    params: AddAnimationInput | UpdateAnimationInput,
) -> NativeObject:
    """Apply settings and follow effects returned by WPS conversions."""
    if params.exit is not None:
        if params.exit and not 1 <= effect.EffectType <= 53:
            raise ValueError(
                "Exit animations require an entrance/exit effect (1-53)"
            )
        effect.Exit = MSO_TRUE if params.exit else MSO_FALSE
    timing = effect.Timing
    for attribute, value in (
        ("Duration", params.duration),
        ("TriggerDelayTime", params.delay),
        ("RepeatCount", params.repeat_count),
    ):
        if value is not None:
            setattr(timing, attribute, value)
    for attribute, value in (
        ("AutoReverse", params.auto_reverse),
        ("RewindAtEnd", params.rewind),
        ("SmoothStart", params.smooth_start),
        ("SmoothEnd", params.smooth_end),
    ):
        if value is not None:
            setattr(timing, attribute, MSO_TRUE if value else MSO_FALSE)
    if params.direction is not None:
        effect.EffectParameters.Direction = (
            ANIM_DIRECTION_MAP[params.direction]
            if isinstance(params.direction, str)
            else params.direction
        )
    conversions: list[tuple[str, tuple[int, ...]]] = []
    if (
        isinstance(params, UpdateAnimationInput)
        and params.build_level is not None
    ):
        conversions.append(
            ("ConvertToBuildLevel", (BUILD_LEVEL_MAP[params.build_level],))
        )
    if params.text_unit_effect is not None:
        conversions.append(
            (
                "ConvertToTextUnitEffect",
                (TEXT_UNIT_EFFECT_MAP[params.text_unit_effect],),
            )
        )
    if params.animate_in_reverse is not None:
        conversions.append(
            (
                "ConvertToAnimateInReverse",
                (MSO_TRUE if params.animate_in_reverse else MSO_FALSE,),
            )
        )
    background = params.animate_background
    if (
        isinstance(params, AddAnimationInput)
        and background is None
        and params.build_level not in (None, "none")
    ):
        background = False
    if background is not None:
        conversions.append(
            (
                "ConvertToAnimateBackground",
                (MSO_TRUE if background else MSO_FALSE,),
            )
        )
    if params.after_effect is not None:
        arguments: tuple[int, ...] = (AFTER_EFFECT_MAP[params.after_effect],)
        if params.dim_color is not None:
            arguments += (hex_to_int(params.dim_color),)
        conversions.append(("ConvertToAfterEffect", arguments))
    for method, arguments in conversions:
        effect = getattr(sequence, method)(effect, *arguments)
        if effect is None:
            raise RuntimeError(
                f"WPS {method} returned no effect; "
                f"inspect animations before continuing"
            )
    return effect


def _animation_info(effect: NativeObject) -> JsonObject:
    """Describe main and interactive effects with one output schema."""
    effect_type, trigger_type = (
        int(effect.EffectType),
        int(effect.Timing.TriggerType),
    )
    is_exit = bool(effect.Exit)
    category = "unknown"
    for start, end, name in (
        (1, 53, "entrance"),
        (54, 82, "emphasis"),
        (83, 85, "media"),
        (86, 149, "motion_path"),
    ):
        if start <= effect_type <= end:
            category = name
            break
    result: JsonObject = {
        "index": effect.Index,
        "shape_name": effect.Shape.Name,
        "effect_type": effect_type,
        "effect_name": ANIMATION_EFFECT_NAMES.get(
            effect_type, f"unknown({effect_type})"
        ),
        "trigger_type": trigger_type,
        "trigger_name": ANIMATION_TRIGGER_NAMES.get(
            trigger_type, f"unknown({trigger_type})"
        ),
        "duration": effect.Timing.Duration,
        "exit": is_exit,
        "category": "exit" if is_exit else category,
    }
    for key, owner, attribute, names in (
        ("direction", "EffectParameters", "Direction", ANIM_DIRECTION_NAMES),
        (
            "after_effect",
            "EffectInformation",
            "AfterEffect",
            AFTER_EFFECT_NAMES,
        ),
        (
            "build_level",
            "EffectInformation",
            "BuildByLevelEffect",
            BUILD_LEVEL_NAMES,
        ),
        (
            "text_unit_effect",
            "EffectInformation",
            "TextUnitEffect",
            TEXT_UNIT_EFFECT_NAMES,
        ),
    ):
        try:
            value = int(getattr(getattr(effect, owner), attribute))
            result[key], result[key + "_name"] = (
                value,
                names.get(value, f"unknown({value})"),
            )
        except Exception:
            result[key], result[key + "_name"] = None, None
    for key, attribute in (
        ("animate_in_reverse", "AnimateTextInReverse"),
        ("animate_background", "AnimateBackground"),
    ):
        try:
            result[key] = bool(getattr(effect.EffectInformation, attribute))
        except Exception:
            result[key] = None
    return result


@tool(
    "wps_ppt_animation",
    component="wpp",
    annotations={"title": "Set Slide Transition", "idempotentHint": True},
    operation="transition",
)
def set_slide_transition(params: SetSlideTransitionInput) -> JsonObject:
    """Set a transition and optional slide advance timing."""
    slide = get_session("wpp").get_document().Slides(params.slide_index)
    transition = slide.SlideShowTransition
    effect = (
        TRANSITION_EFFECT_MAP[params.effect]
        if isinstance(params.effect, str)
        else params.effect
    )
    transition.EntryEffect = effect
    for attribute, value in (
        ("Duration", params.duration),
        ("AdvanceTime", params.advance_time),
    ):
        if value is not None:
            setattr(transition, attribute, value)
    for attribute, value in (
        ("AdvanceOnClick", params.advance_on_click),
        ("AdvanceOnTime", params.advance_on_time),
    ):
        if value is not None:
            setattr(transition, attribute, MSO_TRUE if value else MSO_FALSE)
    return {
        "success": True,
        "slide_index": params.slide_index,
        "effect": effect,
    }


@tool(
    "wps_ppt_animation",
    component="wpp",
    annotations={"title": "Add Animation"},
    operation="add",
)
def add_animation(params: AddAnimationInput) -> JsonObject:
    """Add an entrance, exit, emphasis or motion path effect."""
    slide = get_session("wpp").get_document().Slides(params.slide_index)
    shape = _get_shape(slide, params.shape_name_or_index)
    effect_type = (
        ANIMATION_EFFECT_MAP[params.effect]
        if isinstance(params.effect, str)
        else params.effect
    )
    sequence = slide.TimeLine.MainSequence
    sequence_index: int | None = None
    trigger_shape: NativeObject = None
    if params.trigger_shape is not None:
        trigger_shape = _get_shape(slide, params.trigger_shape)
        sequences = slide.TimeLine.InteractiveSequences
        for index in range(1, sequences.Count + 1):
            candidate = sequences(index)
            if (
                candidate.Count
                and candidate(1).Timing.TriggerShape.Id == trigger_shape.Id
            ):
                sequence, sequence_index = candidate, index
                break
        if sequence_index is None:
            sequence = sequences.Add()
            sequence_index = int(sequences.Count)
    effect = sequence.AddEffect(
        shape,
        effect_type,
        BUILD_LEVEL_MAP[params.build_level] if params.build_level else 0,
        TRIGGER_MAP[params.trigger],
    )
    if trigger_shape is not None:
        effect.Timing.TriggerShape = trigger_shape
    effect = _configure_animation(sequence, effect, params)
    result: JsonObject = {
        "success": True,
        "shape_name": shape.Name,
        "effect": effect_type,
        "exit": params.exit,
        "animation_index": effect.Index,
    }
    if sequence_index is not None:
        assert trigger_shape is not None
        result.update(
            sequence_index=sequence_index,
            effect_index=effect.Index,
            trigger_shape_name=trigger_shape.Name,
        )
    return result


@tool(
    "wps_ppt_animation",
    component="wpp",
    annotations={
        "title": "List Animations",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="list",
)
def list_animations(params: ListAnimationsInput) -> JsonObject:
    """List main and interactive effects with sequence indices."""
    timeline = (
        get_session("wpp").get_document().Slides(params.slide_index).TimeLine
    )
    main = timeline.MainSequence
    animations = [
        _animation_info(main(index)) for index in range(1, main.Count + 1)
    ]
    interactive: list[JsonObject] = []
    sequences = timeline.InteractiveSequences
    for sequence_index in range(1, sequences.Count + 1):
        sequence = sequences(sequence_index)
        for effect_index in range(1, sequence.Count + 1):
            effect = sequence(effect_index)
            entry = _animation_info(effect)
            entry.update(
                sequence_index=sequence_index,
                effect_index=effect_index,
                trigger_shape_name=effect.Timing.TriggerShape.Name,
            )
            interactive.append(entry)
    return {
        "success": True,
        "slide_index": params.slide_index,
        "main_sequence_count": main.Count,
        "animations": animations,
        "interactive_sequences": interactive,
        "interactive_count": len(interactive),
    }


@tool(
    "wps_ppt_animation",
    component="wpp",
    annotations={"title": "Remove Animation", "destructiveHint": True},
    operation="remove",
)
def remove_animation(params: RemoveAnimationInput) -> JsonObject:
    """Delete an effect from a main or interactive sequence."""
    slide = get_session("wpp").get_document().Slides(params.slide_index)
    sequence = _animation_sequence(slide, params.sequence_index)
    if params.animation_index > sequence.Count:
        raise ValueError(f"animation_index must not exceed {sequence.Count}")
    sequence(params.animation_index).Delete()
    return {
        "success": True,
        "slide_index": params.slide_index,
        "remaining_count": sequence.Count,
    }


@tool(
    "wps_ppt_animation",
    component="wpp",
    annotations={
        "title": "Clear Animations",
        "destructiveHint": True,
        "idempotentHint": True,
    },
    operation="clear",
)
def clear_animations(params: ClearAnimationsInput) -> JsonObject:
    """Clear all animation sequences and optionally the transition."""
    slide = get_session("wpp").get_document().Slides(params.slide_index)
    timeline = slide.TimeLine
    main = timeline.MainSequence
    count = main.Count
    for index in range(count, 0, -1):
        main(index).Delete()
    interactive_count = 0
    for index in range(timeline.InteractiveSequences.Count, 0, -1):
        sequence = timeline.InteractiveSequences(index)
        interactive_count += sequence.Count
        for effect_index in range(sequence.Count, 0, -1):
            sequence(effect_index).Delete()
    if params.clear_transitions:
        slide.SlideShowTransition.EntryEffect = 0
    return {
        "success": True,
        "slide_index": params.slide_index,
        "cleared_count": count,
        "interactive_cleared": interactive_count,
    }


@tool(
    "wps_ppt_animation",
    component="wpp",
    annotations={"title": "Update Animation"},
    operation="update",
)
def update_animation(params: UpdateAnimationInput) -> JsonObject:
    """Update one selected effect without retrying failed mutations."""
    slide = get_session("wpp").get_document().Slides(params.slide_index)
    sequence = _animation_sequence(slide, params.sequence_index)
    if params.animation_index > sequence.Count:
        raise ValueError(f"animation_index must not exceed {sequence.Count}")
    if params.move_to is not None and params.move_to > sequence.Count:
        raise ValueError(f"move_to must not exceed {sequence.Count}")
    if params.trigger == "on_shape_click" and params.sequence_index is None:
        raise ValueError(
            "on_shape_click requires an existing interactive sequence"
        )
    effect = sequence(params.animation_index)
    effect_type = (
        (
            ANIMATION_EFFECT_MAP[params.effect]
            if isinstance(params.effect, str)
            else params.effect
        )
        if params.effect is not None
        else effect.EffectType
    )
    if params.exit and not 1 <= effect_type <= 53:
        raise ValueError(
            "Exit animations require an entrance/exit effect (1-53)"
        )
    if params.effect is not None:
        effect.EffectType = effect_type
    if params.trigger is not None:
        effect.Timing.TriggerType = TRIGGER_MAP[params.trigger]
    effect = _configure_animation(sequence, effect, params)
    if params.move_to is not None:
        effect.MoveTo(params.move_to)
    result = _animation_info(effect)
    result.update(
        success=True,
        animation_index=effect.Index,
        delay=effect.Timing.TriggerDelayTime,
    )
    return result
