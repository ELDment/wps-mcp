"""Control only the selected presentation's slide show."""

from typing import Literal, Self

from pydantic import Field, model_validator

from wps_mcp.models import DocumentInput, JsonObject, NativeObject
from wps_mcp.native.objects import capabilities_for
from wps_mcp.native.runtime import get_session
from wps_mcp.native.state import temporary_attributes
from wps_mcp.presentation.constants import (
    MSO_FALSE,
    MSO_TRUE,
    PP_SHOW_ALL,
    PP_SHOW_SLIDE_RANGE,
    PP_SHOW_TYPE_KIOSK,
    PP_SHOW_TYPE_SPEAKER,
    PP_SHOW_TYPE_WINDOW,
    SLIDESHOW_STATE_NAMES,
)
from wps_mcp.protocol.definitions import tool

SHOW_TYPE_MAP = {
    "speaker": PP_SHOW_TYPE_SPEAKER,
    "window": PP_SHOW_TYPE_WINDOW,
    "kiosk": PP_SHOW_TYPE_KIOSK,
}


class SlideShowStartInput(DocumentInput):
    """Optional range and playback mode for a new slide show."""

    start_slide: int | None = Field(default=None, ge=1)
    end_slide: int | None = Field(default=None, ge=1)
    loop: bool | None = None
    show_type: Literal["speaker", "window", "kiosk"] = "speaker"

    @model_validator(mode="after")
    def validate_range(self) -> Self:
        """Require an ascending range when both endpoints exist."""
        if (
            self.start_slide is not None
            and self.end_slide is not None
            and self.start_slide > self.end_slide
        ):
            raise ValueError("start_slide must not exceed end_slide")
        return self


class SlideShowGotoInput(DocumentInput):
    """An absolute slide index in the selected presentation."""

    slide_index: int = Field(ge=1)


def _show_presentation() -> NativeObject:
    """Require a reliable native show lifecycle before any playback action."""
    presentation = get_session("wpp").get_document()
    capabilities_for(presentation).require("slideshow")
    return presentation


def _show_window(required: bool = True) -> NativeObject:
    """Find the slide show belonging to the target deck."""
    session = get_session("wpp")
    presentation = _show_presentation()
    windows = session.get_application().SlideShowWindows
    for index in range(1, windows.Count + 1):
        window = windows(index)
        if window.Presentation.FullName == presentation.FullName:
            return window
    if required:
        raise RuntimeError(
            "No slide show is running for the selected presentation"
        )
    return None


@tool(
    "wps_ppt_slideshow",
    component="wpp",
    annotations={"title": "Start Slide Show"},
    operation="start",
)
def slideshow_start(params: SlideShowStartInput) -> JsonObject:
    """Start a show and restore its stored settings afterwards."""
    presentation = _show_presentation()
    total = presentation.Slides.Count
    start, end = params.start_slide or 1, params.end_slide or total
    if not 1 <= start <= end <= total:
        raise ValueError(
            f"Slide range must satisfy 1 <= start <= end <= {total}"
        )
    changes = {
        "ShowType": SHOW_TYPE_MAP[params.show_type],
        "RangeType": PP_SHOW_ALL,
    }
    if params.start_slide is not None or params.end_slide is not None:
        changes.update(
            RangeType=PP_SHOW_SLIDE_RANGE, StartingSlide=start, EndingSlide=end
        )
    if params.loop is not None:
        changes["LoopUntilStopped"] = MSO_TRUE if params.loop else MSO_FALSE
    with temporary_attributes(presentation.SlideShowSettings, **changes):
        window = presentation.SlideShowSettings.Run()
    return {
        "success": True,
        "show_type": params.show_type,
        "current_slide": window.View.CurrentShowPosition,
        "total_slides": total,
        "start_slide": start,
        "end_slide": end,
    }


@tool(
    "wps_ppt_slideshow",
    component="wpp",
    annotations={"title": "Stop Slide Show", "idempotentHint": True},
    operation="stop",
)
def slideshow_stop(params: DocumentInput) -> JsonObject:
    """Stop the selected presentation's slide show, if it is running."""
    window = _show_window(required=False)
    if window is None:
        return {"success": True, "message": "No slide show was running."}
    window.View.Exit()
    return {"success": True, "message": "Slide show ended."}


@tool(
    "wps_ppt_slideshow",
    component="wpp",
    annotations={"title": "Slide Show Next"},
    operation="next",
)
def slideshow_next(params: DocumentInput) -> JsonObject:
    """Advance to the next animation or slide in the selected show."""
    view = _show_window().View
    view.Next()
    return {
        "success": True,
        "current_slide": view.CurrentShowPosition,
        "state": SLIDESHOW_STATE_NAMES.get(
            view.State, f"unknown({view.State})"
        ),
    }


@tool(
    "wps_ppt_slideshow",
    component="wpp",
    annotations={"title": "Slide Show Previous"},
    operation="previous",
)
def slideshow_previous(params: DocumentInput) -> JsonObject:
    """Return to the previous slide in the selected show."""
    view = _show_window().View
    view.Previous()
    return {
        "success": True,
        "current_slide": view.CurrentShowPosition,
        "state": SLIDESHOW_STATE_NAMES.get(
            view.State, f"unknown({view.State})"
        ),
    }


@tool(
    "wps_ppt_slideshow",
    component="wpp",
    annotations={"title": "Slide Show Go To Slide", "idempotentHint": True},
    operation="goto",
)
def slideshow_goto(params: SlideShowGotoInput) -> JsonObject:
    """Navigate to an absolute slide in the selected show."""
    total = get_session("wpp").get_document().Slides.Count
    if params.slide_index > total:
        raise ValueError(f"slide_index must not exceed {total}")
    view = _show_window().View
    view.GotoSlide(params.slide_index)
    return {
        "success": True,
        "current_slide": view.CurrentShowPosition,
        "state": SLIDESHOW_STATE_NAMES.get(
            view.State, f"unknown({view.State})"
        ),
    }


@tool(
    "wps_ppt_slideshow",
    component="wpp",
    annotations={
        "title": "Get Slide Show Status",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="status",
)
def slideshow_get_status(params: DocumentInput) -> JsonObject:
    """Read the selected presentation's show state."""
    window = _show_window(required=False)
    if window is None:
        return {"running": False}
    view = window.View
    return {
        "running": True,
        "current_slide": view.CurrentShowPosition,
        "state": view.State,
        "state_name": SLIDESHOW_STATE_NAMES.get(
            view.State, f"unknown({view.State})"
        ),
        "pointer_type": view.PointerType,
    }
