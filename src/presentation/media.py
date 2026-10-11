"""Insert media and configure verified playback on the target deck."""

from pathlib import Path
from typing import Literal, Self

from pydantic import Field, model_validator

from src.models import DocumentInput, JsonObject
from src.native.runtime import get_session
from src.presentation.constants import MSO_FALSE, MSO_TRUE
from src.presentation.shape_lookup import resolve_shape
from src.protocol.definitions import tool


class AddMediaInput(DocumentInput):
    """Common audio/video insertion parameters, in points."""

    slide_index: int = Field(ge=1)
    kind: Literal["audio", "video"]
    file_path: str = Field(min_length=1)
    left: float = 100.0
    top: float = 100.0
    width: float | None = Field(default=None, gt=0)
    height: float | None = Field(default=None, gt=0)
    link_to_file: bool = False

    @model_validator(mode="after")
    def validate_dimensions(self) -> Self:
        """Require both media dimensions or omit them together."""
        if (self.width is None) != (self.height is None):
            raise ValueError("Provide both width and height, or omit both")
        return self


class SetMediaSettingsInput(DocumentInput):
    """Playback settings; all times are milliseconds."""

    slide_index: int = Field(ge=1)
    shape_name_or_index: str | int
    volume: float | None = Field(default=None, ge=0, le=1)
    muted: bool | None = None
    start_point: int | None = Field(default=None, ge=0)
    end_point: int | None = Field(default=None, gt=0)
    fade_in: int | None = Field(default=None, ge=0)
    fade_out: int | None = Field(default=None, ge=0)
    loop: bool | None = None
    hide_while_not_playing: bool | None = None

    @model_validator(mode="after")
    def validate_trim(self) -> Self:
        """Require the supplied trim start to precede the trim end."""
        if (
            self.start_point is not None
            and self.end_point is not None
            and self.start_point >= self.end_point
        ):
            raise ValueError("start_point must be before end_point")
        return self


@tool(
    "wps_ppt_media",
    operation="add",
    component="wpp",
    annotations={"openWorldHint": True},
)
def add_media(params: AddMediaInput) -> JsonObject:
    """Use the same WPS insertion contract for audio and video."""
    path = Path(params.file_path).resolve(strict=True)
    if not path.is_file():
        raise ValueError(f"Media path must be a file: {path}")
    slide = get_session("wpp").get_document().Slides(params.slide_index)
    dimensions = () if params.width is None else (params.width, params.height)
    shape = slide.Shapes.AddMediaObject2(
        str(path),
        MSO_TRUE if params.link_to_file else MSO_FALSE,
        MSO_FALSE if params.link_to_file else MSO_TRUE,
        params.left,
        params.top,
        *dimensions,
    )
    return {
        "success": True,
        "shape_name": shape.Name,
        "file_path": str(path),
        "kind": params.kind,
    }


@tool(
    "wps_ppt_media",
    component="wpp",
    annotations={"title": "Set Media Settings", "idempotentHint": True},
    operation="configure",
)
def set_media_settings(params: SetMediaSettingsInput) -> JsonObject:
    """Set playback, fades and verified native trim points."""
    slide = get_session("wpp").get_document().Slides(params.slide_index)
    shape = resolve_shape(slide, params.shape_name_or_index)
    media = shape.MediaFormat
    if params.start_point is not None or params.end_point is not None:
        try:
            current_start, current_end = (
                int(media.StartPoint),
                int(media.EndPoint),
            )
            length = int(media.Length)
        except Exception as error:
            raise RuntimeError(
                "This WPS media object does not expose playback trimming"
            ) from error
        start = (
            current_start if params.start_point is None else params.start_point
        )
        end = current_end if params.end_point is None else params.end_point
        if not 0 <= start < end <= length:
            raise ValueError(
                f"Trim points must satisfy 0 <= start < end <= {length} ms"
            )
        # Expand the interval first, so disjoint old/new ranges remain valid.
        changes = [("StartPoint", start), ("EndPoint", end)]
        if start >= current_end:
            changes.reverse()
        try:
            for attribute, value in changes:
                if getattr(media, attribute) != value:
                    setattr(media, attribute, value)
            if int(media.StartPoint) != start or int(media.EndPoint) != end:
                raise RuntimeError(
                    "WPS did not retain the requested trim points"
                )
        except Exception as error:
            raise RuntimeError(
                "WPS could not apply media trimming; "
                "inspect the current trim before retrying"
            ) from error
    for attribute, setting in (
        ("Volume", params.volume),
        ("Muted", params.muted),
        ("FadeInDuration", params.fade_in),
        ("FadeOutDuration", params.fade_out),
    ):
        if setting is not None:
            setattr(media, attribute, setting)
    if params.loop is not None or params.hide_while_not_playing is not None:
        play = shape.AnimationSettings.PlaySettings
        if params.loop is not None:
            play.LoopUntilStopped = MSO_TRUE if params.loop else MSO_FALSE
        if params.hide_while_not_playing is not None:
            play.HideWhileNotPlaying = (
                MSO_TRUE if params.hide_while_not_playing else MSO_FALSE
            )
    return {"success": True, "shape_name": shape.Name}
