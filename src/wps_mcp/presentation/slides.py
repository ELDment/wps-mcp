"""Create and organize slides, navigate decks and edit speaker notes."""

from __future__ import annotations

import logging
from typing import Annotated, Self

from pydantic import Field, model_validator

from wps_mcp.common.colors import RgbColor
from wps_mcp.models import DocumentInput, JsonObject, NativeObject
from wps_mcp.native.objects import is_optional_member_error
from wps_mcp.native.runtime import get_session
from wps_mcp.presentation.constants import MSO_FALSE, MSO_TRUE, PP_LAYOUT_BLANK
from wps_mcp.presentation.documents import (
    PresentationInput,
    resolve_presentation,
)
from wps_mcp.presentation.navigation import goto_slide as nav_goto_slide
from wps_mcp.presentation.text.formatting import apply_font_format
from wps_mcp.presentation.text.models import TextFormatSpec
from wps_mcp.protocol.definitions import tool


class SetSlideHiddenInput(DocumentInput):
    """Input for setting slide visibility."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    hidden: bool = Field(..., description="True to hide, False to show")


@tool(
    "wps_ppt_slides",
    component="wpp",
    annotations={"title": "Set Slide Hidden", "idempotentHint": True},
    operation="set_hidden",
)
def set_slide_hidden(params: SetSlideHiddenInput) -> JsonObject:
    """Set a slide as hidden or visible in the slideshow."""
    nav_goto_slide(params.slide_index)
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    slide.SlideShowTransition.Hidden = MSO_TRUE if params.hidden else MSO_FALSE
    result = {
        "success": True,
        "slide_index": params.slide_index,
        "hidden": params.hidden,
    }
    return result


logger = logging.getLogger(__name__)
LAYOUT_NAME_MAP = {
    "title": 1,
    "text": 2,
    "two_column_text": 3,
    "table": 4,
    "title_only": 11,
    "blank": 12,
    "section_header": 33,
    "comparison": 34,
    "content_with_caption": 35,
    "picture_with_caption": 36,
}
SlideIndex = Annotated[int, Field(ge=1)]


class AddSlideInput(DocumentInput):
    """Choose a built-in, custom or inherited slide layout."""

    position: int | None = Field(default=None, ge=1)
    layout: int | str | None = None
    layout_name: str | None = Field(default=None, min_length=1)
    design_index: int | None = Field(default=None, ge=1)
    like_slide_index: int | None = Field(default=None, ge=1)
    count: int = Field(default=1, ge=1)

    @model_validator(mode="after")
    def normalize_layout(self) -> Self:
        """Normalize numeric layout strings and resolve name aliases."""
        if isinstance(self.layout, str):
            value = self.layout.strip()
            if not value:
                raise ValueError(
                    "layout must be a layout name or numeric constant"
                )
            if value.lstrip("+-").isdigit():
                self.layout = int(value)
            else:
                self.layout_name = self.layout_name or value
                self.layout = None
        return self


class DeleteSlideInput(DocumentInput):
    """Delete one slide, an explicit list or an inclusive range."""

    slide_index: int | None = Field(default=None, ge=1)
    slide_indices: list[SlideIndex] | None = Field(default=None, min_length=1)
    from_index: int | None = Field(default=None, ge=1)
    to_index: int | None = Field(default=None, ge=1)

    @model_validator(mode="after")
    def validate_selection(self) -> Self:
        """Require exactly one slide selector or an ascending range."""
        ranged = self.from_index is not None or self.to_index is not None
        if (
            sum(
                (
                    self.slide_index is not None,
                    self.slide_indices is not None,
                    ranged,
                )
            )
            != 1
        ):
            raise ValueError(
                "Provide slide_index, slide_indices or from_index+to_index"
            )
        if ranged:
            if self.from_index is None or self.to_index is None:
                raise ValueError("Both from_index and to_index are required")
            if self.from_index > self.to_index:
                raise ValueError("from_index must not exceed to_index")
        return self


class SlideSelectionInput(DocumentInput):
    """Select one slide or a nonempty list of slides."""

    slide_index: int | None = Field(default=None, ge=1)
    slide_indices: list[SlideIndex] | None = Field(default=None, min_length=1)

    @model_validator(mode="after")
    def validate_selection(self) -> Self:
        """Require one single-slide or multiple-slide selector."""
        if (self.slide_index is None) == (self.slide_indices is None):
            raise ValueError(
                "Provide exactly one of slide_index or slide_indices"
            )
        return self


class DuplicateSlideInput(DocumentInput):
    """Duplicate a slide repeatedly; insert_at=-1 appends."""

    slide_index: int = Field(ge=1)
    insert_at: int | None = None
    count: int = Field(default=1, ge=1)

    @model_validator(mode="after")
    def validate_position(self) -> Self:
        """Accept a positive insertion index or append sentinel."""
        if (
            self.insert_at is not None
            and self.insert_at != -1
            and (self.insert_at < 1)
        ):
            raise ValueError("insert_at must be positive or -1")
        return self


class MoveSlideInput(SlideSelectionInput):
    """Move slides as a contiguous block, preserving their order."""

    new_position: int = Field(ge=1)


class CopySlideInput(SlideSelectionInput):
    """Copy slides in order within or between open presentations."""

    source_presentation_index: int | None = Field(default=None, ge=1)
    source_presentation_name: str | None = Field(default=None, min_length=1)
    to_presentation_index: int | None = Field(default=None, ge=1)
    to_presentation_name: str | None = Field(default=None, min_length=1)
    insert_at: int | None = None

    @model_validator(mode="after")
    def validate_targets(self) -> Self:
        """Reject conflicting deck selectors and insertion positions."""
        if (
            self.insert_at is not None
            and self.insert_at != -1
            and (self.insert_at < 1)
        ):
            raise ValueError("insert_at must be positive or -1")
        if (
            self.source_presentation_index is not None
            and self.source_presentation_name is not None
        ):
            raise ValueError(
                "Specify a source presentation index or name, not both"
            )
        if (
            self.to_presentation_index is not None
            and self.to_presentation_name is not None
        ):
            raise ValueError(
                "Specify a destination presentation index or name, not both"
            )
        return self


class SlideInput(DocumentInput):
    """Select a slide by its one-based index."""

    slide_index: int = Field(ge=1)


class SetSlideNotesInput(SlideInput):
    """Update notes text and optional printed notes formatting."""

    notes_text: str | None = None
    font_name: str | None = Field(default=None, min_length=1)
    font_name_fareast: str | None = Field(default=None, min_length=1)
    font_size: float | None = Field(default=None, gt=0)
    bold: bool | None = None
    italic: bool | None = None
    color: RgbColor | None = None


@tool(
    "wps_ppt_slides",
    component="wpp",
    annotations={"title": "Add Slide"},
    operation="add",
)
def add_slide(params: AddSlideInput) -> JsonObject:
    """Add a new slide to the active presentation.

    To add a slide that looks just like an existing one, pass
    like_slide_index — it inherits that slide's exact design + layout
    with no layout-name ambiguity (this only copies the look, not the
    content; use wps_ppt_slides/duplicate for a full copy). This is the
    recommended path.

    Otherwise pass layout_name (a friendly name such as 'blank',
    'title', 'title_only', or a custom layout name from the slide
    master) or layout (a PpSlideLayout integer constant; a name passed
    here is treated like layout_name). Position is 1-based; omit to
    append at the end. Use design_index to pick a layout from a specific
    slide master/design. Use count to create multiple slides at once;
    when count > 1, returns a slides list instead of
    a single slide_index.

    The response includes the actually-applied layout_name and
    design_name. If layout_name matched layouts in multiple designs, the
    response also includes layout_ambiguous=true and a warning naming
    the candidates.
    """
    position = params.position
    app = get_session("wpp").get_application()
    pres = resolve_presentation(app, document=params.document)
    if position is None:
        position = pres.Slides.Count + 1
    if position < 1 or position > pres.Slides.Count + 1:
        raise ValueError(
            f"Position {position} out of range (1-{pres.Slides.Count + 1})"
        )
    use_custom_layout = False
    custom_layout = None
    friendly_layout = None
    layout_val = None
    resolved_layout_name = None
    resolved_design_name = None
    layout_ambiguous = False
    ambiguous_designs = None
    if params.like_slide_index is not None:
        if (
            params.like_slide_index < 1
            or params.like_slide_index > pres.Slides.Count
        ):
            raise ValueError(
                f"like_slide_index {params.like_slide_index} "
                f"out of range (1-{pres.Slides.Count})"
            )
        src_slide = pres.Slides(params.like_slide_index)
        custom_layout = src_slide.CustomLayout
        use_custom_layout = True
        resolved_layout_name = custom_layout.Name
        try:
            resolved_design_name = src_slide.Design.Name
        except Exception as error:
            if not is_optional_member_error(error):
                raise
            resolved_design_name = None
    elif params.layout_name:
        friendly_key = params.layout_name.lower().strip().replace(" ", "_")
        if friendly_key in LAYOUT_NAME_MAP:
            friendly_layout = LAYOUT_NAME_MAP[friendly_key]
        else:
            if params.design_index is not None and (
                params.design_index < 1
                or params.design_index > pres.Designs.Count
            ):
                raise ValueError(
                    f"Design index {params.design_index} "
                    f"out of range (1-{pres.Designs.Count})"
                )
            matches: list[tuple[int, str, NativeObject]] = []
            search = (
                [params.design_index]
                if params.design_index is not None
                else range(1, pres.Designs.Count + 1)
            )
            available: dict[str, list[str]] = {}
            for index in search:
                design = pres.Designs(index)
                layouts = design.SlideMaster.CustomLayouts
                available[design.Name] = []
                for layout_index in range(1, layouts.Count + 1):
                    candidate = layouts(layout_index)
                    available[design.Name].append(candidate.Name)
                    if candidate.Name == params.layout_name:
                        matches.append((index, design.Name, candidate))
            if not matches:
                raise ValueError(
                    f"Layout '{params.layout_name}' not found. "
                    f"Available custom layouts by design: {available}"
                )
            chosen = matches[0]
            if len(matches) > 1:
                try:
                    default_master_name = pres.SlideMaster.Name
                    for mt in matches:
                        if (
                            pres.Designs(mt[0]).SlideMaster.Name
                            == default_master_name
                        ):
                            chosen = mt
                            break
                except Exception as error:
                    if not is_optional_member_error(error):
                        raise
                    logger.debug(
                        "Optional native presentation "
                        "property unavailable: %s",
                        error,
                    )
                layout_ambiguous = True
                ambiguous_designs = [m[1] for m in matches]
            custom_layout = chosen[2]
            resolved_design_name = chosen[1]
            use_custom_layout = True
            resolved_layout_name = custom_layout.Name
    else:
        layout_val = (
            params.layout if params.layout is not None else PP_LAYOUT_BLANK
        )
    created_slides = []
    for i in range(params.count):
        insert_pos = position + i
        if use_custom_layout:
            slide = pres.Slides.AddSlide(
                Index=insert_pos, pCustomLayout=custom_layout
            )
        elif friendly_layout is not None:
            slide = pres.Slides.Add(Index=insert_pos, Layout=friendly_layout)
        else:
            slide = pres.Slides.Add(Index=insert_pos, Layout=layout_val)
        created_slides.append(
            {"slide_index": slide.SlideIndex, "slide_id": slide.SlideID}
        )
    nav_goto_slide(created_slides[-1]["slide_index"])
    first_slide = pres.Slides(created_slides[0]["slide_index"])
    final_layout_name = resolved_layout_name
    final_design_name = resolved_design_name
    try:
        if final_layout_name is None:
            final_layout_name = first_slide.CustomLayout.Name
        if final_design_name is None:
            final_design_name = first_slide.Design.Name
    except Exception as error:
        if not is_optional_member_error(error):
            raise
        logger.debug(
            "Optional native presentation property unavailable: %s", error
        )
    result = {
        "success": True,
        "slides_created": len(created_slides),
        "slides": created_slides,
        "layout": first_slide.Layout,
        "layout_name": final_layout_name,
        "design_name": final_design_name,
    }
    if layout_ambiguous:
        result["layout_ambiguous"] = True
        result["warnings"] = [
            f"Layout '{params.layout_name}' exists in multiple designs "
            f"({', '.join(ambiguous_designs or [])}); "
            f"used design '{final_design_name}'. "
            "Pass design_index or like_slide_index to select a specific one."
        ]
    if params.count == 1:
        result["slide_index"] = created_slides[0]["slide_index"]
        result["slide_id"] = created_slides[0]["slide_id"]
    return result


@tool(
    "wps_ppt_slides",
    component="wpp",
    annotations={"title": "Delete Slide", "destructiveHint": True},
    operation="delete",
)
def delete_slide(params: DeleteSlideInput) -> JsonObject:
    """Delete one or more slides in a single call.

    Provide exactly one of:
      - slide_index: a single 1-based index, or
      - slide_indices: a list of 1-based indices, or
      - from_index + to_index: an inclusive 1-based range.

    For the list/range forms, slides are removed highest-index-first
    internally, so every index you pass refers to the CURRENT numbering
    — no manual bookkeeping for index shifting. Deleting all slides is
    rejected (a presentation must keep at least one).
    """
    app = get_session("wpp").get_application()
    pres = resolve_presentation(app, document=params.document)
    total = pres.Slides.Count
    if params.slide_index is not None:
        targets = [params.slide_index]
    elif params.slide_indices is not None:
        targets = list(params.slide_indices or [])
    else:
        assert params.from_index is not None and params.to_index is not None
        targets = list(range(params.from_index, params.to_index + 1))
    targets = sorted(set(targets))
    if not targets:
        raise ValueError("No slides to delete")
    out_of_range = [i for i in targets if i < 1 or i > total]
    if out_of_range:
        raise ValueError(
            f"Slide index(es) {out_of_range} out of range (1-{total})"
        )
    if len(targets) >= total:
        raise ValueError(
            "Cannot delete every slide. "
            "A presentation must keep at least one slide "
            f"(requested {len(targets)} of {total})"
        )
    nav_goto_slide(targets[0])
    for idx in sorted(targets, reverse=True):
        pres.Slides(idx).Delete()
    result = {
        "success": True,
        "deleted_indices": targets,
        "deleted_count": len(targets),
        "remaining_count": pres.Slides.Count,
    }
    if params.slide_index is not None:
        result["deleted_index"] = params.slide_index
    return result


@tool(
    "wps_ppt_slides",
    component="wpp",
    annotations={"title": "Duplicate Slide"},
    operation="duplicate",
)
def duplicate_slide(params: DuplicateSlideInput) -> JsonObject:
    """Duplicate a slide with optional position and copy count.

    - insert_at omitted: copies go immediately after the source
      (historical behavior).
    - insert_at = -1: append at the end.
    - insert_at = N: the first copy ends up at 1-based position N; extra
      copies (count > 1) follow consecutively.

    This collapses the old "duplicate then move" two-call dance into
    one. Returns new_slide_indices / new_slide_ids (plus new_slide_index
    for a single copy).
    """
    app = get_session("wpp").get_application()
    pres = resolve_presentation(app, document=params.document)
    if params.slide_index < 1 or params.slide_index > pres.Slides.Count:
        raise ValueError(
            f"Slide index {params.slide_index} "
            f"out of range (1-{pres.Slides.Count})"
        )
    if (
        params.insert_at is not None
        and params.insert_at != -1
        and (params.insert_at > pres.Slides.Count + 1)
    ):
        raise ValueError(
            f"insert_at {params.insert_at} out "
            f"of range (1-{pres.Slides.Count + 1})"
        )
    nav_goto_slide(params.slide_index)
    src_id = pres.Slides(params.slide_index).SlideID
    new_indices = []
    new_ids = []
    for i in range(params.count):
        cur_src_idx = pres.Slides.FindBySlideID(src_id).SlideIndex
        new_slide = pres.Slides(cur_src_idx).Duplicate()(1)
        new_id = new_slide.SlideID
        if params.insert_at is None:
            target = cur_src_idx + 1 + i
        elif params.insert_at == -1:
            target = pres.Slides.Count
        else:
            target = min(params.insert_at + i, pres.Slides.Count)
        if new_slide.SlideIndex != target:
            new_slide.MoveTo(target)
        new_indices.append(new_slide.SlideIndex)
        new_ids.append(new_id)
    nav_goto_slide(new_indices[-1])
    result = {
        "success": True,
        "count": len(new_indices),
        "new_slide_indices": new_indices,
        "new_slide_ids": new_ids,
    }
    if params.count == 1:
        result["new_slide_index"] = new_indices[0]
        result["new_slide_id"] = new_ids[0]
    return result


@tool(
    "wps_ppt_slides",
    component="wpp",
    annotations={"title": "Move Slide", "idempotentHint": True},
    operation="move",
)
def move_slide(params: MoveSlideInput) -> JsonObject:
    """Move slides to a new position within the presentation.

    Provide exactly one of:
      - slide_index: a single slide, or
      - slide_indices: a block of slides (relative order preserved).

    new_position is the 1-based destination of the first moved slide; a
    block ends up contiguous at new_position..new_position+k-1. All
    indices refer to the CURRENT numbering — the move is anchored on
    slide IDs so index shifting is handled internally.
    """
    app = get_session("wpp").get_application()
    pres = resolve_presentation(app, document=params.document)
    total = pres.Slides.Count
    sources = (
        [params.slide_index]
        if params.slide_index is not None
        else list(params.slide_indices or [])
    )
    sources = sorted(set(sources))
    out_of_range = [i for i in sources if i < 1 or i > total]
    if out_of_range:
        raise ValueError(
            f"Slide index(es) {out_of_range} out of range (1-{total})"
        )
    k = len(sources)
    max_start = total - k + 1
    if params.new_position < 1 or params.new_position > max_start:
        raise ValueError(
            f"new_position {params.new_position} out of range (1-{max_start}) "
            f"for a block of {k} slide(s)"
        )
    nav_goto_slide(sources[0])
    all_ids = [pres.Slides(i).SlideID for i in range(1, total + 1)]
    src_ids = [all_ids[i - 1] for i in sources]
    selected = set(src_ids)
    others = [slide_id for slide_id in all_ids if slide_id not in selected]
    final_ids = (
        others[: params.new_position - 1]
        + src_ids
        + others[params.new_position - 1 :]
    )
    for f in range(1, total + 1):
        cur_idx = pres.Slides.FindBySlideID(final_ids[f - 1]).SlideIndex
        if cur_idx != f:
            pres.Slides(cur_idx).MoveTo(toPos=f)
    result = {
        "success": True,
        "moved_slide_ids": src_ids,
        "moved_count": k,
        "new_start_position": params.new_position,
    }
    if params.slide_index is not None:
        result["moved_from"] = params.slide_index
        result["moved_to"] = params.new_position
    return result


@tool(
    "wps_ppt_slides",
    component="wpp",
    annotations={"title": "Copy Slide"},
    operation="copy",
)
def copy_slide(params: CopySlideInput) -> JsonObject:
    """Copy slides within or between open presentations.

    Provide exactly one of slide_index or slide_indices for the source.
    By default the source and destination are the active presentation;
    set source_presentation_* / to_presentation_* (by index or name) to
    copy between different open files — e.g. pull layout slides from a
    template deck into the deck you are building.

    insert_at is the 1-based destination position for the first copy;
    omit it (or use -1) to append at the end. Copies are made via the
    clipboard, which preserves formatting and carries the source design
    across files.

    Both decks must be open (use wps_documents/open with component=wpp).
    Returns new_slide_indices and the destination presentation name.
    """
    app = get_session("wpp").get_application()
    src_pres = resolve_presentation(
        app,
        presentation_index=params.source_presentation_index,
        presentation_name=params.source_presentation_name,
        document=(
            params.document
            if params.source_presentation_index is None
            and params.source_presentation_name is None
            else None
        ),
        enforce_target=False,
    )
    dst_pres = resolve_presentation(
        app,
        presentation_index=params.to_presentation_index,
        presentation_name=params.to_presentation_name,
        document=params.document,
    )
    sources = (
        [params.slide_index]
        if params.slide_index is not None
        else list(params.slide_indices or [])
    )
    src_total = src_pres.Slides.Count
    out_of_range = [i for i in sources if i < 1 or i > src_total]
    if out_of_range:
        raise ValueError(
            f"Source slide index(es) {out_of_range} "
            f"out of range (1-{src_total})"
        )
    append = params.insert_at is None or params.insert_at == -1
    if not append:
        assert params.insert_at is not None
        if (
            params.insert_at < 1
            or params.insert_at > dst_pres.Slides.Count + 1
        ):
            raise ValueError(
                f"insert_at {params.insert_at} out of "
                f"range (1-{dst_pres.Slides.Count + 1})"
            )
    src_ids = [src_pres.Slides(i).SlideID for i in sources]
    new_indices = []
    for j, sid in enumerate(src_ids):
        src_pres.Slides.FindBySlideID(sid).Copy()
        if (
            params.insert_at is None
            or append
            or params.insert_at + j > dst_pres.Slides.Count
        ):
            rng = dst_pres.Slides.Paste()
        else:
            rng = dst_pres.Slides.Paste(params.insert_at + j)
        new_indices.append(rng(1).SlideIndex)
    try:
        if dst_pres.Windows.Count:
            dst_pres.Windows(1).View.GotoSlide(new_indices[-1])
    except Exception as error:
        if not is_optional_member_error(error):
            raise
        logger.debug(
            "Optional native presentation property unavailable: %s", error
        )
    return {
        "success": True,
        "copied_count": len(new_indices),
        "new_slide_indices": new_indices,
        "target_presentation": dst_pres.Name,
        "source_presentation": src_pres.Name,
    }


@tool(
    "wps_ppt_slides",
    component="wpp",
    annotations={
        "title": "List Slides",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="list",
)
def list_slides(params: PresentationInput) -> JsonObject:
    """List all slides in the active or specified presentation.

    Returns each slide's index, ID, name, layout, hidden status, shape
    count, and whether it has speaker notes.
    """
    app = get_session("wpp").get_application()
    pres = resolve_presentation(
        app,
        presentation_index=params.presentation_index,
        presentation_name=params.presentation_name,
        document=params.document,
    )
    slides = []
    for i in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(i)
        layout_name = ""
        try:
            layout_name = slide.CustomLayout.Name
        except Exception as error:
            if not is_optional_member_error(error):
                raise
            logger.debug(
                "Optional native presentation property unavailable: %s", error
            )
        has_notes = False
        try:
            notes_text = slide.NotesPage.Shapes.Placeholders(
                2
            ).TextFrame.TextRange.Text
            has_notes = len(notes_text.strip()) > 0
        except Exception as error:
            if not is_optional_member_error(error):
                raise
            logger.debug(
                "Optional native presentation property unavailable: %s", error
            )
        slides.append(
            {
                "index": slide.SlideIndex,
                "slide_id": slide.SlideID,
                "name": slide.Name,
                "layout": slide.Layout,
                "layout_name": layout_name,
                "hidden": bool(slide.SlideShowTransition.Hidden),
                "shapes_count": slide.Shapes.Count,
                "has_notes": has_notes,
            }
        )
    return {"slides_count": pres.Slides.Count, "slides": slides}


@tool(
    "wps_ppt_slides",
    component="wpp",
    annotations={
        "title": "Get Slide Info",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="get",
)
def get_slide_info(params: SlideInput) -> JsonObject:
    """Get detailed information about a specific slide.

    Returns layout, shapes count, title text, speaker notes, transition
    settings, background info, and design name.
    """
    app = get_session("wpp").get_application()
    pres = resolve_presentation(app, document=params.document)
    if params.slide_index < 1 or params.slide_index > pres.Slides.Count:
        raise ValueError(
            f"Slide index {params.slide_index} "
            f"out of range (1-{pres.Slides.Count})"
        )
    slide = pres.Slides(params.slide_index)
    trans = slide.SlideShowTransition
    layout_name = ""
    try:
        layout_name = slide.CustomLayout.Name
    except Exception as error:
        if not is_optional_member_error(error):
            raise
        logger.debug(
            "Optional native presentation property unavailable: %s", error
        )
    title_text = ""
    has_title = bool(slide.Shapes.HasTitle)
    if has_title:
        try:
            title_text = slide.Shapes.Title.TextFrame.TextRange.Text
        except Exception as error:
            if not is_optional_member_error(error):
                raise
            logger.debug(
                "Optional native presentation property unavailable: %s", error
            )
    notes_text = ""
    try:
        notes_text = slide.NotesPage.Shapes.Placeholders(
            2
        ).TextFrame.TextRange.Text
    except Exception as error:
        if not is_optional_member_error(error):
            raise
        logger.debug(
            "Optional native presentation property unavailable: %s", error
        )
    design_name = ""
    try:
        design_name = slide.Design.Name
    except Exception as error:
        if not is_optional_member_error(error):
            raise
        logger.debug(
            "Optional native presentation property unavailable: %s", error
        )
    return {
        "index": slide.SlideIndex,
        "slide_id": slide.SlideID,
        "slide_number": slide.SlideNumber,
        "name": slide.Name,
        "layout": slide.Layout,
        "layout_name": layout_name,
        "hidden": bool(trans.Hidden),
        "shapes_count": slide.Shapes.Count,
        "has_title": has_title,
        "title_text": title_text,
        "notes_text": notes_text,
        "follow_master_background": bool(slide.FollowMasterBackground),
        "transition_effect": trans.EntryEffect,
        "advance_on_click": bool(trans.AdvanceOnClick),
        "advance_on_time": bool(trans.AdvanceOnTime),
        "advance_time": trans.AdvanceTime,
        "design_name": design_name,
    }


@tool(
    "wps_ppt_slides",
    component="wpp",
    annotations={"title": "Set Slide Notes", "idempotentHint": True},
    operation="set_notes",
)
def set_slide_notes(params: SetSlideNotesInput) -> JsonObject:
    """Set the speaker notes text and/or formatting for a slide.

    Replaces any existing notes with the provided text. Optionally
    applies font formatting (font_name, font_size, bold, italic, color).
    If notes_text is omitted, only formatting is
    applied to existing notes.

    **Important**: Font formatting affects printed notes pages and PDF
    export only. The Notes pane and Presenter View ignore these settings
    (Presenter View has its own A+/A- zoom controls).
    """
    formatting = TextFormatSpec.model_validate(
        params.model_dump(exclude={"document", "slide_index", "notes_text"})
    )
    app = get_session("wpp").get_application()
    pres = resolve_presentation(app, document=params.document)
    if params.slide_index < 1 or params.slide_index > pres.Slides.Count:
        raise ValueError(
            f"Slide index {params.slide_index} "
            f"out of range (1-{pres.Slides.Count})"
        )
    slide = pres.Slides(params.slide_index)
    text_range = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange
    if params.notes_text is not None:
        text_range.Text = params.notes_text
    apply_font_format(text_range.Font, formatting)
    return {"success": True}


@tool(
    "wps_ppt_slides",
    component="wpp",
    annotations={
        "title": "Get Slide Notes",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="notes",
)
def get_slide_notes(params: SlideInput) -> JsonObject:
    """Get the speaker notes text for a slide.

    Returns the notes text. If no notes exist, returns an empty string.
    """
    app = get_session("wpp").get_application()
    pres = resolve_presentation(app, document=params.document)
    if params.slide_index < 1 or params.slide_index > pres.Slides.Count:
        raise ValueError(
            f"Slide index {params.slide_index} "
            f"out of range (1-{pres.Slides.Count})"
        )
    slide = pres.Slides(params.slide_index)
    try:
        notes_text = slide.NotesPage.Shapes.Placeholders(
            2
        ).TextFrame.TextRange.Text
    except Exception as error:
        if not is_optional_member_error(error):
            raise
        notes_text = ""
    return {"slide_index": params.slide_index, "notes_text": notes_text}


@tool(
    "wps_ppt_slides",
    component="wpp",
    annotations={"title": "Go To Slide", "idempotentHint": True},
    operation="goto",
)
def goto_slide(params: SlideInput) -> JsonObject:
    """Navigate the active window to display a specific slide.

    Changes which slide is shown in the WPS Presentation editor. Useful
    for jumping to a slide you want to view or edit.
    """
    pres = get_session("wpp").get_document(params.document)
    if params.slide_index < 1 or params.slide_index > pres.Slides.Count:
        raise ValueError(
            f"Slide index {params.slide_index} "
            f"out of range (1-{pres.Slides.Count})"
        )
    window = pres.Windows(1) if pres.Windows.Count else None
    if window is None:
        raise RuntimeError(
            "The target presentation has no window to "
            "navigate (opened with with_window=False?)."
        )
    window.View.GotoSlide(params.slide_index)
    return {"success": True, "active_slide_index": params.slide_index}
