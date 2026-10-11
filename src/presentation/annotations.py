"""Slide comments and shape hyperlinks."""

from typing import Literal

from pydantic import Field

from src.models import DocumentInput, JsonObject
from src.native.runtime import get_session
from src.presentation.constants import (
    PP_ACTION_HYPERLINK,
    PP_ACTION_NONE,
    PP_MOUSE_CLICK,
    PP_MOUSE_OVER,
)
from src.presentation.shape_lookup import resolve_shape as _get_shape
from src.protocol.definitions import tool


class AddCommentInput(DocumentInput):
    """Input for adding a comment to a slide."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    text: str = Field(..., description="Comment text")
    author: str = Field(default="AI Agent", description="Comment author name")
    author_initials: str = Field(
        default="AI", description="Comment author initials"
    )
    left: float = Field(default=0, description="Horizontal position in points")
    top: float = Field(default=0, description="Vertical position in points")


class ListCommentsInput(DocumentInput):
    """Input for listing comments on a slide."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")


class DeleteCommentInput(DocumentInput):
    """Input for deleting a comment from a slide."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    comment_index: int = Field(..., ge=1, description="1-based comment index")


@tool(
    "wps_ppt_annotations",
    component="wpp",
    annotations={"title": "Add Slide Comment"},
    operation="add_comment",
)
def add_comment(params: AddCommentInput) -> JsonObject:
    """Add a comment to a slide.

    Provide text, author name, and optional position. Uses the WPS
    Presentation Comments.Add API.
    """
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    slide.Comments.Add(
        params.left,
        params.top,
        params.author,
        params.author_initials,
        params.text,
    )
    return {"success": True, "text": params.text, "author": params.author}


@tool(
    "wps_ppt_annotations",
    component="wpp",
    annotations={
        "title": "List Slide Comments",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="comments",
)
def list_comments(params: ListCommentsInput) -> JsonObject:
    """List all comments on a slide.

    Returns index, author, text, datetime, and
    position for each comment.
    """
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    comments_col = slide.Comments
    count = comments_col.Count
    comments = []
    for i in range(1, count + 1):
        comment = comments_col(i)
        try:
            dt_str = str(comment.DateTime)
        except Exception:
            dt_str = None
        comments.append(
            {
                "index": i,
                "author": comment.Author,
                "author_initials": comment.AuthorInitials,
                "text": comment.Text,
                "date_time": dt_str,
                "left": round(comment.Left, 2),
                "top": round(comment.Top, 2),
            }
        )
    return {
        "slide_index": params.slide_index,
        "comments_count": count,
        "comments": comments,
    }


@tool(
    "wps_ppt_annotations",
    component="wpp",
    annotations={"title": "Delete Slide Comment", "destructiveHint": True},
    operation="delete_comment",
)
def delete_comment(params: DeleteCommentInput) -> JsonObject:
    """Delete a comment from a slide by its 1-based index.

    Use wps_ppt_annotations/comments to find the comment index first.
    """
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    comments_col = slide.Comments
    if params.comment_index < 1 or params.comment_index > comments_col.Count:
        raise ValueError(
            f"Comment index {params.comment_index} "
            f"out of range (1-{comments_col.Count})"
        )
    comments_col(params.comment_index).Delete()
    return {"success": True}


ACTION_ON_MAP: dict[str, int] = {
    "click": PP_MOUSE_CLICK,
    "mouseover": PP_MOUSE_OVER,
}


class AddHyperlinkInput(DocumentInput):
    """Input for adding a hyperlink to a shape."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: str | int = Field(
        ...,
        description="Shape name or one-based index; "
        "names remain stable after insertions.",
    )
    address: str = Field(
        ..., description="Hyperlink URL, file path, or mailto: address"
    )
    sub_address: str | None = Field(
        default=None,
        description="Sub-address for slide "
        "links (e.g. '3,,' to link to slide 3)",
    )
    screen_tip: str | None = Field(
        default=None, description="Tooltip text shown on hover"
    )
    action_on: Literal["click", "mouseover"] = Field(
        default="click", description="Trigger action: 'click' or 'mouseover'"
    )


class GetHyperlinksInput(DocumentInput):
    """Input for listing hyperlinks on a slide."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")


class RemoveHyperlinkInput(DocumentInput):
    """Input for removing a hyperlink from a shape."""

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: str | int = Field(
        ...,
        description="Shape name or one-based index; "
        "names remain stable after insertions.",
    )
    action_on: Literal["click", "mouseover"] = Field(
        default="click",
        description="Which action to remove: 'click' or 'mouseover'",
    )


@tool(
    "wps_ppt_annotations",
    component="wpp",
    annotations={"title": "Add Hyperlink", "idempotentHint": True},
    operation="add_link",
)
def add_hyperlink(params: AddHyperlinkInput) -> JsonObject:
    """Add a hyperlink to a shape.

    Supports URLs, file paths, mailto: links, and slide links (via
    sub_address). Set action_on to 'click' (default) or 'mouseover' for
    the trigger type. For slide links, use sub_address like '3,,' to
    link to slide 3.
    """
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    shape = _get_shape(slide, params.shape_name_or_index)
    action_key = params.action_on
    action_idx = ACTION_ON_MAP[action_key]
    action_setting = shape.ActionSettings(action_idx)
    action_setting.Action = PP_ACTION_HYPERLINK
    action_setting.Hyperlink.Address = params.address
    if params.sub_address is not None:
        action_setting.Hyperlink.SubAddress = params.sub_address
    if params.screen_tip is not None:
        action_setting.Hyperlink.ScreenTip = params.screen_tip
    return {
        "success": True,
        "shape_name": shape.Name,
        "address": params.address,
        "sub_address": params.sub_address,
        "action_on": action_key,
    }


@tool(
    "wps_ppt_annotations",
    component="wpp",
    annotations={
        "title": "Get Hyperlinks",
        "readOnlyHint": True,
        "idempotentHint": True,
    },
    operation="links",
)
def get_hyperlinks(params: GetHyperlinksInput) -> JsonObject:
    """Get all hyperlinks on a slide.

    Returns address, sub-address, and type for each hyperlink found on
    the slide.
    """
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    hyperlinks = []
    for i in range(1, slide.Hyperlinks.Count + 1):
        hl = slide.Hyperlinks(i)
        hyperlinks.append(
            {
                "index": i,
                "address": hl.Address,
                "sub_address": hl.SubAddress,
                "type": hl.Type,
            }
        )
    return {
        "success": True,
        "slide_index": params.slide_index,
        "hyperlinks_count": slide.Hyperlinks.Count,
        "hyperlinks": hyperlinks,
    }


@tool(
    "wps_ppt_annotations",
    component="wpp",
    annotations={"title": "Remove Hyperlink", "idempotentHint": True},
    operation="remove_link",
)
def remove_hyperlink(params: RemoveHyperlinkInput) -> JsonObject:
    """Remove a hyperlink from a shape.

    Clears the click or mouseover action on the specified shape. Set
    action_on to 'click' (default) or 'mouseover' to
    choose which to remove.
    """
    pres = get_session("wpp").get_document()
    slide = pres.Slides(params.slide_index)
    shape = _get_shape(slide, params.shape_name_or_index)
    action_key = params.action_on
    action_idx = ACTION_ON_MAP[action_key]
    shape.ActionSettings(action_idx).Action = PP_ACTION_NONE
    return {"success": True, "shape_name": shape.Name, "action_on": action_key}
