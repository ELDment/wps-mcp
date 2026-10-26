"""Manage Writer revision tracking, revision decisions, and comments."""

from typing import Annotated, Self

from pydantic import Field, model_validator

from wps_mcp.common.dates import date_text
from wps_mcp.models import DocumentInput, JsonObject, NativeObject
from wps_mcp.native.runtime import get_session
from wps_mcp.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)
from wps_mcp.writer.access import range_details, resolve_item, resolve_range
from wps_mcp.writer.models import ExplicitRangeInput, Index


class ReviewListInput(DocumentInput):
    """Paginate a one-based revision or comment collection."""

    start_index: Index = 1
    count: Annotated[int, Field(ge=1, le=1000, strict=True)] = 100


class RevisionTrackingInput(DocumentInput):
    """Enable or disable tracking for subsequent document edits."""

    enabled: bool


class ReviewSelectionInput(DocumentInput):
    """Select unique original review indices or the entire collection."""

    indices: list[Index] | None = Field(
        default=None, min_length=1, max_length=10000
    )
    all: bool = False

    @model_validator(mode="after")
    def validate_selection(self) -> Self:
        """Require unique indices or an explicit entire-collection choice."""
        if self.all == (self.indices is not None):
            raise ValueError("Specify either indices or all=true, exclusively")
        if self.indices is not None and len(set(self.indices)) != len(
            self.indices
        ):
            raise ValueError("indices must not contain duplicates")
        return self


class AddCommentInput(ExplicitRangeInput):
    """Anchor literal comment text and optional reviewer attribution."""

    text: str = Field(min_length=1)
    author: str | None = Field(default=None, min_length=1)
    initials: str | None = None


class UpdateCommentInput(DocumentInput):
    """Change comment text or attribution without moving its anchor."""

    index: Index
    text: str | None = Field(default=None, min_length=1)
    author: str | None = Field(default=None, min_length=1)
    initials: str | None = None

    @model_validator(mode="after")
    def validate_changes(self) -> Self:
        """Reject updates that do not change any supported property."""
        if self.text is None and self.author is None and self.initials is None:
            raise ValueError("Specify text, author or initials to update")
        return self


def select_review_items(
    collection: NativeObject, params: ReviewSelectionInput, label: str
) -> list[tuple[int, NativeObject]]:
    """Resolve original review proxies before any reverse-order edits."""
    count = int(collection.Count)
    indices = list(range(1, count + 1)) if params.all else params.indices or []
    if any(index > count for index in indices):
        raise ValueError(f"{label} index must be between 1 and {count}")
    return [
        (index, collection.Item(index))
        for index in sorted(indices, reverse=True)
    ]


@tool(
    "wps_writer_revisions",
    operation="state",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def revision_state(params: DocumentInput) -> JsonObject:
    """Read native revision tracking and document revision count."""
    document = get_session("writer").get_document(params.document)
    return {
        "enabled": bool(document.TrackRevisions),
        "count": int(document.Revisions.Count),
    }


@tool(
    "wps_writer_revisions",
    operation="configure",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def configure_revisions(params: RevisionTrackingInput) -> JsonObject:
    """Configure tracking for future edits without resolving revisions."""
    document = get_session("writer").get_document(params.document)
    document.TrackRevisions = params.enabled
    return {
        "enabled": bool(document.TrackRevisions),
        "count": int(document.Revisions.Count),
    }


@tool(
    "wps_writer_revisions",
    operation="list",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def list_revisions(params: ReviewListInput) -> JsonObject:
    """List revision metadata with native UTF-16 text boundaries.

    Each item includes its author, date, native type, and text range.
    """
    document = get_session("writer").get_document(params.document)
    revisions = document.Revisions
    count = int(revisions.Count)
    stop = min(count + 1, params.start_index + params.count)
    result: list[JsonObject] = []
    for index in range(params.start_index, stop):
        revision = revisions.Item(index)
        result.append(
            {
                "index": index,
                "author": str(revision.Author),
                "date": date_text(revision.Date),
                "type": int(revision.Type),
                **range_details(revision.Range),
            }
        )
    return {"revisions": result, "total": count, "has_more": stop <= count}


def resolve_revisions(
    params: ReviewSelectionInput, *, accept: bool
) -> JsonObject:
    """Apply a revision decision once without replaying native failures."""
    document = get_session("writer").get_document(params.document)
    revisions = document.Revisions
    if params.all:
        count = int(revisions.Count)
        if accept:
            revisions.AcceptAll()
        else:
            revisions.RejectAll()
    else:
        targets = select_review_items(revisions, params, "revision")
        count = len(targets)
        for _, revision in targets:
            if accept:
                revision.Accept()
            else:
                revision.Reject()
    return {"count": count, "decision": "accepted" if accept else "rejected"}


@tool(
    "wps_writer_revisions",
    operation="accept",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def accept_revisions(params: ReviewSelectionInput) -> JsonObject:
    """Accept selected original revisions or an explicit entire set.

    Set all=true explicitly to select the entire collection.
    """
    return resolve_revisions(params, accept=True)


@tool(
    "wps_writer_revisions",
    operation="reject",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def reject_revisions(params: ReviewSelectionInput) -> JsonObject:
    """Reject selected original revisions or an explicit entire set.

    Set all=true explicitly to select the entire collection.
    """
    return resolve_revisions(params, accept=False)


def comment_details(comment: NativeObject) -> JsonObject:
    """Detach the comment body and its separate document anchor range."""
    return {
        "index": int(comment.Index),
        "text": str(comment.Range.Text),
        "author": str(comment.Author),
        "initials": str(comment.Initial),
        "date": date_text(comment.Date),
        "scope": range_details(comment.Scope),
    }


@tool(
    "wps_writer_comments",
    operation="list",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def list_comments(params: ReviewListInput) -> JsonObject:
    """Paginate native comments with their separate document anchors.

    Use start_index and count to paginate one-based comment indices.
    """
    document = get_session("writer").get_document(params.document)
    comments = document.Comments
    count = int(comments.Count)
    stop = min(count + 1, params.start_index + params.count)
    return {
        "comments": [
            comment_details(comments.Item(index))
            for index in range(params.start_index, stop)
        ],
        "total": count,
        "has_more": stop <= count,
    }


@tool(
    "wps_writer_comments",
    operation="add",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def add_comment(params: AddCommentInput) -> JsonObject:
    """Add a range comment with optional author and initials."""
    document = get_session("writer").get_document(params.document)
    text_range = resolve_range(document, params.start, params.end)
    comment = document.Comments.Add(text_range, params.text)
    if params.author is not None:
        comment.Author = params.author
    if params.initials is not None:
        comment.Initial = params.initials
    return comment_details(comment)


@tool(
    "wps_writer_comments",
    operation="update",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def update_comment(params: UpdateCommentInput) -> JsonObject:
    """Update comment text or attribution while preserving its anchor."""
    document = get_session("writer").get_document(params.document)
    comment = resolve_item(document.Comments, params.index, "comment")
    if params.text is not None:
        comment.Range.Text = params.text
    if params.author is not None:
        comment.Author = params.author
    if params.initials is not None:
        comment.Initial = params.initials
    return comment_details(comment)


@tool(
    "wps_writer_comments",
    operation="delete",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def delete_comments(params: ReviewSelectionInput) -> JsonObject:
    """Delete selected original comments or an explicit entire set.

    Set all=true explicitly to select the entire collection.
    """
    document = get_session("writer").get_document(params.document)
    targets = select_review_items(document.Comments, params, "comment")
    for _, comment in targets:
        comment.Delete()
    return {
        "deleted_count": len(targets),
        "indices": sorted(index for index, _ in targets),
    }
