"""Manage Writer bookmarks without altering their anchored text."""

from pydantic import Field

from src.models import DocumentInput, JsonObject
from src.native.runtime import get_session
from src.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)
from src.writer.access import range_details, resolve_range
from src.writer.models import RangeInput


class BookmarkInput(DocumentInput):
    """Identify a bookmark by its validated native name."""

    name: str = Field(min_length=1, max_length=40, pattern="^[^\\W\\d]\\w*$")


class AddBookmarkInput(RangeInput):
    """Anchor a named bookmark to a native text interval."""

    name: str = Field(min_length=1, max_length=40, pattern="^[^\\W\\d]\\w*$")
    replace_existing: bool = False


@tool(
    "wps_writer_bookmarks",
    operation="list",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def read_bookmarks(params: DocumentInput) -> JsonObject:
    """List document bookmarks with their names, text and native ranges."""
    document = get_session("writer").get_document(params.document)
    bookmarks = document.Bookmarks
    result = []
    for index in range(1, int(bookmarks.Count) + 1):
        bookmark = bookmarks.Item(index)
        result.append(
            {"name": str(bookmark.Name), **range_details(bookmark.Range)}
        )
    return {"bookmarks": result}


@tool(
    "wps_writer_bookmarks",
    operation="add",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def add_bookmark(params: AddBookmarkInput) -> JsonObject:
    """Create a named bookmark over a validated UTF-16 interval.

    Set replace_existing explicitly to move an existing bookmark.
    """
    document = get_session("writer").get_document(params.document)
    text_range = resolve_range(document, params.start, params.end)
    if document.Bookmarks.Exists(params.name) and (
        not params.replace_existing
    ):
        raise ValueError(
            "Bookmark already exists; set replace_existing to update its range"
        )
    bookmark = document.Bookmarks.Add(params.name, text_range)
    return {"name": str(bookmark.Name), **range_details(bookmark.Range)}


@tool(
    "wps_writer_bookmarks",
    operation="delete",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def delete_bookmark(params: BookmarkInput) -> JsonObject:
    """Delete a bookmark without deleting its text."""
    document = get_session("writer").get_document(params.document)
    if not document.Bookmarks.Exists(params.name):
        raise ValueError("Bookmark does not exist")
    document.Bookmarks.Item(params.name).Delete()
    return {"name": params.name, "deleted": True}
