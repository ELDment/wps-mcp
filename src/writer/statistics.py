"""Read native Writer document and structure statistics."""

from src.models import DocumentInput, JsonObject
from src.native.runtime import get_session
from src.protocol.definitions import READ_ANNOTATIONS, tool


@tool(
    "wps_writer_statistics",
    operation="read",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def read_statistics(params: DocumentInput) -> JsonObject:
    """Read native document counts and UTF-16 content boundaries.

    Includes WPS page, word, character, and document structure counts.
    """
    document = get_session("writer").get_document(params.document)
    return {
        "document": str(document.Name),
        "words": int(document.ComputeStatistics(0)),
        "lines": int(document.ComputeStatistics(1)),
        "pages": int(document.ComputeStatistics(2)),
        "characters": int(document.ComputeStatistics(3)),
        "characters_with_spaces": int(document.ComputeStatistics(5)),
        "paragraphs": int(document.Paragraphs.Count),
        "tables": int(document.Tables.Count),
        "inline_images": int(document.InlineShapes.Count),
        "sections": int(document.Sections.Count),
        "range_end": int(document.Content.End),
    }
