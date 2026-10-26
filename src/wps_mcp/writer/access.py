"""Resolve bounded Writer ranges and one-based native collections."""

from wps_mcp.models import JsonObject, NativeObject


def resolve_range(
    document: NativeObject, start: int, end: int | None = None
) -> NativeObject:
    """Resolve a native range without silently clamping an invalid position."""
    content = document.Content
    resolved_end = int(content.End) if end is None else end
    if (
        start < int(content.Start)
        or resolved_end < start
        or resolved_end > int(content.End)
    ):
        raise ValueError(
            f"Range must be within {content.Start}..{content.End} "
            "with start <= end"
        )
    return document.Range(start, resolved_end)


def resolve_item(
    collection: NativeObject, index: int, label: str
) -> NativeObject:
    """Validate a one-based index before retrieving its native object."""
    count = int(collection.Count)
    if index < 1 or index > count:
        raise ValueError(f"{label} index must be between 1 and {count}")
    return collection.Item(index)


def range_details(text_range: NativeObject) -> JsonObject:
    """Detach native UTF-16 boundaries and text into a JSON object."""
    return {
        "start": int(text_range.Start),
        "end": int(text_range.End),
        "text": str(text_range.Text),
    }
