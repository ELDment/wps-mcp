"""Convert and validate one-based UTF-16 spans for WPS text ranges."""

from __future__ import annotations

import re
from collections.abc import Sequence

from wps_mcp.common.text import normalize_text
from wps_mcp.presentation.text.models import RunSpec, SpanSpec


def utf16_length(value: str) -> int:
    """Count UTF-16 units, preserving existing unpaired surrogates."""
    return len(value.encode("utf-16-le", errors="surrogatepass")) // 2


def find_spans(
    text: str, needle: str, *, match_case: bool, whole_words: bool
) -> list[tuple[int, int]]:
    """Find literal Unicode matches and return one-based UTF-16 spans."""
    if not needle:
        raise ValueError("Search text must not be empty")
    pattern = re.escape(needle)
    if whole_words:
        pattern = rf"(?<!\w){pattern}(?!\w)"
    return [
        (
            utf16_length(text[: match.start()]) + 1,
            utf16_length(match.group()),
        )
        for match in re.finditer(pattern, text, 0 if match_case else re.I)
    ]


def python_span(value: str, start: int, length: int) -> tuple[int, int]:
    """Convert a native span without splitting UTF-16 surrogate pairs."""
    encoded = value.encode("utf-16-le", errors="surrogatepass")
    begin, end = (start - 1) * 2, (start - 1 + length) * 2
    if start < 1 or length < 0 or begin > len(encoded) or end > len(encoded):
        raise ValueError(
            f"Span ({start}, {length}) exceeds text length "
            f"{len(encoded) // 2} UTF-16 units"
        )
    try:
        prefix = encoded[:begin].decode("utf-16-le")
        fragment = encoded[begin:end].decode("utf-16-le")
    except UnicodeDecodeError as error:
        raise ValueError(
            "Span boundaries must not split a UTF-16 surrogate pair"
        ) from error
    return len(prefix), len(prefix) + len(fragment)


def resolve_span(
    full_text: str, shape_name: str, spec: SpanSpec
) -> tuple[int, int]:
    """Resolve a search or explicit span into validated UTF-16 positions."""
    if spec.search_text is not None:
        search_text = normalize_text(spec.search_text)
        position = 0
        for occurrence in range(spec.occurrence):
            position = full_text.find(search_text, position)
            if position == -1:
                raise ValueError(
                    f"{search_text!r}: {occurrence} matches in {shape_name!r}"
                )
            if occurrence < spec.occurrence - 1:
                position += len(search_text)
        return utf16_length(full_text[:position]) + 1, utf16_length(
            search_text
        )
    start, length = spec.start, spec.length
    if start is None or length is None:
        raise ValueError("Provide both start and length")
    python_span(full_text, start, length)
    return start, length


def run_offsets(runs: Sequence[RunSpec]) -> list[tuple[int, int]]:
    """Compute UTF-16 spans after normalizing each run independently."""
    spans = []
    start = 1
    for run in runs:
        length = utf16_length(normalize_text(run.text))
        spans.append((start, length))
        start += length
    return spans
