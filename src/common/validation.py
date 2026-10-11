"""Suggest meaningful alternatives for unknown input fields."""

import difflib
from collections.abc import Iterable


def suggest(key: str, known: Iterable[str]) -> str | None:
    """Suggest a known field using common aliases and spelling similarity."""
    candidates = list(known)
    sibling = {
        "font_color": "color",
        "line_color": "color",
        "fill_color": "color",
        "line_visible": "visible",
        "line_weight": "weight",
        "fill_transparency": "transparency",
    }.get(key)
    if sibling in candidates:
        return sibling
    close = difflib.get_close_matches(key, candidates, n=1, cutoff=0.6)
    return close[0] if close else None
