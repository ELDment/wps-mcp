"""Inspect slide typography and repair short wrapped lines when requested."""

from __future__ import annotations

import logging
import unicodedata
from contextlib import nullcontext

from src.errors import WpsError
from src.models import JsonObject, NativeObject
from src.native.runtime import get_session
from src.native.state import temporary_attributes
from src.presentation.constants import (
    PP_AUTO_SIZE_NONE,
    PP_AUTO_SIZE_TEXT_TO_FIT_SHAPE,
)
from src.presentation.navigation import goto_slide
from src.presentation.text_measurement import FIT_TOLERANCE_PT
from src.presentation.text_models import CheckTypographyInput
from src.presentation.text_spans import utf16_length
from src.protocol.definitions import tool

logger = logging.getLogger(__name__)

_OFF_SLIDE_TOLERANCE = 0.5


def _char_type(c: str) -> str:
    """Classify a character for word-boundary detection."""
    if c in "、。，．！？）」』】〉》・":
        return "punct_close"
    if c in "（「『【〈《":
        return "punct_open"
    cp = ord(c)
    if cp < 128:
        return "latin" if c.isalpha() else "digit" if c.isdigit() else "ascii"
    if 12448 <= cp <= 12543 or 12784 <= cp <= 12799:
        return "katakana"
    if 12352 <= cp <= 12447:
        return "hiragana"
    if unicodedata.category(c).startswith("Lo"):
        return "kanji"
    return "other"


def _find_best_vbreak(
    prev_line_text: str, widow_text: str
) -> tuple[int, str, str] | None:
    """Choose a soft break and return its Python offset and split lines."""
    candidates = []
    for i in range(1, len(prev_line_text)):
        pt = _char_type(prev_line_text[i - 1])
        ct = _char_type(prev_line_text[i])
        score = 0
        if pt == "punct_close":
            score = 5
        elif ct == "punct_open":
            score = 5
        elif pt != ct:
            if pt == "hiragana" and ct == "katakana":
                score = 4
            elif pt == "hiragana" and ct == "kanji":
                score = 3
            elif pt == "katakana" and ct in ("hiragana", "kanji"):
                score = 3
            elif pt == "kanji" and ct == "hiragana":
                score = 2
            elif pt == "latin" and ct in ("hiragana", "katakana", "kanji"):
                score = 3
            elif pt in ("hiragana", "katakana", "kanji") and ct == "latin":
                score = 3
            else:
                score = 1
        if score > 0:
            candidates.append((i, score))
    if not candidates:
        return None
    min_pos = len(prev_line_text) // 2
    valid = [(p, s) for p, s in candidates if p >= min_pos]
    if not valid:
        valid = [
            (p, s) for p, s in candidates if p >= len(prev_line_text) // 3
        ]
    if not valid:
        return None
    valid.sort(key=lambda x: (x[1], x[0]), reverse=True)
    pos = valid[0][0]
    return (pos, prev_line_text[:pos], prev_line_text[pos:] + widow_text)


def _right_neighbor_gap(shape: NativeObject, slide: NativeObject) -> float:
    """Find the nearest gap to a vertically overlapping shape on the right."""
    s_right = shape.Left + shape.Width
    s_top = shape.Top
    s_bottom = shape.Top + shape.Height
    min_gap = float("inf")
    for j in range(1, slide.Shapes.Count + 1):
        other = slide.Shapes(j)
        if other.Name == shape.Name:
            continue
        if other.Top + other.Height <= s_top or other.Top >= s_bottom:
            continue
        if other.Left > s_right - 1:
            gap = other.Left - s_right
            if gap < min_gap:
                min_gap = gap
    return min_gap


def _off_slide_edges(
    shape: NativeObject, slide_w: float, slide_h: float
) -> list[str]:
    """Which slide edges a shape hangs over, by name, or an empty list."""
    if not slide_w or not slide_h:
        return []
    try:
        left, top = (shape.Left, shape.Top)
        width, height = (shape.Width, shape.Height)
    except Exception:
        return []
    edges = []
    if left < -_OFF_SLIDE_TOLERANCE:
        edges.append("left")
    if top < -_OFF_SLIDE_TOLERANCE:
        edges.append("top")
    if left + width > slide_w + _OFF_SLIDE_TOLERANCE:
        edges.append("right")
    if top + height > slide_h + _OFF_SLIDE_TOLERANCE:
        edges.append("bottom")
    return edges


def _short_lines(
    shape: NativeObject, max_chars: int, max_words: int
) -> list[JsonObject]:
    """Find short wrapped lines and short lines after soft breaks."""
    text_range = shape.TextFrame.TextRange
    issues = []
    for index in range(2, text_range.Lines().Count + 1):
        previous = text_range.Lines(index - 1)
        previous_text = str(previous.Text)
        if previous_text.endswith("\r"):
            continue
        current = text_range.Lines(index)
        text = str(current.Text).rstrip("\r\n\v")
        if not text:
            continue
        latin = sum(
            character.isascii() and character.isalpha() for character in text
        )
        short = (
            len(text.split()) <= max_words
            if latin > len(text) * 0.5
            else len(text) <= max_chars
        )
        if short:
            explicit = previous_text.endswith(("\n", "\v"))
            issue: JsonObject = {
                "type": "short_after_vbreak" if explicit else "widow",
                "line_index": index,
                "line_text": text,
                "char_count": len(text),
                "prev_line_text": previous_text.rstrip("\r\n\v"),
                "previous_start": int(previous.Start),
            }
            if explicit:
                issue["issue_type"] = "short_after_vbreak"
            issues.append(issue)
    return issues


@tool("wps_ppt_text", component="wpp", operation="check")
def check_typography(params: CheckTypographyInput) -> JsonObject:
    """Detect overflow, shrinking, off-slide geometry and short lines."""
    session = get_session("wpp")
    presentation = session.get_document()
    indices = (
        [params.slide_index]
        if params.slide_index is not None
        else list(range(1, presentation.Slides.Count + 1))
    )
    if any(index > presentation.Slides.Count for index in indices):
        raise ValueError("slide_index exceeds slide count")
    slide_width = float(presentation.PageSetup.SlideWidth)
    slide_height = float(presentation.PageSetup.SlideHeight)
    issues: list[JsonObject] = []
    fixed: list[JsonObject] = []
    for index in indices:
        goto_slide(session.get_application(), index)
        slide = presentation.Slides(index)
        for shape_index in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(shape_index)
            identity = {"slide_index": index, "shape_name": shape.Name}
            edges = _off_slide_edges(shape, slide_width, slide_height)
            if edges:
                issues.append(
                    {
                        **identity,
                        "type": "off_slide",
                        "edges": edges,
                        "fixable": False,
                        "shape_width": round(shape.Width, 2),
                    }
                )
            if (
                not shape.HasTextFrame
                or not shape.TextFrame.TextRange.Text.strip()
            ):
                continue
            try:
                frame = shape.TextFrame2
                shrinking = frame.AutoSize == PP_AUTO_SIZE_TEXT_TO_FIT_SHAPE
                with (
                    temporary_attributes(frame, AutoSize=PP_AUTO_SIZE_NONE)
                    if shrinking
                    else nullcontext()
                ):
                    natural_height = float(frame.TextRange.BoundHeight)
                    available_height = float(
                        shape.Height - frame.MarginTop - frame.MarginBottom
                    )
                if natural_height > available_height + FIT_TOLERANCE_PT:
                    issues.append(
                        {
                            **identity,
                            "type": "auto_shrink" if shrinking else "overflow",
                            "shape_width": round(shape.Width, 2),
                            "fixable": False,
                            "natural_height": round(natural_height, 2),
                            "available_height": round(available_height, 2),
                        }
                    )
            except WpsError:
                raise
            except Exception as error:
                logger.debug(
                    "Cannot measure shape %s", shape.Name, exc_info=True
                )
                issues.append(
                    {
                        **identity,
                        "type": "measurement_unavailable",
                        "detail": str(error),
                        "fixable": False,
                    }
                )
            line_issues = _short_lines(
                shape, params.max_chars, params.max_words
            )
            widows = [
                issue for issue in line_issues if issue["type"] == "widow"
            ]
            if params.fix and widows:
                original_width = float(shape.Width)
                room = max(
                    0.0,
                    min(
                        _right_neighbor_gap(shape, slide) - 2,
                        slide_width - float(shape.Left) - original_width,
                        params.max_expand_pt,
                    ),
                )
                successful_width = None
                with temporary_attributes(shape, Width=original_width):
                    for expansion in range(1, int(room) + 1):
                        shape.Width = original_width + expansion
                        remaining = _short_lines(
                            shape, params.max_chars, params.max_words
                        )
                        if not any(
                            issue["type"] == "widow" for issue in remaining
                        ):
                            successful_width = float(shape.Width)
                            break
                if successful_width is not None:
                    shape.Width = successful_width
                    fixed.append(
                        {
                            **identity,
                            "old_width": round(original_width, 2),
                            "new_width": round(successful_width, 2),
                            "expanded_by": round(
                                successful_width - original_width, 2
                            ),
                        }
                    )
                else:
                    text_range = shape.TextFrame.TextRange
                    for widow in reversed(widows):
                        breakpoint = _find_best_vbreak(
                            widow["prev_line_text"], widow["line_text"]
                        )
                        if breakpoint is None:
                            continue
                        position, before, after = breakpoint
                        # COM line offsets disambiguate repeated wording.
                        native_position = widow[
                            "previous_start"
                        ] + utf16_length(widow["prev_line_text"][:position])
                        text_range.Characters(native_position, 0).InsertBefore(
                            "\v"
                        )
                        fixed.append(
                            {
                                **identity,
                                "fix_method": "soft_return",
                                "before": before,
                                "after": after,
                            }
                        )
                line_issues = _short_lines(
                    shape, params.max_chars, params.max_words
                )
            issues.extend(
                {**identity, "shape_width": round(shape.Width, 2), **issue}
                for issue in line_issues
            )
    result: JsonObject = {"issues": issues, "total": len(issues)}
    if params.fix:
        result.update(
            fixed=fixed, fixed_count=len(fixed), remaining=len(issues)
        )
    return result
