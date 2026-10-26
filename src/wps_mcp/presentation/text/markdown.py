"""Extract slide text as Markdown in visual reading order."""

from __future__ import annotations

import logging
from typing import TypedDict

from wps_mcp.models import NativeObject
from wps_mcp.native.objects import is_optional_member_error, optional_members
from wps_mcp.presentation.constants import (
    MSO_GROUP,
    MSO_PLACEHOLDER,
    MSO_TRUE,
    PP_BULLET_NUMBERED,
)

logger = logging.getLogger(__name__)

_SKIP_PLACEHOLDER_TYPES = {13, 14, 15, 16}


_TITLE_PLACEHOLDER_TYPES = {1, 3, 5}


_SUBTITLE_PLACEHOLDER_TYPES = {4}


class ShapeInfo(TypedDict):
    """Readable shape bounds and semantic roles for Markdown extraction."""

    shape: NativeObject
    top: float
    left: float
    width: float
    height: float
    is_title: bool
    is_subtitle: bool
    has_table: bool


def _is_all_bold(shape: NativeObject) -> bool:
    """Check if ALL text runs in a shape are bold."""
    try:
        tr = shape.TextFrame.TextRange
        text = tr.Text.strip()
        if not text:
            return False
        run_count = tr.Runs().Count
        if run_count == 0:
            return False
        for i in range(1, run_count + 1):
            run = tr.Runs(i)
            if not run.Text.strip():
                continue
            if run.Font.Bold != MSO_TRUE:
                return False
        return True
    except Exception as error:
        if not is_optional_member_error(error):
            raise
        return False


def _runs_to_markdown(paragraph: NativeObject) -> str:
    """Merge adjacent equal styles without dropping literal run text."""
    try:
        run_count = paragraph.Runs().Count
    except Exception as error:
        if not is_optional_member_error(error):
            raise
        logger.debug(
            "Run formatting unavailable; using plain text", exc_info=True
        )
        return str(paragraph.Text).replace("\r", "").replace("\v", "\n")
    if not run_count:
        return str(paragraph.Text).replace("\r", "").replace("\v", "\n")
    merged: list[tuple[str, bool, bool]] = []
    for index in range(1, run_count + 1):
        run = paragraph.Runs(index)
        text = str(run.Text).replace("\r", "").replace("\v", "\n")
        if not text:
            continue
        try:
            bold, italic = (
                run.Font.Bold == MSO_TRUE,
                run.Font.Italic == MSO_TRUE,
            )
        except Exception as error:
            if not is_optional_member_error(error):
                raise
            logger.debug(
                "Font formatting unavailable for run %s", index, exc_info=True
            )
            bold, italic = False, False
        if merged and merged[-1][1:] == (bold, italic):
            text = merged.pop()[0] + text
        merged.append((text, bold, italic))
    parts = []
    for text, bold, italic in merged:
        marker = "*" * (2 * bold + italic)
        parts.append(f"{marker}{text}{marker}")
    return "".join(parts)


def _shape_paragraphs_to_markdown(
    shape: NativeObject, as_heading: str = ""
) -> str:
    """Convert a shape's paragraphs to Markdown text."""
    tr = shape.TextFrame.TextRange
    if as_heading:
        text = tr.Text.replace("\r", " ").replace("\x0b", " ")
        if not text:
            return ""
        return f"{as_heading} {text}"
    lines = []
    para_count = tr.Paragraphs().Count
    for i in range(1, para_count + 1):
        para = tr.Paragraphs(i)
        text = _runs_to_markdown(para)
        if not text:
            lines.append("")
            continue
        indent_level = para.IndentLevel
        bullet_prefix = ""
        try:
            bullet = para.ParagraphFormat.Bullet
            if bullet.Visible == MSO_TRUE:
                indent = "  " * max(0, indent_level - 1)
                if bullet.Type == PP_BULLET_NUMBERED:
                    bullet_prefix = f"{indent}1. "
                else:
                    bullet_prefix = f"{indent}- "
        except Exception as error:
            if not is_optional_member_error(error):
                raise
            logger.debug("Bullet formatting unavailable", exc_info=True)
        lines.append(f"{bullet_prefix}{text}")
    return "\n".join(lines)


def _table_to_markdown(shape: NativeObject) -> str:
    """Read every table cell, escaping Markdown separators."""
    table = shape.Table
    rows = []
    for row_index in range(1, table.Rows.Count + 1):
        cells = []
        for column_index in range(1, table.Columns.Count + 1):
            text = str(
                table.Cell(
                    row_index, column_index
                ).Shape.TextFrame.TextRange.Text
            )
            cells.append(
                text.replace("\r", " ").replace("\v", " ").replace("|", "\\|")
            )
        rows.append("| " + " | ".join(cells) + " |")
        if row_index == 1:
            rows.append(
                "| " + " | ".join(["---"] * table.Columns.Count) + " |"
            )
    return "\n".join(rows)


def _collect_text_shapes(slide: NativeObject) -> list[ShapeInfo]:
    """Collect readable text and tables while retaining their visual order."""
    shapes: list[ShapeInfo] = []

    def process_shape(
        shape: NativeObject, offset_top: float = 0.0, offset_left: float = 0.0
    ) -> None:
        shape_type = shape.Type
        if shape_type == MSO_GROUP:
            for index in range(1, shape.GroupItems.Count + 1):
                process_shape(
                    shape.GroupItems(index),
                    offset_top + shape.Top,
                    offset_left + shape.Left,
                )
            return
        placeholder_type = (
            shape.PlaceholderFormat.Type
            if shape_type == MSO_PLACEHOLDER
            else None
        )
        if placeholder_type in _SKIP_PLACEHOLDER_TYPES:
            return
        has_table = bool(shape.HasTable)
        if not has_table and not (
            shape.HasTextFrame and shape.TextFrame.HasText
        ):
            return
        shapes.append(
            {
                "shape": shape,
                "top": float(shape.Top + offset_top),
                "left": float(shape.Left + offset_left),
                "width": float(shape.Width),
                "height": float(shape.Height),
                "is_title": placeholder_type in _TITLE_PLACEHOLDER_TYPES,
                "is_subtitle": placeholder_type in _SUBTITLE_PLACEHOLDER_TYPES,
                "has_table": has_table,
            }
        )

    for index in range(1, slide.Shapes.Count + 1):
        process_shape(slide.Shapes(index))
    return shapes


def _group_into_rows(
    shape_infos: list[ShapeInfo], threshold: float = 0.4
) -> list[list[ShapeInfo]]:
    """Group shapes into rows based on vertical overlap."""
    if not shape_infos:
        return []
    sorted_shapes = sorted(shape_infos, key=lambda s: (s["top"], s["left"]))
    rows = []
    used = set()
    for i, s in enumerate(sorted_shapes):
        if i in used:
            continue
        row = [s]
        used.add(i)
        s_top = s["top"]
        s_bottom = s_top + s["height"]
        for j in range(i + 1, len(sorted_shapes)):
            if j in used:
                continue
            other = sorted_shapes[j]
            o_top = other["top"]
            o_bottom = o_top + other["height"]
            overlap_top = max(s_top, o_top)
            overlap_bottom = min(s_bottom, o_bottom)
            overlap = max(0, overlap_bottom - overlap_top)
            shorter_height = min(s["height"], other["height"])
            if shorter_height > 0 and overlap / shorter_height >= threshold:
                row.append(other)
                used.add(j)
        row.sort(key=lambda x: x["left"])
        rows.append(row)
    rows.sort(key=lambda row: sum(s["top"] for s in row) / len(row))
    return rows


def _shape_info_to_markdown(
    info: ShapeInfo, subheading_level: str = "##"
) -> str:
    """Convert a single shape_info dict to Markdown text."""
    shape = info["shape"]
    if info["has_table"]:
        return _table_to_markdown(shape)
    if info["is_title"]:
        return _shape_paragraphs_to_markdown(shape, as_heading="#")
    if info["is_subtitle"]:
        return _shape_paragraphs_to_markdown(shape)
    if _is_all_bold(shape):
        return _shape_paragraphs_to_markdown(
            shape, as_heading=subheading_level
        )
    return _shape_paragraphs_to_markdown(shape)


def _group_into_columns(
    shapes: list[ShapeInfo], threshold: float = 50.0
) -> list[list[ShapeInfo]]:
    """Group shapes by X-position proximity into columns."""
    if not shapes:
        return []
    sorted_shapes = sorted(shapes, key=lambda s: s["left"])
    columns = []
    current_col = [sorted_shapes[0]]
    col_avg_left = sorted_shapes[0]["left"]
    for s in sorted_shapes[1:]:
        if abs(s["left"] - col_avg_left) <= threshold:
            current_col.append(s)
            col_avg_left = sum(x["left"] for x in current_col) / len(
                current_col
            )
        else:
            columns.append(current_col)
            current_col = [s]
            col_avg_left = s["left"]
    columns.append(current_col)
    for col in columns:
        col.sort(key=lambda s: s["top"])
    columns.sort(key=lambda col: sum(s["left"] for s in col) / len(col))
    rows = _group_into_rows(shapes)
    singletons = {id(column[0]) for column in columns if len(column) == 1}
    if any(
        id(shape) in singletons
        for row in rows
        if len(row) > 1
        for shape in row
    ):
        return [[s for row in rows for s in row]]
    return columns


def slide_to_markdown(slide: NativeObject, slide_index: int) -> str:
    """Convert a single slide to pseudo-Markdown."""
    hidden = ""
    with optional_members():
        if slide.SlideShowTransition.Hidden:
            hidden = " (hidden)"
    parts = [f"== Slide {slide_index}{hidden} =="]
    shape_infos = _collect_text_shapes(slide)
    if not shape_infos:
        parts.append("(no text)")
        return "\n".join(parts)
    rows = _group_into_rows(shape_infos)
    has_multi_shape_rows = any(len(row) > 1 for row in rows)
    if not has_multi_shape_rows:
        for row in rows:
            md = _shape_info_to_markdown(row[0])
            if md.strip():
                parts.append(md)
    else:
        pending_column_shapes: list[ShapeInfo] = []

        def flush_columns() -> None:
            """Group pending column shapes by X and append to parts."""
            if not pending_column_shapes:
                return
            columns = _group_into_columns(pending_column_shapes)
            for col_idx, col in enumerate(columns):
                if col_idx > 0:
                    parts.append("")
                for info in col:
                    md = _shape_info_to_markdown(info, subheading_level="###")
                    if md.strip():
                        parts.append(md)
            pending_column_shapes.clear()

        for row in rows:
            if len(row) == 1:
                flush_columns()
                md = _shape_info_to_markdown(row[0])
                if md.strip():
                    parts.append(md)
            else:
                pending_column_shapes.extend(row)
        flush_columns()
    return "\n".join(parts)
