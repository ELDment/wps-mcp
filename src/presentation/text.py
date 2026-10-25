"""Text editing and layout tools using WPS's one-based UTF-16 positions."""

from __future__ import annotations

from collections.abc import Iterator, Sequence
from pathlib import Path

from src.common.colors import hex_to_int, int_to_hex
from src.common.text import normalize_text
from src.models import JsonObject, NativeObject
from src.native.runtime import get_session
from src.presentation.constants import (
    MSO_FALSE,
    MSO_TRUE,
    PP_BULLET_NONE,
)
from src.presentation.navigation import goto_slide
from src.presentation.shape_lookup import (
    PATH_SEPARATOR,
    resolve_shape,
    walk_group_children,
)
from src.presentation.text_enums import (
    ALIGNMENT_MAP,
    AUTO_SIZE_MAP,
    BULLET_TYPE_MAP,
    NUMBERED_STYLE_MAP,
    ORIENTATION_MAP,
    VERTICAL_ANCHOR_MAP,
)
from src.presentation.text_formatting import apply_text_format
from src.presentation.text_markdown import slide_to_markdown
from src.presentation.text_measurement import measure_text
from src.presentation.text_models import (
    FindReplaceTextInput,
    FormatTextInput,
    FormatTextRangeInput,
    GetAllTextInput,
    GetTextInput,
    SetBulletInput,
    SetParagraphFormatInput,
    SetTextframeInput,
    SetTextInput,
    ShapeTextInput,
    TextRangeSpec,
)
from src.presentation.text_spans import (
    find_spans,
    python_span,
    resolve_span,
    run_offsets,
    utf16_length,
)
from src.protocol.definitions import tool


def _searchable_shapes(
    slide: NativeObject, include_groups: bool
) -> Iterator[tuple[NativeObject, str]]:
    """Every shape on the slide that a text search should look at."""
    for si in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(si)
        yield (shape, shape.Name)
        if include_groups:
            yield from walk_group_children(shape, shape.Name + PATH_SEPARATOR)


def _text_frame_of(
    params: ShapeTextInput,
) -> tuple[NativeObject, NativeObject]:
    session = get_session("wpp")
    presentation = session.get_document()
    if params.slide_index > presentation.Slides.Count:
        raise ValueError(f"Slide index out of range: {params.slide_index}")
    goto_slide(session.get_application(), params.slide_index)
    shape = resolve_shape(
        presentation.Slides(params.slide_index), params.shape_name_or_index
    )
    if not shape.HasTextFrame:
        raise ValueError(f"Shape {shape.Name!r} does not have a text frame")
    return shape, shape.TextFrame.TextRange


@tool("wps_ppt_text", component="wpp", operation="set")
def set_text(params: SetTextInput) -> JsonObject:
    """Replace text spans or formatted runs using native UTF-16 offsets."""
    shape, text_range = _text_frame_of(params)
    result: JsonObject = {
        "status": "success",
        "slide_index": params.slide_index,
        "shape_name": shape.Name,
        "position_unit": "UTF-16",
    }
    if params.runs is not None:
        if any(run.highlight_color is not None for run in params.runs):
            _ = shape.TextFrame2.TextRange.Font.Highlight
        spans = run_offsets(params.runs)
        text_range.Text = "".join(
            normalize_text(run.text) for run in params.runs
        )
        warnings = set()
        for (start, length), run in zip(spans, params.runs, strict=True):
            applied = apply_text_format(shape, run, start, length)
            warnings.update(applied.get("warnings", []))
        result["runs"] = [
            {
                "index": index,
                "text": run.text,
                "start": start,
                "length": length,
            }
            for index, ((start, length), run) in enumerate(
                zip(spans, params.runs, strict=True), 1
            )
        ]
        if warnings:
            result["warnings"] = sorted(warnings)
    else:
        assert params.text is not None
        text = normalize_text(params.text)
        if params.start is not None or params.search_text is not None:
            start, length = resolve_span(text_range.Text, shape.Name, params)
            text_range.Characters(Start=start, Length=length).Text = text
            result.update(
                text=text_range.Text,
                start=start,
                replaced_length=length,
                written_length=utf16_length(text),
            )
        else:
            text_range.Text = text
    result.update(
        text_length=text_range.Length,
        paragraph_count=text_range.Paragraphs().Count,
    )
    return result


@tool(
    "wps_ppt_text",
    component="wpp",
    annotations={"readOnlyHint": True},
    operation="get",
)
def get_text(params: GetTextInput) -> JsonObject:
    """Read text, UTF-16 positions, formatting and optional measurements."""
    shape, text_range = _text_frame_of(params)
    paragraphs = []
    for index in range(1, text_range.Paragraphs().Count + 1):
        paragraph = text_range.Paragraphs(index)
        paragraphs.append(
            {
                "index": index,
                "text": paragraph.Text,
                "indent_level": paragraph.IndentLevel,
                "alignment": paragraph.ParagraphFormat.Alignment,
            }
        )
    runs = []
    for index in range(1, text_range.Runs().Count + 1):
        run = text_range.Runs(index)
        font = run.Font
        runs.append(
            {
                "index": index,
                "text": run.Text,
                "start": run.Start,
                "length": run.Length,
                "font_name": font.Name,
                "font_name_fareast": font.NameFarEast,
                "font_size": font.Size,
                "bold": font.Bold == MSO_TRUE,
                "italic": font.Italic == MSO_TRUE,
                "underline": font.Underline == MSO_TRUE,
                "color_hex": int_to_hex(font.Color.RGB),
            }
        )
    result: JsonObject = {
        "status": "success",
        "shape_name": shape.Name,
        "text": text_range.Text,
        "text_length": text_range.Length,
        "position_unit": "UTF-16",
        "paragraph_count": len(paragraphs),
        "paragraphs": paragraphs,
        "runs": runs,
    }
    if params.measure:
        result["measurement"] = measure_text(shape, text_range)
    return result


@tool("wps_ppt_text", component="wpp", operation="format")
def format_text(params: FormatTextInput) -> JsonObject:
    """Set font properties on all text in a shape."""
    shape, _ = _text_frame_of(params)
    return apply_text_format(shape, params)


@tool("wps_ppt_text", component="wpp", operation="format_range")
def format_text_range(params: FormatTextRangeInput) -> JsonObject:
    """Validate all spans before applying font properties in order."""
    shape, text_range = _text_frame_of(params)
    specs: Sequence[TextRangeSpec | FormatTextRangeInput] = params.ranges or [
        params
    ]
    spans = [resolve_span(text_range.Text, shape.Name, spec) for spec in specs]
    if any(spec.highlight_color is not None for spec in specs) or (
        params.base is not None and params.base.highlight_color is not None
    ):
        _ = shape.TextFrame2.TextRange.Font.Highlight
    warnings = set()
    if params.base is not None:
        base_result = apply_text_format(shape, params.base)
        warnings.update(base_result.get("warnings", []))
    applied = []
    for spec, (start, length) in zip(specs, spans, strict=True):
        result = apply_text_format(shape, spec, start, length)
        warnings.update(result.get("warnings", []))
        applied.append(result)
    if params.ranges is None:
        return applied[0]
    result = {
        "status": "success",
        "shape_name": shape.Name,
        "count": len(applied),
        "ranges": applied,
        "position_unit": "UTF-16",
    }
    if warnings:
        result["warnings"] = sorted(warnings)
    return result


@tool("wps_ppt_text", component="wpp", operation="paragraph")
def set_paragraph_format(params: SetParagraphFormatInput) -> JsonObject:
    """Set alignment, spacing, and indentation on one or all paragraphs."""
    shape, text_range = _text_frame_of(params)
    if (
        params.paragraph_index is not None
        and params.paragraph_index > text_range.Paragraphs().Count
    ):
        raise ValueError("paragraph_index exceeds paragraph count")
    target = (
        text_range
        if params.paragraph_index is None
        else text_range.Paragraphs(params.paragraph_index)
    )
    paragraph_format = target.ParagraphFormat
    if params.alignment is not None:
        paragraph_format.Alignment = ALIGNMENT_MAP[params.alignment]
    for field, rule, attribute, multiplier in (
        ("line_spacing", "LineRuleWithin", "SpaceWithin", True),
        ("space_before", "LineRuleBefore", "SpaceBefore", False),
        ("space_after", "LineRuleAfter", "SpaceAfter", False),
    ):
        value = getattr(params, field)
        if value is not None:
            setattr(
                paragraph_format, rule, MSO_TRUE if multiplier else MSO_FALSE
            )
            setattr(paragraph_format, attribute, value)
    if params.indent_level is not None:
        target.IndentLevel = params.indent_level
    if params.first_line_indent is not None:
        paragraph_format.FirstLineIndent = params.first_line_indent
    return {
        "status": "success",
        "shape_name": shape.Name,
        "paragraph_index": params.paragraph_index or "all",
    }


@tool("wps_ppt_text", component="wpp", operation="bullets")
def set_bullet(params: SetBulletInput) -> JsonObject:
    """Set list numbering, nesting, font and color overrides."""
    shape, text_range = _text_frame_of(params)
    if (
        params.paragraph_index is not None
        and params.paragraph_index > text_range.Paragraphs().Count
    ):
        raise ValueError("paragraph_index exceeds paragraph count")
    target = (
        text_range
        if params.paragraph_index is None
        else text_range.Paragraphs(params.paragraph_index)
    )
    bullet = target.ParagraphFormat.Bullet
    bullet_type = BULLET_TYPE_MAP[params.bullet_type]
    bullet.Visible = MSO_FALSE if bullet_type == PP_BULLET_NONE else MSO_TRUE
    if bullet_type != PP_BULLET_NONE:
        bullet.Type = bullet_type
    if params.bullet_char is not None:
        bullet.Character = ord(params.bullet_char)
    if params.bullet_start_value is not None:
        bullet.StartValue = params.bullet_start_value
    if params.indent_level is not None:
        target.IndentLevel = params.indent_level
    if params.numbered_style is not None:
        bullet.Style = NUMBERED_STYLE_MAP[params.numbered_style]
    if params.size is not None:
        bullet.RelativeSize = params.size
    if params.use_text_color is not None:
        bullet.UseTextColor = MSO_TRUE if params.use_text_color else MSO_FALSE
    if params.use_text_font is not None:
        bullet.UseTextFont = MSO_TRUE if params.use_text_font else MSO_FALSE
    if params.color is not None:
        bullet.Font.Color.RGB = hex_to_int(params.color)
    if params.font_name is not None:
        bullet.Font.Name = params.font_name
    return {
        "status": "success",
        "shape_name": shape.Name,
        "paragraph_index": params.paragraph_index or "all",
        "bullet_type": params.bullet_type,
        "numbered_style": params.numbered_style,
        "indent_level": params.indent_level,
        "color_hex": params.color,
        "size": params.size,
        "font_name": params.font_name,
        "use_text_color": False
        if params.color is not None
        else params.use_text_color,
        "use_text_font": False
        if params.font_name is not None
        else params.use_text_font,
    }


@tool("wps_ppt_text", component="wpp", operation="replace")
def find_replace_text(params: FindReplaceTextInput) -> JsonObject:
    """Find or replace original matches using one-based UTF-16 spans."""
    presentation = get_session("wpp").get_document()
    indices = params.slide_indices or list(
        range(1, presentation.Slides.Count + 1)
    )
    if any(index > presentation.Slides.Count for index in indices):
        raise ValueError(
            f"slide_indices exceeds slide count {presentation.Slides.Count}"
        )
    find_only = params.replace_text is None or params.dry_run
    find_text = normalize_text(params.find_text)
    hits = []
    for index in dict.fromkeys(indices):
        slide = presentation.Slides(index)
        for shape, path in _searchable_shapes(slide, params.include_groups):
            if params.shape_name is not None and params.shape_name not in (
                shape.Name,
                path,
            ):
                continue
            if not shape.HasTextFrame:
                continue
            text_range = shape.TextFrame.TextRange
            spans = find_spans(
                str(text_range.Text),
                find_text,
                match_case=params.match_case,
                whole_words=params.whole_words,
            )
            replacement = params.replace_text
            if not find_only:
                assert replacement is not None
                replacement = normalize_text(replacement)
                # Reverse writes preserve original spans, including deletions.
                for start, length in reversed(spans):
                    text_range.Characters(
                        Start=start, Length=length
                    ).Text = replacement
            offset = 0
            for start, length in spans:
                original_length = length
                if not find_only:
                    assert replacement is not None
                    start += offset
                    length = utf16_length(replacement)
                    offset += length - original_length
                hit: JsonObject = {
                    "slide_index": index,
                    "shape_name": shape.Name,
                    "shape_path": path,
                    "start": start,
                    "length": length,
                }
                if params.context_chars:
                    text = text_range.Text
                    begin, end = python_span(text, start, length)
                    hit["context"] = (
                        text[max(0, begin - params.context_chars) : begin]
                        + "["
                        + text[begin:end]
                        + "]"
                        + text[end : end + params.context_chars]
                    )
                hits.append(hit)
    return {
        "status": "success",
        "mode": "find" if find_only else "replace",
        "find_text": params.find_text,
        "replace_text": params.replace_text,
        "match_count": len(hits),
        "matches": hits,
        "position_unit": "UTF-16",
    }


@tool("wps_ppt_text", component="wpp", operation="frame")
def set_textframe(params: SetTextframeInput) -> JsonObject:
    """Set auto-fit, margins, wrapping, orientation, and vertical alignment."""
    shape, _ = _text_frame_of(params)
    frame = shape.TextFrame
    for field, attribute in (
        ("margin_left", "MarginLeft"),
        ("margin_right", "MarginRight"),
        ("margin_top", "MarginTop"),
        ("margin_bottom", "MarginBottom"),
    ):
        value = getattr(params, field)
        if value is not None:
            setattr(frame, attribute, value)
    if params.word_wrap is not None:
        frame.WordWrap = MSO_TRUE if params.word_wrap else MSO_FALSE
    if params.orientation is not None:
        frame.Orientation = ORIENTATION_MAP[params.orientation]
    if params.auto_size is not None:
        shape.TextFrame2.AutoSize = AUTO_SIZE_MAP[params.auto_size]
    if params.vertical_anchor is not None:
        frame.VerticalAnchor = VERTICAL_ANCHOR_MAP[params.vertical_anchor]
    return {"status": "success", "shape_name": shape.Name}


@tool(
    "wps_ppt_text",
    component="wpp",
    annotations={"readOnlyHint": True},
    operation="all",
)
def get_all_text(params: GetAllTextInput) -> JsonObject | str:
    """Extract Markdown and optionally save a UTF-8 file with CRLF."""
    presentation = get_session("wpp").get_document()
    indices = params.slide_indices or list(
        range(1, presentation.Slides.Count + 1)
    )
    if any(index > presentation.Slides.Count for index in indices):
        raise ValueError(
            f"slide_indices exceeds slide count {presentation.Slides.Count}"
        )
    text = "\n\n".join(
        slide_to_markdown(presentation.Slides(index), index)
        for index in indices
    )
    if params.output_path is None:
        return text
    path = Path(params.output_path).resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\r\n") as output:
        output.write(text)
    return {
        "status": "success",
        "output_path": str(path),
        "slide_count": len(indices),
    }
