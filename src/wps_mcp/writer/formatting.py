"""Apply shared Writer character and paragraph formatting."""

from wps_mcp.common.colors import hex_to_int
from wps_mcp.models import NativeObject
from wps_mcp.writer.models import CharacterFormat, ParagraphFormat

ALIGNMENT_VALUES = {
    "left": 0,
    "center": 1,
    "right": 2,
    "justify": 3,
    "distribute": 4,
}


UNDERLINE_VALUES = {"none": 0, "single": 1, "double": 3}


def apply_character(
    text_range: NativeObject, character: CharacterFormat | None
) -> None:
    """Apply only supplied font properties using Writer native values."""
    if character is None:
        return
    font = text_range.Font
    property_names = {
        "font_name": "Name",
        "font_name_far_east": "NameFarEast",
        "font_size": "Size",
        "bold": "Bold",
        "italic": "Italic",
    }
    for field_name, property_name in property_names.items():
        value = getattr(character, field_name)
        if value is not None:
            setattr(
                font,
                property_name,
                (-1 if value else 0) if isinstance(value, bool) else value,
            )
    if character.underline is not None:
        font.Underline = UNDERLINE_VALUES[character.underline]
    if character.color is not None:
        font.Color = hex_to_int(character.color)


def apply_paragraph(
    text_range: NativeObject, paragraph: ParagraphFormat | None
) -> None:
    """Apply supplied paragraph properties without changing other fields."""
    if paragraph is None:
        return
    paragraph_format = text_range.ParagraphFormat
    property_names = {
        "left_indent": "LeftIndent",
        "right_indent": "RightIndent",
        "first_line_indent": "FirstLineIndent",
        "space_before": "SpaceBefore",
        "space_after": "SpaceAfter",
        "keep_with_next": "KeepWithNext",
        "keep_together": "KeepTogether",
        "page_break_before": "PageBreakBefore",
    }
    for field_name, property_name in property_names.items():
        value = getattr(paragraph, field_name)
        if value is not None:
            setattr(
                paragraph_format,
                property_name,
                (-1 if value else 0) if isinstance(value, bool) else value,
            )
    if paragraph.alignment is not None:
        paragraph_format.Alignment = ALIGNMENT_VALUES[paragraph.alignment]
    if paragraph.line_spacing is not None:
        paragraph_format.LineSpacingRule = 4
        paragraph_format.LineSpacing = paragraph.line_spacing
