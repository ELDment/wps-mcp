"""Organize slide sections with validated one-based indices."""

from typing import Literal, Self

from pydantic import Field, model_validator

from src.models import DocumentInput, JsonObject
from src.native.objects import supports_feature
from src.native.runtime import get_session
from src.protocol.definitions import tool


class AddSectionInput(DocumentInput):
    """Start a named section at a slide."""

    name: str = Field(min_length=1)
    slide_index: int = Field(ge=1)


class ManageSectionInput(DocumentInput):
    """Rename, move or remove a section without deleting slides."""

    section_index: int = Field(ge=1)
    action: Literal["rename", "move", "delete"]
    new_name: str | None = Field(default=None, min_length=1)
    move_to_index: int | None = Field(default=None, ge=1)

    @model_validator(mode="after")
    def validate_action(self) -> Self:
        """Require the arguments used by rename and move actions."""
        if self.action == "rename" and self.new_name is None:
            raise ValueError("new_name is required for rename")
        if self.action == "move" and self.move_to_index is None:
            raise ValueError("move_to_index is required for move")
        return self


@tool("wps_ppt_sections", component="wpp", operation="add")
def add_section(params: AddSectionInput) -> JsonObject:
    """Add a section, warning when an existing section becomes empty."""
    presentation = get_session("wpp").get_document(params.document)
    if params.slide_index > presentation.Slides.Count:
        raise ValueError(
            f"Slide index {params.slide_index} out "
            f"of range (1-{presentation.Slides.Count})"
        )
    sections = presentation.SectionProperties
    prior = None
    for index in range(1, sections.Count + 1):
        if sections.FirstSlide(index) == params.slide_index:
            prior = str(sections.Name(index))
            break
    index = sections.AddBeforeSlide(params.slide_index, params.name)
    result: JsonObject = {
        "success": True,
        "section_index": index,
        "name": params.name,
        "slide_index": params.slide_index,
        "slides_count": sections.SlidesCount(index),
        "sections_count": sections.Count,
    }
    if prior is not None:
        result["warning"] = (
            f"Section {prior!r} also started at slide {params.slide_index} "
            "and is now empty. Use wps_ppt_sections/manage "
            "to remove or move it."
        )
    return result


@tool(
    "wps_ppt_sections",
    component="wpp",
    annotations={"readOnlyHint": True},
    operation="list",
)
def list_sections(params: DocumentInput) -> JsonObject:
    """List sections; empty sections have null slide boundaries."""
    presentation = get_session("wpp").get_document(params.document)
    sections = presentation.SectionProperties
    result: list[JsonObject] = []
    for index in range(1, sections.Count + 1):
        first = sections.FirstSlide(index)
        count = sections.SlidesCount(index)
        empty = count == 0 or first < 1
        result.append(
            {
                "index": index,
                "name": sections.Name(index),
                "first_slide": None if empty else first,
                "last_slide": None if empty else first + count - 1,
                "slides_count": count,
                "empty": empty,
            }
        )
    return {
        "success": True,
        "sections_count": sections.Count,
        "slides_total": presentation.Slides.Count,
        "sections": result,
    }


@tool(
    "wps_ppt_sections",
    component="wpp",
    annotations={"destructiveHint": True},
    operation="manage",
)
def manage_section(params: ManageSectionInput) -> JsonObject:
    """Rename, move or remove a section while retaining slides."""
    presentation = get_session("wpp").get_document(params.document)
    sections = presentation.SectionProperties
    if params.section_index > sections.Count:
        raise ValueError(f"Section index out of range (1-{sections.Count})")
    result: JsonObject = {
        "success": True,
        "action": params.action,
        "section_index": params.section_index,
    }
    if params.action == "rename":
        sections.Rename(params.section_index, params.new_name)
        result["new_name"] = params.new_name
    elif params.action == "move":
        if (
            params.move_to_index is None
            or params.move_to_index > sections.Count
        ):
            raise ValueError(
                f"Section destination out of range (1-{sections.Count})"
            )
        sections.Move(params.section_index, params.move_to_index)
        result["moved_to"] = params.move_to_index
    else:
        result["deleted_section"] = sections.Name(params.section_index)
        if (
            not supports_feature(sections, "delete_first_section")
            and params.section_index == 1
            and sections.Count > 1
        ):
            slides = presentation.Slides
            identities = [
                slides(index).SlideID for index in range(1, slides.Count + 1)
            ]
            sections.Move(1, 2)
            try:
                sections.Delete(2, False)
            finally:
                for index, identity in enumerate(identities, 1):
                    slides.FindBySlideID(identity).MoveTo(index)
        else:
            sections.Delete(params.section_index, False)
    return result
