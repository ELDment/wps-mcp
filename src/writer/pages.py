"""Configure Writer sections, page geometry, and headers or footers."""

from typing import Annotated, Literal, Self

from pydantic import Field, model_validator

from src.models import DocumentInput, JsonObject, NativeObject
from src.native.runtime import get_session
from src.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)
from src.writer.access import resolve_item, resolve_range
from src.writer.models import Index, Position, PositivePoints

HEADER_FOOTER_VALUES = {"primary": 1, "first_page": 2, "even_pages": 3}


class PageSetupInput(DocumentInput):
    """Configure one section using page dimensions in points."""

    section_index: Index = 1
    orientation: Literal["portrait", "landscape"] | None = None
    page_width: PositivePoints | None = None
    page_height: PositivePoints | None = None
    margin_top: Annotated[float, Field(ge=0)] | None = None
    margin_bottom: Annotated[float, Field(ge=0)] | None = None
    margin_left: Annotated[float, Field(ge=0)] | None = None
    margin_right: Annotated[float, Field(ge=0)] | None = None
    header_distance: Annotated[float, Field(ge=0)] | None = None
    footer_distance: Annotated[float, Field(ge=0)] | None = None

    @model_validator(mode="after")
    def validate_setup(self) -> Self:
        """Require at least one explicit page-setup change."""
        if not self.model_dump(
            exclude={"document", "section_index"}, exclude_none=True
        ):
            raise ValueError("Specify at least one page setup property")
        return self


class HeaderFooterInput(DocumentInput):
    """Select a section and its header or footer variant."""

    section_index: Index = 1
    location: Literal["header", "footer"]
    variant: Literal["primary", "first_page", "even_pages"] = "primary"


class WriteHeaderFooterInput(HeaderFooterInput):
    """Replace header or footer text with explicit link behavior."""

    text: str
    link_to_previous: bool = False


class BreakInput(DocumentInput):
    """Insert one supported break at a native text position."""

    position: Position
    kind: Literal[
        "page", "line", "section_next_page", "section_continuous"
    ] = "page"


@tool(
    "wps_writer_pages",
    operation="sections",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def read_sections(params: DocumentInput) -> JsonObject:
    """List sections and current page dimensions/margins in points."""
    document = get_session("writer").get_document(params.document)
    result = []
    for index in range(1, int(document.Sections.Count) + 1):
        section = document.Sections.Item(index)
        setup = section.PageSetup
        result.append(
            {
                "index": index,
                "start": int(section.Range.Start),
                "end": int(section.Range.End),
                "orientation": "landscape"
                if int(setup.Orientation) == 1
                else "portrait",
                "page_width": float(setup.PageWidth),
                "page_height": float(setup.PageHeight),
                "margin_top": float(setup.TopMargin),
                "margin_bottom": float(setup.BottomMargin),
                "margin_left": float(setup.LeftMargin),
                "margin_right": float(setup.RightMargin),
            }
        )
    return {"sections": result}


@tool(
    "wps_writer_pages",
    operation="setup",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def set_page_setup(params: PageSetupInput) -> JsonObject:
    """Set section dimensions, margins, and orientation in points.

    Header and footer distances also use points.
    """
    document = get_session("writer").get_document(params.document)
    setup = resolve_item(
        document.Sections, params.section_index, "section"
    ).PageSetup
    page_width = float(setup.PageWidth)
    page_height = float(setup.PageHeight)
    if params.orientation is not None and int(setup.Orientation) != (
        1 if params.orientation == "landscape" else 0
    ):
        page_width, page_height = (page_height, page_width)
    page_width = (
        params.page_width if params.page_width is not None else page_width
    )
    page_height = (
        params.page_height if params.page_height is not None else page_height
    )
    margin_left = (
        params.margin_left
        if params.margin_left is not None
        else float(setup.LeftMargin)
    )
    margin_right = (
        params.margin_right
        if params.margin_right is not None
        else float(setup.RightMargin)
    )
    margin_top = (
        params.margin_top
        if params.margin_top is not None
        else float(setup.TopMargin)
    )
    margin_bottom = (
        params.margin_bottom
        if params.margin_bottom is not None
        else float(setup.BottomMargin)
    )
    if (
        margin_left + margin_right >= page_width
        or margin_top + margin_bottom >= page_height
    ):
        raise ValueError(
            "Page margins must leave positive content width and height"
        )
    if params.orientation is not None:
        setup.Orientation = 1 if params.orientation == "landscape" else 0
    property_names = {
        "page_width": "PageWidth",
        "page_height": "PageHeight",
        "margin_top": "TopMargin",
        "margin_bottom": "BottomMargin",
        "margin_left": "LeftMargin",
        "margin_right": "RightMargin",
        "header_distance": "HeaderDistance",
        "footer_distance": "FooterDistance",
    }
    for field_name, property_name in property_names.items():
        value = getattr(params, field_name)
        if value is not None:
            setattr(setup, property_name, value)
    return {
        "section_index": params.section_index,
        "page_setup": params.model_dump(
            exclude={"document", "section_index"}, exclude_none=True
        ),
    }


def resolve_header_footer(
    document: NativeObject, params: HeaderFooterInput
) -> tuple[NativeObject, NativeObject]:
    """Resolve a section and its selected native header or footer."""
    section = resolve_item(document.Sections, params.section_index, "section")
    collection = (
        section.Headers if params.location == "header" else section.Footers
    )
    return (section, collection.Item(HEADER_FOOTER_VALUES[params.variant]))


@tool(
    "wps_writer_pages",
    operation="read_header_footer",
    component="writer",
    annotations=READ_ANNOTATIONS,
)
def read_header_footer(params: HeaderFooterInput) -> JsonObject:
    """Read one header or footer variant and its preceding-section link."""
    document = get_session("writer").get_document(params.document)
    section, item = resolve_header_footer(document, params)
    return {
        "section_index": params.section_index,
        "location": params.location,
        "variant": params.variant,
        "text": str(item.Range.Text),
        "link_to_previous": bool(item.LinkToPrevious),
    }


@tool(
    "wps_writer_pages",
    operation="write_header_footer",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def write_header_footer(params: WriteHeaderFooterInput) -> JsonObject:
    """Replace header/footer text.

    Defaults to unlinking the preceding section; enabling link_to_previous
    also edits linked sections. First/even variants enable their page
    settings.
    """
    document = get_session("writer").get_document(params.document)
    section, item = resolve_header_footer(document, params)
    if params.link_to_previous and params.section_index == 1:
        raise ValueError("The first section has no preceding section")
    if params.variant == "first_page":
        section.PageSetup.DifferentFirstPageHeaderFooter = -1
    if params.variant == "even_pages":
        section.PageSetup.OddAndEvenPagesHeaderFooter = -1
    # Enabling a variant can replace its native header/footer object.
    _, item = resolve_header_footer(document, params)
    if params.section_index > 1:
        item.LinkToPrevious = params.link_to_previous
    item.Range.Text = params.text
    return {
        "section_index": params.section_index,
        "location": params.location,
        "variant": params.variant,
        "text": str(item.Range.Text),
        "link_to_previous": bool(item.LinkToPrevious),
    }


@tool(
    "wps_writer_pages",
    operation="insert_break",
    component="writer",
    annotations=WRITE_ANNOTATIONS,
)
def insert_break(params: BreakInput) -> JsonObject:
    """Insert a supported native break at a UTF-16 position.

    Supported kinds are page, line, next-page section, and continuous section.
    """
    document = get_session("writer").get_document(params.document)
    text_range = resolve_range(document, params.position, params.position)
    text_range.InsertBreak(
        {
            "page": 7,
            "line": 6,
            "section_next_page": 2,
            "section_continuous": 3,
        }[params.kind]
    )
    return {"position": params.position, "kind": params.kind}
