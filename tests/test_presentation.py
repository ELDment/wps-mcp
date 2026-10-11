"""Verify presentation contracts using deterministic native stand-ins."""

from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from pydantic import ValidationError

from src.models import DocumentInput, JsonObject
from src.native.context import CALL_DOCUMENT
from src.native.selection import select_document
from src.presentation import (
    app,
    placeholders,
    sections,
    slides,
)
from src.presentation import (
    document as presentation_document,
)
from src.presentation import (
    export as presentation_export,
)


class NativeSlides:
    """Model index shifts while exposing the native collection's methods."""

    def __init__(
        self, count: int, clipboard: dict[str, Mock] | None = None
    ) -> None:
        self.items: list[Mock] = []
        self.clipboard = clipboard if clipboard is not None else {}
        self.next_id = 101
        self.native = Mock(side_effect=self.item)
        self.native.FindBySlideID.side_effect = self.find
        self.native.Paste.side_effect = self.paste
        for _ in range(count):
            self.insert(len(self.items) + 1)

    def item(self, index: int) -> Mock:
        return self.items[index - 1]

    def find(self, slide_id: int) -> Mock:
        return next(slide for slide in self.items if slide.SlideID == slide_id)

    def refresh(self) -> None:
        self.native.Count = len(self.items)
        for index, slide in enumerate(self.items, 1):
            slide.SlideIndex = index

    def insert(self, index: int, source: Mock | None = None) -> Mock:
        slide = Mock()
        slide.SlideID = self.next_id
        slide.content = (
            source.content if source is not None else f"content-{self.next_id}"
        )
        self.next_id += 1
        self.items.insert(index - 1, slide)
        self.refresh()

        def move(position: int | None = None, **options: int) -> None:
            destination = options.get("toPos", position)
            assert destination is not None
            self.items.remove(slide)
            self.items.insert(destination - 1, slide)
            self.refresh()

        def delete() -> None:
            self.items.remove(slide)
            self.refresh()

        def duplicate() -> Mock:
            copied = self.insert(slide.SlideIndex + 1, slide)
            return Mock(return_value=copied)

        def copy() -> None:
            self.clipboard["slide"] = slide

        slide.MoveTo.side_effect = move
        slide.Delete.side_effect = delete
        slide.Duplicate.side_effect = duplicate
        slide.Copy.side_effect = copy
        return slide

    def paste(self, index: int | None = None) -> Mock:
        copied = self.insert(
            index if index is not None else len(self.items) + 1,
            self.clipboard["slide"],
        )
        return Mock(return_value=copied)


@pytest.fixture
def native_session(monkeypatch: pytest.MonkeyPatch) -> SimpleNamespace:
    store = NativeSlides(5)
    document = Mock()
    document.Name = "source.pptx"
    document.FullName = "C:\\decks\\source.pptx"
    document.Slides = store.native
    document.Windows.Count = 0
    documents = [document]
    application = Mock()
    application.Presentations = Mock(
        side_effect=lambda index: documents[index - 1]
    )
    application.Presentations.Count = len(documents)
    session = SimpleNamespace(
        documents=documents,
        store=store,
        get_document=Mock(
            side_effect=lambda value=None: select_document(documents, value)
        ),
        get_application=Mock(return_value=application),
        get_window=Mock(return_value=None),
    )
    for module in (
        app,
        presentation_document,
        slides,
        placeholders,
        presentation_export,
        sections,
    ):
        monkeypatch.setattr(module, "get_session", Mock(return_value=session))
    monkeypatch.setattr(slides, "nav_goto_slide", Mock())
    monkeypatch.setattr(placeholders, "goto_slide", Mock())
    return session


@pytest.mark.parametrize(
    "options",
    [
        {},
        {"slide_index": 1, "slide_indices": [2]},
        {"slide_indices": []},
        {"slide_indices": [0]},
        {"from_index": 2},
        {"from_index": 3, "to_index": 2},
    ],
)
def test_delete_rejects_ambiguous_or_invalid_selection(
    options: JsonObject,
) -> None:
    with pytest.raises(ValidationError):
        slides.DeleteSlideInput(**options)


def test_explicit_document_cannot_conflict_with_index(
    native_session: SimpleNamespace,
) -> None:
    other = Mock(Name="other.pptx", FullName="C:\\decks\\other.pptx")
    native_session.documents.append(other)
    application = native_session.get_application()
    application.Presentations.Count = 2
    with pytest.raises(ValueError, match="conflicts"):
        presentation_document.resolve_presentation(
            application, 1, document=other.FullName
        )


def test_source_presentation_may_differ_from_call_document(
    native_session: SimpleNamespace,
) -> None:
    other = Mock(Name="other.pptx", FullName="C:\\decks\\other.pptx")
    native_session.documents.append(other)
    application = native_session.get_application()
    application.Presentations.Count = 2
    token = CALL_DOCUMENT.set(other.FullName)
    try:
        result = presentation_document.resolve_presentation(
            application, 1, enforce_target=False
        )
        assert result is native_session.documents[0]
        with pytest.raises(ValueError, match="conflicts"):
            presentation_document.resolve_presentation(application, 1)
    finally:
        CALL_DOCUMENT.reset(token)


def test_delete_uses_original_indices_and_keeps_one_slide(
    native_session: SimpleNamespace,
) -> None:
    result = slides.delete_slide(
        slides.DeleteSlideInput(slide_indices=[2, 4, 2])
    )
    assert result["deleted_indices"] == [2, 4]
    assert [slide.SlideID for slide in native_session.store.items] == [
        101,
        103,
        105,
    ]
    with pytest.raises(ValueError, match="Cannot delete every slide"):
        slides.delete_slide(slides.DeleteSlideInput(from_index=1, to_index=3))
    assert len(native_session.store.items) == 3


def test_move_noncontiguous_slides_preserves_relative_order(
    native_session: SimpleNamespace,
) -> None:
    result = slides.move_slide(
        slides.MoveSlideInput(slide_indices=[5, 1, 4], new_position=2)
    )
    assert result["moved_slide_ids"] == [101, 104, 105]
    assert [slide.SlideID for slide in native_session.store.items] == [
        102,
        101,
        104,
        105,
        103,
    ]


def test_move_rejects_impossible_block_before_mutation(
    native_session: SimpleNamespace,
) -> None:
    with pytest.raises(ValueError, match="out of range"):
        slides.move_slide(
            slides.MoveSlideInput(slide_indices=[1, 3, 5], new_position=4)
        )
    assert [slide.SlideID for slide in native_session.store.items] == [
        101,
        102,
        103,
        104,
        105,
    ]


def test_duplicate_before_source_tracks_identity(
    native_session: SimpleNamespace,
) -> None:
    result = slides.duplicate_slide(
        slides.DuplicateSlideInput(slide_index=4, insert_at=1, count=3)
    )
    assert result["new_slide_indices"] == [1, 2, 3]
    assert [slide.content for slide in native_session.store.items[:3]] == [
        "content-104"
    ] * 3
    assert native_session.store.find(104).SlideIndex == 7


def test_copy_within_presentation_anchors_sources_by_id(
    native_session: SimpleNamespace,
) -> None:
    result = slides.copy_slide(
        slides.CopySlideInput(slide_indices=[4, 2], insert_at=1)
    )
    assert result["new_slide_indices"] == [1, 2]
    assert [slide.content for slide in native_session.store.items[:2]] == [
        "content-104",
        "content-102",
    ]


def test_copy_to_explicit_document_navigates_only_destination(
    native_session: SimpleNamespace,
) -> None:
    source = native_session.documents[0]
    target_store = NativeSlides(1, native_session.store.clipboard)
    target = Mock(Name="target.pptx", FullName="C:\\decks\\target.pptx")
    target.Slides = target_store.native
    target.Windows.Count = 1
    native_session.documents.append(target)
    native_session.get_application().Presentations.Count = 2
    result = slides.copy_slide(
        slides.CopySlideInput(
            source_presentation_index=1,
            document=target.FullName,
            slide_index=4,
        )
    )
    assert target_store.items[1].content == "content-104"
    assert result["target_presentation"] == "target.pptx"
    source.Windows.assert_not_called()
    target.Windows(1).View.GotoSlide.assert_called_once_with(2)


def test_preview_cleans_temporary_directory_after_invalid_output(
    native_session: SimpleNamespace,
) -> None:
    exported: list[Path] = []

    def export(file_path: str, image_format: str) -> None:
        exported.append(Path(file_path))
        Path(file_path).write_bytes(b"not a png")

    native_session.store.items[0].Export.side_effect = export
    with pytest.raises(RuntimeError, match="valid PNG"):
        presentation_export.get_slide_preview(
            presentation_export.GetSlidePreviewInput(slide_index=1)
        )
    assert exported and not exported[0].parent.exists()


def test_preview_returns_bytes_and_preserves_explicit_target(
    native_session: SimpleNamespace,
) -> None:
    payload = b"\x89PNG\r\n\x1a\nimage"

    def export(file_path: str, image_format: str) -> None:
        Path(file_path).write_bytes(payload)

    native_session.store.items[0].Export.side_effect = export
    document = native_session.documents[0].FullName
    assert (
        presentation_export.get_slide_preview(
            presentation_export.GetSlidePreviewInput(document=document)
        )
        == payload
    )
    native_session.get_document.assert_called_once_with(document)


def test_placeholder_reports_collection_index_and_prefers_exact_type(
    native_session: SimpleNamespace,
) -> None:
    centered = Mock()
    centered.PlaceholderFormat.Type = 3
    title = Mock()
    title.PlaceholderFormat.Type = 1
    collection = Mock(
        side_effect=lambda index: [centered, title][index - 1], Count=2
    )
    slide = native_session.store.items[0]
    slide.Shapes.Placeholders = collection
    assert placeholders.resolve_placeholder(slide, None, "title") == (2, title)
    title.HasTextFrame = False
    title.Name = "title"
    title.Left, title.Top, title.Width, title.Height, title.Rotation = (
        1,
        2,
        3,
        4,
        0,
    )
    result = placeholders.get_placeholder(
        placeholders.GetPlaceholderInput(
            slide_index=1,
            placeholder_type="title",
        )
    )
    assert result["placeholder"]["index"] == 2
    assert result["placeholder"]["type"] == 1


def test_placeholder_normalizes_crlf_without_doubling_paragraphs(
    native_session: SimpleNamespace,
) -> None:
    placeholder = Mock(HasTextFrame=True)
    slide = native_session.store.items[0]
    slide.Shapes.Placeholders = Mock(return_value=placeholder, Count=1)
    placeholders.set_placeholder_text(
        placeholders.SetPlaceholderTextInput(
            slide_index=1,
            placeholder_index=1,
            text="  first\r\nsecond\nthird  ",
        )
    )
    assert placeholder.TextFrame.TextRange.Text == "  first\rsecond\rthird  "


def test_properties_use_snake_case_keys_and_preserve_whitespace(
    native_session: SimpleNamespace,
) -> None:
    result = presentation_document.set_properties(
        presentation_document.SetPropertiesInput(title="  Title  ")
    )
    assert result["set_names"] == ["title"]
    assert (
        native_session.documents[0].BuiltInDocumentProperties("Title").Value
        == "  Title  "
    )


def test_section_invalid_destination_does_not_mutate(
    native_session: SimpleNamespace,
) -> None:
    section_properties = native_session.documents[0].SectionProperties
    section_properties.Count = 2
    with pytest.raises(ValueError, match="destination"):
        sections.manage_section(
            sections.ManageSectionInput(
                section_index=1,
                action="move",
                move_to_index=3,
            )
        )
    section_properties.Move.assert_not_called()


def test_empty_sections_have_null_slide_boundaries(
    native_session: SimpleNamespace,
) -> None:
    section_properties = native_session.documents[0].SectionProperties
    section_properties.Count = 1
    section_properties.FirstSlide.return_value = -1
    section_properties.SlidesCount.return_value = 0
    result = sections.list_sections(DocumentInput())
    section = result["sections"][0]
    assert section["first_slide"] is None and section["last_slide"] is None
    assert section["empty"] is True


def test_save_routes_to_shared_lifecycle(
    native_session: SimpleNamespace,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    save = Mock(return_value={"saved": True})
    monkeypatch.setattr(presentation_document, "save_document", save)
    presentation_document.save_presentation(
        presentation_document.PresentationInput(presentation_index=1)
    )
    assert (
        save.call_args.args[0].document == native_session.documents[0].FullName
    )
    assert save.call_args.args[0].component == "wpp"


def test_pdf_save_as_exports_without_native_save_as(
    native_session: SimpleNamespace,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    export = Mock(return_value={"file_path": "output.pdf"})
    monkeypatch.setattr(presentation_document, "export_pdf", export)
    presentation_document.save_presentation_as(
        presentation_document.SavePresentationAsInput(
            file_path="output.pdf",
            format="pdf",
        )
    )
    assert (
        export.call_args.args[0].document
        == native_session.documents[0].FullName
    )
    native_session.documents[0].SaveAs.assert_not_called()


def test_missing_templates_directory_raises(tmp_path: Path) -> None:
    with pytest.raises(NotADirectoryError):
        presentation_document.list_templates(
            presentation_document.ListTemplatesInput(
                templates_dir=str(tmp_path / "missing"),
            )
        )


def test_template_discovery_includes_uppercase_extensions(
    tmp_path: Path,
) -> None:
    (tmp_path / "b.POTX").write_bytes(b"template")
    (tmp_path / "a.dpt").write_bytes(b"template")
    (tmp_path / "skip.pptx").write_bytes(b"presentation")
    result = presentation_document.list_templates(
        presentation_document.ListTemplatesInput(templates_dir=str(tmp_path))
    )
    assert [template["name"] for template in result["templates"]] == ["a", "b"]


@pytest.mark.parametrize(
    "layout, expected", [("blank", "blank"), ("12", None)]
)
def test_layout_name_normalization_does_not_mutate_input(
    layout: str,
    expected: str | None,
) -> None:
    data = {"layout": layout}
    params = slides.AddSlideInput.model_validate(data)
    assert params.layout_name == expected
    assert data == {"layout": layout}


def test_add_slide_inherits_exact_layout_before_inserting(
    native_session: SimpleNamespace,
) -> None:
    source = native_session.store.items[2]
    source.CustomLayout.Name = "Custom"
    source.Design.Name = "Design"

    def add(**options: object) -> Mock:
        assert options["pCustomLayout"] is source.CustomLayout
        index = options["Index"]
        assert isinstance(index, int)
        created = native_session.store.insert(index)
        created.Layout = 12
        return created

    native_session.store.native.AddSlide.side_effect = add
    result = slides.add_slide(
        slides.AddSlideInput(
            position=1,
            like_slide_index=3,
            layout_name="ignored",
            count=2,
        )
    )
    assert result["slides_created"] == 2
    assert result["layout_name"] == "Custom"
    assert result["design_name"] == "Design"
    assert source.SlideIndex == 5
    native_session.store.native.Add.assert_not_called()


def test_add_slide_ambiguous_layout_prefers_default_master(
    native_session: SimpleNamespace,
) -> None:
    document = native_session.documents[0]
    designs = [Mock(), Mock()]
    for index, design in enumerate(designs):
        design.Name = f"design-{index}"
        design.SlideMaster.Name = f"master-{index}"
        design.SlideMaster.CustomLayouts.Count = 1
        design.SlideMaster.CustomLayouts(1).Name = "Custom"
    document.Designs = Mock(
        side_effect=lambda index: designs[index - 1], Count=2
    )
    document.SlideMaster.Name = "master-1"

    def add(**options: object) -> Mock:
        assert options["pCustomLayout"] is designs[
            1
        ].SlideMaster.CustomLayouts(1)
        index = options["Index"]
        assert isinstance(index, int)
        created = native_session.store.insert(index)
        created.Layout = 12
        return created

    document.Slides.AddSlide.side_effect = add
    result = slides.add_slide(slides.AddSlideInput(layout_name="Custom"))
    assert result["layout_ambiguous"] is True
    assert result["design_name"] == "design-1"
    assert "design-0" in result["warning"]


def test_invalid_notes_color_is_rejected_before_text_changes(
    native_session: SimpleNamespace,
) -> None:
    with pytest.raises(ValueError):
        slides.set_slide_notes(
            slides.SetSlideNotesInput(
                slide_index=1,
                notes_text="new text",
                color="invalid",
            )
        )
    native_session.get_document.assert_not_called()


def test_image_save_as_routes_to_shared_export(
    native_session: SimpleNamespace,
    monkeypatch: pytest.MonkeyPatch,
) -> None:

    export_images = Mock(return_value={"files": ["images/slide_1.png"]})
    monkeypatch.setattr(presentation_export, "export_images", export_images)
    result = presentation_document.save_presentation_as(
        presentation_document.SavePresentationAsInput(
            file_path="images",
            format="png",
        )
    )
    assert result == {"files": ["images/slide_1.png"]}
    forwarded = export_images.call_args.args[0]
    assert forwarded.format == "png" and forwarded.output_dir == "images"
    assert forwarded.document == native_session.documents[0].FullName
    native_session.documents[0].SaveAs.assert_not_called()


def test_inconsistent_save_as_format_is_rejected_before_save(
    native_session: SimpleNamespace,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    save = Mock()
    monkeypatch.setattr(presentation_document, "save_document", save)
    with pytest.raises(ValueError, match="format must match"):
        presentation_document.save_presentation_as(
            presentation_document.SavePresentationAsInput(
                file_path="output.pptx",
                format="dps",
            )
        )
    save.assert_not_called()


def test_active_window_info_respects_explicit_document(
    native_session: SimpleNamespace,
) -> None:
    document = native_session.documents[0]
    document.Windows.Count = 1
    window = document.Windows(1)
    window.Caption = "Presentation"
    window.ViewType = 9
    window.View.Slide.SlideIndex = 3
    window.Selection.Type = 0
    result = app.get_active_window_info(
        DocumentInput(document=document.FullName)
    )
    assert result["active_slide_index"] == 3
    native_session.get_document.assert_called_once_with(document.FullName)
