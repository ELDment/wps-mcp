"""Verify writer references behavior and native preflight boundaries."""

from collections.abc import Iterator
from contextlib import contextmanager, nullcontext
from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from pydantic import ValidationError

from src.models import DocumentInput
from src.native.context import CALL_DOCUMENT
from src.writer import references


@pytest.fixture
def document(monkeypatch: pytest.MonkeyPatch) -> Mock:
    current = Mock()
    current.FullName = "references.docx"
    current.Content = SimpleNamespace(Start=0, End=100)
    current.Range.return_value = Mock(Start=10, End=10, Text="")
    current.GetCrossReferenceItems.return_value = ("Target A", "Target B")
    current.Fields.Count = 0
    current.Fields.Item.return_value = Mock(
        Type=3,
        Code=SimpleNamespace(Start=11, Text=" REF target "),
        Result=SimpleNamespace(Start=25, End=31, Text="target"),
    )
    current.application = Mock()
    current.application.Selection.Document.FullName = current.FullName

    def insert_reference(*arguments: object) -> None:
        current.Fields.Count = 1

    current.application.Selection.InsertCrossReference.side_effect = (
        insert_reference
    )
    current.session = SimpleNamespace(
        get_document=Mock(return_value=current),
        get_application=Mock(return_value=current.application),
        command_window=Mock(side_effect=nullcontext),
    )
    monkeypatch.setattr(
        references, "get_session", lambda component: current.session
    )
    return current


@pytest.fixture
def toc(document: Mock) -> Mock:
    current = Mock(
        Range=SimpleNamespace(Start=0, End=12, Text="Chapter\t1\r"),
        UpperHeadingLevel=1,
        LowerHeadingLevel=3,
        IncludePageNumbers=True,
        RightAlignPageNumbers=True,
        UseHyperlinks=True,
    )
    document.TablesOfContents.Count = 1
    document.TablesOfContents.Item.return_value = current
    document.TablesOfContents.Add.return_value = current
    return current


@pytest.fixture
def note(document: Mock) -> Mock:
    current = Mock(
        Index=1,
        Reference=SimpleNamespace(Start=7, End=8, Text="1"),
        Range=SimpleNamespace(Start=1, End=8, Text="note\r"),
    )
    for collection in (document.Footnotes, document.Endnotes):
        collection.Count = 1
        collection.Item.return_value = current
        collection.Add.return_value = current
    return current


@pytest.fixture
def field(document: Mock) -> Mock:
    current = Mock(
        Type=3,
        Locked=False,
        Code=SimpleNamespace(Text=" REF target "),
        Result=SimpleNamespace(Start=3, End=9, Text="target"),
    )
    current.Update.return_value = True
    document.Fields.Count = 1
    document.Fields.Item.return_value = current
    return current


def test_toc_add_uses_collapsed_range_and_explicit_document(
    document: Mock, toc: Mock
) -> None:
    result = references.add_toc(
        references.AddTocInput(
            document="D:/report.docx",
            position=10,
            upper_heading_level=2,
            lower_heading_level=4,
            include_page_numbers=False,
            use_outline_levels=True,
        )
    )
    document.session.get_document.assert_called_once_with("D:/report.docx")
    document.Range.assert_called_once_with(10, 10)
    document.TablesOfContents.Add.assert_called_once_with(
        Range=document.Range.return_value,
        UseHeadingStyles=True,
        UpperHeadingLevel=2,
        LowerHeadingLevel=4,
        UseFields=False,
        RightAlignPageNumbers=True,
        IncludePageNumbers=False,
        UseHyperlinks=True,
        UseOutlineLevels=True,
    )
    assert result["start"] == 0
    assert (
        "use_outline_levels" not in result
    )  # WPS exposes it on Add, not on the TOC object.


@pytest.mark.parametrize(
    "values",
    [
        {"upper_heading_level": 3, "lower_heading_level": 2},
        {"upper_heading_level": 0},
        {"lower_heading_level": 10},
        {"lower_heading_level": True},
        {"position": True},
        {"unknown": 1},
    ],
)
def test_toc_model_rejects_invalid_levels_and_positions(
    values: dict[str, object],
) -> None:
    with pytest.raises(ValidationError):
        references.AddTocInput.model_validate({"position": 0} | values)


def test_toc_rejects_position_before_native_write(document: Mock) -> None:
    with pytest.raises(ValueError, match="within"):
        references.add_toc(references.AddTocInput(position=101))
    document.TablesOfContents.Add.assert_not_called()


def test_toc_list_has_one_based_indices(document: Mock, toc: Mock) -> None:
    result = references.list_tocs(DocumentInput(document="chosen.docx"))
    assert result["tables_of_contents"][0]["toc_index"] == 1
    assert result["tables_of_contents"][0]["text"] == "Chapter\t1\r"
    document.TablesOfContents.Item.assert_called_once_with(1)
    document.session.get_document.assert_called_once_with("chosen.docx")


@pytest.mark.parametrize("mode", ["all", "page_numbers"])
def test_toc_update_selects_the_requested_method(
    document: Mock, toc: Mock, mode: str
) -> None:
    params = references.UpdateTocInput.model_validate(
        {"toc_index": 1, "mode": mode}
    )
    result = references.update_toc(params)
    assert result["mode"] == mode
    assert toc.Update.call_count == (mode == "all")
    assert toc.UpdatePageNumbers.call_count == (mode == "page_numbers")
    document.TablesOfContents.Item.assert_called_once_with(1)


def test_toc_rejects_bad_index_before_update(
    document: Mock, toc: Mock
) -> None:
    with pytest.raises(ValueError, match="between 1 and 1"):
        references.update_toc(references.UpdateTocInput(toc_index=2))
    toc.Update.assert_not_called()
    toc.UpdatePageNumbers.assert_not_called()


def test_toc_delete_and_native_failures_propagate(
    document: Mock, toc: Mock
) -> None:
    assert (
        references.delete_toc(references.TocInput(toc_index=1))["deleted"]
        is True
    )
    toc.Delete.assert_called_once_with()
    toc.Update.side_effect = RuntimeError("Native failure")
    with pytest.raises(RuntimeError, match="Native failure"):
        references.update_toc(references.UpdateTocInput(toc_index=1))
    toc.Update.assert_called_once_with()


@pytest.mark.parametrize("kind", ["footnote", "endnote"])
def test_note_add_selects_kind_and_preserves_text(
    document: Mock, note: Mock, kind: str
) -> None:
    params = references.AddNoteInput.model_validate(
        {
            "document": "notes.docx",
            "kind": kind,
            "position": 10,
            "text": "  note 😀\r",
        }
    )
    result = references.add_note(params)
    chosen = document.Footnotes if kind == "footnote" else document.Endnotes
    other = document.Endnotes if kind == "footnote" else document.Footnotes
    chosen.Add.assert_called_once_with(
        Range=document.Range.return_value, Text="  note 😀\r"
    )
    other.Add.assert_not_called()
    assert result["reference"]["start"] == 7
    assert result["content"]["start"] == 1
    assert result["note_index"] == 1
    document.session.get_document.assert_called_once_with("notes.docx")


def test_note_add_custom_reference(document: Mock, note: Mock) -> None:
    references.add_note(
        references.AddNoteInput(
            kind="footnote", position=10, text="custom", reference_mark="*"
        )
    )
    assert document.Footnotes.Add.call_args.kwargs["Reference"] == "*"


def test_note_list_update_delete_preserve_reference(
    document: Mock, note: Mock
) -> None:
    result = references.list_notes(references.NoteKindInput(kind="endnote"))
    assert result["notes"][0]["content"]["text"] == "note\r"
    changed = references.update_note(
        references.UpdateNoteInput(
            kind="endnote", note_index=1, text=" changed \r"
        )
    )
    assert changed["content"]["text"] == " changed \r"
    assert note.Reference.Text == "1"
    assert references.delete_note(
        references.NoteInput(kind="endnote", note_index=1)
    )["deleted"]
    note.Delete.assert_called_once_with()
    document.Footnotes.Item.assert_not_called()


def test_note_invalid_index_and_position_do_not_mutate(
    document: Mock, note: Mock
) -> None:
    with pytest.raises(ValueError, match="between"):
        references.update_note(
            references.UpdateNoteInput(kind="footnote", note_index=2, text="x")
        )
    assert note.Range.Text == "note\r"
    with pytest.raises(ValueError, match="within"):
        references.add_note(
            references.AddNoteInput(kind="footnote", position=101, text="x")
        )
    document.Footnotes.Add.assert_not_called()


@pytest.mark.parametrize(
    "values",
    [
        {"kind": "comment", "note_index": 1},
        {"kind": "footnote", "note_index": 0},
        {"kind": "footnote", "note_index": True},
    ],
)
def test_note_models_reject_non_note_targets(
    values: dict[str, object],
) -> None:
    with pytest.raises(ValidationError):
        references.NoteInput.model_validate(values)


@pytest.mark.parametrize(
    (
        "reference_type",
        "reference_kind",
        "reference_item",
        "native_type",
        "native_kind",
    ),
    [
        ("bookmark", "text", "Target A", 2, -1),
        ("bookmark", "page_number", "Target B", 2, 7),
        ("heading", "text", 1, 1, -1),
        ("heading", "number_full_context", 2, 1, -4),
        ("footnote", "number", 1, 3, 5),
        ("footnote", "number_formatted", 1, 3, 16),
        ("endnote", "number", 1, 4, 6),
        ("endnote", "number_formatted", 1, 4, 17),
    ],
)
def test_cross_reference_insertion_maps_only_supported_native_enums(
    document: Mock,
    reference_type: str,
    reference_kind: str,
    reference_item: int | str,
    native_type: int,
    native_kind: int,
) -> None:
    params = references.InsertReferenceInput.model_validate(
        {
            "document": "references.docx",
            "position": 10,
            "reference_type": reference_type,
            "reference_kind": reference_kind,
            "reference_item": reference_item,
        }
    )
    result = references.insert_reference(params)
    document.session.get_document.assert_called_once_with("references.docx")
    document.GetCrossReferenceItems.assert_called_once_with(native_type)
    (
        document.application.Selection.InsertCrossReference
    ).assert_called_once_with(
        native_type, native_kind, reference_item, True, False, False, " "
    )
    document.Range.return_value.Select.assert_called_once_with()
    document.Range.return_value.InsertCrossReference.assert_not_called()
    (
        document.application.Selection.Range.Duplicate.Select
    ).assert_called_once_with()
    assert document.Repaginate.call_count == (reference_kind == "page_number")
    assert result["inserted"] is True
    assert result["field_index"] == 1
    assert result["text"] == "target"


@pytest.mark.parametrize(
    "values",
    [
        {"reference_type": "macro"},
        {"reference_type": 99},
        {"reference_kind": "DDE"},
        {"reference_type": "footnote", "reference_kind": "text"},
        {"reference_type": "heading", "reference_kind": "number_formatted"},
        {"reference_type": "bookmark", "reference_item": 1},
        {"reference_type": "heading", "reference_item": "1"},
        {"reference_item": True},
        {"reference_item": 0},
        {"field_code": "DDE executable"},
    ],
)
def test_cross_reference_models_reject_invalid_combinations(
    values: dict[str, object],
) -> None:
    defaults = {
        "position": 0,
        "reference_type": "heading",
        "reference_kind": "text",
        "reference_item": 1,
    }
    with pytest.raises(ValidationError):
        references.InsertReferenceInput.model_validate(defaults | values)


@pytest.mark.parametrize(
    ("reference_type", "reference_item"),
    [("bookmark", "Missing"), ("heading", 3)],
)
def test_cross_reference_target_validation_precedes_write(
    document: Mock, reference_type: str, reference_item: int | str
) -> None:
    params = references.InsertReferenceInput.model_validate(
        {
            "position": 10,
            "reference_type": reference_type,
            "reference_kind": "page_number",
            "reference_item": reference_item,
        }
    )
    with pytest.raises(ValueError, match="not found|between"):
        references.insert_reference(params)
    document.application.Selection.InsertCrossReference.assert_not_called()
    document.session.command_window.assert_not_called()
    document.Repaginate.assert_not_called()


def test_reference_targets_keep_native_order_and_bookmark_names(
    document: Mock,
) -> None:
    document.GetCrossReferenceItems.return_value = ("  Chapter 2", "Chapter 1")
    headings = references.list_reference_targets(
        references.ReferenceTargetsInput(reference_type="heading")
    )
    assert headings["targets"] == [
        {"reference_item": 1, "text": "  Chapter 2"},
        {"reference_item": 2, "text": "Chapter 1"},
    ]
    document.GetCrossReferenceItems.return_value = ("BookmarkB", "BookmarkA")
    bookmarks = references.list_reference_targets(
        references.ReferenceTargetsInput(reference_type="bookmark")
    )
    assert bookmarks["targets"][0]["reference_item"] == "BookmarkB"
    document.GetCrossReferenceItems.return_value = None
    document.Endnotes.Count = 0
    assert (
        references.list_reference_targets(
            references.ReferenceTargetsInput(reference_type="endnote")
        )["targets"]
        == []
    )


def test_linux_heading_targets_follow_builtin_style_order(
    document: Mock, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(references, "sys", SimpleNamespace(platform="linux"))
    document.GetCrossReferenceItems.return_value = None
    document.Styles.Item.side_effect = lambda identifier: SimpleNamespace(
        NameLocal=f"标题 {-identifier - 1}"
    )
    paragraphs = [
        SimpleNamespace(
            OutlineLevel=3,
            Range=SimpleNamespace(
                Text="Third first\r", Style=SimpleNamespace(NameLocal="标题 3")
            ),
        ),
        SimpleNamespace(
            OutlineLevel=2,
            Range=SimpleNamespace(
                Text="Custom outline\r",
                Style=SimpleNamespace(NameLocal="正文"),
            ),
        ),
        SimpleNamespace(
            OutlineLevel=1,
            Range=SimpleNamespace(
                Text="First last\r", Style=SimpleNamespace(NameLocal="标题 1")
            ),
        ),
    ]
    document.Paragraphs.Count = len(paragraphs)
    document.Paragraphs.Item.side_effect = lambda index: paragraphs[index - 1]
    assert references.list_reference_targets(
        references.ReferenceTargetsInput(reference_type="heading")
    )["targets"] == [
        {"reference_item": 1, "text": "Third first"},
        {"reference_item": 2, "text": "First last"},
    ]


@pytest.mark.parametrize(
    ("reference_type", "collection_name"),
    [
        ("bookmark", "Bookmarks"),
        ("footnote", "Footnotes"),
        ("endnote", "Endnotes"),
    ],
)
def test_linux_reference_targets_use_native_collection_order(
    document: Mock,
    monkeypatch: pytest.MonkeyPatch,
    reference_type: str,
    collection_name: str,
) -> None:
    monkeypatch.setattr(references, "sys", SimpleNamespace(platform="linux"))
    document.GetCrossReferenceItems.return_value = None
    items = [
        SimpleNamespace(
            Name="Second", Range=SimpleNamespace(Text="  note 2\r")
        ),
        SimpleNamespace(Name="First", Range=SimpleNamespace(Text="note 1\r")),
    ]
    collection = getattr(document, collection_name)
    collection.Count = 2
    collection.Item.side_effect = lambda index: items[index - 1]
    params = references.ReferenceTargetsInput.model_validate(
        {"reference_type": reference_type}
    )
    targets = references.list_reference_targets(params)["targets"]
    assert targets == (
        [
            {"reference_item": "Second", "text": "Second"},
            {"reference_item": "First", "text": "First"},
        ]
        if reference_type == "bookmark"
        else [
            {"reference_item": 1, "text": "  note 2"},
            {"reference_item": 2, "text": "note 1"},
        ]
    )


def test_windows_empty_native_targets_do_not_guess_collection_order(
    document: Mock, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(references, "sys", SimpleNamespace(platform="win32"))
    document.GetCrossReferenceItems.return_value = None
    assert references.reference_targets(document, "heading") == []
    document.Paragraphs.Item.assert_not_called()


def test_linux_native_target_array_takes_precedence_over_collections(
    document: Mock, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(references, "sys", SimpleNamespace(platform="linux"))
    document.GetCrossReferenceItems.return_value = ("Native order",)
    assert references.reference_targets(document, "heading") == [
        "Native order"
    ]
    document.Paragraphs.Item.assert_not_called()


def test_linux_target_lookup_does_not_hide_native_errors(
    document: Mock, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(references, "sys", SimpleNamespace(platform="linux"))
    document.GetCrossReferenceItems.side_effect = RuntimeError(
        "Native failure"
    )
    with pytest.raises(RuntimeError, match="Native failure"):
        references.reference_targets(document, "heading")
    document.GetCrossReferenceItems.assert_called_once_with(1)
    document.Paragraphs.Item.assert_not_called()


@pytest.mark.parametrize("values", ["a single string", (None,), (1,)])
def test_reference_targets_reject_unexpected_native_array(
    document: Mock, values: object
) -> None:
    document.GetCrossReferenceItems.return_value = values
    with pytest.raises(RuntimeError, match="target array"):
        references.list_reference_targets(
            references.ReferenceTargetsInput(reference_type="heading")
        )


def test_reference_list_preserves_document_field_indices(
    document: Mock, field: Mock
) -> None:
    document.Fields.Count = 3
    document.Fields.Item.side_effect = [
        SimpleNamespace(Type=33),
        field,
        SimpleNamespace(Type=45),
    ]
    result = references.list_references(
        DocumentInput(document="references.docx")
    )
    assert result["story"] == "main"
    assert len(result["references"]) == 1
    assert result["references"][0]["field_index"] == 2
    assert result["references"][0]["code"] == " REF target "


def test_reference_update_rejects_non_reference_and_locked_fields(
    document: Mock, field: Mock
) -> None:
    field.Type = 45
    with pytest.raises(ValueError, match="not a supported"):
        references.update_reference(
            references.ReferenceFieldInput(field_index=1)
        )
    field.Type = 3
    field.Locked = True
    with pytest.raises(ValueError, match="locked"):
        references.update_reference(
            references.ReferenceFieldInput(field_index=1)
        )
    field.Update.assert_not_called()
    document.Repaginate.assert_not_called()


def test_reference_page_update_repaginates_and_detects_native_failure(
    document: Mock, field: Mock
) -> None:
    field.Type = 37
    result = references.update_reference(
        references.ReferenceFieldInput(field_index=1)
    )
    assert result["updated"]
    document.Repaginate.assert_called_once_with()
    field.Update.assert_called_once_with()
    field.Update.return_value = False
    with pytest.raises(RuntimeError, match="could not update"):
        references.update_reference(
            references.ReferenceFieldInput(field_index=1)
        )


def test_reference_delete_rejects_arbitrary_fields(
    document: Mock, field: Mock
) -> None:
    field.Type = 59
    with pytest.raises(ValueError, match="not a supported"):
        references.delete_reference(
            references.ReferenceFieldInput(field_index=1)
        )
    field.Delete.assert_not_called()
    field.Type = 72
    assert references.delete_reference(
        references.ReferenceFieldInput(field_index=1)
    )["deleted"]
    field.Delete.assert_called_once_with()


def test_cross_reference_does_not_retry_native_failure(document: Mock) -> None:
    document.application.Selection.InsertCrossReference.side_effect = (
        RuntimeError("Native failure")
    )
    with pytest.raises(RuntimeError, match="Native failure"):
        references.insert_reference(
            references.InsertReferenceInput(
                position=10,
                reference_type="heading",
                reference_kind="text",
                reference_item=1,
            )
        )
    assert document.application.Selection.InsertCrossReference.call_count == 1
    (
        document.application.Selection.Range.Duplicate.Select
    ).assert_called_once_with()


@pytest.mark.parametrize("existing_count", [0, 1])
def test_cross_reference_rejects_silent_noop_even_with_an_existing_reference(
    document: Mock, existing_count: int
) -> None:
    document.Fields.Count = existing_count
    document.application.Selection.InsertCrossReference.side_effect = None
    with pytest.raises(references.WpsError, match="did not create"):
        references.insert_reference(
            references.InsertReferenceInput(
                position=10,
                reference_type="heading",
                reference_kind="text",
                reference_item=1,
            )
        )
    (
        document.application.Selection.Range.Duplicate.Select
    ).assert_called_once_with()


@pytest.mark.parametrize(("field_type", "position"), [(3, 50), (45, 11)])
def test_cross_reference_requires_a_reference_at_the_requested_position(
    document: Mock, field_type: int, position: int
) -> None:
    document.Fields.Item.return_value.Type = field_type
    document.Fields.Item.return_value.Code.Start = position
    with pytest.raises(references.WpsError, match="did not create"):
        references.insert_reference(
            references.InsertReferenceInput(
                position=10,
                reference_type="heading",
                reference_kind="text",
                reference_item=1,
            )
        )
    (
        document.application.Selection.Range.Duplicate.Select
    ).assert_called_once_with()


def test_cross_reference_rejects_selection_from_another_document(
    document: Mock,
) -> None:
    document.application.Selection.Document.FullName = "other.docx"
    with pytest.raises(references.WpsError, match="activate the target"):
        references.insert_reference(
            references.InsertReferenceInput(
                position=10,
                reference_type="heading",
                reference_kind="text",
                reference_item=1,
            )
        )
    document.Range.return_value.Select.assert_not_called()
    document.application.Selection.InsertCrossReference.assert_not_called()


@pytest.mark.parametrize("native_failure", [False, True])
def test_cross_reference_restores_selection_window_and_context_in_order(
    document: Mock, native_failure: bool
) -> None:
    events: list[str] = []

    @contextmanager
    def command_window() -> Iterator[None]:
        assert CALL_DOCUMENT.get() == document.FullName
        events.append("activate_target")
        try:
            yield
        finally:
            events.append("restore_window")

    def insert(*arguments: object) -> None:
        events.append("insert")
        if native_failure:
            raise RuntimeError("Native failure")
        document.Fields.Count = 1

    document.session.command_window = command_window
    document.Range.return_value.Select.side_effect = lambda: events.append(
        "select_target"
    )
    document.application.Selection.InsertCrossReference.side_effect = insert
    document.application.Selection.Range.Duplicate.Select.side_effect = (
        lambda: events.append("restore_selection")
    )
    token = CALL_DOCUMENT.set("original.docx")
    try:
        params = references.InsertReferenceInput(
            document="references.docx",
            position=10,
            reference_type="heading",
            reference_kind="text",
            reference_item=1,
        )
        if native_failure:
            with pytest.raises(RuntimeError, match="Native failure"):
                references.insert_reference(params)
        else:
            assert references.insert_reference(params)["inserted"]
        assert CALL_DOCUMENT.get() == "original.docx"
    finally:
        CALL_DOCUMENT.reset(token)
    assert events == [
        "activate_target",
        "select_target",
        "insert",
        "restore_selection",
        "restore_window",
    ]


@pytest.mark.parametrize("native_failure", [False, True])
def test_cross_reference_reports_restore_failure_without_masking_native_error(
    document: Mock, native_failure: bool, caplog: pytest.LogCaptureFixture
) -> None:
    selection = document.application.Selection
    selection.Range.Duplicate.Select.side_effect = RuntimeError(
        "Restore failure"
    )
    if native_failure:
        selection.InsertCrossReference.side_effect = RuntimeError(
            "Native failure"
        )
    token = CALL_DOCUMENT.set("original.docx")
    try:
        with pytest.raises(
            RuntimeError if native_failure else references.WpsError,
            match="Native failure"
            if native_failure
            else "restore the Writer selection",
        ):
            references.insert_reference(
                references.InsertReferenceInput(
                    position=10,
                    reference_type="heading",
                    reference_kind="text",
                    reference_item=1,
                )
            )
        assert CALL_DOCUMENT.get() == "original.docx"
    finally:
        CALL_DOCUMENT.reset(token)
    if native_failure:
        assert "Restore failure" in caplog.text
