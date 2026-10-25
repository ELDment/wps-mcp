"""Verify writer review behavior and native preflight boundaries."""

from types import SimpleNamespace
from unittest.mock import Mock, PropertyMock, call

import pytest
from pydantic import ValidationError

from src.models import DocumentInput
from src.writer import review
from src.writer import styles as writer_styles
from src.writer.models import CharacterFormat, ParagraphFormat


@pytest.fixture
def document(monkeypatch: pytest.MonkeyPatch) -> Mock:
    current = Mock()
    current.Content = SimpleNamespace(Start=0, End=20)
    current.TrackRevisions = False
    current.Revisions.Count = 0
    current.Comments.Count = 0
    current.Styles.Count = 0
    session = SimpleNamespace(get_document=Mock(return_value=current))
    for module in (review, writer_styles):
        monkeypatch.setattr(module, "get_session", lambda component: session)
    return current


def make_comment(index: int = 1) -> Mock:
    comment = Mock()
    comment.Index = index
    comment.Range = SimpleNamespace(Text=" comment \r")
    comment.Scope = SimpleNamespace(Start=2, End=4, Text="😀")
    comment.Author = "Original Author"
    comment.Initial = "OA"
    comment.Date = "2026-10-11 12:00:00"
    return comment


def make_style(
    name: str = "Custom", kind: int = 1, *, built_in: bool = False
) -> Mock:
    style = Mock()
    style.NameLocal = name
    style.Type = kind
    style.BuiltIn = built_in
    style.InUse = False
    style.BaseStyle = ""
    style.NextParagraphStyle = name
    style.AutomaticallyUpdate = False
    style.Font = SimpleNamespace(
        Name="Arial",
        NameFarEast="宋体",
        Size=12.0,
        Bold=0,
        Italic=0,
        Underline=0,
        Color=0,
    )
    style.ParagraphFormat = SimpleNamespace(
        Alignment=0,
        LeftIndent=0.0,
        RightIndent=0.0,
        FirstLineIndent=0.0,
        SpaceBefore=0.0,
        SpaceAfter=0.0,
        LineSpacing=12.0,
        LineSpacingRule=0,
        KeepWithNext=0,
        KeepTogether=0,
        PageBreakBefore=0,
    )
    return style


@pytest.mark.parametrize(
    "values",
    [
        {},
        {"indices": []},
        {"indices": [1, 1]},
        {"indices": [True]},
        {"indices": [1], "all": True},
    ],
)
def test_revision_selection_requires_an_explicit_unambiguous_target(
    values: dict[str, object],
) -> None:
    with pytest.raises(ValidationError):
        review.ReviewSelectionInput.model_validate(values)


def test_strict_models_reject_unknown_fields_and_empty_updates() -> None:
    with pytest.raises(ValidationError):
        review.RevisionTrackingInput.model_validate({"enabled": 1})
    with pytest.raises(ValidationError):
        review.AddCommentInput.model_validate(
            {"start": 0, "end": 2, "text": "x", "op": "add"}
        )
    with pytest.raises(ValidationError):
        review.UpdateCommentInput(index=1)
    with pytest.raises(ValidationError):
        writer_styles.UpdateStyleInput(
            name="Custom", character=CharacterFormat()
        )
    with pytest.raises(ValidationError):
        review.AddCommentInput(start=3, end=2, text="x")


def test_revision_tracking_does_not_resolve_existing_changes(
    document: Mock,
) -> None:
    document.Revisions.Count = 2
    assert review.revision_state(DocumentInput()) == {
        "enabled": False,
        "count": 2,
    }
    assert review.configure_revisions(
        review.RevisionTrackingInput(enabled=True)
    ) == {
        "enabled": True,
        "count": 2,
    }
    document.Revisions.AcceptAll.assert_not_called()
    document.Revisions.RejectAll.assert_not_called()


def test_revision_list_preserves_native_ranges_and_paginates(
    document: Mock,
) -> None:
    document.Revisions.Count = 3
    document.Revisions.Item.return_value = SimpleNamespace(
        Author="Author",
        Date="2026-10-11",
        Type=1,
        Range=SimpleNamespace(Start=2, End=4, Text="😀"),
    )
    result = review.list_revisions(
        review.ReviewListInput(start_index=2, count=1)
    )
    assert result == {
        "revisions": [
            {
                "index": 2,
                "author": "Author",
                "date": "2026-10-11",
                "type": 1,
                "start": 2,
                "end": 4,
                "text": "😀",
            }
        ],
        "total": 3,
        "has_more": True,
    }
    document.Revisions.Item.assert_called_once_with(2)


def test_revision_batch_rejects_all_invalid_indices_before_native_lookup(
    document: Mock,
) -> None:
    document.Revisions.Count = 2
    with pytest.raises(ValueError, match="revision index"):
        review.accept_revisions(review.ReviewSelectionInput(indices=[1, 3]))
    document.Revisions.Item.assert_not_called()


def test_revision_batch_resolves_every_proxy_before_mutation(
    document: Mock,
) -> None:
    first = Mock()
    document.Revisions.Count = 2
    document.Revisions.Item.side_effect = [
        first,
        RuntimeError("native lookup failed"),
    ]
    with pytest.raises(RuntimeError, match="native lookup failed"):
        review.reject_revisions(review.ReviewSelectionInput(indices=[1, 2]))
    first.Reject.assert_not_called()


def test_selected_revisions_are_processed_in_reverse_original_order(
    document: Mock,
) -> None:
    revisions = [Mock(), Mock(), Mock()]
    document.Revisions.Count = len(revisions)
    document.Revisions.Item.side_effect = lambda index: revisions[index - 1]
    events: list[int] = []
    revisions[0].Accept.side_effect = lambda: events.append(1)
    revisions[2].Accept.side_effect = lambda: events.append(3)
    result = review.accept_revisions(
        review.ReviewSelectionInput(indices=[1, 3])
    )
    assert result == {"count": 2, "decision": "accepted"}
    assert events == [3, 1]
    revisions[1].Accept.assert_not_called()
    assert document.Revisions.Item.call_args_list == [call(3), call(1)]


def test_revision_error_is_propagated_without_replaying_or_continuing(
    document: Mock,
) -> None:
    revisions = [Mock(), Mock()]
    document.Revisions.Count = 2
    document.Revisions.Item.side_effect = lambda index: revisions[index - 1]
    revisions[1].Reject.side_effect = RuntimeError(
        "WPS rejected the operation"
    )
    with pytest.raises(RuntimeError, match="WPS rejected"):
        review.reject_revisions(review.ReviewSelectionInput(indices=[1, 2]))
    revisions[1].Reject.assert_called_once_with()
    revisions[0].Reject.assert_not_called()


@pytest.mark.parametrize("accept", [True, False])
def test_all_revisions_use_native_collection_decision(
    document: Mock, accept: bool
) -> None:
    document.Revisions.Count = 5
    operation = review.accept_revisions if accept else review.reject_revisions
    result = operation(review.ReviewSelectionInput(all=True))
    assert result["count"] == 5
    expected = (
        document.Revisions.AcceptAll
        if accept
        else document.Revisions.RejectAll
    )
    expected.assert_called_once_with()
    document.Revisions.Item.assert_not_called()


def test_add_comment_preflights_range_before_mutation(document: Mock) -> None:
    with pytest.raises(ValueError, match="within"):
        review.add_comment(review.AddCommentInput(start=19, end=21, text="x"))
    document.Comments.Add.assert_not_called()


def test_add_comment_preserves_text_and_changes_only_comment_attribution(
    document: Mock,
) -> None:
    comment = make_comment()
    document.Comments.Add.return_value = comment
    result = review.add_comment(
        review.AddCommentInput(
            start=2, end=4, text=" literal \r", author="Reviewer", initials=""
        )
    )
    document.Range.assert_called_once_with(2, 4)
    document.Comments.Add.assert_called_once_with(
        document.Range.return_value, " literal \r"
    )
    assert result["author"] == "Reviewer"
    assert result["initials"] == ""
    assert result["scope"] == {"start": 2, "end": 4, "text": "😀"}


def test_update_comment_preserves_anchor_and_optional_author(
    document: Mock,
) -> None:
    comment = make_comment()
    document.Comments.Count = 1
    document.Comments.Item.return_value = comment
    result = review.update_comment(
        review.UpdateCommentInput(index=1, text="  replacement  ")
    )
    assert comment.Range.Text == "  replacement  "
    assert result["author"] == "Original Author"
    assert result["scope"] == {"start": 2, "end": 4, "text": "😀"}


def test_comment_delete_validates_entire_batch_before_lookup(
    document: Mock,
) -> None:
    document.Comments.Count = 1
    with pytest.raises(ValueError, match="comment index"):
        review.delete_comments(review.ReviewSelectionInput(indices=[1, 2]))
    document.Comments.Item.assert_not_called()


def test_comment_delete_holds_original_targets_as_indices_shift(
    document: Mock,
) -> None:
    first, second = make_comment(1), make_comment(2)
    remaining = [first, second]
    document.Comments.Count = 2
    document.Comments.Item.side_effect = lambda index: remaining[index - 1]
    first.Delete.side_effect = lambda: remaining.remove(first)
    second.Delete.side_effect = lambda: remaining.remove(second)
    assert review.delete_comments(review.ReviewSelectionInput(all=True)) == {
        "deleted_count": 2,
        "indices": [1, 2],
    }
    assert remaining == []


def test_empty_collections_return_empty_pages(document: Mock) -> None:
    assert review.list_comments(review.ReviewListInput()) == {
        "comments": [],
        "total": 0,
        "has_more": False,
    }
    assert review.list_revisions(review.ReviewListInput(start_index=9)) == {
        "revisions": [],
        "total": 0,
        "has_more": False,
    }


def test_style_list_filters_without_reading_unsupported_formatting(
    document: Mock,
) -> None:
    styles = [
        make_style("Normal", built_in=True),
        make_style("Table", 3),
        make_style("Emphasis", 2),
    ]
    document.Styles.Count = len(styles)
    document.Styles.Item.side_effect = lambda index: styles[index - 1]
    result = writer_styles.list_styles(
        writer_styles.ListStylesInput(kind="character", built_in=False)
    )
    assert result == {
        "styles": [
            {
                "name": "Emphasis",
                "kind": "character",
                "built_in": False,
                "in_use": False,
            }
        ],
        "total": 3,
    }


def test_style_read_preserves_native_undefined_format_sentinel(
    document: Mock,
) -> None:
    style = make_style()
    style.Font.Bold = 9999999
    style.BaseStyle = SimpleNamespace(NameLocal="Normal")
    document.Styles.Item.return_value = style
    result = writer_styles.read_style(
        writer_styles.ReadStyleInput(name="Custom")
    )
    assert result["character"]["bold"] == 9999999
    assert result["base_style"] == "Normal"
    assert result["paragraph"]["line_spacing_rule"] == 0


def test_style_read_handles_empty_native_base_variant(document: Mock) -> None:
    style = make_style("Normal", built_in=True)
    style.BaseStyle = None
    document.Styles.Item.return_value = style
    assert (
        writer_styles.read_style(writer_styles.ReadStyleInput(name="Normal"))[
            "base_style"
        ]
        is None
    )


@pytest.mark.parametrize(
    ("root_id", "local_name", "kind", "style_type"),
    [(-1, "正文", "paragraph", 1), (-66, "默认段落字体", "character", 2)],
)
def test_native_root_style_read_and_inheritance_skip_unavailable_base_property(
    document: Mock, root_id: int, local_name: str, kind: str, style_type: int
) -> None:
    root = make_style(local_name, style_type, built_in=True)
    inaccessible = PropertyMock(
        side_effect=RuntimeError("WPS root BaseStyle access denied")
    )
    type(root).BaseStyle = inaccessible
    document.Styles.Item.side_effect = lambda key: {
        root_id: root,
        local_name: root,
    }[key]
    assert (
        writer_styles.read_style(
            writer_styles.ReadStyleInput(name=local_name)
        )["base_style"]
        is None
    )
    created = make_style("Custom", style_type)
    document.Styles.Add.return_value = created
    writer_styles.create_style(
        writer_styles.CreateStyleInput.model_validate(
            {"name": "Custom", "kind": kind, "base_style": local_name}
        )
    )
    assert created.BaseStyle == local_name
    inaccessible.assert_not_called()
    assert call(root_id) in document.Styles.Item.call_args_list


@pytest.mark.parametrize("built_in", [False, True])
@pytest.mark.parametrize("operation", ["read", "create"])
def test_nonroot_base_property_failures_are_not_suppressed(
    document: Mock, built_in: bool, operation: str
) -> None:
    style = make_style("Heading", built_in=built_in)
    failure = PropertyMock(
        side_effect=RuntimeError("native BaseStyle access failed")
    )
    type(style).BaseStyle = failure
    root = make_style("正文", built_in=True)
    document.Styles.Item.side_effect = lambda key: {
        -1: root,
        "Heading": style,
    }[key]
    with pytest.raises(RuntimeError, match="native BaseStyle access failed"):
        if operation == "read":
            writer_styles.read_style(
                writer_styles.ReadStyleInput(name="Heading")
            )
        else:
            writer_styles.create_style(
                writer_styles.CreateStyleInput(
                    name="Custom", base_style="Heading"
                )
            )
    failure.assert_called_once_with()
    document.Styles.Add.assert_not_called()


def test_custom_style_named_normal_is_not_treated_as_a_native_root(
    document: Mock,
) -> None:
    style = make_style("Normal")
    failure = PropertyMock(
        side_effect=RuntimeError("custom style access failed")
    )
    type(style).BaseStyle = failure
    document.Styles.Item.return_value = style
    with pytest.raises(RuntimeError, match="custom style access failed"):
        writer_styles.read_style(writer_styles.ReadStyleInput(name="Normal"))
    failure.assert_called_once_with()
    document.Styles.Item.assert_called_once_with("Normal")


def test_root_lookup_errors_are_propagated(document: Mock) -> None:
    style = make_style("正文", built_in=True)
    document.Styles.Item.side_effect = [
        style,
        RuntimeError("root lookup failed"),
    ]
    with pytest.raises(RuntimeError, match="root lookup failed"):
        writer_styles.read_style(writer_styles.ReadStyleInput(name="正文"))


def test_style_create_resolves_missing_dependencies_before_add(
    document: Mock,
) -> None:
    document.Styles.Item.side_effect = RuntimeError("Unknown base style")
    with pytest.raises(RuntimeError, match="Unknown base"):
        writer_styles.create_style(
            writer_styles.CreateStyleInput(name="Custom", base_style="Missing")
        )
    document.Styles.Add.assert_not_called()


def test_style_create_rejects_duplicate_local_name(document: Mock) -> None:
    document.Styles.Count = 1
    document.Styles.Item.return_value = make_style("CUSTOM")
    with pytest.raises(ValueError, match="already exists"):
        writer_styles.create_style(
            writer_styles.CreateStyleInput(name="Custom")
        )
    document.Styles.Add.assert_not_called()


def test_character_style_rejects_paragraph_settings_before_creation(
    document: Mock,
) -> None:
    with pytest.raises(ValueError, match="Character styles"):
        writer_styles.create_style(
            writer_styles.CreateStyleInput(
                name="Emphasis",
                kind="character",
                paragraph=ParagraphFormat(alignment="center"),
            )
        )
    document.Styles.Add.assert_not_called()


def test_character_style_rejects_even_an_empty_paragraph_object(
    document: Mock,
) -> None:
    with pytest.raises(ValueError, match="Character styles"):
        writer_styles.create_style(
            writer_styles.CreateStyleInput(
                name="Emphasis", kind="character", paragraph=ParagraphFormat()
            )
        )
    document.Styles.Add.assert_not_called()


def test_style_create_reuses_shared_font_and_paragraph_contract(
    document: Mock,
) -> None:
    base = make_style("Normal", built_in=True)
    style = make_style()
    document.Styles.Item.return_value = base
    document.Styles.Add.return_value = style
    result = writer_styles.create_style(
        writer_styles.CreateStyleInput(
            name="Custom",
            base_style="Normal",
            next_style="Custom",
            automatically_update=False,
            character=CharacterFormat(bold=True, color="#123456"),
            paragraph=ParagraphFormat(alignment="center", line_spacing=18.0),
        )
    )
    document.Styles.Add.assert_called_once_with("Custom", 1)
    assert result["name"] == "Custom"
    assert style.BaseStyle == "Normal"
    assert style.NextParagraphStyle == "Custom"
    assert style.Font.Bold == -1
    assert style.Font.Color == 0x563412
    assert style.ParagraphFormat.Alignment == 1
    assert style.ParagraphFormat.LineSpacingRule == 4
    assert style.ParagraphFormat.LineSpacing == 18.0


def test_style_update_preflights_inheritance_cycle_before_any_write(
    document: Mock,
) -> None:
    original = make_style("Original")
    dependent = make_style("Dependent")
    dependent.BaseStyle = "Original"
    styles = {"Original": original, "Dependent": dependent}
    document.Styles.Item.side_effect = lambda name: styles[name]
    with pytest.raises(ValueError, match="cycle"):
        writer_styles.update_style(
            writer_styles.UpdateStyleInput(
                name="Original",
                base_style="Dependent",
                character=CharacterFormat(bold=True),
            )
        )
    assert original.BaseStyle == ""
    assert original.Font.Bold == 0


def test_style_cycle_validation_uses_resolved_local_name(
    document: Mock,
) -> None:
    original = make_style("正文", built_in=True)
    dependent = make_style("Dependent")
    dependent.BaseStyle = "正文"
    styles = {"Normal": original, "正文": original, "Dependent": dependent}
    document.Styles.Item.side_effect = lambda name: styles[name]
    with pytest.raises(ValueError, match="cycle"):
        writer_styles.update_style(
            writer_styles.UpdateStyleInput(
                name="Normal", base_style="Dependent"
            )
        )
    assert original.BaseStyle == ""


def test_style_update_rejects_incompatible_next_style_before_any_write(
    document: Mock,
) -> None:
    style = make_style()
    next_style = make_style("Emphasis", 2)
    document.Styles.Item.side_effect = [style, next_style]
    with pytest.raises(ValueError, match="paragraph style"):
        writer_styles.update_style(
            writer_styles.UpdateStyleInput(
                name="Custom", next_style="Emphasis", automatically_update=True
            )
        )
    assert style.AutomaticallyUpdate is False


def test_style_update_applies_false_values(document: Mock) -> None:
    style = make_style()
    style.Font.Bold = -1
    style.AutomaticallyUpdate = True
    document.Styles.Item.return_value = style
    writer_styles.update_style(
        writer_styles.UpdateStyleInput(
            name="Custom",
            automatically_update=False,
            character=CharacterFormat(bold=False),
        )
    )
    assert style.Font.Bold == 0
    assert style.AutomaticallyUpdate is False


def test_apply_style_validates_range_and_type_before_assignment(
    document: Mock,
) -> None:
    with pytest.raises(ValueError, match="within"):
        writer_styles.apply_style(
            writer_styles.ApplyStyleInput(name="Custom", end=21)
        )
    document.Styles.Item.assert_not_called()
    text_range = SimpleNamespace(Start=0, End=2, Style="original")
    document.Range.return_value = text_range
    document.Styles.Item.return_value = make_style("Table", 3)
    with pytest.raises(ValueError, match="text range"):
        writer_styles.apply_style(
            writer_styles.ApplyStyleInput(name="Table", end=2)
        )
    assert text_range.Style == "original"


def test_apply_style_uses_resolved_native_object(document: Mock) -> None:
    style = make_style()
    text_range = SimpleNamespace(Start=2, End=4, Style=None)
    document.Styles.Item.return_value = style
    document.Range.return_value = text_range
    assert writer_styles.apply_style(
        writer_styles.ApplyStyleInput(name="Custom", start=2, end=4)
    ) == {
        "name": "Custom",
        "start": 2,
        "end": 4,
    }
    assert text_range.Style is style


def test_builtin_style_cannot_be_deleted(document: Mock) -> None:
    style = make_style("Normal", built_in=True)
    document.Styles.Item.return_value = style
    with pytest.raises(ValueError, match="Built-in"):
        writer_styles.delete_style(writer_styles.StyleInput(name="Normal"))
    style.Delete.assert_not_called()


def test_custom_style_can_be_deleted(document: Mock) -> None:
    style = make_style()
    document.Styles.Item.return_value = style
    assert writer_styles.delete_style(
        writer_styles.StyleInput(name="Custom")
    ) == {
        "name": "Custom",
        "deleted": True,
    }
    style.Delete.assert_called_once_with()
