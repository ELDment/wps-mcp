"""Verify writer behavior and native preflight boundaries."""

from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from pydantic import ValidationError

from src.models import DocumentInput
from src.writer import (
    bookmarks,
    images,
    models,
    pages,
    statistics,
    styles,
    tables,
)
from src.writer import text as text_tools


@pytest.fixture
def document(monkeypatch: pytest.MonkeyPatch) -> Mock:
    current = Mock()
    current.Content = SimpleNamespace(Start=0, End=20)
    current.Name = "Document1"
    session = SimpleNamespace(get_document=Mock(return_value=current))
    for module in (text_tools, tables, pages, bookmarks, images, statistics):
        monkeypatch.setattr(module, "get_session", lambda component: session)
    return current


def test_inputs_preserve_whitespace_and_require_strict_positions() -> None:
    params = text_tools.InsertTextInput(position=0, text="  text \r\n")
    assert params.text == "  text \r\n"
    with pytest.raises(ValidationError):
        text_tools.InsertTextInput.model_validate(
            {"position": True, "text": "x"}
        )
    with pytest.raises(ValidationError):
        text_tools.InsertTextInput.model_validate(
            {"position": 0, "text": "x", "unknown": 1}
        )
    with pytest.raises(ValidationError):
        models.CharacterFormat(font_size=float("nan"))


@pytest.mark.parametrize("name", ["Normal", "正文", -1, -2])
def test_style_read_accepts_names_and_builtin_identifiers(
    name: str | int,
) -> None:
    params = styles.ReadStyleInput.model_validate({"name": name})
    assert params.name == name
    assert type(params.name) is type(name)


@pytest.mark.parametrize("name", ["", 0, 1, True, -1.0, None])
def test_style_read_rejects_invalid_or_coerced_identifiers(
    name: object,
) -> None:
    with pytest.raises(ValidationError):
        styles.ReadStyleInput.model_validate({"name": name})


@pytest.mark.parametrize(
    "model",
    [styles.StyleInput, styles.CreateStyleInput, styles.UpdateStyleInput],
)
def test_style_mutations_still_require_names(
    model: type[styles.StyleInput],
) -> None:
    fields: dict[str, object] = {"name": -1}
    if model is styles.UpdateStyleInput:
        fields["character"] = {"bold": True}
    with pytest.raises(ValidationError):
        model.model_validate(fields)


def test_nested_format_rejects_coercion() -> None:
    with pytest.raises(ValidationError):
        text_tools.FormatTextInput.model_validate({"character": {"bold": 1}})
    with pytest.raises(ValidationError):
        text_tools.FormatTextInput(character=models.CharacterFormat())


def test_range_rejects_out_of_bounds_before_com(document: Mock) -> None:
    with pytest.raises(ValueError, match="within"):
        text_tools.read_text(models.RangeInput(start=21))
    document.Range.assert_not_called()


def test_read_uses_utf16_native_offsets(document: Mock) -> None:
    document.Range.return_value = SimpleNamespace(Start=2, End=5, Text="😀\r")
    result = text_tools.read_text(models.RangeInput(start=2, end=5))
    document.Range.assert_called_once_with(2, 5)
    assert result["text"] == "😀\r"
    assert result["position_unit"] == "UTF-16"
    assert result["end_exclusive"] is True


def test_replace_preflights_match_limit(
    monkeypatch: pytest.MonkeyPatch, document: Mock
) -> None:
    matches = [
        {"start": 1, "end": 2, "text": "a"},
        {"start": 4, "end": 5, "text": "a"},
    ]
    monkeypatch.setattr(text_tools, "find_matches", Mock(return_value=matches))
    with pytest.raises(ValueError, match="max_results"):
        text_tools.replace_text(
            text_tools.ReplaceTextInput(
                text="a", replacement="b", max_results=1
            )
        )
    document.Range.assert_not_called()


def test_replace_edits_backwards(
    monkeypatch: pytest.MonkeyPatch, document: Mock
) -> None:
    matches = [
        {"start": 1, "end": 2, "text": "a"},
        {"start": 4, "end": 5, "text": "a"},
    ]
    monkeypatch.setattr(text_tools, "find_matches", Mock(return_value=matches))
    first = SimpleNamespace(Text="a")
    second = SimpleNamespace(Text="a")
    document.Range.side_effect = [second, first]
    result = text_tools.replace_text(
        text_tools.ReplaceTextInput(text="a", replacement="  b  ")
    )
    assert result["replaced_count"] == 2
    assert [call.args for call in document.Range.call_args_list] == [
        (4, 5),
        (1, 2),
    ]
    assert first.Text == second.Text == "  b  "


def test_search_rejects_nonadvancing_native_match(document: Mock) -> None:
    search = SimpleNamespace(End=20)
    candidate = SimpleNamespace(Start=0, End=0, Text="", Find=Mock())
    candidate.Find.Execute.return_value = True
    document.Range.side_effect = [search, candidate]
    with pytest.raises(RuntimeError, match="non-advancing"):
        text_tools.find_text(text_tools.FindTextInput(text="x"))


def test_table_preflights_merged_cells_before_any_write(
    document: Mock,
) -> None:
    first = SimpleNamespace(
        RowIndex=1, ColumnIndex=1, Range=SimpleNamespace(Text="first")
    )
    merged = SimpleNamespace(
        RowIndex=1, ColumnIndex=1, Range=SimpleNamespace(Text="merged")
    )
    table = Mock()
    table.Rows.Count = 1
    table.Columns.Count = 2
    table.Cell.side_effect = [first, merged]
    document.Tables.Count = 1
    document.Tables.Item.return_value = table
    with pytest.raises(ValueError, match="merged cell"):
        tables.write_table(
            tables.WriteTableInput(table_index=1, data=[[0, False]])
        )
    assert first.Range.Text == "first"
    assert merged.Range.Text == "merged"


def test_table_values_and_end_markers() -> None:
    assert tables.cell_text(False) == "false"
    assert tables.cell_text(0) == "0"
    assert tables.cell_text(None) == ""
    assert tables.cell_text(" text \r") == " text \r"
    with pytest.raises(ValidationError, match="end-of-cell"):
        tables.AddTableInput(position=0, data=[["bad\x07"]])
    with pytest.raises(ValidationError, match="rectangular"):
        tables.AddTableInput(position=0, data=[[1, 2], [3]])


def test_read_table_only_removes_native_terminal_marker(
    document: Mock,
) -> None:
    cell = SimpleNamespace(
        RowIndex=1,
        ColumnIndex=1,
        Range=SimpleNamespace(Text=" text \r\r\x07", Start=1, End=10),
    )
    table = Mock()
    table.Range.Cells.Count = 1
    table.Range.Cells.Item.return_value = cell
    document.Tables.Count = 1
    document.Tables.Item.return_value = table
    result = tables.read_table(tables.TableInput(table_index=1))
    assert result["cells"][0]["text"] == " text \r"


def test_page_setup_rejects_impossible_margins_before_mutation(
    document: Mock,
) -> None:
    setup = SimpleNamespace(
        PageWidth=100.0,
        PageHeight=200.0,
        Orientation=0,
        LeftMargin=10.0,
        RightMargin=10.0,
        TopMargin=10.0,
        BottomMargin=10.0,
    )
    document.Sections.Count = 1
    document.Sections.Item.return_value = SimpleNamespace(PageSetup=setup)
    with pytest.raises(ValueError, match="positive content"):
        pages.set_page_setup(pages.PageSetupInput(margin_left=95.0))
    assert setup.LeftMargin == 10.0
    assert setup.Orientation == 0


def test_change_table_uses_snake_case_action(document: Mock) -> None:
    table = Mock()
    table.Rows.Count = 2
    document.Tables.Count = 1
    document.Tables.Item.return_value = table
    result = tables.change_table(
        tables.ChangeTableInput(table_index=1, action="insert_row")
    )
    table.Rows.Add.assert_called_once_with()
    table.Columns.Add.assert_not_called()
    assert result["position"] == 3


def test_read_empty_bookmarks(document: Mock) -> None:
    document.Bookmarks.Count = 0
    assert bookmarks.read_bookmarks(DocumentInput()) == {"bookmarks": []}


@pytest.mark.parametrize("page_width", [None, 150.0])
def test_page_orientation_preflights_rotated_height(
    document: Mock, page_width: float | None
) -> None:
    setup = SimpleNamespace(
        PageWidth=100.0,
        PageHeight=200.0,
        Orientation=0,
        LeftMargin=10.0,
        RightMargin=10.0,
        TopMargin=60.0,
        BottomMargin=60.0,
    )
    document.Sections.Count = 1
    document.Sections.Item.return_value = SimpleNamespace(PageSetup=setup)
    with pytest.raises(ValueError, match="positive content"):
        pages.set_page_setup(
            pages.PageSetupInput(
                orientation="landscape", page_width=page_width
            )
        )
    assert setup.Orientation == 0
    assert setup.PageWidth == 100.0


@pytest.mark.parametrize("variant", ["primary", "first_page", "even_pages"])
def test_first_section_link_is_rejected_before_header_mutations(
    document: Mock, variant: str
) -> None:
    section = Mock()
    section.PageSetup = SimpleNamespace(
        DifferentFirstPageHeaderFooter=0, OddAndEvenPagesHeaderFooter=0
    )
    header = SimpleNamespace(
        LinkToPrevious=False, Range=SimpleNamespace(Text="existing")
    )
    section.Headers.Item.return_value = header
    document.Sections.Count = 1
    document.Sections.Item.return_value = section
    params = pages.WriteHeaderFooterInput.model_validate(
        {
            "location": "header",
            "variant": variant,
            "text": "replacement",
            "link_to_previous": True,
        }
    )
    with pytest.raises(ValueError, match="no preceding section"):
        pages.write_header_footer(params)
    assert vars(section.PageSetup) == {
        "DifferentFirstPageHeaderFooter": 0,
        "OddAndEvenPagesHeaderFooter": 0,
    }
    assert header.Range.Text == "existing"
    assert header.LinkToPrevious is False
