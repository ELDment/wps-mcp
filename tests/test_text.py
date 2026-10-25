"""Text behavior and palette guarantees without an installed WPS process."""

from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, PropertyMock

import pytest
from pydantic import ValidationError

from src.models import NativeObject
from src.presentation import (
    text,
    text_formatting,
    text_markdown,
    text_measurement,
    text_models,
    text_spans,
    themes,
    typography,
)
from src.presentation.constants import MSO_FALSE, MSO_TRUE


def make_text_range(content: str) -> NativeObject:
    """Emulate WPS UTF-16 offsets without production conversion helpers."""
    native = Mock()
    state = [content]
    writes: list[tuple[int, int, str]] = []

    def whole_text(*values: str) -> str:
        if values:
            state[0] = values[0]
        return state[0]

    type(native).Text = PropertyMock(side_effect=whole_text)
    type(native).Length = PropertyMock(
        side_effect=lambda: (
            len(state[0].encode("utf-16-le", errors="surrogatepass")) // 2
        )
    )
    native.Start = 1
    native.writes = writes
    native.Font = SimpleNamespace(
        Name="Old",
        NameFarEast="Old East",
        Size=20.0,
        Bold=MSO_FALSE,
        Italic=MSO_FALSE,
        Underline=MSO_FALSE,
        Color=SimpleNamespace(RGB=0, ObjectThemeColor=0),
    )
    native.Paragraphs.return_value = SimpleNamespace(Count=1)

    def characters(*positions: int, **named: int) -> NativeObject:
        start = named.get("Start", positions[0] if positions else 1)
        length = named.get("Length", positions[1] if len(positions) > 1 else 0)
        span = Mock()
        span.Start, span.Length, span.Font = start, length, native.Font

        def span_text(*values: str) -> str:
            encoded = state[0].encode("utf-16-le", errors="surrogatepass")
            begin, end = (start - 1) * 2, (start - 1 + length) * 2
            if values:
                writes.append((start, length, values[0]))
                encoded = (
                    encoded[:begin]
                    + values[0].encode("utf-16-le", errors="surrogatepass")
                    + encoded[end:]
                )
                state[0] = encoded.decode("utf-16-le", errors="surrogatepass")
            return encoded[begin:end].decode(
                "utf-16-le", errors="surrogatepass"
            )

        def insert_before(value: str) -> None:
            encoded = state[0].encode("utf-16-le", errors="surrogatepass")
            begin = (start - 1) * 2
            writes.append((start, 0, value))
            state[0] = (
                encoded[:begin] + value.encode("utf-16-le") + encoded[begin:]
            ).decode("utf-16-le", errors="surrogatepass")

        type(span).Text = PropertyMock(side_effect=span_text)
        span.InsertBefore.side_effect = insert_before
        return span

    def find(
        needle: str, after: int, match_case: int, whole_words: int
    ) -> NativeObject:
        haystack = state[0] if match_case == MSO_TRUE else state[0].lower()
        lookup = needle if match_case == MSO_TRUE else needle.lower()
        encoded = haystack.encode("utf-16-le")
        if after >= len(encoded) // 2:
            raise RuntimeError("WPS Find rejects After >= Length")
        python_after = len(encoded[: after * 2].decode("utf-16-le"))
        index = haystack.find(lookup, python_after)
        if index == -1:
            return None
        return SimpleNamespace(
            Start=len(haystack[:index].encode("utf-16-le")) // 2 + 1,
            Length=len(needle.encode("utf-16-le")) // 2,
        )

    native.Characters.side_effect = characters
    native.Find.side_effect = find
    return native


def make_shape(content: str) -> NativeObject:
    return SimpleNamespace(
        Name="Body",
        HasTextFrame=MSO_TRUE,
        TextFrame=SimpleNamespace(TextRange=make_text_range(content)),
    )


def test_text_models_preserve_literal_whitespace() -> None:
    params = text_models.SetTextInput(
        slide_index=1, shape_name_or_index=" Body ", text="  x \n"
    )
    assert params.text == "  x \n"
    assert params.shape_name_or_index == " Body "
    assert text_models.RunSpec(text=" \n").text == " \n"
    assert (
        text_models.FindReplaceTextInput(
            find_text=" ", replace_text="  "
        ).find_text
        == " "
    )
    assert (
        text_models.TextRangeSpec(search_text=" ", bold=True).search_text
        == " "
    )


@pytest.mark.parametrize("value", [float("nan"), float("inf"), -1.0, 0.0])
def test_font_size_rejects_nonfinite_or_nonpositive_values(
    value: float,
) -> None:
    with pytest.raises(ValidationError):
        text_models.TextFormatSpec(font_size=value)


@pytest.mark.parametrize(
    "fields",
    [
        {"search_text": "", "bold": True},
        {"start": 0, "length": 1},
        {"start": 1, "length": -1},
        {"search_text": "x", "start": 1, "length": 1},
        {"start": 1, "length": 1, "occurrence": 2},
        {"start": 1},
    ],
)
def test_span_validation_rejects_ambiguous_or_invalid_input(
    fields: dict[str, object],
) -> None:
    with pytest.raises(ValidationError):
        text_models.TextRangeSpec.model_validate(fields)


def test_find_rejects_legacy_scope_field() -> None:
    with pytest.raises(ValidationError):
        text_models.FindReplaceTextInput.model_validate(
            {"find_text": "x", "slide_index": 2}
        )


@pytest.mark.parametrize(
    ("text_value", "needle", "match_case", "whole_words", "expected"),
    [
        ("😀a😀a", "a", True, False, [(3, 1), (6, 1)]),
        ("Cat cat concatenate", "cat", False, True, [(1, 3), (5, 3)]),
        ("Cat cat concatenate", "cat", True, True, [(5, 3)]),
        ("a.b axb", "a.b", True, False, [(1, 3)]),
        ("中文😀", "missing", False, False, []),
    ],
)
def test_literal_search_preserves_unicode_offsets_and_boundaries(
    text_value: str,
    needle: str,
    match_case: bool,
    whole_words: bool,
    expected: list[tuple[int, int]],
) -> None:
    assert (
        text_spans.find_spans(
            text_value, needle, match_case=match_case, whole_words=whole_words
        )
        == expected
    )


def test_colors_are_checked_before_native_access() -> None:
    for fields in (
        {"color": "unknown"},
        {"highlight_color": "clear"},
        {"font_color_theme": "missing"},
    ):
        with pytest.raises(ValidationError):
            text_models.TextFormatSpec.model_validate(fields)


def test_resolve_span_preserves_spaces_and_normalizes_paragraph_breaks() -> (
    None
):
    assert text_spans.resolve_span(
        "x x x", "Body", text_models.SpanSpec(search_text=" x", occurrence=2)
    ) == (
        4,
        2,
    )
    assert text_spans.resolve_span(
        "first\rsecond",
        "Body",
        text_models.SpanSpec(search_text="first\nsecond"),
    ) == (1, 12)
    assert text_spans.resolve_span(
        "abc", "Body", text_models.SpanSpec(start=4, length=0)
    ) == (4, 0)
    with pytest.raises(ValueError, match="exceeds text length"):
        text_spans.resolve_span(
            "abc", "Body", text_models.SpanSpec(start=3, length=2)
        )
    with pytest.raises(ValueError, match="matches"):
        text_spans.resolve_span(
            "abc", "Body", text_models.SpanSpec(search_text="missing")
        )


def test_run_offsets_account_for_crlf_and_soft_breaks() -> None:
    runs = [
        text_models.RunSpec(text=" a\r\n"),
        text_models.RunSpec(text="b\v "),
    ]
    assert text_spans.run_offsets(runs) == [(1, 3), (4, 3)]


def test_set_text_keeps_spaces_and_paragraph_separators(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    shape = make_shape("old")
    monkeypatch.setattr(
        text,
        "_text_frame_of",
        lambda params: (shape, shape.TextFrame.TextRange),
    )
    result = text.set_text(
        text_models.SetTextInput(
            slide_index=1, shape_name_or_index="Body", text=" a\r\nb \v"
        )
    )
    assert shape.TextFrame.TextRange.Text == " a\rb \v"
    assert result["text_length"] == 6


def test_runs_keep_separate_boundary_breaks(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    shape = make_shape("old")
    monkeypatch.setattr(
        text,
        "_text_frame_of",
        lambda params: (shape, shape.TextFrame.TextRange),
    )
    result = text.set_text(
        text_models.SetTextInput(
            slide_index=1,
            shape_name_or_index="Body",
            runs=[
                text_models.RunSpec(text="a\r"),
                text_models.RunSpec(text="\nb"),
            ],
        )
    )
    assert shape.TextFrame.TextRange.Text == "a\r\rb"
    assert [item["start"] for item in result["runs"]] == [1, 3]
    assert [item["length"] for item in result["runs"]] == [2, 2]


def test_formatting_invalid_later_range_leaves_base_unmodified(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    shape = make_shape("a b c")
    monkeypatch.setattr(
        text,
        "_text_frame_of",
        lambda params: (shape, shape.TextFrame.TextRange),
    )
    apply = Mock()
    monkeypatch.setattr(text, "apply_text_format", apply)
    params = text_models.FormatTextRangeInput(
        slide_index=1,
        shape_name_or_index="Body",
        base=text_models.TextFormatSpec(bold=True),
        ranges=[
            text_models.TextRangeSpec(search_text="a", italic=True),
            text_models.TextRangeSpec(start=5, length=3),
        ],
    )
    with pytest.raises(ValueError, match="exceeds text length"):
        text.format_text_range(params)
    apply.assert_not_called()


def test_font_names_and_boolean_flags_apply_consistently() -> None:
    shape = make_shape("Hello")
    result = text_formatting.apply_text_format(
        shape,
        text_models.TextFormatSpec(
            font_name="Aptos",
            font_name_fareast="Microsoft YaHei",
            bold=False,
            italic=True,
            font_size=1.0,
        ),
    )
    font = shape.TextFrame.TextRange.Font
    assert (font.Name, font.NameFarEast) == ("Aptos", "Microsoft YaHei")
    assert (font.Bold, font.Italic) == (MSO_FALSE, MSO_TRUE)
    assert isinstance(result["warnings"], list)


def test_unavailable_highlight_api_does_not_change_font() -> None:
    shape = make_shape("Hello")
    with pytest.raises(AttributeError):
        text_formatting.apply_text_format(
            shape,
            text_models.TextFormatSpec(bold=True, highlight_color="#FFFF00"),
        )
    assert shape.TextFrame.TextRange.Font.Bold == MSO_FALSE


@pytest.mark.parametrize(
    ("content", "replacement", "expected", "count"),
    [
        ("aaaa", "", "", 4),
        ("aa", "aaa", "aaaaaa", 2),
        ("a a", "x", "x x", 2),
    ],
)
def test_replace_original_matches_once(
    monkeypatch: pytest.MonkeyPatch,
    content: str,
    replacement: str,
    expected: str,
    count: int,
) -> None:
    shape = make_shape(content)
    slide = SimpleNamespace(Shapes=Mock(return_value=shape))
    slide.Shapes.Count = 1
    slides = Mock(return_value=slide)
    slides.Count = 1
    session = SimpleNamespace(
        get_document=lambda: SimpleNamespace(Slides=slides)
    )
    monkeypatch.setattr(text, "get_session", lambda component: session)
    result = text.find_replace_text(
        text_models.FindReplaceTextInput(
            find_text="a", replace_text=replacement
        )
    )
    assert shape.TextFrame.TextRange.Text == expected
    assert result["match_count"] == count
    positions = [
        start
        for start, length, replacement in shape.TextFrame.TextRange.writes
    ]
    assert positions == sorted(positions, reverse=True)


def test_find_dry_run_preserves_text_and_reports_context(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    shape = make_shape(" a a ")
    slide = SimpleNamespace(Shapes=Mock(return_value=shape))
    slide.Shapes.Count = 1
    slides = Mock(return_value=slide)
    slides.Count = 1
    session = SimpleNamespace(
        get_document=lambda: SimpleNamespace(Slides=slides)
    )
    monkeypatch.setattr(text, "get_session", lambda component: session)
    result = text.find_replace_text(
        text_models.FindReplaceTextInput(
            find_text="a",
            replace_text="",
            dry_run=True,
            context_chars=1,
        )
    )
    assert result["mode"] == "find"
    assert result["matches"][0]["context"] == " [a] "
    assert shape.TextFrame.TextRange.Text == " a a "
    assert not shape.TextFrame.TextRange.writes


@pytest.mark.parametrize(
    ("orientation", "width", "height", "expected"),
    [
        ("horizontal", 100.0, 80.6, True),
        ("horizontal", 100.0, 80.4, False),
        ("vertical", 80.6, 100.0, True),
        ("vertical", 80.4, 100.0, False),
    ],
)
def test_measurement_uses_flow_axis_and_rounding_tolerance(
    orientation: str,
    width: float,
    height: float,
    expected: bool,
) -> None:
    result = text_measurement.build_measurement(
        [],
        width,
        height,
        100.0,
        100.0,
        {"left": 10.0, "right": 10.0, "top": 10.0, "bottom": 10.0},
        True,
        "none",
        orientation,
    )
    assert result["overflows"] is expected


def test_missing_bounds_never_claim_text_fits() -> None:
    result = text_measurement.build_measurement(
        [], None, None, 100.0, 100.0, None, True, "none", "horizontal"
    )
    assert result["overflows"] is None


@pytest.mark.parametrize(
    "primary",
    ["#FFFFFF", "#000000", "#FF0000", "#00FF00", "#0000FF", "#abcdef"],
)
def test_generated_palette_accents_have_required_contrast(
    primary: str,
) -> None:
    palette = themes.generate_palette_from_primary(primary)
    assert len(palette) == 10
    assert all(
        themes.contrast_ratio(palette[f"accent{index}"], "#FFFFFF") >= 3.0
        for index in range(1, 7)
    )


def test_all_seventeen_presets_meet_stated_contrast() -> None:
    assert len(themes.PRESET_PALETTES) == 17
    for palette in themes.PRESET_PALETTES.values():
        for index in range(1, 7):
            assert (
                themes.contrast_ratio(
                    palette[f"accent{index}"], palette["light1"]
                )
                >= 3.0
            )


def test_theme_invalid_color_fails_before_connecting(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    session = Mock()
    monkeypatch.setattr(themes, "get_session", session)
    with pytest.raises(ValueError):
        themes.set_theme_colors(
            themes.SetThemeColorsInput(accent1="#123456", accent2="invalid")
        )
    session.assert_not_called()


def test_theme_layers_and_all_designs(monkeypatch: pytest.MonkeyPatch) -> None:
    schemes = [Mock(), Mock()]
    slots = [
        {index: SimpleNamespace(RGB=0) for index in range(1, 13)}
        for _ in schemes
    ]
    for scheme, colors in zip(schemes, slots, strict=True):
        scheme.side_effect = colors.__getitem__
    designs = Mock(
        side_effect=lambda index: SimpleNamespace(
            SlideMaster=SimpleNamespace(
                Theme=SimpleNamespace(ThemeColorScheme=schemes[index - 1])
            ),
        )
    )
    designs.Count = 2
    session = SimpleNamespace(
        get_document=lambda: SimpleNamespace(Designs=designs)
    )
    monkeypatch.setattr(themes, "get_session", lambda component: session)
    result = themes.set_theme_colors(
        themes.SetThemeColorsInput(
            primary="#FF0000",
            preset="corporate_blue",
            accent1="#010203",
        )
    )
    assert result["designs_updated"] == 2
    for colors in slots:
        assert themes.int_to_hex(colors[5].RGB) == "#010203"
        assert (
            themes.int_to_hex(colors[6].RGB)
            == themes.PRESET_PALETTES["corporate_blue"]["accent2"]
        )


def test_unknown_theme_and_empty_update_raise() -> None:
    with pytest.raises(ValueError, match="Unknown preset"):
        themes.set_theme_colors(themes.SetThemeColorsInput(preset="missing"))
    with pytest.raises(ValueError, match="No colors"):
        themes.set_theme_colors(themes.SetThemeColorsInput())


def typography_session(
    monkeypatch: pytest.MonkeyPatch, shape: NativeObject
) -> None:
    shapes = Mock(return_value=shape)
    shapes.Count = 1
    slides = Mock(return_value=SimpleNamespace(Shapes=shapes))
    slides.Count = 1
    presentation = SimpleNamespace(
        Slides=slides,
        PageSetup=SimpleNamespace(SlideWidth=100.0, SlideHeight=100.0),
    )
    session = SimpleNamespace(
        get_document=lambda: presentation, get_application=Mock()
    )
    monkeypatch.setattr(text, "get_session", lambda component: session)
    monkeypatch.setattr(text, "goto_slide", lambda application, index: None)
    monkeypatch.setattr(typography, "get_session", lambda component: session)
    monkeypatch.setattr(
        typography, "goto_slide", lambda application, index: None
    )


def test_typography_restores_autosize_after_measurement_failure(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    shape = make_shape("Body")
    shape.Width, shape.Height, shape.Left, shape.Top = 50.0, 30.0, 0.0, 0.0
    bounds = Mock()
    type(bounds).BoundHeight = PropertyMock(
        side_effect=RuntimeError("bounds unavailable")
    )
    shape.TextFrame2 = SimpleNamespace(
        AutoSize=typography.PP_AUTO_SIZE_TEXT_TO_FIT_SHAPE,
        TextRange=bounds,
        MarginTop=0.0,
        MarginBottom=0.0,
    )
    typography_session(monkeypatch, shape)
    monkeypatch.setattr(
        typography, "_short_lines", lambda current, characters, words: []
    )
    result = typography.check_typography(text_models.CheckTypographyInput())
    assert (
        shape.TextFrame2.AutoSize == typography.PP_AUTO_SIZE_TEXT_TO_FIT_SHAPE
    )
    assert result["issues"][0]["type"] == "measurement_unavailable"


def test_typography_restores_width_after_fix_measurement_failure(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    shape = make_shape("Body")
    shape.Width, shape.Height, shape.Left, shape.Top = 50.0, 30.0, 0.0, 0.0
    shape.TextFrame2 = SimpleNamespace(
        AutoSize=typography.PP_AUTO_SIZE_NONE,
        TextRange=SimpleNamespace(BoundHeight=20.0),
        MarginTop=0.0,
        MarginBottom=0.0,
    )
    typography_session(monkeypatch, shape)
    check_lines = Mock(
        side_effect=[
            [{"type": "widow", "line_text": "a"}],
            RuntimeError("line read failed"),
        ]
    )
    monkeypatch.setattr(typography, "_short_lines", check_lines)
    with pytest.raises(RuntimeError, match="line read failed"):
        typography.check_typography(text_models.CheckTypographyInput(fix=True))
    assert shape.Width == 50.0


def test_typography_keeps_width_fix_inside_slide(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    shape = make_shape("Body")
    shape.Width, shape.Height, shape.Left, shape.Top = 50.0, 30.0, 49.0, 0.0
    shape.TextFrame2 = SimpleNamespace(
        AutoSize=typography.PP_AUTO_SIZE_NONE,
        TextRange=SimpleNamespace(BoundHeight=20.0),
        MarginTop=0.0,
        MarginBottom=0.0,
    )
    typography_session(monkeypatch, shape)
    check_lines = Mock(
        side_effect=[
            [{"type": "widow", "line_text": "a"}],
            [],
            [],
        ]
    )
    monkeypatch.setattr(typography, "_short_lines", check_lines)
    result = typography.check_typography(
        text_models.CheckTypographyInput(fix=True)
    )
    assert shape.Width == 51.0
    assert shape.Left + shape.Width == 100.0
    assert result["fixed_count"] == 1


def test_markdown_merges_styles_and_preserves_run_spaces() -> None:
    runs = [
        SimpleNamespace(
            Text=" first ",
            Font=SimpleNamespace(Bold=MSO_TRUE, Italic=MSO_FALSE),
        ),
        SimpleNamespace(
            Text="second ",
            Font=SimpleNamespace(Bold=MSO_TRUE, Italic=MSO_FALSE),
        ),
    ]
    collection = Mock(
        side_effect=lambda index=None: (
            SimpleNamespace(Count=2) if index is None else runs[index - 1]
        )
    )
    assert (
        text_markdown._runs_to_markdown(SimpleNamespace(Runs=collection))
        == "** first second **"
    )


def test_markdown_export_creates_directory_and_uses_crlf(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    presentation = SimpleNamespace(Slides=Mock())
    presentation.Slides.Count = 1
    monkeypatch.setattr(
        text,
        "get_session",
        lambda component: SimpleNamespace(get_document=lambda: presentation),
    )
    monkeypatch.setattr(
        text, "slide_to_markdown", lambda slide, index: " first\nsecond "
    )
    output = tmp_path / "nested" / "deck.md"
    result = text.get_all_text(
        text_models.GetAllTextInput(output_path=str(output))
    )
    assert output.read_bytes() == b" first\r\nsecond "
    assert isinstance(result, dict)
    assert result["slide_count"] == 1


def test_headers_footers_keep_text_and_enable_auto_date(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    settings = SimpleNamespace(
        Footer=SimpleNamespace(Visible=MSO_FALSE, Text=""),
        SlideNumber=SimpleNamespace(Visible=MSO_FALSE),
        DateAndTime=SimpleNamespace(UseFormat=MSO_FALSE, Text=""),
    )
    slides = Mock(return_value=SimpleNamespace(HeadersFooters=settings))
    slides.Count = 1
    monkeypatch.setattr(
        themes,
        "get_session",
        lambda component: SimpleNamespace(
            get_document=lambda: SimpleNamespace(Slides=slides)
        ),
    )
    themes.set_headers_footers(
        themes.SetHeadersFootersInput(footer_text=" date ", date_format=1)
    )
    assert settings.Footer.Text == " date "
    assert settings.Footer.Visible == MSO_TRUE
    assert settings.DateAndTime.UseFormat == MSO_TRUE
    themes.set_headers_footers(
        themes.SetHeadersFootersInput(date_format=1, date_fixed_text=" fixed ")
    )
    assert settings.DateAndTime.UseFormat == MSO_FALSE
    assert settings.DateAndTime.Text == " fixed "


@pytest.mark.parametrize(
    ("needle", "occurrence", "expected"),
    [
        ("😀", 1, (2, 2)),
        ("B", 1, (4, 1)),
        ("中文", 1, (5, 2)),
        ("😀", 2, (7, 2)),
    ],
)
def test_emoji_search_resolves_native_offsets(
    needle: str,
    occurrence: int,
    expected: tuple[int, int],
) -> None:
    assert (
        text_spans.resolve_span(
            "A😀B中文😀",
            "Body",
            text_models.SpanSpec(search_text=needle, occurrence=occurrence),
        )
        == expected
    )


@pytest.mark.parametrize(("start", "length"), [(2, 1), (3, 0), (3, 1)])
def test_explicit_range_cannot_split_surrogate_pair(
    start: int, length: int
) -> None:
    with pytest.raises(ValueError, match="surrogate pair"):
        text_spans.resolve_span(
            "A😀B", "Body", text_models.SpanSpec(start=start, length=length)
        )


def test_utf16_range_can_reach_final_character_and_append() -> None:
    assert text_spans.resolve_span(
        "A😀B", "Body", text_models.SpanSpec(start=4, length=1)
    ) == (4, 1)
    assert text_spans.resolve_span(
        "A😀B", "Body", text_models.SpanSpec(start=5, length=0)
    ) == (5, 0)
    with pytest.raises(ValueError, match="exceeds text length"):
        text_spans.resolve_span(
            "A😀B", "Body", text_models.SpanSpec(start=5, length=1)
        )


@pytest.mark.parametrize("search", [True, False])
def test_set_text_replaces_complete_emoji(
    monkeypatch: pytest.MonkeyPatch, search: bool
) -> None:
    shape = make_shape("A😀B中文")
    monkeypatch.setattr(
        text,
        "_text_frame_of",
        lambda params: (shape, shape.TextFrame.TextRange),
    )
    fields = {"search_text": "😀"} if search else {"start": 2, "length": 2}
    params = text_models.SetTextInput.model_validate(
        {
            "slide_index": 1,
            "shape_name_or_index": "Body",
            "text": "🚀!",
            **fields,
        }
    )
    result = text.set_text(params)
    assert shape.TextFrame.TextRange.Text == "A🚀!B中文"
    assert (
        result["start"],
        result["replaced_length"],
        result["written_length"],
    ) == (2, 2, 3)
    assert result["text_length"] == 7
    assert result["position_unit"] == "UTF-16"


def test_emoji_run_keeps_following_run_format_at_native_position(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    shape = make_shape("old")
    native = shape.TextFrame.TextRange
    monkeypatch.setattr(text, "_text_frame_of", lambda params: (shape, native))
    result = text.set_text(
        text_models.SetTextInput(
            slide_index=1,
            shape_name_or_index="Body",
            runs=[
                text_models.RunSpec(text="A😀", bold=True),
                text_models.RunSpec(text="B中文", color="#123456"),
            ],
        )
    )
    assert native.Text == "A😀B中文"
    assert [(run["start"], run["length"]) for run in result["runs"]] == [
        (1, 3),
        (4, 3),
    ]
    assert native.Characters.call_args_list[0].kwargs == {
        "Start": 1,
        "Length": 3,
    }
    assert native.Characters.call_args_list[1].kwargs == {
        "Start": 4,
        "Length": 3,
    }
    assert native.Characters(Start=4, Length=3).Text == "B中文"
    assert result["text_length"] == 6


def test_format_text_search_after_emoji_targets_correct_characters(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    shape = make_shape("A😀B中文")
    monkeypatch.setattr(
        text,
        "_text_frame_of",
        lambda params: (shape, shape.TextFrame.TextRange),
    )
    result = text.format_text_range(
        text_models.FormatTextRangeInput(
            slide_index=1,
            shape_name_or_index="Body",
            search_text="B中",
            bold=True,
        )
    )
    assert (result["start"], result["length"], result["formatted_text"]) == (
        4,
        2,
        "B中",
    )


@pytest.mark.parametrize(
    ("needle", "start", "length", "context"),
    [
        ("😀", 2, 2, "A[😀]B"),
        ("B", 4, 1, "😀[B]中"),
        ("中", 5, 1, "B[中]文"),
    ],
)
def test_emoji_dry_run_context_uses_complete_python_characters(
    monkeypatch: pytest.MonkeyPatch,
    needle: str,
    start: int,
    length: int,
    context: str,
) -> None:
    shape = make_shape("A😀B中文")
    typography_session(monkeypatch, shape)
    result = text.find_replace_text(
        text_models.FindReplaceTextInput(
            find_text=needle,
            replace_text="",
            dry_run=True,
            context_chars=1,
        )
    )
    match = result["matches"][0]
    assert (match["start"], match["length"], match["context"]) == (
        start,
        length,
        context,
    )
    assert shape.TextFrame.TextRange.Text == "A😀B中文"
    assert not shape.TextFrame.TextRange.writes


@pytest.mark.parametrize(
    ("content", "needle", "replacement", "expected", "spans"),
    [
        ("😀a😀a", "a", "🚀", "😀🚀😀🚀", [(3, 2), (7, 2)]),
        ("😀😀", "😀", "", "", [(1, 0), (1, 0)]),
        ("😀A😀", "😀", "B", "BAB", [(1, 1), (3, 1)]),
        ("A😀B😀C", "😀", "😀😀", "A😀😀B😀😀C", [(2, 4), (7, 4)]),
    ],
)
def test_emoji_replacements_keep_native_offsets_and_context(
    monkeypatch: pytest.MonkeyPatch,
    content: str,
    needle: str,
    replacement: str,
    expected: str,
    spans: list[tuple[int, int]],
) -> None:
    shape = make_shape(content)
    typography_session(monkeypatch, shape)
    result = text.find_replace_text(
        text_models.FindReplaceTextInput(
            find_text=needle,
            replace_text=replacement,
            context_chars=1,
        )
    )
    assert shape.TextFrame.TextRange.Text == expected
    assert [
        (match["start"], match["length"]) for match in result["matches"]
    ] == spans
    assert all(
        f"[{replacement}]" in match["context"] for match in result["matches"]
    )
    assert result["position_unit"] == "UTF-16"


def test_typography_converts_python_break_to_native_offset(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    shape = make_shape("😀ABC中")
    shape.Width, shape.Height, shape.Left, shape.Top = 100.0, 30.0, 0.0, 0.0
    shape.TextFrame2 = SimpleNamespace(
        AutoSize=typography.PP_AUTO_SIZE_NONE,
        TextRange=SimpleNamespace(BoundHeight=20.0),
        MarginTop=0.0,
        MarginBottom=0.0,
    )
    typography_session(monkeypatch, shape)
    line_issue = {
        "type": "widow",
        "line_text": "中",
        "prev_line_text": "😀ABC",
        "previous_start": 1,
    }
    monkeypatch.setattr(
        typography, "_short_lines", Mock(side_effect=[[line_issue], []])
    )
    monkeypatch.setattr(
        typography,
        "_find_best_vbreak",
        lambda previous, widow: (3, "😀AB", "C中"),
    )
    result = typography.check_typography(
        text_models.CheckTypographyInput(fix=True)
    )
    assert shape.TextFrame.TextRange.Text == "😀AB\vC中"
    assert shape.TextFrame.TextRange.writes == [(5, 0, "\v")]
    assert result["fixed_count"] == 1
