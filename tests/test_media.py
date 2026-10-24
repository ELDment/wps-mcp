"""Verify presentation media, publication and native boundaries."""

from collections.abc import Iterator
from contextlib import ExitStack, closing, contextmanager
from pathlib import Path
from types import ModuleType, SimpleNamespace
from unittest.mock import Mock

import pypdfium2 as pdfium
import pytest
from PIL import Image
from pydantic import ValidationError

from src.models import DocumentInput, NativeObject
from src.native import clipboard
from src.presentation import (
    animation,
    export,
    history,
    media,
    slideshow,
    smartart,
)


@pytest.fixture
def session(monkeypatch: pytest.MonkeyPatch) -> Mock:
    native = Mock()
    for module in (animation, history, export, media, slideshow, smartart):
        monkeypatch.setattr(module, "get_session", lambda component: native)
    return native


class TrimMedia:
    """Enforce the native requirement that the start must precede the end."""

    def __init__(self, start: int = 100, end: int = 500) -> None:
        self.points = {"StartPoint": start, "EndPoint": end, "Length": 1000}
        self.writes: list[tuple[str, int]] = []

    def __getattr__(self, name: str) -> int:
        if name not in self.points:
            raise AttributeError(name)
        return self.points[name]

    def __setattr__(self, name: str, value: NativeObject) -> None:
        if name not in ("StartPoint", "EndPoint"):
            object.__setattr__(self, name, value)
            return
        candidate = self.points | {name: value}
        if candidate["StartPoint"] >= candidate["EndPoint"]:
            raise ValueError("Invalid intermediate trim interval")
        self.points[name] = value
        self.writes.append((name, value))


def test_media_changes_real_trim_points(
    monkeypatch: pytest.MonkeyPatch,
    session: Mock,
) -> None:
    native = TrimMedia()
    shape = SimpleNamespace(Name="clip", MediaFormat=native)
    monkeypatch.setattr(
        media, "resolve_shape", lambda slide, identifier: shape
    )
    result = media.set_media_settings(
        media.SetMediaSettingsInput(
            slide_index=1,
            shape_name_or_index="clip",
            start_point=700,
            end_point=900,
        )
    )
    assert native.points["StartPoint"] == 700
    assert native.points["EndPoint"] == 900
    assert native.writes == [("EndPoint", 900), ("StartPoint", 700)]
    assert result["success"] is True


def test_media_moves_trim_backwards_without_invalid_intermediate_state(
    monkeypatch: pytest.MonkeyPatch,
    session: Mock,
) -> None:
    native = TrimMedia(700, 900)
    monkeypatch.setattr(
        media,
        "resolve_shape",
        lambda slide, identifier: SimpleNamespace(
            Name="clip",
            MediaFormat=native,
        ),
    )
    media.set_media_settings(
        media.SetMediaSettingsInput(
            slide_index=1,
            shape_name_or_index="clip",
            start_point=100,
            end_point=300,
        )
    )
    assert native.writes == [("StartPoint", 100), ("EndPoint", 300)]


def test_unsupported_media_trim_fails_before_other_changes(
    monkeypatch: pytest.MonkeyPatch,
    session: Mock,
) -> None:
    native = SimpleNamespace(Volume=0.5)
    monkeypatch.setattr(
        media,
        "resolve_shape",
        lambda slide, identifier: SimpleNamespace(
            Name="clip",
            MediaFormat=native,
        ),
    )
    with pytest.raises(
        RuntimeError, match="does not expose playback trimming"
    ):
        media.set_media_settings(
            media.SetMediaSettingsInput(
                slide_index=1,
                shape_name_or_index="clip",
                start_point=100,
                volume=1.0,
            )
        )
    assert native.Volume == 0.5


def test_media_rejects_trim_beyond_duration(
    monkeypatch: pytest.MonkeyPatch,
    session: Mock,
) -> None:
    native = TrimMedia()
    monkeypatch.setattr(
        media,
        "resolve_shape",
        lambda slide, identifier: SimpleNamespace(
            Name="clip",
            MediaFormat=native,
        ),
    )
    with pytest.raises(ValueError, match="1000 ms"):
        media.set_media_settings(
            media.SetMediaSettingsInput(
                slide_index=1,
                shape_name_or_index="clip",
                end_point=1200,
            )
        )
    assert native.writes == []


def test_media_rejects_unapplied_trim(
    monkeypatch: pytest.MonkeyPatch,
    session: Mock,
) -> None:
    class IgnoredTrim(TrimMedia):
        def __setattr__(self, name: str, value: NativeObject) -> None:
            if name not in ("StartPoint", "EndPoint"):
                super().__setattr__(name, value)

    native = IgnoredTrim()
    monkeypatch.setattr(
        media,
        "resolve_shape",
        lambda slide, identifier: SimpleNamespace(
            Name="clip",
            MediaFormat=native,
        ),
    )
    with pytest.raises(RuntimeError, match="could not apply media trimming"):
        media.set_media_settings(
            media.SetMediaSettingsInput(
                slide_index=1,
                shape_name_or_index="clip",
                start_point=200,
            )
        )


def test_media_requires_both_dimensions() -> None:
    with pytest.raises(ValidationError, match="both width and height"):
        media.AddMediaInput(
            kind="video", slide_index=1, file_path="clip.mp4", width=320
        )


def test_animation_conversion_is_never_replayed(session: Mock) -> None:
    effect = Mock(EffectType=1)
    sequence = Mock(Count=2, return_value=effect)
    failure = RuntimeError("native conversion failed after mutation")
    sequence.ConvertToTextUnitEffect.side_effect = failure
    slide = Mock(TimeLine=Mock(MainSequence=sequence))
    session.get_document.return_value.Slides.return_value = slide
    with pytest.raises(RuntimeError) as caught:
        animation.update_animation(
            animation.UpdateAnimationInput(
                slide_index=1,
                animation_index=2,
                text_unit_effect="by_word",
            )
        )
    assert caught.value is failure
    sequence.ConvertToTextUnitEffect.assert_called_once()
    sequence.FindFirstAnimationFor.assert_not_called()


def test_animation_invalid_move_fails_before_effect_mutation(
    session: Mock,
) -> None:
    sequence = Mock(Count=2)
    timeline = session.get_document.return_value.Slides.return_value.TimeLine
    timeline.MainSequence = sequence
    with pytest.raises(ValueError, match="move_to"):
        animation.update_animation(
            animation.UpdateAnimationInput(
                slide_index=1,
                animation_index=2,
                move_to=3,
                duration=4,
            )
        )
    sequence.assert_not_called()


def test_pdf_range_export_preserves_print_options(
    session: Mock, tmp_path: Path
) -> None:
    presentation = session.get_document.return_value
    presentation.Slides.Count = 3

    def render(path: str, format_type: int, **options: NativeObject) -> None:
        with closing(pdfium.PdfDocument.new()) as document:
            for width in (100, 200, 300):
                with closing(document.new_page(width, 100)):
                    pass
            document.save(path)

    presentation.ExportAsFixedFormat.side_effect = render
    target = tmp_path / "selection.pdf"
    result = export.export_pdf(
        export.ExportPDFInput(
            file_path=str(target),
            slide_range_start=2,
            slide_range_end=3,
        )
    )
    with closing(pdfium.PdfDocument(str(target))) as document:
        assert len(document) == 2
        assert document.get_page_size(0) == (200, 100)
    assert result["total_slides"] == 3
    presentation.PrintOptions.Ranges.ClearAll.assert_not_called()
    assert presentation.ExportAsFixedFormat.call_args.kwargs == {
        "PrintHiddenSlides": -1,
        "PrintRange": None,
        "RangeType": 1,
    }
    assert sorted(path.name for path in tmp_path.iterdir()) == [
        "selection.pdf"
    ]


def test_pdf_export_preserves_existing_file(
    session: Mock, tmp_path: Path
) -> None:
    session.get_document.return_value.Slides.Count = 1
    target = tmp_path / "existing.pdf"
    target.write_bytes(b"user document")
    with pytest.raises(FileExistsError):
        export.export_pdf(export.ExportPDFInput(file_path=str(target)))
    assert target.read_bytes() == b"user document"
    session.get_document.return_value.ExportAsFixedFormat.assert_not_called()


def test_image_export_never_publishes_invalid_output(
    session: Mock, tmp_path: Path
) -> None:
    presentation = session.get_document.return_value
    presentation.Slides.Count = 1
    with pytest.raises(RuntimeError, match="did not create"):
        export.export_images(
            export.ExportImagesInput(output_dir=str(tmp_path))
        )
    assert list(tmp_path.iterdir()) == []


def test_image_export_preserves_concurrent_destination(
    monkeypatch: pytest.MonkeyPatch,
    session: Mock,
    tmp_path: Path,
) -> None:
    presentation = session.get_document.return_value
    presentation.Slides.Count = 1

    def render(path: str, image_format: str, *dimensions: int) -> None:
        with Image.new("RGB", (4, 4), "white") as image:
            image.save(path, format=image_format)

    presentation.Slides.return_value.Export.side_effect = render
    publish = export.publish_file

    def race(source: Path, target: Path) -> None:
        target.write_bytes(b"concurrent user data")
        publish(source, target)

    monkeypatch.setattr(export, "publish_file", race)
    with pytest.raises(RuntimeError, match="publication stopped"):
        export.export_images(
            export.ExportImagesInput(output_dir=str(tmp_path))
        )
    assert (tmp_path / "Slide1.png").read_bytes() == b"concurrent user data"


@pytest.mark.parametrize(
    "file_name", ["../slide.png", "C:slide.png", "slide.png:stream"]
)
def test_image_export_rejects_filename_paths(file_name: str) -> None:
    with pytest.raises(
        ValidationError, match="without directories or streams"
    ):
        export.ExportImagesInput(
            output_dir="images", slide_index=1, file_name=file_name
        )


def test_clipboard_temp_files_follow_runtime_lifespan(
    monkeypatch: pytest.MonkeyPatch,
    session: Mock,
) -> None:
    presentation = session.get_document.return_value
    presentation.Slides.Count = 2

    def render(path: str, image_format: str, *dimensions: int) -> None:
        with Image.new("RGB", (4, 4), "white") as image:
            image.save(path, format=image_format)

    presentation.Slides.return_value.Export.side_effect = render
    write_clipboard = Mock(return_value="text/uri-list")
    monkeypatch.setattr(export, "ensure_clipboard_available", Mock())
    monkeypatch.setattr(export, "copy_images", write_clipboard)
    with ExitStack() as resources:
        monkeypatch.setattr(
            export, "get_runtime", lambda: SimpleNamespace(resources=resources)
        )
        result = export.copy_to_clipboard(
            export.CopyToClipboardInput(slide_indices=[1, 2])
        )
        directory = Path(result["temp_dir"])
        assert len(list(directory.iterdir())) == 2
        assert write_clipboard.call_args.args == (
            [directory / "Slide1.png", directory / "Slide2.png"],
        )
        assert result["method"] == "text/uri-list"
    assert not directory.exists()


def test_slideshow_navigation_uses_selected_document(session: Mock) -> None:
    first = Mock(Presentation=SimpleNamespace(FullName="first.pptx"))
    target = Mock(Presentation=SimpleNamespace(FullName="target.pptx"))
    target.View.CurrentShowPosition, target.View.State = 2, 1
    session.get_document.return_value.FullName = "target.pptx"
    session.get_application.return_value.SlideShowWindows = Mock(
        Count=2,
        side_effect=[first, target],
    )
    result = slideshow.slideshow_next(DocumentInput(document="target.pptx"))
    first.View.Next.assert_not_called()
    target.View.Next.assert_called_once()
    assert result["current_slide"] == 2


def test_slideshow_start_restores_stored_settings_on_failure(
    session: Mock,
) -> None:
    presentation = session.get_document.return_value
    presentation.Slides.Count = 3
    settings = SimpleNamespace(
        ShowType=1,
        RangeType=1,
        StartingSlide=1,
        EndingSlide=3,
        LoopUntilStopped=0,
        Run=Mock(side_effect=RuntimeError("failed")),
    )
    presentation.SlideShowSettings = settings
    with pytest.raises(RuntimeError, match="failed"):
        slideshow.slideshow_start(
            slideshow.SlideShowStartInput(start_slide=2, loop=True)
        )
    assert (
        settings.ShowType,
        settings.RangeType,
        settings.StartingSlide,
        settings.EndingSlide,
        settings.LoopUntilStopped,
    ) == (1, 1, 1, 3, 0)


def test_history_command_restores_activation_when_native_action_fails(
    session: Mock,
) -> None:
    events: list[str] = []

    @contextmanager
    def window() -> Iterator[None]:
        events.append("activate")
        try:
            yield
        finally:
            events.append("restore")

    session.command_window.side_effect = window
    session.get_application.return_value.CommandBars.ExecuteMso.side_effect = (
        RuntimeError("failed")
    )
    with pytest.raises(RuntimeError, match="failed"):
        history.undo(history.HistoryInput())
    assert events == ["activate", "restore"]


def test_smartart_node_creation_failure_is_not_reported_as_success(
    session: Mock,
) -> None:
    layout = SimpleNamespace(Name="Process")
    session.get_application.return_value.SmartArtLayouts.return_value = layout
    shape = Mock()
    shape.SmartArt.AllNodes.Count = 1
    shape.SmartArt.AllNodes.return_value.AddNode.side_effect = RuntimeError(
        "unsupported node"
    )
    shapes = session.get_document.return_value.Slides.return_value.Shapes
    shapes.AddSmartArt.return_value = shape
    with pytest.raises(RuntimeError, match="unsupported node"):
        smartart.add_smartart(
            smartart.AddSmartArtInput(slide_index=1, node_texts=["A", "B"])
        )
    shape.SmartArt.AllNodes.return_value.AddNode.assert_called_once()


def test_smartart_invalid_node_does_not_append_elsewhere(
    monkeypatch: pytest.MonkeyPatch,
    session: Mock,
) -> None:
    shape = Mock(Type=smartart.MSO_SMART_ART)
    shape.SmartArt.AllNodes.Count = 1
    monkeypatch.setattr(
        smartart, "_get_shape", lambda slide, identifier: shape
    )
    with pytest.raises(ValueError, match="node_index"):
        smartart.modify_smartart(
            smartart.ModifySmartArtInput(
                slide_index=1,
                shape_name_or_index="diagram",
                action="add_node",
                node_index=2,
            )
        )
    shape.SmartArt.AllNodes.assert_not_called()


def test_interactive_animation_reports_exact_sequence(
    monkeypatch: pytest.MonkeyPatch,
    session: Mock,
) -> None:
    trigger = SimpleNamespace(Id=42, Name="button")
    shape = SimpleNamespace(Id=7, Name="target")
    monkeypatch.setattr(
        animation,
        "_get_shape",
        lambda slide, identifier: trigger if identifier == "button" else shape,
    )
    first = Mock(Count=1)
    first.return_value.Timing.TriggerShape.Id = 5
    target = Mock(Count=1)
    target.return_value.Timing.TriggerShape.Id = 42
    target.AddEffect.return_value.EffectType = 1
    target.AddEffect.return_value.Index = 2
    sequences = Mock(Count=2, side_effect=[first, target])
    timeline = session.get_document.return_value.Slides.return_value.TimeLine
    timeline.InteractiveSequences = sequences
    result = animation.add_animation(
        animation.AddAnimationInput(
            slide_index=1,
            shape_name_or_index="target",
            trigger="on_shape_click",
            trigger_shape="button",
        )
    )
    assert result["sequence_index"] == 2
    assert result["effect_index"] == 2
    assert result["trigger_shape_name"] == "button"
    first.AddEffect.assert_not_called()
    sequences.Add.assert_not_called()


def test_clipboard_close_failure_preserves_write_failure() -> None:
    native = Mock(spec=ModuleType)
    native.OpenClipboard = Mock()
    native.CloseClipboard = Mock(side_effect=RuntimeError("close"))
    failure = ValueError("write")
    with pytest.raises(ValueError) as caught:
        with clipboard._windows_clipboard(native):
            raise failure
    assert caught.value is failure


def test_image_export_honors_composed_call_document(
    session: Mock, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    wrong = Mock()
    correct = Mock()
    correct.Slides.Count = 1
    session.get_document.side_effect = lambda document=None: (
        correct if document == "chosen" else wrong
    )

    def render(
        slide: NativeObject,
        path: Path,
        image_format: str,
        width: int | None,
        height: int | None,
    ) -> None:
        path.write_bytes(b"verified image")

    monkeypatch.setattr(export, "_export_slide_image", render)
    result = export.export_images(
        export.ExportImagesInput(document="chosen", output_dir=str(tmp_path))
    )
    assert result["files_count"] == 1
    wrong.Slides.assert_not_called()
    correct.Slides.assert_called_once_with(1)
