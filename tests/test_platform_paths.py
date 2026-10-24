"""Verify filesystem identity and desktop services across platforms."""

import ntpath
import posixpath
import struct
import subprocess
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from PIL import Image

from src.errors import DocumentNotFoundError
from src.native import clipboard, selection
from src.presentation import document, export


def test_linux_document_selection_preserves_case(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(selection, "path", posixpath)
    lower = SimpleNamespace(Name="report.docx", FullName="/docs/report.docx")
    upper = SimpleNamespace(Name="Report.docx", FullName="/docs/Report.docx")
    documents = [lower, upper]
    assert selection.select_document(documents, "/docs/report.docx") is lower
    assert selection.select_document(documents, "/docs/Report.docx") is upper
    assert selection.select_document(documents, "Report") is upper
    with pytest.raises(DocumentNotFoundError):
        selection.select_document(documents, "/DOCS/Report.docx")


def test_linux_backslash_is_a_filename_character(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(selection, "path", posixpath)
    native = SimpleNamespace(Name=r"a\b.docx", FullName=r"/docs/a\b.docx")
    assert selection.select_document([native], r"a\b.docx") is native
    assert selection.document_key(r"/docs/a\b.docx") != (
        selection.document_key("/docs/a/b.docx")
    )
    assert selection.document_key("/docs/report.docx ") != (
        selection.document_key("/docs/report.docx")
    )


def test_windows_document_selection_remains_case_insensitive(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(selection, "path", ntpath)
    native = SimpleNamespace(
        Name="Report.docx", FullName=r"C:\Docs\Report.docx"
    )
    assert selection.select_document([native], "c:/docs/report.docx") is native
    assert selection.select_document([native], "REPORT") is native


def test_linux_templates_use_xdg_data_home(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setattr(document, "sys", SimpleNamespace(platform="linux"))
    monkeypatch.delenv("APPDATA", raising=False)
    monkeypatch.delenv("WPS_TEMPLATES_DIR", raising=False)
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path))
    templates = tmp_path / "Kingsoft" / "office6" / "templates"
    templates.mkdir(parents=True)
    (templates / "report.POTX").write_bytes(b"template")
    result = document.list_templates(document.ListTemplatesInput())
    assert result["templates_dir"] == str(templates.resolve())
    assert result["templates"][0]["file_name"] == "report.POTX"


def test_explicit_template_directory_precedes_xdg(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setattr(document, "sys", SimpleNamespace(platform="linux"))
    monkeypatch.delenv("APPDATA", raising=False)
    monkeypatch.setenv("WPS_TEMPLATES_DIR", str(tmp_path))
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "unused"))
    result = document.list_templates(document.ListTemplatesInput())
    assert result["templates_dir"] == str(tmp_path.resolve())


@pytest.mark.parametrize("count", [1, 2])
def test_windows_clipboard_preserves_bitmap_and_file_drop_formats(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, count: int
) -> None:
    native = Mock()
    monkeypatch.setattr(clipboard, "sys", SimpleNamespace(platform="win32"))
    monkeypatch.setattr(clipboard, "import_module", Mock(return_value=native))
    paths = [tmp_path / f"Slide{index}.png" for index in range(count)]
    for path in paths:
        with Image.new("RGBA", (4, 4), (255, 0, 0, 128)) as image:
            image.save(path)
    method = clipboard.copy_images(paths)
    native.OpenClipboard.assert_called_once_with()
    native.EmptyClipboard.assert_called_once_with()
    native.CloseClipboard.assert_called_once_with()
    format_type, payload = native.SetClipboardData.call_args.args
    if count == 1:
        assert method == "CF_DIB"
        assert format_type == 8
        assert struct.unpack("<I", payload[:4])[0] == 40
    else:
        assert method == "CF_HDROP"
        assert format_type == 15
        assert payload[20:].decode("utf-16-le") == (
            "\0".join(str(path) for path in paths) + "\0\0"
        )


@pytest.mark.parametrize("wayland", [True, False])
@pytest.mark.parametrize("count", [1, 2])
def test_linux_clipboard_sends_png_or_uri_list(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    wayland: bool,
    count: int,
) -> None:
    monkeypatch.setattr(clipboard, "sys", SimpleNamespace(platform="linux"))
    monkeypatch.setenv("DISPLAY", ":0")
    if wayland:
        monkeypatch.setenv("WAYLAND_DISPLAY", "wayland-0")
    else:
        monkeypatch.delenv("WAYLAND_DISPLAY", raising=False)
    monkeypatch.setattr(clipboard.shutil, "which", lambda name: "/bin/" + name)
    run = Mock()
    monkeypatch.setattr(clipboard.subprocess, "run", run)
    paths = [tmp_path / f"Slide {index}.png" for index in range(count)]
    for path in paths:
        path.write_bytes(b"PNG payload")
    method = clipboard.copy_images(paths)
    run.assert_called_once()
    arguments = run.call_args.args[0]
    assert arguments[0] == ("/bin/wl-copy" if wayland else "/bin/xclip")
    assert run.call_args.kwargs["check"] is True
    assert "shell" not in run.call_args.kwargs
    payload = run.call_args.kwargs["input"]
    if count == 1:
        assert method == "image/png"
        assert payload == b"PNG payload"
    else:
        assert method == "text/uri-list"
        assert payload.decode("utf-8").splitlines() == [
            path.resolve().as_uri() for path in paths
        ]
    assert method in arguments


def test_missing_clipboard_helper_fails_before_native_rendering(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(clipboard, "sys", SimpleNamespace(platform="linux"))
    monkeypatch.setenv("WAYLAND_DISPLAY", "wayland-0")
    monkeypatch.setattr(clipboard.shutil, "which", lambda name: None)
    native = Mock()
    monkeypatch.setattr(export, "get_session", native)
    with pytest.raises(NotImplementedError, match="wl-copy"):
        export.copy_to_clipboard(export.CopyToClipboardInput())
    native.assert_not_called()


@pytest.mark.parametrize("failure", ["timeout", "exit"])
def test_clipboard_failure_does_not_retry(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, failure: str
) -> None:
    monkeypatch.setattr(clipboard, "sys", SimpleNamespace(platform="linux"))
    monkeypatch.delenv("WAYLAND_DISPLAY", raising=False)
    monkeypatch.setenv("DISPLAY", ":0")
    monkeypatch.setattr(clipboard.shutil, "which", lambda name: "/bin/xclip")
    error = (
        subprocess.TimeoutExpired("xclip", 10)
        if failure == "timeout"
        else subprocess.CalledProcessError(1, "xclip")
    )
    run = Mock(side_effect=error)
    monkeypatch.setattr(clipboard.subprocess, "run", run)
    image = tmp_path / "image.png"
    image.write_bytes(b"PNG payload")
    with pytest.raises(RuntimeError, match="Clipboard transfer"):
        clipboard.copy_images([image])
    run.assert_called_once()
