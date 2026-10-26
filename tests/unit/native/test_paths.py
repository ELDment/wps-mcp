"""Verify filesystem identity and desktop services across platforms."""

import base64
import json
import ntpath
import os
import posixpath
import struct
import subprocess
import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
from urllib.parse import unquote, urlsplit

import pytest
from PIL import Image

from wps_mcp.errors import DocumentNotFoundError
from wps_mcp.native import clipboard, discovery, files, selection
from wps_mcp.presentation import documents as document
from wps_mcp.presentation import exports as export


@pytest.mark.parametrize("explicit", [False, True])
def test_macos_temporary_files_follow_container_override(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, explicit: bool
) -> None:
    monkeypatch.setattr(files, "sys", SimpleNamespace(platform="darwin"))
    container = tmp_path / "container"
    monkeypatch.setenv("WPS_DRIVER_DIR", str(container))
    requested = tmp_path / "requested"
    requested.mkdir()
    with files.temporary_directory(
        prefix="native-", dir=requested if explicit else None
    ) as directory:
        path = Path(directory)
        assert path.parent == (requested if explicit else container / "tmp")
        assert path.is_dir()
    assert not path.exists()
    if explicit:
        assert not container.exists()


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
    monkeypatch.setattr(discovery, "sys", SimpleNamespace(platform="linux"))
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
    monkeypatch.setattr(discovery, "sys", SimpleNamespace(platform="linux"))
    monkeypatch.delenv("APPDATA", raising=False)
    monkeypatch.setenv("WPS_TEMPLATES_DIR", str(tmp_path))
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "unused"))
    result = document.list_templates(document.ListTemplatesInput())
    assert result["templates_dir"] == str(tmp_path.resolve())


def test_macos_templates_discover_bundled_locales_recursively(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setattr(discovery, "sys", SimpleNamespace(platform="darwin"))
    monkeypatch.delenv("WPS_TEMPLATES_DIR", raising=False)
    monkeypatch.setattr(discovery, "application_bundle", lambda: tmp_path)
    root = tmp_path / "Contents/Resources/office6/mui"
    expected = []
    for locale, name in (("zh_CN", "report.dpt"), ("en_US", "slides.POTX")):
        path = root / locale / "templates/wpp/general" / name
        path.parent.mkdir(parents=True)
        path.write_bytes(b"template")
        expected.append(str(path.resolve()))
    (root / "zh_CN/templates/ignored.pptx").write_bytes(b"presentation")
    (root / "en_US/outside.potx").write_bytes(b"not in templates")
    result = document.list_templates(document.ListTemplatesInput())
    assert result["templates_dir"] == str(root.resolve())
    assert result["count"] == 2
    assert [item["file_path"] for item in result["templates"]] == expected


@pytest.mark.parametrize("explicit", [False, True])
def test_macos_template_override_keeps_existing_scan_semantics(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, explicit: bool
) -> None:
    monkeypatch.setattr(discovery, "sys", SimpleNamespace(platform="darwin"))
    bundle = Mock(
        side_effect=AssertionError("override must not inspect bundle")
    )
    monkeypatch.setattr(discovery, "application_bundle", bundle)
    nested = tmp_path / "nested"
    nested.mkdir()
    (tmp_path / "top.potx").write_bytes(b"template")
    (nested / "hidden.dpt").write_bytes(b"template")
    if explicit:
        monkeypatch.setenv("WPS_TEMPLATES_DIR", str(nested))
        params = document.ListTemplatesInput(templates_dir=str(tmp_path))
    else:
        monkeypatch.setenv("WPS_TEMPLATES_DIR", str(tmp_path))
        params = document.ListTemplatesInput()
    result = document.list_templates(params)
    assert result["templates_dir"] == str(tmp_path.resolve())
    assert [item["file_name"] for item in result["templates"]] == ["top.potx"]
    bundle.assert_not_called()


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


@pytest.mark.parametrize("count", [1, 2])
def test_macos_clipboard_uses_native_formats_and_path_arguments(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, count: int
) -> None:
    monkeypatch.setattr(clipboard, "sys", SimpleNamespace(platform="darwin"))
    monkeypatch.setattr(clipboard.os, "access", Mock(return_value=True))
    run = Mock()
    monkeypatch.setattr(clipboard.subprocess, "run", run)
    paths = [tmp_path / f"Slide {index}; '中文'.png" for index in range(count)]
    for path in paths:
        path.write_bytes(b"PNG payload")
    method = clipboard.copy_images(paths)
    run.assert_called_once()
    arguments = run.call_args.args[0]
    assert arguments[:5] == [
        "/usr/bin/osascript",
        "-l",
        "JavaScript",
        "-e",
        clipboard.MACOS_SCRIPT,
    ]
    assert arguments[5:] == ["--", *(str(path.resolve()) for path in paths)]
    assert all(str(path) not in clipboard.MACOS_SCRIPT for path in paths)
    assert method == ("public.png" if count == 1 else "public.file-url")
    assert run.call_args.kwargs["input"] is None
    assert run.call_args.kwargs["check"] is True
    assert run.call_args.kwargs["timeout"] == 10
    assert "shell" not in run.call_args.kwargs


def test_missing_macos_clipboard_helper_fails_before_rendering(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(clipboard, "sys", SimpleNamespace(platform="darwin"))
    monkeypatch.setattr(clipboard.os, "access", Mock(return_value=False))
    native = Mock()
    monkeypatch.setattr(export, "get_session", native)
    with pytest.raises(NotImplementedError, match="osascript"):
        export.copy_to_clipboard(export.CopyToClipboardInput())
    native.assert_not_called()


@pytest.mark.skipif(
    sys.platform != "darwin" or os.environ.get("WPS_NATIVE_TESTS") != "1",
    reason="Set WPS_NATIVE_TESTS=1 for isolated macOS pasteboard validation",
)
@pytest.mark.parametrize("count", [1, 2])
def test_native_macos_clipboard_on_private_pasteboard(
    tmp_path: Path, count: int
) -> None:
    paths = [tmp_path / f"image {index}; '中文'.png" for index in range(count)]
    for path in paths:
        with Image.new("RGBA", (4, 4), (255, 0, 0, 128)) as image:
            image.save(path)
    script = clipboard.MACOS_SCRIPT.replace(
        "$.NSPasteboard.generalPasteboard",
        "$.NSPasteboard.pasteboardWithUniqueName",
    )
    # Inspect a private board without changing the general clipboard.
    script = (
        script.rsplit("}", 1)[0]
        + """
    try {
        var result = [];
        var saved = pasteboard.pasteboardItems;
        for (var i = 0; i < Number(saved.count); i++) {
            var item = saved.objectAtIndex(i);
            result.push(paths.length === 1
                ? ObjC.unwrap(item.dataForType($.NSPasteboardTypePNG)
                    .base64EncodedStringWithOptions(0))
                : ObjC.unwrap(item.stringForType($.NSPasteboardTypeFileURL)));
        }
        return JSON.stringify(result);
    } finally { pasteboard.releaseGlobally; }
}
"""
    )
    assert "generalPasteboard" not in script
    result = subprocess.run(
        [
            "/usr/bin/osascript",
            "-l",
            "JavaScript",
            "-e",
            script,
            "--",
            *(str(path.resolve()) for path in paths),
        ],
        check=True,
        capture_output=True,
        text=True,
        timeout=10,
    )
    values = json.loads(result.stdout)
    assert len(values) == count
    if count == 1:
        assert base64.b64decode(values[0]) == paths[0].read_bytes()
    else:
        assert [unquote(urlsplit(value).path) for value in values] == [
            str(path.resolve()) for path in paths
        ]


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


@pytest.mark.parametrize("platform", ["linux", "darwin"])
@pytest.mark.parametrize("failure", ["timeout", "exit"])
def test_clipboard_failure_does_not_retry(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    failure: str,
    platform: str,
) -> None:
    monkeypatch.setattr(clipboard, "sys", SimpleNamespace(platform=platform))
    monkeypatch.setattr(clipboard.os, "access", Mock(return_value=True))
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
