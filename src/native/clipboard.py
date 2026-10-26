"""Publish rendered images through the current desktop clipboard."""

import logging
import os
import shutil
import struct
import subprocess
import sys
from collections.abc import Iterator, Sequence
from contextlib import contextmanager
from importlib import import_module
from io import BytesIO
from pathlib import Path
from types import ModuleType

from PIL import Image

logger = logging.getLogger(__name__)

MACOS_SCRIPT = """
ObjC.import('AppKit');
function run(paths) {
    var items;
    if (paths.length === 1) {
        var data = $.NSData.dataWithContentsOfFile($(paths[0]));
        if (data.isNil()) throw new Error('Cannot read clipboard image');
        var item = $.NSPasteboardItem.alloc.init;
        if (!item.setDataForType(data, $.NSPasteboardTypePNG))
            throw new Error('Cannot prepare clipboard image');
        items = [item];
    } else {
        items = paths.map(function(path) {
            return $.NSURL.fileURLWithPath($(path));
        });
    }
    var pasteboard = $.NSPasteboard.generalPasteboard;
    pasteboard.clearContents;
    if (!pasteboard.writeObjects($(items)))
        throw new Error('Cannot write clipboard');
}
"""


def _linux_command(mime_type: str) -> list[str]:
    """Choose the native clipboard transport for the active desktop."""
    if os.environ.get("WAYLAND_DISPLAY"):
        executable = shutil.which("wl-copy")
        if executable is not None:
            return [executable, "--type", mime_type]
        raise NotImplementedError(
            "Wayland clipboard requires wl-copy from wl-clipboard"
        )
    if os.environ.get("DISPLAY"):
        executable = shutil.which("xclip")
        if executable is not None:
            return [
                executable,
                "-selection",
                "clipboard",
                "-target",
                mime_type,
                "-in",
            ]
        raise NotImplementedError("X11 clipboard requires xclip")
    raise NotImplementedError(
        "Linux clipboard requires an active Wayland or X11 desktop"
    )


def ensure_clipboard_available() -> None:
    """Reject unavailable desktop clipboard support before rendering."""
    if sys.platform == "win32":
        import_module("win32clipboard")
    elif sys.platform == "linux":
        _linux_command("image/png")
    elif sys.platform == "darwin":
        if not os.access("/usr/bin/osascript", os.X_OK):
            raise NotImplementedError("macOS clipboard requires osascript")
    else:
        raise NotImplementedError(
            f"Clipboard export is unsupported on {sys.platform}"
        )


@contextmanager
def _windows_clipboard(clipboard: ModuleType) -> Iterator[None]:
    """Release Windows clipboard access without masking a write failure."""
    clipboard.OpenClipboard()
    failed = False
    try:
        yield
    except BaseException:
        failed = True
        raise
    finally:
        try:
            clipboard.CloseClipboard()
        except Exception:
            if not failed:
                raise
            logger.exception(
                "Could not close clipboard after an operation failed"
            )


def _copy_windows(paths: Sequence[Path]) -> str:
    """Write a bitmap or file drop list using the Windows clipboard."""
    clipboard = import_module("win32clipboard")
    if len(paths) == 1:
        with (
            Image.open(paths[0]) as source,
            source.convert("RGBA") as rgba,
            Image.new("RGB", rgba.size, "white") as canvas,
            rgba.getchannel("A") as alpha,
            BytesIO() as buffer,
        ):
            canvas.paste(rgba, mask=alpha)
            canvas.save(buffer, format="BMP")
            payload = buffer.getvalue()[14:]
        format_type, method = 8, "CF_DIB"
    else:
        payload = struct.pack("<IiiII", 20, 0, 0, 0, 1)
        payload += ("\0".join(str(path) for path in paths) + "\0\0").encode(
            "utf-16-le"
        )
        format_type, method = 15, "CF_HDROP"
    with _windows_clipboard(clipboard):
        clipboard.EmptyClipboard()
        clipboard.SetClipboardData(format_type, payload)
    return method


def copy_images(paths: Sequence[Path]) -> str:
    """Copy PNG data or file URIs; callers retain referenced files."""
    if not paths:
        raise ValueError("Clipboard export requires at least one image")
    ensure_clipboard_available()
    if sys.platform == "win32":
        return _copy_windows(paths)
    if sys.platform == "darwin":
        mime_type = "public.png" if len(paths) == 1 else "public.file-url"
        command = [
            "/usr/bin/osascript",
            "-l",
            "JavaScript",
            "-e",
            MACOS_SCRIPT,
            "--",
            *(str(path.resolve()) for path in paths),
        ]
        payload = None
    else:
        if len(paths) == 1:
            mime_type = "image/png"
            payload = paths[0].read_bytes()
        else:
            mime_type = "text/uri-list"
            payload = (
                "\r\n".join(path.resolve().as_uri() for path in paths) + "\r\n"
            ).encode("utf-8")
        command = _linux_command(mime_type)
    try:
        subprocess.run(
            command,
            input=payload,
            check=True,
            timeout=10,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except subprocess.TimeoutExpired as error:
        raise RuntimeError(
            "Clipboard transfer timed out; inspect the clipboard "
            "before repeating the operation"
        ) from error
    except subprocess.CalledProcessError as error:
        raise RuntimeError(
            f"Clipboard transfer failed with exit code {error.returncode}"
        ) from error
    return mime_type
