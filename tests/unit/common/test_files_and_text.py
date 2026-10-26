"""Check native text boundaries and completed-file publication."""

from pathlib import Path
from unittest.mock import Mock

import pytest

from wps_mcp.common.files import protect_existing_file, publish_file
from wps_mcp.common.text import normalize_text


@pytest.mark.parametrize("line_break", ["\r", "\n", "\r\n"])
def test_text_normalization_preserves_unicode_and_whitespace(
    line_break: str,
) -> None:
    assert normalize_text(f"  中文😀{line_break} tail ") == "  中文😀\r tail "
    assert normalize_text("first\rsecond\r\n", newline="\n") == (
        "first\nsecond\n"
    )


def test_rich_text_run_boundaries_remain_independent() -> None:
    runs = ["first\r", "\nsecond"]
    assert "".join(normalize_text(run) for run in runs) == "first\r\rsecond"


def test_publication_preserves_concurrent_destination(tmp_path: Path) -> None:
    source, target = tmp_path / "complete.pdf", tmp_path / "target.pdf"
    source.write_bytes(b"validated output")
    target.write_bytes(b"concurrent document")
    with pytest.raises(FileExistsError):
        publish_file(source, target)
    assert target.read_bytes() == b"concurrent document"
    assert source.read_bytes() == b"validated output"


def test_explicit_publication_replaces_destination(tmp_path: Path) -> None:
    source, target = tmp_path / "complete.pdf", tmp_path / "target.pdf"
    source.write_bytes(b"validated output")
    target.write_bytes(b"previous document")
    publish_file(source, target, overwrite=True)
    assert target.read_bytes() == b"validated output"
    assert not source.exists()


@pytest.mark.parametrize("failed", [False, True])
def test_backup_cleanup_preserves_save_failure(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    caplog: pytest.LogCaptureFixture,
    failed: bool,
) -> None:
    target = tmp_path / "document.docx"
    target.write_bytes(b"original")
    cleanup = Mock(side_effect=PermissionError("backup busy"))
    monkeypatch.setattr(Path, "unlink", cleanup)
    expected = RuntimeError if failed else PermissionError
    with pytest.raises(
        expected, match="WPS save failed" if failed else "busy"
    ):
        with protect_existing_file(target):
            target.write_bytes(b"new content")
            if failed:
                raise ValueError("native save rejected")
    assert target.read_bytes() == (b"original" if failed else b"new content")
    cleanup.assert_called_once()
    if failed:
        assert "Could not remove recovery copy" in caplog.text
