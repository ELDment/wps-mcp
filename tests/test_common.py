"""Check native text boundaries and completed-file publication."""

from pathlib import Path

import pytest

from src.common.files import publish_file
from src.common.text import normalize_text


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
