"""Rendering failures release every native and Python resource."""

import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from src.common.rendering import render_pdf_page


def test_render_failure_releases_all_resources(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    events: list[str] = []
    image = Mock()
    image.save.side_effect = RuntimeError("encoding failed")
    image.close.side_effect = lambda: events.append("image")
    bitmap = Mock()
    bitmap.to_pil.return_value = image
    bitmap.close.side_effect = lambda: events.append("bitmap")
    page = Mock()
    page.get_width.return_value = 100
    page.get_height.return_value = 100
    page.render.return_value = bitmap
    page.close.side_effect = lambda: events.append("page")

    class Pdf:
        def __enter__(self) -> "Pdf":
            return self

        def __exit__(self, *args: object) -> None:
            events.append("pdf")

        def __len__(self) -> int:
            return 1

        def __getitem__(self, index: int) -> Mock:
            return page

    monkeypatch.setitem(
        sys.modules,
        "pypdfium2",
        SimpleNamespace(PdfDocument=lambda path: Pdf()),
    )
    with pytest.raises(RuntimeError, match="encoding failed"):
        render_pdf_page(Path("unused.pdf"))
    assert events == ["image", "bitmap", "page", "pdf"]
