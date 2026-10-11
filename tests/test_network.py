"""Bounded network reads and temporary image ownership."""

from email.message import Message
from io import BytesIO
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, Mock
from urllib.error import HTTPError

import pytest
from pydantic import ValidationError

from src.presentation import pictures


def test_download_closes_response_and_uses_timeout(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    response = MagicMock(
        url="https://example.test/picture.png",
        headers={"Content-Type": "image/png"},
    )
    response.__enter__.return_value = response
    response.read.return_value = b"image"
    opener = Mock(return_value=response)
    monkeypatch.setattr(pictures.urllib.request, "urlopen", opener)
    assert pictures._download_resource(
        "https://example.test/picture.png", max_bytes=10
    ) == (
        b"image",
        "image/png",
    )
    assert opener.call_args.kwargs["timeout"] == pictures.HTTP_TIMEOUT_SECONDS
    response.read.assert_called_once_with(11)
    response.__exit__.assert_called_once()


@pytest.mark.parametrize("declared", [True, False])
def test_download_rejects_oversized_body_and_closes(
    monkeypatch: pytest.MonkeyPatch, declared: bool
) -> None:
    headers = {"Content-Length": "11"} if declared else {}
    response = MagicMock(url="https://example.test/image", headers=headers)
    response.__enter__.return_value = response
    response.read.return_value = b"x" * 11
    monkeypatch.setattr(
        pictures.urllib.request, "urlopen", Mock(return_value=response)
    )
    with pytest.raises(ValueError, match="byte limit"):
        pictures._download_resource("https://example.test/image", max_bytes=10)
    response.__exit__.assert_called_once()
    if declared:
        response.read.assert_not_called()


def test_download_closes_http_error_body(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    with BytesIO(b"not found") as body:
        error = HTTPError(
            "https://example.test/image", 404, "missing", Message(), body
        )
        monkeypatch.setattr(
            pictures.urllib.request, "urlopen", Mock(side_effect=error)
        )
        with pytest.raises(HTTPError):
            pictures._download_resource("https://example.test/image")
        assert body.closed


@pytest.mark.parametrize(
    "url",
    ["file:///C:/private.txt", "ftp://example.test/image", "relative.png"],
)
def test_only_http_urls_are_downloaded(
    monkeypatch: pytest.MonkeyPatch, url: str
) -> None:
    opener = Mock()
    monkeypatch.setattr(pictures.urllib.request, "urlopen", opener)
    with pytest.raises(ValueError, match="HTTP"):
        pictures._download_resource(url)
    opener.assert_not_called()


def test_picture_temp_file_is_removed_on_insertion_failure() -> None:
    paths: list[Path] = []

    def fail_insert(path: str, *args: object) -> None:
        paths.append(Path(path))
        assert paths[0].read_bytes() == b"image payload"
        raise RuntimeError("insert failed")

    slide = SimpleNamespace(
        Shapes=SimpleNamespace(AddPicture=Mock(side_effect=fail_insert))
    )
    with pytest.raises(RuntimeError, match="insert failed"):
        pictures._insert_picture(
            slide,
            b"image payload",
            ".png",
            0.0,
            0.0,
            None,
            None,
            False,
            "front",
        )
    assert paths and not paths[0].exists()
    assert not paths[0].parent.exists()


def test_picture_temp_file_is_removed_after_success() -> None:
    paths: list[Path] = []
    picture = SimpleNamespace(
        Name="Picture", ZOrderPosition=1, Width=100.0, Height=50.0
    )

    def insert(path: str, *args: object) -> SimpleNamespace:
        paths.append(Path(path))
        assert paths[0].is_file()
        return picture

    slide = SimpleNamespace(
        Shapes=SimpleNamespace(AddPicture=Mock(side_effect=insert))
    )
    result = pictures._insert_picture(
        slide, b"picture", ".png", 0.0, 0.0, 80.0, 40.0, False, "front"
    )
    assert result["width"] == 80.0
    assert result["height"] == 40.0
    assert not paths[0].exists()


def test_metadata_cached_only_after_valid_download(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(pictures, "_icon_cache", None)
    monkeypatch.setattr(pictures, "_icon_cache_time", 0.0)
    download = Mock(
        return_value=(
            b')]}\'\n{"icons":[{"name":"home"}]}',
            "application/json",
        )
    )
    monkeypatch.setattr(pictures, "_download_resource", download)
    assert pictures._fetch_icon_metadata() == [{"name": "home"}]
    assert pictures._fetch_icon_metadata() == [{"name": "home"}]
    download.assert_called_once()


def test_invalid_metadata_is_not_cached(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(pictures, "_icon_cache", None)
    download = Mock(
        return_value=(b'{"icons":[{"name":42}]}', "application/json")
    )
    monkeypatch.setattr(pictures, "_download_resource", download)
    with pytest.raises(ValueError, match="Invalid icon"):
        pictures._fetch_icon_metadata()
    assert pictures._icon_cache is None


def test_icon_search_filters_missing_names(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        pictures,
        "_fetch_icon_metadata",
        lambda: [
            {"name": "home", "tags": ["house"], "popularity": 2000},
            {"name": "home_work", "tags": ["house"], "popularity": 1000},
        ],
    )
    monkeypatch.setattr(
        pictures, "_icon_is_on_cdn", lambda name: name != "home"
    )
    result = pictures.search_icons(
        pictures.SearchIconsInput(query="home", max_results=1)
    )
    assert [icon["name"] for icon in result["icons"]] == ["home_work"]


def test_download_models_reject_ambiguous_fit_and_path_injection() -> None:
    with pytest.raises(ValidationError, match="both width and height"):
        pictures.AddPictureFromUrlInput(
            slide_index=1, url="https://example.test/image", fit=True
        )
    with pytest.raises(ValidationError):
        pictures.AddSvgIconInput(slide_index=1, icon_name="../private")
