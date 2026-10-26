"""Read bounded web assets and search the Material icon catalog."""

import json
import threading
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from contextlib import closing
from urllib.parse import urlsplit

from pydantic import Field

from wps_mcp.models import InputModel, JsonObject
from wps_mcp.presentation.constants import ICON_PACKAGE_BASE
from wps_mcp.protocol.definitions import tool

HTTP_TIMEOUT_SECONDS = 15


MAX_DOWNLOAD_BYTES = 20 * 1024 * 1024


MAX_METADATA_BYTES = 8 * 1024 * 1024


def download_resource(
    url: str, *, max_bytes: int = MAX_DOWNLOAD_BYTES, method: str = "GET"
) -> tuple[bytes, str]:
    """Download a bounded HTTP response and close it on every exit path."""
    parsed = urlsplit(url)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        raise ValueError("Only absolute HTTP or HTTPS URLs are supported")
    request = urllib.request.Request(url, method=method)
    try:
        with urllib.request.urlopen(
            request, timeout=HTTP_TIMEOUT_SECONDS
        ) as response:
            if urlsplit(response.url).scheme not in {"http", "https"}:
                raise ValueError("The response redirected to a non-HTTP URL")
            content_type = response.headers.get("Content-Type", "")
            if method == "HEAD":
                return (b"", content_type)
            length = response.headers.get("Content-Length")
            if length is not None and int(length) > max_bytes:
                raise ValueError(
                    f"Response exceeds the {max_bytes}-byte limit"
                )
            data = response.read(max_bytes + 1)
            if len(data) > max_bytes:
                raise ValueError(
                    f"Response exceeds the {max_bytes}-byte limit"
                )
            return (data, content_type)
    except urllib.error.HTTPError as error:
        with closing(error):
            raise


_icon_cache: list[JsonObject] | None = None


_icon_cache_time = 0.0


_ICON_CACHE_TTL = 86400


_icon_cache_lock = threading.Lock()


_ICON_METADATA_URL = "https://fonts.google.com/metadata/icons"


def _fetch_icon_metadata() -> list[JsonObject]:
    """Fetch public icon metadata only when a caller explicitly searches."""
    global _icon_cache, _icon_cache_time
    with _icon_cache_lock:
        now = time.monotonic()
        if (
            _icon_cache is not None
            and now - _icon_cache_time < _ICON_CACHE_TTL
        ):
            return _icon_cache
        payload, _ = download_resource(
            _ICON_METADATA_URL, max_bytes=MAX_METADATA_BYTES
        )
        raw = payload.decode("utf-8")
        if raw.startswith(")]}'"):
            raw = raw.partition("\n")[2]
        metadata = json.loads(raw)
        if not isinstance(metadata, dict) or not isinstance(
            metadata.get("icons"), list
        ):
            raise ValueError("Invalid icon metadata response")
        icons = metadata["icons"]
        if any(
            not isinstance(icon, dict) or not isinstance(icon.get("name"), str)
            for icon in icons
        ):
            raise ValueError("Invalid icon entry in metadata response")
        _icon_cache = icons
        _icon_cache_time = time.monotonic()
        return icons


_icon_on_cdn_cache: dict[str, bool] = {}


def _icon_is_on_cdn(name: str) -> bool:
    """Exclude known missing icons; leave transient availability unresolved."""
    with _icon_cache_lock:
        cached = _icon_on_cdn_cache.get(name)
    if cached is not None:
        return cached
    try:
        download_resource(
            f"{ICON_PACKAGE_BASE}/outlined/{name}.svg", method="HEAD"
        )
        available = True
    except urllib.error.HTTPError as error:
        if error.code != 404:
            return True
        available = False
    except (urllib.error.URLError, TimeoutError, OSError):
        return True
    with _icon_cache_lock:
        _icon_on_cdn_cache[name] = available
    return available


class SearchIconsInput(InputModel):
    """Input for searching Material Symbols icons by keyword."""

    query: str = Field(
        ...,
        description="Icon search keywords; multiple words narrow the search.",
        min_length=1,
    )
    max_results: int = Field(
        default=20,
        ge=1,
        le=100,
        description="Maximum number of results to return (default: 20)",
    )


@tool(
    "wps_ppt_pictures",
    component=None,
    annotations={
        "title": "Search Material Icons",
        "readOnlyHint": True,
        "idempotentHint": True,
        "openWorldHint": True,
    },
    operation="search_icons",
)
def search_icons(params: SearchIconsInput) -> JsonObject:
    """Search Google's Material Symbols icon library by keyword."""
    icons = _fetch_icon_metadata()
    query_lower = params.query.lower().strip()
    query_words = query_lower.split()
    results = []
    for icon in icons:
        name = icon.get("name", "")
        tags = [t.lower() for t in icon.get("tags", [])]
        categories = [c.lower() for c in icon.get("categories", [])]
        popularity = icon.get("popularity", 0)
        score = 0
        if name == query_lower:
            score += 100
        elif query_lower.replace(" ", "_") == name:
            score += 90
        elif query_lower in name:
            score += 50
        if len(query_words) > 1 and all(w in name for w in query_words):
            score += 40
        for word in query_words:
            if word in name:
                score += 30
            for tag in tags:
                if word == tag:
                    score += 20
                elif word in tag:
                    score += 10
            for cat in categories:
                if word in cat:
                    score += 5
        if score > 0:
            score += min(popularity / 1000, 10)
            results.append(
                {
                    "name": name,
                    "categories": icon.get("categories", []),
                    "tags": icon.get("tags", [])[:8],
                    "popularity": popularity,
                    "score": round(score, 2),
                }
            )
    results.sort(key=lambda x: x["score"], reverse=True)
    kept: list[JsonObject] = []
    with ThreadPoolExecutor(max_workers=8) as pool:
        for start in range(0, len(results), params.max_results + 4):
            batch = results[start : start + params.max_results + 4]
            for icon, available in zip(
                batch,
                pool.map(_icon_is_on_cdn, [item["name"] for item in batch]),
                strict=True,
            ):
                if available:
                    kept.append(icon)
            if len(kept) >= params.max_results:
                break
    kept = kept[: params.max_results]
    return {
        "success": True,
        "query": params.query,
        "count": len(kept),
        "icons": kept,
    }
