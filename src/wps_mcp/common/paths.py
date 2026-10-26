"""Validate local inputs and output destinations before native access."""

from collections.abc import Mapping, Set
from pathlib import Path


def local_path(file_path: str, *, exists: bool = False) -> Path:
    """Require an absolute local path and preserve meaningful whitespace."""
    if (
        not file_path.strip()
        or "://" in file_path
        or file_path.startswith(("\\\\", "//"))
    ):
        raise ValueError("file_path must be an absolute local filesystem path")
    path = Path(file_path).expanduser()
    if not path.is_absolute():
        raise ValueError("file_path must be absolute")
    path = path.resolve()
    if str(path).startswith(("\\\\", "//")):
        raise ValueError(
            "Network paths are not supported by the local document service"
        )
    if exists and not path.is_file():
        raise ValueError(f"File does not exist: {path}")
    return path


def destination(
    file_path: str,
    extensions: Mapping[str, int | None] | Set[str],
    overwrite: bool = False,
) -> Path:
    """Validate the destination before creating its parent or calling WPS."""
    path = local_path(file_path)
    if path.suffix.lower() not in extensions:
        raise ValueError(
            f"Supported output extensions: {', '.join(extensions)}"
        )
    if path.exists() and (not path.is_file() or not overwrite):
        raise ValueError(
            f"Output already exists; set overwrite to replace it: {path}"
        )
    path.parent.mkdir(parents=True, exist_ok=True)
    return path
