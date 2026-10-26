"""Stage native resources in locations accessible to the desktop backend."""

import sys
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from shutil import rmtree
from tempfile import mkdtemp

from src.errors import UnknownOutcomeError
from src.native.discovery import driver_directory


@contextmanager
def temporary_directory(
    *, prefix: str, dir: str | Path | None = None
) -> Iterator[str]:
    """Stage native files, retaining them until unknown work is inspected.

    Retained directories have no finalizer: native work can outlive both the
    Python worker and the MCP process, so they must be removed manually.
    """
    if dir is None and sys.platform == "darwin":
        dir = driver_directory() / "tmp"
        dir.mkdir(parents=True, exist_ok=True)
    directory = mkdtemp(prefix=prefix, dir=dir)
    retained = False
    try:
        yield directory
    except UnknownOutcomeError as error:
        retained = True
        raise UnknownOutcomeError(
            f"{error} Native temporary files are preserved at {directory}. "
            "Wait for WPS to finish and inspect the result before "
            "removing this directory manually."
        ) from error
    finally:
        if not retained:
            rmtree(directory)
