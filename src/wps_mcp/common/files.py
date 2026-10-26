"""Validate native outputs and preserve existing bytes after failed writes."""

import logging
import os
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from shutil import copy2
from tempfile import NamedTemporaryFile
from zipfile import ZipFile

from wps_mcp.errors import UnknownOutcomeError
from wps_mcp.models import JsonObject

logger = logging.getLogger(__name__)


def verify_output(file_path: str | Path) -> JsonObject:
    """Require a non-empty output file and return its path and byte count."""
    path = Path(file_path)
    if not path.is_file() or path.stat().st_size == 0:
        raise RuntimeError("WPS did not produce a non-empty output file")
    return {"file_path": str(path), "bytes": path.stat().st_size}


def publish_file(
    source: Path, target: Path, *, overwrite: bool = False
) -> None:
    """Publish a validated file staged on the destination filesystem.

    An exclusive hard link fails if another writer creates the target.
    Replacement uses an atomic rename only when explicitly requested.
    """
    if overwrite:
        os.replace(source, target)
    else:
        os.link(source, target)


def verify_saved_file(path: Path) -> None:
    """Reject empty files and OOXML files with the wrong document payload."""
    verify_output(path)
    part = {
        ".docx": "word/document.xml",
        ".xlsx": "xl/workbook.xml",
        ".pptx": "ppt/presentation.xml",
    }.get(path.suffix.lower())
    if part is not None:
        with ZipFile(path) as archive:
            if part not in archive.namelist():
                raise RuntimeError(f"WPS output is missing {part}")


@contextmanager
def protect_existing_file(
    path: Path, *, overwrite: bool = True
) -> Iterator[None]:
    """Restore existing bytes after failure; never replay a native write.

    COM state can already have changed when a call fails. If restoring the file
    is blocked by WPS or the filesystem, retain a named recovery copy.
    An unknown native outcome leaves the destination untouched and retains
    the recovery copy until the user can confirm native work has finished.

    Default SaveAs reserves a new pathname exclusively before suppressing WPS
    alerts. The reservation is closed so WPS can replace it; this does not lock
    out another process that later modifies or replaces the reserved file.
    """
    backup: Path | None = None
    reservation: os.stat_result | None = None
    retain_backup = False
    failed = False
    try:
        if not overwrite:
            with path.open("xb") as stream:
                reservation = os.fstat(stream.fileno())
        elif path.is_file():
            with NamedTemporaryFile(
                prefix=f".{path.name}.recovery-", dir=path.parent, delete=False
            ) as stream:
                backup = Path(stream.name)
            copy2(path, backup)
        try:
            yield
        except UnknownOutcomeError as error:
            retain_backup = True
            recovery = (
                f"Original bytes are preserved at {backup}. "
                if backup is not None
                else ""
            )
            raise UnknownOutcomeError(
                "WPS save outcome is unknown; native work may still be "
                "running. The destination was not rolled back or removed: "
                f"{path}. {recovery}"
                "Wait for WPS to finish and inspect the document before "
                "restoring any recovery copy or retrying. "
                f"Detail: {error}"
            ) from error
        except Exception as error:
            if backup is not None:
                try:
                    os.replace(backup, path)
                except OSError:
                    retain_backup = True
                    raise RuntimeError(
                        "WPS save failed and the destination "
                        "could not be restored. "
                        f"Original bytes are preserved at {backup}. "
                        "Native state may have changed; "
                        "inspect before retrying."
                    ) from error
            if reservation is not None:
                try:
                    current = path.lstat()
                    if (
                        current.st_dev == reservation.st_dev
                        and current.st_ino == reservation.st_ino
                        and current.st_mtime_ns == reservation.st_mtime_ns
                        and current.st_ctime_ns == reservation.st_ctime_ns
                        and current.st_size == 0
                    ):
                        path.unlink()
                except FileNotFoundError:
                    pass
                except OSError:
                    logger.warning(
                        "Could not remove the unused save reservation",
                        exc_info=True,
                    )
            raise RuntimeError(
                "WPS save failed. Existing destination bytes were preserved; "
                "native state may have changed. "
                "Inspect the document before retrying."
            ) from error
    except BaseException:
        failed = True
        raise
    finally:
        if backup is not None and not retain_backup:
            try:
                backup.unlink(missing_ok=True)
            except OSError:
                if not failed:
                    raise
                logger.warning(
                    "Could not remove recovery copy %s", backup, exc_info=True
                )
