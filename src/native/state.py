"""Guard native state changes and data leaving an apartment thread."""

import logging
import math
from collections.abc import Iterator
from contextlib import contextmanager
from typing import Any, TypeVar

from src.errors import WpsError
from src.models import NativeObject

logger = logging.getLogger(__name__)
T = TypeVar("T")


def ensure_plain_result(value: T) -> T:
    """Reject native references and cycles before data leaves the STA."""
    ancestors: set[int] = set()

    def visit(item: object) -> None:
        if isinstance(item, float) and not math.isfinite(item):
            raise WpsError("WPS results must contain finite numbers")
        if item is None or type(item) in (str, int, float, bool, bytes):
            return
        if isinstance(item, (dict, list, tuple)):
            identity = id(item)
            if identity in ancestors:
                raise WpsError("WPS results must not contain cycles")
            ancestors.add(identity)
            try:
                if isinstance(item, dict):
                    if any(not isinstance(key, str) for key in item):
                        raise WpsError("WPS result keys must be strings")
                    for child in item.values():
                        visit(child)
                else:
                    for child in item:
                        visit(child)
            finally:
                ancestors.remove(identity)
            return
        raise WpsError(
            "Only plain data or image bytes may leave the WPS thread"
        )

    visit(value)
    return value


@contextmanager
def temporary_attributes(
    target: NativeObject, **changes: Any
) -> Iterator[NativeObject]:
    """Restore all changed properties, including after a setter failure."""
    originals: list[tuple[str, Any]] = []
    failed = False
    try:
        for name, value in changes.items():
            originals.append((name, getattr(target, name)))
            setattr(target, name, value)
        yield target
    except BaseException:
        failed = True
        raise
    finally:
        errors = []
        for name, value in reversed(originals):
            try:
                setattr(target, name, value)
            except Exception as error:
                errors.append(f"{name}: {error}")
        if errors:
            detail = "; ".join(errors)
            if failed:
                logger.error("Native state restoration failed: %s", detail)
            else:
                raise WpsError(f"Native state restoration failed: {detail}")
