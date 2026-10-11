"""Keep service ownership and document targeting local to each call."""

from __future__ import annotations

from contextvars import ContextVar
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from src.native.runtime import Runtime


CALL_DOCUMENT: ContextVar[str | None] = ContextVar(
    "wps_document", default=None
)


CURRENT_RUNTIME: ContextVar[Runtime | None] = ContextVar(
    "wps_runtime", default=None
)
