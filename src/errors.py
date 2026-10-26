"""Actionable failures at the WPS execution boundary."""

import sys
from importlib import import_module


class WpsError(RuntimeError):
    """An operation failed and may have partially changed its document."""


class WpsUnavailableError(WpsError):
    """A required native WPS component is unavailable."""


class DocumentNotFoundError(WpsError):
    """The requested document is missing or its identity is ambiguous."""


class OperationTimeoutError(WpsError):
    """A queued operation was cancelled or a native call is unresolved."""


class UnknownOutcomeError(WpsError):
    """A native request may still be running; rollback is not yet safe."""


class NativeError(WpsError):
    """A native API rejected a call, optionally with an HRESULT."""

    def __init__(self, message: str, hresult: int | None = None) -> None:
        self.hresult = hresult
        super().__init__(message)


def native_errors() -> tuple[type[Exception], ...]:
    """Recognize native errors without importing Windows modules on Linux."""
    if sys.platform == "win32":
        return NativeError, import_module("pywintypes").com_error
    return (NativeError,)
