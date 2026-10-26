"""Actionable failures at the WPS execution boundary."""

import sys
from importlib import import_module

from wps_mcp.models import JsonObject


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


_ERROR_TYPES: dict[str, type[Exception]] = {
    kind.__name__: kind
    for kind in (
        WpsError,
        WpsUnavailableError,
        DocumentNotFoundError,
        OperationTimeoutError,
        UnknownOutcomeError,
        ValueError,
        TypeError,
        NotImplementedError,
        OSError,
    )
}


def detach_error(error: BaseException) -> Exception:
    """Copy approved error data without retaining native frames or objects."""
    if isinstance(error, NativeError):
        code = error.hresult if type(error.hresult) is int else None
        return NativeError(str(error), code)
    kind = _ERROR_TYPES.get(type(error).__name__)
    if kind is not type(error):
        if isinstance(error, ValueError):
            kind = ValueError
        elif isinstance(error, TypeError):
            kind = TypeError
        else:
            kind = WpsError
    assert kind is not None
    return kind(str(error))


def encode_error(error: BaseException) -> JsonObject:
    """Serialize the same detached failure used by in-process workers."""
    detached = detach_error(error)
    return {
        "type": type(detached).__name__,
        "message": str(detached),
        "hresult": detached.hresult
        if isinstance(detached, NativeError)
        else None,
    }


def decode_error(payload: JsonObject) -> Exception:
    """Validate remote error data and rebuild only approved exception types."""
    if (
        set(payload) != {"type", "message", "hresult"}
        or not isinstance(payload["type"], str)
        or not isinstance(payload["message"], str)
        or (
            payload["hresult"] is not None
            and type(payload["hresult"]) is not int
        )
    ):
        raise WpsError("Malformed native error")
    if payload["type"] == "NativeError":
        return NativeError(payload["message"], payload["hresult"])
    if payload["hresult"] is not None:
        raise WpsError("Malformed native error")
    return _ERROR_TYPES.get(payload["type"], WpsError)(payload["message"])
