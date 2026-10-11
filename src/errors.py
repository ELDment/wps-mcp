"""Actionable failures at the WPS execution boundary."""


class WpsError(RuntimeError):
    """An operation failed and may have partially changed its document."""


class WpsUnavailableError(WpsError):
    """A required native WPS component is unavailable."""


class DocumentNotFoundError(WpsError):
    """The requested document is missing or its identity is ambiguous."""


class OperationTimeoutError(WpsError):
    """A queued operation was cancelled or a native call is unresolved."""
