"""Manage sessions and temporary resources for one service lifespan."""

from __future__ import annotations

import logging
from collections.abc import Callable
from contextlib import ExitStack
from contextvars import Token
from threading import Lock, Thread
from types import TracebackType

from src.errors import WpsError
from src.models import Component
from src.native.components import SPECS
from src.native.context import CURRENT_RUNTIME
from src.native.session import WpsSession

logger = logging.getLogger(__name__)


class Runtime:
    """Own the sessions and temporary resources of one service."""

    def __init__(
        self, session_factory: Callable[[Component], WpsSession] = WpsSession
    ) -> None:
        self.resources = ExitStack()
        self.sessions = {
            component: session_factory(component) for component in SPECS
        }
        self._lock = Lock()
        self._closed = False
        self._token: Token[Runtime | None] | None = None
        self.cleanup_thread: Thread | None = None

    def __enter__(self) -> Runtime:
        if self._closed or self._token is not None:
            raise WpsError("Runtime cannot be reused or entered twice")
        self._token = CURRENT_RUNTIME.set(self)
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        try:
            self.close()
        except Exception:
            if exc is None:
                raise
            logger.exception(
                "Runtime cleanup failed while handling an operation error"
            )
        finally:
            if self._token is not None:
                CURRENT_RUNTIME.reset(self._token)
                self._token = None

    def close(self) -> None:
        """Stop sessions, retaining resources until native calls finish."""
        with self._lock:
            if self._closed:
                return
            self._closed = True
        pending: list[WpsSession] = []
        errors: list[str] = []
        for session in self.sessions.values():
            try:
                if not session.close():
                    pending.append(session)
            except Exception as error:
                errors.append(str(error))
        if pending:
            logger.warning(
                "WPS calls are still running; temporary resources remain "
                "owned until completion"
            )

            def finish() -> None:
                try:
                    for session in pending:
                        try:
                            if session.worker.thread is not None:
                                session.worker.thread.join()
                            session.close()
                        except Exception:
                            logger.exception("Deferred WPS cleanup failed")
                finally:
                    try:
                        self.resources.close()
                    except Exception:
                        logger.exception(
                            "Deferred temporary resource cleanup failed"
                        )

            self.cleanup_thread = Thread(
                target=finish, name="WPS cleanup", daemon=True
            )
            self.cleanup_thread.start()
        else:
            self.resources.close()
        if errors:
            raise WpsError("; ".join(errors))


def get_runtime() -> Runtime:
    """Return the runtime attached to the current service context."""
    runtime = CURRENT_RUNTIME.get()
    if runtime is None:
        raise WpsError(
            "WPS runtime is unavailable outside its service lifespan"
        )
    return runtime


def get_session(component: Component) -> WpsSession:
    """Return the component session owned by the active runtime."""
    return get_runtime().sessions[component]
