"""Serial native execution with cancellation and deterministic ownership."""

from collections.abc import Callable
from concurrent.futures import CancelledError, Future, TimeoutError
from contextlib import AbstractContextManager, nullcontext
from contextvars import Context, ContextVar, copy_context
from dataclasses import dataclass, field
from math import isfinite
from queue import Empty, Queue
from threading import Event, Lock, Thread, current_thread
from typing import Any, ParamSpec, TypeVar, cast

import anyio

from src.errors import (
    DocumentNotFoundError,
    OperationTimeoutError,
    WpsError,
    WpsUnavailableError,
)

P = ParamSpec("P")
T = TypeVar("T")


class PendingCalls:
    """Cancel only work that has not entered the native host."""

    def __init__(self) -> None:
        self._lock = Lock()
        self._futures: list[Future[Any]] = []
        self._cancelled = False

    def add(self, future: Future[Any]) -> None:
        """Track queued work or cancel it if its request already ended."""
        with self._lock:
            if not self._cancelled:
                self._futures.append(future)
                return
        future.cancel()

    def cancel(self) -> None:
        """Cancel queued futures while letting in-flight native work finish."""
        with self._lock:
            self._cancelled = True
            futures = tuple(self._futures)
        for future in futures:
            future.cancel()


PENDING_CALLS: ContextVar[PendingCalls | None] = ContextVar(
    "wps_pending_calls", default=None
)


@dataclass(slots=True)
class WorkItem:
    """Carry one operation, its call context and completion signals."""

    function: Callable[..., Any]
    arguments: tuple[Any, ...]
    keywords: dict[str, Any]
    context: Context
    future: Future[Any] = field(default_factory=Future)
    started: Event = field(default_factory=Event)


class StaWorker:
    """Own one thread for its entire lifetime; never retry native writes."""

    def __init__(
        self,
        name: str,
        scope: Callable[[], AbstractContextManager[Any]] = nullcontext,
        *,
        queue_timeout: float = 60.0,
        call_timeout: float = 30.0,
        startup_timeout: float = 5.0,
    ) -> None:
        timeouts = (queue_timeout, call_timeout, startup_timeout)
        if any(not isfinite(timeout) or timeout <= 0 for timeout in timeouts):
            raise ValueError("Timeouts must be finite and positive")
        self.name = name
        self._scope = scope
        self._queue_timeout = queue_timeout
        self._call_timeout = call_timeout
        self._startup_timeout = startup_timeout
        self._queue: Queue[WorkItem | None] = Queue()
        self._lock = Lock()
        self._ready = Event()
        self._finished = Event()
        self._startup_error: str | None = None
        self._cleanup_error: str | None = None
        self._stopping = False
        self._unresolved: Future[Any] | None = None
        self.thread: Thread | None = None

    def _start(self) -> None:
        with self._lock:
            if self._stopping:
                raise WpsError(f"{self.name} worker is closed")
            if self.thread is None:
                self.thread = Thread(
                    target=self._run, name=self.name, daemon=True
                )
                self.thread.start()
        if not self._ready.wait(self._startup_timeout):
            raise OperationTimeoutError(
                f"{self.name} worker is still initializing"
            )
        if self._startup_error is not None:
            raise WpsError(
                f"{self.name} initialization failed: {self._startup_error}"
            )
        if self._finished.is_set():
            raise WpsError(f"{self.name} worker has terminated")

    def _run(self) -> None:
        try:
            with self._scope():
                self._ready.set()
                while (item := self._queue.get()) is not None:
                    if item.future.set_running_or_notify_cancel():
                        item.started.set()
                        try:
                            result = item.context.run(
                                item.function, *item.arguments, **item.keywords
                            )
                        except BaseException as error:
                            # Recreate errors here: tracebacks and causes can
                            # retain COM proxies owned by this STA thread.
                            if type(error) in (
                                WpsError,
                                WpsUnavailableError,
                                DocumentNotFoundError,
                                OperationTimeoutError,
                            ):
                                error_type = type(error)
                            elif isinstance(error, ValueError):
                                error_type = ValueError
                            elif isinstance(error, TypeError):
                                error_type = TypeError
                            else:
                                error_type = WpsError
                            item.future.set_exception(error_type(str(error)))
                        else:
                            item.future.set_result(result)
                            del result
                    item.started.set()
                    del item
        except BaseException as error:
            if self._ready.is_set():
                self._cleanup_error = str(error)
            else:
                self._startup_error = str(error)
        finally:
            self._ready.set()
            self._finished.set()
            # A failing resource scope must not leave callers waiting forever.
            with self._lock:
                self._stopping = True
                self._cancel_queued()

    def _cancel_queued(self) -> None:
        """Drain pending work while the caller holds the lifecycle lock."""
        while True:
            try:
                item = self._queue.get_nowait()
            except Empty:
                return
            if item is not None:
                item.future.cancel()
                item.started.set()

    def execute(
        self, function: Callable[P, T], *args: P.args, **kwargs: P.kwargs
    ) -> T:
        """Run once on the owned thread and wait within the timeout budget."""
        if current_thread() is self.thread:
            return function(*args, **kwargs)
        self._start()
        item = WorkItem(function, args, kwargs, copy_context())
        started = item.started

        def wake(future: Future[Any]) -> None:
            started.set()

        item.future.add_done_callback(wake)
        with self._lock:
            if self._stopping:
                raise WpsError(f"{self.name} worker is closed")
            if self._unresolved is not None and not self._unresolved.done():
                raise OperationTimeoutError(
                    f"{self.name} is completing a timed-out call. "
                    "No new work was queued; "
                    "inspect its result before repeating a write."
                )
            tracker = PENDING_CALLS.get()
            if tracker is not None:
                tracker.add(item.future)
            self._queue.put(item)
        if not item.started.wait(self._queue_timeout) and item.future.cancel():
            raise OperationTimeoutError(
                "Queue timeout: operation cancelled before execution"
            )
        try:
            return cast(T, item.future.result(timeout=self._call_timeout))
        except TimeoutError:
            if item.future.done():
                return cast(T, item.future.result())
            with self._lock:
                self._unresolved = item.future
            raise OperationTimeoutError(
                "WPS exceeded the call timeout and may still finish. "
                "Inspect the document before repeating a write."
            ) from None
        except CancelledError:
            raise WpsError("Operation cancelled before execution") from None

    def close(self, timeout: float = 2.0) -> bool:
        """Stop new work and report whether resource release completed."""
        with self._lock:
            if not self._stopping:
                self._stopping = True
                self._cancel_queued()
                self._queue.put(None)
            thread = self.thread
        if thread is None:
            self._finished.set()
        elif thread is not current_thread():
            thread.join(timeout)
        if self._cleanup_error:
            raise WpsError(
                f"{self.name} resource cleanup failed: {self._cleanup_error}"
            )
        return self._finished.is_set()


async def run_offloaded(function: Callable[..., T], *args: Any) -> T:
    """Propagate call context and cancel queued work on caller disconnect."""
    tracker = PendingCalls()
    token = PENDING_CALLS.set(tracker)
    try:
        return await anyio.to_thread.run_sync(
            function, *args, abandon_on_cancel=True
        )
    except anyio.get_cancelled_exc_class():
        tracker.cancel()
        raise
    finally:
        PENDING_CALLS.reset(token)
