"""Own WPS application proxies on a single apartment thread."""

from __future__ import annotations

import gc
import sys
import time
from collections.abc import Callable, Iterator
from contextlib import AbstractContextManager, ExitStack, contextmanager
from importlib import import_module
from threading import current_thread
from typing import Any, ParamSpec, TypeVar

from wps_mcp.errors import WpsError, WpsUnavailableError
from wps_mcp.models import Component, NativeObject
from wps_mcp.native.capabilities import validate_visibility
from wps_mcp.native.components import BUSY_CODES
from wps_mcp.native.session import LocalSession
from wps_mcp.native.state import ensure_plain_result
from wps_mcp.native.worker import StaWorker

P = ParamSpec("P")
T = TypeVar("T")


class ComSession(LocalSession):
    """Borrow one WPS application; own its proxies on one STA thread."""

    def __init__(
        self,
        component: Component,
        *,
        connector: Callable[[str, bool], NativeObject] | None = None,
        scope: Callable[[], AbstractContextManager[Any]] | None = None,
        queue_timeout: float = 60.0,
        call_timeout: float = 30.0,
    ) -> None:
        super().__init__(component)
        self._connector = connector
        self._scope = scope
        self.worker = StaWorker(
            self.spec.name,
            self._native_scope,
            queue_timeout=queue_timeout,
            call_timeout=call_timeout,
        )

    @property
    def started(self) -> bool:
        """Report whether the COM apartment thread has started."""
        return self.worker.thread is not None

    @contextmanager
    def _native_scope(self) -> Iterator[None]:
        with ExitStack() as resources:
            if self._scope is not None:
                resources.enter_context(self._scope())
            else:
                if sys.platform != "win32":
                    raise WpsUnavailableError(
                        "Native WPS automation requires Windows"
                    )
                import pythoncom

                pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
                resources.callback(pythoncom.CoUninitialize)
            try:
                yield
            finally:
                self.application = None
                gc.collect()

    def _assert_thread(self) -> None:
        if current_thread() is not self.worker.thread:
            raise WpsError(
                "Native objects may only be accessed on their WPS worker"
            )

    def execute(
        self, function: Callable[P, T], *args: P.args, **kwargs: P.kwargs
    ) -> T:
        """Run once on the STA and detach the result before returning it."""

        def operation() -> T:
            deadline = time.monotonic() + 10.0
            while self.application is not None:
                try:
                    _ = self.application.Name
                    break
                except Exception as error:
                    if getattr(error, "hresult", None) not in BUSY_CODES:
                        self.application = None
                        break
                    if time.monotonic() >= deadline:
                        raise WpsError(
                            f"{self.spec.name} is busy; close its dialog"
                        ) from error
                    time.sleep(0.1)
            return ensure_plain_result(function(*args, **kwargs))

        return self.worker.execute(operation)

    def connect(
        self, visible: bool | None = None, allow_launch: bool = True
    ) -> NativeObject:
        """Borrow or launch WPS on the STA, optionally setting visibility."""
        self._assert_thread()
        validate_visibility(self.component, visible)
        if self.application is None:
            if self._connector is not None:
                self.application = self._connector(
                    self.spec.prog_id, allow_launch
                )
            else:
                client = import_module("win32com.client")

                try:
                    self.application = client.GetActiveObject(
                        self.spec.prog_id
                    )
                except Exception as error:
                    if getattr(error, "hresult", None) in BUSY_CODES:
                        raise WpsUnavailableError(
                            f"{self.spec.name} is busy"
                        ) from error
                    if not allow_launch:
                        raise WpsUnavailableError(
                            f"{self.spec.name} is not running"
                        ) from error
                    try:
                        self.application = client.Dispatch(self.spec.prog_id)
                    except Exception as launch_error:
                        raise WpsUnavailableError(
                            f"Cannot start {self.spec.prog_id}"
                        ) from launch_error
        application = self.application
        if application is None:
            raise WpsUnavailableError(
                f"No application returned for {self.spec.prog_id}"
            )
        if visible is not None:
            application.Visible = visible
        return application

    def close(self, timeout: float = 2.0) -> bool:
        """Release proxies after pending native operations finish."""
        return self.worker.close(timeout)

    def wait(self) -> None:
        """Wait for the apartment worker to release all native objects."""
        if self.worker.thread is not None:
            self.worker.thread.join()
