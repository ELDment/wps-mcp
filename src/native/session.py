"""Own WPS application proxies on a single apartment thread."""

from __future__ import annotations

import gc
import logging
import sys
import time
from collections.abc import Callable, Iterator
from contextlib import AbstractContextManager, ExitStack, contextmanager
from threading import current_thread
from typing import Any, ParamSpec, TypeVar

from src.errors import WpsError, WpsUnavailableError
from src.models import Component, JsonObject, NativeObject
from src.native.components import BUSY_CODES, SPECS
from src.native.context import CALL_DOCUMENT
from src.native.selection import document_key, select_document
from src.native.state import ensure_plain_result
from src.native.worker import StaWorker

logger = logging.getLogger(__name__)
P = ParamSpec("P")
T = TypeVar("T")


class WpsSession:
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
        self.component = component
        self.spec = SPECS[component]
        self.application: NativeObject | None = None
        self.target_name: str | None = None
        self._connector = connector
        self._scope = scope
        self.worker = StaWorker(
            self.spec.name,
            self._native_scope,
            queue_timeout=queue_timeout,
            call_timeout=call_timeout,
        )

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
        if self.component == "wpp" and visible is False:
            raise ValueError("WPP cannot be hidden; use visible=true")
        if self.application is None:
            if self._connector is not None:
                self.application = self._connector(
                    self.spec.prog_id, allow_launch
                )
            else:
                import win32com.client

                try:
                    self.application = win32com.client.GetActiveObject(
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
                        self.application = win32com.client.Dispatch(
                            self.spec.prog_id
                        )
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

    def get_application(self, allow_launch: bool = False) -> NativeObject:
        """Get the native application, launching only when requested."""
        return self.connect(allow_launch=allow_launch)

    def get_documents(self) -> list[NativeObject]:
        """Read the open document collection on its owning thread."""
        collection = getattr(self.get_application(), self.spec.collection)
        return [
            collection.Item(index) for index in range(1, collection.Count + 1)
        ]

    def get_document(self, document: str | int | None = None) -> NativeObject:
        """Resolve the explicit, per-call or session document selector."""
        wanted = document if document is not None else CALL_DOCUMENT.get()
        return select_document(
            self.get_documents(),
            wanted if wanted is not None else self.target_name,
        )

    def set_target(self, document: str | int) -> JsonObject:
        """Persist a resolved document identity for subsequent calls."""
        selected = select_document(self.get_documents(), document)
        self.target_name = str(selected.FullName)
        return {"name": str(selected.Name), "full_name": self.target_name}

    def get_window(self) -> NativeObject | None:
        """Return the selected document's first window, if one exists."""
        document = self.get_document()
        return document.Windows.Item(1) if document.Windows.Count else None

    def activate_window(self) -> NativeObject:
        """Activate the selected document window or fail explicitly."""
        window = self.get_window()
        if window is None:
            raise WpsError("The selected document has no editor window")
        window.Activate()
        return window

    @contextmanager
    def command_window(self) -> Iterator[NativeObject]:
        """Activate the target temporarily and restore the prior window."""
        application = self.get_application()
        previous = application.ActiveWindow
        failed = False
        try:
            window = self.activate_window()
            if self.component == "wpp" and document_key(
                application.ActiveWindow.Presentation.FullName
            ) != document_key(self.get_document().FullName):
                raise WpsError("Could not activate the target presentation")
            yield window
        except BaseException:
            failed = True
            raise
        finally:
            if previous is not None:
                try:
                    previous.Activate()
                except Exception as error:
                    if failed:
                        logger.error(
                            "Could not restore the previous window: %s", error
                        )
                    else:
                        raise WpsError(
                            "Could not restore the previous window"
                        ) from error

    def close(self, timeout: float = 2.0) -> bool:
        """Release proxies on their owning thread after pending work ends."""
        return self.worker.close(timeout)
