"""Define execution ownership and shared native document selection."""

from __future__ import annotations

import logging
from abc import ABC, abstractmethod
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from typing import ParamSpec, Protocol, TypeVar

from wps_mcp.errors import WpsError
from wps_mcp.models import Component, JsonObject, NativeObject
from wps_mcp.native.components import SPECS, ComponentSpec
from wps_mcp.native.context import CALL_DOCUMENT
from wps_mcp.native.selection import document_key, select_document

logger = logging.getLogger(__name__)
P = ParamSpec("P")
T = TypeVar("T")


class ExecutionSession(Protocol):
    """Execute an operation without exposing native references to callers."""

    component: Component
    spec: ComponentSpec

    @property
    def started(self) -> bool:
        """Report whether this execution context has started."""
        ...

    def execute(
        self, function: Callable[P, T], *args: P.args, **kwargs: P.kwargs
    ) -> T:
        """Execute a single operation without automatic replay."""
        ...

    def close(self, timeout: float = 2.0) -> bool:
        """Stop accepting work and report completed cleanup."""
        ...

    def wait(self) -> None:
        """Wait for outstanding work before releasing shared resources."""
        ...


class LocalSession(ABC):
    """Resolve documents only inside the native object's owning context."""

    def __init__(self, component: Component) -> None:
        self.component: Component = component
        self.spec = SPECS[component]
        self.application: NativeObject | None = None
        self.target_name: str | None = None

    @property
    @abstractmethod
    def started(self) -> bool:
        """Report whether native execution has started."""

    @abstractmethod
    def connect(
        self, visible: bool | None = None, allow_launch: bool = True
    ) -> NativeObject:
        """Connect inside the backend's native execution context."""

    @abstractmethod
    def execute(
        self, function: Callable[P, T], *args: P.args, **kwargs: P.kwargs
    ) -> T:
        """Execute inside the backend's native execution context."""

    @abstractmethod
    def close(self, timeout: float = 2.0) -> bool:
        """Release native references without discarding user documents."""

    @abstractmethod
    def wait(self) -> None:
        """Wait until outstanding native operations finish."""

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
