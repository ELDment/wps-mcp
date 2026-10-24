"""Own Qt and WPS RPC objects on a Linux worker process's main thread."""

from __future__ import annotations

import gc
import logging
import os
import shutil
from collections.abc import Callable
from importlib import import_module
from pathlib import Path
from tempfile import mkdtemp
from threading import current_thread, main_thread
from typing import Any, ParamSpec, TypeVar

from src.errors import WpsError, WpsUnavailableError
from src.models import Component, NativeObject
from src.native.discovery import EXECUTABLES, linux_requirements
from src.native.rpc import check_hresult, wrap
from src.native.session import LocalSession
from src.native.state import ensure_plain_result

logger = logging.getLogger(__name__)
P = ParamSpec("P")
T = TypeVar("T")


def prepare_profile() -> Path:
    """Copy desktop settings without changing the user's WPS configuration."""
    source = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    state = (
        Path(
            os.environ.get("XDG_STATE_HOME", Path.home() / ".local" / "state")
        )
        / "wps-mcp"
    )
    state.mkdir(parents=True, exist_ok=True, mode=0o700)
    profile = Path(mkdtemp(prefix="rpc-", dir=state))
    target = profile / "Kingsoft"
    target.mkdir(mode=0o700)
    original = source / "Kingsoft" / "Office.conf"
    content = (
        original.read_text(encoding="utf-8") if original.is_file() else ""
    )
    keys = {
        "common\\do_not_detect_file_association_while_startup": "true",
        "wpsoffice\\Application%20Settings\\AppComponentMode": (
            "prome_independ"
        ),
        "wpsoffice\\Application%20Settings\\AppComponentModeInstall": (
            "prome_independ"
        ),
    }
    lines = content.splitlines()
    section = ""
    output: list[str] = []
    inserted = False
    for line in [*lines, "[wps-mcp-end]"]:
        if line.startswith("["):
            if section == "[6.0]":
                output.extend(f"{key}={value}" for key, value in keys.items())
                inserted = True
            section = line.strip()
        if line == "[wps-mcp-end]":
            break
        if section == "[6.0]" and line.partition("=")[0] in keys:
            continue
        output.append(line)
    if not inserted:
        output.extend(["[6.0]", *(f"{k}={v}" for k, v in keys.items())])
    config = target / "Office.conf"
    config.write_text(
        "\n".join(output) + "\n", encoding="utf-8", newline="\r\n"
    )
    config.chmod(0o600)
    os.environ["XDG_CONFIG_HOME"] = str(profile)
    return profile


class RpcSession(LocalSession):
    """Keep one component's RPC references inside the worker main thread."""

    def __init__(self, component: Component) -> None:
        super().__init__(component)
        self._qt: Any = None
        self._rpc: Any = None
        self._profile: Path | None = None
        self._closed = False
        self._launched = False
        self._previous_config = os.environ.get("XDG_CONFIG_HOME")

    @property
    def started(self) -> bool:
        """Report whether the native RPC client has been initialized."""
        return self._rpc is not None

    def _assert_thread(self) -> None:
        if current_thread() is not main_thread():
            raise WpsError("Linux WPS RPC must run on its process main thread")
        if self._closed:
            raise WpsError("Linux WPS session is closed")

    def connect(
        self, visible: bool | None = None, allow_launch: bool = True
    ) -> NativeObject:
        """Launch WPS with a private multi-component profile."""
        self._assert_thread()
        if self.component == "wpp" and visible is False:
            raise ValueError("WPP cannot be hidden; use visible=true")
        if self.application is None:
            if not allow_launch:
                raise WpsUnavailableError(
                    "No managed WPS instance; create, open or connect first"
                )
            directory, issues = linux_requirements(self.component)
            if issues or directory is None:
                raise WpsUnavailableError("; ".join(issues))
            if self._rpc is not None:
                raise WpsUnavailableError(
                    "RPC initialization failed; reconnect the MCP service"
                )
            self._profile = prepare_profile()
            try:
                common = import_module("pywpsrpc.common")
                executable = EXECUTABLES[self.component]
                module = import_module(f"pywpsrpc.rpc{executable}api")
                self._qt = common.QtApp(["wps-mcp"])
                factory = getattr(
                    module, f"create{executable.capitalize()}RpcInstance"
                )
                self._rpc = check_hresult(factory(), "Create RPC instance")
                check_hresult(
                    self._rpc.setStartTimeout(15_000_000), "Set RPC timeout"
                )
                check_hresult(
                    self._rpc.setProcessPath(str(directory / executable)),
                    "Set WPS executable",
                )
                self._launched = True
                application = check_hresult(
                    getattr(
                        self._rpc, f"get{executable.capitalize()}Application"
                    )(),
                    "Connect WPS RPC",
                )
                if application is None:
                    raise WpsUnavailableError("WPS returned no application")
                self.application = wrap(
                    application, getattr(module, executable + "api")
                )
            except (ImportError, OSError) as error:
                raise WpsUnavailableError(
                    f"Cannot load WPS RPC and Qt5 libraries: {error}"
                ) from None
        application = self.application
        if application is None:
            raise WpsUnavailableError("WPS returned no application")
        if visible is not None:
            application.Visible = visible
        return application

    def execute(
        self, function: Callable[P, T], *args: P.args, **kwargs: P.kwargs
    ) -> T:
        """Run one complete operation without retries or thread migration."""
        self._assert_thread()
        return ensure_plain_result(function(*args, **kwargs))

    def close(self, timeout: float = 2.0) -> bool:
        """Quit empty managed instances and preserve open user documents."""
        if self._closed:
            return True
        self._assert_thread()
        self._closed = True
        empty = not self._launched
        try:
            if self.application is not None:
                collection = getattr(self.application, self.spec.collection)
                empty = collection.Count == 0
                if empty:
                    self.application.Quit()
                del collection
        except Exception:
            empty = False
            logger.exception("Cleanup left the managed application running")
        finally:
            self.application = None
            self._rpc = None
            gc.collect()
            self._qt = None
            gc.collect()
            if empty and self._profile is not None:
                shutil.rmtree(self._profile)
            if self._profile is not None and os.environ.get(
                "XDG_CONFIG_HOME"
            ) == str(self._profile):
                if self._previous_config is None:
                    os.environ.pop("XDG_CONFIG_HOME", None)
                else:
                    os.environ["XDG_CONFIG_HOME"] = self._previous_config
        return True

    def wait(self) -> None:
        """Return immediately because this process executes synchronously."""
