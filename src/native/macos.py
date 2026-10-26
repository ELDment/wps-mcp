"""Call the built-in macOS WPS driver over its local Unix HTTP socket."""

from __future__ import annotations

import http.client
import json
import logging
import os
import re
import socket
import stat
import struct
import subprocess
import sys
import time
from collections.abc import Callable
from pathlib import Path
from shutil import copy2
from threading import RLock, current_thread
from typing import Any, ParamSpec, TypeVar
from uuid import uuid4

from src.common.paths import local_path
from src.errors import UnknownOutcomeError, WpsError, WpsUnavailableError
from src.models import Component, JsonObject
from src.native.components import SPECS
from src.native.discovery import (
    EXECUTABLES,
    SOCKET_PATTERN,
    application_bundle,
    discover_sockets,
    driver_directory,
)
from src.native.files import temporary_directory
from src.native.javascript import ObjectBridge
from src.native.session import LocalSession
from src.native.state import ensure_plain_result
from src.native.worker import StaWorker

P = ParamSpec("P")
T = TypeVar("T")
logger = logging.getLogger(__name__)

MAX_RESPONSE = 16 * 1024 * 1024


def unix_socket() -> socket.socket:
    """Create a local channel only on interpreters with Unix socket support."""
    family = getattr(socket, "AF_UNIX", None)
    if family is None:
        raise WpsUnavailableError("The WPS driver requires Unix sockets")
    return socket.socket(family, socket.SOCK_STREAM)


def socket_process(path: Path) -> int | None:
    """Ignore stale endpoints and identify the process behind a live socket."""
    with unix_socket() as channel:
        channel.settimeout(2)
        try:
            channel.connect(str(path))
        except (FileNotFoundError, ConnectionRefusedError):
            return None
        try:
            # macOS SOL_LOCAL / LOCAL_PEERPID identifies the actual listener.
            process = struct.unpack("i", channel.getsockopt(0, 2, 4))[0]
        except OSError as error:
            raise WpsUnavailableError(
                f"Cannot identify the WPS editor process for {path}"
            ) from error
        if process <= 0:
            raise WpsUnavailableError("WPS editor returned an invalid PID")
        return process


def live_sockets(directory: Path) -> dict[Path, int]:
    """Associate responsive editor endpoints with their owning processes."""
    return {
        path: pid
        for path in discover_sockets(directory)
        if (pid := socket_process(path)) is not None
    }


class UnixHttpConnection(http.client.HTTPConnection):
    """Use HTTP framing while connecting exclusively to a Unix socket."""

    def __init__(self, path: Path, timeout: float) -> None:
        super().__init__("localhost", timeout=timeout)
        self.path = path

    def connect(self) -> None:
        """Connect without resolving a hostname or opening a TCP port."""
        self.sock = unix_socket()
        self.sock.settimeout(self.timeout)
        self.sock.connect(str(self.path))


class MacDriver:
    """Serialize native calls and never replay a write after a lost reply."""

    def __init__(self, directory: Path | None = None) -> None:
        self.directory = (
            directory if directory is not None else driver_directory()
        )
        self._lock = RLock()
        self._sessions: dict[Path, tuple[int, int, str]] = {}
        self._process: int | None = None
        self._unresolved = False

    def _select(
        self, component: Component, socket_path: str | None
    ) -> tuple[Path, int]:
        if sys.platform != "darwin":
            raise WpsUnavailableError("This native driver requires macOS")
        if component not in EXECUTABLES:
            raise ValueError("component must be writer, et or wpp")
        socket_path = socket_path or os.environ.get("WPS_DRIVER_SOCKET")
        if socket_path is not None:
            selected = local_path(socket_path)
            if not SOCKET_PATTERN.fullmatch(selected.name):
                raise ValueError("socket_path must be a WPS editor socket")
            try:
                is_socket = stat.S_ISSOCK(selected.lstat().st_mode)
            except FileNotFoundError:
                is_socket = False
            if not is_socket or (pid := socket_process(selected)) is None:
                raise WpsUnavailableError(
                    f"WPS editor socket is not running: {selected}"
                )
            return selected, pid
        processes = live_sockets(self.directory)
        if self._process in processes.values():
            processes = {
                path: pid
                for path, pid in processes.items()
                if pid == self._process
            }
        if not processes:
            raise WpsUnavailableError(
                "No editor driver is running. Open a document in WPS "
                "Writer, ET or Presentation, then retry."
            )
        if len(set(processes.values())) != 1:
            raise WpsUnavailableError(
                "Multiple WPS processes are available; set WPS_DRIVER_SOCKET "
                "to the intended editor endpoint."
            )
        available = list(processes)
        prefix = f"wpsdriverlocal_{EXECUTABLES[component]}_"
        matching = [p for p in available if p.name.startswith(prefix)]
        choices = matching or available
        return choices[0], processes[choices[0]]

    def _request(
        self, path: Path, route: str, body: JsonObject, timeout: float
    ) -> JsonObject:
        connection = UnixHttpConnection(path, timeout)
        try:
            connection.request(
                "POST",
                route,
                json.dumps(body, ensure_ascii=True, allow_nan=False).encode(),
                {"Content-Type": "application/json"},
            )
            response = connection.getresponse()
            payload = response.read(MAX_RESPONSE + 1)
            if len(payload) > MAX_RESPONSE:
                raise ValueError("Native response exceeds 16 MiB")
            result = json.loads(payload)
            if not isinstance(result, dict):
                raise ValueError("Native response must be a JSON object")
        except (OSError, http.client.HTTPException, ValueError) as error:
            self._unresolved = True
            raise UnknownOutcomeError(
                "WPS reply could not be confirmed; the operation may have "
                "completed or may still be running. "
                "Inspect WPS before restarting the MCP client. "
                f"No retry was sent. Detail: {error}"
            ) from error
        finally:
            connection.close()
        if (
            response.status != 200
            or result.get("status") != 0
            or result.get("isSuccess") is not True
        ):
            raise WpsError(
                f"WPS driver HTTP {response.status}: {result.get('value')}. "
                "Earlier statements may have run; inspect before retrying."
            )
        return result

    def execute(
        self,
        component: Component,
        script: str,
        args: list[Any] | None = None,
        *,
        socket_path: str | None = None,
        timeout: float = 30.0,
    ) -> JsonObject:
        """Run a synchronous function body with explicit app and args bindings.

        Return plain JSON-compatible data. Scripts execute inside the native
        application and can use all objects exposed by WPS.
        """
        if not isinstance(script, str) or not script.strip():
            raise ValueError("script must be a non-empty function body")
        if not 0 < timeout <= 120:
            raise ValueError("timeout must be greater than 0 and at most 120")
        arguments = json.dumps(
            args if args is not None else [],
            ensure_ascii=True,
            allow_nan=False,
        )
        with self._lock:
            if self._unresolved:
                raise UnknownOutcomeError(
                    "A prior native call has an unknown outcome. Inspect "
                    "WPS and restart the MCP client before sending more work."
                )
            path, process = self._select(component, socket_path)
            self._process = process
            inode = path.stat().st_ino
            previous = self._sessions.get(path)
            if previous is None or previous[:2] != (inode, process):
                created = self._request(
                    path, "/session", {"capabilities": {}}, timeout
                )
                value = created.get("value")
                session_id = (
                    value.get("sessionId") if isinstance(value, dict) else None
                )
                if not isinstance(session_id, str) or not re.fullmatch(
                    r"[a-zA-Z0-9-]+", session_id
                ):
                    raise WpsError("WPS did not return a valid sessionId")
                self._sessions[path] = inode, process, session_id
            session_id = self._sessions[path][2]
            prog_id = json.dumps(SPECS[component].prog_id)
            wrapped = (
                "(function(){var app=CreateObject(" + prog_id + ");"
                "var value=(function(app,args){\n"
                + script
                + "\n})(app,"
                + arguments
                + ");"
                "if(value && typeof value.then==='function')"
                "{throw new Error('Async scripts are unsupported');}"
                "return JSON.stringify({result:"
                "value===undefined?null:value});})()"
            )
            reply = self._request(
                path,
                f"/session/{session_id}/execute/script/evaluate",
                {"script": wrapped},
                timeout,
            )
            try:
                decoded = json.loads(reply["value"])
                if not isinstance(decoded, dict) or set(decoded) != {"result"}:
                    raise ValueError("Unexpected evaluation envelope")
            except (KeyError, TypeError, ValueError) as error:
                self._unresolved = True
                raise UnknownOutcomeError(
                    "WPS executed the script but its return value was not "
                    "valid JSON. Return plain data; do not retry writes."
                ) from error
            return {
                "socket_path": str(path),
                "result": decoded["result"],
            }


def open_document(file_path: str) -> None:
    """Ask WPS to open an existing file and load the appropriate editor."""
    if sys.platform != "darwin":
        raise WpsUnavailableError("Opening macOS WPS requires macOS")
    path = local_path(file_path, exists=True)
    formats = {suffix for spec in SPECS.values() for suffix in spec.formats}
    if path.suffix.lower() not in formats:
        raise ValueError(f"Supported document formats: {sorted(formats)}")
    bundle = application_bundle()
    if bundle is None:
        raise WpsUnavailableError("WPS was not found in Applications")
    subprocess.run(
        ["/usr/bin/open", "-a", str(bundle), str(path)],
        check=True,
        capture_output=True,
        text=True,
        timeout=20,
    )


class MacExecutor:
    """Serialize all components sharing one WPS process and JS runtime."""

    def __init__(self) -> None:
        self.driver = MacDriver()
        self.worker = StaWorker("WPS macOS", call_timeout=120.0)


class MacSession(LocalSession):
    """Reuse domain operations through scoped JavaScript object references."""

    def __init__(
        self, component: Component, executor: MacExecutor | None = None
    ) -> None:
        super().__init__(component)
        self.executor = executor if executor is not None else MacExecutor()
        self._bridge: ObjectBridge | None = None
        self._scope_key = "wpsMcp" + uuid4().hex
        self._started = False

    @property
    def started(self) -> bool:
        """Report whether this component has connected during an operation."""
        return self._started

    def _bootstrap(self) -> None:
        """Load one editor from a private copy of a bundled WPS template."""
        if os.environ.get("WPS_DRIVER_SOCKET") is not None:
            self.executor.driver._select(self.component, None)
            return
        if live_sockets(self.executor.driver.directory):
            return
        bundle = application_bundle()
        if bundle is None:
            raise WpsUnavailableError("Install WPS Office for macOS")
        templates = bundle / "Contents/Resources/office6/mui"
        source = next(
            (
                path
                for suffix in (".docx", ".doc", ".wps")
                for path in sorted(
                    templates.glob(f"*/templates/newfile{suffix}")
                )
            ),
            None,
        )
        if source is None:
            raise WpsUnavailableError("Open a WPS document editor first")
        # Retain the bootstrap file until WPS has opened and closed its copy.
        with temporary_directory(prefix="wps-mcp-bootstrap-") as directory:
            target = Path(directory) / ("bootstrap" + source.suffix)
            copy2(source, target)
            open_document(str(target))
            deadline = time.monotonic() + 15
            while not live_sockets(self.executor.driver.directory):
                if time.monotonic() >= deadline:
                    raise WpsUnavailableError("WPS editor startup timed out")
                time.sleep(0.1)
            while True:
                result = self.executor.driver.execute(
                    "writer",
                    "var c=app.Documents;for(var i=c.Count;i>=1;i--)"
                    "{var d=c.Item(i);if(d.FullName===args[0])"
                    "{d.Close(0);return true;}}return false;",
                    [str(target)],
                )
                if result["result"] is True:
                    break
                if time.monotonic() >= deadline:
                    raise WpsUnavailableError(
                        "WPS did not open its startup document in time"
                    )
                time.sleep(0.1)

    def _activate(self) -> None:
        """Bind the component's engine before shared document selection."""
        self.executor.driver.execute(
            self.component,
            "if(Application.Name===app.Name)return null;"
            "var docs=app[args[0]];"
            "if(docs.Count && app.ActiveWindow)"
            "{app.ActiveWindow.Activate();return null;}"
            "for(var i=1;i<=docs.Count;i++){var d=docs.Item(i);"
            "if(d.Windows.Count){d.Windows.Item(1).Activate();break;}}"
            "return null;",
            [self.spec.collection],
        )

    def connect(
        self, visible: bool | None = None, allow_launch: bool = True
    ) -> Any:
        """Resolve an explicitly bound application inside this operation."""
        if current_thread() is not self.executor.worker.thread:
            raise WpsError(
                "Native objects may only be accessed on their WPS worker"
            )
        if self._bridge is None:
            raise WpsError("Native objects require an active WPS operation")
        if self.component == "wpp" and visible is False:
            raise ValueError("WPP cannot be hidden; use visible=true")
        if self.application is None:
            if allow_launch:
                self._bootstrap()
            self._activate()
            self.application = self._bridge.request("begin")
            self._started = True
        application = self.application
        if application is None:
            raise WpsUnavailableError(
                "WPS did not return an application object"
            )
        if visible is not None:
            application.Visible = visible
        return application

    def execute(
        self, function: Callable[P, T], *args: P.args, **kwargs: P.kwargs
    ) -> T:
        """Run a complete domain operation without replaying native writes."""

        def operation() -> T:
            bridge = ObjectBridge(
                self.executor.driver, self.component, self._scope_key
            )
            self._bridge = bridge
            failed = False
            try:
                return ensure_plain_result(function(*args, **kwargs))
            except BaseException:
                failed = True
                raise
            finally:
                try:
                    if (
                        self.application is not None
                        and not self.executor.driver._unresolved
                    ):
                        bridge.request("end")
                except Exception:
                    if not failed:
                        raise
                    logger.exception("Could not release WPS object scope")
                finally:
                    bridge.closed = True
                    self.application = None
                    self._bridge = None

        return self.executor.worker.execute(operation)

    def close(self, timeout: float = 2.0) -> bool:
        """Release the bridge worker without closing user documents."""
        return self.executor.worker.close(timeout)

    def wait(self) -> None:
        """Wait for unresolved native work before releasing temporary files."""
        thread = self.executor.worker.thread
        if thread is not None and thread is not current_thread():
            thread.join()
