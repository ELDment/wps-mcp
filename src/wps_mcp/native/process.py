"""Run Linux operations in a persistent, serial native worker process."""

from __future__ import annotations

import base64
import binascii
import logging
import socket
import subprocess
import sys
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from typing import ParamSpec, TypeVar, cast

from wps_mcp.errors import WpsError, WpsUnavailableError, decode_error
from wps_mcp.models import Component, InputModel, JsonObject
from wps_mcp.native.components import SPECS
from wps_mcp.native.context import CALL_DOCUMENT, CALL_OPERATION
from wps_mcp.native.transport import receive_message, send_message
from wps_mcp.native.worker import StaWorker

P = ParamSpec("P")
T = TypeVar("T")
logger = logging.getLogger(__name__)


class ProcessSession:
    """Own a private RPC process without sharing native object references."""

    def __init__(self, component: Component) -> None:
        self.component: Component = component
        self.spec = SPECS[component]
        self.worker = StaWorker(self.spec.name, self._scope, call_timeout=60.0)
        self._process: subprocess.Popen[bytes] | None = None
        self._channel: socket.socket | None = None
        self._request_id = 0
        self._failed = False

    @property
    def started(self) -> bool:
        """Report whether an operation has started the process owner."""
        return self.worker.thread is not None

    @contextmanager
    def _scope(self) -> Iterator[None]:
        """Release the worker only after the last native operation finishes."""
        try:
            yield
        finally:
            channel, process = self._channel, self._process
            if channel is not None:
                try:
                    if not self._failed:
                        send_message(channel, {"kind": "close"})
                except (OSError, WpsError):
                    pass
                finally:
                    channel.close()
            if process is not None:
                try:
                    exit_code = process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    process.terminate()
                    try:
                        exit_code = process.wait(timeout=3)
                    except subprocess.TimeoutExpired:
                        process.kill()
                        exit_code = process.wait()
                if exit_code != 0:
                    logger.warning(
                        "WPS %s worker exited with status %s; inspect open "
                        "documents before reconnecting",
                        self.component,
                        exit_code,
                    )
            self._channel = None

    def _start(self) -> socket.socket:
        """Start once; a crashed native process is never replaced silently."""
        if self._failed:
            raise WpsUnavailableError(
                "The WPS worker failed; inspect documents and reconnect MCP"
            )
        if self._channel is not None:
            return self._channel
        if sys.platform != "linux" or sys.version_info[:2] != (3, 11):
            raise WpsUnavailableError(
                "Linux RPC requires CPython 3.11; start with "
                "uvx --python 3.11 wps-mcp"
            )
        parent, child = socket.socketpair()
        try:
            self._process = subprocess.Popen(
                [
                    sys.executable,
                    "-P",
                    "-m",
                    "wps_mcp.protocol.host",
                    str(child.fileno()),
                    self.component,
                ],
                pass_fds=(child.fileno(),),
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=None,
            )
        except BaseException:
            parent.close()
            raise
        finally:
            child.close()
        self._channel = parent
        return parent

    def _exchange(self, request: JsonObject) -> object:
        """Execute once and restore only plain data or a detached error."""
        channel = self._start()
        self._request_id += 1
        request["id"] = self._request_id
        try:
            send_message(channel, request)
            response = receive_message(channel)
        except (OSError, WpsError) as error:
            self._failed = True
            raise WpsError(str(error)) from None
        if (
            type(response.get("id")) is not int
            or response["id"] != self._request_id
        ):
            self._failed = True
            raise WpsError("Native response does not match its request")
        if "error" in response:
            failure = response["error"]
            if set(response) != {"id", "error"} or not isinstance(
                failure, dict
            ):
                self._failed = True
                raise WpsError("Malformed native error")
            try:
                restored_error = decode_error(failure)
            except WpsError:
                self._failed = True
                raise
            raise restored_error
        if set(response) == {"id", "result"}:
            return response["result"]
        if (
            set(response) == {"id", "result", "encoding"}
            and response["encoding"] == "base64"
            and isinstance(response["result"], str)
        ):
            try:
                return base64.b64decode(response["result"], validate=True)
            except (ValueError, binascii.Error):
                pass
        self._failed = True
        raise WpsError("Malformed native result")

    def execute(
        self, function: Callable[P, T], *args: P.args, **kwargs: P.kwargs
    ) -> T:
        """Dispatch a registered operation; never serialize Python code."""
        if kwargs or len(args) != 1 or not isinstance(args[0], InputModel):
            raise TypeError(
                "Linux sessions execute registered tools with one input model"
            )
        operation = CALL_OPERATION.get()
        if operation is None:
            raise TypeError("Linux sessions require a registered operation")
        request = {
            "kind": "call",
            "operation": operation,
            "params": args[0].model_dump(mode="json", exclude_unset=True),
            "document": CALL_DOCUMENT.get(),
        }
        return cast(T, self.worker.execute(self._exchange, request))

    def close(self, timeout: float = 2.0) -> bool:
        """Cancel queued calls while allowing started operations to finish."""
        return self.worker.close(timeout)

    def wait(self) -> None:
        """Wait until the process owner has released its channel."""
        if self.worker.thread is not None:
            self.worker.thread.join()
