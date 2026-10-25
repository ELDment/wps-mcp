"""Exercise private native transport without launching WPS."""

import base64
import json
import socket
import struct
import subprocess
import sys
from collections.abc import Iterator
from contextlib import contextmanager, nullcontext
from typing import cast
from unittest.mock import Mock

import pytest

from src.errors import NativeError, WpsError, WpsUnavailableError
from src.models import Component, InputModel, JsonObject
from src.native import process, transport
from src.native.context import CALL_DOCUMENT
from src.native.process import ProcessSession
from src.native.runtime import Runtime
from src.protocol import host
from src.protocol.definitions import ToolDefinition


class ProcessInput(InputModel):
    value: str


class AutoProcessInput(ProcessInput):
    component: Component


def echo_input(params: ProcessInput) -> JsonObject:
    return {"value": params.value, "document": CALL_DOCUMENT.get()}


def echo_bytes(params: ProcessInput) -> bytes:
    return params.value.encode("utf-8")


def echo_component(params: AutoProcessInput) -> JsonObject:
    return {"component": params.component}


def fail_native(params: ProcessInput) -> JsonObject:
    raise NativeError(params.value, -2147467259)


@contextmanager
def channels() -> Iterator[tuple[socket.socket, socket.socket]]:
    parent, child = socket.socketpair()
    parent.settimeout(2)
    child.settimeout(2)
    try:
        yield parent, child
    finally:
        parent.close()
        child.close()


def test_transport_preserves_consecutive_frame_boundaries() -> None:
    first = {"id": 1, "text": "中文😀", "values": [None, True, 3.25]}
    second = {"id": 2, "value": "next"}
    with channels() as (parent, child):
        transport.send_message(parent, first)
        transport.send_message(parent, second)
        assert transport.receive_message(child) == first
        assert transport.receive_message(child) == second


def test_transport_reads_fragmented_header_and_body() -> None:
    payload = json.dumps({"id": 1, "value": "fragmented"}).encode("utf-8")
    frame = struct.pack("!I", len(payload)) + payload
    channel = Mock(spec=socket.socket)
    channel.recv.side_effect = [
        frame[index : index + 1] for index in range(len(frame))
    ]
    assert transport.receive_message(channel) == {
        "id": 1,
        "value": "fragmented",
    }
    assert channel.recv.call_count == len(frame)


@pytest.mark.parametrize("size", [0, transport.MAX_FRAME_SIZE + 1])
def test_transport_rejects_length_before_reading_body(size: int) -> None:
    channel = Mock(spec=socket.socket)
    channel.recv.return_value = struct.pack("!I", size)
    with pytest.raises(WpsError, match="length"):
        transport.receive_message(channel)
    channel.recv.assert_called_once_with(4)


def test_transport_rejects_large_outgoing_frame_before_sending(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(transport, "MAX_FRAME_SIZE", 20)
    channel = Mock(spec=socket.socket)
    with pytest.raises(WpsError, match="exceeds"):
        transport.send_message(channel, {"value": "x" * 21})
    channel.sendall.assert_not_called()


@pytest.mark.parametrize("fragment", [b"\x00\x00", b"\x00\x00\x00\x05{}"])
def test_transport_reports_disconnect_before_frame_completion(
    fragment: bytes,
) -> None:
    with channels() as (parent, child):
        parent.sendall(fragment)
        parent.shutdown(socket.SHUT_WR)
        with pytest.raises(WpsError, match="may have completed"):
            transport.receive_message(child)


@pytest.mark.parametrize(
    "payload", [b"[]", b"{", b"\xff", b'{"value":NaN}', b'{"value":1e309}']
)
def test_transport_rejects_malformed_json(payload: bytes) -> None:
    with channels() as (parent, child):
        parent.sendall(struct.pack("!I", len(payload)) + payload)
        with pytest.raises(WpsError):
            transport.receive_message(child)


@pytest.fixture
def registered_runtime(monkeypatch: pytest.MonkeyPatch) -> Runtime:
    functions = (echo_input, echo_bytes, fail_native, echo_component)
    definitions = {
        function.__name__: ToolDefinition(
            function.__name__,
            function,
            "auto" if function is echo_component else "writer",
            {},
        )
        for function in functions
    }
    monkeypatch.setattr(host, "DEFINITIONS", definitions)
    session = Mock()
    session.execute.side_effect = lambda function, params: function(params)
    runtime = Mock(spec=Runtime)
    runtime.sessions = {"writer": session}
    return cast(Runtime, runtime)


def request(function: str = "echo_input", **params: object) -> JsonObject:
    return {
        "kind": "call",
        "id": 7,
        "module": __name__,
        "function": function,
        "params": {"value": "中文😀", **params},
        "document": "selected.docx",
    }


def test_dispatch_revalidates_and_restores_document_context(
    registered_runtime: Runtime,
) -> None:
    token = CALL_DOCUMENT.set("outer.docx")
    try:
        response = host.dispatch(request(), registered_runtime, "writer")
        assert response == {
            "id": 7,
            "result": {"value": "中文😀", "document": "selected.docx"},
        }
        assert CALL_DOCUMENT.get() == "outer.docx"
    finally:
        CALL_DOCUMENT.reset(token)


@pytest.mark.parametrize(
    "changes",
    [
        {"module": "os", "function": "system"},
        {"function": "unregistered_function"},
        {"id": True},
        {"id": 0},
        {"document": {"unexpected": "context"}},
        {"params": {"value": "text", "extra": True}},
        {"params": {"value": 42}},
    ],
)
def test_dispatch_rejects_unknown_code_and_invalid_inputs_before_execution(
    registered_runtime: Runtime, changes: JsonObject
) -> None:
    response = host.dispatch(request() | changes, registered_runtime, "writer")
    assert "error" in response
    cast(
        Mock, registered_runtime.sessions["writer"]
    ).execute.assert_not_called()


def test_dispatch_checks_auto_component_before_execution(
    registered_runtime: Runtime,
) -> None:
    response = host.dispatch(
        request("echo_component", component="et"), registered_runtime, "writer"
    )
    assert response["error"]["type"] == "ValueError"
    cast(
        Mock, registered_runtime.sessions["writer"]
    ).execute.assert_not_called()


def test_dispatch_encodes_binary_results(registered_runtime: Runtime) -> None:
    response = host.dispatch(
        request("echo_bytes"), registered_runtime, "writer"
    )
    assert response["encoding"] == "base64"
    assert (
        base64.b64decode(response["result"], validate=True)
        == "中文😀".encode()
    )


def test_dispatch_preserves_native_error_code_and_context(
    registered_runtime: Runtime,
) -> None:
    original = CALL_DOCUMENT.get()
    response = host.dispatch(
        request("fail_native"), registered_runtime, "writer"
    )
    assert response == {
        "id": 7,
        "error": {
            "type": "NativeError",
            "message": "中文😀",
            "hresult": -2147467259,
        },
    }
    assert CALL_DOCUMENT.get() == original


def test_host_serves_registered_call_then_stops_on_close(
    monkeypatch: pytest.MonkeyPatch, registered_runtime: Runtime
) -> None:
    monkeypatch.setattr(host, "load_definitions", Mock())
    monkeypatch.setattr(
        host, "Runtime", lambda factory: nullcontext(registered_runtime)
    )
    with channels() as (parent, child):
        transport.send_message(parent, request())
        transport.send_message(parent, {"kind": "close"})
        host.serve(child, "writer")
        assert transport.receive_message(parent) == {
            "id": 7,
            "result": {"value": "中文😀", "document": "selected.docx"},
        }


@pytest.mark.parametrize("binary", [False, True])
def test_process_exchange_round_trips_plain_and_binary_results(
    binary: bool,
) -> None:
    session = ProcessSession("writer")
    result = b"\x89PNG\x00\xff" if binary else {"saved": True}
    response: JsonObject = {"id": 1, "result": result}
    if isinstance(result, bytes):
        response.update(
            encoding="base64", result=base64.b64encode(result).decode()
        )
    with channels() as (parent, child):
        session._channel = parent
        transport.send_message(child, response)
        assert session._exchange({"kind": "call"}) == result
        assert transport.receive_message(child) == {"kind": "call", "id": 1}
    assert session.close()


def test_process_exchange_restores_native_error() -> None:
    session = ProcessSession("writer")
    with channels() as (parent, child):
        session._channel = parent
        transport.send_message(
            child,
            {
                "id": 1,
                "error": {
                    "type": "NativeError",
                    "message": "rejected",
                    "hresult": -1,
                },
            },
        )
        with pytest.raises(NativeError, match="rejected") as caught:
            session._exchange({"kind": "call"})
        assert caught.value.hresult == -1
        assert session._failed is False
    assert session.close()


@pytest.mark.parametrize(
    "response",
    [
        {"id": 2, "result": None},
        {"id": True, "result": None},
        {"id": 1},
        {"id": 1, "error": "invalid"},
        {"id": 1, "error": {"type": "ValueError", "message": 42}},
        {
            "id": 1,
            "error": {
                "type": "NativeError",
                "message": "bad",
                "hresult": True,
            },
        },
        {"id": 1, "result": None, "unexpected": True},
        {"id": 1, "encoding": "base64", "result": "!!!"},
    ],
)
def test_malformed_response_poisoning_prevents_reusing_worker(
    response: JsonObject,
) -> None:
    session = ProcessSession("writer")
    with channels() as (parent, child):
        session._channel = parent
        transport.send_message(child, response)
        with pytest.raises(WpsError):
            session._exchange({"kind": "call"})
        assert session._failed is True
        with pytest.raises(WpsUnavailableError, match="reconnect"):
            session._start()
    assert session.close()


def test_worker_disconnect_never_spawns_a_replacement(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    session = ProcessSession("writer")
    launch = Mock()
    monkeypatch.setattr(process.subprocess, "Popen", launch)
    with channels() as (parent, child):
        session._channel = parent
        child.shutdown(socket.SHUT_WR)
        with pytest.raises(WpsError, match="may have completed"):
            session._exchange({"kind": "call"})
        with pytest.raises(WpsUnavailableError, match="reconnect"):
            session._exchange({"kind": "call"})
    launch.assert_not_called()
    assert session.close()


def test_process_request_contains_only_registered_function_identity(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    session = ProcessSession("writer")
    execute = Mock(return_value={"done": True})
    monkeypatch.setattr(session.worker, "execute", execute)
    token = CALL_DOCUMENT.set("selected.docx")
    try:
        assert session.execute(echo_input, ProcessInput(value="text")) == {
            "done": True
        }
    finally:
        CALL_DOCUMENT.reset(token)
    assert execute.call_args.args[1] == {
        "kind": "call",
        "module": __name__,
        "function": "echo_input",
        "params": {"value": "text"},
        "document": "selected.docx",
    }
    with pytest.raises(TypeError, match="one input model"):
        session.execute(lambda: None)
    assert execute.call_count == 1
    assert session.close()


def test_process_scope_sends_close_before_waiting_for_exit() -> None:
    session = ProcessSession("writer")
    child_process = Mock()
    child_process.wait.return_value = 0
    session._process = child_process
    with channels() as (parent, child):
        session._channel = parent
        with session._scope():
            pass
        assert transport.receive_message(child) == {"kind": "close"}
    assert session._channel is None
    child_process.wait.assert_called_once_with(timeout=5)
    child_process.terminate.assert_not_called()


def test_process_reports_crash_during_native_cleanup(
    caplog: pytest.LogCaptureFixture,
) -> None:
    session = ProcessSession("writer")
    session._process = Mock()
    session._process.wait.return_value = -11
    with session._scope():
        pass
    assert "writer worker exited with status -11" in caplog.text


@pytest.mark.skipif(sys.platform != "linux", reason="POSIX inherited socket")
def test_host_socket_does_not_survive_native_child_exec(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    with channels() as (parent, child):
        descriptor = child.detach()
        monkeypatch.setattr(host.sys, "argv", ["host", str(descriptor), "et"])
        monkeypatch.setattr(host.os, "dup2", Mock())

        def spawn_native(channel: socket.socket, component: str) -> None:
            assert component == "et"
            result = subprocess.run(
                [
                    sys.executable,
                    "-c",
                    "import os, sys\n"
                    "try: os.fstat(int(sys.argv[1]))\n"
                    "except OSError: sys.exit(0)\n"
                    "sys.exit(1)",
                    str(channel.fileno()),
                ],
                close_fds=False,
                check=False,
                timeout=10,
            )
            assert result.returncode == 0

        host.os.set_inheritable(descriptor, True)
        monkeypatch.setattr(host, "serve", spawn_native)
        host.main()
        assert parent.recv(1) == b""
