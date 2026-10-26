"""Verify Unix HTTP framing, native binding and failure behavior."""

import json
import socket
import socketserver
import stat
import struct
import sys
import tempfile
from collections.abc import Iterator
from functools import partial
from http.server import BaseHTTPRequestHandler
from pathlib import Path
from threading import Thread
from types import SimpleNamespace
from typing import Any
from unittest.mock import MagicMock, Mock

import pytest

from src.errors import UnknownOutcomeError, WpsError, WpsUnavailableError
from src.models import Component
from src.native import macos
from src.native.macos import MacDriver, MacSession, discover_sockets


@pytest.fixture
def driver_host(
    monkeypatch: pytest.MonkeyPatch,
) -> Iterator[tuple[Path, list[dict[str, Any]]]]:
    if sys.platform == "win32":
        pytest.skip("Requires a Unix-domain HTTP server")
    monkeypatch.setattr(sys, "platform", "darwin")
    monkeypatch.delenv("WPS_DRIVER_SOCKET", raising=False)
    calls: list[dict[str, Any]] = []

    def process(path: Path) -> int | None:
        with macos.unix_socket() as channel:
            try:
                channel.connect(str(path))
            except (FileNotFoundError, ConnectionRefusedError):
                return None
        return 4242

    # Exercise actual Unix transport without depending on macOS socket options.
    monkeypatch.setattr(macos, "socket_process", process)

    class Handler(BaseHTTPRequestHandler):
        protocol_version = "HTTP/1.1"

        def log_message(self, format: str, *args: Any) -> None:
            pass

        def do_POST(self) -> None:
            body = json.loads(
                self.rfile.read(int(self.headers["Content-Length"]))
            )
            calls.append({"path": self.path, "body": body})
            if self.path == "/session":
                value: Any = {"sessionId": "test-session"}
            elif "LOSE_REPLY" in body["script"]:
                self.connection.shutdown(socket.SHUT_RDWR)
                self.close_connection = True
                return
            elif "NATIVE_ERROR" in body["script"]:
                payload = json.dumps(
                    {"isSuccess": False, "status": 13, "value": "native error"}
                ).encode()
                self.send_response(500)
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            elif "BAD_JSON" in body["script"]:
                value = "undefined"
            else:
                value = json.dumps(
                    {"result": "你好\nline 2"}, ensure_ascii=False
                )
            payload = json.dumps(
                {"isSuccess": True, "status": 0, "value": value}
            ).encode()
            self.send_response(200)
            self.send_header("Content-Length", str(len(payload)))
            # Keep alive to catch clients that incorrectly wait for EOF.
            self.send_header("Connection", "keep-alive")
            self.end_headers()
            self.wfile.write(payload)

    with tempfile.TemporaryDirectory(
        prefix="wps-test-", dir="/tmp"
    ) as directory:
        path = Path(directory) / "wpsdriverlocal_wps_1"
        host = socketserver.UnixStreamServer(str(path), Handler)
        thread = Thread(target=host.serve_forever, daemon=True)
        thread.start()
        try:
            yield Path(directory), calls
        finally:
            host.shutdown()
            host.server_close()
            thread.join(timeout=2)


def test_native_component_binding_and_http_keepalive(driver_host: Any) -> None:
    directory, calls = driver_host
    driver = MacDriver(directory)
    args = ['quotes " and newline\n汉字']
    first = driver.execute("et", "return args[0];", args)
    second = driver.execute("writer", "return app.Name;")
    assert first["result"] == "你好\nline 2"
    assert second["result"] == "你好\nline 2"
    assert [item["path"] for item in calls] == [
        "/session",
        "/session/test-session/execute/script/evaluate",
        "/session/test-session/execute/script/evaluate",
    ]
    assert 'CreateObject("KET.Application")' in calls[1]["body"]["script"]
    assert json.dumps(args, ensure_ascii=True) in calls[1]["body"]["script"]
    assert 'CreateObject("KWPS.Application")' in calls[2]["body"]["script"]


def test_lost_reply_is_not_replayed_and_blocks_later_calls(
    driver_host: Any,
) -> None:
    directory, calls = driver_host
    driver = MacDriver(directory)
    with pytest.raises(UnknownOutcomeError, match="No retry was sent"):
        driver.execute("writer", "LOSE_REPLY")
    with pytest.raises(UnknownOutcomeError, match="unknown outcome"):
        driver.execute("writer", "return app.Name;")
    assert len(calls) == 2


def test_unknown_result_also_blocks_other_component_sockets(
    driver_host: Any,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    directory, calls = driver_host
    driver = MacDriver(directory)
    other = macos.unix_socket()
    try:
        other.bind(str(directory / "wpsdriverlocal_et_1"))
        with pytest.raises(WpsError, match="No retry was sent"):
            driver.execute("writer", "LOSE_REPLY")
        select = Mock(side_effect=AssertionError("must not reconnect"))
        monkeypatch.setattr(driver, "_select", select)
        with pytest.raises(WpsError, match="unknown outcome"):
            driver.execute("et", "return app.Name;")
        select.assert_not_called()
        assert len(calls) == 2
    finally:
        other.close()


@pytest.mark.parametrize("script", ["NATIVE_ERROR", "BAD_JSON"])
def test_native_errors_never_report_success(
    driver_host: Any, script: str
) -> None:
    directory, calls = driver_host
    with pytest.raises(WpsError):
        MacDriver(directory).execute("writer", script)
    assert len(calls) == 2


def test_stale_and_non_editor_sockets_do_not_block_discovery(
    driver_host: Any,
) -> None:
    directory, _ = driver_host
    hub = macos.unix_socket()
    second = macos.unix_socket()
    try:
        hub.bind(str(directory / "wpsdriverlocal_wpsoffice_1"))
        second.bind(str(directory / "wpsdriverlocal_wps_2"))
        (directory / "wpsdriverlocal_et_1").write_text("not a socket")
        assert len(discover_sockets(directory)) == 2
        result = MacDriver(directory).execute("writer", "return app.Name;")
        assert result["socket_path"] == str(directory / "wpsdriverlocal_wps_1")
        assert (directory / "wpsdriverlocal_wps_2").exists()
    finally:
        hub.close()
        second.close()


@pytest.mark.parametrize("process", [4242, 4243])
def test_process_identity_controls_ambiguous_endpoint_selection(
    driver_host: Any, monkeypatch: pytest.MonkeyPatch, process: int
) -> None:
    directory, calls = driver_host
    first = directory / "wpsdriverlocal_wps_1"
    second = directory / "wpsdriverlocal_wps_2"
    monkeypatch.setattr(
        macos,
        "live_sockets",
        Mock(return_value={first: 4242, second: process}),
    )
    driver = MacDriver(directory)
    if process == 4242:
        assert driver._select("writer", None) == (first, 4242)
    else:
        with pytest.raises(WpsUnavailableError, match="Multiple"):
            driver.execute("writer", "return app.Name;")
    assert calls == []


@pytest.fixture
def selection_driver(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> MacDriver:
    monkeypatch.setattr(macos, "sys", SimpleNamespace(platform="darwin"))
    monkeypatch.delenv("WPS_DRIVER_SOCKET", raising=False)
    driver = MacDriver(tmp_path)
    monkeypatch.setattr(
        driver,
        "_request",
        Mock(
            side_effect=[
                {"value": {"sessionId": "fake-session"}},
                {"value": '{"result":null}'},
            ]
        ),
    )
    return driver


@pytest.mark.parametrize("previous_process", [None, 9999])
@pytest.mark.parametrize("component", ["et", "wpp"])
def test_unbound_process_rejects_mixed_component_endpoints(
    selection_driver: MacDriver,
    monkeypatch: pytest.MonkeyPatch,
    previous_process: int | None,
    component: Component,
) -> None:
    driver = selection_driver
    driver._process = previous_process
    main = driver.directory / "wpsdriverlocal_wpp_1"
    child = driver.directory / "wpsdriverlocal_et_2"
    main.touch()
    child.touch()
    monkeypatch.setattr(
        macos, "live_sockets", Mock(return_value={main: 101, child: 202})
    )
    request = driver._request
    assert isinstance(request, Mock)

    with pytest.raises(WpsUnavailableError, match="WPS_DRIVER_SOCKET"):
        driver.execute(component, "return null;")

    assert driver._process == previous_process
    request.assert_not_called()


@pytest.mark.parametrize(
    ("component", "endpoint"),
    [("writer", "wpp"), ("wpp", "wpp"), ("et", "et")],
)
def test_one_process_prefers_component_endpoint_with_fallback(
    selection_driver: MacDriver,
    monkeypatch: pytest.MonkeyPatch,
    component: Component,
    endpoint: str,
) -> None:
    driver = selection_driver
    presentation = driver.directory / "wpsdriverlocal_wpp_1"
    spreadsheet = driver.directory / "wpsdriverlocal_et_1"
    monkeypatch.setattr(
        macos,
        "live_sockets",
        Mock(return_value={presentation: 101, spreadsheet: 101}),
    )

    assert driver._select(component, None) == (
        driver.directory / f"wpsdriverlocal_{endpoint}_1",
        101,
    )


def test_bound_process_falls_back_to_main_editor_when_et_is_a_child(
    selection_driver: MacDriver, monkeypatch: pytest.MonkeyPatch
) -> None:
    driver = selection_driver
    main = driver.directory / "wpsdriverlocal_wpp_1"
    child = driver.directory / "wpsdriverlocal_et_2"
    main.touch()
    monkeypatch.setattr(macos, "live_sockets", Mock(return_value={main: 101}))
    driver.execute("wpp", "return null;")
    assert driver._process == 101
    monkeypatch.setattr(
        macos, "live_sockets", Mock(return_value={main: 101, child: 202})
    )
    assert driver._select("et", None) == (main, 101)


def test_bound_process_wins_between_matching_component_endpoints(
    selection_driver: MacDriver, monkeypatch: pytest.MonkeyPatch
) -> None:
    driver = selection_driver
    driver._process = 101
    main = driver.directory / "wpsdriverlocal_et_1"
    child = driver.directory / "wpsdriverlocal_et_2"
    monkeypatch.setattr(
        macos, "live_sockets", Mock(return_value={child: 202, main: 101})
    )
    assert driver._select("et", None) == (main, 101)


def test_departed_process_rebinds_only_when_replacement_is_unambiguous(
    selection_driver: MacDriver, monkeypatch: pytest.MonkeyPatch
) -> None:
    driver = selection_driver
    driver._process = 101
    replacement = driver.directory / "wpsdriverlocal_et_2"
    other = driver.directory / "wpsdriverlocal_et_3"
    monkeypatch.setattr(
        macos,
        "live_sockets",
        Mock(return_value={replacement: 202, other: 303}),
    )
    with pytest.raises(WpsUnavailableError, match="Multiple"):
        driver.execute("et", "return null;")
    assert driver._process == 101
    replacement.touch()
    monkeypatch.setattr(
        macos, "live_sockets", Mock(return_value={replacement: 202})
    )
    driver.execute("et", "return null;")
    assert driver._process == 202


@pytest.mark.parametrize("environment", [False, True])
def test_explicit_endpoint_uses_its_live_process(
    driver_host: Any, monkeypatch: pytest.MonkeyPatch, environment: bool
) -> None:
    directory, _ = driver_host
    path = directory / "wpsdriverlocal_wps_1"
    socket_path: str | None = str(path)
    if environment:
        monkeypatch.setenv("WPS_DRIVER_SOCKET", socket_path)
        socket_path = None
    driver = MacDriver(directory)
    driver._process = 9999
    assert driver._select("et", socket_path) == (
        path.resolve(),
        4242,
    )


def test_explicit_stale_socket_fails_before_sending(
    driver_host: Any,
) -> None:
    directory, calls = driver_host
    path = directory / "wpsdriverlocal_wpp_2"
    with macos.unix_socket() as channel:
        channel.bind(str(path))
    with pytest.raises(WpsUnavailableError, match="not running"):
        MacDriver(directory).execute("wpp", "return 1;", socket_path=str(path))
    assert path.exists()
    assert calls == []


def test_changed_process_recreates_native_session(
    driver_host: Any, monkeypatch: pytest.MonkeyPatch
) -> None:
    directory, calls = driver_host
    path = directory / "wpsdriverlocal_wps_1"
    select = Mock(side_effect=[(path, 4242), (path, 4243)])
    driver = MacDriver(directory)
    monkeypatch.setattr(driver, "_select", select)
    driver.execute("writer", "return 1;")
    driver.execute("writer", "return 2;")
    assert [call["path"] for call in calls].count("/session") == 2


def test_malformed_script_result_blocks_future_writes(
    driver_host: Any,
) -> None:
    directory, calls = driver_host
    driver = MacDriver(directory)
    with pytest.raises(UnknownOutcomeError, match="not valid JSON"):
        driver.execute("writer", "BAD_JSON")
    with pytest.raises(UnknownOutcomeError, match="unknown outcome"):
        driver.execute("et", "return 1;")
    assert len(calls) == 2


@pytest.mark.parametrize(
    "error", [None, FileNotFoundError(), ConnectionRefusedError()]
)
def test_socket_process_uses_peer_pid_and_ignores_stale_paths(
    monkeypatch: pytest.MonkeyPatch, error: Exception | None
) -> None:
    channel = MagicMock(spec=socket.socket)
    channel.__enter__.return_value = channel
    channel.connect.side_effect = error
    channel.getsockopt.return_value = struct.pack("i", 4242)
    factory = Mock(return_value=channel)
    monkeypatch.setattr(socket, "socket", factory)
    monkeypatch.setattr(socket, "AF_UNIX", 1, raising=False)
    actual = macos.socket_process(Path("/tmp/wpsdriverlocal_wps_1"))
    if error is None:
        assert actual == 4242
        channel.getsockopt.assert_called_once_with(0, 2, 4)
    else:
        assert actual is None
        channel.getsockopt.assert_not_called()
    channel.__exit__.assert_called_once()


@pytest.mark.parametrize("process", [None, 0, -1])
def test_unavailable_peer_identity_is_not_treated_as_a_known_process(
    monkeypatch: pytest.MonkeyPatch, process: int | None
) -> None:
    channel = MagicMock(spec=socket.socket)
    channel.__enter__.return_value = channel
    if process is None:
        channel.getsockopt.side_effect = OSError("unsupported option")
    else:
        channel.getsockopt.return_value = struct.pack("i", process)
    monkeypatch.setattr(socket, "socket", Mock(return_value=channel))
    monkeypatch.setattr(socket, "AF_UNIX", 1, raising=False)
    with pytest.raises(WpsUnavailableError):
        macos.socket_process(Path("/tmp/wpsdriverlocal_wps_1"))


def test_invalid_arguments_do_not_send_requests(driver_host: Any) -> None:
    directory, calls = driver_host
    driver = MacDriver(directory)
    for timeout in [0, -1, 121, float("nan")]:
        with pytest.raises(ValueError):
            driver.execute("writer", "return 1;", timeout=timeout)
    with pytest.raises(ValueError):
        driver.execute("writer", "return args;", [float("nan")])
    assert calls == []


@pytest.fixture
def bootstrap_session(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> Iterator[MacSession]:
    monkeypatch.delenv("WPS_DRIVER_SOCKET", raising=False)
    template = tmp_path / "Contents/Resources/office6/mui/en/templates"
    template.mkdir(parents=True)
    (template / "newfile.docx").write_bytes(b"bundled template")
    monkeypatch.setattr(macos, "application_bundle", lambda: tmp_path)
    monkeypatch.setattr(
        macos,
        "temporary_directory",
        partial(tempfile.TemporaryDirectory, dir=tmp_path),
    )
    monkeypatch.setattr(macos, "open_document", Mock())
    monkeypatch.setattr(
        macos,
        "time",
        SimpleNamespace(monotonic=Mock(return_value=0), sleep=Mock()),
    )
    session = MacSession("writer")
    session.executor.driver.directory = tmp_path
    try:
        yield session
    finally:
        assert session.close()


@pytest.mark.parametrize("process", [4242, None])
def test_bootstrap_checks_explicit_endpoint_before_default_discovery(
    bootstrap_session: MacSession,
    monkeypatch: pytest.MonkeyPatch,
    process: int | None,
) -> None:
    session = bootstrap_session
    directory = session.executor.driver.directory
    endpoint = directory / "alternate/wpsdriverlocal_wps_1"
    session.executor.driver.directory = directory / "default"
    monkeypatch.setenv("WPS_DRIVER_SOCKET", str(endpoint))
    monkeypatch.setattr(macos, "sys", SimpleNamespace(platform="darwin"))
    monkeypatch.setattr(
        Path,
        "lstat",
        Mock(return_value=SimpleNamespace(st_mode=stat.S_IFSOCK)),
    )
    peer = Mock(return_value=process)
    live = Mock(return_value={})
    bundle = Mock(return_value=None)
    launch = Mock()
    monkeypatch.setattr(macos, "socket_process", peer)
    monkeypatch.setattr(macos, "live_sockets", live)
    monkeypatch.setattr(macos, "application_bundle", bundle)
    monkeypatch.setattr(macos, "open_document", launch)

    if process is None:
        with pytest.raises(WpsUnavailableError, match="not running"):
            session._bootstrap()
    else:
        session._bootstrap()

    peer.assert_called_once_with(endpoint.resolve())
    live.assert_not_called()
    bundle.assert_not_called()
    launch.assert_not_called()


@pytest.mark.parametrize("suffix", [".docx", ".doc", ".wps"])
def test_bootstrap_waits_for_live_socket_and_opened_document(
    bootstrap_session: MacSession, monkeypatch: pytest.MonkeyPatch, suffix: str
) -> None:
    session = bootstrap_session
    template = (
        session.executor.driver.directory
        / "Contents/Resources/office6/mui/en/templates/newfile.docx"
    )
    template.rename(template.with_suffix(suffix))
    endpoint = session.executor.driver.directory / "wpsdriverlocal_wps_1"
    live = Mock(side_effect=[{}, {}, {endpoint: 4242}])
    monkeypatch.setattr(macos, "live_sockets", live)
    targets: list[Path] = []

    def execute(
        component: str, script: str, args: list[str]
    ) -> dict[str, Any]:
        target = Path(args[0])
        assert target.suffix == suffix
        assert target.read_bytes() == b"bundled template"
        assert "d.FullName===args[0]" in script
        targets.append(target)
        return {"result": len(targets) == 2}

    monkeypatch.setattr(session.executor.driver, "execute", execute)
    session._bootstrap()
    assert live.call_count == 3
    assert len(targets) == 2
    assert targets[0] == targets[1]
    assert not targets[0].exists()


def test_bootstrap_stale_endpoint_times_out_without_sending_script(
    bootstrap_session: MacSession, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(macos, "live_sockets", Mock(return_value={}))
    monkeypatch.setattr(
        macos,
        "time",
        SimpleNamespace(monotonic=Mock(side_effect=[0, 16]), sleep=Mock()),
    )
    execute = Mock()
    monkeypatch.setattr(bootstrap_session.executor.driver, "execute", execute)
    with pytest.raises(WpsUnavailableError, match="startup timed out"):
        bootstrap_session._bootstrap()
    execute.assert_not_called()
