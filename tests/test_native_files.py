"""Keep native inputs alive when the remote operation has not completed."""

import gc
import http.client
import json
import socket
from contextlib import ExitStack
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from threading import Event, Thread
from types import SimpleNamespace
from typing import Any
from unittest.mock import Mock

import pytest

from src.errors import UnknownOutcomeError
from src.native import files, macos
from src.native.worker import StaWorker
from src.presentation import pictures, themes


@pytest.mark.parametrize("failed", [False, True])
def test_temporary_files_are_cleaned_after_known_outcome(
    tmp_path: Path, failed: bool
) -> None:
    error = ValueError("native operation rejected")
    directory: Path | None = None
    try:
        with files.temporary_directory(prefix="native-", dir=tmp_path) as name:
            directory = Path(name)
            (directory / "input.xml").write_bytes(b"native input")
            if failed:
                raise error
    except ValueError as caught:
        assert caught is error
    assert directory is not None
    assert not directory.exists()


def test_temporary_files_follow_transferred_exit_stack(tmp_path: Path) -> None:
    with ExitStack() as owner:
        with ExitStack() as scope:
            name = scope.enter_context(
                files.temporary_directory(prefix="native-", dir=tmp_path)
            )
            owner.enter_context(scope.pop_all())
        assert Path(name).is_dir()
    assert not Path(name).exists()


@pytest.mark.parametrize("failure", ["timeout", "disconnect"])
@pytest.mark.parametrize("operation", ["theme", "picture"])
def test_unknown_native_call_retains_input_for_delayed_read(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    failure: str,
    operation: str,
) -> None:
    monkeypatch.setattr(files, "sys", SimpleNamespace(platform="darwin"))
    monkeypatch.setenv("WPS_DRIVER_DIR", str(tmp_path))
    started, release, finished = Event(), Event(), Event()
    paths: list[Path] = []
    reads: list[bytes] = []
    calls: list[str] = []

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, format: str, *args: Any) -> None:
            pass

        def do_POST(self) -> None:
            self.rfile.read(int(self.headers["Content-Length"]))
            calls.append(self.path)
            if self.path == "/session":
                value: Any = {"sessionId": "delayed-input"}
            else:
                started.set()
                if failure == "disconnect":
                    self.connection.shutdown(socket.SHUT_RDWR)
                    self.close_connection = True
                if not release.wait(5):
                    return
                try:
                    reads.append(paths[0].read_bytes())
                finally:
                    finished.set()
                value = json.dumps({"result": None})
            payload = json.dumps(
                {"status": 0, "isSuccess": True, "value": value}
            ).encode()
            try:
                self.send_response(200)
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
            except OSError:
                pass

    # Local TCP exercises the real timeout/disconnect path on Windows too.
    host = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    thread = Thread(target=host.serve_forever, daemon=True)
    thread.start()
    driver = macos.MacDriver(tmp_path)
    monkeypatch.setattr(driver, "_select", Mock(return_value=(tmp_path, 4242)))
    monkeypatch.setattr(
        macos,
        "UnixHttpConnection",
        lambda path, timeout: http.client.HTTPConnection(
            "127.0.0.1", host.server_port, timeout=timeout
        ),
    )

    def load(path: str, *args: object) -> None:
        paths.append(Path(path))
        driver.execute("wpp", "return args[0];", [path], timeout=0.2)

    worker = StaWorker("delayed native input")
    try:
        with pytest.raises(UnknownOutcomeError) as caught:
            if operation == "theme":
                worker.execute(
                    themes._load_theme_colors,
                    SimpleNamespace(Load=load),
                    dict.fromkeys(range(1, 13), 0),
                )
            else:
                worker.execute(
                    pictures._insert_picture,
                    SimpleNamespace(Shapes=SimpleNamespace(AddPicture=load)),
                    b"native image payload",
                    ".png",
                    0.0,
                    0.0,
                    None,
                    None,
                    False,
                    "front",
                )
        assert started.is_set() and not finished.is_set()
        assert worker.close()
        gc.collect()
        assert paths[0].is_file()
        assert str(paths[0].parent) in str(caught.value)
        assert "Wait for WPS to finish" in str(caught.value)
        assert "No retry was sent" in str(caught.value)
        assert len(calls) == 2
        expected = paths[0].read_bytes()
        release.set()
        assert finished.wait(3)
        assert reads == [expected]
        assert paths[0].read_bytes() == expected
    finally:
        release.set()
        host.shutdown()
        host.server_close()
        thread.join(timeout=2)
        assert worker.close()
