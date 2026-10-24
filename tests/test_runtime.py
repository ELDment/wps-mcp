"""Exercise native ownership and cancellation without launching WPS."""

import ntpath
from collections.abc import Iterator
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager, nullcontext
from threading import Event, get_ident
from types import SimpleNamespace
from typing import Literal, TypedDict
from unittest.mock import Mock

import anyio
import pytest

from src.errors import (
    DocumentNotFoundError,
    OperationTimeoutError,
    WpsError,
)
from src.models import Component
from src.native import selection
from src.native.context import CALL_DOCUMENT
from src.native.runtime import Runtime, get_runtime
from src.native.selection import select_document
from src.native.state import ensure_plain_result, temporary_attributes
from src.native.windows import ComSession
from src.native.worker import StaWorker, run_offloaded


class WorkerTimeouts(TypedDict, total=False):
    queue_timeout: float
    call_timeout: float
    startup_timeout: float


@pytest.mark.parametrize("timeout", [0.0, -1.0, float("inf"), float("nan")])
@pytest.mark.parametrize(
    "name", ["queue_timeout", "call_timeout", "startup_timeout"]
)
def test_worker_rejects_invalid_timeouts(
    name: Literal["queue_timeout", "call_timeout", "startup_timeout"],
    timeout: float,
) -> None:
    timeouts: WorkerTimeouts = {}
    timeouts[name] = timeout
    with pytest.raises(ValueError, match="finite and positive"):
        StaWorker("invalid", **timeouts)


def test_native_scope_runs_and_releases_on_one_thread() -> None:
    events: list[tuple[str, int]] = []

    @contextmanager
    def scope() -> Iterator[None]:
        events.append(("enter", get_ident()))
        try:
            yield
        finally:
            events.append(("exit", get_ident()))

    worker = StaWorker("test", scope)
    try:
        native_thread = worker.execute(get_ident)
        assert (
            worker.execute(lambda: worker.execute(get_ident)) == native_thread
        )
    finally:
        assert worker.close()
    assert events == [("enter", native_thread), ("exit", native_thread)]
    assert native_thread != get_ident()
    assert worker.close()
    with pytest.raises(WpsError, match="closed"):
        worker.execute(get_ident)


def test_timeout_does_not_replay_or_replace_worker() -> None:
    entered, release, completed = Event(), Event(), Event()
    calls: list[int] = []

    def native_write() -> None:
        calls.append(get_ident())
        entered.set()
        try:
            assert release.wait(3)
        finally:
            completed.set()

    worker = StaWorker("timeout", call_timeout=0.03)
    try:
        with pytest.raises(OperationTimeoutError, match="may still finish"):
            worker.execute(native_write)
        assert entered.is_set()
        with pytest.raises(OperationTimeoutError, match="No new work"):
            worker.execute(lambda: calls.append(0))
        assert not worker.close(timeout=0.001)
    finally:
        release.set()
        assert completed.wait(1)
        assert worker.close()
    assert len(calls) == 1


def test_pending_queue_timeout_prevents_mutation() -> None:
    entered, release = Event(), Event()
    applied: list[bool] = []
    worker = StaWorker("queue", queue_timeout=0.03, call_timeout=3)

    def block() -> None:
        entered.set()
        assert release.wait(2)

    try:
        with ThreadPoolExecutor(max_workers=1) as executor:
            first = executor.submit(worker.execute, block)
            assert entered.wait(1)
            try:
                with pytest.raises(
                    OperationTimeoutError, match="before execution"
                ):
                    worker.execute(lambda: applied.append(True))
            finally:
                release.set()
            first.result(timeout=1)
        assert applied == []
    finally:
        worker.close()


def test_exception_traceback_does_not_cross_thread() -> None:
    worker = StaWorker("error")

    def fail() -> None:
        raise ValueError("native failure") from RuntimeError("native cause")

    try:
        with pytest.raises(ValueError, match="native failure") as caught:
            worker.execute(fail)
        assert caught.value.__cause__ is None
        traceback = caught.value.__traceback__
        frames = []
        while traceback is not None:
            frames.append(traceback.tb_frame.f_code.co_name)
            traceback = traceback.tb_next
        assert "fail" not in frames
    finally:
        worker.close()


def test_exception_subclasses_cannot_terminate_worker() -> None:
    from src.presentation.style import ApplyStyleInput

    worker = StaWorker("validation", call_timeout=1)
    try:
        with pytest.raises(ValueError, match="decode"):
            worker.execute(lambda: b"\xff".decode("utf-8"))
        invalid = {
            "slide_index": 1,
            "shapes": [1],
            "operations": [{"kind": "fill", "options": {"unknown": True}}],
        }
        with pytest.raises(ValueError, match="unknown"):
            worker.execute(ApplyStyleInput.model_validate, invalid)
        assert worker.execute(lambda: "healthy") == "healthy"
    finally:
        assert worker.close()


def test_cancellation_leaves_running_call_owned_until_completion() -> None:
    entered, release, completed = Event(), Event(), Event()
    worker = StaWorker("cancel", call_timeout=3)

    def native_write() -> None:
        entered.set()
        assert release.wait(3)
        completed.set()

    async def scenario() -> None:
        async with anyio.create_task_group() as group:
            group.start_soon(run_offloaded, worker.execute, native_write)
            assert await anyio.to_thread.run_sync(entered.wait, 1)
            group.cancel_scope.cancel()

    try:
        anyio.run(scenario)
        assert not completed.is_set()
        assert not worker.close(timeout=0.001)
    finally:
        release.set()
        assert worker.close()
    assert completed.is_set()


def test_failed_scope_initialization_does_not_execute() -> None:
    @contextmanager
    def broken() -> Iterator[None]:
        raise RuntimeError("init failed")
        yield

    worker = StaWorker("startup", broken)
    with pytest.raises(WpsError, match="init failed"):
        worker.execute(lambda: pytest.fail("must not execute"))
    assert worker.close()


def test_document_selection_never_uses_foreground_heuristics(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(selection, "path", ntpath)
    first = SimpleNamespace(Name="same.docx", FullName=r"C:\one\same.docx")
    second = SimpleNamespace(Name="same.docx", FullName=r"C:\two\same.docx")
    assert select_document([first, second], r"c:/ONE/same.docx") is first
    for wanted in (None, "same", "same.docx", "missing", True, 0, " "):
        with pytest.raises(DocumentNotFoundError):
            select_document([first, second], wanted)
    with pytest.raises(DocumentNotFoundError):
        select_document([first], r"C:\missing\same.docx")


def test_plain_result_rejects_native_objects_and_cycles() -> None:
    valid = {"values": [1, "text", None, b"PNG", {"saved": True}]}
    assert ensure_plain_result(valid) is valid
    with pytest.raises(WpsError, match="Only plain"):
        ensure_plain_result({"nested": [SimpleNamespace()]})
    recursive: list[object] = []
    recursive.append(recursive)
    with pytest.raises(WpsError, match="cycles"):
        ensure_plain_result(recursive)


def test_runtime_isolates_sessions_and_restores_context() -> None:
    with Runtime() as first:
        with Runtime() as second:
            assert get_runtime() is second
            assert first.sessions["wpp"] is not second.sessions["wpp"]
        assert get_runtime() is first
    with pytest.raises(WpsError, match="outside"):
        get_runtime()


def test_deferred_cleanup_waits_for_all_workers_despite_errors() -> None:
    events: list[str] = []

    def factory(component: Component) -> ComSession:
        session = Mock(spec=ComSession)
        session.worker = Mock()
        session.wait.side_effect = lambda: events.append(component)
        session.close.side_effect = [
            False,
            WpsError("cleanup") if component == "writer" else True,
        ]
        return session

    runtime = Runtime(factory)
    runtime.resources.callback(lambda: events.append("resources"))
    runtime.close()
    assert runtime.cleanup_thread is not None
    runtime.cleanup_thread.join(timeout=1)
    assert not runtime.cleanup_thread.is_alive()
    assert events == ["writer", "et", "wpp", "resources"]


def test_session_context_and_native_result_boundary() -> None:
    document = SimpleNamespace(Name="test", FullName="test")
    collection = SimpleNamespace(Count=1, Item=lambda index: document)
    application = SimpleNamespace(Name="Fake WPS", Documents=collection)
    session = ComSession(
        "writer",
        connector=lambda prog_id, launch: application,
        scope=nullcontext,
    )
    token = CALL_DOCUMENT.set("test")
    try:
        assert session.execute(lambda: session.get_document().Name) == "test"
        with pytest.raises(WpsError, match="Only plain"):
            session.execute(session.get_document)
        with pytest.raises(WpsError, match="only be accessed"):
            session.get_application()
    finally:
        CALL_DOCUMENT.reset(token)
        session.close()
    assert session.application is None


@pytest.mark.parametrize("component", ["writer", "et", "wpp"])
@pytest.mark.parametrize("visible", [None, True, False])
@pytest.mark.parametrize("initial_visible", [False, True])
def test_session_visibility(
    component: Component, visible: bool | None, initial_visible: bool
) -> None:
    application = SimpleNamespace(Name="Fake WPS", Visible=initial_visible)
    connector = Mock(return_value=application)
    session = ComSession(component, connector=connector, scope=nullcontext)
    try:
        if component == "wpp" and visible is False:
            with pytest.raises(ValueError, match="use visible=true"):
                session.execute(lambda: session.connect(visible))
            connector.assert_not_called()
            assert application.Visible is initial_visible
        else:
            actual = session.execute(lambda: session.connect(visible).Visible)
            assert actual is (initial_visible if visible is None else visible)
            connector.assert_called_once_with(session.spec.prog_id, True)
    finally:
        assert session.close()


@pytest.mark.parametrize("initial_visible", [False, True])
def test_wpp_rejects_hiding_an_existing_connection(
    initial_visible: bool,
) -> None:
    application = SimpleNamespace(Name="Fake WPS", Visible=initial_visible)
    connector = Mock(return_value=application)
    session = ComSession("wpp", connector=connector, scope=nullcontext)
    try:
        assert session.execute(lambda: session.connect().Visible) is (
            initial_visible
        )
        with pytest.raises(ValueError, match="use visible=true"):
            session.execute(lambda: session.connect(False))
        assert session.execute(lambda: session.get_application().Visible) is (
            initial_visible
        )
        connector.assert_called_once_with(session.spec.prog_id, True)
    finally:
        assert session.close()


def test_temporary_properties_restore_when_body_or_setter_fails() -> None:
    state = SimpleNamespace(visible=True, alerts=7)
    with pytest.raises(ValueError, match="failure"):
        with temporary_attributes(state, visible=False, alerts=0):
            assert state.visible is False
            raise ValueError("failure")
    assert vars(state) == {"visible": True, "alerts": 7}

    class State:
        visible = True

        @property
        def read_only(self) -> int:
            return 1

    partial = State()
    with pytest.raises(AttributeError):
        with temporary_attributes(partial, visible=False, read_only=2):
            pytest.fail("must fail before body")
    assert partial.visible is True
