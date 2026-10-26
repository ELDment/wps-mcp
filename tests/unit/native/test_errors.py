"""Preserve safe failure semantics across worker and process boundaries."""

import pytest

from wps_mcp.errors import (
    DocumentNotFoundError,
    NativeError,
    OperationTimeoutError,
    UnknownOutcomeError,
    WpsError,
    WpsUnavailableError,
    decode_error,
    detach_error,
    encode_error,
)
from wps_mcp.models import JsonObject
from wps_mcp.native.javascript import JsError
from wps_mcp.native.objects import is_optional_member_error
from wps_mcp.native.worker import StaWorker


@pytest.mark.parametrize(
    "kind",
    [
        WpsError,
        WpsUnavailableError,
        DocumentNotFoundError,
        OperationTimeoutError,
        UnknownOutcomeError,
        ValueError,
        TypeError,
        NotImplementedError,
        OSError,
    ],
)
def test_failure_type_survives_thread_and_process_round_trip(
    kind: type[Exception],
) -> None:
    error = kind("native failure")

    def fail() -> None:
        raise error from RuntimeError("native cause")

    worker = StaWorker("error contract")
    try:
        with pytest.raises(kind) as caught:
            worker.execute(fail)
        assert type(caught.value) is kind
        assert caught.value is not error
        assert caught.value.__cause__ is None
        decoded = decode_error(encode_error(error))
        assert type(decoded) is kind
        assert str(decoded) == "native failure"
        assert decoded.__traceback__ is None
        assert decoded.__cause__ is None
        assert decoded.__context__ is None
    finally:
        assert worker.close()


def test_native_error_preserves_only_message_and_hresult() -> None:
    error = NativeError("native call rejected", -2147467259)
    decoded = decode_error(encode_error(error))
    assert type(decoded) is NativeError
    assert decoded.hresult == error.hresult
    assert decoded is not error
    assert decoded.__traceback__ is None


def test_unknown_subclasses_cannot_transfer_arbitrary_attributes() -> None:
    class CustomError(WpsError):
        native: object

    error = CustomError("private failure")
    error.native = object()
    detached = detach_error(error)
    assert type(detached) is WpsError
    assert not hasattr(detached, "native")
    assert (
        type(
            decode_error(
                {
                    "type": "CustomError",
                    "message": "private failure",
                    "hresult": None,
                }
            )
        )
        is WpsError
    )


@pytest.mark.parametrize(
    "payload",
    [
        {},
        {"type": "WpsError", "message": "missing code"},
        {"type": 1, "message": "bad type", "hresult": None},
        {"type": "WpsError", "message": [], "hresult": None},
        {"type": "NativeError", "message": "bad code", "hresult": True},
        {"type": "WpsError", "message": "bad code", "hresult": 1},
        {"type": "WpsError", "message": "extra", "hresult": None, "x": 1},
    ],
)
def test_malformed_remote_errors_are_rejected(payload: JsonObject) -> None:
    with pytest.raises(WpsError, match="Malformed native error"):
        decode_error(payload)


@pytest.mark.parametrize(
    ("error", "optional"),
    [
        (AttributeError("member missing"), True),
        (NativeError("missing", -2147352573), True),
        (NativeError("unknown", -2147352570), True),
        (NativeError("not implemented", -2147467263), True),
        (NativeError("disconnected", -2147467259), False),
        (UnknownOutcomeError("reply lost"), False),
        (OperationTimeoutError("call running"), False),
        (WpsUnavailableError("disconnected"), False),
        (RuntimeError("bug"), False),
        (ValueError("invalid"), False),
        (
            JsError(
                "get", "Name", NativeError("Error: Function excute faild")
            ),
            True,
        ),
        (JsError("get", "Name", NativeError("other failure")), False),
    ],
)
def test_optional_member_failures_do_not_hide_uncertain_execution(
    error: Exception,
    optional: bool,
) -> None:
    assert is_optional_member_error(error) is optional
