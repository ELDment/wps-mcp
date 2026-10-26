"""Exchange bounded JSON frames over a private local socket."""

import json
import math
import socket
import struct
from typing import Any, NoReturn

from wps_mcp.errors import WpsError
from wps_mcp.models import JsonObject

MAX_FRAME_SIZE = 32 * 1024 * 1024


def send_message(channel: socket.socket, message: JsonObject) -> None:
    """Send one complete frame without exposing a network listener."""
    payload = json.dumps(
        message, ensure_ascii=True, allow_nan=False, separators=(",", ":")
    ).encode("utf-8")
    if len(payload) > MAX_FRAME_SIZE:
        raise WpsError("Native request or result exceeds 32 MiB")
    channel.sendall(struct.pack("!I", len(payload)) + payload)


def _read_exact(channel: socket.socket, size: int) -> bytes:
    """Read a complete frame or report an ambiguous worker disconnect."""
    chunks = bytearray()
    while len(chunks) < size:
        chunk = channel.recv(size - len(chunks))
        if not chunk:
            raise WpsError(
                "WPS worker disconnected; native changes may have completed. "
                "Inspect the document before repeating a write."
            )
        chunks.extend(chunk)
    return bytes(chunks)


def receive_message(channel: socket.socket) -> JsonObject:
    """Reject malformed or oversized messages before allocating the body."""
    size = struct.unpack("!I", _read_exact(channel, 4))[0]
    if not 0 < size <= MAX_FRAME_SIZE:
        raise WpsError("Invalid native message length")
    try:
        value: Any = json.loads(
            _read_exact(channel, size),
            parse_constant=_invalid_number,
            parse_float=_finite_float,
        )
    except (ValueError, UnicodeError) as error:
        raise WpsError("Invalid native JSON message") from error
    if not isinstance(value, dict):
        raise WpsError("Native messages must be JSON objects")
    return value


def _invalid_number(value: str) -> NoReturn:
    """Reject JSON extensions such as NaN and Infinity."""
    raise ValueError(f"Non-finite JSON number: {value}")


def _finite_float(value: str) -> float:
    """Reject finite-looking JSON literals that overflow a Python float."""
    result = float(value)
    if not math.isfinite(result):
        _invalid_number(value)
    return result
