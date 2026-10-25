"""Execute registered Linux operations over an inherited private socket."""

import base64
import logging
import os
import socket
import sys
from typing import get_type_hints

from src.errors import WpsError
from src.models import Component, InputModel, JsonObject
from src.native.components import SPECS
from src.native.context import CALL_DOCUMENT
from src.native.linux import RpcSession
from src.native.runtime import Runtime
from src.native.transport import receive_message, send_message
from src.protocol.catalog import load_definitions
from src.protocol.definitions import DEFINITIONS


def dispatch(
    request: JsonObject, runtime: Runtime, component: Component
) -> JsonObject:
    """Resolve only known operations and revalidate inputs in the worker."""
    identity = request.get("id")
    try:
        if type(identity) is not int or identity <= 0:
            raise ValueError("Request id must be a positive integer")
        if request.get("document") is not None and not isinstance(
            request["document"], str
        ):
            raise ValueError("Request document must be a string or null")
        definition = next(
            (
                definition
                for definition in DEFINITIONS.values()
                if definition.function.__module__ == request.get("module")
                and definition.function.__qualname__ == request.get("function")
            ),
            None,
        )
        if definition is None:
            raise ValueError("Unknown native operation")
        model = get_type_hints(definition.function)["params"]
        if not isinstance(model, type) or not issubclass(model, InputModel):
            raise TypeError("Native operation has no input model")
        params = model.model_validate(request.get("params"))
        target = (
            getattr(params, "component", None)
            if definition.component == "auto"
            else definition.component
        )
        if target != component:
            raise ValueError("Operation belongs to a different WPS component")
        token = CALL_DOCUMENT.set(request.get("document"))
        try:
            result = runtime.sessions[component].execute(
                definition.function, params
            )
        finally:
            CALL_DOCUMENT.reset(token)
        if isinstance(result, bytes):
            return {
                "id": identity,
                "encoding": "base64",
                "result": base64.b64encode(result).decode("ascii"),
            }
        return {"id": identity, "result": result}
    except Exception as error:
        return {
            "id": identity,
            "error": {
                "type": type(error).__name__,
                "message": str(error),
                "hresult": getattr(error, "hresult", None),
            },
        }


def serve(channel: socket.socket, component: Component) -> None:
    """Keep Qt and all native operations on this process's main thread."""
    load_definitions()
    with Runtime(RpcSession) as runtime:
        while True:
            try:
                request = receive_message(channel)
            except WpsError:
                return
            if request.get("kind") == "close":
                return
            if request.get("kind") != "call":
                raise WpsError("Unknown native message kind")
            response = dispatch(request, runtime, component)
            send_message(channel, response)


def main() -> None:
    """Use a private inherited descriptor; reserve stdout for no protocol."""
    logging.basicConfig(level=logging.INFO, stream=sys.stderr)
    if len(sys.argv) != 3 or sys.argv[2] not in SPECS:
        raise SystemExit("A private socket and WPS component are required")
    os.dup2(sys.stderr.fileno(), sys.stdout.fileno())
    with socket.socket(fileno=int(sys.argv[1])) as channel:
        channel.set_inheritable(False)
        serve(channel, sys.argv[2])


if __name__ == "__main__":
    main()
