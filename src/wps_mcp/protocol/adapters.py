"""Validate tool schemas and dispatch requests across the native boundary."""

from __future__ import annotations

import inspect
from collections.abc import Awaitable, Callable
from functools import reduce
from operator import or_
from typing import Annotated, Any, Literal, cast

from mcp.server.mcpserver import Image
from mcp.server.mcpserver.exceptions import ToolError
from pydantic import Field, create_model

from wps_mcp.errors import WpsError
from wps_mcp.models import Component, InputModel, JsonObject, ToolResult
from wps_mcp.native.context import CALL_DOCUMENT, CALL_OPERATION
from wps_mcp.native.runtime import get_runtime
from wps_mcp.native.worker import run_offloaded
from wps_mcp.protocol.definitions import ToolDefinition


def make_adapter(
    definitions: list[ToolDefinition],
) -> Callable[..., Awaitable[ToolResult | Image]]:
    """Build a typed asynchronous dispatcher for one tool operation group."""
    if not definitions:
        raise ValueError("A tool must have at least one operation")
    if any(item.name != definitions[0].name for item in definitions):
        raise ValueError("An operation group must have one tool name")
    variants = []
    choices = {item.operation: item for item in definitions}
    if len(choices) != len(definitions):
        raise ValueError("Duplicate operation in tool group")
    for item in definitions:
        fields: dict[str, Any] = {"operation": (Literal[item.operation], ...)}
        variants.append(
            create_model(
                f"{item.name}_{item.operation}",
                __base__=item.input_model,
                **fields,
            )
        )
    input_type = (
        Annotated[reduce(or_, variants), Field(discriminator="operation")]
        if len(variants) > 1
        else variants[0]
    )

    async def invoke(params: InputModel) -> ToolResult | Image:
        operation = str(params.model_dump()["operation"])
        selected = choices[operation]
        params = selected.input_model.model_validate(
            params.model_dump(exclude={"operation"}, exclude_unset=True)
        )
        token = CALL_DOCUMENT.set(getattr(params, "document", None))
        operation_token = CALL_OPERATION.set(selected.identity)
        try:
            component = (
                cast(Component, params.model_dump()["component"])
                if selected.component == "auto"
                else selected.component
            )
            if component is None:
                result = await run_offloaded(selected.function, params)
            else:
                session = get_runtime().sessions[component]

                def execute_native() -> ToolResult:
                    return session.execute(selected.function, params)

                result = await run_offloaded(execute_native)
            if isinstance(result, dict) and result.get("error"):
                raise WpsError(str(result["error"]))
            return (
                Image(data=result, format="png")
                if isinstance(result, bytes)
                else result
            )
        except ToolError:
            raise
        except (
            WpsError,
            ValueError,
            TypeError,
            OSError,
            NotImplementedError,
        ) as error:
            raise ToolError(f"{selected.identity}: {error}") from None
        finally:
            CALL_OPERATION.reset(operation_token)
            CALL_DOCUMENT.reset(token)

    # The SDK introspects this concrete model to build and validate its schema.
    invoke.__name__ = definitions[0].name
    invoke.__doc__ = "\n\n".join(
        f"{item.operation}: {inspect.getdoc(item.function) or item.operation}"
        for item in definitions
    )
    returns = {item.output_type for item in definitions}
    output_type = JsonObject if returns == {JsonObject} else ToolResult | Image
    if returns == {bytes}:
        output_type = Image
    invoke.__annotations__ = {
        "params": input_type,
        "return": output_type,
    }
    return invoke
