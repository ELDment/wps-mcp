"""Validate tool schemas and dispatch requests across the native boundary."""

from __future__ import annotations

import inspect
from collections.abc import Awaitable, Callable
from functools import reduce
from operator import or_
from typing import Annotated, Any, Literal, cast, get_type_hints

from mcp.server.mcpserver import Image
from mcp.server.mcpserver.exceptions import ToolError
from pydantic import Field, create_model

from src.errors import WpsError
from src.models import Component, InputModel, JsonObject, ToolResult
from src.native.context import CALL_DOCUMENT
from src.native.runtime import get_session
from src.native.worker import run_offloaded
from src.protocol.definitions import ToolDefinition


def make_adapter(
    definition: ToolDefinition | list[ToolDefinition],
) -> Callable[..., Awaitable[ToolResult | Image]]:
    """Build a typed asynchronous dispatcher for one tool operation group."""
    definitions = definition if isinstance(definition, list) else [definition]
    if not definitions:
        raise ValueError("A tool must have at least one operation")
    grouped = definitions[0].operation is not None
    if any((item.operation is not None) != grouped for item in definitions):
        raise ValueError("A tool cannot mix grouped and ungrouped operations")
    models: dict[str | None, type[InputModel]] = {}
    variants = []
    choices = {item.operation: item for item in definitions}
    if len(choices) != len(definitions):
        raise ValueError("Duplicate operation in tool group")
    for item in definitions:
        signature = inspect.signature(item.function)
        model = get_type_hints(item.function).get("params")
        if list(signature.parameters) != ["params"]:
            raise TypeError(
                f"{item.name} must accept exactly one params model"
            )
        if not isinstance(model, type) or not issubclass(model, InputModel):
            raise TypeError(
                f"{item.name} requires a concrete InputModel annotation"
            )
        models[item.operation] = model
        if grouped:
            if "operation" in model.model_fields:
                raise ValueError(
                    "operation is reserved for the registry discriminator"
                )
            fields: dict[str, Any] = {
                "operation": (Literal[item.operation], ...)
            }
            tagged = create_model(
                f"{item.name}_{item.operation}",
                __base__=model,
                **fields,
            )
            variants.append(tagged)
        else:
            variants.append(model)
    input_type = (
        Annotated[reduce(or_, variants), Field(discriminator="operation")]
        if len(variants) > 1
        else variants[0]
    )

    async def invoke(params: InputModel) -> ToolResult | Image:
        operation = getattr(params, "operation", None) if grouped else None
        selected = choices[operation]
        if grouped:
            params = models[operation].model_validate(
                params.model_dump(exclude={"operation"}, exclude_unset=True)
            )
        token = CALL_DOCUMENT.set(getattr(params, "document", None))
        try:
            component = (
                cast(Component, params.model_dump()["component"])
                if selected.component == "auto"
                else selected.component
            )
            if component is None:
                result = await run_offloaded(selected.function, params)
            else:
                session = get_session(component)

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
            label = selected.name + (f"/{operation}" if operation else "")
            raise ToolError(f"{label}: {error}") from None
        finally:
            CALL_DOCUMENT.reset(token)

    # The SDK introspects this concrete model to build and validate its schema.
    invoke.__name__ = definitions[0].name
    invoke.__doc__ = "\n\n".join(
        f"{item.operation}: {inspect.getdoc(item.function) or item.operation}"
        if grouped
        else (inspect.getdoc(item.function) or item.name)
        for item in definitions
    )
    returns = {get_type_hints(item.function)["return"] for item in definitions}
    output_type = JsonObject if returns == {JsonObject} else ToolResult | Image
    if returns == {bytes}:
        output_type = Image
    invoke.__annotations__ = {
        "params": input_type,
        "return": output_type,
    }
    return invoke
