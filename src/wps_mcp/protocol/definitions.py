"""Describe domain tools without importing the MCP SDK or native runtime."""

from __future__ import annotations

import inspect
import re
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Literal, TypeVar, get_type_hints

from wps_mcp.models import Component, InputModel, JsonObject, ToolResult

F = TypeVar("F", bound=Callable[..., ToolResult])
ToolComponent = Component | Literal["auto"] | None


READ_ANNOTATIONS = {
    "readOnlyHint": True,
    "destructiveHint": False,
    "openWorldHint": False,
}


WRITE_ANNOTATIONS = {
    "readOnlyHint": False,
    "destructiveHint": True,
    "openWorldHint": False,
}


@dataclass(frozen=True, slots=True)
class ToolDefinition:
    """Describe an operation, its native component and MCP hints."""

    name: str
    function: Callable[..., ToolResult]
    component: ToolComponent
    annotations: JsonObject
    operation: str
    input_model: type[InputModel] = field(init=False)
    output_type: object = field(init=False)

    def __post_init__(self) -> None:
        """Validate the operation contract once at registration."""
        if not re.fullmatch(r"[a-z][a-z0-9_]*", self.name):
            raise ValueError(f"Invalid tool name: {self.name}")
        if not re.fullmatch(r"[a-z][a-z0-9_]*", self.operation):
            raise ValueError(f"Invalid operation: {self.operation}")
        if list(inspect.signature(self.function).parameters) != ["params"]:
            raise TypeError(
                f"{self.name} must accept exactly one params model"
            )
        hints = get_type_hints(self.function)
        model = hints.get("params")
        if not isinstance(model, type) or not issubclass(model, InputModel):
            raise TypeError(
                f"{self.name} requires a concrete InputModel annotation"
            )
        if "operation" in model.model_fields:
            raise ValueError("operation is reserved for the registry")
        if "return" not in hints:
            raise TypeError(f"{self.name} requires a return annotation")
        object.__setattr__(self, "input_model", model)
        object.__setattr__(self, "output_type", hints["return"])

    @property
    def identity(self) -> str:
        """Identify an operation independently of its Python module."""
        return f"{self.name}/{self.operation}"


DEFINITIONS: dict[str, ToolDefinition] = {}


def tool(
    name: str,
    *,
    component: ToolComponent,
    annotations: JsonObject | None = None,
    operation: str,
) -> Callable[[F], F]:
    """Attach operation metadata without wrapping the business function."""
    if not re.fullmatch(r"[a-z][a-z0-9_]*", name):
        raise ValueError(f"Tool name must use snake_case: {name}")

    def register(function: F) -> F:
        key = f"{name}/{operation}"
        if key in DEFINITIONS:
            raise ValueError(f"Duplicate operation: {key}")
        defaults: JsonObject = {
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        }
        DEFINITIONS[key] = ToolDefinition(
            name,
            function,
            component,
            defaults | (annotations or {}),
            operation,
        )
        return function

    return register
