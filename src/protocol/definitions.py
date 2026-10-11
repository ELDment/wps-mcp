"""Describe domain tools without importing the MCP SDK or native runtime."""

from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass
from typing import Literal, TypeVar

from src.models import Component, JsonObject, ToolResult

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
    operation: str | None = None


DEFINITIONS: dict[str, ToolDefinition] = {}


def tool(
    name: str,
    *,
    component: ToolComponent,
    annotations: JsonObject | None = None,
    operation: str | None = None,
) -> Callable[[F], F]:
    """Attach operation metadata without wrapping the business function."""
    if not re.fullmatch(r"[a-z][a-z0-9_]*", name):
        raise ValueError(f"Tool name must use snake_case: {name}")

    def register(function: F) -> F:
        if operation is not None and not re.fullmatch(
            r"[a-z][a-z0-9_]*", operation
        ):
            raise ValueError(f"Invalid operation: {operation}")
        key = name if operation is None else f"{name}/{operation}"
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
