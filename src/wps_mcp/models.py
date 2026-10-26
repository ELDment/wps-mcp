"""Shared, strict inputs and explicit native/JSON boundaries."""

from typing import Any, Literal, TypeAlias

from pydantic import BaseModel, ConfigDict, Field

Component: TypeAlias = Literal["writer", "et", "wpp"]
NativeObject: TypeAlias = Any
JsonObject: TypeAlias = dict[str, Any]
ToolResult: TypeAlias = JsonObject | list[JsonObject] | str | bytes
CellValue: TypeAlias = str | int | float | bool | None


class InputModel(BaseModel):
    """Reject unknown/coerced inputs without altering user text."""

    model_config = ConfigDict(
        extra="forbid", strict=True, allow_inf_nan=False, validate_default=True
    )


class DocumentInput(InputModel):
    """Select an open document by full path or unambiguous name."""

    document: str | None = Field(default=None, min_length=1)
