"""Validate worksheet and native table or pivot names."""

import re
from typing import Annotated

from pydantic import AfterValidator

from src.spreadsheet.addresses import parse_address


def validate_sheet_name(value: str) -> str:
    """Reject unsupported worksheet names without trimming user text."""
    if (
        not value.strip()
        or len(value) > 31
        or any(character in value for character in "\\/?*[]:")
    ):
        raise ValueError(
            "Worksheet names must be 1–31 characters and "
            "cannot contain \\ / ? * [ ] :"
        )
    if (
        value.startswith("'")
        or value.endswith("'")
        or any(ord(character) < 32 for character in value)
    ):
        raise ValueError(
            "Worksheet names cannot start/end with an "
            "apostrophe or contain control characters"
        )
    return value


def validate_object_name(value: str) -> str:
    """Use names accepted by native tables without implicit native renaming."""
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]{0,254}", value):
        raise ValueError(
            "Names require 1–255 ASCII letters, digits, underscores or dots"
        )
    if value.casefold() in {"r", "c"} or re.fullmatch(
        r"R[1-9][0-9]*C[1-9][0-9]*", value, re.I
    ):
        raise ValueError(
            "Names cannot be cell references or reserved R/C identifiers"
        )
    try:
        parse_address(value)
    except ValueError:
        return value
    raise ValueError("Names cannot be cell references")


SheetName = Annotated[str, AfterValidator(validate_sheet_name)]
ObjectName = Annotated[str, AfterValidator(validate_object_name)]
