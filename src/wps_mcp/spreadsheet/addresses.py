"""Parse explicit rectangular A1 worksheet addresses."""

import re
from typing import Annotated

from pydantic import AfterValidator


def column_number(value: str) -> int:
    """Convert A1 column letters to their one-based numeric index."""
    result = 0
    for character in value.upper():
        result = result * 26 + ord(character) - ord("A") + 1
    return result


def parse_address(address: str) -> tuple[int, int, int, int]:
    """Parse one bounded A1 rectangle into inclusive row and column bounds."""
    match = re.fullmatch(
        r"\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)"
        r"(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]*))?",
        address,
    )
    if match is None:
        raise ValueError(
            "Use one rectangular A1 address, such as A1 or A1:C10"
        )
    first_column = column_number(match[1])
    first_row = int(match[2])
    last_column = column_number(match[3] or match[1])
    last_row = int(match[4] or match[2])
    if first_column > last_column or first_row > last_row:
        raise ValueError("Range end must not precede its start")
    if last_row > 1048576 or last_column > 16384:
        raise ValueError(
            "Range exceeds the maximum supported worksheet dimensions"
        )
    return (first_row, first_column, last_row, last_column)


def validate_address(value: str) -> str:
    """Validate an A1 rectangle while preserving its supplied spelling."""
    parse_address(value)
    return value


A1Address = Annotated[str, AfterValidator(validate_address)]
