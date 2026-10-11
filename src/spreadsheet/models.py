"""Define shared bounded worksheet and range selectors."""

from pydantic import (
    Field,
    field_validator,
)

from src.models import DocumentInput
from src.spreadsheet.addresses import A1Address

SheetReference = int | str


MAX_CELLS = 100000


class SheetInput(DocumentInput):
    """Select a worksheet by name or one-based index."""

    sheet: SheetReference = Field(
        default=1, description="Worksheet name or one-based index"
    )

    @field_validator("sheet")
    @classmethod
    def validate_sheet(cls, value: SheetReference) -> SheetReference:
        """Reject empty worksheet names and nonpositive indices."""
        if isinstance(value, int) and value < 1:
            raise ValueError("Worksheet index must be at least 1")
        if isinstance(value, str) and (not value):
            raise ValueError("Worksheet name must not be empty")
        return value


class RangeInput(SheetInput):
    """Select an explicit document range using native coordinates."""

    address: A1Address = Field(
        description=(
            "One rectangular A1 range, e.g. A1:C10; no names or whole columns"
        )
    )
