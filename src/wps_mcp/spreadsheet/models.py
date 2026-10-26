"""Define shared bounded worksheet and range selectors."""

from typing import Annotated

from pydantic import Field

from wps_mcp.models import DocumentInput
from wps_mcp.spreadsheet.addresses import A1Address

SheetReference = (
    Annotated[int, Field(ge=1)] | Annotated[str, Field(min_length=1)]
)


MAX_CELLS = 100000


class SheetInput(DocumentInput):
    """Select a worksheet by name or one-based index."""

    sheet: SheetReference = Field(
        default=1, description="Worksheet name or one-based index"
    )


class RangeInput(SheetInput):
    """Select an explicit document range using native coordinates."""

    address: A1Address = Field(
        description=(
            "One rectangular A1 range, e.g. A1:C10; no names or whole columns"
        )
    )
