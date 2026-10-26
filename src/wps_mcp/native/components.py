"""Describe registered WPS applications and their native file formats."""

from collections.abc import Mapping
from dataclasses import dataclass

from wps_mcp.models import Component

BUSY_CODES = {-2147418111, -2147417846}


@dataclass(frozen=True, slots=True)
class ComponentSpec:
    """Describe one COM application and the formats it can save."""

    name: str
    prog_id: str
    collection: str
    formats: Mapping[str, int | None]


SPECS: dict[Component, ComponentSpec] = {
    "writer": ComponentSpec(
        "WPS Writer",
        "KWPS.Application",
        "Documents",
        {".docx": 12, ".doc": 0, ".wps": None, ".rtf": 6, ".txt": 2},
    ),
    "et": ComponentSpec(
        "WPS ET",
        "KET.Application",
        "Workbooks",
        {".xlsx": 51, ".xls": 56, ".et": None, ".csv": 62},
    ),
    "wpp": ComponentSpec(
        "WPS WPP",
        "KWPP.Application",
        "Presentations",
        {".pptx": 24, ".ppt": 1, ".dps": None},
    ),
}
