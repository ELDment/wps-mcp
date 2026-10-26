"""Format native date objects and OLE Automation date values consistently."""

import math
from datetime import datetime, timedelta

OLE_EPOCH = datetime(1899, 12, 30)


def date_text(value: object) -> str:
    """Preserve date text while interpreting numeric OLE DATE fields.

    OLE DATE stores the date as a signed day count and the time as the
    absolute fractional day. For example, -1.25 is 1899-12-29 at 06:00,
    while both -0.25 and 0.25 represent 1899-12-30 at 06:00.
    """
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            serial = float(value)
            if not math.isfinite(serial):
                raise ValueError("OLE DATE must be finite")
            days = math.trunc(serial)
            moment = OLE_EPOCH + timedelta(
                days=days, seconds=abs(serial - days) * 86400
            )
        except OverflowError as error:
            raise ValueError(
                "OLE DATE is outside the supported date range"
            ) from error
        return str(moment)
    return str(value)
