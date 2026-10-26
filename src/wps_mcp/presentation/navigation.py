"""Navigate presentation windows without changing document targeting."""

import logging

from wps_mcp.native.objects import is_optional_member_error
from wps_mcp.native.runtime import get_session

logger = logging.getLogger(__name__)


def goto_slide(slide_index: int) -> None:
    """Navigate the selected window without changing the target deck."""
    try:
        window = get_session("wpp").get_window()
        if window is not None:
            window.View.GotoSlide(slide_index)
    except Exception as error:
        if not is_optional_member_error(error):
            raise
        logger.debug("Could not navigate the presentation window: %s", error)
