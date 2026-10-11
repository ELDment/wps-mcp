"""Navigate presentation windows without changing document targeting."""

import logging

from src.models import NativeObject
from src.native.runtime import get_session

logger = logging.getLogger(__name__)


def goto_slide(application: NativeObject, slide_index: int) -> None:
    """Navigate the selected window without changing the target deck."""
    try:
        window = get_session("wpp").get_window()
        if window is not None:
            window.View.GotoSlide(slide_index)
    except Exception as error:
        logger.debug("Could not navigate the presentation window: %s", error)
