"""Scoped native undo and redo commands."""

from pydantic import Field

from wps_mcp.models import DocumentInput, JsonObject
from wps_mcp.native.runtime import get_session
from wps_mcp.protocol.definitions import tool


class HistoryInput(DocumentInput):
    """A bounded number of native history commands."""

    times: int = Field(
        default=1, ge=1, description="Number of undo operations to perform"
    )


def _run_history_command(command: str, times: int) -> int:
    """Run target-window history commands and restore activation."""
    session = get_session("wpp")
    app = session.get_application()
    with session.command_window():
        count = 0
        for _ in range(times):
            if not app.CommandBars.GetEnabledMso(command):
                break
            app.CommandBars.ExecuteMso(command)
            count += 1
    return count


@tool(
    "wps_ppt_history",
    component="wpp",
    annotations={"title": "Undo"},
    operation="undo",
)
def undo(params: HistoryInput) -> JsonObject:
    """Undo recent actions in WPS Presentation.

    Performs one or more undo operations. Stops early if no more actions
    can be undone. Returns the number of actions actually undone.
    """
    return {
        "success": True,
        "actions_undone": _run_history_command("Undo", params.times),
    }


@tool(
    "wps_ppt_history",
    component="wpp",
    annotations={"title": "Redo"},
    operation="redo",
)
def redo(params: HistoryInput) -> JsonObject:
    """Redo recently undone actions in WPS Presentation.

    Performs one or more redo operations. Stops early if no more actions
    can be redone. Returns the number of actions actually redone.
    """
    return {
        "success": True,
        "actions_redone": _run_history_command("Redo", params.times),
    }


@tool(
    "wps_ppt_history",
    component="wpp",
    annotations={"title": "Start Undo Entry"},
    operation="begin",
)
def start_undo_entry(params: DocumentInput) -> JsonObject:
    """Start a new undo entry in WPS Presentation.

    Call this before a batch of operations so that all subsequent
    changes can be undone with a single Ctrl+Z. Useful for grouping
    multiple tool calls into one undoable action.
    """
    session = get_session("wpp")
    with session.command_window():
        session.get_application().StartNewUndoEntry()
    return {"success": True}
