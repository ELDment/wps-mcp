"""Validate rule expressions and anchor relative native formulas."""

import logging
import math
import re
from collections.abc import Iterator
from contextlib import contextmanager
from typing import Annotated, Literal

from pydantic import AfterValidator

from src.errors import WpsError
from src.models import NativeObject
from src.native.context import CALL_DOCUMENT
from src.native.runtime import get_session
from src.native.selection import document_key
from src.spreadsheet.models import RangeInput

Comparison = Literal[
    "between",
    "not_between",
    "equal",
    "not_equal",
    "greater",
    "less",
    "greater_equal",
    "less_equal",
]


OPERATORS = {
    "between": 1,
    "not_between": 2,
    "equal": 3,
    "not_equal": 4,
    "greater": 5,
    "less": 6,
    "greater_equal": 7,
    "less_equal": 8,
}


LOGGER = logging.getLogger(__name__)


def validate_comparison(
    operator: Comparison | None, formula1: str, formula2: str | None
) -> None:
    """Reject incomplete comparisons before a native rule can be modified."""
    if operator is None:
        raise ValueError("operator is required for this rule type")
    if (operator in {"between", "not_between"}) != (formula2 is not None):
        raise ValueError("formula2 is required only for between/not_between")
    for formula in (formula1, formula2):
        if formula is None or formula.startswith("="):
            continue
        if not re.fullmatch(
            r"[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?",
            formula,
        ):
            raise ValueError(
                "Comparison bounds must be finite numbers or "
                "formulas beginning with '='"
            )
        if not math.isfinite(float(formula)):
            raise ValueError("Comparison bounds must be finite numbers")


@contextmanager
def formula_range(params: RangeInput, target: NativeObject) -> Iterator[None]:
    """Anchor relative formulas and restore the target window selection.

    Select the target range first cell, then restore its own sheet and
    selection before restoring the previously active window.
    """
    session = get_session("et")
    document = session.get_document(params.document)
    token = CALL_DOCUMENT.set(str(document.FullName))
    try:
        with session.command_window() as window:
            application = session.get_application()
            if document_key(
                application.ActiveWorkbook.FullName
            ) != document_key(document.FullName):
                raise WpsError(
                    "Could not activate the requested workbook for "
                    "relative formulas"
                )
            previous_sheet = window.ActiveSheet
            previous_selection = application.Selection
            failed = False
            try:
                target.Parent.Activate()
                target.Cells.Item(1, 1).Select()
                yield
            except BaseException:
                failed = True
                raise
            finally:
                errors = []
                for native in (previous_sheet, previous_selection):
                    if native is None:
                        continue
                    try:
                        if native is previous_sheet:
                            native.Activate()
                        else:
                            native.Select()
                    except Exception as error:
                        errors.append(str(error))
                if errors:
                    detail = "; ".join(errors)
                    if failed:
                        LOGGER.error(
                            "Could not restore the target workbook "
                            "selection: %s",
                            detail,
                        )
                    else:
                        raise WpsError(
                            "Could not restore the target workbook selection: "
                            f"{detail}"
                        )
    finally:
        CALL_DOCUMENT.reset(token)


def validate_formula(value: str) -> str:
    """Reject blank expressions and control characters before native writes."""
    if (
        not value.strip()
        or (value.startswith("=") and not value[1:].strip())
        or any(ord(character) < 32 for character in value)
    ):
        raise ValueError(
            "Formulas must not be blank or contain control characters"
        )
    return value


RuleFormula = Annotated[str, AfterValidator(validate_formula)]
