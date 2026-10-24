"""Inspect and configure bounded native cell validation."""

from typing import Literal, Self

from pydantic import Field, model_validator

from src.errors import NativeError
from src.models import JsonObject, NativeObject
from src.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)
from src.spreadsheet.access import resolve_range
from src.spreadsheet.addresses import parse_address
from src.spreadsheet.formulas import (
    OPERATORS,
    Comparison,
    RuleFormula,
    formula_range,
    validate_comparison,
)
from src.spreadsheet.models import RangeInput

VALIDATION_TYPES = {
    "any": 0,
    "whole": 1,
    "decimal": 2,
    "list": 3,
    "date": 4,
    "time": 5,
    "text_length": 6,
    "custom": 7,
}


ALERT_STYLES = {"stop": 1, "warning": 2, "information": 3}


class SetValidationInput(RangeInput):
    """Define cell validation, prompts, and native alert behavior."""

    validation_type: Literal[
        "whole", "decimal", "list", "date", "time", "text_length", "custom"
    ]
    operator: Comparison | None = None
    formula1: RuleFormula = Field(min_length=1, max_length=255)
    formula2: RuleFormula | None = Field(
        default=None, min_length=1, max_length=255
    )
    alert_style: Literal["stop", "warning", "information"] = "stop"
    ignore_blank: bool = True
    in_cell_dropdown: bool | None = None
    show_input: bool = True
    input_title: str = Field(default="", max_length=32)
    input_message: str = Field(default="", max_length=255)
    show_error: bool = True
    error_title: str = Field(default="", max_length=32)
    error_message: str = Field(default="", max_length=225)

    @model_validator(mode="after")
    def validate_rule(self) -> Self:
        """Reject incompatible native rule types, operators, and formulas."""
        if self.validation_type in {"list", "custom"}:
            if self.operator is not None or self.formula2 is not None:
                raise ValueError(
                    "List/custom validation does not accept operator "
                    "or formula2"
                )
            if (
                self.validation_type == "custom"
                and not self.formula1.startswith("=")
            ):
                raise ValueError(
                    "Custom validation formula1 must begin with '='"
                )
        else:
            validate_comparison(self.operator, self.formula1, self.formula2)
            for formula in (self.formula1, self.formula2):
                if formula is None or formula.startswith("="):
                    continue
                number = float(formula)
                if (
                    self.validation_type in {"whole", "text_length"}
                    and not number.is_integer()
                ):
                    raise ValueError(
                        "Whole-number and text-length bounds must be integers"
                    )
                if self.validation_type == "text_length" and number < 0:
                    raise ValueError("Text-length bounds must not be negative")
                if self.validation_type == "time" and not 0 <= number <= 1:
                    raise ValueError(
                        "Time bounds must be serial fractions between 0 and 1"
                    )
        if (
            self.validation_type != "list"
            and self.in_cell_dropdown is not None
        ):
            raise ValueError(
                "in_cell_dropdown is only valid for list validation"
            )
        return self


class ReadValidationInput(RangeInput):
    """Inspect exactly one cell to avoid mixed validation results."""

    @model_validator(mode="after")
    def validate_single_cell(self) -> Self:
        """Require exactly one cell to avoid mixed native validation."""
        first_row, first_column, last_row, last_column = parse_address(
            self.address
        )
        if (first_row, first_column) != (last_row, last_column):
            raise ValueError(
                "Read one cell at a time to avoid ambiguous mixed validation"
            )
        return self


def validation_type(target: NativeObject) -> int | None:
    """Confirm missing validation independently of generic native errors."""
    try:
        kind = target.Validation.Type
    except Exception as error:
        info = getattr(error, "excepinfo", None)
        if not (
            (isinstance(error, NativeError) and error.hresult == -2147483645)
            or (
                getattr(error, "hresult", None) == -2147352567
                and isinstance(info, tuple)
                and len(info) > 5
                and info[5] in {-2146827284, -2147024809}
            )
        ):
            raise
        try:
            marked = target.SpecialCells(-4174)
        except Exception as lookup_error:
            lookup_info = getattr(lookup_error, "excepinfo", None)
            if (
                isinstance(lookup_error, NativeError)
                and lookup_error.hresult == -2146827284
            ) or (
                getattr(lookup_error, "hresult", None) == -2147352567
                and isinstance(lookup_info, tuple)
                and len(lookup_info) > 5
                and lookup_info[5] == -2146827284
            ):
                return None
            raise
        # A single-cell SpecialCells query can search the worksheet's used
        # range.
        if (
            marked is None
            or target.Application.Intersect(target, marked) is None
        ):
            return None
        raise ValueError(
            "The range contains validation that cannot be read uniformly; "
            "select a uniform range or clear it explicitly"
        ) from error
    if kind is None:
        raise ValueError(
            "The range has mixed validation; select a uniform "
            "range or clear it explicitly"
        )
    return int(kind)


@tool(
    "wps_et_validation",
    operation="read",
    component="et",
    annotations=READ_ANNOTATIONS,
)
def read_validation(params: ReadValidationInput) -> JsonObject:
    """Read one cell validation rule, prompts, and dropdown settings."""
    sheet, target = resolve_range(params)
    validation = target.Validation
    kind = validation_type(target)
    result: JsonObject = {
        "sheet": str(sheet.Name),
        "address": str(target.Address),
        "present": kind is not None,
    }
    if kind is None:
        return result
    operator = int(validation.Operator) if kind in {1, 2, 4, 5, 6} else None
    result.update(
        validation_type=next(
            (key for key, value in VALIDATION_TYPES.items() if value == kind),
            "other",
        ),
        native_type=kind,
        operator=next(
            (key for key, value in OPERATORS.items() if value == operator),
            None,
        ),
        formula1=str(validation.Formula1) if kind != 0 else None,
        formula2=str(validation.Formula2) if operator in {1, 2} else None,
        alert_style=next(
            (
                key
                for key, value in ALERT_STYLES.items()
                if value == int(validation.AlertStyle)
            ),
            "other",
        ),
        ignore_blank=bool(validation.IgnoreBlank),
        in_cell_dropdown=bool(validation.InCellDropdown)
        if kind == 3
        else None,
        show_input=bool(validation.ShowInput),
        input_title=str(validation.InputTitle),
        input_message=str(validation.InputMessage),
        show_error=bool(validation.ShowError),
        error_title=str(validation.ErrorTitle),
        error_message=str(validation.ErrorMessage),
    )
    return result


@tool(
    "wps_et_validation",
    operation="set",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def set_validation(params: SetValidationInput) -> JsonObject:
    """Add or modify validated cell rules without deleting them first.

    The target must have uniform validation; use native Add or Modify.
    """
    sheet, target = resolve_range(params)
    validation = target.Validation
    current = validation_type(target)
    arguments: JsonObject = {
        "Type": VALIDATION_TYPES[params.validation_type],
        "AlertStyle": ALERT_STYLES[params.alert_style],
        "Operator": OPERATORS[params.operator]
        if params.operator is not None
        else 1,
        "Formula1": params.formula1,
        "Formula2": params.formula2 or "",
    }
    with formula_range(params, target):
        if current is None:
            validation.Add(**arguments)
        else:
            validation.Modify(**arguments)
        for field, property_name in (
            ("ignore_blank", "IgnoreBlank"),
            ("show_input", "ShowInput"),
            ("input_title", "InputTitle"),
            ("input_message", "InputMessage"),
            ("show_error", "ShowError"),
            ("error_title", "ErrorTitle"),
            ("error_message", "ErrorMessage"),
        ):
            setattr(validation, property_name, getattr(params, field))
        if params.validation_type == "list":
            validation.InCellDropdown = (
                True
                if params.in_cell_dropdown is None
                else params.in_cell_dropdown
            )
    return {
        "sheet": str(sheet.Name),
        "address": str(target.Address),
        "set": True,
    }


@tool(
    "wps_et_validation",
    operation="clear",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def clear_validation(params: RangeInput) -> JsonObject:
    """Remove native validation while preserving target cell values."""
    sheet, target = resolve_range(params)
    target.Validation.Delete()
    return {
        "sheet": str(sheet.Name),
        "address": str(target.Address),
        "cleared": True,
    }
