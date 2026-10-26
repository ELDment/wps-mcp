"""Manage conditional formats with verified native replacements."""

from typing import Literal, Self

from pydantic import Field, model_validator

from wps_mcp.common.colors import RgbColor, hex_to_int, int_to_hex
from wps_mcp.errors import WpsError
from wps_mcp.models import JsonObject, NativeObject
from wps_mcp.native.objects import matches_member_error, optional_property
from wps_mcp.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)
from wps_mcp.spreadsheet.access import resolve_range
from wps_mcp.spreadsheet.formulas import (
    OPERATORS,
    Comparison,
    RuleFormula,
    formula_range,
    validate_comparison,
)
from wps_mcp.spreadsheet.models import RangeInput


class ConditionalRuleInput(RangeInput):
    """Define a supported conditional predicate and optional style."""

    rule_type: Literal["cell_value", "expression"]
    operator: Comparison | None = None
    formula1: RuleFormula = Field(min_length=1, max_length=8192)
    formula2: RuleFormula | None = Field(
        default=None, min_length=1, max_length=8192
    )
    font_color: RgbColor | None = None
    fill_color: RgbColor | None = None
    bold: bool | None = None
    italic: bool | None = None
    strikethrough: bool | None = None
    stop_if_true: bool | None = None

    @model_validator(mode="after")
    def validate_rule(self) -> Self:
        """Reject incompatible native rule types, operators, and formulas."""
        if self.rule_type == "expression":
            if not self.formula1.startswith("=") or len(self.formula1) == 1:
                raise ValueError(
                    "Expression formula1 must begin with '=' and "
                    "contain a formula"
                )
            if self.operator is not None or self.formula2 is not None:
                raise ValueError(
                    "Expression rules do not accept operator or formula2"
                )
        else:
            validate_comparison(self.operator, self.formula1, self.formula2)
        return self


class ConditionalIndexInput(RangeInput):
    """Select one conditional rule in its complete target range."""

    index: int = Field(
        ge=1, description="One-based rule index returned by list"
    )


class UpdateConditionalInput(ConditionalRuleInput):
    """Replace one conditional rule while retaining native style."""

    index: int = Field(ge=1)


class ListConditionalInput(RangeInput):
    """Bound conditional-rule enumeration for a rectangle."""

    max_items: int = Field(default=200, ge=1, le=1000)


def describe_condition(rule: NativeObject, index: int) -> JsonObject:
    """Describe a rule without implying unsupported types are editable."""
    kind = int(rule.Type)
    result: JsonObject = {
        "index": index,
        "native_type": kind,
        "rule_type": {1: "cell_value", 2: "expression"}.get(kind, "other"),
        "editable": kind in {1, 2},
        "applies_to": str(rule.AppliesTo.Address),
    }
    if kind not in {1, 2}:
        return result
    operator = int(rule.Operator) if kind == 1 else None
    result.update(
        operator=next(
            (key for key, value in OPERATORS.items() if value == operator),
            None,
        ),
        formula1=str(rule.Formula1),
        formula2=str(rule.Formula2) if operator in {1, 2} else None,
        font_color=int_to_hex(int(rule.Font.Color))
        if rule.Font.Color is not None
        else None,
        fill_color=int_to_hex(int(rule.Interior.Color))
        if rule.Interior.Color is not None
        else None,
        bold=rule.Font.Bold,
        italic=rule.Font.Italic,
        strikethrough=rule.Font.Strikethrough,
        stop_if_true=bool(rule.StopIfTrue),
    )
    return result


def conditional_rule(target: NativeObject, index: int) -> NativeObject:
    """Require a complete native rule range before modifying the rule."""
    if index > int(target.FormatConditions.Count):
        raise ValueError(
            "Conditional format index is out of range; list the rules again"
        )
    rule = target.FormatConditions.Item(index)
    if (
        str(rule.AppliesTo.Address).casefold()
        != str(target.Address).casefold()
    ):
        raise ValueError(
            "Select the rule's complete applies_to range before changing it"
        )
    return rule


def style_condition(rule: NativeObject, params: ConditionalRuleInput) -> None:
    """Apply only explicitly supplied conditional-format properties."""
    if params.font_color is not None:
        rule.Font.Color = hex_to_int(params.font_color)
    if params.fill_color is not None:
        rule.Interior.Color = hex_to_int(params.fill_color)
    for field, property_name in (
        ("bold", "Bold"),
        ("italic", "Italic"),
        ("strikethrough", "Strikethrough"),
    ):
        value = getattr(params, field)
        if value is not None:
            setattr(rule.Font, property_name, value)
    if params.stop_if_true is not None:
        rule.StopIfTrue = params.stop_if_true


def condition_style(rule: NativeObject) -> dict[str, JsonObject]:
    """Snapshot conditional styling before staging a replacement rule."""
    result: dict[str, JsonObject] = {}
    objects = [
        (
            "Font",
            rule.Font,
            ("Bold", "Italic", "Strikethrough", "Underline", "ThemeFont"),
        ),
        ("Interior", rule.Interior, ("Pattern",)),
        *(
            (str(index), rule.Borders.Item(index), ("LineStyle", "Weight"))
            for index in (7, 8, 9, 10)
        ),
    ]
    for name, native, properties in objects:
        attributes = {key: getattr(native, key) for key in properties}
        if name == "Interior" and attributes["Pattern"] in {4000, 4001}:
            raise ValueError(
                "Gradient fills cannot be replaced losslessly; "
                "edit this rule in WPS"
            )
        for prefix in ("", "Pattern") if name == "Interior" else ("",):
            color_index = getattr(native, prefix + "ColorIndex")
            color = getattr(native, prefix + "Color")
            theme = None
            try:
                theme = optional_property(native, prefix + "ThemeColor")
            except Exception as error:
                if not matches_member_error(
                    error,
                    "get",
                    prefix + "ThemeColor",
                    native_code=-2147483645,
                    dispatch_codes={-2146827284, -2147024809},
                ):
                    raise
            if theme is not None and int(theme) > 0:
                attributes[prefix + "ThemeColor"] = theme
            elif color_index is not None and int(color_index) < 0:
                attributes[prefix + "ColorIndex"] = color_index
            elif color is not None:
                attributes[prefix + "Color"] = color
            attributes[prefix + "TintAndShade"] = getattr(
                native, prefix + "TintAndShade"
            )
        # Colors and weights can turn native fills or borders on implicitly.
        final_property = "Pattern" if name == "Interior" else "LineStyle"
        if final_property in attributes:
            attributes[final_property] = attributes.pop(final_property)
        result[name] = {
            key: value
            for key, value in attributes.items()
            if value is not None
        }
    return result


def verify_condition(
    target: NativeObject,
    params: ConditionalRuleInput,
    priority: int,
    count: int,
) -> None:
    """Verify rule predicates and priorities in the actual collection.

    Stale WPS proxies can outlive deleted rules, so read the collection again.
    """
    conditions = target.FormatConditions
    if int(conditions.Count) != count:
        raise WpsError("WPS returned an unexpected conditional-format count")
    for index in range(1, count + 1):
        rule = conditions.Item(index)
        if int(rule.Priority) != priority:
            continue
        if (
            int(rule.Type) == (1 if params.rule_type == "cell_value" else 2)
            and str(rule.Formula1).removeprefix("=")
            == params.formula1.removeprefix("=")
            and (
                params.operator is None
                or int(rule.Operator) == OPERATORS[params.operator]
            )
            and (
                params.formula2 is None
                or str(rule.Formula2).removeprefix("=")
                == params.formula2.removeprefix("=")
            )
            and str(rule.AppliesTo.Address) == str(target.Address)
        ):
            return
        break
    raise WpsError(
        "WPS did not retain the requested "
        "conditional-format predicate and priority"
    )


@tool(
    "wps_et_conditional_formats",
    operation="list",
    component="et",
    annotations=READ_ANNOTATIONS,
)
def list_conditional_formats(params: ListConditionalInput) -> JsonObject:
    """List bounded conditional formats using relative rule indices."""
    sheet, target = resolve_range(params)
    rules = target.FormatConditions
    count = int(rules.Count)
    return {
        "sheet": str(sheet.Name),
        "rules": [
            describe_condition(rules.Item(index), index)
            for index in range(1, min(count, params.max_items) + 1)
        ],
        "count": count,
        "truncated": count > params.max_items,
    }


@tool(
    "wps_et_conditional_formats",
    operation="add",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def add_conditional_format(params: ConditionalRuleInput) -> JsonObject:
    """Append a supported rule while preserving existing formats."""
    sheet, target = resolve_range(params)
    arguments: JsonObject = {
        "Type": 1 if params.rule_type == "cell_value" else 2,
        "Operator": OPERATORS[params.operator]
        if params.operator is not None
        else 3,
        "Formula1": params.formula1,
        "Formula2": params.formula2 or "",
    }
    with formula_range(params, target):
        rule = target.FormatConditions.Add(**arguments)
        style_condition(rule, params)
    return {
        "sheet": str(sheet.Name),
        "address": str(target.Address),
        "added": True,
    }


@tool(
    "wps_et_conditional_formats",
    operation="update",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def update_conditional_format(params: UpdateConditionalInput) -> JsonObject:
    """Verify a styled replacement before deleting the original rule.

    WPS Modify can silently remove a rule. Preparation failures clean up only
    the new rule. Once deletion of the old rule starts, failures retain the
    verified replacement and report partial state instead of discarding the
    only remaining valid rule.
    """
    sheet, target = resolve_range(params)
    rule = conditional_rule(target, params.index)
    if int(rule.Type) not in {1, 2}:
        raise ValueError("Only cell_value and expression rules can be updated")
    original_count = int(target.FormatConditions.Count)
    priority = int(rule.Priority)
    stop_if_true = bool(rule.StopIfTrue)
    number_format = rule.NumberFormat
    preserved = condition_style(rule)
    with formula_range(params, target):
        replacement = None
        try:
            replacement = target.FormatConditions.Add(
                Type=1 if params.rule_type == "cell_value" else 2,
                Operator=OPERATORS[params.operator]
                if params.operator is not None
                else 3,
                Formula1=params.formula1,
                Formula2=params.formula2 or "",
            )
            if int(target.FormatConditions.Count) != original_count + 1:
                raise WpsError(
                    "WPS did not append one replacement conditional format"
                )
            for name, attributes in preserved.items():
                native = (
                    replacement.Borders.Item(int(name))
                    if name.isdigit()
                    else getattr(replacement, name)
                )
                for property_name, value in attributes.items():
                    setattr(native, property_name, value)
            if number_format is not None:
                replacement.NumberFormat = number_format
            replacement.StopIfTrue = stop_if_true
            style_condition(replacement, params)
            replacement.Priority = priority
            verify_condition(target, params, priority, original_count + 1)
        except Exception as error:
            if replacement is not None:
                try:
                    replacement.Delete()
                    if int(target.FormatConditions.Count) != original_count:
                        raise WpsError(
                            "The replacement cleanup did not restore the "
                            "rule count"
                        )
                except Exception as cleanup_error:
                    raise WpsError(
                        "Conditional-format preparation failed and cleanup "
                        "could not be confirmed; "
                        "the old rule was not intentionally deleted. "
                        f"Inspect before retrying: {cleanup_error}"
                    ) from error
            raise
        try:
            rule.Delete()
            verify_condition(target, params, priority, original_count)
        except Exception as error:
            raise WpsError(
                "A verified replacement was created, but old-rule removal "
                "or final verification failed. No replacement cleanup was "
                "attempted; inspect both rules before retrying."
            ) from error
    return {
        "sheet": str(sheet.Name),
        "address": str(target.Address),
        "updated": params.index,
    }


@tool(
    "wps_et_conditional_formats",
    operation="delete",
    component="et",
    annotations=WRITE_ANNOTATIONS,
)
def delete_conditional_format(params: ConditionalIndexInput) -> JsonObject:
    """Delete exactly one rule, requiring its complete applies_to range."""
    sheet, target = resolve_range(params)
    conditional_rule(target, params.index).Delete()
    return {
        "sheet": str(sheet.Name),
        "address": str(target.Address),
        "deleted": params.index,
    }
