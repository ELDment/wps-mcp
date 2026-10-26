"""ET rule contracts reject invalid requests before native mutation."""

from types import SimpleNamespace
from unittest.mock import Mock, PropertyMock

import pytest
from pydantic import ValidationError

from wps_mcp.errors import NativeError, WpsError
from wps_mcp.native.context import CALL_DOCUMENT
from wps_mcp.native.javascript import JsError
from wps_mcp.native.windows import ComSession
from wps_mcp.spreadsheet import conditional, formulas, models, validation
from wps_mcp.spreadsheet import names as defined_names


@pytest.fixture
def target(monkeypatch: pytest.MonkeyPatch) -> Mock:
    current = Mock()
    current.Address = "$A$1:$B$3"
    current.FormatConditions.Count = 2
    current.Validation.Type = 1
    current.Validation.Operator = 1
    current.Validation.Formula1 = "1"
    current.Validation.Formula2 = "100"
    current.Validation.AlertStyle = 1
    current.Validation.IgnoreBlank = True
    current.Validation.ShowInput = True
    current.Validation.InputTitle = "Input"
    current.Validation.InputMessage = "Enter a value"
    current.Validation.ShowError = True
    current.Validation.ErrorTitle = "Invalid"
    current.Validation.ErrorMessage = "Use 1..100"
    current.SpecialCells.return_value = None
    document = Mock()
    document.Name = "Selected.xlsx"
    document.FullName = "D:/Selected.xlsx"
    other = Mock()
    other.Name = "Other.xlsx"
    other.FullName = "D:/Other.xlsx"
    previous_sheet, previous_selection = Mock(), Mock()
    window = Mock(ActiveSheet=previous_sheet, Selection=previous_selection)
    previous_window = Mock()
    application = Mock(
        ActiveWindow=previous_window,
        ActiveWorkbook=other,
        Selection=previous_selection,
    )
    document.Windows.Count = 1
    document.Windows.Item.return_value = window

    def activate_window() -> None:
        application.ActiveWindow = window
        application.ActiveWorkbook = document

    def restore_window() -> None:
        application.ActiveWindow = previous_window
        application.ActiveWorkbook = other

    window.Activate.side_effect = activate_window
    previous_window.Activate.side_effect = restore_window
    current.Parent.Activate.side_effect = lambda: setattr(
        window, "ActiveSheet", current.Parent
    )
    current.Cells.Item.return_value.Select.side_effect = lambda: setattr(
        window, "Selection", current.Cells.Item.return_value
    )
    previous_sheet.Activate.side_effect = lambda: setattr(
        window, "ActiveSheet", previous_sheet
    )
    previous_selection.Select.side_effect = lambda: setattr(
        window, "Selection", previous_selection
    )
    session = ComSession("et")
    session.target_name = document.FullName
    monkeypatch.setattr(
        session, "get_application", lambda allow_launch=False: application
    )
    monkeypatch.setattr(session, "get_documents", lambda: [document, other])
    for module in (formulas, defined_names):
        monkeypatch.setattr(module, "get_session", lambda component: session)
    current.context = SimpleNamespace(
        document=document,
        window=window,
        application=application,
        previous_window=previous_window,
        previous_sheet=previous_sheet,
        previous_selection=previous_selection,
    )
    sheet = SimpleNamespace(Name="Data")
    for module in (conditional, validation):
        monkeypatch.setattr(
            module, "resolve_range", lambda *args: (sheet, current)
        )
    return current


def test_formula_scope_targets_and_restores_the_selected_workbook(
    target: Mock,
) -> None:
    context = target.context
    token = CALL_DOCUMENT.set("previous-call")
    try:
        with formulas.formula_range(
            models.RangeInput(
                document=context.document.FullName, address="B2:B4"
            ),
            target,
        ):
            assert CALL_DOCUMENT.get() == context.document.FullName
            assert context.application.ActiveWindow is context.window
            assert context.window.ActiveSheet is target.Parent
            assert context.window.Selection is target.Cells.Item.return_value
        assert CALL_DOCUMENT.get() == "previous-call"
    finally:
        CALL_DOCUMENT.reset(token)
    assert context.window.ActiveSheet is context.previous_sheet
    assert context.window.Selection is context.previous_selection
    assert context.application.ActiveWindow is context.previous_window
    context.previous_sheet.Activate.assert_called_once()
    context.previous_selection.Select.assert_called_once()
    context.previous_window.Selection.Select.assert_not_called()
    target.Cells.Item.assert_called_with(1, 1)


def test_formula_scope_restores_after_initial_selection_failure(
    target: Mock,
) -> None:
    target.Cells.Item.return_value.Select.side_effect = RuntimeError(
        "selection failed"
    )
    with (
        pytest.raises(RuntimeError, match="selection failed"),
        formulas.formula_range(models.RangeInput(address="B2:B4"), target),
    ):
        pytest.fail("Body must not run after failed selection")
    assert target.context.window.ActiveSheet is target.context.previous_sheet
    assert target.context.window.Selection is target.context.previous_selection
    assert (
        target.context.application.ActiveWindow
        is target.context.previous_window
    )


def test_formula_scope_never_borrows_wrong_workbook_selection(
    target: Mock,
) -> None:
    target.context.window.Activate.side_effect = None
    with (
        pytest.raises(WpsError, match="requested workbook"),
        formulas.formula_range(models.RangeInput(address="B2:B4"), target),
    ):
        pytest.fail("Body must not run in another workbook")
    target.Parent.Activate.assert_not_called()
    target.Cells.Item.return_value.Select.assert_not_called()


def test_formula_scope_restoration_error_is_reported_after_success(
    target: Mock,
) -> None:
    target.context.previous_sheet.Activate.side_effect = RuntimeError(
        "restore failed"
    )
    with pytest.raises(
        WpsError, match="restore the target workbook selection"
    ):
        with formulas.formula_range(
            models.RangeInput(address="B2:B4"), target
        ):
            pass
    target.context.previous_selection.Select.assert_called_once()
    assert (
        target.context.application.ActiveWindow
        is target.context.previous_window
    )


def test_formula_scope_keeps_original_failure_and_logs_restoration_error(
    target: Mock, caplog: pytest.LogCaptureFixture
) -> None:
    target.context.previous_selection.Select.side_effect = RuntimeError(
        "restore failed"
    )
    with pytest.raises(RuntimeError, match="write failed"):
        with formulas.formula_range(
            models.RangeInput(address="B2:B4"), target
        ):
            raise RuntimeError("write failed")
    assert "restore failed" in caplog.text
    assert (
        target.context.application.ActiveWindow
        is target.context.previous_window
    )


@pytest.mark.parametrize(
    "values",
    [
        {"rule_type": "expression", "formula1": "A1>1"},
        {"rule_type": "expression", "formula1": "="},
        {"rule_type": "expression", "formula1": "=A1>1", "operator": "equal"},
        {"rule_type": "expression", "formula1": "=A1>1", "formula2": "2"},
        {"rule_type": "cell_value", "formula1": "1"},
        {"rule_type": "cell_value", "formula1": "1", "operator": "between"},
        {
            "rule_type": "cell_value",
            "formula1": "1",
            "operator": "equal",
            "formula2": "2",
        },
        {"rule_type": "cell_value", "formula1": "\n1", "operator": "equal"},
        {
            "rule_type": "cell_value",
            "formula1": "not a number",
            "operator": "equal",
        },
        {"rule_type": "cell_value", "formula1": "1e999", "operator": "equal"},
        {"rule_type": "cell_value", "formula1": "=  ", "operator": "equal"},
        {"rule_type": "expression", "formula1": "=A1>1", "fill_color": "red"},
    ],
)
def test_conditional_preflight(
    values: dict[str, object], target: Mock
) -> None:
    with pytest.raises(ValidationError):
        conditional.ConditionalRuleInput.model_validate(
            {"address": "A1:B3", **values}
        )
    target.FormatConditions.Add.assert_not_called()
    target.FormatConditions.Delete.assert_not_called()


def test_conditional_add_preserves_other_rules(target: Mock) -> None:
    rule = target.FormatConditions.Add.return_value
    conditional.add_conditional_format(
        conditional.ConditionalRuleInput(
            address="A1:B3",
            rule_type="expression",
            formula1="=A1>10",
            fill_color="#123456",
            bold=False,
            stop_if_true=False,
        )
    )
    target.FormatConditions.Add.assert_called_once_with(
        Type=2, Operator=3, Formula1="=A1>10", Formula2=""
    )
    assert rule.Interior.Color == 0x563412
    assert rule.Font.Bold is False
    assert rule.StopIfTrue is False
    target.FormatConditions.Delete.assert_not_called()


def test_conditional_add_anchors_relative_formula_without_rewriting_it(
    target: Mock,
) -> None:
    def add(**arguments: object) -> Mock:
        assert (
            target.context.application.ActiveWorkbook
            is target.context.document
        )
        assert (
            target.context.window.Selection is target.Cells.Item.return_value
        )
        assert arguments["Formula1"] == "=B2>15"
        return Mock()

    target.FormatConditions.Add.side_effect = add
    conditional.add_conditional_format(
        conditional.ConditionalRuleInput(
            address="B2:B4", rule_type="expression", formula1="=B2>15"
        )
    )
    assert target.context.window.Selection is target.context.previous_selection


def test_validation_set_anchors_relative_formula_without_rewriting_it(
    target: Mock,
) -> None:
    def modify(**arguments: object) -> None:
        assert (
            target.context.application.ActiveWorkbook
            is target.context.document
        )
        assert (
            target.context.window.Selection is target.Cells.Item.return_value
        )
        assert arguments["Formula1"] == "=B2>15"

    target.Validation.Modify.side_effect = modify
    validation.set_validation(
        validation.SetValidationInput(
            address="B2:B4", validation_type="custom", formula1="=B2>15"
        )
    )
    assert target.context.window.Selection is target.context.previous_selection


@pytest.fixture
def conditional_replacement(target: Mock) -> tuple[Mock, Mock]:
    old, new, sibling = Mock(), Mock(), Mock()
    members = [sibling, old]
    priorities = {id(sibling): 1, id(old): 2, id(new): 0}
    target.events = []
    for rule in (old, new, sibling):
        rule.Type = 1
        rule.Operator = 5
        rule.Formula1 = "=15"
        rule.Formula2 = ""
        rule.StopIfTrue = True
        rule.NumberFormat = "0.00"
        rule.AppliesTo.Address = target.Address
        rule.Font = Mock(
            Bold=False,
            Italic=True,
            Strikethrough=False,
            Underline=2,
            ThemeFont=0,
            ColorIndex=1,
            Color=0x123456,
            ThemeColor=0,
            TintAndShade=0,
        )
        rule.Interior = Mock(
            Pattern=1,
            ColorIndex=6,
            Color=0x00FFFF,
            ThemeColor=0,
            TintAndShade=0,
            PatternColorIndex=-4105,
            PatternColor=0,
            PatternThemeColor=0,
            PatternTintAndShade=0,
        )
        rule.Borders.Item.side_effect = {
            index: Mock(
                LineStyle=-4142,
                Weight=2,
                ColorIndex=-4105,
                Color=0,
                ThemeColor=0,
                TintAndShade=0,
            )
            for index in (7, 8, 9, 10)
        }.__getitem__

    def priority(*values: object, rule: Mock) -> int:
        previous = priorities[id(rule)]
        if values:
            desired = int(str(values[0]))
            for member in members:
                current = priorities[id(member)]
                if member is not rule and previous < current <= desired:
                    priorities[id(member)] -= 1
                elif member is not rule and desired <= current < previous:
                    priorities[id(member)] += 1
            priorities[id(rule)] = desired
        return priorities[id(rule)]

    for rule in (old, new, sibling):
        type(rule).Priority = PropertyMock(
            side_effect=lambda *values, rule=rule: priority(*values, rule=rule)
        )
    type(target.FormatConditions).Count = PropertyMock(
        side_effect=lambda: len(members)
    )
    target.FormatConditions.Item.side_effect = lambda index: sorted(
        members, key=lambda member: member.Priority
    )[index - 1]

    def append(**arguments: object) -> Mock:
        target.events.append("add")
        for member in members:
            priorities[id(member)] += 1
        priorities[id(new)] = 1
        for key, value in arguments.items():
            setattr(new, key, value)
        members.append(new)
        return new

    def remove(rule: Mock) -> None:
        target.events.append("delete_old" if rule is old else "delete_new")
        if rule in members:
            members.remove(rule)
            for member in members:
                if member.Priority > priorities[id(rule)]:
                    priorities[id(member)] -= 1

    target.FormatConditions.Add.side_effect = append
    old.Delete.side_effect = lambda: remove(old)
    new.Delete.side_effect = lambda: remove(new)
    return old, new


def test_conditional_linux_rgb_style_has_no_theme(
    conditional_replacement: tuple[Mock, Mock],
) -> None:
    old, _ = conditional_replacement
    type(old.Font).ThemeColor = PropertyMock(
        side_effect=NativeError("not a theme color", -2147483645)
    )
    snapshot = conditional.condition_style(old)
    assert snapshot["Font"]["Color"] == 0x123456
    assert "ThemeColor" not in snapshot["Font"]


@pytest.mark.parametrize("member", ["ThemeColor", "PatternThemeColor"])
def test_conditional_macos_rgb_style_has_no_theme(
    conditional_replacement: tuple[Mock, Mock], member: str
) -> None:
    old, _ = conditional_replacement
    native = old.Font if member == "ThemeColor" else old.Interior
    setattr(
        type(native),
        member,
        PropertyMock(
            side_effect=JsError(
                "get", member, NativeError("Error: Function excute faild")
            )
        ),
    )
    snapshot = conditional.condition_style(old)
    style = snapshot["Font" if member == "ThemeColor" else "Interior"]
    assert member not in style
    if member == "ThemeColor":
        assert style["Color"] == 0x123456
    else:
        assert style["PatternColorIndex"] == -4105


@pytest.mark.parametrize("member", ["ThemeColor", "PatternThemeColor"])
@pytest.mark.parametrize(
    "mismatch", ["operation", "member", "detail", "hresult", "legacy_hresult"]
)
def test_conditional_macos_theme_failure_never_replaces_rule(
    target: Mock,
    conditional_replacement: tuple[Mock, Mock],
    member: str,
    mismatch: str,
) -> None:
    old, new = conditional_replacement
    error = JsError(
        "set" if mismatch == "operation" else "get",
        "Color" if mismatch == "member" else member,
        NativeError(
            "Access denied"
            if mismatch == "detail"
            else "Error: Function excute faild",
            {"hresult": -2147024891, "legacy_hresult": -2147483645}.get(
                mismatch
            ),
        ),
    )
    native = old.Font if member == "ThemeColor" else old.Interior
    setattr(type(native), member, PropertyMock(side_effect=error))
    with pytest.raises(JsError) as caught:
        conditional.update_conditional_format(
            conditional.UpdateConditionalInput(
                address="A1:B3",
                index=2,
                rule_type="expression",
                formula1="=A1>1",
            )
        )
    assert caught.value is error
    target.FormatConditions.Add.assert_not_called()
    old.Delete.assert_not_called()
    new.Delete.assert_not_called()


def test_conditional_update_changes_one_rule(
    target: Mock, conditional_replacement: tuple[Mock, Mock]
) -> None:
    old, new = conditional_replacement
    conditional.update_conditional_format(
        conditional.UpdateConditionalInput(
            address="A1:B3",
            index=2,
            rule_type="expression",
            formula1="=A1>1",
            bold=True,
        )
    )
    target.FormatConditions.Add.assert_called_once_with(
        Type=2, Operator=3, Formula1="=A1>1", Formula2=""
    )
    old.Modify.assert_not_called()
    old.Delete.assert_called_once()
    new.Delete.assert_not_called()
    assert target.events == ["add", "delete_old"]
    assert target.FormatConditions.Count == 2
    assert new.Priority == 2 and new.StopIfTrue is True
    assert new.Font.Bold is True and new.Font.Italic is True
    assert new.Font.Underline == 2 and new.Font.Color == 0x123456
    assert new.Interior.Color == 0x00FFFF and new.Interior.Pattern == 1
    assert new.NumberFormat == "0.00"
    assert new.Borders.Item(7).LineStyle == -4142
    target.FormatConditions.Delete.assert_not_called()


def test_conditional_snapshot_failure_precedes_new_rule(
    target: Mock, conditional_replacement: tuple[Mock, Mock]
) -> None:
    old, new = conditional_replacement
    type(old.Font).ThemeFont = PropertyMock(
        side_effect=RuntimeError("style unavailable")
    )
    with pytest.raises(RuntimeError, match="style unavailable"):
        conditional.update_conditional_format(
            conditional.UpdateConditionalInput(
                address="A1:B3",
                index=2,
                rule_type="expression",
                formula1="=TRUE",
            )
        )
    target.FormatConditions.Add.assert_not_called()
    old.Delete.assert_not_called()
    new.Delete.assert_not_called()


def test_conditional_staging_failure_cleans_new_and_preserves_old(
    target: Mock, conditional_replacement: tuple[Mock, Mock]
) -> None:
    old, new = conditional_replacement
    type(new.Font).Bold = PropertyMock(
        side_effect=RuntimeError("style write failed")
    )
    with pytest.raises(RuntimeError, match="style write failed"):
        conditional.update_conditional_format(
            conditional.UpdateConditionalInput(
                address="A1:B3",
                index=2,
                rule_type="expression",
                formula1="=TRUE",
            )
        )
    old.Delete.assert_not_called()
    new.Delete.assert_called_once()
    assert target.events == ["add", "delete_new"]
    assert target.FormatConditions.Count == 2
    assert old.Priority == 2


def test_conditional_predicate_is_verified_in_collection(
    target: Mock, conditional_replacement: tuple[Mock, Mock]
) -> None:
    old, new = conditional_replacement
    type(new).Formula1 = PropertyMock(return_value="=WRONG")
    with pytest.raises(WpsError, match="did not retain"):
        conditional.update_conditional_format(
            conditional.UpdateConditionalInput(
                address="A1:B3",
                index=2,
                rule_type="expression",
                formula1="=TRUE",
            )
        )
    old.Delete.assert_not_called()
    new.Delete.assert_called_once()
    assert target.FormatConditions.Count == 2


def test_conditional_old_delete_failure_retains_verified_replacement(
    target: Mock, conditional_replacement: tuple[Mock, Mock]
) -> None:
    old, new = conditional_replacement
    old.Delete.side_effect = RuntimeError("delete failed")
    with pytest.raises(WpsError, match="No replacement cleanup"):
        conditional.update_conditional_format(
            conditional.UpdateConditionalInput(
                address="A1:B3",
                index=2,
                rule_type="expression",
                formula1="=TRUE",
            )
        )
    old.Delete.assert_called_once()
    new.Delete.assert_not_called()
    assert target.FormatConditions.Count == 3


def test_conditional_cleanup_failure_is_explicit(
    target: Mock, conditional_replacement: tuple[Mock, Mock]
) -> None:
    old, new = conditional_replacement
    type(new.Font).Bold = PropertyMock(
        side_effect=RuntimeError("style write failed")
    )
    new.Delete.side_effect = RuntimeError("cleanup failed")
    with pytest.raises(WpsError, match="cleanup could not be confirmed"):
        conditional.update_conditional_format(
            conditional.UpdateConditionalInput(
                address="A1:B3",
                index=2,
                rule_type="expression",
                formula1="=TRUE",
            )
        )
    old.Delete.assert_not_called()
    assert target.FormatConditions.Count == 3


def test_conditional_update_rejects_partial_rule_target(target: Mock) -> None:
    rule = target.FormatConditions.Item.return_value
    rule.AppliesTo.Address = "$A$1:$B$100"
    with pytest.raises(ValueError, match="complete"):
        conditional.delete_conditional_format(
            conditional.ConditionalIndexInput(address="A1:B3", index=1)
        )
    rule.Delete.assert_not_called()


def test_conditional_index_and_unsupported_type_are_guarded(
    target: Mock,
) -> None:
    with pytest.raises(ValueError, match="index"):
        conditional.delete_conditional_format(
            conditional.ConditionalIndexInput(address="A1:B3", index=3)
        )
    rule = target.FormatConditions.Item.return_value
    rule.Type = 3
    rule.AppliesTo.Address = target.Address
    with pytest.raises(ValueError, match="Only cell_value"):
        conditional.update_conditional_format(
            conditional.UpdateConditionalInput(
                address="A1:B3",
                index=1,
                rule_type="expression",
                formula1="=TRUE",
            )
        )
    rule.Modify.assert_not_called()


def test_conditional_list_handles_newer_native_types(target: Mock) -> None:
    rule = target.FormatConditions.Item.return_value
    rule.Type = 3
    rule.AppliesTo.Address = target.Address
    result = conditional.list_conditional_formats(
        conditional.ListConditionalInput(address="A1:B3", max_items=1)
    )
    assert result["truncated"] is True
    assert result["rules"][0]["native_type"] == 3
    assert result["rules"][0]["editable"] is False


def test_conditional_native_failure_is_not_replayed(target: Mock) -> None:
    target.FormatConditions.Add.side_effect = RuntimeError(
        "native write failed"
    )
    with pytest.raises(RuntimeError, match="native write failed"):
        conditional.add_conditional_format(
            conditional.ConditionalRuleInput(
                address="A1:B3", rule_type="expression", formula1="=TRUE"
            )
        )
    assert target.FormatConditions.Add.call_count == 1
    target.FormatConditions.Delete.assert_not_called()


@pytest.mark.parametrize(
    "values",
    [
        {"validation_type": "whole", "formula1": "1"},
        {"validation_type": "decimal", "formula1": "1", "operator": "between"},
        {"validation_type": "custom", "formula1": "A1>1"},
        {
            "validation_type": "custom",
            "formula1": "=A1>1",
            "operator": "equal",
        },
        {"validation_type": "list", "formula1": "a,b", "formula2": "c,d"},
        {"validation_type": "list", "formula1": "="},
        {"validation_type": "list", "formula1": "a" * 256},
        {
            "validation_type": "date",
            "formula1": "1",
            "operator": "equal",
            "in_cell_dropdown": True,
        },
        {
            "validation_type": "list",
            "formula1": "a,b",
            "error_title": "x" * 33,
        },
        {"validation_type": "whole", "formula1": "1.5", "operator": "equal"},
        {
            "validation_type": "text_length",
            "formula1": "-1",
            "operator": "equal",
        },
        {"validation_type": "time", "formula1": "2", "operator": "equal"},
        {
            "validation_type": "decimal",
            "formula1": "Infinity",
            "operator": "equal",
        },
        {"validation_type": "custom", "formula1": "=  "},
    ],
)
def test_validation_preflight(values: dict[str, object], target: Mock) -> None:
    with pytest.raises(ValidationError):
        validation.SetValidationInput.model_validate(
            {"address": "A1:B3", **values}
        )
    target.Validation.Delete.assert_not_called()
    target.Validation.Add.assert_not_called()
    target.Validation.Modify.assert_not_called()


@pytest.mark.parametrize(
    "kind", ["whole", "decimal", "date", "time", "text_length"]
)
def test_validation_comparison_types(kind: str, target: Mock) -> None:
    params = validation.SetValidationInput.model_validate(
        {
            "address": "A1:B3",
            "validation_type": kind,
            "operator": "between",
            "formula1": "1",
            "formula2": "1" if kind == "time" else "10",
        }
    )
    validation.set_validation(params)
    target.Validation.Modify.assert_called_once_with(
        Type=validation.VALIDATION_TYPES[kind],
        AlertStyle=1,
        Operator=1,
        Formula1="1",
        Formula2="1" if kind == "time" else "10",
    )
    target.Validation.Delete.assert_not_called()


def test_validation_list_prompts_and_false_settings(target: Mock) -> None:
    validation.set_validation(
        validation.SetValidationInput(
            address="A1:B3",
            validation_type="list",
            formula1="Open,Closed",
            in_cell_dropdown=False,
            ignore_blank=False,
            show_input=False,
            show_error=False,
            input_title="Status",
            error_message="Choose a listed value",
        )
    )
    assert target.Validation.InCellDropdown is False
    assert target.Validation.IgnoreBlank is False
    assert target.Validation.ShowInput is False
    assert target.Validation.ShowError is False
    assert target.Validation.InputTitle == "Status"
    assert target.Validation.ErrorMessage == "Choose a listed value"
    target.Validation.Modify.assert_called_once_with(
        Type=3, AlertStyle=1, Operator=1, Formula1="Open,Closed", Formula2=""
    )


class MissingValidationError(Exception):
    hresult = -2147352567
    excepinfo = (0, "ET", "No validation", None, 0, -2146827284)


class WpsMissingValidationError(Exception):
    hresult = -2147352567
    excepinfo = (0, None, None, None, 0, -2147024809)


@pytest.mark.parametrize(
    "error",
    [
        WpsMissingValidationError(),
        NativeError("No type", -2147483645),
        JsError("get", "Type", NativeError("Error: Function excute faild")),
    ],
)
def test_wps_empty_validation_requires_independent_absence_check(
    target: Mock,
    error: Exception,
) -> None:
    type(target.Validation).Type = PropertyMock(side_effect=error)
    assert (
        validation.read_validation(
            validation.ReadValidationInput(address="A1")
        )["present"]
        is False
    )
    target.SpecialCells.assert_called_once_with(-4174)


@pytest.mark.parametrize(
    "error",
    [
        MissingValidationError(),
        NativeError("No cells", -2146827284),
        JsError(
            "call", "SpecialCells", NativeError("Error: Function excute faild")
        ),
    ],
)
def test_missing_validation_accepts_native_no_special_cells(
    target: Mock,
    error: Exception,
) -> None:
    type(target.Validation).Type = PropertyMock(
        side_effect=WpsMissingValidationError()
    )
    target.SpecialCells.side_effect = error
    assert (
        validation.read_validation(
            validation.ReadValidationInput(address="A1")
        )["present"]
        is False
    )


def test_missing_validation_ignores_special_cells_outside_target(
    target: Mock,
) -> None:
    type(target.Validation).Type = PropertyMock(
        side_effect=WpsMissingValidationError()
    )
    marked = Mock()
    target.SpecialCells.return_value = marked
    target.Application.Intersect.return_value = None
    assert (
        validation.read_validation(
            validation.ReadValidationInput(address="A1")
        )["present"]
        is False
    )
    target.Application.Intersect.assert_called_once_with(target, marked)


@pytest.mark.parametrize(
    "error",
    [
        WpsMissingValidationError(),
        NativeError("No type", -2147483645),
        JsError("get", "Type", NativeError("Error: Function excute faild")),
    ],
)
def test_invalid_native_type_with_existing_rules_does_not_replace_them(
    target: Mock,
    error: Exception,
) -> None:
    type(target.Validation).Type = PropertyMock(side_effect=error)
    target.SpecialCells.return_value = Mock()
    target.Application.Intersect.return_value = Mock()
    with pytest.raises(ValueError, match="cannot be read uniformly"):
        validation.set_validation(
            validation.SetValidationInput(
                address="A1:B3", validation_type="list", formula1="a,b"
            )
        )
    target.Validation.Add.assert_not_called()
    target.Validation.Modify.assert_not_called()
    target.Validation.Delete.assert_not_called()


@pytest.mark.parametrize(
    "error",
    [
        RuntimeError("RPC failure"),
        WpsMissingValidationError(),
        NativeError("Disconnected", -2147467259),
    ],
)
def test_special_cells_failure_is_not_misreported_as_absence(
    target: Mock, error: Exception
) -> None:
    type(target.Validation).Type = PropertyMock(
        side_effect=WpsMissingValidationError()
    )
    target.SpecialCells.side_effect = error
    with pytest.raises(type(error)):
        validation.read_validation(
            validation.ReadValidationInput(address="A1")
        )


@pytest.mark.parametrize(
    "operation,member", [("get", "Type"), ("call", "SpecialCells")]
)
@pytest.mark.parametrize(
    "mismatch", ["operation", "member", "detail", "hresult", "legacy_hresult"]
)
def test_validation_macos_unrelated_failure_never_adds_rule(
    target: Mock, operation: str, member: str, mismatch: str
) -> None:
    legacy = -2147483645 if member == "Type" else -2146827284
    error = JsError(
        "set" if mismatch == "operation" else operation,
        "Count" if mismatch == "member" else member,
        NativeError(
            "Access denied"
            if mismatch == "detail"
            else "Error: Function excute faild",
            {"hresult": -2147024891, "legacy_hresult": legacy}.get(mismatch),
        ),
    )
    type(target.Validation).Type = PropertyMock(
        side_effect=error
        if member == "Type"
        else JsError(
            "get", "Type", NativeError("Error: Function excute faild")
        )
    )
    target.SpecialCells.side_effect = error
    with pytest.raises(JsError) as caught:
        validation.set_validation(
            validation.SetValidationInput(
                address="A1:B3", validation_type="list", formula1="a,b"
            )
        )
    assert caught.value is error
    target.Validation.Add.assert_not_called()
    target.Validation.Modify.assert_not_called()
    target.Validation.Delete.assert_not_called()


def test_validation_absent_adds_without_delete(target: Mock) -> None:
    type(target.Validation).Type = PropertyMock(
        side_effect=MissingValidationError()
    )
    validation.set_validation(
        validation.SetValidationInput(
            address="A1:B3", validation_type="custom", formula1="=A1>0"
        )
    )
    target.Validation.Add.assert_called_once_with(
        Type=7, AlertStyle=1, Operator=1, Formula1="=A1>0", Formula2=""
    )
    target.Validation.Modify.assert_not_called()
    target.Validation.Delete.assert_not_called()


def test_validation_read_absent_and_unrelated_native_errors(
    target: Mock,
) -> None:
    type(target.Validation).Type = PropertyMock(
        side_effect=MissingValidationError()
    )
    assert (
        validation.read_validation(
            validation.ReadValidationInput(address="A1")
        )["present"]
        is False
    )
    type(target.Validation).Type = PropertyMock(
        side_effect=RuntimeError("connection lost")
    )
    with pytest.raises(RuntimeError, match="connection lost"):
        validation.read_validation(
            validation.ReadValidationInput(address="A1")
        )


def test_validation_read_and_mixed_ranges(target: Mock) -> None:
    result = validation.read_validation(
        validation.ReadValidationInput(address="A1")
    )
    assert result["validation_type"] == "whole"
    assert result["formula2"] == "100"
    assert result["operator"] == "between"
    with pytest.raises(ValidationError, match="one cell"):
        validation.ReadValidationInput(address="A1:B3")
    target.Validation.Type = None
    with pytest.raises(ValueError, match="mixed"):
        validation.set_validation(
            validation.SetValidationInput(
                address="A1:B3", validation_type="list", formula1="a,b"
            )
        )
    target.Validation.Modify.assert_not_called()


def test_validation_modify_failure_preserves_rule_without_replay(
    target: Mock,
) -> None:
    target.Validation.Modify.side_effect = RuntimeError("unsupported modify")
    with pytest.raises(RuntimeError, match="unsupported modify"):
        validation.set_validation(
            validation.SetValidationInput(
                address="A1:B3", validation_type="list", formula1="a,b"
            )
        )
    target.Validation.Modify.assert_called_once()
    target.Validation.Delete.assert_not_called()
    target.Validation.Add.assert_not_called()


@pytest.fixture
def workbook(monkeypatch: pytest.MonkeyPatch) -> Mock:
    book = Mock()
    book.Name = "selected.xlsx"
    sheet = Mock()
    sheet.Name = "Data's sheet"
    sheet.Rows.Count = 1048576
    sheet.Columns.Count = 16384
    sheet.Names.Count = 0
    sheet.Range.return_value.Address = "$A$1:$B$3"
    sheet.Range.return_value.Rows.Count = 3
    sheet.Range.return_value.Columns.Count = 2
    sheet.Range.return_value.Value2 = (("A", 1), ("B", 2), ("C", 3))
    book.Worksheets.Count = 1
    book.Worksheets.Item.return_value = sheet
    book.Names.Count = 0
    session = Mock()
    session.get_document.return_value = book
    for module in (formulas, defined_names):
        monkeypatch.setattr(module, "get_session", lambda component: session)
    return book


@pytest.mark.parametrize(
    "name",
    [
        "A1",
        "R1C1",
        "R",
        "c",
        "Sheet!Range",
        "[book]Area",
        "=NOW()",
        "_xlnm.Print_Area",
    ],
)
def test_name_rejects_reference_formula_and_reserved_names(name: str) -> None:
    with pytest.raises(ValidationError):
        defined_names.NameInput(name=name)


@pytest.mark.parametrize(
    "values",
    [
        {"scope": "sheet"},
        {"scope": "workbook", "sheet": 1},
        {"scope": "sheet", "sheet": False},
    ],
)
def test_name_scope_is_explicit(values: dict[str, object]) -> None:
    with pytest.raises(ValidationError):
        defined_names.NameScopeInput.model_validate(values)


def make_name(
    name: str, reference: str = "='Data''s sheet'!$A$1:$B$3"
) -> Mock:
    item = Mock()
    item.Name = name
    item.RefersTo = reference
    item.Visible = True
    return item


def test_names_distinguish_global_local_and_formula(workbook: Mock) -> None:
    names = [
        make_name("Area"),
        make_name("'Data''s sheet'!Area"),
        make_name("Computed", "=SUM(A1:A3)"),
    ]
    workbook.Names.Count = 3
    workbook.Names.Item.side_effect = lambda index: names[index - 1]
    result = defined_names.list_names(defined_names.ListNamesInput())
    assert [name["name"] for name in result["names"]] == ["Area", "Computed"]
    assert result["names"][1]["kind"] == "formula_or_external"
    for name in names:
        assert "RefersToRange" not in name.__dict__


def test_set_name_constructs_quoted_absolute_local_reference(
    workbook: Mock,
) -> None:
    workbook.Names.Add.return_value = make_name("Area")
    defined_names.set_name(
        defined_names.SetNameInput(
            document="selected.xlsx",
            name="Area",
            source_sheet=1,
            address="A1:B3",
        )
    )
    workbook.Names.Add.assert_called_once_with(
        Name="Area", RefersTo="='Data''s sheet'!$A$1:$B$3", Visible=True
    )


def test_set_sheet_name_uses_sheet_collection(workbook: Mock) -> None:
    sheet = workbook.Worksheets.Item.return_value
    sheet.Names.Add.return_value = make_name("'Data''s sheet'!Area")
    defined_names.set_name(
        defined_names.SetNameInput(
            name="Area",
            scope="sheet",
            sheet=1,
            source_sheet=1,
            address="A1:B3",
        )
    )
    sheet.Names.Add.assert_called_once_with(
        Name="'Data''s sheet'!Area",
        RefersTo="='Data''s sheet'!$A$1:$B$3",
        Visible=True,
    )
    workbook.Names.Add.assert_not_called()


@pytest.mark.parametrize(
    "reference",
    [
        "=NOW()",
        "='[foreign.xlsx]Data'!$A$1",
        "='C:\\Data'!$A$1",
        "=Data!A1:A1048577",
    ],
)
def test_formula_and_external_names_cannot_be_read_changed_or_deleted(
    workbook: Mock, reference: str
) -> None:
    name = make_name("Area", reference)
    workbook.Names.Count = 1
    workbook.Names.Item.return_value = name
    with pytest.raises(ValueError, match="formula or external"):
        defined_names.set_name(
            defined_names.SetNameInput(
                name="Area", source_sheet=1, address="A1:B3"
            )
        )
    with pytest.raises(ValueError, match="formula or external"):
        defined_names.read_name(defined_names.ReadNameInput(name="Area"))
    with pytest.raises(ValueError, match="Only local"):
        defined_names.delete_name(defined_names.NameInput(name="Area"))
    assert name.RefersTo == reference
    name.Delete.assert_not_called()


def test_read_name_resolves_bound_values_in_selected_workbook(
    workbook: Mock,
) -> None:
    workbook.Names.Count = 1
    workbook.Names.Item.return_value = make_name("Area")
    result = defined_names.read_name(
        defined_names.ReadNameInput(name="Area", include_values=True)
    )
    assert result["values"] == [["A", 1], ["B", 2], ["C", 3]]
    with pytest.raises(ValueError, match="cell limit"):
        defined_names.read_name(
            defined_names.ReadNameInput(
                name="Area", include_values=True, max_cells=2
            )
        )


def test_delete_sheet_name_does_not_delete_workbook_name(
    workbook: Mock,
) -> None:
    global_name = make_name("Area")
    local_name = make_name("'Data''s sheet'!Area")
    workbook.Names.Count = 1
    workbook.Names.Item.return_value = global_name
    sheet = workbook.Worksheets.Item.return_value
    sheet.Names.Count = 1
    sheet.Names.Item.return_value = local_name
    defined_names.delete_name(
        defined_names.NameInput(name="Area", scope="sheet", sheet=1)
    )
    local_name.Delete.assert_called_once()
    global_name.Delete.assert_not_called()
