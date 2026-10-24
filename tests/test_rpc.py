"""Exercise Linux binding differences without importing its native modules."""

from collections.abc import Callable
from datetime import date, datetime, time
from types import SimpleNamespace
from typing import Any

import pytest

from src.errors import NativeError
from src.native import rpc
from src.native.rpc import (
    RpcMethod,
    RpcObject,
    check_hresult,
    decode_bstr,
    unwrap,
    wrap,
)


def method(signature: str, function: Callable[..., Any]) -> Callable[..., Any]:
    function.__doc__ = signature
    function.__name__ = signature.split("(", 1)[0]
    return function


@pytest.mark.parametrize("status", [-2147467259, 2147500037, -2147418111])
def test_failed_hresult_preserves_signed_status(status: int) -> None:
    with pytest.raises(NativeError, match="HRESULT") as captured:
        check_hresult((status, None), "Create")
    assert captured.value.hresult == (
        status if status < 0 else status - 0x100000000
    )


def test_status_false_and_outputs_follow_hresult_semantics() -> None:
    assert check_hresult(1, "NoChange") is None
    assert check_hresult((1, False), "Exists") is False
    assert check_hresult((0, 4, 5), "Coordinates") == (4, 5)
    with pytest.raises(NativeError, match="invalid HRESULT"):
        check_hresult(None, "Unknown")


def test_method_results_distinguish_objects_integers_and_status() -> None:
    api = SimpleNamespace()
    native = SimpleNamespace(Name="Document")
    create = RpcMethod(
        method(
            "Add(self) -> (int, wpsapi._Document|None)", lambda: (0, native)
        ),
        api,
    )
    assert create().native is native
    count = RpcMethod(method("getCount(self) -> int", lambda: 23), api)
    assert count() == 23
    statistic = RpcMethod(
        method("ComputeStatistics(self) -> (int, int)", lambda: (0, 42)),
        api,
    )
    assert statistic() == 42
    direct = RpcMethod(
        method("Range(self) -> etapi.IRange|None", lambda: native), api
    )
    assert direct().native is native
    save = RpcMethod(method("Save(self) -> int", lambda: 0), api)
    assert save() is None


def test_native_write_failure_is_never_retried_or_ignored() -> None:
    calls = []

    def save() -> int:
        calls.append("save")
        return -2147467259

    wrapped = RpcMethod(method("Save(self) -> int", save), SimpleNamespace())
    with pytest.raises(NativeError, match="Save failed"):
        wrapped()
    assert calls == ["save"]


def test_missing_native_lcid_is_supplied_without_replaying() -> None:
    calls = []

    def calculate(lcid: int, /) -> int:
        calls.append(lcid)
        return 0

    calculate_method = RpcMethod(
        method("Calculate(self, lcid: int) -> int", calculate),
        SimpleNamespace(),
    )
    assert calculate_method() is None
    assert calculate_method(2052) is None
    assert calls == [0, 2052]


def test_named_required_arguments_become_positional_for_sip() -> None:
    calls = []

    def add(index: int, layout: int, /) -> tuple[int, None]:
        calls.append((index, layout))
        return 0, None

    add_method = RpcMethod(
        method("Add(self, Index: int, Layout: int) -> (int, Any)", add),
        SimpleNamespace(),
    )
    assert add_method(Layout=12, Index=1) is None
    assert calls == [(1, 12)]


def test_native_keywords_follow_com_case_insensitive_contract() -> None:
    calls = []

    def add(name: str, refers_to: str = "") -> int:
        calls.append((name, refers_to))
        return 0

    wrapped = RpcMethod(
        method("Add(self, name: str, refers_to: str = '') -> int", add),
        SimpleNamespace(),
    )
    wrapped(Name="Sales", REFERS_TO="=A1")
    assert calls == [("Sales", "=A1")]
    with pytest.raises(TypeError, match="Duplicate"):
        wrapped(name="A", Name="B")
    with pytest.raises(TypeError, match="Duplicate"):
        wrapped("A", NAME="B")
    ambiguous = RpcMethod(
        method("Add(self, name: str, Name: str) -> int", add),
        SimpleNamespace(),
    )
    with pytest.raises(TypeError, match="Ambiguous"):
        ambiguous(NAME="A")
    assert calls == [("Sales", "=A1")]


@pytest.mark.parametrize(
    "text", ["", "中文", "abc\r", "😀", "A😀中𠀀B", "😀😀😀尾", "尾\0"]
)
@pytest.mark.parametrize("suffix", ["\0", "错", "🧨"])
def test_bstr_recovery_ignores_exact_uninitialized_suffix(
    text: str, suffix: str
) -> None:
    extra = sum(ord(character) > 0xFFFF for character in text)
    native_value = text + suffix * extra
    assert decode_bstr(native_value) == text


def test_variant_matrices_decode_each_native_string_once() -> None:
    matrix = [["😀中?", 7], [False, None]]
    assert wrap(matrix, SimpleNamespace()) == [["😀中", 7], [False, None]]
    assert wrap(None, SimpleNamespace(), "str") == ""
    assert wrap(None, SimpleNamespace(), "Any") is None
    with pytest.raises(NativeError, match="invalid BSTR"):
        decode_bstr("😀")


def test_qstring_extension_results_are_already_correct() -> None:
    getter = RpcMethod(
        method("getDocumentFieldValue(self) -> str", lambda: "😀尾"),
        SimpleNamespace(),
    )
    assert getter() == "😀尾"


@pytest.mark.parametrize("text", ["before\0after", "\ud800", "\udfff"])
def test_unrepresentable_input_is_rejected_before_binding(text: str) -> None:
    calls = []
    writer = RpcMethod(
        method(
            "TypeText(self, Text: str) -> int",
            lambda value: calls.append(value),
        ),
        SimpleNamespace(),
    )
    with pytest.raises(ValueError, match="NUL or isolated surrogates"):
        writer(text)
    assert calls == []


def test_arguments_unwrap_recursively_without_decoding_input_strings() -> None:
    native = SimpleNamespace()
    proxy = RpcObject(native, SimpleNamespace())
    assert unwrap((proxy, ("😀", 1))) == [native, ["😀", 1]]


def test_method_arguments_and_properties_use_native_enum_types() -> None:
    enum_type = type("MsoTriState", (int,), {})
    api = SimpleNamespace(MsoTriState=enum_type)
    calls = []

    def add(visible: int) -> tuple[int, None]:
        calls.append(visible)
        return 0, None

    add_method = RpcMethod(
        method("Add(self, visible: wppapi.MsoTriState) -> (int, Any)", add),
        api,
    )
    add_method(True)
    add_method(visible=0)
    assert [type(value) for value in calls] == [enum_type, enum_type]
    assert calls == [-1, 0]
    native = SimpleNamespace()
    native.setVisible = method(
        "setVisible(self, prop: wppapi.MsoTriState) -> Any",
        lambda value: setattr(native, "Visible", value),
    )
    proxy = RpcObject(native, api)
    proxy.Visible = True
    assert type(native.Visible) is enum_type
    assert native.Visible == -1
    assert wrap(enum_type(-1), api) == -1
    assert type(wrap(enum_type(-1), api)) is int


@pytest.mark.parametrize(
    "value", [date(2026, 10, 24), datetime(2026, 10, 24, 12), time(12, 30)]
)
def test_native_temporal_values_remain_plain_values(value: object) -> None:
    assert wrap(value, SimpleNamespace()) is value


def test_declared_getters_supply_optional_arguments_and_lcid() -> None:
    fields = SimpleNamespace(Count=3)
    calls = []

    def data_fields(index: Any = None) -> tuple[int, Any]:
        calls.append(index)
        return 0, fields

    native = SimpleNamespace(
        get_DataFields=method(
            "get_DataFields(self, Index: Any = None) -> (int, Any)",
            data_fields,
        ),
        get_Count=method(
            "get_Count(self, lcid: int) -> (int, int)", lambda lcid: (0, 3)
        ),
    )
    proxy = RpcObject(native, SimpleNamespace())
    assert proxy.DataFields.native is fields
    assert proxy.Count == 3
    assert calls == [None]


def test_required_getter_parameters_remain_explicit_calls() -> None:
    calls = []

    def information(kind: int) -> tuple[int, int]:
        calls.append(kind)
        return 0, 25

    proxy = RpcObject(
        SimpleNamespace(
            get_Information=method(
                "get_Information(self, Type: int) -> (int, int)", information
            )
        ),
        SimpleNamespace(),
    )
    getter = proxy.Information
    assert calls == []
    assert getter(4) == 25
    assert calls == [4]


def test_optional_getter_keeps_its_parameterized_call_contract() -> None:
    calls = []

    def resize(rows: Any = None, columns: Any = None) -> tuple[int, Any]:
        calls.append((rows, columns))
        return 0, SimpleNamespace(RowCount=rows, ColumnCount=columns)

    proxy = RpcObject(
        SimpleNamespace(
            get_Resize=method(
                "get_Resize(self, RowSize: Any = None, "
                "ColumnSize: Any = None) -> (int, etapi.IRange|None)",
                resize,
            )
        ),
        SimpleNamespace(),
    )
    resized = proxy.Resize(4, 5)
    assert resized.RowCount == 4
    assert resized.ColumnCount == 5
    assert calls == [(None, None), (4, 5)]


def test_range_parent_alias_is_limited_to_declared_et_range_type() -> None:
    range_type = type("IRange", (), {})
    native: Any = range_type()
    sheet = SimpleNamespace(Name="Sheet1")
    native.Worksheet = sheet
    api = SimpleNamespace(IRange=range_type)
    assert RpcObject(native, api).Parent.native is sheet
    with pytest.raises(AttributeError):
        _ = RpcObject(SimpleNamespace(Worksheet=sheet), api).Parent


def test_raw_property_setter_is_called_instead_of_python_assignment() -> None:
    calls = []

    def assign(lcid: int, value: Any, /) -> int:
        calls.append((lcid, value))
        return 0

    native = SimpleNamespace(
        put_Value2=method(
            "put_Value2(self, lcid: int, RHS: Any) -> int", assign
        )
    )
    proxy = RpcObject(native, SimpleNamespace())
    proxy.Value2 = "😀"
    assert calls == [(0, "😀")]
    assert not hasattr(native, "Value2")


def test_failing_raw_property_setter_does_not_create_shadow_attribute() -> (
    None
):
    native = SimpleNamespace(
        put_Name=method(
            "put_Name(self, RHS: str) -> int", lambda value: -2147467259
        )
    )
    proxy = RpcObject(native, SimpleNamespace())
    with pytest.raises(NativeError):
        proxy.Name = "new"
    assert not hasattr(native, "Name")


def test_collection_item_alias_and_multidimensional_calls() -> None:
    items = [SimpleNamespace(Name="one"), SimpleNamespace(Name="two")]
    native = SimpleNamespace(
        Count=2,
        get_Item=method(
            "get_Item(self, Index: Any) -> (int, Any)",
            lambda index: (0, items[index - 1]),
        ),
    )
    proxy = RpcObject(native, SimpleNamespace())
    assert proxy.Item(1).Name == "one"
    assert proxy(2).Name == "two"
    assert proxy[1].Name == "one"
    assert [item.Name for item in proxy] == ["one", "two"]
    assert len(proxy) == 2
    assert bool(RpcObject(SimpleNamespace(), SimpleNamespace()))
    cells = RpcObject(
        SimpleNamespace(
            Item=method(
                "Item(self, RowIndex: Any, ColumnIndex: Any = None) -> Any",
                lambda row, column: SimpleNamespace(Value2=row * 10 + column),
            )
        ),
        SimpleNamespace(),
    )
    assert cells(2, 3).Value2 == 23


def test_address_method_and_empty_bstr_property_match_com_access() -> None:
    native_type = type(
        "Native",
        (),
        {"Name": property(lambda self: None, doc="getName(self) -> str")},
    )
    native: Any = native_type()
    native.Address = method(
        "Address(self, RowAbsolute: Any = None) -> str", lambda: "$A$1:$B$2"
    )
    proxy = RpcObject(native, SimpleNamespace())
    assert proxy.Name == ""
    assert proxy.Address == "$A$1:$B$2"


def test_property_hresult_errors_preserve_native_failure() -> None:
    def fail(instance: Any) -> None:
        raise AttributeError("Call 'get_Text()' failed with 0x80004005")

    native_type = type("Native", (), {"Text": property(fail)})
    with pytest.raises(NativeError) as captured:
        _ = RpcObject(native_type(), SimpleNamespace()).Text
    assert captured.value.hresult == -2147467259


def test_unknown_properties_cannot_be_written_only_to_python() -> None:
    native = SimpleNamespace()
    proxy = RpcObject(native, SimpleNamespace())
    with pytest.raises(AttributeError, match="no writable Missing"):
        proxy.Missing = "not a native write"
    assert not hasattr(native, "Missing")


def matrix_range(
    rows: int, columns: int, fail_at: tuple[int, int] | None = None
) -> tuple[RpcObject, list[tuple[int, int, Any]]]:
    writes = []

    def cell(row: int, column: int) -> Any:
        def assign(instance: Any, value: Any) -> None:
            if (row, column) == fail_at:
                raise AttributeError(
                    "Call 'put_Value2()' failed with 0x80004005"
                )
            writes.append((row, column, value))

        cell_type = type(
            "Cell",
            (),
            {
                "Value2": property(fset=assign),
                "Formula": property(fset=assign),
            },
        )
        return cell_type()

    native = SimpleNamespace(
        Rows=SimpleNamespace(Count=rows),
        Columns=SimpleNamespace(Count=columns),
        Cells=SimpleNamespace(
            Item=method(
                "Item(self, RowIndex: Any, ColumnIndex: Any) -> Any", cell
            )
        ),
    )
    return RpcObject(native, SimpleNamespace()), writes


def test_matrix_writes_validate_all_inputs_before_editing() -> None:
    proxy, writes = matrix_range(2, 2)
    with pytest.raises(ValueError, match="dimensions"):
        proxy.Value2 = [[1, 2]]
    with pytest.raises(ValueError, match="NUL"):
        proxy.Value2 = [[1, 2], [3, "bad\0value"]]
    with pytest.raises(ValueError, match="scalar"):
        proxy.Value2 = [[1, 2], [3, object()]]
    assert writes == []
    proxy.Value2 = (("😀", 7), (False, None))
    assert writes == [(1, 1, "😀"), (1, 2, 7), (2, 1, False), (2, 2, None)]


def test_matrix_failure_reports_partial_write_and_stops() -> None:
    proxy, writes = matrix_range(2, 2, fail_at=(2, 1))
    with pytest.raises(NativeError, match="after 2 of 4 cells") as captured:
        proxy.Value2 = [[1, 2], [3, 4]]
    assert captured.value.hresult == -2147467259
    assert writes == [(1, 1, 1), (1, 2, 2)]


def test_formula_matrix_uses_scalar_formula_setters() -> None:
    proxy, writes = matrix_range(1, 2)
    proxy.Formula = (("=SUM(A1,5)", None),)
    assert writes == [(1, 1, "=SUM(A1,5)"), (1, 2, None)]


@pytest.mark.parametrize(
    ("owner", "name"),
    [
        ("Style", "BaseStyle"),
        ("Style", "NextParagraphStyle"),
        ("Range", "Style"),
        ("Paragraph", "Style"),
    ],
)
def test_style_variant_transfers_one_native_release(
    monkeypatch: pytest.MonkeyPatch, owner: str, name: str
) -> None:
    owners = []
    released = []
    dispatch_type = type("IDispatch", (), {"__module__": "pywpsrpc.common"})
    style_type = type("Style", (), {})
    owner_type = style_type if owner == "Style" else type(owner, (), {})
    raw: Any = dispatch_type()
    raw.owns_reference = True
    native: Any = owner_type()
    setattr(native, name, raw)

    def cast(value: Any, target: type[Any]) -> Any:
        assert value is raw and target is style_type
        typed: Any = target()
        typed.NameLocal = "正文"
        typed.owns_reference = True
        owners.append(typed)
        return typed

    def setdeleted(value: Any) -> None:
        assert value is raw
        value.owns_reference = False

    modules = {
        "pywpsrpc.common": SimpleNamespace(IDispatch=dispatch_type),
        "pywpsrpc.sip": SimpleNamespace(cast=cast, setdeleted=setdeleted),
    }
    monkeypatch.setattr(rpc, "import_module", lambda name: modules[name])
    api = SimpleNamespace(**{owner: owner_type, "Style": style_type})
    result = getattr(RpcObject(native, api), name)
    assert result.NameLocal == "正文"
    for wrapper in [raw, *owners]:
        if wrapper.owns_reference:
            released.append("Release")
    assert released == ["Release"]


def test_defined_name_properties_supply_lcidin() -> None:
    calls = []

    def read_name(locale: int, /) -> tuple[int, str]:
        calls.append(locale)
        return 0, "Sales"

    native = SimpleNamespace(
        get_Name=method("get_Name(self, lcidIn: int) -> (int, str)", read_name)
    )
    assert RpcObject(native, SimpleNamespace()).Name == "Sales"
    assert calls == [0]


def test_chart_series_uses_owned_collection_and_typed_item(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    dispatch_type = type("IDispatch", (), {"__module__": "pywpsrpc.common"})
    chart_type = type("Chart", (), {})
    collection_type = type("SeriesCollection", (), {})
    field = SimpleNamespace(Name="Revenue")
    raw: Any = dispatch_type()
    events = []

    def cast(value: Any, target: type[Any]) -> Any:
        assert value is raw and target is collection_type
        typed: Any = target()
        typed.Count = 1
        typed.Item = method(
            "Item(self, Index: Any) -> (int, wppapi.Series|None)",
            lambda index: (0, field),
        )
        events.append("cast")
        return typed

    modules = {
        "pywpsrpc.common": SimpleNamespace(IDispatch=dispatch_type),
        "pywpsrpc.sip": SimpleNamespace(
            cast=cast, setdeleted=lambda value: events.append("transfer")
        ),
    }
    monkeypatch.setattr(rpc, "import_module", lambda name: modules[name])
    native: Any = chart_type()
    native.SeriesCollection = method(
        "SeriesCollection(self, Index: Any = None) -> (int, IDispatch|None)",
        lambda: (0, raw),
    )
    api = SimpleNamespace(Chart=chart_type, SeriesCollection=collection_type)
    assert RpcObject(native, api).SeriesCollection(1).Name == "Revenue"
    assert events == ["cast", "transfer"]


def test_unrelated_dispatch_variants_are_not_cast(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def unexpected_import(name: str) -> None:
        raise AssertionError(name)

    monkeypatch.setattr(rpc, "import_module", unexpected_import)
    native = SimpleNamespace(BaseStyle=SimpleNamespace())
    assert RpcObject(native, SimpleNamespace()).BaseStyle.native is (
        native.BaseStyle
    )


def test_pivot_field_selection_uses_typed_collection_item() -> None:
    pivot_type = type("IPivotTable", (), {})
    native: Any = pivot_type()
    calls = []
    field = SimpleNamespace(Orientation=1)
    collection = SimpleNamespace(
        Count=1,
        Item=method(
            "Item(self, Index: Any) -> (int, etapi.IPivotField|None)",
            lambda index: (0, field),
        ),
    )

    def fields(index: Any = None) -> tuple[int, Any]:
        calls.append(index)
        return 0, collection

    native.PivotFields = method(
        "PivotFields(self, Index: Any = None) "
        "-> (int, etapi.IPivotFields|None)",
        fields,
    )
    proxy = RpcObject(native, SimpleNamespace(IPivotTable=pivot_type))
    assert proxy.PivotFields().Count == 1
    assert proxy.PivotFields("Region").Orientation == 1
    assert proxy.PivotFields(Index="Region").Orientation == 1
    assert calls == [None, None, None]


@pytest.mark.parametrize("name", ["Value2", "Formula"])
def test_known_safearray_failure_reads_scalar_cells(name: str) -> None:
    def fail(instance: Any) -> None:
        raise AttributeError(f"Call 'get_{name}()' failed with 0x80000003")

    range_type = type("IRange", (), {name: property(fail)})
    native: Any = range_type()
    native.Rows = SimpleNamespace(Count=1)
    native.Columns = SimpleNamespace(Count=2)
    native.Cells = SimpleNamespace(
        Item=method(
            "Item(self, RowIndex: Any, ColumnIndex: Any) -> Any",
            lambda row, column: SimpleNamespace(**{name: column}),
        )
    )
    proxy = RpcObject(native, SimpleNamespace(IRange=range_type))
    assert getattr(proxy, name) == [[1, 2]]
    native.Columns.Count = 1
    with pytest.raises(NativeError) as captured:
        getattr(proxy, name)
    assert captured.value.hresult == -2147483645
