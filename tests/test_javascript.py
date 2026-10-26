"""Check native reference ownership and JavaScript argument boundaries."""

import json
import subprocess
from collections.abc import Callable, Iterator
from concurrent.futures import ThreadPoolExecutor
from datetime import UTC, date, datetime
from shutil import which
from threading import Event, get_ident
from types import SimpleNamespace
from typing import Any, cast
from unittest.mock import Mock

import pytest

from src.errors import NativeError, WpsError
from src.models import Component
from src.native import runtime
from src.native.context import CALL_DOCUMENT
from src.native.javascript import (
    BRIDGE,
    JsError,
    JsObject,
    ObjectBridge,
    encode,
)
from src.native.macos import MacDriver, MacExecutor, MacSession
from src.native.objects import optional_property, read_collection


@pytest.fixture
def bridge() -> ObjectBridge:
    driver = Mock(spec=MacDriver)
    driver.execute.return_value = {"result": None}
    return ObjectBridge(driver, "writer", "operation-one")


def sent(bridge: ObjectBridge) -> tuple[str, dict[str, Any]]:
    _, command, data = cast(Mock, bridge.driver.execute).call_args.args[2]
    return command, data


@pytest.fixture
def native_engine() -> Iterator[
    tuple[ObjectBridge, Callable[[dict[str, Any]], dict[str, Any]]]
]:
    node = which("node")
    if node is None:
        pytest.skip("Node.js is needed to execute the isolated bridge script")
    program = """
const lines=require('readline').createInterface({input:process.stdin});
const documents=[], stats={adds:0,opens:0,saves:0,writes:0,
    mutations:0,calls:[]};
let expire=false;
function createDocument() {
    const doc=Object.assign(function(){}, {
        FullName:'Untitled'+(documents.length+1), Saved:false,
        Close(){},
        Range(start,end){return {Start:start,End:end,Text:'range'};},
        SaveAs(name){stats.saves++;doc.FullName=name;expire=true;},
        MakeObject(){stats.mutations++;expire=true;return {Name:'ephemeral'};}
    });
    doc.SaveAs2=doc.SaveAs;
    Object.defineProperty(doc,'Title',{get(){return 'title';},
        set(value){stats.writes++;}});
    documents.push(doc);expire=true;return doc;
}
const collection={get Count(){return documents.length;},
    Item(index){return documents[index-1];},
    Add(){stats.adds++;return createDocument();},
    Open(name){stats.opens++;const doc=createDocument();doc.FullName=name;
        return doc;}};
const styles=[{NameLocal:'Normal',Type:1,BuiltIn:true,InUse:false},
    {NameLocal:'Emphasis',Type:2,BuiltIn:false,InUse:true}];
function opaque(value) {
    return new Proxy(value,{
        has(target,name) {
            if(['FullName','Saved','Close'].includes(name))
                throw new Error('Unsafe native member probe: '+name);
            return Reflect.has(target,name);
        },
        get(target,name,receiver) {
            if(name==='FullName') throw new Error('Not a document');
            return Reflect.get(target,name,receiver);
        }
    });
}
const opaqueValue=opaque(Object.assign(function(){},{Name:'opaque'}));
const opaqueMethod=opaque(function(){return opaqueValue;});
const chartSeries={Count:2,Item(index){return {
    XValues:[['Q1',undefined],['Q2',true,null]],Values:[[index,index+1],[0,false]],
    NativeObject:[opaqueValue],NativeFunction:[opaqueMethod],PlainObject:[{}],
    get Name(){
        throw new Error('Unsafe getter must not be queried');
    }
}}};
const app={Documents:collection,Workbooks:collection,Presentations:collection,
    Styles:{Count:styles.length,Item(index){return styles[index-1];}},
    Series:chartSeries,
    Unrelated:{Documents:{Add:opaqueMethod,Open:opaqueMethod,Item:opaqueMethod},
        Font:opaqueValue,Modify:opaqueMethod,Workbook:opaqueValue}};
Object.defineProperties(app,{
    ActiveDocument:{get(){return documents[0];}},
    ActiveWorkbook:{get(){return documents[0];}},
    ActivePresentation:{get(){return documents[0];}}
});
lines.on('line',line=>{
    const input=JSON.parse(line);
    if(input.stats) return console.log(JSON.stringify(stats));
    if(input.expire) {delete globalThis.scope;return console.log('{}');}
    const bridge=new Function('app','args',input.script);
    stats.calls.push(input.args[1]);
    const result=bridge(app,input.args);
    console.log(JSON.stringify({result,socket_path:'/tmp/wpsdriverlocal_wps_1'}));
    if(expire) {delete globalThis.scope;expire=false;}
});
"""
    process = subprocess.Popen(
        [node, "-e", program],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )

    def exchange(data: dict[str, Any]) -> dict[str, Any]:
        assert process.stdin is not None and process.stdout is not None
        process.stdin.write(json.dumps(data) + "\n")
        process.stdin.flush()
        result = process.stdout.readline()
        if not result:
            assert process.stderr is not None
            pytest.fail(process.stderr.read())
        return json.loads(result)

    def execute(
        component: Component,
        script: str,
        args: list[Any],
        *,
        socket_path: str | None = None,
    ) -> dict[str, Any]:
        return exchange({"script": script, "args": args})

    driver = Mock(spec=MacDriver)
    driver.execute.side_effect = execute
    try:
        yield ObjectBridge(driver, "writer", "scope"), exchange
    finally:
        process.kill()
        process.communicate(timeout=5)


@pytest.mark.parametrize("method", ["Add", "Open"])
def test_context_recovery_locates_created_document_without_replaying(
    native_engine: Any, method: str
) -> None:
    bridge, exchange = native_engine
    application = bridge.request("begin")
    collection = application.Documents
    create = getattr(collection, method)
    document = create() if method == "Add" else create("opened.docx")
    expected = "Untitled1" if method == "Add" else "opened.docx"
    assert document.FullName == expected
    assert bridge._recipes[document._id] == {
        "kind": "document",
        "full_name": expected,
    }
    text_range = document.Range(0, 5)
    first_id = text_range._id
    exchange({"expire": True})
    assert text_range.Text == "range"
    assert document.Range(0, 1)._id > first_id
    stats = exchange({"stats": True})
    assert stats["adds"] + stats["opens"] == 1
    assert stats["calls"].count("restore") == 2


def test_recovery_does_not_replay_unknown_mutating_reference(
    native_engine: Any,
) -> None:
    bridge, exchange = native_engine
    application = bridge.request("begin")
    document = application.Documents.Add()
    temporary = document.MakeObject()
    with pytest.raises(JsError, match="cannot be restored safely"):
        _ = temporary.Name
    stats = exchange({"stats": True})
    assert stats["adds"] == 1
    assert stats["mutations"] == 1


def test_generic_objects_and_methods_never_probe_document_members(
    native_engine: Any,
) -> None:
    bridge, _ = native_engine
    unrelated = bridge.request("begin").Unrelated
    assert unrelated.Font.Name == "opaque"
    assert unrelated.Modify().Name == "opaque"
    assert unrelated.Documents.Add().Name == "opaque"
    assert unrelated.Documents.Open().Name == "opaque"
    assert unrelated.Documents.Item(1).Name == "opaque"
    assert unrelated.Documents[1].Name == "opaque"
    assert unrelated.Documents(1).Name == "opaque"
    # Workbook belongs to ET, so Writer must retain its ordinary getter path.
    workbook = unrelated.Workbook
    assert workbook.Name == "opaque"
    assert bridge._recipes[workbook._id]["kind"] == "get"


@pytest.mark.parametrize(
    ("component", "collection", "getter"),
    [
        ("writer", "Documents", "ActiveDocument"),
        ("et", "Workbooks", "ActiveWorkbook"),
        ("wpp", "Presentations", "ActivePresentation"),
    ],
)
def test_component_binding_tags_only_known_document_sources(
    native_engine: Any, component: Component, collection: str, getter: str
) -> None:
    bridge, exchange = native_engine
    bridge.component = component
    application = bridge.request("begin")
    documents = getattr(application, collection)
    document = documents.Add()
    for reference in (
        documents.Item(1),
        documents[1],
        documents(1),
        documents()(1),
        getattr(application, getter),
    ):
        assert bridge._recipes[reference._id] == {
            "kind": "document",
            "full_name": "Untitled1",
        }
    exchange({"expire": True})
    assert document.FullName == "Untitled1"
    assert exchange({"stats": True})["adds"] == 1


@pytest.mark.parametrize("lookup", ["index", "default", "method"])
def test_read_only_collection_paths_recover_lazily(
    native_engine: Any, lookup: str
) -> None:
    bridge, exchange = native_engine
    styles = bridge.request("begin").Styles
    style = (
        styles[1]
        if lookup == "index"
        else styles()[1]
        if lookup == "default"
        else styles.Item(1)
    )
    exchange({"expire": True})
    assert style.NameLocal == "Normal"
    assert exchange({"stats": True})["calls"].count("restore") == 1


def test_ambiguous_document_recovery_rejects_current_write(
    native_engine: Any,
) -> None:
    bridge, exchange = native_engine
    application = bridge.request("begin")
    first = application.Documents.Add()
    second = application.Documents.Add()
    second.SaveAs("Untitled1")
    with pytest.raises(JsError, match="uniquely restored"):
        first.Title = "must not be written"
    assert exchange({"stats": True})["writes"] == 0


@pytest.mark.parametrize("method", ["SaveAs", "SaveAs2"])
def test_save_as_refreshes_all_document_locators_before_context_recovery(
    native_engine: Any, method: str
) -> None:
    bridge, exchange = native_engine
    application = bridge.request("begin")
    first = application.Documents.Add()
    alias = application.Documents.Item(1)
    getattr(first, method)("renamed.docx")
    for reference in (first, alias):
        assert bridge._recipes[reference._id]["full_name"] == "renamed.docx"
        assert reference.FullName == "renamed.docx"
    assert exchange({"stats": True})["saves"] == 1


def test_expired_scope_restores_before_executing_current_write_once(
    native_engine: Any,
) -> None:
    bridge, exchange = native_engine
    document = bridge.request("begin").Documents.Add()
    document.Title = "new title"
    exchange({"expire": True})
    bridge.request("end")
    stats = exchange({"stats": True})
    assert stats["writes"] == 1
    assert stats["calls"].count("restore") == 1
    assert stats["calls"][-1] == "end"


def test_transport_failure_never_triggers_reference_recovery(
    bridge: ObjectBridge,
) -> None:
    execute = cast(Mock, bridge.driver.execute)
    execute.side_effect = WpsError("Unconfirmed native response")
    with pytest.raises(WpsError, match="Unconfirmed"):
        JsObject(bridge, 0, "Application").Visible = True
    execute.assert_called_once()


def test_recovery_is_bounded_and_only_accepts_explicit_expired_reply(
    bridge: ObjectBridge,
) -> None:
    execute = cast(Mock, bridge.driver.execute)
    execute.return_value = {
        "result": {"kind": "error", "message": "Native object scope expired"}
    }
    with pytest.raises(JsError):
        bridge.request("get", id=0, name="Name")
    execute.assert_called_once()
    execute.reset_mock()
    execute.side_effect = [
        {"result": {"kind": "expired"}},
        {"result": None},
        {"result": {"kind": "expired"}},
    ]
    with pytest.raises(JsError, match="changed again"):
        bridge.request("set", id=0, name="Visible", value=True)
    assert [call.args[2][1] for call in execute.call_args_list] == [
        "set",
        "restore",
        "set",
    ]


def test_collect_reads_all_scalar_fields_in_one_native_request(
    native_engine: Any,
) -> None:
    bridge, exchange = native_engine
    styles = bridge.request("begin").Styles
    before = len(exchange({"stats": True})["calls"])
    result = read_collection(styles, ("NameLocal", "Type", "BuiltIn", "InUse"))
    assert result == [
        {"NameLocal": "Normal", "Type": 1, "BuiltIn": True, "InUse": False},
        {"NameLocal": "Emphasis", "Type": 2, "BuiltIn": False, "InUse": True},
    ]
    assert exchange({"stats": True})["calls"][before:] == ["collect"]


def test_collect_preserves_non_javascript_native_behavior() -> None:
    items = [SimpleNamespace(Name="Normal", InUse=False)]
    collection = SimpleNamespace(Count=1, Item=lambda index: items[index - 1])
    assert read_collection(collection, ("Name", "InUse")) == [
        {"Name": "Normal", "InUse": False}
    ]


def test_collect_snapshots_nested_arrays_without_reading_unsafe_fields(
    native_engine: Any,
) -> None:
    bridge, exchange = native_engine
    series = bridge.request("begin").Series
    before = len(exchange({"stats": True})["calls"])
    assert read_collection(series, ("XValues", "Values")) == [
        {
            "XValues": [["Q1", None], ["Q2", True, None]],
            "Values": [[index, index + 1], [0, False]],
        }
        for index in (1, 2)
    ]
    assert exchange({"stats": True})["calls"][before:] == ["collect"]


@pytest.mark.parametrize(
    "field", ["NativeObject", "NativeFunction", "PlainObject"]
)
def test_collect_rejects_objects_nested_inside_arrays(
    native_engine: Any, field: str
) -> None:
    bridge, _ = native_engine
    series = bridge.request("begin").Series
    with pytest.raises(JsError, match="plain value or array"):
        read_collection(series, (field,))


def test_scope_stays_on_its_endpoint_when_component_socket_appears() -> None:
    writer = "/tmp/wpsdriverlocal_wps_1"
    presentation = "/tmp/wpsdriverlocal_wpp_1"
    requested: list[str | None] = []
    scopes: dict[str, str] = {}

    def execute(
        component: Component,
        script: str,
        args: list[Any],
        *,
        socket_path: str | None = None,
    ) -> dict[str, Any]:
        selected = socket_path or (presentation if requested else writer)
        requested.append(socket_path)
        key, command, _ = args
        result: Any = None
        if command == "begin":
            scopes[key] = selected
            result = {"kind": "object", "id": 0, "name": "Application"}
        else:
            assert scopes[key] == selected, "Native object scope expired"
            if command == "get":
                result = {"kind": "value", "value": True}
            elif command == "end":
                del scopes[key]
        return {"socket_path": selected, "result": result}

    driver = Mock(spec=MacDriver)
    driver.execute.side_effect = execute
    first = ObjectBridge(driver, "wpp", "cold-start")
    application = first.request("begin")
    application.Visible = True
    assert application.Visible is True
    first.request("end")
    second = ObjectBridge(driver, "wpp", "next-operation")
    second.request("begin")
    second.request("end")
    assert requested == [None, writer, writer, writer, None, presentation]
    assert scopes == {}


def test_object_arguments_preserve_identity_without_recreation(
    bridge: ObjectBridge,
) -> None:
    layout = JsObject(bridge, 17, "CustomLayout")
    add = JsObject(bridge, 5, "AddSlide")
    add(Index=2, pCustomLayout=layout)
    assert sent(bridge) == (
        "call",
        {
            "id": 5,
            "name": "AddSlide",
            "args": [2, {"kind": "ref", "id": 17}],
        },
    )
    cast(Mock, bridge.driver.execute).assert_called_once()


def test_cross_scope_reference_rejected_before_native_call(
    bridge: ObjectBridge,
) -> None:
    foreign = ObjectBridge(bridge.driver, "writer", "operation-two")
    source = JsObject(foreign, 2, "Shape")
    method = JsObject(bridge, 4, "BeginConnect")
    with pytest.raises(ValueError, match="cross operation scopes"):
        method(source, 1)
    target = JsObject(bridge, 3, "Target")
    with pytest.raises(ValueError, match="cross operation scopes"):
        target.Layout = source
    cast(Mock, bridge.driver.execute).assert_not_called()


@pytest.mark.parametrize(
    ("keywords", "expected"),
    [
        (
            {"Type": 1, "Operator": 5, "Formula1": "10", "Formula2": ""},
            [1, 5, "10", ""],
        ),
        (
            {"Type": 1, "AlertStyle": 1, "Operator": 5, "Formula1": "10"},
            [1, 1, 5, "10"],
        ),
        ({"BeforeRow": 3}, [3]),
        (
            {"Name": "total", "RefersTo": "=A1", "Visible": True},
            ["total", "=A1", True],
        ),
    ],
)
def test_add_overloads_choose_the_matching_native_signature(
    bridge: ObjectBridge, keywords: dict[str, Any], expected: list[Any]
) -> None:
    JsObject(bridge, 4, "Add")(**keywords)
    assert sent(bridge)[1]["args"] == expected


@pytest.mark.parametrize(
    ("component", "keywords", "expected"),
    [
        (
            "writer",
            {"ReadOnly": True, "AddToRecentFiles": False},
            ["sample.docx", {"kind": "missing"}, True, False],
        ),
        (
            "et",
            {"UpdateLinks": 0, "ReadOnly": False},
            ["sample.docx", 0, False],
        ),
        (
            "wpp",
            {"WithWindow": True},
            ["sample.docx", {"kind": "missing"}, {"kind": "missing"}, True],
        ),
    ],
)
def test_open_uses_component_specific_parameter_positions(
    bridge: ObjectBridge,
    component: Component,
    keywords: dict[str, Any],
    expected: list[Any],
) -> None:
    bridge.component = component
    JsObject(bridge, 4, "Open")("sample.docx", **keywords)
    assert sent(bridge)[1]["args"] == expected


def test_sparse_arguments_keep_explicit_null_distinct_from_missing(
    bridge: ObjectBridge,
) -> None:
    JsObject(bridge, 7, "ExportAsFixedFormat")(
        "deck.pdf", 2, PrintRange=None, RangeType=1
    )
    assert sent(bridge)[1]["args"] == [
        "deck.pdf",
        2,
        *({"kind": "missing"} for _ in range(5)),
        None,
        1,
    ]


@pytest.mark.parametrize(
    ("args", "keywords", "message"),
    [
        ((1,), {"Index": 2, "Layout": 12}, "Duplicate"),
        ((), {"Index": 1, "index": 2, "Layout": 12}, "Duplicate"),
        ((), {"typo": 1}, "Unsupported"),
    ],
)
def test_invalid_named_arguments_fail_before_sending(
    bridge: ObjectBridge,
    args: tuple[Any, ...],
    keywords: dict[str, Any],
    message: str,
) -> None:
    with pytest.raises(TypeError, match=message):
        JsObject(bridge, 7, "Add")(*args, **keywords)
    cast(Mock, bridge.driver.execute).assert_not_called()


def test_recursive_values_keep_dates_arrays_and_nulls(
    bridge: ObjectBridge,
) -> None:
    moment = datetime(2026, 10, 26, 12, 30, tzinfo=UTC)
    assert encode((None, True, date(2026, 10, 26), moment), bridge) == [
        None,
        True,
        {"kind": "date", "value": "2026-10-26"},
        {"kind": "date", "value": moment.isoformat()},
    ]
    payload = {
        "kind": "array",
        "value": [
            {"kind": "value", "value": None},
            {"kind": "value", "value": False},
            {"kind": "date", "value": "2026-10-26T12:30:00Z"},
            {
                "kind": "array",
                "value": [
                    {"kind": "value", "value": "中文\n"},
                    {"kind": "value", "value": 42},
                ],
            },
        ],
    }
    assert bridge.decode(payload) == [None, False, moment, ["中文\n", 42]]


def test_get_set_and_index_do_not_cache_mutable_native_state(
    bridge: ObjectBridge,
) -> None:
    cast(Mock, bridge.driver.execute).side_effect = [
        {"result": {"kind": "value", "value": 1}},
        {"result": None},
        {"result": {"kind": "value", "value": 2}},
        {"result": {"kind": "value", "value": 0.25}},
        {"result": None},
        {"result": {"kind": "value", "value": 0.75}},
    ]
    native = JsObject(bridge, 3, "Adjustments")
    assert native.Count == 1
    native.Count = 2
    assert native.Count == 2
    assert native[1] == 0.25
    native[1] = 0.75
    assert native[1] == 0.75
    commands = [
        call.args[2][1]
        for call in cast(Mock, bridge.driver.execute).call_args_list
    ]
    assert commands == ["get", "set", "get", "index", "setindex", "index"]


def test_collection_default_call_and_returned_reference_remain_callable(
    bridge: ObjectBridge,
) -> None:
    cast(Mock, bridge.driver.execute).side_effect = [
        {"result": {"kind": "object", "id": 8, "name": ""}},
        {"result": {"kind": "object", "id": 9, "name": ""}},
        {"result": {"kind": "value", "value": "Copied slide"}},
    ]
    duplicate = JsObject(bridge, 4, "Duplicate")
    assert duplicate()(1).Name == "Copied slide"
    calls = [
        call.args[2][2]
        for call in cast(Mock, bridge.driver.execute).call_args_list
    ]
    assert calls[0]["args"] == []
    assert calls[1]["args"] == [1]
    assert calls[2] == {"id": 9, "name": "Name"}


def test_slide_id_lookup_scans_collection_and_preserves_references() -> None:
    node = which("node")
    if node is None:
        pytest.skip("Node.js is needed to execute the isolated bridge script")
    script = """
const bridge=new Function('app','args',require('fs').readFileSync(0,'utf8'));
const slides=[{SlideID:10,Name:'first'},{SlideID:20,Name:'second'},
    {SlideID:30,Name:'third'}], visits=[];
const app={Slides:{Count:slides.length,
    Item(index){visits.push(index);return slides[index-1];},
    FindBySlideID(){return null;}}};
const run=(command,data={})=>bridge(app,['scope',command,data]);
const root=run('begin');
const collection=run('get',{id:root.id,name:'Slides'});
const method=run('get',{id:collection.id,name:'FindBySlideID'});
const names=[20,10,30,999].map(id=>{
    const slide=run('call',{id:method.id,name:'FindBySlideID',args:[id]});
    return slide===null?null:run('get',{id:slide.id,name:'Name'}).value;
});
run('end');
process.stdout.write(JSON.stringify({names,visits}));
"""
    result = subprocess.run(
        [node, "-e", script],
        input=BRIDGE,
        capture_output=True,
        text=True,
        check=True,
        timeout=10,
    )
    assert json.loads(result.stdout) == {
        "names": ["second", "first", "third", None],
        "visits": [1, 2, 1, 1, 2, 3, 1, 2, 3],
    }


def test_missing_members_and_native_failures_remain_distinct(
    bridge: ObjectBridge,
) -> None:
    cast(Mock, bridge.driver.execute).side_effect = [
        {"result": {"kind": "missing", "name": "Unsupported"}},
        {
            "result": {
                "kind": "error",
                "message": "rejected",
                "number": -2147467259,
            }
        },
    ]
    native = JsObject(bridge, 3, "Range")
    assert getattr(native, "Unsupported", None) is None
    with pytest.raises(JsError, match="get Type: rejected") as caught:
        _ = native.Type
    assert caught.value.operation == "get"
    assert caught.value.member == "Type"
    assert caught.value.hresult == -2147467259
    assert caught.value.detail == "rejected"


@pytest.mark.parametrize("value", [None, False, 0, "Series 1"])
@pytest.mark.parametrize(
    "name",
    [
        "Name",
        "HasTitleMaster",
        "TransparencyColor",
        "ThemeColor",
        "PatternThemeColor",
        "Direction",
    ],
)
def test_optional_property_preserves_non_javascript_values(
    value: Any, name: str
) -> None:
    assert optional_property(SimpleNamespace(**{name: value}), name) == value


@pytest.mark.parametrize(
    "name",
    [
        "Name",
        "HasTitleMaster",
        "TransparencyColor",
        "ThemeColor",
        "PatternThemeColor",
        "Direction",
    ],
)
def test_unsafe_optional_properties_never_query_native_getters(
    bridge: ObjectBridge, name: str
) -> None:
    execute = cast(Mock, bridge.driver.execute)
    execute.side_effect = AssertionError("Unsafe getter must not be queried")
    assert optional_property(JsObject(bridge, 3, "NativeObject"), name) is None
    execute.assert_not_called()


def test_optional_property_accepts_only_its_native_getter_failure(
    bridge: ObjectBridge,
) -> None:
    cast(Mock, bridge.driver.execute).return_value = {
        "result": {
            "kind": "error",
            "message": "Error: Function excute faild",
            "number": None,
        }
    }
    assert optional_property(JsObject(bridge, 3, "Series"), "Caption") is None
    cast(Mock, bridge.driver.execute).assert_called_once()


@pytest.mark.parametrize(
    "error",
    [
        JsError("set", "Name", NativeError("Error: Function excute faild")),
        JsError("get", "Count", NativeError("Error: Function excute faild")),
        JsError("get", "Name", NativeError("Access denied")),
        JsError("get", "Name", NativeError("Error: Function execute failed")),
        JsError("get", "Name", NativeError("Error: Function excute faild", 0)),
        JsError(
            "get",
            "Name",
            NativeError("Error: Function excute faild", -2147024891),
        ),
        NativeError("COM failure", -2147467259),
        AttributeError("missing property"),
        WpsError("unconfirmed reply"),
    ],
)
def test_optional_property_propagates_unrelated_failures(
    error: Exception,
) -> None:
    class Target:
        def __getattr__(self, name: str) -> Any:
            raise error

    with pytest.raises(type(error)) as caught:
        optional_property(Target(), "Name")
    assert caught.value is error


def test_closed_scope_never_sends_more_native_work(
    bridge: ObjectBridge,
) -> None:
    native = JsObject(bridge, 3, "Range")
    bridge.closed = True
    for operation in (
        lambda: native.Name,
        lambda: native(1),
        lambda: native[1],
        lambda: setattr(native, "Text", "new"),
    ):
        with pytest.raises(NativeError, match="outlive"):
            operation()
    cast(Mock, bridge.driver.execute).assert_not_called()


def test_session_releases_refs_even_when_result_contains_native_object() -> (
    None
):
    executor = MacExecutor()
    driver = Mock(spec=MacDriver)
    driver._unresolved = False
    executor.driver = driver
    driver.execute.side_effect = [
        {"result": None},
        {"result": {"kind": "object", "id": 0, "name": "Application"}},
        {"result": None},
    ]
    session = MacSession("writer", executor)
    references: list[JsObject] = []

    def operation() -> Any:
        application = session.connect(allow_launch=False)
        references.append(application)
        return {"native": application}

    try:
        with pytest.raises(WpsError, match="Only plain"):
            session.execute(operation)
        assert session.application is None
        assert session._bridge is None
        assert session.started
        with pytest.raises(NativeError, match="outlive"):
            _ = references[0].Name
        assert [
            call.args[2][1]
            for call in driver.execute.call_args_list
            if call.args[1] == BRIDGE
        ] == [
            "begin",
            "end",
        ]
    finally:
        assert session.close()


def test_runtime_components_share_one_serial_worker_and_isolate_lifespans(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(runtime, "sys", SimpleNamespace(platform="darwin"))
    entered, release, second_started = Event(), Event(), Event()
    events: list[str] = []

    def first() -> int:
        events.append("first-start")
        entered.set()
        assert release.wait(3)
        events.append("first-end")
        return get_ident()

    def second() -> int:
        second_started.set()
        events.append("second")
        return get_ident()

    with runtime.Runtime() as one, runtime.Runtime() as two:
        writer, et = one.sessions["writer"], one.sessions["et"]
        assert isinstance(writer, MacSession) and isinstance(et, MacSession)
        other = two.sessions["writer"]
        assert isinstance(other, MacSession)
        assert writer.executor is et.executor
        assert writer.executor is not other.executor
        with ThreadPoolExecutor(max_workers=2) as pool:
            first_call = pool.submit(writer.execute, first)
            assert entered.wait(1)
            second_call = pool.submit(et.execute, second)
            try:
                assert not second_started.wait(0.05)
            finally:
                release.set()
            assert first_call.result(timeout=2) == second_call.result(
                timeout=2
            )
    assert events == ["first-start", "first-end", "second"]


def test_session_wait_holds_resources_until_worker_finishes() -> None:
    session = MacSession("writer")
    entered, release = Event(), Event()

    def operation() -> bool:
        entered.set()
        assert release.wait(3)
        return True

    try:
        with ThreadPoolExecutor(max_workers=2) as pool:
            work = pool.submit(session.execute, operation)
            assert entered.wait(1)
            assert not session.close(timeout=0)
            waiting = pool.submit(session.wait)
            try:
                assert not waiting.done()
            finally:
                release.set()
            assert work.result(timeout=2)
            assert waiting.result(timeout=2) is None
    finally:
        release.set()
        session.close()


def test_scope_cleanup_does_not_replace_the_original_operation_error() -> None:
    executor = MacExecutor()
    driver = Mock(spec=MacDriver)
    driver._unresolved = False
    executor.driver = driver
    driver.execute.side_effect = [
        {"result": None},
        {"result": {"kind": "object", "id": 0, "name": "Application"}},
        {
            "result": {
                "kind": "error",
                "message": "cleanup failure",
                "number": None,
            }
        },
    ]
    session = MacSession("writer", executor)

    def operation() -> None:
        session.connect(allow_launch=False)
        raise ValueError("body failure")

    try:
        with pytest.raises(ValueError, match="body failure"):
            session.execute(operation)
        assert session.application is None
        assert session._bridge is None
        assert driver.execute.call_count == 3
        assert driver.execute.call_args_list[0].args[2] == ["Documents"]
    finally:
        assert session.close()


def test_shared_executor_defers_resources_and_stops_all_components(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(runtime, "sys", SimpleNamespace(platform="darwin"))
    original_close = MacSession.close
    monkeypatch.setattr(
        MacSession, "close", lambda self: original_close(self, timeout=0)
    )
    owner = runtime.Runtime()
    writer = owner.sessions["writer"]
    assert isinstance(writer, MacSession)
    entered, release = Event(), Event()
    cleanup = Mock()
    owner.resources.callback(cleanup)

    def operation() -> bool:
        entered.set()
        assert release.wait(3)
        return True

    with ThreadPoolExecutor(max_workers=1) as pool:
        call = pool.submit(writer.execute, operation)
        assert entered.wait(1)
        try:
            owner.close()
            assert owner.cleanup_thread is not None
            cleanup.assert_not_called()
            for session in owner.sessions.values():
                with pytest.raises(WpsError, match="closed"):
                    session.execute(lambda: None)
        finally:
            release.set()
        assert call.result(timeout=2)
    assert owner.cleanup_thread is not None
    owner.cleanup_thread.join(timeout=2)
    assert not owner.cleanup_thread.is_alive()
    cleanup.assert_called_once()
    owner.close()


def test_session_wait_inside_worker_does_not_join_itself() -> None:
    session = MacSession("writer")
    try:
        assert session.execute(session.wait) is None
    finally:
        assert session.close()


@pytest.mark.parametrize("active_window", [False, True])
@pytest.mark.parametrize("selector", ["Budget", "/reports/./Budget.xlsx"])
def test_cross_component_activation_leaves_selection_to_shared_contract(
    monkeypatch: pytest.MonkeyPatch,
    active_window: bool,
    selector: str,
) -> None:
    node = which("node")
    if node is None:
        pytest.skip("Node.js is needed to execute the isolated bridge script")
    session = MacSession("et")
    driver = Mock(spec=MacDriver)
    session.executor.driver = driver
    token = CALL_DOCUMENT.set(selector)
    try:
        session._activate()
    finally:
        CALL_DOCUMENT.reset(token)
        assert session.close()
    _, script, args = driver.execute.call_args.args
    program = """
const input=JSON.parse(require('fs').readFileSync(0,'utf8'));
let activated=false;
const window={Activate(){activated=true;}};
const document={Name:'Budget.xlsx',FullName:'/reports/Budget.xlsx',
    Windows:{Count:1,Item(){return window;}}};
globalThis.Application={Name:'WPP'};
const app={Name:'ET',ActiveWindow:input.active?window:null,
    Workbooks:{Count:1,Item(){return document;}}};
new Function('app','args',input.script)(app,input.args);
process.stdout.write(JSON.stringify({activated}));
"""
    result = subprocess.run(
        [node, "-e", program],
        input=json.dumps(
            {"script": script, "args": args, "active": active_window}
        ),
        capture_output=True,
        text=True,
        check=True,
        timeout=10,
    )
    assert json.loads(result.stdout) == {"activated": True}


def test_session_rejects_native_access_from_another_thread() -> None:
    session = MacSession("writer")
    try:
        with pytest.raises(WpsError, match="WPS worker"):
            session.connect(allow_launch=False)
    finally:
        assert session.close()
