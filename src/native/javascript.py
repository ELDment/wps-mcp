"""Adapt WPS JavaScript objects to the shared native Python contract."""

from __future__ import annotations

from collections.abc import Iterator
from datetime import date, datetime
from typing import TYPE_CHECKING, Any

from src.errors import NativeError
from src.native.components import SPECS

if TYPE_CHECKING:
    from src.models import Component
    from src.native.macos import MacDriver

# Only named parameters used by the shared operations need positional binding.
SIGNATURES = {
    "AddShape": "Type Left Top Width Height",
    "AddTextbox": "Orientation Left Top Width Height",
    "AddLine": "BeginX BeginY EndX EndY",
    "AddPicture": "FileName LinkToFile SaveWithDocument Left Top Width Height",
    "AddTable": "NumRows NumColumns Left Top Width Height",
    "AddSlide": "Index pCustomLayout",
    "MoveTo": "toPos",
    "Characters": "Start Length",
    "TwoColorGradient": "Style Variant",
    "Sort": (
        "Key1 Order1 Key2 Type Order2 Key3 Order3 Header "
        "OrderCustom MatchCase Orientation"
    ),
    "AutoFilter": "Field Criteria1 Operator Criteria2 VisibleDropDown",
    "Find": (
        "What After LookIn LookAt SearchOrder SearchDirection "
        "MatchCase MatchByte SearchFormat"
    ),
    "Cut": "Destination",
    "Modify": "Type AlertStyle Operator Formula1 Formula2",
    "ExportAsFixedFormat": (
        "Path FixedFormatType Intent FrameSlides HandoutOrder "
        "OutputType PrintHiddenSlides PrintRange RangeType"
    ),
    "OpenDataSource": (
        "Name Format ConfirmConversions ReadOnly LinkToSource "
        "AddToRecentFiles PasswordDocument PasswordTemplate Revert "
        "WritePasswordDocument WritePasswordTemplate Connection "
        "SQLStatement SQLStatement1 OpenExclusive SubType"
    ),
}
OVERLOADS = {
    "Add": (
        "Template NewTemplate DocumentType Visible",
        (
            "Range UseHeadingStyles UpperHeadingLevel "
            "LowerHeadingLevel UseFields TableID "
            "RightAlignPageNumbers IncludePageNumbers AddedStyles "
            "UseHyperlinks HidePageNumbersInWeb UseOutlineLevels"
        ),
        "Range Reference Text",
        "Type AlertStyle Operator Formula1 Formula2",
        "Type Operator Formula1 Formula2",
        "Name RefersTo Visible",
        "Before After Count Type",
        "SourceType Source LinkSource XlListObjectHasHeaders Destination",
        "Index Layout",
        "BeforeRow",
        "BeforeColumn",
    ),
    "Execute": (
        (
            "FindText MatchCase MatchWholeWord MatchWildcards "
            "MatchSoundsLike MatchAllWordForms Forward Wrap Format "
            "ReplaceWith Replace"
        ),
        "Pause",
    ),
}
OPEN_SIGNATURES = {
    "writer": "FileName ConfirmConversions ReadOnly AddToRecentFiles",
    "et": "Filename UpdateLinks ReadOnly",
    "wpp": "FileName ReadOnly Untitled WithWindow",
}

BRIDGE = """
var key=args[0], command=args[1], input=args[2];
var readMethods=['Item','Range','Characters','FindBySlideID'];
var documentGetters={Documents:['ActiveDocument','Document'],
    Workbooks:['ActiveWorkbook','Workbook'],
    Presentations:['ActivePresentation','Presentation']};
function documentCollection(id) {
    var recipe=state.recipes[id];
    if(!recipe) return false;
    if(recipe.kind==='item' && !recipe.args.length)
        return documentCollection(recipe.id);
    return recipe.kind==='get' && recipe.name===state.collection
        && state.recipes[recipe.id]
        && state.recipes[recipe.id].kind==='app';
}
function property(object, name) {
    if(!(name in object)) throw new Error('Missing native member: '+name);
    var value=object[name];
    if(['Address','AddressLocal','DataFields'].includes(name)
        && typeof value==='function')
        value=Reflect.apply(value,object,[]);
    return value;
}
function readCall(object, owner, name, values) {
    if(!readMethods.includes(name))
        throw new Error('Native call cannot be restored safely: '+name);
    if(name==='FindBySlideID') {
        for(var i=1;i<=owner.Count;i++) {
            var slide=owner.Item(i);
            if(slide.SlideID===values[0]) return slide;
        }
        return null;
    }
    if(name==='Range' && values.length===2) {
        values=values.map(v=>v && typeof v.Address==='function'
            ?v.Address():v);
    }
    return Reflect.apply(object,owner,values);
}
function resolve(id) {
    if(id in state.objects) return state.objects[id];
    var recipe=state.recipes[id], value, owner;
    if(!recipe) throw new Error(
        'Native reference cannot be restored safely; no write was replayed');
    if(state.resolving.has(id)) throw new Error('Cyclic native reference');
    state.resolving.add(id);
    try {
        if(recipe.kind==='app') value=app;
        else if(recipe.kind==='document') {
            var documents=app[state.collection], matches=[];
            for(var i=1;i<=documents.Count;i++) {
                var document=documents.Item(i);
                if(String(document.FullName)===recipe.full_name)
                    matches.push(document);
            }
            if(matches.length!==1)
                throw new Error('Native document cannot be uniquely restored');
            value=matches[0];
        } else {
            owner=resolve(recipe.id);
            if(recipe.kind==='get') value=property(owner,recipe.name);
            else if(recipe.kind==='item') {
                var values=unpack(recipe.args);
                value=values.length?Reflect.apply(owner.Item,owner,values):owner;
            } else if(recipe.kind==='call')
                value=readCall(owner,state.owners[recipe.id],
                    recipe.name,unpack(recipe.args));
            else throw new Error('Unknown native reference recipe');
        }
        state.objects[id]=value;state.owners[id]=owner;
        return value;
    } finally {state.resolving.delete(id);}
}
function pack(value, owner, name, recipe) {
    if(value===undefined || value===null) return {kind:'value',value:null};
    if(Array.isArray(value)) return {kind:'array',value:value.map(v=>pack(v))};
    if(value instanceof Date) return {kind:'date',value:value.toISOString()};
    if(typeof value!=='object' && typeof value!=='function')
        return {kind:'value',value:value};
    if(value===app) recipe={kind:'app'};
    else if(recipe && recipe.kind==='document')
        recipe={kind:'document',full_name:String(value.FullName)};
    var id=state.objects.length;
    state.objects.push(value);state.owners.push(owner);
    state.recipes.push(recipe||null);
    return {kind:'object',id:id,name:name||'',recipe:recipe||null};
}
function unpack(value) {
    if(!value || typeof value!=='object') return value;
    if(value.kind==='ref') return resolve(value.id);
    if(value.kind==='missing') return undefined;
    if(value.kind==='date') return new Date(value.value);
    if(Array.isArray(value)) return value.map(unpack);
    var result={};for(var k in value) result[k]=unpack(value[k]);return result;
}
function plainField(value) {
    if(value===undefined || value===null) return null;
    if(Array.isArray(value)) return value.map(plainField);
    if(['string','boolean','number'].includes(typeof value)) return value;
    throw new Error('Collection field must be a plain value or array');
}
try {
    if(command==='begin') {
        globalThis[key]={objects:[],owners:[],recipes:[],resolving:new Set(),
            collection:input.collection};
        var state=globalThis[key];return pack(app);
    }
    if(command==='end') {delete globalThis[key];return null;}
    if(command==='restore') {
        globalThis[key]={objects:new Array(input.recipes.length),
            owners:new Array(input.recipes.length),recipes:input.recipes,
            collection:input.collection,resolving:new Set()};
        return null;
    }
    var state=globalThis[key];
    if(!state) return {kind:'expired'};
    var object=resolve(input.id), name=input.name;
    if(command==='get') {
        if(!(name in object)) return {kind:'missing',name:name};
        var recipe=(documentGetters[state.collection]||[]).includes(name)
            ?{kind:'document'}:{kind:'get',id:input.id,name:name};
        return pack(property(object,name),object,name,
            recipe);
    }
    if(command==='collect') {
        var rows=[];
        for(var i=1;i<=object.Count;i++) {
            var item=object.Item(i), row={};
            for(var field of input.fields) {
                row[field]=plainField(property(item,field));
            }
            rows.push(row);
        }
        return {kind:'value',value:rows};
    }
    if(command==='set') {
        if(!(name in object)) return {kind:'missing',name:name};
        var value=unpack(input.value);
        if(name==='LinkToPrevious') {
            var text=object.Range.Text;
            if(typeof value==='boolean') value=value?-1:0;
        }
        object[name]=value;return null;
    }
    if(command==='call') {
        var owner=state.owners[input.id], values=unpack(input.args);
        var method=state.recipes[input.id];
        var ownerId=method && method.kind==='get'?method.id:null;
        var documentMethod=ownerId!==null && documentCollection(ownerId);
        if('Item' in object) {
            return pack(values.length
                ?Reflect.apply(object.Item,object,values):object,
                undefined,undefined,
                values.length && documentCollection(input.id)
                    ?{kind:'document'}
                    :{kind:'item',id:input.id,args:input.args});
        }
        if(readMethods.includes(name)) {
            var value=readCall(object,owner,name,values);
            return value===null?null:pack(value,undefined,undefined,
                name==='Item' && documentMethod?{kind:'document'}
                    :{kind:'call',id:input.id,name:name,args:input.args});
        }
        if(name==='Sort' && owner.Worksheet) {
            var sort=owner.Worksheet.Sort;
            sort.SortFields.Clear();sort.SortFields.Add(values[0],0,values[1]);
            sort.SetRange(owner);sort.Header=values[7];
            sort.MatchCase=!!values[9];sort.Orientation=values[10];sort.Apply();
            return null;
        }
        var saving=['SaveAs','SaveAs2'].includes(name)
            && ownerId!==null && state.recipes[ownerId]
            && state.recipes[ownerId].kind==='document';
        var oldName=saving?String(owner.FullName):null;
        var result=pack(Reflect.apply(object,owner,values),undefined,undefined,
            ['Add','Open'].includes(name) && documentMethod
                ?{kind:'document'}:null);
        if(saving) result.rename={old_name:oldName,
            full_name:String(owner.FullName)};
        return result;
    }
    if(command==='index') return pack(object.Item(input.index),
        undefined,undefined,documentCollection(input.id)?{kind:'document'}
            :{kind:'item',id:input.id,args:[input.index]});
    if(command==='setindex') {
        object.Item(input.index,unpack(input.value));return null;
    }
    throw new Error('Unknown native bridge command');
} catch(error) {
    return {kind:'error',message:String(error),
        number:error.number||error.hresult||null};
}
"""


class JsError(NativeError):
    """Preserve the operation and original detail of a native JS failure."""

    def __init__(
        self, operation: str, member: str, error: NativeError
    ) -> None:
        super().__init__(f"{operation} {member}: {error}", error.hresult)
        self.operation = operation
        self.member = member
        self.detail = str(error)


def encode(value: Any, bridge: ObjectBridge) -> Any:
    """Transfer references only within their original operation scope."""
    if isinstance(value, JsObject):
        if value._bridge is not bridge:
            raise ValueError("Native objects cannot cross operation scopes")
        return {"kind": "ref", "id": value._id}
    if isinstance(value, (datetime, date)):
        return {"kind": "date", "value": value.isoformat()}
    if isinstance(value, (list, tuple)):
        return [encode(item, bridge) for item in value]
    if isinstance(value, dict):
        return {key: encode(item, bridge) for key, item in value.items()}
    return value


class ObjectBridge:
    """Own a short-lived WPS reference table for one complete operation."""

    def __init__(
        self, driver: MacDriver, component: Component, key: str
    ) -> None:
        self.driver = driver
        self.component: Component = component
        self.key = key
        self.closed = False
        self._socket_path: str | None = None
        self._recipes: list[dict[str, Any] | None] = []

    def request(self, command: str, **data: Any) -> Any:
        """Execute one native access and decode plain values or references."""
        if self.closed:
            raise NativeError("Native objects cannot outlive their operation")
        if command == "begin":
            data["collection"] = SPECS[self.component].collection
        try:
            result = self._send(command, data)
            if isinstance(result, dict) and result.get("kind") == "expired":
                self.decode(
                    self._send(
                        "restore",
                        {
                            "recipes": self._recipes,
                            "collection": SPECS[self.component].collection,
                        },
                    )
                )
                result = self._send(command, data)
            return self.decode(result)
        except NativeError as error:
            raise JsError(command, data.get("name", ""), error) from error

    def _send(self, command: str, data: dict[str, Any]) -> Any:
        reply = self.driver.execute(
            self.component,
            BRIDGE,
            [self.key, command, data],
            socket_path=self._socket_path,
        )
        if self._socket_path is None:
            path = reply.get("socket_path")
            if isinstance(path, str):
                self._socket_path = path
        return reply["result"]

    def decode(self, result: Any) -> Any:
        """Detach arrays and primitives without serializing native objects."""
        if result is None:
            return None
        if rename := result.get("rename"):
            for recipe in self._recipes:
                if (
                    recipe is not None
                    and recipe.get("kind") == "document"
                    and recipe.get("full_name") == rename["old_name"]
                ):
                    recipe["full_name"] = rename["full_name"]
        kind = result["kind"]
        if kind == "expired":
            raise NativeError(
                "Native context changed again during recovery; "
                "the current command was not executed"
            )
        if kind == "missing":
            raise AttributeError(f"WPS has no member {result['name']}")
        if kind == "error":
            raise NativeError(result["message"], result.get("number"))
        if kind == "object":
            identity = result["id"]
            while len(self._recipes) <= identity:
                self._recipes.append(None)
            self._recipes[identity] = result.get("recipe")
            return JsObject(self, identity, result["name"])
        if kind == "array":
            return [self.decode(item) for item in result["value"]]
        if kind == "date":
            return datetime.fromisoformat(
                result["value"].replace("Z", "+00:00")
            )
        return result["value"]


class JsObject:
    """Expose native properties, methods and callable collections uniformly."""

    def __init__(self, bridge: ObjectBridge, identity: int, name: str) -> None:
        object.__setattr__(self, "_bridge", bridge)
        object.__setattr__(self, "_id", identity)
        object.__setattr__(self, "_name", name)

    def __getattr__(self, name: str) -> Any:
        try:
            return self._bridge.request("get", id=self._id, name=name)
        except AttributeError as error:
            raise AttributeError(
                f"{self._name or 'Object'}.{name}: {error}"
            ) from error

    def __setattr__(self, name: str, value: Any) -> None:
        self._bridge.request(
            "set", id=self._id, name=name, value=encode(value, self._bridge)
        )

    def __call__(self, *args: Any, **kwargs: Any) -> Any:
        """Bind supported named parameters to native positional arguments."""
        values = [encode(value, self._bridge) for value in args]
        if kwargs:
            candidates = (
                (OPEN_SIGNATURES[self._bridge.component],)
                if self._name == "Open"
                else OVERLOADS.get(
                    self._name, (SIGNATURES.get(self._name, ""),)
                )
            )
            named = {name.casefold(): value for name, value in kwargs.items()}
            if len(named) != len(kwargs):
                raise TypeError("Duplicate case-insensitive named argument")
            signatures = [
                candidate.lower().split() for candidate in candidates
            ]
            signature = min(
                (s for s in signatures if named.keys() <= set(s)),
                key=len,
                default=None,
            )
            if signature is None:
                raise TypeError(
                    f"Unsupported named arguments for {self._name}: "
                    f"{list(kwargs)}"
                )
            for name, value in named.items():
                index = signature.index(name)
                if index < len(args):
                    raise TypeError(f"Duplicate argument: {name}")
                while len(values) <= index:
                    values.append({"kind": "missing"})
                values[index] = encode(value, self._bridge)
        return self._bridge.request(
            "call", id=self._id, name=self._name, args=values
        )

    def __getitem__(self, index: Any) -> Any:
        return self._bridge.request("index", id=self._id, index=index)

    def __setitem__(self, index: Any, value: Any) -> None:
        self._bridge.request(
            "setindex",
            id=self._id,
            index=index,
            value=encode(value, self._bridge),
        )

    def __iter__(self) -> Iterator[Any]:
        return (self.Item(index) for index in range(1, int(self.Count) + 1))

    def __len__(self) -> int:
        return int(self.Count)

    def __bool__(self) -> bool:
        return True
