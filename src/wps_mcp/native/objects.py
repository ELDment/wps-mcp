"""Keep backend capabilities and object reads behind one native boundary."""

from collections.abc import Iterator
from contextlib import contextmanager
from typing import Any

from wps_mcp.errors import NativeError, native_errors
from wps_mcp.models import NativeObject
from wps_mcp.native.capabilities import (
    BackendCapabilities,
    get_capabilities,
)
from wps_mcp.native.javascript import JsError, JsObject
from wps_mcp.native.rpc import RpcObject


def capabilities_for(target: NativeObject) -> BackendCapabilities:
    """Identify the binding without reading properties on native proxies."""
    if isinstance(target, JsObject):
        return get_capabilities("js")
    if isinstance(target, RpcObject):
        return get_capabilities("rpc")
    return get_capabilities("com")


def is_member_unavailable(
    error: Exception, operation: str, member: str
) -> bool:
    """Match the binding's unnumbered optional-member failure exactly."""
    return (
        isinstance(error, JsError)
        and error.operation == operation
        and error.member == member
        and error.hresult is None
        and error.detail == "Error: Function excute faild"
    )


def matches_member_error(
    error: Exception,
    operation: str,
    member: str,
    *,
    native_code: int,
    dispatch_codes: set[int],
) -> bool:
    """Match explicit binding failures while preserving domain checks."""
    info = getattr(error, "excepinfo", None)
    return (
        is_member_unavailable(error, operation, member)
        or (type(error) is NativeError and error.hresult == native_code)
        or (
            getattr(error, "hresult", None) == -2147352567
            and isinstance(info, tuple)
            and len(info) > 5
            and info[5] in dispatch_codes
        )
    )


def is_optional_member_error(error: Exception) -> bool:
    """Accept only known missing-member failures for optional metadata."""
    if isinstance(error, AttributeError):
        return True
    if isinstance(error, JsError):
        return is_member_unavailable(error, error.operation, error.member)
    if not isinstance(error, native_errors()):
        return False
    code = getattr(error, "hresult", None)
    if code == -2147352567:
        info = getattr(error, "excepinfo", None)
        code = info[5] if isinstance(info, tuple) and len(info) > 5 else None
    return code in {-2147352573, -2147352570, -2147467263}


@contextmanager
def optional_members() -> Iterator[None]:
    """Ignore known missing metadata members; propagate other failures."""
    try:
        yield
    except Exception as error:
        if not is_optional_member_error(error):
            raise


def optional_property(target: NativeObject, name: str) -> Any:
    """Read optional metadata without invoking known unsafe getters."""
    if not capabilities_for(target).can_read(name):
        return None
    try:
        return getattr(target, name)
    except NativeError as error:
        if is_member_unavailable(error, "get", name):
            return None
        raise


def read_collection(
    collection: NativeObject, fields: tuple[str, ...]
) -> list[dict[str, Any]]:
    """Read plain collection fields with one evaluation when supported."""
    if isinstance(collection, JsObject):
        return collection._bridge.request(
            "collect", id=collection._id, fields=fields
        )
    result = []
    for index in range(1, int(collection.Count) + 1):
        item = collection.Item(index)
        result.append({name: getattr(item, name) for name in fields})
    return result


def read_reference_targets(document: NativeObject, kind: int) -> list[str]:
    """Read Writer targets in native order, applying the binding's fallback."""
    if kind not in {1, 2, 3, 4}:
        raise ValueError(f"Unsupported cross-reference type: {kind}")
    values = document.GetCrossReferenceItems(kind)
    if values is None:
        if capabilities_for(document).reference_fallback == "collections":
            return _collection_reference_targets(document, kind)
        return []
    if not isinstance(values, (tuple, list)) or any(
        not isinstance(item, str) for item in values
    ):
        raise RuntimeError(
            "WPS returned an invalid cross-reference target array"
        )
    return list(values)


def _collection_reference_targets(
    document: NativeObject, kind: int
) -> list[str]:
    """Recover omitted arrays from built-in headings or native collections."""
    if kind == 1:
        headings = {
            str(document.Styles.Item(identifier).NameLocal)
            for identifier in range(-2, -11, -1)
        }
        paragraphs = document.Paragraphs
        result = []
        for index in range(1, int(paragraphs.Count) + 1):
            text_range = paragraphs.Item(index).Range
            if str(text_range.Style.NameLocal) in headings:
                result.append(str(text_range.Text).rstrip("\r"))
        return result
    collection = getattr(
        document, {2: "Bookmarks", 3: "Footnotes", 4: "Endnotes"}[kind]
    )
    result = []
    for index in range(1, int(collection.Count) + 1):
        item = collection.Item(index)
        result.append(
            str(item.Name) if kind == 2 else str(item.Range.Text).rstrip("\r")
        )
    return result
