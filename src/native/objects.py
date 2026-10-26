"""Keep backend capabilities and object reads behind one native boundary."""

from typing import Any, Literal

from src.errors import NativeError
from src.models import NativeObject
from src.native.javascript import JsError, JsObject
from src.native.rpc import RpcObject

Feature = Literal[
    "chart_workbook",
    "chart_activation",
    "chart_series_names",
    "connection_sites",
    "freeform_builder",
    "freeform_nodes",
    "delete_first_section",
    "slideshow",
    "smartart_nodes",
    "theme_color_assignment",
]

JS_LIMITATIONS: dict[Feature, str] = {
    "chart_workbook": (
        "macOS WPS does not expose ChartData.Workbook through its "
        "JavaScript driver; embedded chart data was not changed"
    ),
    "chart_activation": "The chart data editor cannot be safely activated",
    "chart_series_names": "Chart series names cannot be safely read",
    "connection_sites": (
        "macOS WPS does not expose ConnectionSites coordinates; "
        "use numeric connection site indices instead of directions"
    ),
    "freeform_builder": (
        "macOS WPS supports freeform creation with straight segments "
        "only; Bezier control points are unavailable"
    ),
    "freeform_nodes": (
        "macOS WPS cannot read or edit freeform nodes through its "
        "JavaScript driver; the shape was not changed"
    ),
    "delete_first_section": "The first section requires move-before-delete",
    "slideshow": (
        "macOS WPS slide show automation is unsafe in the tested build: "
        "exiting leaves invalid windows and can crash WPS. "
        "Use WPS playback controls; no show action was sent."
    ),
    "smartart_nodes": (
        "macOS WPS does not expose SmartArt nodes; "
        "only diagram layout, color and style can be changed"
    ),
    "theme_color_assignment": "Theme colors require a native scheme import",
}
RPC_LIMITATIONS: dict[Feature, str] = {
    "chart_workbook": (
        "Linux WPS RPC cannot safely edit embedded chart workbooks; "
        "chart data was not changed"
    ),
}


def _limitations(target: NativeObject) -> dict[Feature, str]:
    if isinstance(target, JsObject):
        return JS_LIMITATIONS
    if isinstance(target, RpcObject):
        return RPC_LIMITATIONS
    return {}


def supports_feature(target: NativeObject, feature: Feature) -> bool:
    """Check known binding limitations without probing unsafe members."""
    return feature not in _limitations(target)


def require_feature(target: NativeObject, feature: Feature) -> None:
    """Reject unavailable native behavior before beginning a mutation."""
    reason = _limitations(target).get(feature)
    if reason is not None:
        raise NotImplementedError(reason)


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


def optional_property(target: NativeObject, name: str) -> Any:
    """Read optional metadata without invoking known unsafe getters."""
    if isinstance(target, JsObject) and name in {
        "Name",
        "HasTitleMaster",
        "TransparencyColor",
        "ThemeColor",
        "PatternThemeColor",
        "Direction",
    }:
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
