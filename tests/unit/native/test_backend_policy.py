"""Verify one capability contract across COM, RPC and JavaScript."""

from contextlib import nullcontext
from types import SimpleNamespace
from typing import cast, get_args
from unittest.mock import Mock

import pytest

from wps_mcp.documents.models import ConnectInput
from wps_mcp.errors import WpsUnavailableError
from wps_mcp.models import Component
from wps_mcp.native import discovery, macos, runtime
from wps_mcp.native.capabilities import Feature, get_capabilities
from wps_mcp.native.javascript import JsObject
from wps_mcp.native.linux import RpcSession
from wps_mcp.native.objects import (
    capabilities_for,
    optional_property,
    read_reference_targets,
)
from wps_mcp.native.rpc import RpcObject
from wps_mcp.native.windows import ComSession


@pytest.mark.parametrize(
    ("platform", "backend", "blocked"),
    [
        ("win32", "com", set()),
        ("linux", "rpc", {"chart_workbook"}),
        (
            "darwin",
            "js",
            {
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
                "animation_direction",
                "title_master",
                "transparency_color",
                "theme_color_read",
            },
        ),
    ],
)
def test_all_registered_features_have_explicit_backend_policy(
    platform: str, backend: str, blocked: set[str]
) -> None:
    profile = get_capabilities(platform)
    assert profile is get_capabilities(backend)
    assert profile.platform == platform and profile.backend == backend
    for feature in get_args(Feature):
        if feature in blocked:
            restriction = profile.limitation(feature)
            assert restriction is not None and restriction.reason
            assert not profile.allows(feature)
            with pytest.raises(NotImplementedError) as caught:
                profile.require(feature)
            assert str(caught.value) == restriction.reason
        else:
            assert profile.allows(feature)
            profile.require(feature)


@pytest.mark.parametrize("identifier", ["unknown", "", "windows"])
def test_unknown_backends_are_not_treated_as_windows(identifier: str) -> None:
    with pytest.raises(WpsUnavailableError, match="Unsupported"):
        get_capabilities(identifier)


@pytest.mark.parametrize("backend", ["com", "rpc", "js"])
def test_unknown_features_are_not_implicitly_allowed(backend: str) -> None:
    with pytest.raises(ValueError, match="Unknown native capability"):
        get_capabilities(backend).allows(cast(Feature, "not_registered"))


@pytest.mark.parametrize("host", ["win32", "linux", "darwin"])
@pytest.mark.parametrize(
    ("binding", "backend"),
    [(None, "com"), (RpcObject, "rpc"), (JsObject, "js")],
)
def test_object_binding_selects_policy_without_native_access(
    monkeypatch: pytest.MonkeyPatch,
    host: str,
    binding: type | None,
    backend: str,
) -> None:
    monkeypatch.setattr(runtime, "sys", SimpleNamespace(platform=host))
    native = Mock(spec=binding) if binding is not None else Mock()
    assert capabilities_for(native) is get_capabilities(backend)
    assert native.mock_calls == []


@pytest.mark.parametrize("binding", [None, RpcObject, JsObject])
@pytest.mark.parametrize("member", ["Name", "ThemeColor", "Direction"])
def test_unsafe_getter_policy_stays_scoped_to_its_binding(
    binding: type | None, member: str
) -> None:
    native = Mock(spec=binding) if binding is not None else Mock()
    setattr(native, member, 0)
    assert optional_property(native, member) == (
        None if binding is JsObject else 0
    )
    assert native.mock_calls == []


@pytest.mark.parametrize("binding", [None, RpcObject, JsObject])
def test_reference_fallback_is_selected_by_binding(
    binding: type | None,
) -> None:
    document = Mock(spec=binding) if binding is not None else Mock()
    document.GetCrossReferenceItems = Mock(return_value=None)
    document.Footnotes = Mock(Count=1)
    document.Footnotes.Item.return_value.Range.Text = "Footnote\r"
    expected = ["Footnote"] if binding is RpcObject else []
    assert read_reference_targets(document, 3) == expected
    assert document.Footnotes.Item.call_count == int(binding is RpcObject)


def test_windows_discovery_reports_policy_without_claiming_verified_support(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(discovery, "sys", SimpleNamespace(platform="win32"))
    registry = Mock()
    registry.OpenKey.return_value = nullcontext()
    monkeypatch.setattr(discovery, "import_module", lambda name: registry)
    result = discovery.detect_components()
    assert result["backend"] == "com"
    assert (
        result["limitations"] == get_capabilities("win32").tool_limitations()
    )
    assert result["limitations"] == []
    assert all(item["ready"] for item in result["components"])
    assert all(
        not item["native_execution_verified"] for item in result["components"]
    )


def test_runtime_rejects_unknown_platform_before_creating_sessions(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        runtime, "sys", SimpleNamespace(platform="unsupported")
    )
    with pytest.raises(WpsUnavailableError, match="Unsupported"):
        runtime.Runtime()


@pytest.mark.parametrize("component", ["writer", "et", "wpp"])
def test_shared_visibility_rule_is_applied_at_input_boundary(
    component: Component,
) -> None:
    if component == "wpp":
        with pytest.raises(ValueError, match="WPP cannot be hidden"):
            ConnectInput(component=component, visible=False)
    else:
        assert (
            ConnectInput(component=component, visible=False).visible is False
        )


def test_visibility_rule_also_protects_direct_backend_calls(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    connector = Mock()
    windows = ComSession("wpp", connector=connector, scope=nullcontext)
    try:
        with pytest.raises(ValueError, match="WPP cannot be hidden"):
            windows.execute(lambda: windows.connect(False))
        connector.assert_not_called()
    finally:
        windows.close()
    linux = RpcSession("wpp")
    with pytest.raises(ValueError, match="WPP cannot be hidden"):
        linux.connect(False)
    assert linux.application is None and not linux.started
    apple = macos.MacSession("wpp")
    monkeypatch.setattr(
        macos, "current_thread", lambda: apple.executor.worker.thread
    )
    apple._bridge = Mock()
    with pytest.raises(ValueError, match="WPP cannot be hidden"):
        apple.connect(False)
    assert apple.application is None and not apple.started
