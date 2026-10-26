"""Define backend capability policy and its public explanations once."""

from collections.abc import Mapping
from dataclasses import dataclass
from types import MappingProxyType
from typing import Literal, cast

from wps_mcp.errors import WpsUnavailableError
from wps_mcp.models import Component, JsonObject

Backend = Literal["com", "rpc", "js"]


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
    "animation_direction",
    "title_master",
    "transparency_color",
    "theme_color_read",
]


@dataclass(frozen=True, slots=True)
class Limitation:
    """Explain a known restriction, its operations and unsafe members."""

    reason: str
    tools: tuple[tuple[str, str], ...] = ()
    members: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class Capability:
    """Declare a known feature policy explicitly for every backend."""

    com: Limitation | None
    rpc: Limitation | None
    js: Limitation | None


# None permits the native path; it does not verify the installed WPS build.
_CAPABILITIES: Mapping[Feature, Capability] = MappingProxyType(
    {
        "chart_workbook": Capability(
            com=None,
            rpc=Limitation(
                "Linux WPS RPC cannot safely edit embedded chart workbooks; "
                "chart data was not changed",
                (("wps_ppt_charts", "set_data"),),
            ),
            js=Limitation(
                "macOS WPS does not expose ChartData.Workbook through its "
                "JavaScript driver; embedded chart data was not changed",
                (("wps_ppt_charts", "set_data"),),
            ),
        ),
        "chart_activation": Capability(
            com=None,
            rpc=None,
            js=Limitation("The chart data editor cannot be safely activated"),
        ),
        "chart_series_names": Capability(
            com=None,
            rpc=None,
            js=Limitation(
                "Series names are null; "
                "new charts may have empty cached labels.",
                (("wps_ppt_charts", "add get_data set_type"),),
                ("Name",),
            ),
        ),
        "connection_sites": Capability(
            com=None,
            rpc=None,
            js=Limitation(
                "macOS WPS does not expose ConnectionSites coordinates; "
                "use numeric connection site indices instead of directions",
                (("wps_ppt_connectors", "add format"),),
            ),
        ),
        "freeform_builder": Capability(
            com=None,
            rpc=None,
            js=Limitation(
                "macOS WPS supports freeform creation with straight segments "
                "only; Bezier control points are unavailable",
                (("wps_ppt_freeform", "build"),),
            ),
        ),
        "freeform_nodes": Capability(
            com=None,
            rpc=None,
            js=Limitation(
                "macOS WPS cannot read or edit freeform nodes through its "
                "JavaScript driver; the shape was not changed",
                (
                    (
                        "wps_ppt_freeform",
                        "nodes move_node insert_node delete_node "
                        "set_editing set_segment",
                    ),
                ),
            ),
        ),
        "delete_first_section": Capability(
            com=None,
            rpc=None,
            js=Limitation("The first section requires move-before-delete"),
        ),
        "slideshow": Capability(
            com=None,
            rpc=None,
            js=Limitation(
                "macOS WPS slide show automation is unsafe "
                "in the tested build: "
                "exiting leaves invalid windows and can crash WPS. "
                "Use WPS playback controls; no show action was sent.",
                (
                    (
                        "wps_ppt_slideshow",
                        "start stop next previous goto status",
                    ),
                ),
            ),
        ),
        "smartart_nodes": Capability(
            com=None,
            rpc=None,
            js=Limitation(
                "macOS WPS does not expose SmartArt nodes; "
                "only diagram layout, color and style can be changed",
                (("wps_ppt_smartart", "add modify"),),
            ),
        ),
        "theme_color_assignment": Capability(
            com=None,
            rpc=None,
            js=Limitation(
                "Theme colors require a native scheme import named WPS MCP",
                (("wps_ppt_themes", "set_colors"),),
            ),
        ),
        "animation_direction": Capability(
            com=None,
            rpc=None,
            js=Limitation(
                "Animation direction metadata is null.",
                (("wps_ppt_animation", "add update list"),),
                ("Direction",),
            ),
        ),
        "title_master": Capability(
            com=None,
            rpc=None,
            js=Limitation(
                "has_title_master is null; unsafe getter is not queried.",
                (("wps_ppt_placeholders", "master"),),
                ("HasTitleMaster",),
            ),
        ),
        "transparency_color": Capability(
            com=None,
            rpc=None,
            js=Limitation(
                "Transparency color is null; unsafe getter is not queried.",
                (("wps_ppt_pictures", "format"),),
                ("TransparencyColor",),
            ),
        ),
        "theme_color_read": Capability(
            com=None,
            rpc=None,
            js=Limitation(
                "Replacements preserve readable colors; "
                "theme bindings are lost.",
                (("wps_et_conditional_formats", "update"),),
                ("ThemeColor", "PatternThemeColor"),
            ),
        ),
    }
)


@dataclass(frozen=True, slots=True)
class BackendCapabilities:
    """Query known binding policy without probing the native application."""

    platform: str
    backend: Backend
    reference_fallback: Literal["empty", "collections"]
    session_scope: str

    def limitation(self, feature: Feature) -> Limitation | None:
        """Reject unknown features instead of implicitly permitting them."""
        if feature not in _CAPABILITIES:
            raise ValueError(f"Unknown native capability: {feature}")
        return cast(
            Limitation | None, getattr(_CAPABILITIES[feature], self.backend)
        )

    def allows(self, feature: Feature) -> bool:
        """Permit native access without claiming the WPS build was verified."""
        return self.limitation(feature) is None

    def require(self, feature: Feature) -> None:
        """Reject a restricted operation before it touches native state."""
        restriction = self.limitation(feature)
        if restriction is not None:
            raise NotImplementedError(restriction.reason)

    def can_read(self, member: str) -> bool:
        """Avoid unsafe optional getters only on their affected backend."""
        return not any(
            member in restriction.members
            for feature in _CAPABILITIES
            if (restriction := self.limitation(feature)) is not None
        )

    def tool_limitations(self) -> list[JsonObject]:
        """Build discovery hints from the policies enforced at runtime."""
        return [
            {
                "tool": tool,
                "operation": operation,
                "reason": restriction.reason,
            }
            for feature in _CAPABILITIES
            if (restriction := self.limitation(feature)) is not None
            for tool, operations in restriction.tools
            for operation in operations.split()
        ]


_BACKENDS = {
    identifier: profile
    for profile in (
        BackendCapabilities(
            "win32",
            "com",
            "empty",
            "Existing or newly launched WPS instances",
        ),
        BackendCapabilities(
            "linux",
            "rpc",
            "collections",
            "MCP-managed WPS instances",
        ),
        BackendCapabilities(
            "darwin",
            "js",
            "empty",
            "Existing or newly launched WPS instances",
        ),
    )
    for identifier in (profile.platform, profile.backend)
}


def get_capabilities(identifier: str) -> BackendCapabilities:
    """Resolve a platform or backend through the same registered profile."""
    try:
        return _BACKENDS[identifier]
    except KeyError:
        raise WpsUnavailableError(
            f"Unsupported WPS platform or backend: {identifier}"
        ) from None


def validate_visibility(component: Component, visible: bool | None) -> None:
    """Enforce the shared application visibility rule at both boundaries."""
    if component == "wpp" and visible is False:
        raise ValueError("WPP cannot be hidden; use visible=true")
