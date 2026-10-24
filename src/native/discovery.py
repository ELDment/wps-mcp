"""Inspect platform prerequisites without importing or starting native WPS."""

import os
import platform
import sys
from importlib import import_module
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

from src.models import Component, JsonObject
from src.native.components import SPECS

RPC_VERSION = "2.4.0"
EXECUTABLES: dict[Component, str] = {
    "writer": "wps",
    "et": "et",
    "wpp": "wpp",
}


def office_directory() -> Path | None:
    """Locate native executables rather than shell launchers."""
    configured = os.environ.get("WPS_OFFICE_DIR")
    candidates = (
        [Path(configured)]
        if configured
        else [
            Path("/opt/kingsoft/wps-office/office6"),
            Path("/usr/lib/wps-office/office6"),
        ]
    )
    return next(
        (path for path in candidates if (path / "wps").is_file()), None
    )


def linux_requirements(component: Component) -> tuple[Path | None, list[str]]:
    """Explain unmet prerequisites before loading unsafe native bindings."""
    directory = office_directory()
    issues = []
    if platform.python_implementation() != "CPython" or (
        sys.version_info[:2] != (3, 11)
    ):
        issues.append("Use CPython 3.11: uvx --python 3.11 wps-mcp")
    if platform.machine().lower() not in {"x86_64", "amd64"}:
        issues.append("Linux RPC currently supports x86_64 only")
    try:
        installed = version("pywpsrpc")
    except PackageNotFoundError:
        installed = None
    if installed != RPC_VERSION:
        issues.append(f"Install pywpsrpc=={RPC_VERSION} in this interpreter")
    executable = EXECUTABLES[component]
    if directory is None or not os.access(directory / executable, os.X_OK):
        issues.append("Install WPS Office or set WPS_OFFICE_DIR to office6")
    elif not (directory / f"librpc{executable}api_wpsqt.so").is_file():
        issues.append(f"WPS is missing librpc{executable}api_wpsqt.so")
    if not os.environ.get("DISPLAY"):
        issues.append("Run in a desktop session with DISPLAY and X authority")
    return directory, issues


def detect_components() -> JsonObject:
    """Report installation and execution prerequisites separately."""
    components = []
    for component, spec in SPECS.items():
        registered = False
        issues: list[str] = []
        executable: str | None = None
        if sys.platform == "win32":
            registry = import_module("winreg")
            try:
                with registry.OpenKey(
                    registry.HKEY_CLASSES_ROOT, spec.prog_id + "\\CLSID"
                ):
                    registered = True
            except FileNotFoundError:
                pass
        elif sys.platform == "linux":
            directory, issues = linux_requirements(component)
            if directory is not None:
                path = directory / EXECUTABLES[component]
                registered = path.is_file()
                executable = str(path)
        else:
            issues.append("Supported platforms are Windows and Linux")
        components.append(
            {
                "component": component,
                "prog_id": spec.prog_id if sys.platform == "win32" else None,
                "registered": registered,
                "executable": executable,
                "formats": list(spec.formats),
                "ready": registered and not issues,
                "issues": issues,
                "native_execution_verified": False,
            }
        )
    return {
        "platform": sys.platform,
        "backend": {"linux": "rpc", "win32": "com"}.get(sys.platform),
        "components": components,
        "session_scope": (
            "MCP-managed WPS instances"
            if sys.platform == "linux"
            else "Existing or newly launched WPS instances"
        ),
        "limitations": (
            [
                {
                    "tool": "wps_ppt_charts",
                    "operation": "set_data",
                    "reason": (
                        "The Linux RPC embedded chart workbook interface "
                        "is unavailable; data writes are refused."
                    ),
                }
            ]
            if sys.platform == "linux"
            else []
        ),
        "note": (
            "Discovery does not verify every feature. Linux uses an isolated "
            "multi-component profile; existing windows are not adopted."
        ),
    }
