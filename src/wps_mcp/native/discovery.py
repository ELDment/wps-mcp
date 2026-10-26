"""Inspect platform prerequisites without importing or starting native WPS."""

import os
import platform
import re
import stat
import sys
from importlib import import_module
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

from wps_mcp.common.paths import local_path
from wps_mcp.errors import WpsUnavailableError
from wps_mcp.models import Component, JsonObject
from wps_mcp.native.capabilities import get_capabilities
from wps_mcp.native.components import SPECS

RPC_VERSION = "2.4.0"
EXECUTABLES: dict[Component, str] = {
    "writer": "wps",
    "et": "et",
    "wpp": "wpp",
}


SOCKET_PATTERN = re.compile(r"wpsdriverlocal_(wps|et|wpp)_[0-9]+")


def driver_directory() -> Path:
    """Locate the application container, allowing an explicit path override."""
    configured = os.environ.get("WPS_DRIVER_DIR")
    if configured is not None:
        return local_path(configured)
    return Path.home() / "Library/Containers/com.kingsoft.wpsoffice.mac/Data"


def application_bundle() -> Path | None:
    """Find the installed WPS bundle without launching it."""
    for parent in (Path("/Applications"), Path.home() / "Applications"):
        candidate = parent / "wpsoffice.app"
        if (candidate / "Contents/Info.plist").is_file():
            return candidate
    return None


def discover_sockets(directory: Path) -> list[Path]:
    """Find editor endpoints; the home-page endpoint cannot run scripts."""
    result = []
    for path in sorted(directory.glob("wpsdriverlocal_*")):
        if not SOCKET_PATTERN.fullmatch(path.name):
            continue
        try:
            if stat.S_ISSOCK(path.lstat().st_mode):
                result.append(path)
        except FileNotFoundError:
            continue
    return result


def office_directory() -> Path | None:
    """Locate native executables rather than shell launchers."""
    configured = os.environ.get("WPS_OFFICE_DIR")
    candidates = (
        [local_path(configured)]
        if configured is not None
        else [
            Path("/opt/kingsoft/wps-office/office6"),
            Path("/usr/lib/wps-office/office6"),
        ]
    )
    return next(
        (path for path in candidates if (path / "wps").is_file()), None
    )


def xdg_directory(variable: str, fallback: Path) -> Path:
    """Use an absolute XDG override, ignoring empty and relative values."""
    configured = os.environ.get(variable)
    if configured and Path(configured).is_absolute():
        return Path(configured)
    return fallback


def template_directory(directory: str | None = None) -> tuple[Path, bool]:
    """Locate template roots, retaining explicit directory scan semantics."""
    bundled = False
    if directory is None:
        directory = os.environ.get("WPS_TEMPLATES_DIR")
    if directory is None:
        candidates: list[str] = []
        app_data = os.environ.get("APPDATA")
        if app_data and sys.platform == "win32":
            candidates.append(
                str(Path(app_data) / "kingsoft" / "office6" / "templates")
            )
        if sys.platform == "linux":
            data_root = xdg_directory(
                "XDG_DATA_HOME", Path.home() / ".local" / "share"
            )
            candidates.append(
                str(data_root / "Kingsoft" / "office6" / "templates")
            )
            candidates.append("/opt/kingsoft/wps-office/templates")
        directory = next(
            (
                candidate
                for candidate in candidates
                if candidate and Path(candidate).is_dir()
            ),
            None,
        )
        if directory is None and sys.platform == "darwin":
            bundle = application_bundle()
            if bundle is not None:
                root = bundle / "Contents/Resources/office6/mui"
                if any(root.glob("*/templates")):
                    directory = str(root)
                    bundled = True
    if directory is None:
        raise FileNotFoundError(
            "No templates directory found; specify templates_dir"
        )
    root = local_path(directory)
    if not root.is_dir():
        raise NotADirectoryError(f"Templates directory not found: {root}")
    return root, bundled


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
    try:
        capabilities = get_capabilities(sys.platform)
    except WpsUnavailableError:
        capabilities = None
    bundle: Path | None = None
    macos_info: JsonObject = {}
    if sys.platform == "darwin":
        bundle = application_bundle()
        directory = driver_directory()
        macos_info = {
            "application": str(bundle) if bundle else None,
            "driver_directory": str(directory),
            "editor_sockets": [
                str(path) for path in discover_sockets(directory)
            ],
        }
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
        elif sys.platform == "darwin":
            if bundle is None:
                issues.append("Install WPS Office in Applications")
            else:
                path = bundle / "Contents/MacOS/wpsoffice"
                registered = path.is_file()
                executable = str(path)
                if not os.access(path, os.X_OK):
                    issues.append("WPS application executable is unavailable")
                frameworks = bundle / "Contents/Frameworks/office6"
                for name in (
                    "wpsdriverlocal",
                    "wpsdriverbase",
                    "wpsdriverkso",
                    "wpsdriver" + EXECUTABLES[component],
                ):
                    if not (frameworks / f"{name}.framework" / name).is_file():
                        issues.append(
                            f"WPS is missing {name}.framework/{name}"
                        )
        else:
            issues.append("Supported platforms are Windows, Linux and macOS")
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
        "backend": capabilities.backend if capabilities is not None else None,
        "components": components,
        "session_scope": (
            capabilities.session_scope
            if capabilities is not None
            else "Unavailable"
        ),
        "limitations": capabilities.tool_limitations()
        if capabilities is not None
        else [],
        "note": (
            "Discovery does not verify every feature. Linux uses an isolated "
            "multi-component profile; existing windows are not adopted. "
            "macOS uses the built-in JS driver; editor sockets appear after "
            "an editor is loaded. Ready does not require an open document. "
            "Cross-component calls may "
            "activate an editor window."
        ),
        **macos_info,
    }
