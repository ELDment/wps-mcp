"""Inspect platform prerequisites without importing or starting native WPS."""

import os
import platform
import re
import stat
import sys
from importlib import import_module
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

from src.common.paths import local_path
from src.models import Component, JsonObject
from src.native.components import SPECS

RPC_VERSION = "2.4.0"
EXECUTABLES: dict[Component, str] = {
    "writer": "wps",
    "et": "et",
    "wpp": "wpp",
}


MACOS_LIMITATIONS = (
    (
        "wps_ppt_animation",
        "add update list",
        "Animation direction metadata is null.",
    ),
    (
        "wps_et_conditional_formats",
        "update",
        "Replacements preserve readable colors; theme bindings are lost.",
    ),
    (
        "wps_ppt_charts",
        "set_data",
        "Embedded chart workbook writes are unavailable.",
    ),
    (
        "wps_ppt_charts",
        "add get_data set_type",
        "Series names are null; new charts may have empty cached labels.",
    ),
    ("wps_ppt_freeform", "build", "Only straight segments are supported."),
    (
        "wps_ppt_freeform",
        "nodes move_node insert_node delete_node set_editing set_segment",
        "Node operations are refused: native edits silently do nothing.",
    ),
    (
        "wps_ppt_slideshow",
        "start stop next previous goto status",
        "Disabled: native playback can leave invalid windows and crash WPS.",
    ),
    (
        "wps_ppt_smartart",
        "add modify",
        "Node editing is unavailable; node counts are null.",
    ),
    (
        "wps_ppt_connectors",
        "add format",
        "Use numeric connection sites; directions are unavailable.",
    ),
    (
        "wps_ppt_placeholders",
        "master",
        "has_title_master is null; unsafe getter is not queried.",
    ),
    (
        "wps_ppt_pictures",
        "format",
        "Transparency color is null; unsafe getter is not queried.",
    ),
    (
        "wps_ppt_themes",
        "set_colors",
        "WPS loads a generated scheme named WPS MCP.",
    ),
)


SOCKET_PATTERN = re.compile(r"wpsdriverlocal_(wps|et|wpp)_[0-9]+")


def driver_directory() -> Path:
    """Locate the application container, allowing an explicit path override."""
    configured = os.environ.get("WPS_DRIVER_DIR")
    if configured:
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


def template_directory(directory: str | None = None) -> tuple[Path, bool]:
    """Locate template roots, retaining explicit directory scan semantics."""
    bundled = False
    if directory is None:
        candidates = [os.environ.get("WPS_TEMPLATES_DIR")]
        app_data = os.environ.get("APPDATA")
        if app_data and sys.platform == "win32":
            candidates.append(
                str(Path(app_data) / "kingsoft" / "office6" / "templates")
            )
        if sys.platform == "linux":
            data_home = os.environ.get("XDG_DATA_HOME")
            data_root = (
                Path(data_home)
                if data_home and Path(data_home).is_absolute()
                else Path.home() / ".local" / "share"
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
    root = Path(directory).resolve()
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
        "backend": {"linux": "rpc", "win32": "com", "darwin": "js"}.get(
            sys.platform
        ),
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
            else [
                {"tool": name, "operation": operation, "reason": reason}
                for name, operations, reason in MACOS_LIMITATIONS
                for operation in operations.split()
            ]
            if sys.platform == "darwin"
            else []
        ),
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
