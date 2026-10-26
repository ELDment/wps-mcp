"""Check Linux prerequisites and profile ownership without native WPS."""

import os
import stat
from importlib.metadata import PackageNotFoundError
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from wps_mcp.errors import WpsUnavailableError
from wps_mcp.native import discovery, linux, macos


@pytest.fixture
def linux_office(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Path:
    monkeypatch.setattr(
        discovery,
        "sys",
        SimpleNamespace(platform="linux", version_info=(3, 11, 17)),
    )
    monkeypatch.setattr(
        discovery,
        "platform",
        SimpleNamespace(
            python_implementation=lambda: "CPython", machine=lambda: "x86_64"
        ),
    )
    monkeypatch.setattr(
        discovery, "version", lambda name: discovery.RPC_VERSION
    )
    monkeypatch.setenv("DISPLAY", ":0")
    monkeypatch.setenv("WPS_OFFICE_DIR", str(tmp_path))
    for executable in ("wps", "et", "wpp"):
        path = tmp_path / executable
        path.write_bytes(b"native executable placeholder")
        path.chmod(0o700)
        (tmp_path / f"librpc{executable}api_wpsqt.so").write_bytes(b"library")
    return tmp_path


def test_linux_discovery_does_not_import_or_start_wps(
    monkeypatch: pytest.MonkeyPatch, linux_office: Path
) -> None:
    load = Mock(side_effect=AssertionError("must not import native libraries"))
    monkeypatch.setattr(discovery, "import_module", load)
    result = discovery.detect_components()
    assert result["backend"] == "rpc"
    assert result["session_scope"] == "MCP-managed WPS instances"
    assert result["limitations"][0]["tool"] == "wps_ppt_charts"
    assert result["limitations"][0]["operation"] == "set_data"
    assert len(result["components"]) == 3
    for component in result["components"]:
        assert component["registered"] is True
        assert component["ready"] is True
        assert component["issues"] == []
        assert component["native_execution_verified"] is False
        assert component["prog_id"] is None
        assert Path(component["executable"]).parent == linux_office
    load.assert_not_called()


def test_linux_discovery_distinguishes_installed_and_ready(
    monkeypatch: pytest.MonkeyPatch, linux_office: Path
) -> None:
    monkeypatch.delenv("DISPLAY")
    result = discovery.detect_components()
    assert all(item["registered"] for item in result["components"])
    assert not any(item["ready"] for item in result["components"])
    assert all("DISPLAY" in item["issues"][0] for item in result["components"])


def test_linux_requirements_report_missing_binding(
    monkeypatch: pytest.MonkeyPatch, linux_office: Path
) -> None:
    def missing(name: str) -> str:
        raise PackageNotFoundError(name)

    monkeypatch.setattr(discovery, "version", missing)
    directory, issues = discovery.linux_requirements("writer")
    assert directory == linux_office
    assert issues == ["Install pywpsrpc==2.4.0 in this interpreter"]


@pytest.mark.parametrize("installed", ["2.3.12", "2.4.1"])
def test_linux_requirements_reject_unvalidated_binding_versions(
    monkeypatch: pytest.MonkeyPatch, linux_office: Path, installed: str
) -> None:
    monkeypatch.setattr(discovery, "version", lambda name: installed)
    assert discovery.linux_requirements("writer")[1] == [
        "Install pywpsrpc==2.4.0 in this interpreter"
    ]


def test_linux_requirements_accumulate_runtime_and_component_failures(
    monkeypatch: pytest.MonkeyPatch, linux_office: Path
) -> None:
    monkeypatch.setattr(
        discovery,
        "sys",
        SimpleNamespace(platform="linux", version_info=(3, 12, 3)),
    )
    monkeypatch.setattr(
        discovery,
        "platform",
        SimpleNamespace(
            python_implementation=lambda: "CPython", machine=lambda: "aarch64"
        ),
    )
    (linux_office / "librpcetapi_wpsqt.so").unlink()
    monkeypatch.delenv("DISPLAY")
    _, issues = discovery.linux_requirements("et")
    assert len(issues) == 4
    assert any("CPython 3.11" in issue for issue in issues)
    assert any("x86_64" in issue for issue in issues)
    assert any("librpcetapi_wpsqt.so" in issue for issue in issues)
    assert any("DISPLAY" in issue for issue in issues)


def test_configured_office_directory_does_not_fall_back_silently(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setenv("WPS_OFFICE_DIR", str(tmp_path / "missing"))
    assert discovery.office_directory() is None


def test_missing_component_does_not_report_ready(linux_office: Path) -> None:
    (linux_office / "et").unlink()
    result = discovery.detect_components()
    components = {item["component"]: item for item in result["components"]}
    assert components["writer"]["ready"] is True
    assert components["et"]["registered"] is False
    assert components["et"]["ready"] is False
    assert components["et"]["issues"]


def test_unsupported_platform_reports_no_ready_components(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        discovery, "sys", SimpleNamespace(platform="unsupported")
    )
    result = discovery.detect_components()
    assert not any(item["ready"] for item in result["components"])
    assert all(item["issues"] for item in result["components"])


@pytest.fixture
def macos_office(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Path:
    monkeypatch.setattr(discovery, "sys", SimpleNamespace(platform="darwin"))
    bundle = tmp_path / "wpsoffice.app"
    executable = bundle / "Contents/MacOS/wpsoffice"
    executable.parent.mkdir(parents=True)
    executable.write_bytes(b"native executable placeholder")
    executable.chmod(0o700)
    for name in (
        "wpsdriverlocal",
        "wpsdriverbase",
        "wpsdriverkso",
        "wpsdriverwps",
        "wpsdriveret",
        "wpsdriverwpp",
    ):
        library = (
            bundle / f"Contents/Frameworks/office6/{name}.framework/{name}"
        )
        library.parent.mkdir(parents=True)
        library.write_bytes(b"native library placeholder")
    monkeypatch.setattr(discovery, "application_bundle", lambda: bundle)
    monkeypatch.setattr(
        discovery, "driver_directory", lambda: tmp_path / "container"
    )
    monkeypatch.setattr(discovery, "discover_sockets", Mock(return_value=[]))
    return bundle


def test_macos_discovery_is_ready_without_running_wps(
    monkeypatch: pytest.MonkeyPatch, macos_office: Path
) -> None:
    launch = Mock(side_effect=AssertionError("discovery must not launch WPS"))
    execute = Mock(side_effect=AssertionError("discovery must not execute JS"))
    monkeypatch.setattr(macos.subprocess, "run", launch)
    monkeypatch.setattr(macos.MacDriver, "execute", execute)
    result = discovery.detect_components()
    assert result["backend"] == "js"
    assert result["application"] == str(macos_office)
    assert result["editor_sockets"] == []
    for item in result["components"]:
        assert item["registered"] is True
        assert item["ready"] is True
        assert item["issues"] == []
        assert item["native_execution_verified"] is False
        assert item["prog_id"] is None
    launch.assert_not_called()
    execute.assert_not_called()


def test_macos_discovery_reports_component_specific_missing_driver(
    macos_office: Path,
) -> None:
    (
        macos_office
        / "Contents/Frameworks/office6/wpsdriveret.framework/wpsdriveret"
    ).unlink()
    result = discovery.detect_components()
    components = {item["component"]: item for item in result["components"]}
    assert components["et"]["registered"] is True
    assert components["et"]["ready"] is False
    assert "wpsdriveret" in components["et"]["issues"][0]
    assert components["writer"]["ready"] is True
    assert components["wpp"]["ready"] is True


def test_macos_discovery_separates_socket_presence_from_installation(
    monkeypatch: pytest.MonkeyPatch, macos_office: Path
) -> None:
    path = macos_office.parent / "container/wpsdriverlocal_wps_1"
    monkeypatch.setattr(
        discovery, "discover_sockets", Mock(return_value=[path])
    )
    monkeypatch.setattr(discovery, "application_bundle", lambda: None)
    result = discovery.detect_components()
    assert result["editor_sockets"] == [str(path)]
    assert result["application"] is None
    assert not any(item["registered"] for item in result["components"])
    assert not any(item["ready"] for item in result["components"])


def test_private_profile_changes_only_its_copy(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    config_root = tmp_path / "configuration"
    state_root = tmp_path / "state"
    original = config_root / "Kingsoft" / "Office.conf"
    original.parent.mkdir(parents=True)
    content = (
        b"[General]\r\nLanguage=zh_CN\r\n"
        b"[6.0]\r\ncommon\\AcceptedEULA=true\r\n"
        b"wpsoffice\\Application%20Settings\\AppComponentMode=prome_fushion\r\n"
        b"wpsoffice\\Application%20Settings\\AppComponentModeInstall=prome_fushion\r\n"
        b"[Other]\r\nUnrelated=value\r\n"
    )
    original.write_bytes(content)
    monkeypatch.setenv("XDG_CONFIG_HOME", str(config_root))
    monkeypatch.setenv("XDG_STATE_HOME", str(state_root))
    profile = linux.prepare_profile()
    assert original.read_bytes() == content
    assert profile.parent == state_root / "wps-mcp"
    assert os.environ["XDG_CONFIG_HOME"] == str(profile)
    copied = profile / "Kingsoft" / "Office.conf"
    rendered = copied.read_text(encoding="utf-8")
    assert rendered.count("AppComponentMode=prome_independ") == 1
    assert rendered.count("AppComponentModeInstall=prome_independ") == 1
    assert "prome_fushion" not in rendered
    assert "common\\AcceptedEULA=true" in rendered
    assert "[Other]\nUnrelated=value" in rendered
    if os.name != "nt":
        assert stat.S_IMODE(profile.stat().st_mode) == 0o700
        assert stat.S_IMODE(copied.stat().st_mode) == 0o600


@pytest.mark.parametrize(
    "original_text", [None, "[General]\nLanguage=zh_CN\n"]
)
def test_private_profile_adds_mode_without_accepting_license(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, original_text: str | None
) -> None:
    config_root = tmp_path / "configuration"
    if original_text is not None:
        original = config_root / "Kingsoft" / "Office.conf"
        original.parent.mkdir(parents=True)
        original.write_text(original_text, encoding="utf-8")
    monkeypatch.setenv("XDG_CONFIG_HOME", str(config_root))
    monkeypatch.setenv("XDG_STATE_HOME", str(tmp_path / "state"))
    profile = linux.prepare_profile()
    rendered = (profile / "Kingsoft" / "Office.conf").read_text(
        encoding="utf-8"
    )
    assert "[6.0]" in rendered
    assert "AppComponentMode=prome_independ" in rendered
    assert "AcceptedEULA" not in rendered


def test_rpc_rejects_missing_prerequisites_before_loading_bindings(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        linux, "linux_requirements", lambda component: (None, ["missing WPS"])
    )
    load = Mock(side_effect=AssertionError("must not load bindings"))
    monkeypatch.setattr(linux, "import_module", load)
    session = linux.RpcSession("writer")
    with pytest.raises(WpsUnavailableError, match="missing WPS"):
        session.connect()
    load.assert_not_called()
    assert session._profile is None
    assert session.close()


@pytest.mark.parametrize("open_documents", [0, 1])
def test_rpc_close_quits_and_removes_profile_only_when_empty(
    tmp_path: Path, open_documents: int
) -> None:
    profile = tmp_path / "profile"
    profile.mkdir()
    session = linux.RpcSession("writer")
    application = Mock()
    application.Documents.Count = open_documents
    session.application = application
    session._profile = profile
    assert session.close()
    assert session.close()
    if open_documents:
        application.Quit.assert_not_called()
        assert profile.is_dir()
    else:
        application.Quit.assert_called_once_with()
        assert not profile.exists()
    assert session.application is None


def test_rpc_uncertain_launch_retains_profile_and_restores_environment(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    previous = str(tmp_path / "original")
    monkeypatch.setenv("XDG_CONFIG_HOME", previous)
    session = linux.RpcSession("writer")
    profile = tmp_path / "profile"
    profile.mkdir()
    session._profile = profile
    session._launched = True
    monkeypatch.setenv("XDG_CONFIG_HOME", str(profile))
    assert session.close()
    assert profile.is_dir()
    assert os.environ["XDG_CONFIG_HOME"] == previous


@pytest.mark.parametrize("previous", [None, "old-config"])
def test_rpc_profile_cleanup_failure_still_restores_environment(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    previous: str | None,
) -> None:
    if previous is None:
        monkeypatch.delenv("XDG_CONFIG_HOME", raising=False)
    else:
        monkeypatch.setenv("XDG_CONFIG_HOME", previous)
    session = linux.RpcSession("writer")
    profile = tmp_path / "profile"
    profile.mkdir()
    session._profile = profile
    monkeypatch.setenv("XDG_CONFIG_HOME", str(profile))
    monkeypatch.setattr(
        linux.shutil, "rmtree", Mock(side_effect=PermissionError("busy"))
    )
    with pytest.raises(PermissionError, match="busy"):
        session.close()
    assert os.environ.get("XDG_CONFIG_HOME") == previous
    assert session.application is None


@pytest.mark.parametrize("configured", ["", "relative/config"])
def test_invalid_xdg_directories_use_defaults(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    configured: str,
) -> None:
    monkeypatch.setattr(Path, "home", lambda: tmp_path)
    monkeypatch.setenv("XDG_CONFIG_HOME", configured)
    monkeypatch.setenv("XDG_STATE_HOME", configured)
    profile = linux.prepare_profile()
    assert profile.parent == tmp_path / ".local/state/wps-mcp"
    assert profile.is_dir()


def test_failed_profile_construction_removes_partial_directory(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path / "config"))
    monkeypatch.setenv("XDG_STATE_HOME", str(tmp_path / "state"))
    monkeypatch.setattr(
        Path, "write_text", Mock(side_effect=OSError("write failed"))
    )
    with pytest.raises(OSError, match="write failed"):
        linux.prepare_profile()
    assert list((tmp_path / "state/wps-mcp").iterdir()) == []
    assert os.environ["XDG_CONFIG_HOME"] == str(tmp_path / "config")


@pytest.mark.parametrize(
    "variable", ["WPS_OFFICE_DIR", "WPS_DRIVER_DIR", "WPS_TEMPLATES_DIR"]
)
@pytest.mark.parametrize("configured", ["", "relative/path"])
def test_explicit_native_directories_require_absolute_local_paths(
    monkeypatch: pytest.MonkeyPatch,
    variable: str,
    configured: str,
) -> None:
    monkeypatch.setenv(variable, configured)
    resolve = {
        "WPS_OFFICE_DIR": discovery.office_directory,
        "WPS_DRIVER_DIR": discovery.driver_directory,
        "WPS_TEMPLATES_DIR": discovery.template_directory,
    }[variable]
    with pytest.raises(ValueError, match="absolute"):
        resolve()


def test_discovery_limitations_share_runtime_capability_reasons() -> None:
    from typing import get_args

    from wps_mcp.native.capabilities import Feature, get_capabilities

    for platform in ("win32", "linux", "darwin"):
        policy = get_capabilities(platform)
        reported = policy.tool_limitations()
        assert len(reported) == len(
            {(item["tool"], item["operation"]) for item in reported}
        )
        for feature in get_args(Feature):
            limitation = policy.limitation(feature)
            if limitation is None:
                continue
            for tool, operations in limitation.tools:
                for operation in operations.split():
                    assert {
                        "tool": tool,
                        "operation": operation,
                        "reason": limitation.reason,
                    } in reported, feature
