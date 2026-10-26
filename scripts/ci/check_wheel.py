"""Install and exercise the built wheel outside the source checkout."""

import argparse
import json
import os
import subprocess
import sys
from email.parser import BytesParser
from importlib import import_module
from importlib.metadata import distribution
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile

import anyio
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from mcp.types import TextContent


async def check_stdio(expected_version: str, *, console: bool) -> None:
    """Initialize a real stdio server without starting native applications."""
    executable = Path(sys.executable)
    command = (
        str(
            executable.with_name(
                "wps-mcp.exe" if os.name == "nt" else "wps-mcp"
            )
        )
        if console
        else str(executable)
    )
    parameters = StdioServerParameters(
        command=command,
        args=[] if console else ["-I", "-m", "wps_mcp"],
        cwd=Path.cwd(),
        env=dict(os.environ),
    )
    with anyio.fail_after(30):
        async with stdio_client(parameters) as (read, write):
            async with ClientSession(read, write) as client:
                initialized = await client.initialize()
                assert initialized.server_info.version == expected_version
                catalog = await client.list_tools()
                assert len(catalog.tools) == 45
                detected = await client.call_tool(
                    "wps_components", {"params": {"operation": "detect"}}
                )
                assert not detected.is_error
                assert detected.structured_content is not None
                rejected = await client.call_tool(
                    "wps_documents",
                    {
                        "params": {
                            "operation": "create",
                            "component": "writer",
                            "typo": True,
                        }
                    },
                )
                assert rejected.is_error
                hidden = await client.call_tool(
                    "wps_components",
                    {
                        "params": {
                            "operation": "connect",
                            "component": "wpp",
                            "visible": False,
                        }
                    },
                )
                assert hidden.is_error
                assert any(
                    isinstance(content, TextContent)
                    and "WPP cannot be hidden; use visible=true"
                    in content.text
                    for content in hidden.content
                )


def check_installed(expected_version: str) -> None:
    """Prove imports and both entry points come from the installed wheel."""
    package = import_module("wps_mcp")
    metadata = distribution("wps-mcp")
    prefix = Path(sys.prefix).resolve()
    assert package.__file__ is not None
    assert Path(package.__file__).resolve().is_relative_to(prefix)
    assert Path(str(metadata.locate_file(""))).resolve().is_relative_to(prefix)
    assert package.__version__ == metadata.version == expected_version
    direct_url = json.loads(metadata.read_text("direct_url.json") or "{}")
    assert "archive_info" in direct_url
    assert not direct_url.get("dir_info", {}).get("editable", False)
    assert direct_url["url"].endswith(".whl")
    catalog = import_module("wps_mcp.protocol.catalog")
    definitions = import_module("wps_mcp.protocol.definitions")
    catalog.load_definitions()
    assert len(definitions.DEFINITIONS) == 252

    async def scenario() -> None:
        await check_stdio(expected_version, console=False)
        await check_stdio(expected_version, console=True)

    anyio.run(scenario)
    print(
        f"Wheel {metadata.version}: isolated module and console checks passed"
    )


def main() -> None:
    """Create a clean environment, install one wheel, and verify its entry."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--wheel", type=Path)
    parser.add_argument("--installed-version", help=argparse.SUPPRESS)
    arguments = parser.parse_args()
    if arguments.installed_version is not None:
        check_installed(arguments.installed_version)
        return
    wheels = (
        [arguments.wheel]
        if arguments.wheel
        else sorted(Path("dist").glob("*.whl"))
    )
    if len(wheels) != 1:
        parser.error(
            "Specify --wheel when dist does not contain exactly one wheel"
        )
    wheel = wheels[0].resolve(strict=True)
    with ZipFile(wheel) as archive:
        records = [
            name
            for name in archive.namelist()
            if name.endswith(".dist-info/METADATA")
        ]
        if len(records) != 1:
            raise ValueError(
                "The wheel must contain exactly one metadata record"
            )
        metadata = BytesParser().parsebytes(archive.read(records[0]))
        if metadata["Name"] != "wps-mcp" or not metadata["Version"]:
            raise ValueError("Expected a wps-mcp wheel with a version")
        expected_version = metadata["Version"]
    environment = {
        key: value
        for key, value in os.environ.items()
        if key not in {"PYTHONPATH", "PYTHONHOME", "VIRTUAL_ENV"}
    }
    environment["PYTHONDONTWRITEBYTECODE"] = "1"
    with TemporaryDirectory(prefix="wps-wheel-check-") as directory:
        root = Path(directory)
        venv = root / "environment"
        subprocess.run(
            [
                "uv",
                "venv",
                "--no-config",
                "--python",
                sys.executable,
                str(venv),
            ],
            cwd=root,
            env=environment,
            check=True,
            timeout=60,
        )
        python = venv / (
            "Scripts/python.exe" if os.name == "nt" else "bin/python"
        )
        subprocess.run(
            [
                "uv",
                "pip",
                "install",
                "--no-config",
                "--python",
                str(python),
                str(wheel),
            ],
            cwd=root,
            env=environment,
            check=True,
            timeout=300,
        )
        subprocess.run(
            [
                str(python),
                "-I",
                str(Path(__file__).resolve()),
                "--installed-version",
                expected_version,
            ],
            cwd=root,
            env=environment,
            check=True,
            timeout=90,
        )


if __name__ == "__main__":
    main()
