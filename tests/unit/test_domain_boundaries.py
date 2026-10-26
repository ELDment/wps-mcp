"""Verify shared domain foundations have no tool-registration side effects."""

import os
import subprocess
import sys
from pathlib import Path

import pytest


@pytest.mark.parametrize(
    ("domain", "modules"),
    [
        ("writer", ("models", "access", "formatting")),
        (
            "spreadsheet",
            (
                "models",
                "addresses",
                "matrices",
                "access",
                "object_names",
                "table_access",
                "formulas",
            ),
        ),
    ],
)
def test_foundations_import_without_registering_or_starting_native_tools(
    domain: str, modules: tuple[str, ...]
) -> None:
    source_root = Path(__file__).resolve().parents[2] / "src"
    script = "\n".join(
        [
            "from importlib import import_module",
            "from wps_mcp.protocol.definitions import DEFINITIONS",
            "from wps_mcp.native.context import CURRENT_RUNTIME",
            f"modules = {modules!r}",
            "for module in modules:",
            f"    import_module('wps_mcp.{domain}.' + module)",
            "assert not DEFINITIONS, sorted(DEFINITIONS)",
            "assert CURRENT_RUNTIME.get() is None",
        ]
    )
    result = subprocess.run(
        [sys.executable, "-c", script],
        env=os.environ | {"PYTHONPATH": str(source_root)},
        capture_output=True,
        text=True,
        timeout=20,
        check=False,
    )
    assert result.returncode == 0, result.stdout + result.stderr
