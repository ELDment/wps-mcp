"""Keep infrastructure and domain foundations free of tool dependencies."""

import ast
from pathlib import Path

import pytest

PACKAGE = Path(__file__).resolve().parents[2] / "src" / "wps_mcp"


@pytest.mark.parametrize(
    ("directory", "forbidden"),
    [
        (
            "common",
            (
                "mcp",
                "wps_mcp.native",
                "wps_mcp.protocol",
                "wps_mcp.documents",
                "wps_mcp.writer",
                "wps_mcp.spreadsheet",
                "wps_mcp.presentation",
            ),
        ),
        (
            "native",
            (
                "mcp",
                "wps_mcp.protocol",
                "wps_mcp.documents",
                "wps_mcp.writer",
                "wps_mcp.spreadsheet",
                "wps_mcp.presentation",
            ),
        ),
        (
            "documents/service.py",
            ("mcp", "wps_mcp.protocol", "wps_mcp.presentation"),
        ),
        (
            "documents",
            (
                "wps_mcp.native.javascript",
                "wps_mcp.native.macos",
                "wps_mcp.native.rpc",
            ),
        ),
        (
            "writer",
            (
                "wps_mcp.native.javascript",
                "wps_mcp.native.macos",
                "wps_mcp.native.rpc",
            ),
        ),
        (
            "spreadsheet",
            (
                "wps_mcp.native.javascript",
                "wps_mcp.native.macos",
                "wps_mcp.native.rpc",
            ),
        ),
        (
            "presentation",
            (
                "wps_mcp.native.javascript",
                "wps_mcp.native.macos",
                "wps_mcp.native.rpc",
            ),
        ),
        (
            "protocol/definitions.py",
            ("mcp", "wps_mcp.native", "wps_mcp.documents"),
        ),
    ],
)
def test_infrastructure_imports_follow_layer_boundaries(
    directory: str, forbidden: tuple[str, ...]
) -> None:
    target = PACKAGE / directory
    paths = [target] if target.is_file() else sorted(target.rglob("*.py"))
    assert paths
    violations = []
    for path in paths:
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                imports = [alias.name for alias in node.names]
            elif isinstance(node, ast.ImportFrom) and node.module:
                imports = [node.module]
            else:
                continue
            for imported in imports:
                if any(
                    imported == prefix or imported.startswith(prefix + ".")
                    for prefix in forbidden
                ):
                    violations.append(
                        f"{path.relative_to(PACKAGE)}:{node.lineno}: "
                        f"{imported}"
                    )
    assert not violations, "\n".join(violations)


@pytest.mark.parametrize(
    "domain", ["documents", "writer", "spreadsheet", "presentation"]
)
def test_domain_rules_do_not_inspect_the_host_platform(domain: str) -> None:
    forbidden = {
        "sys.platform",
        "os.name",
        "platform.system",
        "platform.platform",
    }
    violations = []
    for path in (PACKAGE / domain).rglob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        aliases = {}
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                aliases.update(
                    (item.asname or item.name, item.name)
                    for item in node.names
                )
            elif isinstance(node, ast.ImportFrom) and node.module:
                aliases.update(
                    (item.asname or item.name, f"{node.module}.{item.name}")
                    for item in node.names
                )
        for node in ast.walk(tree):
            if isinstance(node, ast.Name):
                name = aliases.get(node.id, node.id)
            elif isinstance(node, ast.Attribute) and isinstance(
                node.value, ast.Name
            ):
                name = (
                    aliases.get(node.value.id, node.value.id) + "." + node.attr
                )
            else:
                continue
            if name in forbidden:
                violations.append(
                    f"{path.relative_to(PACKAGE)}:{node.lineno}: {name}"
                )
    assert not violations, "\n".join(violations)
