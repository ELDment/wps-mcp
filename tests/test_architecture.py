"""Keep infrastructure and domain foundations free of tool dependencies."""

import ast
from pathlib import Path

import pytest

PACKAGE = Path(__file__).resolve().parents[1] / "src"


@pytest.mark.parametrize(
    ("directory", "forbidden"),
    [
        (
            "common",
            (
                "mcp",
                "src.native",
                "src.protocol",
                "src.documents",
                "src.writer",
                "src.spreadsheet",
                "src.presentation",
            ),
        ),
        (
            "native",
            (
                "mcp",
                "src.protocol",
                "src.documents",
                "src.writer",
                "src.spreadsheet",
                "src.presentation",
            ),
        ),
        ("documents/service.py", ("mcp", "src.protocol")),
        (
            "documents",
            ("src.native.javascript", "src.native.macos", "src.native.rpc"),
        ),
        (
            "writer",
            ("src.native.javascript", "src.native.macos", "src.native.rpc"),
        ),
        (
            "spreadsheet",
            ("src.native.javascript", "src.native.macos", "src.native.rpc"),
        ),
        (
            "presentation",
            ("src.native.javascript", "src.native.macos", "src.native.rpc"),
        ),
        (
            "protocol/definitions.py",
            ("mcp", "src.native", "src.documents"),
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
