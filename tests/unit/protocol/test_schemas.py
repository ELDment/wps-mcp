"""Published schemas retain constraints without client reference resolution."""

from copy import deepcopy

import pytest
from jsonschema import Draft202012Validator
from pydantic import ValidationError

from tests.schemas import assert_inline_schema
from wps_mcp.models import JsonObject
from wps_mcp.protocol.catalog import inline_schema
from wps_mcp.protocol.server import WpsServer, build_server


@pytest.fixture(scope="module")
def server() -> WpsServer:
    return build_server()


def test_published_schemas_preserve_all_operation_constraints(
    server: WpsServer,
) -> None:
    for tool in server._tool_manager.list_tools():
        original = tool.fn_metadata.arg_model.model_json_schema()
        published = tool.parameters
        Draft202012Validator.check_schema(published)
        assert_inline_schema(published)
        params = published["properties"]["params"]
        variants = params.get("oneOf", [params])
        for variant in variants:
            source = original["$defs"][variant["title"]]
            assert variant["required"] == source["required"]
            assert variant["additionalProperties"] is False
            assert variant["properties"].keys() == source["properties"].keys()
            assert (
                variant["properties"]["operation"]
                == (source["properties"]["operation"])
            )
        # Publishing a schema must not replace the runtime's typed model.
        assert "$defs" in original


@pytest.mark.parametrize(
    ("name", "params", "valid"),
    [
        ("wps_components", {"operation": "detect"}, True),
        (
            "wps_components",
            {"operation": "connect", "component": "writer"},
            True,
        ),
        ("wps_components", {"operation": "connect"}, False),
        ("wps_components", {"operation": "missing"}, False),
        (
            "wps_components",
            {"operation": "detect", "component": "writer"},
            False,
        ),
        (
            "wps_components",
            {"operation": "connect", "component": "missing"},
            False,
        ),
        (
            "wps_components",
            {
                "operation": "connect",
                "component": "writer",
                "visible": "false",
            },
            False,
        ),
        ("wps_writer_statistics", {"operation": "read"}, True),
        ("wps_writer_statistics", {"operation": "read", "typo": True}, False),
    ],
)
def test_published_schema_accepts_the_same_calls(
    server: WpsServer, name: str, params: JsonObject, valid: bool
) -> None:
    tool = server._tool_manager.get_tool(name)
    assert tool is not None
    original = tool.fn_metadata.arg_model.model_json_schema()
    for schema in (original, tool.parameters):
        assert (
            Draft202012Validator(schema).is_valid({"params": params}) is valid
        )


@pytest.mark.parametrize(
    ("change", "valid"),
    [
        ({"kind": "glow", "options": {"radius": 2, "color": "#fff"}}, True),
        ({"kind": "glow", "options": {"radius": -1}}, False),
        ({"kind": "glow", "options": {}}, False),
        ({"kind": "glow", "options": {"radius": 2, "color": "red"}}, False),
        ({"kind": "text", "options": {"bold": None}}, True),
        ({"kind": "text", "options": {"bold": "true"}}, False),
        ({"kind": "missing", "options": {}}, False),
    ],
)
def test_nested_style_schema_retains_validation(
    server: WpsServer, change: JsonObject, valid: bool
) -> None:
    tool = server._tool_manager.get_tool("wps_ppt_style")
    assert tool is not None
    arguments = {
        "params": {
            "operation": "apply",
            "slide_index": 1,
            "shapes": ["Title"],
            "operations": [change],
        }
    }
    original = tool.fn_metadata.arg_model.model_json_schema()
    for schema in (original, tool.parameters):
        assert Draft202012Validator(schema).is_valid(arguments) is valid
    if valid:
        tool.fn_metadata.validate_arguments(arguments)
    else:
        with pytest.raises(ValidationError):
            tool.fn_metadata.validate_arguments(arguments)


@pytest.mark.parametrize("keyword", ["default", "const", "enum", "examples"])
def test_schema_expansion_preserves_literal_data_and_source(
    keyword: str,
) -> None:
    literal = {"$ref": "literal", "$defs": {}, "discriminator": "data"}
    value = [literal] if keyword in {"enum", "examples"} else literal
    schema = {
        "$defs": {"Value": {"type": "object", keyword: value}},
        "type": "object",
        "properties": {
            "$ref": {"$ref": "#/$defs/Value"},
            "$defs": {"$ref": "#/$defs/Value"},
            "discriminator": {"$ref": "#/$defs/Value"},
        },
    }
    original = deepcopy(schema)
    expanded = inline_schema(schema)
    assert schema == original
    assert expanded["properties"] == {
        "$ref": {"type": "object", keyword: value},
        "$defs": {"type": "object", keyword: value},
        "discriminator": {"type": "object", keyword: value},
    }
    assert_inline_schema(expanded)


@pytest.mark.parametrize(
    "schema",
    [
        {"$defs": {}},
        {"properties": {"value": {"$ref": "#/$defs/Value"}}},
        {
            "items": {
                "oneOf": [{"discriminator": {"propertyName": "operation"}}]
            }
        },
    ],
)
def test_inline_schema_check_rejects_schema_keywords(
    schema: JsonObject,
) -> None:
    with pytest.raises(AssertionError):
        assert_inline_schema(schema)


@pytest.mark.parametrize(
    ("schema", "message"),
    [
        ({"$ref": "https://example.com/schema"}, "bare local"),
        ({"$ref": "#/$defs/Missing"}, "Unknown"),
        ({"$ref": "#/$defs/Value", "maxLength": 3}, "bare local"),
        (
            {
                "$defs": {"Loop": {"$ref": "#/$defs/Loop"}},
                "properties": {"value": {"$ref": "#/$defs/Loop"}},
            },
            "Recursive",
        ),
    ],
)
def test_unsupported_schema_references_fail_explicitly(
    schema: JsonObject, message: str
) -> None:
    with pytest.raises(ValueError, match=message):
        inline_schema(schema)
