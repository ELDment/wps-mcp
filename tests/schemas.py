"""Assertions for published JSON Schema nodes, excluding literal data."""

from wps_mcp.models import JsonObject


def assert_inline_schema(schema: JsonObject | bool) -> None:
    """Reject references and discriminators only at schema positions."""
    if isinstance(schema, bool):
        return
    assert not {"$ref", "$defs", "discriminator"}.intersection(schema), schema
    for key in ("properties", "patternProperties", "dependentSchemas"):
        for child in schema.get(key, {}).values():
            assert_inline_schema(child)
    for key in (
        "items",
        "additionalProperties",
        "contains",
        "propertyNames",
        "not",
        "if",
        "then",
        "else",
        "unevaluatedProperties",
        "unevaluatedItems",
        "contentSchema",
    ):
        if key in schema:
            assert_inline_schema(schema[key])
    for key in ("oneOf", "anyOf", "allOf", "prefixItems"):
        for child in schema.get(key, []):
            assert_inline_schema(child)
