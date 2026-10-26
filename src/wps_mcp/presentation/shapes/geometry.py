"""Select geometry update targets and apply validated coordinates."""

from __future__ import annotations

from wps_mcp.models import NativeObject
from wps_mcp.presentation.shapes.models import UpdateShapeInput


def select_targets(
    available: list[str],
    shape_names: list[str] | None,
    all_shapes: bool,
    exclude: list[str] | None,
) -> list[int | str]:
    """Work out which shapes an update should touch."""
    excluded = set(exclude or ())
    if all_shapes:
        return [i for i, name in enumerate(available) if name not in excluded]
    first_at: dict[str, int] = {}
    for i, name in enumerate(available):
        first_at.setdefault(name, i)
    picked = []
    for name in shape_names or ():
        if name in excluded:
            continue
        picked.append(first_at.get(name, name))
    return picked


def apply_geometry(shape: NativeObject, params: UpdateShapeInput) -> None:
    """Apply validated absolute coordinates and relative offsets."""
    for field in ("left", "top", "width", "height", "rotation"):
        value = getattr(params, field)
        if value is not None:
            setattr(shape, field.title(), value)
    for field in ("left", "top", "width", "height"):
        offset = getattr(params, "d" + field)
        if offset is not None:
            setattr(
                shape, field.title(), getattr(shape, field.title()) + offset
            )
