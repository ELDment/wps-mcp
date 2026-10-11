"""Resolve shapes and group paths without ambiguous writes."""

from collections.abc import Iterator, Sequence

from src.models import NativeObject
from src.presentation.constants import MSO_GROUP

PATH_SEPARATOR = "/"


class ShapeNotFoundError(ValueError):
    """No matching shape exists on the slide."""


class AmbiguousShapeError(ValueError):
    """Several group members share the requested name."""


def _direct_children(shape: NativeObject) -> list[NativeObject]:
    """Read direct group members, propagating collection errors."""
    if shape.Type != MSO_GROUP:
        return []
    items = shape.GroupItems
    return [items(index) for index in range(1, items.Count + 1)]


def walk_group_children(
    shape: NativeObject, prefix: str = ""
) -> Iterator[tuple[NativeObject, str]]:
    """Yield descendants and full paths in depth-first order."""
    for child in _direct_children(shape):
        path = prefix + child.Name
        yield child, path
        yield from walk_group_children(child, path + PATH_SEPARATOR)


def _match_path(
    candidates: Sequence[NativeObject], name: str
) -> tuple[NativeObject, str] | None:
    """Match exact names first, including any path separators."""
    for shape in candidates:
        if shape.Name == name:
            return shape, name
    for shape in candidates:
        prefix = shape.Name + PATH_SEPARATOR
        if name.startswith(prefix):
            match = _match_path(_direct_children(shape), name[len(prefix) :])
            if match is not None:
                child, path = match
                return child, prefix + path
    return None


def group_names(slide: NativeObject) -> list[str]:
    """Return top-level group names for discovery and error messages."""
    return [
        slide.Shapes(index).Name
        for index in range(1, slide.Shapes.Count + 1)
        if slide.Shapes(index).Type == MSO_GROUP
    ]


def resolve_with_path(
    slide: NativeObject, name_or_index: str | int
) -> tuple[NativeObject, str]:
    """Resolve an index, exact name, unique child name or group path."""
    if isinstance(name_or_index, bool):
        raise ShapeNotFoundError(
            "A shape index must be an integer, not a boolean"
        )
    if isinstance(name_or_index, int):
        if not 1 <= name_or_index <= slide.Shapes.Count:
            raise ShapeNotFoundError(
                f"Shape index {name_or_index} out "
                f"of range (1-{slide.Shapes.Count})"
            )
        shape = slide.Shapes(name_or_index)
        return shape, shape.Name
    shapes = [
        slide.Shapes(index) for index in range(1, slide.Shapes.Count + 1)
    ]
    for shape in shapes:
        if shape.Name == name_or_index:
            return shape, shape.Name
    matches = [
        (child, path)
        for shape in shapes
        for child, path in walk_group_children(
            shape, shape.Name + PATH_SEPARATOR
        )
        if child.Name == name_or_index
    ]
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        paths = ", ".join(sorted(path for _, path in matches))
        raise AmbiguousShapeError(
            f"Shape '{name_or_index}' matches several group members: {paths}. "
            "Use a complete group path."
        )
    if PATH_SEPARATOR in name_or_index:
        match = _match_path(shapes, name_or_index)
        if match is not None:
            return match
    groups = group_names(slide)
    hint = (
        f". Inspect groups with "
        f"wps_ppt_shapes/group_items: {', '.join(groups)}"
        if groups
        else ""
    )
    raise ShapeNotFoundError(
        f"Shape '{name_or_index}' not found on slide{hint}"
    )


def resolve_shape(
    slide: NativeObject, name_or_index: str | int
) -> NativeObject:
    """Resolve a native shape by name, group path or index."""
    return resolve_with_path(slide, name_or_index)[0]


def require_top_level(
    slide: NativeObject, names: Sequence[str], slide_index: int, tool_name: str
) -> None:
    """Reject nested shapes before native top-level operations."""
    top_level = {
        slide.Shapes(index).Name for index in range(1, slide.Shapes.Count + 1)
    }
    for name in names:
        if name in top_level:
            continue
        try:
            _, path = resolve_with_path(slide, name)
        except ShapeNotFoundError as error:
            raise ValueError(
                f"Shape '{name}' not found on slide {slide_index}"
            ) from error
        raise ValueError(
            f"Shape '{name}' is inside group path '{path}'. "
            f"{tool_name} requires "
            "top-level shapes; ungroup it first with wps_ppt_shapes/ungroup."
        )
