"""Adapt pywpsrpc 2.4 objects to the shared native object contract."""

from __future__ import annotations

import inspect
import re
from collections.abc import Callable, Iterator
from dataclasses import dataclass
from datetime import date, datetime, time
from functools import lru_cache
from importlib import import_module
from typing import Any

from wps_mcp.errors import NativeError

HRESULT_PATTERN = re.compile(r"(?:with|:)\s*(0x[0-9a-fA-F]{8})")
SIGNATURE_PATTERN = re.compile(r"^\w+\((.*)\)\s*->\s*(.+)$")
INDEXED_PROPERTIES = frozenset({"Address", "AddressLocal"})
QSTRING_METHODS = frozenset({"getAllDocumentField", "getDocumentFieldValue"})
LOCALE_PARAMETERS = frozenset({"lcid", "lcidin"})


@dataclass(frozen=True, slots=True)
class Parameter:
    """Describe one argument published by a SIP method signature."""

    name: str
    annotation: str
    optional: bool = False


@dataclass(frozen=True, slots=True)
class Signature:
    """Retain argument and return types without evaluating binding text."""

    parameters: tuple[Parameter, ...]
    returns: str


def _split_types(value: str) -> list[str]:
    """Split annotations only at commas outside nested type expressions."""
    result = []
    start = 0
    depth = 0
    for index, character in enumerate(value):
        if character in "([":
            depth += 1
        elif character in ")]":
            depth -= 1
        elif character == "," and depth == 0:
            result.append(value[start:index].strip())
            start = index + 1
    result.append(value[start:].strip())
    return result


@lru_cache(maxsize=2048)
def _signature(documentation: str) -> Signature:
    """Read the generated SIP signature without importing native modules."""
    for line in documentation.splitlines():
        match = SIGNATURE_PATTERN.match(line.strip())
        if match is None:
            continue
        parameters = []
        for argument in _split_types(match[1]):
            name, separator, annotation = argument.partition(":")
            if name.strip() in {"self", "cls", ""}:
                continue
            parameters.append(
                Parameter(
                    name.strip(),
                    annotation.split("=", 1)[0].strip()
                    if separator
                    else "Any",
                    "=" in annotation,
                )
            )
        return Signature(tuple(parameters), match[2].strip())
    return Signature((), "")


def _native_error(error: Exception, operation: str) -> Exception:
    """Preserve native HRESULTs embedded in SIP property exceptions."""
    match = HRESULT_PATTERN.search(str(error))
    if match is None:
        return error
    status = int(match[1], 16)
    if status & 0x80000000:
        status -= 0x100000000
    return NativeError(f"{operation}: {error}", hresult=status)


def check_hresult(result: Any, operation: str) -> Any:
    """Raise for failed HRESULTs and detach successful output parameters."""
    status = result[0] if isinstance(result, tuple) and result else result
    if type(status) is not int:
        raise NativeError(f"{operation} returned an invalid HRESULT")
    if status & 0x80000000:
        signed = status if status < 0 else status - 0x100000000
        raise NativeError(
            f"{operation} failed with HRESULT 0x{status & 0xFFFFFFFF:08X}",
            hresult=signed,
        )
    if isinstance(result, tuple):
        return result[1] if len(result) == 2 else result[1:]
    return None


def decode_bstr(value: str) -> str:
    """Recover the valid prefix from pywpsrpc's UTF-16-sized string.

    Version 2.4 allocates a Python string using the BSTR's UTF-16 unit
    count, then writes Unicode code points. Counting units in the valid
    prefix reaches that allocation length before any unwritten suffix.
    This function must only receive values from that binding converter.
    """
    units = 0
    length = len(value)
    for index, character in enumerate(value):
        units += 2 if ord(character) > 0xFFFF else 1
        if units == length:
            return value[: index + 1]
        if units > length:
            raise NativeError("The RPC binding returned an invalid BSTR")
    return value


def unwrap(value: Any, *, variant: bool = True) -> Any:
    """Convert arguments without changing valid Unicode or literal values."""
    if isinstance(value, RpcObject):
        return value.native
    if variant and type(value) is int and not -(2**31) <= value < 2**31:
        # pywpsrpc stores Python ints in VT_I4, even on a 64-bit host.
        if not -(2**53) <= value <= 2**53:
            raise ValueError("RPC integers outside ±2**53 must use text")
        return float(value)
    if isinstance(value, str):
        if "\0" in value or any(
            0xD800 <= ord(character) <= 0xDFFF for character in value
        ):
            raise ValueError(
                "WPS RPC strings cannot contain NUL or isolated surrogates"
            )
        return value
    if isinstance(value, (list, tuple)):
        return [unwrap(item, variant=variant) for item in value]
    if isinstance(value, dict):
        return {
            key: unwrap(item, variant=variant) for key, item in value.items()
        }
    return value


def _argument(value: Any, annotation: str, api: Any) -> Any:
    """Convert VARIANT values and typed enum constants for SIP."""
    qualified = annotation.strip("'\"").split("|", 1)[0]
    value = unwrap(value, variant=qualified == "Any")
    if not isinstance(value, int):
        return value
    namespace, separator, name = qualified.partition(".")
    if namespace == "ksoapi":
        api = import_module("pywpsrpc.common").ksoapi
    elif not separator or namespace not in {"wpsapi", "etapi", "wppapi"}:
        return value
    enum_type = getattr(api, name, None)
    if not isinstance(enum_type, type) or not issubclass(enum_type, int):
        return value
    if isinstance(value, bool) and name == "MsoTriState":
        value = -1 if value else 0
    return enum_type(value)


def wrap(value: Any, api: Any, annotation: str = "") -> Any:
    """Adapt one native return value, including recursive VARIANT strings."""
    if value is None:
        return "" if annotation.strip("'\"") == "str" else None
    if isinstance(value, RpcObject):
        return value
    if isinstance(value, str):
        return decode_bstr(value)
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return int(value)
    if isinstance(value, (float, bytes, datetime, date, time)):
        return value
    if isinstance(value, list):
        return [wrap(item, api) for item in value]
    if isinstance(value, tuple):
        return tuple(wrap(item, api) for item in value)
    return RpcObject(value, api)


def _style_variant(value: Any, native: Any, api: Any, name: str) -> Any:
    """Transfer documented Writer Style VARIANT references in pywpsrpc 2.4.

    These properties return an owning IDispatch wrapper for a Style pointer.
    SIP cast creates another owning wrapper without adding a native reference;
    invalidating the original transfers its single Release to the typed owner.
    This version-specific conversion is deliberately limited to these fields.
    """
    style_type = getattr(api, "Style", None)
    style_owner = isinstance(style_type, type) and isinstance(
        native, style_type
    )
    range_owner = any(
        isinstance(native_type, type) and isinstance(native, native_type)
        for native_type in (
            getattr(api, "Range", None),
            getattr(api, "Paragraph", None),
        )
    )
    if not (
        style_owner
        and name in {"BaseStyle", "NextParagraphStyle"}
        or range_owner
        and name == "Style"
    ) or not isinstance(style_type, type):
        return value
    return _transfer_dispatch(value, style_type)


def _transfer_dispatch(value: Any, target_type: type[Any]) -> Any:
    """Transfer a verified interface reference without a second Release."""
    if type(value).__module__ != "pywpsrpc.common":
        return value
    common = import_module("pywpsrpc.common")
    if type(value) is not common.IDispatch:
        return value
    sip = import_module("pywpsrpc.sip")
    typed = sip.cast(value, target_type)
    sip.setdeleted(value)
    return typed


class RpcMethod:
    """Call a native method once, normalizing its arguments and result."""

    def __init__(self, method: Callable[..., Any], api: Any) -> None:
        self._method = method
        self._api = api
        self._name = getattr(method, "__name__", "RPC method")
        self._signature = _signature(method.__doc__ or "")

    def __call__(self, *args: Any, **kwargs: Any) -> Any:
        """Check every HRESULT without treating ordinary values as status."""
        signature = self._signature
        positional = [
            _argument(
                value,
                signature.parameters[index].annotation
                if index < len(signature.parameters)
                else "",
                self._api,
            )
            for index, value in enumerate(args)
        ]
        keywords = {}
        for name, value in kwargs.items():
            matches = [
                (index, parameter)
                for index, parameter in enumerate(signature.parameters)
                if parameter.name.casefold() == name.casefold()
            ]
            if len(matches) > 1:
                raise TypeError(f"Ambiguous RPC argument: {name}")
            if matches:
                index, parameter = matches[0]
                canonical = parameter.name
                if index < len(args) or canonical in keywords:
                    raise TypeError(f"Duplicate RPC argument: {name}")
                annotation = parameter.annotation
            else:
                canonical, annotation = name, ""
            keywords[canonical] = _argument(value, annotation, self._api)
        for index, parameter in enumerate(signature.parameters):
            if (
                parameter.name.lower() in LOCALE_PARAMETERS
                and index >= len(positional)
                and parameter.name not in keywords
            ):
                if index == len(positional):
                    positional.append(0)
                else:
                    keywords[parameter.name] = 0
        while len(positional) < len(signature.parameters):
            parameter = signature.parameters[len(positional)]
            if parameter.name not in keywords:
                break
            positional.append(keywords.pop(parameter.name))
        try:
            result = self._method(*positional, **keywords)
        except Exception as error:
            converted = _native_error(error, self._name)
            if converted is error:
                raise
            raise converted from None
        returns = signature.returns
        if returns.startswith("(int,"):
            payload = check_hresult(result, self._name)
            output_annotations = _split_types(returns[1:-1])[1:]
            if len(output_annotations) == 1:
                return wrap(payload, self._api, output_annotations[0])
            return tuple(
                wrap(item, self._api, annotation)
                for item, annotation in zip(
                    payload, output_annotations, strict=True
                )
            )
        if returns == "int" and not (
            self._name.startswith("get")
            or self._name in {"AddRef", "Release", "__len__"}
        ):
            return check_hresult(result, self._name)
        if self._name in QSTRING_METHODS:
            return result
        return wrap(result, self._api, returns)


class RpcObject:
    """Keep native proxies local while matching Office object access."""

    __slots__ = ("_native", "_api", "_getter")

    def __init__(self, native: Any, api: Any) -> None:
        object.__setattr__(self, "_native", native)
        object.__setattr__(self, "_api", api)
        object.__setattr__(self, "_getter", None)

    @property
    def native(self) -> Any:
        """Return the original proxy only for same-process native calls."""
        return self._native

    def __getattr__(self, name: str) -> Any:
        try:
            value = getattr(self._native, name)
        except Exception as error:
            converted = _native_error(error, name)
            if converted is not error:
                if (
                    isinstance(converted, NativeError)
                    and converted.hresult == -2147483645
                    and name in {"Value2", "Formula"}
                    and self._is_type("IRange")
                ):
                    return self._read_matrix(name, converted)
                raise converted from None
            if not isinstance(error, AttributeError):
                raise
            if inspect.getattr_static(self._native, name, None) is not None:
                raise
            return self._get_missing(name)
        if (
            inspect.isbuiltin(value)
            or inspect.ismethod(value)
            or inspect.isfunction(value)
        ):
            method = RpcMethod(value, self._api)
            if name == "PivotFields" and self._is_type("IPivotTable"):
                return self._pivot_fields
            if name == "SeriesCollection" and self._is_type("Chart"):
                return self._chart_series
            if name == "Axes" and self._is_type("Chart"):
                return self._chart_axes
            if name in INDEXED_PROPERTIES:
                return method()
            return method
        descriptor = getattr(type(self._native), name, None)
        annotation = _signature(
            getattr(descriptor, "__doc__", "") or ""
        ).returns
        if not annotation:
            getter = getattr(type(self._native), "get" + name, None)
            annotation = _signature(
                getattr(getter, "__doc__", "") or ""
            ).returns
        value = _style_variant(value, self._native, self._api, name)
        return wrap(value, self._api, annotation)

    def _is_type(self, name: str) -> bool:
        """Match a declared interface type in this component's namespace."""
        native_type = getattr(self._api, name, None)
        return isinstance(native_type, type) and isinstance(
            self._native, native_type
        )

    def _pivot_fields(self, *args: Any, **kwargs: Any) -> Any:
        """Resolve individual fields through the correctly typed collection."""
        if len(args) > 1 or set(kwargs) - {"Index"} or (args and kwargs):
            raise TypeError("PivotFields accepts one optional Index")
        index = args[0] if args else kwargs.get("Index")
        fields = RpcMethod(self._native.PivotFields, self._api)()
        return fields if index is None else fields.Item(index)

    def _chart_series(self, *args: Any, **kwargs: Any) -> Any:
        """Read WPP's declared series collection before selecting an item."""
        if len(args) > 1 or set(kwargs) - {"Index"} or (args and kwargs):
            raise TypeError("SeriesCollection accepts one optional Index")
        index = args[0] if args else kwargs.get("Index")
        raw = check_hresult(
            self._native.SeriesCollection(), "SeriesCollection"
        )
        collection = wrap(
            _transfer_dispatch(raw, self._api.SeriesCollection), self._api
        )
        return collection if index is None else collection.Item(index)

    def _chart_axes(self, *args: Any, **kwargs: Any) -> Any:
        """Transfer the declared Axis or Axes interface returned by WPP."""
        result = RpcMethod(self._native.Axes, self._api)(*args, **kwargs)
        named = {name.casefold(): value for name, value in kwargs.items()}
        axis_type = args[0] if args else named.get("type")
        target = self._api.Axes if axis_type is None else self._api.Axis
        return wrap(_transfer_dispatch(unwrap(result), target), self._api)

    def _read_matrix(self, name: str, error: NativeError) -> Any:
        """Read cells after the known pywpsrpc SAFEARRAY conversion failure."""
        rows = int(self.Rows.Count)
        columns = int(self.Columns.Count)
        if rows * columns <= 1:
            raise error
        return [
            [
                getattr(self.Cells.Item(row, column), name)
                for column in range(1, columns + 1)
            ]
            for row in range(1, rows + 1)
        ]

    def _get_missing(self, name: str) -> Any:
        """Use only declared raw getters and explicit object-model aliases."""
        range_type = getattr(self._api, "IRange", None)
        if (
            name == "Parent"
            and isinstance(range_type, type)
            and isinstance(self._native, range_type)
        ):
            return self.Worksheet
        getter = getattr(self._native, "get_" + name)
        signature = _signature(getter.__doc__ or "")
        method = RpcMethod(getter, self._api)
        if name == "Item" or any(
            not parameter.optional
            and parameter.name.lower() not in LOCALE_PARAMETERS
            for parameter in signature.parameters
        ):
            return method
        if not signature.returns:
            raise AttributeError(f"The RPC getter for {name} has no signature")
        result = method()
        if isinstance(result, RpcObject):
            object.__setattr__(result, "_getter", method)
        return result

    def __setattr__(self, name: str, value: Any) -> None:
        if name in {"Value2", "Value", "Formula"} and isinstance(
            value, (list, tuple)
        ):
            self._write_matrix(name, value)
            return
        setter = getattr(self._native, "set" + name, None)
        if setter is None:
            setter = getattr(self._native, "put_" + name, None)
        descriptor = getattr(type(self._native), name, None)
        if setter is None and not (
            isinstance(descriptor, property) and descriptor.fset is not None
        ):
            raise AttributeError(f"The RPC object has no writable {name}")
        parameters = _signature(
            getattr(setter, "__doc__", "") or ""
        ).parameters
        if setter is not None and parameters:
            parameter = next(
                (
                    parameter
                    for parameter in reversed(parameters)
                    if parameter.name.lower() not in LOCALE_PARAMETERS
                ),
                None,
            )
            if parameter is None:
                raise AttributeError(f"The RPC setter for {name} has no value")
            RpcMethod(setter, self._api)(**{parameter.name: value})
            return
        if not (
            isinstance(descriptor, property) and descriptor.fset is not None
        ):
            raise AttributeError(f"The RPC setter for {name} has no signature")
        converted = unwrap(value)
        try:
            setattr(self._native, name, converted)
        except Exception as error:
            converted_error = _native_error(error, name)
            if converted_error is error:
                raise
            raise converted_error from None

    def __call__(self, *args: Any, **kwargs: Any) -> Any:
        """Resolve the collection's default indexed member."""
        if self._getter is not None:
            return self._getter(*args, **kwargs) if args or kwargs else self
        return self.Item(*args, **kwargs)

    def __getitem__(self, index: Any) -> Any:
        return self.Item(index)

    def __iter__(self) -> Iterator[Any]:
        return (self.Item(index) for index in range(1, int(self.Count) + 1))

    def __len__(self) -> int:
        return int(self.Count)

    def __bool__(self) -> bool:
        return True

    def _write_matrix(
        self, name: str, value: list[Any] | tuple[Any, ...]
    ) -> None:
        """Write verified dimensions cell by cell without unsafe SAFEARRAYs."""
        matrix = unwrap(value)
        rows = int(self.Rows.Count)
        columns = int(self.Columns.Count)
        if len(matrix) != rows or any(
            not isinstance(row, list) or len(row) != columns for row in matrix
        ):
            raise ValueError("RPC matrix dimensions must match the range")
        if any(
            cell is not None and type(cell) not in (str, int, float, bool)
            for row in matrix
            for cell in row
        ):
            raise ValueError("RPC cell values must be scalar")
        completed = 0
        try:
            for row_index, row in enumerate(matrix, 1):
                for column_index, cell_value in enumerate(row, 1):
                    cell = self.Cells.Item(row_index, column_index)
                    setattr(cell, name, cell_value)
                    completed += 1
        except Exception as error:
            raise NativeError(
                f"RPC matrix write failed after {completed} of "
                f"{rows * columns} cells; inspect before repeating: {error}",
                hresult=getattr(error, "hresult", None),
            ) from None
