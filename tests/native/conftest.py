"""Require explicit opt-in for every native WPS test in this directory."""

import os
from pathlib import Path

import pytest


def pytest_collection_modifyitems(items: list[pytest.Item]) -> None:
    """Apply the native marker and one shared environment gate."""
    directory = Path(__file__).resolve().parent
    enabled = os.environ.get("WPS_NATIVE_TESTS") == "1"
    for item in items:
        if item.path.is_relative_to(directory):
            item.add_marker(pytest.mark.native)
            if not enabled:
                item.add_marker(
                    pytest.mark.skip(
                        reason="Set WPS_NATIVE_TESTS=1 on a WPS host"
                    )
                )
