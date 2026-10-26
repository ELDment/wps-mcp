"""Share presentation setup without hiding public operation contracts."""

from collections.abc import Awaitable, Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import anyio
import pytest

from tests.native.client import NativeClient, native_client
from wps_mcp.models import JsonObject


@dataclass
class Deck:
    client: NativeClient
    identity: str
    monkeypatch: pytest.MonkeyPatch

    async def call(
        self, group: str, operation: str, **params: Any
    ) -> JsonObject:
        return await self.client.call(f"wps_ppt_{group}", operation, **params)

    async def document(self, operation: str, **params: Any) -> JsonObject:
        result = await self.client.call(
            "wps_documents", operation, component="wpp", **params
        )
        if operation in {"create", "open", "save"}:
            self.identity = str(result["full_name"])
        return result

    async def shape(self, text: str = "原生测试", **params: Any) -> str:
        options = {
            "shape_type": "rectangle",
            "left": 40.0,
            "top": 50.0,
            "width": 180.0,
            "height": 90.0,
            "text": text,
        }
        result = await self.call(
            "shapes",
            "add",
            document=self.identity,
            slide_index=1,
            **options | params,
        )
        return str(result["shape_name"])

    async def text(self, name: str) -> str:
        return str(
            (
                await self.call(
                    "text",
                    "get",
                    document=self.identity,
                    slide_index=1,
                    shape_name_or_index=name,
                )
            )["text"]
        )


Scenario = Callable[[Deck, Path], Awaitable[None]]


def run_presentation(
    scenario: Scenario, directory: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    async def run() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="wpp"
            )
            deck = Deck(client, str(created["full_name"]), monkeypatch)
            await deck.call(
                "slides", "add", document=deck.identity, layout="blank"
            )
            await scenario(deck, directory)

    anyio.run(run)
