"""Exercise native Writer tools using only test-owned files."""

from pathlib import Path

import anyio

from tests.native.client import NativeClient, native_client


async def builtin_style_names(
    client: NativeClient, identity: str
) -> tuple[str, str]:
    """Read localized built-in names through the public tool contract."""
    names = []
    for identifier in (-1, -2):
        result = await client.call(
            "wps_writer_styles", "read", document=identity, name=identifier
        )
        names.append(str(result["name"]))
    return names[0], names[1]


async def save_and_reopen(
    client: NativeClient, identity: str, target: Path
) -> str:
    """Round-trip DOCX content while retaining cleanup ownership."""
    saved = await client.call(
        "wps_documents",
        "save",
        component="writer",
        document=identity,
        file_path=str(target),
    )
    assert saved["saved"] and target.is_file() and target.stat().st_size > 0
    await client.call(
        "wps_documents",
        "close",
        component="writer",
        document=saved["full_name"],
    )
    reopened = await client.call(
        "wps_documents", "open", component="writer", file_path=str(target)
    )
    return str(reopened["full_name"])


def test_native_writer_revisions_comments_and_persistence(
    tmp_path: Path,
) -> None:
    async def scenario() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="writer"
            )
            identity = created["full_name"]
            await client.call(
                "wps_writer_text",
                "insert",
                document=identity,
                position=0,
                text="Original paragraph.\r",
            )
            await client.call(
                "wps_writer_revisions",
                "configure",
                document=identity,
                enabled=True,
            )
            await client.call(
                "wps_writer_text",
                "insert",
                document=identity,
                position=0,
                text="Accepted change. ",
            )
            revisions = await client.call(
                "wps_writer_revisions", "list", document=identity
            )
            assert revisions["total"] > 0
            assert any(
                "Accepted change." in item["text"]
                for item in revisions["revisions"]
            )
            accepted = await client.call(
                "wps_writer_revisions",
                "accept",
                document=identity,
                indices=[item["index"] for item in revisions["revisions"]],
            )
            assert accepted["count"] == revisions["total"]
            await client.call(
                "wps_writer_text",
                "insert",
                document=identity,
                position=0,
                text="Rejected change. ",
            )
            rejected = await client.call(
                "wps_writer_revisions", "reject", document=identity, all=True
            )
            assert rejected["count"] > 0
            state = await client.call(
                "wps_writer_revisions",
                "configure",
                document=identity,
                enabled=False,
            )
            assert state == {"enabled": False, "count": 0}
            comment = await client.call(
                "wps_writer_comments",
                "add",
                document=identity,
                start=0,
                end=8,
                text="First comment",
                author="Native Test",
                initials="NT",
            )
            changed = await client.call(
                "wps_writer_comments",
                "update",
                document=identity,
                index=comment["index"],
                text="Updated comment 中文",
                author="Native Reviewer",
                initials="NR",
            )
            assert changed["scope"] == comment["scope"]
            assert changed["text"].removesuffix("\r") == "Updated comment 中文"
            comments = await client.call(
                "wps_writer_comments", "list", document=identity
            )
            assert comments["total"] == 1
            identity = await save_and_reopen(
                client, identity, tmp_path / "writer-review.docx"
            )
            text = await client.call(
                "wps_writer_text", "read", document=identity
            )
            assert "Accepted change. Original paragraph." in text["text"]
            assert "Rejected change." not in text["text"]
            assert (
                await client.call(
                    "wps_writer_revisions", "state", document=identity
                )
                == state
            )
            comments = await client.call(
                "wps_writer_comments", "list", document=identity
            )
            assert comments["total"] == 1
            persisted = comments["comments"][0]
            assert (
                persisted["text"].removesuffix("\r") == "Updated comment 中文"
            )
            assert persisted["author"] == "Native Reviewer"
            assert persisted["initials"] == "NR"
            assert persisted["scope"] == comment["scope"]
            removed = await client.call(
                "wps_writer_comments",
                "delete",
                document=identity,
                indices=[persisted["index"]],
            )
            assert removed["deleted_count"] == 1
            assert (
                await client.call(
                    "wps_writer_comments", "list", document=identity
                )
            )["total"] == 0

    anyio.run(scenario)


def test_native_writer_style_inheritance_and_persistence(
    tmp_path: Path,
) -> None:
    async def scenario() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="writer"
            )
            identity = created["full_name"]
            normal, _ = await builtin_style_names(client, identity)
            await client.call(
                "wps_writer_text",
                "insert",
                document=identity,
                position=0,
                text="Styled paragraph.\r",
            )
            await client.call(
                "wps_writer_styles",
                "create",
                document=identity,
                name="NativeBase",
                base_style=normal,
                next_style=normal,
                automatically_update=False,
                character={"font_size": 14.0, "bold": True},
                paragraph={"space_after": 8.0},
            )
            await client.call(
                "wps_writer_styles",
                "create",
                document=identity,
                name="NativeDerived",
                base_style="NativeBase",
                next_style=normal,
                character={"italic": True},
            )
            applied = await client.call(
                "wps_writer_styles",
                "apply",
                document=identity,
                name="NativeDerived",
                start=0,
                end=18,
            )
            assert applied["name"] == "NativeDerived"
            await client.call(
                "wps_writer_styles",
                "update",
                document=identity,
                name="NativeBase",
                character={"font_size": 18.0},
                paragraph={"space_after": 12.0},
            )
            style = await client.call(
                "wps_writer_styles",
                "read",
                document=identity,
                name="NativeDerived",
            )
            assert style["base_style"] == "NativeBase"
            assert style["character"]["font_size"] == 18.0
            assert (
                style["character"]["bold"] != 0
                and style["character"]["italic"] != 0
            )
            assert style["paragraph"]["space_after"] == 12.0
            identity = await save_and_reopen(
                client, identity, tmp_path / "writer-styles.docx"
            )
            styles = await client.call(
                "wps_writer_styles", "list", document=identity, built_in=False
            )
            assert {"NativeBase", "NativeDerived"} <= {
                item["name"] for item in styles["styles"]
            }
            persisted = await client.call(
                "wps_writer_styles",
                "read",
                document=identity,
                name="NativeDerived",
            )
            assert persisted["base_style"] == style["base_style"]
            assert persisted["character"]["font_size"] == 18.0
            assert persisted["paragraph"]["space_after"] == 12.0
            await client.call(
                "wps_writer_styles",
                "apply",
                document=identity,
                name=normal,
                start=0,
                end=18,
            )
            for name in ("NativeDerived", "NativeBase"):
                removed = await client.call(
                    "wps_writer_styles", "delete", document=identity, name=name
                )
                assert removed["deleted"]
            remaining = await client.call(
                "wps_writer_styles", "list", document=identity, built_in=False
            )
            assert not {"NativeBase", "NativeDerived"} & {
                item["name"] for item in remaining["styles"]
            }

    anyio.run(scenario)


def test_native_writer_toc_notes_cross_references_and_persistence(
    tmp_path: Path,
) -> None:
    async def scenario() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="writer"
            )
            identity = created["full_name"]
            _, heading = await builtin_style_names(client, identity)
            await client.call(
                "wps_writer_text",
                "insert",
                document=identity,
                position=0,
                text=(
                    "\rChapter Alpha\rBookmark payload\r"
                    "Notes and references.\r"
                ),
            )
            paragraphs = await client.call(
                "wps_writer_text", "paragraphs", document=identity
            )
            title, bookmarked = paragraphs["paragraphs"][1:3]
            await client.call(
                "wps_writer_styles",
                "apply",
                document=identity,
                name=heading,
                start=title["start"],
                end=title["end"],
            )
            await client.call(
                "wps_writer_bookmarks",
                "add",
                document=identity,
                name="NativeBookmark",
                start=bookmarked["start"],
                end=bookmarked["end"] - 1,
            )
            await client.call(
                "wps_writer_toc",
                "add",
                document=identity,
                position=0,
                lower_heading_level=1,
            )
            tocs = await client.call(
                "wps_writer_toc", "list", document=identity
            )
            assert len(tocs["tables_of_contents"]) == 1
            toc_index = tocs["tables_of_contents"][0]["toc_index"]
            for mode in ("all", "page_numbers"):
                updated = await client.call(
                    "wps_writer_toc",
                    "update",
                    document=identity,
                    toc_index=toc_index,
                    mode=mode,
                )
                assert "Chapter Alpha" in updated["text"]
            for kind in ("footnote", "endnote"):
                text = await client.call(
                    "wps_writer_text", "read", document=identity
                )
                await client.call(
                    "wps_writer_notes",
                    "add",
                    document=identity,
                    kind=kind,
                    position=text["end"] - 1,
                    text=f"Original {kind}",
                )
                notes = await client.call(
                    "wps_writer_notes", "list", document=identity, kind=kind
                )
                assert len(notes["notes"]) == 1
                updated = await client.call(
                    "wps_writer_notes",
                    "update",
                    document=identity,
                    kind=kind,
                    note_index=notes["notes"][0]["note_index"],
                    text=f"Updated {kind} 中文",
                )
                assert f"Updated {kind} 中文" in updated["content"]["text"]
            reference_codes: list[str] = []
            for reference_type, reference_kind, expected in (
                ("bookmark", "text", "Bookmark payload"),
                ("heading", "text", "Chapter Alpha"),
                ("footnote", "number", ""),
                ("endnote", "number_formatted", ""),
            ):
                targets = await client.call(
                    "wps_writer_cross_references",
                    "targets",
                    document=identity,
                    reference_type=reference_type,
                )
                assert targets["targets"]
                target = next(
                    item
                    for item in targets["targets"]
                    if reference_type not in {"bookmark", "heading"}
                    or (
                        "NativeBookmark"
                        if reference_type == "bookmark"
                        else expected
                    )
                    in item["text"]
                )
                text = await client.call(
                    "wps_writer_text", "read", document=identity
                )
                await client.call(
                    "wps_writer_cross_references",
                    "insert",
                    document=identity,
                    position=text["end"] - 1,
                    reference_type=reference_type,
                    reference_kind=reference_kind,
                    reference_item=target["reference_item"],
                )
                listed = await client.call(
                    "wps_writer_cross_references", "list", document=identity
                )
                latest = max(
                    listed["references"],
                    key=lambda item: item["result"]["start"],
                )
                updated = await client.call(
                    "wps_writer_cross_references",
                    "update",
                    document=identity,
                    field_index=latest["field_index"],
                )
                assert updated["updated"] and updated["result"]["text"].strip()
                assert expected in updated["result"]["text"]
                reference_codes.append(updated["code"])
            identity = await save_and_reopen(
                client, identity, tmp_path / "writer-references.docx"
            )
            tocs = await client.call(
                "wps_writer_toc", "list", document=identity
            )
            assert len(tocs["tables_of_contents"]) == 1
            assert "Chapter Alpha" in tocs["tables_of_contents"][0]["text"]
            bookmarks = await client.call(
                "wps_writer_bookmarks", "list", document=identity
            )
            assert any(
                item["name"] == "NativeBookmark"
                and item["text"] == "Bookmark payload"
                for item in bookmarks["bookmarks"]
            )
            persisted = await client.call(
                "wps_writer_cross_references", "list", document=identity
            )
            references = [
                item
                for item in persisted["references"]
                if item["code"] in reference_codes
            ]
            assert len(references) == 4
            assert any(
                "Bookmark payload" in item["result"]["text"]
                for item in references
            )
            assert any(
                "Chapter Alpha" in item["result"]["text"]
                for item in references
            )
            for reference in sorted(
                references, key=lambda item: item["field_index"], reverse=True
            ):
                removed = await client.call(
                    "wps_writer_cross_references",
                    "delete",
                    document=identity,
                    field_index=reference["field_index"],
                )
                assert removed["deleted"]
            remaining = await client.call(
                "wps_writer_cross_references", "list", document=identity
            )
            assert not any(
                item["code"] in reference_codes
                for item in remaining["references"]
            )
            for kind in ("footnote", "endnote"):
                notes = await client.call(
                    "wps_writer_notes", "list", document=identity, kind=kind
                )
                assert len(notes["notes"]) == 1
                assert (
                    f"Updated {kind} 中文"
                    in notes["notes"][0]["content"]["text"]
                )
                removed = await client.call(
                    "wps_writer_notes",
                    "delete",
                    document=identity,
                    kind=kind,
                    note_index=notes["notes"][0]["note_index"],
                )
                assert removed["deleted"]
                assert not (
                    await client.call(
                        "wps_writer_notes",
                        "list",
                        document=identity,
                        kind=kind,
                    )
                )["notes"]
            removed = await client.call(
                "wps_writer_toc",
                "delete",
                document=identity,
                toc_index=tocs["tables_of_contents"][0]["toc_index"],
            )
            assert removed["deleted"]
            assert not (
                await client.call("wps_writer_toc", "list", document=identity)
            )["tables_of_contents"]

    anyio.run(scenario)
