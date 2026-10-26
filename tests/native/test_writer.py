"""Exercise native Writer tools using only test-owned files."""

from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

import anyio
import pytest
from PIL import Image

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


NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}


def document_xml(path: Path) -> ElementTree.Element:
    """Inspect the document saved by WPS without modifying its package."""
    with ZipFile(path) as archive:
        return ElementTree.fromstring(archive.read("word/document.xml"))


def test_native_writer_text_edits_format_and_bookmark_delete(
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
                text="Alpha 😀 Alpha\rSecond line.\r",
            )
            found = await client.call(
                "wps_writer_text",
                "find",
                document=identity,
                text="Alpha",
                match_case=True,
                whole_word=True,
            )
            assert [
                (item["start"], item["end"]) for item in found["matches"]
            ] == [
                (0, 5),
                (9, 14),
            ]
            assert not found["has_more"]
            changed = await client.call(
                "wps_writer_text",
                "replace_matches",
                document=identity,
                text="Alpha",
                replacement="Beta",
                replace_all=True,
            )
            assert changed["replaced_count"] == 2
            replaced = await client.call(
                "wps_writer_text",
                "replace",
                document=identity,
                start=0,
                end=4,
                text="Root",
            )
            assert replaced["text"] == "Root"
            deleted = await client.call(
                "wps_writer_text",
                "delete",
                document=identity,
                start=5,
                end=7,
            )
            assert deleted["deleted"]
            await client.call(
                "wps_writer_text",
                "format",
                document=identity,
                start=0,
                end=4,
                character={
                    "font_size": 14.0,
                    "bold": True,
                    "color": "#123456",
                },
                paragraph={"alignment": "center", "left_indent": 12.0},
            )
            await client.call(
                "wps_writer_bookmarks",
                "add",
                document=identity,
                name="DeleteOnlyAnchor",
                start=0,
                end=4,
            )
            removed = await client.call(
                "wps_writer_bookmarks",
                "delete",
                document=identity,
                name="DeleteOnlyAnchor",
            )
            assert removed["deleted"]
            assert not (
                await client.call(
                    "wps_writer_bookmarks", "list", document=identity
                )
            )["bookmarks"]
            target = tmp_path / "writer-text.docx"
            identity = await save_and_reopen(client, identity, target)
            text = await client.call(
                "wps_writer_text", "read", document=identity
            )
            assert text["text"].startswith("Root  Beta\rSecond line.\r")
            assert "😀" not in text["text"]
            root = document_xml(target)
            run = next(
                run
                for run in root.findall(".//w:r", NAMESPACES)
                if "".join(run.itertext()) == "Root"
            )
            bold = run.find("w:rPr/w:b", NAMESPACES)
            color = run.find("w:rPr/w:color", NAMESPACES)
            size = run.find("w:rPr/w:sz", NAMESPACES)
            assert bold is not None
            assert (
                color is not None
                and color.get(f"{{{NAMESPACES['w']}}}val", "").lower()
                == "123456"
            )
            assert (
                size is not None
                and size.get(f"{{{NAMESPACES['w']}}}val") == "28"
            )
            alignment = root.find(".//w:pPr/w:jc", NAMESPACES)
            assert alignment is not None
            assert alignment.get(f"{{{NAMESPACES['w']}}}val") == "center"

    anyio.run(scenario)


def test_native_writer_tables_image_and_statistics(tmp_path: Path) -> None:
    async def scenario() -> None:
        image_path = tmp_path / "embedded.png"
        Image.new("RGB", (40, 20), "#336699").save(image_path)
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="writer"
            )
            identity = created["full_name"]
            await client.call(
                "wps_writer_tables",
                "add",
                document=identity,
                position=0,
                data=[["标题", "值"], [0, False]],
            )
            listed = await client.call(
                "wps_writer_tables", "list", document=identity
            )
            assert len(listed["tables"]) == 1
            assert listed["tables"][0]["cell_count"] == 4
            table = await client.call(
                "wps_writer_tables",
                "read",
                document=identity,
                table_index=1,
            )
            assert [cell["text"] for cell in table["cells"]] == [
                "标题",
                "值",
                "0",
                "false",
            ]
            written = await client.call(
                "wps_writer_tables",
                "write",
                document=identity,
                table_index=1,
                start_row=2,
                data=[[None, "替换"]],
            )
            assert written["cells_written"] == 2
            await client.call(
                "wps_writer_tables",
                "format_cell",
                document=identity,
                table_index=1,
                row=1,
                column=1,
                character={"bold": True, "color": "#FFFFFF"},
                alignment="center",
                vertical_alignment="center",
                fill_color="#336699",
            )
            for action, expected_count in (
                ("insert_row", 6),
                ("insert_column", 9),
                ("delete_column", 6),
                ("delete_row", 4),
            ):
                await client.call(
                    "wps_writer_tables",
                    "change",
                    document=identity,
                    table_index=1,
                    action=action,
                    position=2,
                )
                table = await client.call(
                    "wps_writer_tables",
                    "read",
                    document=identity,
                    table_index=1,
                )
                assert len(table["cells"]) == expected_count
            assert [cell["text"] for cell in table["cells"]] == [
                "标题",
                "值",
                "",
                "替换",
            ]
            target = tmp_path / "writer-table.docx"
            identity = await save_and_reopen(client, identity, target)
            root = document_xml(target)
            fill = root.find(".//w:tcPr/w:shd", NAMESPACES)
            assert fill is not None
            assert (
                fill.get(f"{{{NAMESPACES['w']}}}fill", "").lower() == "336699"
            )
            vertical = root.find(".//w:tcPr/w:vAlign", NAMESPACES)
            assert vertical is not None
            assert vertical.get(f"{{{NAMESPACES['w']}}}val") == "center"
            await client.call(
                "wps_writer_tables",
                "delete",
                document=identity,
                table_index=1,
            )
            assert not (
                await client.call(
                    "wps_writer_tables", "list", document=identity
                )
            )["tables"]
            added = await client.call(
                "wps_writer_images",
                "add",
                document=identity,
                position=0,
                path=str(image_path),
                width=72.0,
                height=36.0,
                alternative_text="Native embedded image",
            )
            assert added["width"] == pytest.approx(72.0, abs=0.2)
            assert added["height"] == pytest.approx(36.0, abs=0.2)
            identity = await save_and_reopen(
                client, identity, tmp_path / "writer-image.docx"
            )
            statistics = await client.call(
                "wps_writer_statistics", "read", document=identity
            )
            assert statistics["tables"] == 0
            assert statistics["inline_images"] == 1
            assert statistics["sections"] == 1
            assert statistics["range_end"] > 1

    anyio.run(scenario)


def test_native_writer_sections_headers_and_page_setup(tmp_path: Path) -> None:
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
                text="First section.\r",
            )
            await client.call(
                "wps_writer_pages",
                "setup",
                document=identity,
                orientation="landscape",
                page_width=792.0,
                page_height=612.0,
                margin_top=36.0,
                margin_bottom=36.0,
                margin_left=42.0,
                margin_right=42.0,
            )
            first_header = "Original first header"
            await client.call(
                "wps_writer_pages",
                "write_header_footer",
                document=identity,
                location="header",
                text=first_header,
            )
            for kind in (
                "line",
                "page",
                "section_next_page",
                "section_continuous",
            ):
                text = await client.call(
                    "wps_writer_text", "read", document=identity
                )
                await client.call(
                    "wps_writer_pages",
                    "insert_break",
                    document=identity,
                    position=text["end"] - 1,
                    kind=kind,
                )
                text = await client.call(
                    "wps_writer_text", "read", document=identity
                )
                await client.call(
                    "wps_writer_text",
                    "insert",
                    document=identity,
                    position=text["end"] - 1,
                    text=f"After {kind}.\r",
                )
            sections = await client.call(
                "wps_writer_pages", "sections", document=identity
            )
            assert len(sections["sections"]) == 3
            setup = sections["sections"][0]
            assert setup["orientation"] == "landscape"
            for key, value in {
                "page_width": 792.0,
                "page_height": 612.0,
                "margin_top": 36.0,
                "margin_left": 42.0,
            }.items():
                assert setup[key] == pytest.approx(value, abs=0.2)
            variants = (
                ("header", "primary", "独立第二节页眉"),
                ("footer", "first_page", "First page footer"),
                ("footer", "even_pages", "Even page footer"),
            )
            for location, variant, value in variants:
                written = await client.call(
                    "wps_writer_pages",
                    "write_header_footer",
                    document=identity,
                    section_index=2,
                    location=location,
                    variant=variant,
                    text=value,
                    link_to_previous=False,
                )
                assert value in written["text"]
                assert not written["link_to_previous"]
            identity = await save_and_reopen(
                client, identity, tmp_path / "writer-pages.docx"
            )
            for location, variant, value in variants:
                item = await client.call(
                    "wps_writer_pages",
                    "read_header_footer",
                    document=identity,
                    section_index=2,
                    location=location,
                    variant=variant,
                )
                assert value in item["text"]
                assert not item["link_to_previous"]
            preserved = await client.call(
                "wps_writer_pages",
                "read_header_footer",
                document=identity,
                location="header",
            )
            assert first_header in preserved["text"]
            statistics = await client.call(
                "wps_writer_statistics", "read", document=identity
            )
            assert statistics["sections"] == 3
            assert statistics["pages"] >= 2
            assert statistics["words"] > 0

    anyio.run(scenario)
