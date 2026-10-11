"""Exercise mail merge with test-created CSV data and Writer documents."""

from pathlib import Path

import anyio

from tests.native.client import native_client


def test_native_mail_merge_records_save_reopen_and_disconnect(
    tmp_path: Path,
) -> None:
    async def scenario() -> None:
        source_path = tmp_path / "recipients.csv"
        source_path.write_text(
            "Name,City\nAlice,Paris\nBob,Rome\n",
            encoding="utf-8",
            newline="\r\n",
        )
        async with native_client() as client:
            template = await client.call(
                "wps_documents", "create", component="writer"
            )
            template_identity = template["full_name"]
            await client.call(
                "wps_writer_text",
                "insert",
                document=template_identity,
                position=0,
                text="Name: \rCity: \r",
            )
            connection = await client.call(
                "wps_writer_mail_merge",
                "connect",
                document=template_identity,
                file_path=str(source_path),
            )
            assert connection["source"]["fields"] == ["Name", "City"]
            assert connection["source"]["record_count"] == 2
            assert (
                Path(connection["source"]["file_path"]).resolve()
                == source_path.resolve()
            )

            # Insert backwards to keep earlier UTF-16 offsets stable.
            for name, position in (
                ("City", len("Name: \rCity: ")),
                ("Name", len("Name: ")),
            ):
                field = await client.call(
                    "wps_writer_mail_merge",
                    "field",
                    document=template_identity,
                    position=position,
                    name=name,
                )
                assert "MERGEFIELD" in field["code"].upper()
                assert name in field["code"]

            before = await client.call(
                "wps_writer_mail_merge", "inspect", document=template_identity
            )
            assert before["source"]["fields"] == ["Name", "City"]
            assert len(before["fields"]) == 2
            field_codes = [field["code"] for field in before["fields"]]

            for selection in ("second", "all"):
                merged = await client.call(
                    "wps_writer_mail_merge",
                    "execute",
                    document=template_identity,
                    **(
                        {"first_record": 2, "last_record": 2}
                        if selection == "second"
                        else {}
                    ),
                )
                assert merged["full_name"] != template_identity
                assert merged["source_document"] == template_identity
                assert merged["first_record"] == (
                    2 if selection == "second" else 1
                )
                assert merged["last_record"] == 2
                output = await client.call(
                    "wps_writer_text", "read", document=merged["full_name"]
                )
                text = output["text"]
                assert text.count("Name: Bob") == 1
                assert text.count("City: Rome") == 1
                if selection == "all":
                    assert text.count("Name: Alice") == 1
                    assert text.count("City: Paris") == 1
                else:
                    assert "Alice" not in text and "Paris" not in text
                assert "MERGEFIELD" not in text

                output_path = tmp_path / f"merged-{selection}.docx"
                saved = await client.call(
                    "wps_documents",
                    "save",
                    component="writer",
                    document=merged["full_name"],
                    file_path=str(output_path),
                )
                assert saved["saved"] and output_path.is_file()
                await client.call(
                    "wps_documents",
                    "close",
                    component="writer",
                    document=saved["full_name"],
                )
                reopened = await client.call(
                    "wps_documents",
                    "open",
                    component="writer",
                    file_path=str(output_path),
                )
                persisted = await client.call(
                    "wps_writer_text", "read", document=reopened["full_name"]
                )
                assert persisted["text"] == text
                await client.call(
                    "wps_documents",
                    "close",
                    component="writer",
                    document=reopened["full_name"],
                )

                restored = await client.call(
                    "wps_writer_mail_merge",
                    "inspect",
                    document=template_identity,
                )
                assert (
                    restored["source"]["first_record"]
                    == before["source"]["first_record"]
                )
                assert (
                    restored["source"]["last_record"]
                    == before["source"]["last_record"]
                )
                assert [
                    field["code"] for field in restored["fields"]
                ] == field_codes

            disconnected = await client.call(
                "wps_writer_mail_merge",
                "disconnect",
                document=template_identity,
            )
            assert disconnected["disconnected"] is True
            detached = await client.call(
                "wps_writer_mail_merge", "inspect", document=template_identity
            )
            assert detached["source"] is None
            assert [
                field["code"] for field in detached["fields"]
            ] == field_codes
            assert source_path.is_file()

    anyio.run(scenario)
