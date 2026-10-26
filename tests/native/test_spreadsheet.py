"""Opt-in ET workflows; all mutations target newly created workbooks."""

from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

import anyio

from tests.native.client import NativeClient, native_client


async def save_and_reopen(
    client: NativeClient, identity: str, path: Path
) -> str:
    """Round-trip XLSX content while retaining cleanup ownership."""
    saved = await client.call(
        "wps_documents",
        "save",
        component="et",
        document=identity,
        file_path=str(path),
    )
    assert saved["saved"] and path.is_file() and path.stat().st_size > 0
    await client.call(
        "wps_documents", "close", component="et", document=saved["full_name"]
    )
    opened = await client.call(
        "wps_documents", "open", component="et", file_path=str(path)
    )
    assert isinstance(opened["full_name"], str)
    return opened["full_name"]


def test_native_et_rules_validation_and_scoped_names(tmp_path: Path) -> None:
    async def scenario() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="et"
            )
            identity = created["full_name"]
            await client.call(
                "wps_et_sheets",
                "rename",
                document=identity,
                sheet=1,
                name="规则 Data",
            )
            await client.call(
                "wps_et_ranges",
                "write_values",
                document=identity,
                sheet=1,
                address="A1:B4",
                values=[
                    ["Name", "Amount"],
                    ["Alpha", 10],
                    ["Beta", 20],
                    ["Gamma", 30],
                ],
            )
            await client.call(
                "wps_et_conditional_formats",
                "add",
                document=identity,
                sheet=1,
                address="B2:B4",
                rule_type="cell_value",
                operator="greater",
                formula1="15",
                fill_color="#FFE699",
                font_color="#123456",
            )
            conditions = await client.call(
                "wps_et_conditional_formats",
                "list",
                document=identity,
                sheet=1,
                address="B2:B4",
            )
            assert conditions["count"] == 1
            rule = conditions["rules"][0]
            assert rule["rule_type"] == "cell_value"
            assert rule["operator"] == "greater"
            assert rule["fill_color"] == "#FFE699"
            assert rule["font_color"] == "#123456"
            await client.call(
                "wps_et_conditional_formats",
                "update",
                document=identity,
                sheet=1,
                address="B2:B4",
                index=rule["index"],
                rule_type="expression",
                formula1="=B2>15",
                bold=True,
                stop_if_true=True,
            )

            absent = await client.call(
                "wps_et_validation",
                "read",
                document=identity,
                sheet=1,
                address="D2",
            )
            assert absent["present"] is False
            await client.call(
                "wps_et_validation",
                "set",
                document=identity,
                sheet=1,
                address="D2:D4",
                validation_type="whole",
                operator="between",
                formula1="1",
                formula2="100",
                input_title="Count",
                input_message="Enter a whole number",
                error_title="Invalid count",
                error_message="Use 1 through 100",
            )
            validation = await client.call(
                "wps_et_validation",
                "read",
                document=identity,
                sheet=1,
                address="D2",
            )
            assert validation["validation_type"] == "whole"
            assert validation["operator"] == "between"
            assert validation["formula1"].lstrip("=") == "1"
            assert validation["formula2"].lstrip("=") == "100"
            assert validation["input_title"] == "Count"
            await client.call(
                "wps_et_validation",
                "set",
                document=identity,
                sheet=1,
                address="D2:D4",
                validation_type="list",
                formula1="Open,Closed",
                in_cell_dropdown=False,
                ignore_blank=False,
                input_title="Status",
                input_message="Choose a status",
                error_title="Invalid status",
                error_message="Use Open or Closed",
            )

            await client.call(
                "wps_et_names",
                "set",
                document=identity,
                name="SharedArea",
                source_sheet=1,
                address="A2:B3",
            )
            await client.call(
                "wps_et_names",
                "set",
                document=identity,
                name="SharedArea",
                scope="sheet",
                sheet=1,
                source_sheet=1,
                address="A4:B4",
            )
            global_name = await client.call(
                "wps_et_names",
                "read",
                document=identity,
                name="SharedArea",
                include_values=True,
            )
            local_name = await client.call(
                "wps_et_names",
                "read",
                document=identity,
                name="SharedArea",
                scope="sheet",
                sheet=1,
                include_values=True,
            )
            assert global_name["values"] == [["Alpha", 10], ["Beta", 20]]
            assert local_name["values"] == [["Gamma", 30]]
            await client.call(
                "wps_et_names",
                "set",
                document=identity,
                name="SharedArea",
                source_sheet=1,
                address="A3:B4",
            )

            identity = await save_and_reopen(
                client, identity, tmp_path / "et-rules.xlsx"
            )
            conditions = await client.call(
                "wps_et_conditional_formats",
                "list",
                document=identity,
                sheet=1,
                address="B2:B4",
            )
            assert conditions["count"] == 1
            persisted = conditions["rules"][0]
            assert persisted["rule_type"] == "expression"
            assert persisted["formula1"].lstrip("=") == "B2>15"
            assert persisted["fill_color"] == "#FFE699"
            assert bool(persisted["bold"]) and persisted["stop_if_true"]
            validation = await client.call(
                "wps_et_validation",
                "read",
                document=identity,
                sheet=1,
                address="D3",
            )
            assert validation["validation_type"] == "list"
            assert validation["formula1"] == "Open,Closed"
            assert validation["ignore_blank"] is False
            assert validation["in_cell_dropdown"] is False
            assert validation["input_title"] == "Status"
            assert validation["error_message"] == "Use Open or Closed"
            global_name = await client.call(
                "wps_et_names",
                "read",
                document=identity,
                name="SharedArea",
                include_values=True,
            )
            local_name = await client.call(
                "wps_et_names",
                "read",
                document=identity,
                name="SharedArea",
                scope="sheet",
                sheet=1,
                include_values=True,
            )
            assert global_name["values"] == [["Beta", 20], ["Gamma", 30]]
            assert local_name["values"] == [["Gamma", 30]]
            assert (
                global_name["scope"] == "workbook"
                and local_name["scope"] == "sheet"
            )
            await client.call(
                "wps_et_names",
                "delete",
                document=identity,
                name="SharedArea",
                scope="sheet",
                sheet=1,
            )
            remaining = await client.call(
                "wps_et_names", "list", document=identity
            )
            assert any(
                name["name"] == "SharedArea" for name in remaining["names"]
            )
            await client.call(
                "wps_et_names", "delete", document=identity, name="SharedArea"
            )
            for scope in ("workbook", "sheet"):
                remaining = await client.call(
                    "wps_et_names",
                    "list",
                    document=identity,
                    scope=scope,
                    **({"sheet": 1} if scope == "sheet" else {}),
                )
                assert not any(
                    name["name"] == "SharedArea" for name in remaining["names"]
                )
            await client.call(
                "wps_et_conditional_formats",
                "delete",
                document=identity,
                sheet=1,
                address="B2:B4",
                index=persisted["index"],
            )
            assert (
                await client.call(
                    "wps_et_conditional_formats",
                    "list",
                    document=identity,
                    sheet=1,
                    address="B2:B4",
                )
            )["count"] == 0
            await client.call(
                "wps_et_validation",
                "clear",
                document=identity,
                sheet=1,
                address="D2:D4",
            )
            assert (
                await client.call(
                    "wps_et_validation",
                    "read",
                    document=identity,
                    sheet=1,
                    address="D2",
                )
            )["present"] is False

    anyio.run(scenario)


def test_native_et_pivot_layout_refresh_and_persistence(
    tmp_path: Path,
) -> None:
    async def scenario() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="et"
            )
            identity = created["full_name"]
            await client.call(
                "wps_et_sheets",
                "rename",
                document=identity,
                sheet=1,
                name="Source",
            )
            await client.call(
                "wps_et_sheets", "add", document=identity, name="Reports"
            )
            await client.call(
                "wps_et_ranges",
                "write_values",
                document=identity,
                sheet="Source",
                address="A1:D5",
                values=[
                    ["Region", "Category", "Channel", "Amount"],
                    ["East", "A", "Online", 10],
                    ["East", "B", "Store", 20],
                    ["West", "A", "Online", 30],
                    ["West", "B", "Store", 40],
                ],
            )
            await client.call(
                "wps_et_pivot_tables",
                "create",
                document=identity,
                sheet="Source",
                address="A1:D5",
                name="SalesPivot",
                destination_sheet="Reports",
                destination_cell="A4",
                row_fields=["Region"],
                data_fields=[{"field": "Amount", "aggregate": "sum"}],
            )
            pivot = await client.call(
                "wps_et_pivot_tables",
                "read",
                document=identity,
                sheet="Reports",
                pivot="SalesPivot",
            )
            assert pivot["values"][-1][-1] == 100
            assert pivot["data_fields"][0]["aggregate"] == "sum"
            await client.call(
                "wps_et_pivot_tables",
                "configure",
                document=identity,
                sheet="Reports",
                pivot="SalesPivot",
                row_fields=["Region"],
                column_fields=["Category"],
                page_fields=["Channel"],
                data_fields=[{"field": "Amount", "aggregate": "avg"}],
            )
            pivot = await client.call(
                "wps_et_pivot_tables",
                "read",
                document=identity,
                sheet="Reports",
                pivot="SalesPivot",
            )
            orientations = {
                field["name"]: field["orientation"]
                for field in pivot["fields"]
            }
            assert orientations["Region"] == 1
            assert orientations["Category"] == 2
            assert orientations["Channel"] == 3
            assert pivot["data_fields"][0]["aggregate"] == "avg"
            assert pivot["values"][-1][-1] == 25
            await client.call(
                "wps_et_ranges",
                "write_values",
                document=identity,
                sheet="Source",
                address="D2",
                values=[[50]],
            )
            await client.call(
                "wps_et_pivot_tables",
                "refresh",
                document=identity,
                sheet="Reports",
                pivot="SalesPivot",
            )
            refreshed = await client.call(
                "wps_et_pivot_tables",
                "read",
                document=identity,
                sheet="Reports",
                pivot="SalesPivot",
            )
            assert refreshed["values"][-1][-1] == 35
            identity = await save_and_reopen(
                client, identity, tmp_path / "et-pivot.xlsx"
            )
            persisted = await client.call(
                "wps_et_pivot_tables",
                "read",
                document=identity,
                sheet="Reports",
                pivot="SalesPivot",
            )
            assert persisted["values"][-1][-1] == 35
            assert persisted["data_fields"][0]["aggregate"] == "avg"
            await client.call(
                "wps_et_pivot_tables",
                "delete",
                document=identity,
                sheet="Reports",
                pivot="SalesPivot",
            )
            assert (
                await client.call(
                    "wps_et_pivot_tables",
                    "list",
                    document=identity,
                    sheet="Reports",
                )
            )["pivots"] == []
            source = await client.call(
                "wps_et_ranges",
                "read",
                document=identity,
                sheet="Source",
                address="D2:D5",
            )
            assert source["values"] == [[50], [20], [30], [40]]
            cleared = await client.call(
                "wps_et_ranges",
                "read",
                document=identity,
                sheet="Reports",
                address=persisted["address"],
            )
            assert all(
                value is None for row in cleared["values"] for value in row
            )

    anyio.run(scenario)


def test_native_et_structured_table_preserve_and_delete(
    tmp_path: Path,
) -> None:
    async def scenario() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="et"
            )
            identity = created["full_name"]
            values = [
                ["Item", "Quantity", "Price"],
                ["Alpha", 1, 10],
                ["Beta", 2, 20],
                ["Gamma", 3, 30],
                ["Delta", 4, 40],
            ]
            await client.call(
                "wps_et_ranges",
                "write_values",
                document=identity,
                sheet=1,
                address="A1:C5",
                values=values,
            )
            await client.call(
                "wps_et_ranges",
                "write_formulas",
                document=identity,
                sheet=1,
                address="E1:E3",
                formulas=[["=A1"], ["=C2"], ["=C5"]],
            )
            await client.call(
                "wps_et_tables",
                "create",
                document=identity,
                sheet=1,
                address="A1:C4",
                name="SalesTable",
            )
            initial = await client.call(
                "wps_et_tables",
                "read",
                document=identity,
                sheet=1,
                table="SalesTable",
            )
            assert initial["values"] == values[:4]
            adjacent = await client.call(
                "wps_et_ranges",
                "read",
                document=identity,
                sheet=1,
                address="A5:C5",
            )
            assert adjacent["values"] == [values[-1]]
            references = await client.call(
                "wps_et_ranges",
                "read",
                document=identity,
                sheet=1,
                address="E1:E3",
                include_formulas=True,
            )
            assert references["formulas"] == [["=A1"], ["=C2"], ["=C5"]]
            resized = await client.call(
                "wps_et_tables",
                "resize",
                document=identity,
                sheet=1,
                table="SalesTable",
                address="A1:C5",
            )
            assert resized["rows"] == 5 and resized["headers"] == values[0]
            await client.call(
                "wps_et_tables",
                "configure",
                document=identity,
                sheet=1,
                table="SalesTable",
                name="SalesFinal",
                style="TableStyleMedium2",
                banded_rows=False,
                banded_columns=True,
                first_column=True,
                last_column=True,
            )
            path = tmp_path / "et-table.xlsx"
            identity = await save_and_reopen(client, identity, path)
            persisted = await client.call(
                "wps_et_tables",
                "read",
                document=identity,
                sheet=1,
                table="SalesFinal",
            )
            assert persisted["values"] == values
            assert persisted["headers"] == values[0]
            with ZipFile(path) as archive:
                entries = [
                    name
                    for name in archive.namelist()
                    if name.startswith("xl/tables/") and name.endswith(".xml")
                ]
                assert len(entries) == 1
                table = ElementTree.fromstring(archive.read(entries[0]))
            assert table.attrib["name"] == "SalesFinal"
            style = table.find(
                "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}tableStyleInfo"
            )
            assert (
                style is not None
                and style.attrib["name"] == "TableStyleMedium2"
            )
            assert style.attrib.get("showRowStripes", "0") in {"0", "false"}
            for attribute in (
                "showColumnStripes",
                "showFirstColumn",
                "showLastColumn",
            ):
                assert style.attrib[attribute] in {"1", "true"}
            await client.call(
                "wps_et_tables",
                "unlist",
                document=identity,
                sheet=1,
                table="SalesFinal",
            )
            assert (
                await client.call(
                    "wps_et_tables", "list", document=identity, sheet=1
                )
            )["tables"] == []
            preserved = await client.call(
                "wps_et_ranges",
                "read",
                document=identity,
                sheet=1,
                address="A1:C5",
            )
            assert preserved["values"] == values
            await client.call(
                "wps_et_tables",
                "create",
                document=identity,
                sheet=1,
                address="A1:C5",
                name="SalesAgain",
            )
            await client.call(
                "wps_et_tables",
                "delete",
                document=identity,
                sheet=1,
                table="SalesAgain",
                delete_data=True,
            )
            assert (
                await client.call(
                    "wps_et_tables", "list", document=identity, sheet=1
                )
            )["tables"] == []
            cleared = await client.call(
                "wps_et_ranges",
                "read",
                document=identity,
                sheet=1,
                address="A1:C5",
            )
            assert cleared["values"] == [[None, None, None] for _ in range(5)]

    anyio.run(scenario)


SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"


async def save_workbook(
    client: NativeClient, identity: str, path: Path
) -> str:
    """Save workbook state while retaining cleanup ownership."""
    saved = await client.call(
        "wps_documents",
        "save",
        component="et",
        document=identity,
        file_path=str(path),
    )
    assert saved["saved"] and path.stat().st_size > 0
    return str(saved["full_name"])


def read_xml(path: Path, member: str) -> ElementTree.Element:
    with ZipFile(path) as archive:
        return ElementTree.fromstring(archive.read(member))


def required_xml(
    element: ElementTree.Element, path: str
) -> ElementTree.Element:
    child = element.find(path)
    assert child is not None, path
    return child


def test_native_et_sheet_lifecycle_and_frozen_panes(tmp_path: Path) -> None:
    async def scenario() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="et"
            )
            identity = created["full_name"]
            original = await client.call(
                "wps_et_sheets", "list", document=identity
            )
            assert original["sheets"] and original["sheets"][0]["visible"]
            added = await client.call(
                "wps_et_sheets",
                "add",
                document=identity,
                name="Temporary native sheet",
            )
            listed = await client.call(
                "wps_et_sheets", "list", document=identity
            )
            assert len(listed["sheets"]) == len(original["sheets"]) + 1
            assert listed["sheets"][-1]["name"] == added["name"]
            deleted = await client.call(
                "wps_et_sheets",
                "delete",
                document=identity,
                sheet=added["name"],
            )
            assert deleted["deleted"] == added["name"]
            assert (
                await client.call("wps_et_sheets", "list", document=identity)
            )["sheets"] == (original["sheets"])
            await client.call(
                "wps_et_ranges",
                "write_values",
                document=identity,
                address="A1:A2",
                values=[[19], [23]],
            )
            await client.call(
                "wps_et_ranges",
                "write_formulas",
                document=identity,
                address="A3",
                formulas=[["=SUM(A1:A2)"]],
            )
            recalculated = await client.call(
                "wps_et_sheets", "recalculate", document=identity
            )
            assert recalculated["recalculated_sheets"] == len(
                original["sheets"]
            )
            assert (
                await client.call(
                    "wps_et_ranges", "read", document=identity, address="A3"
                )
            )["values"] == [[42]]
            await client.call(
                "wps_et_sheets", "freeze", document=identity, rows=2, columns=1
            )
            frozen = tmp_path / "frozen.xlsx"
            identity = await save_workbook(client, identity, frozen)
            pane = read_xml(frozen, "xl/worksheets/sheet1.xml").find(
                f".//{{{SHEET_NS}}}pane"
            )
            assert pane is not None and pane.get("state", "").startswith(
                "frozen"
            )
            assert float(pane.attrib["xSplit"]) == 1
            assert float(pane.attrib["ySplit"]) == 2
            await client.call(
                "wps_et_sheets", "freeze", document=identity, rows=0, columns=0
            )
            unfrozen = tmp_path / "unfrozen.xlsx"
            identity = await save_workbook(client, identity, unfrozen)
            pane = read_xml(unfrozen, "xl/worksheets/sheet1.xml").find(
                f".//{{{SHEET_NS}}}pane"
            )
            assert pane is None or pane.get("state") not in {
                "frozen",
                "frozenSplit",
            }

    anyio.run(scenario)


def test_native_et_format_merge_and_clear_modes(tmp_path: Path) -> None:
    async def scenario() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="et"
            )
            identity = created["full_name"]
            await client.call(
                "wps_et_ranges",
                "write_values",
                document=identity,
                address="A1:B2",
                values=[["标题", None], [7, 8]],
            )
            await client.call(
                "wps_et_ranges",
                "format",
                document=identity,
                address="A1:B2",
                font_size=14.0,
                bold=True,
                italic=True,
                font_color="#123456",
                fill_color="#FFE699",
                horizontal_alignment="center",
                vertical_alignment="center",
                wrap_text=True,
                border_style="continuous",
                border_color="#ABCDEF",
                column_width=20.0,
                row_height=26.0,
            )
            await client.call(
                "wps_et_ranges", "merge", document=identity, address="A1:B1"
            )
            path = tmp_path / "formatted.xlsx"
            identity = await save_workbook(client, identity, path)
            sheet = read_xml(path, "xl/worksheets/sheet1.xml")
            merged = required_xml(sheet, f".//{{{SHEET_NS}}}mergeCell")
            assert merged.attrib["ref"] == "A1:B1"
            cell = sheet.find(f".//{{{SHEET_NS}}}c[@r='A1']")
            assert cell is not None
            styles = read_xml(path, "xl/styles.xml")
            style = required_xml(styles, f"{{{SHEET_NS}}}cellXfs")[
                int(cell.attrib["s"])
            ]
            font = required_xml(styles, f"{{{SHEET_NS}}}fonts")[
                int(style.attrib["fontId"])
            ]
            assert font.find(f"{{{SHEET_NS}}}b") is not None
            assert font.find(f"{{{SHEET_NS}}}i") is not None
            size = required_xml(font, f"{{{SHEET_NS}}}sz")
            assert float(size.attrib["val"]) == 14
            alignment = style.find(f"{{{SHEET_NS}}}alignment")
            assert alignment is not None
            assert alignment.attrib["horizontal"] == "center"
            assert alignment.attrib["wrapText"] in {"1", "true"}
            await client.call(
                "wps_et_ranges",
                "merge",
                document=identity,
                address="A1:B1",
                merge=False,
            )
            for mode, address in (
                ("formats", "A1:B1"),
                ("contents", "A2"),
                ("all", "B2"),
            ):
                cleared = await client.call(
                    "wps_et_ranges",
                    "clear",
                    document=identity,
                    address=address,
                    mode=mode,
                )
                assert cleared["cleared"] == mode
            remaining = await client.call(
                "wps_et_ranges", "read", document=identity, address="A1:B2"
            )
            assert remaining["values"] == [["标题", None], [None, None]]
            await client.call(
                "wps_et_ranges",
                "format",
                document=identity,
                address="A1:B2",
                auto_fit_columns=True,
                auto_fit_rows=True,
            )
            final = tmp_path / "cleared.xlsx"
            identity = await save_workbook(client, identity, final)
            assert (
                read_xml(final, "xl/worksheets/sheet1.xml").find(
                    f".//{{{SHEET_NS}}}mergeCell"
                )
                is None
            )

    anyio.run(scenario)


def test_native_et_insert_and_delete_bands_preserve_data() -> None:
    async def scenario() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="et"
            )
            identity = created["full_name"]
            values = [["A", "B"], [11, 12], [21, 22]]
            await client.call(
                "wps_et_ranges",
                "write_values",
                document=identity,
                address="A1:B3",
                values=values,
            )
            for axis, inserted_address, expected in (
                ("rows", "A1:B4", [values[0], [None, None], *values[1:]]),
                (
                    "columns",
                    "A1:C3",
                    [[row[0], None, row[1]] for row in values],
                ),
            ):
                await client.call(
                    "wps_et_ranges",
                    "edit_bands",
                    document=identity,
                    axis=axis,
                    action="insert",
                    start=2,
                )
                inserted = await client.call(
                    "wps_et_ranges",
                    "read",
                    document=identity,
                    address=inserted_address,
                )
                assert inserted["values"] == expected
                await client.call(
                    "wps_et_ranges",
                    "edit_bands",
                    document=identity,
                    axis=axis,
                    action="delete",
                    start=2,
                )
                restored = await client.call(
                    "wps_et_ranges", "read", document=identity, address="A1:B3"
                )
                assert restored["values"] == values

    anyio.run(scenario)


def test_native_et_sort_find_filter_and_clear(tmp_path: Path) -> None:
    async def scenario() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="et"
            )
            identity = created["full_name"]
            values = [
                ["Name", "Amount"],
                ["Beta", 20],
                ["Alpha", 10],
                ["Gamma", 30],
            ]
            await client.call(
                "wps_et_ranges",
                "write_values",
                document=identity,
                address="A1:B4",
                values=values,
            )
            await client.call(
                "wps_et_ranges",
                "sort",
                document=identity,
                address="A1:B4",
                key_column=2,
                descending=True,
            )
            sorted_values = await client.call(
                "wps_et_ranges", "read", document=identity, address="A1:B4"
            )
            assert sorted_values["values"] == [
                values[0],
                values[3],
                values[1],
                values[2],
            ]
            found = await client.call(
                "wps_et_ranges",
                "find",
                document=identity,
                address="A1:B4",
                text="Beta",
                whole_cell=True,
                match_case=True,
            )
            assert len(found["matches"]) == 1 and not found["truncated"]
            assert found["matches"][0]["address"] == "$A$3"
            assert found["matches"][0]["value"] == "Beta"
            missing = await client.call(
                "wps_et_ranges",
                "find",
                document=identity,
                address="A1:B4",
                text="Absent",
            )
            assert missing["matches"] == []
            await client.call(
                "wps_et_ranges",
                "filter",
                document=identity,
                address="A1:B4",
                field=2,
                criteria=">15",
            )
            filtered = tmp_path / "filtered.xlsx"
            identity = await save_workbook(client, identity, filtered)
            sheet = read_xml(filtered, "xl/worksheets/sheet1.xml")
            hidden = {
                row.attrib["r"]
                for row in sheet.findall(f".//{{{SHEET_NS}}}row")
                if row.get("hidden") in {"1", "true"}
            }
            assert hidden == {"4"}
            await client.call(
                "wps_et_ranges", "clear_filter", document=identity
            )
            unfiltered = tmp_path / "unfiltered.xlsx"
            identity = await save_workbook(client, identity, unfiltered)
            assert not any(
                row.get("hidden") in {"1", "true"}
                for row in read_xml(
                    unfiltered, "xl/worksheets/sheet1.xml"
                ).findall(f".//{{{SHEET_NS}}}row")
            )
            assert (
                await client.call(
                    "wps_et_ranges", "read", document=identity, address="A1:B4"
                )
            )["values"] == sorted_values["values"]

    anyio.run(scenario)


def test_native_et_chart_source_configuration_and_delete(
    tmp_path: Path,
) -> None:
    async def scenario() -> None:
        async with native_client() as client:
            created = await client.call(
                "wps_documents", "create", component="et"
            )
            identity = created["full_name"]
            await client.call(
                "wps_et_ranges",
                "write_values",
                document=identity,
                address="A1:C4",
                values=[
                    ["Month", "Actual", "Target"],
                    ["Jan", 2, 3],
                    ["Feb", 4, 5],
                    ["Mar", 6, 7],
                ],
            )
            added = await client.call(
                "wps_et_charts",
                "add",
                document=identity,
                address="A1:B4",
                name="NativeSales",
                chart_type="column",
                title="Before update",
            )
            assert added["chart"] == "NativeSales"
            charts = (
                await client.call("wps_et_charts", "list", document=identity)
            )["charts"]
            assert len(charts) == 1 and charts[0]["chart_type"] == 51
            await client.call(
                "wps_et_charts",
                "configure",
                document=identity,
                chart="NativeSales",
                address="A1:C4",
                chart_type="line",
                title="Native trend",
                show_legend=False,
            )
            charts = (
                await client.call("wps_et_charts", "list", document=identity)
            )["charts"]
            assert len(charts) == 1 and charts[0]["chart_type"] == 4
            path = tmp_path / "charts.xlsx"
            identity = await save_workbook(client, identity, path)
            with ZipFile(path) as archive:
                entries = [
                    name
                    for name in archive.namelist()
                    if name.startswith("xl/charts/chart")
                    and name.endswith(".xml")
                ]
                assert len(entries) == 1
                chart = ElementTree.fromstring(archive.read(entries[0]))
            line = chart.find(f".//{{{CHART_NS}}}lineChart")
            assert (
                line is not None
                and len(line.findall(f"{{{CHART_NS}}}ser")) == 2
            )
            assert "Native trend" in "".join(chart.itertext())
            assert chart.find(f".//{{{CHART_NS}}}legend") is None
            removed = await client.call(
                "wps_et_charts",
                "delete",
                document=identity,
                chart="NativeSales",
            )
            assert removed["deleted"] == "NativeSales"
            assert (
                await client.call("wps_et_charts", "list", document=identity)
            )["charts"] == []

    anyio.run(scenario)
