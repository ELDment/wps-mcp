"""Verify writer merge behavior and native preflight boundaries."""

from dataclasses import dataclass
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from pydantic import ValidationError

from wps_mcp.errors import WpsError
from wps_mcp.models import DocumentInput
from wps_mcp.native.context import CALL_DOCUMENT
from wps_mcp.native.windows import ComSession
from wps_mcp.writer import merge


@dataclass
class MergeHarness:
    document: Mock
    source: Mock
    application: Mock
    session: ComSession
    previous_window: Mock
    documents: list[Mock]


@pytest.fixture
def native(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> MergeHarness:
    path = tmp_path / "records.csv"
    path.write_text(
        "Name,City\nAlice,Shanghai\nBob,Beijing\n", encoding="utf-8"
    )
    source = Mock()
    source.Name = str(path)
    source.RecordCount = 2
    source.FirstRecord = 2
    source.LastRecord = 2
    source.FieldNames.Count = 2
    source.FieldNames.Item.side_effect = lambda index: SimpleNamespace(
        Name=["Name", "City"][index - 1]
    )
    document = Mock()
    document.Name = "Template.docx"
    document.FullName = str(tmp_path / "Template.docx")
    document.Saved = False
    document.Content = SimpleNamespace(Start=0, End=20)
    document.MailMerge.State = 2
    document.MailMerge.MainDocumentType = 0
    document.MailMerge.DataSource = source
    document.MailMerge.Fields.Count = 0
    document.MailMerge.Destination = 2
    document.MailMerge.SuppressBlankLines = False
    documents = [document]
    application = Mock()
    application.DisplayAlerts = -1
    application.ActiveDocument = Mock(FullName="Wrong active document.docx")
    previous_window = Mock()
    application.ActiveWindow = previous_window
    previous_window.Activate.side_effect = lambda: setattr(
        application, "ActiveWindow", previous_window
    )
    target_window = Mock()
    target_window.Activate.side_effect = lambda: setattr(
        application, "ActiveWindow", target_window
    )
    document.Windows.Count = 1
    document.Windows.Item.return_value = target_window
    session = ComSession("writer")
    monkeypatch.setattr(
        session, "get_application", lambda allow_launch=False: application
    )
    monkeypatch.setattr(session, "get_documents", lambda: list(documents))
    monkeypatch.setattr(merge, "get_session", lambda component: session)
    return MergeHarness(
        document, source, application, session, previous_window, documents
    )


def test_inspect_reads_schema_without_advancing_record(
    native: MergeHarness,
) -> None:
    native.document.MailMerge.Fields.Count = 1
    native.document.MailMerge.Fields.Item.return_value = SimpleNamespace(
        Type=59,
        Code=SimpleNamespace(Text=" MERGEFIELD Name ", Start=3, End=20),
    )
    result = merge.inspect_merge(
        DocumentInput(document=native.document.FullName)
    )
    assert result["fields"] == [
        {
            "index": 1,
            "type": 59,
            "code": " MERGEFIELD Name ",
            "start": 3,
            "end": 20,
        }
    ]
    assert result["source"]["fields"] == ["Name", "City"]
    assert result["state"] == "main_and_source"
    assert native.source.FirstRecord == native.source.LastRecord == 2
    assert "ConnectString" not in str(result)


def test_inspect_normal_document_does_not_read_source(
    native: MergeHarness,
) -> None:
    native.document.MailMerge.State = 0
    native.document.MailMerge.DataSource = None
    assert merge.inspect_merge(DocumentInput())["source"] is None


def test_field_uses_explicit_template_and_collapsed_range(
    native: MergeHarness,
) -> None:
    wrong = Mock(Name="Other.docx", FullName="Other.docx")
    native.documents.append(wrong)
    native.document.Range.return_value = SimpleNamespace(Start=4, End=4)
    native.document.MailMerge.Fields.Add.return_value = SimpleNamespace(
        Code=SimpleNamespace(Text=" MERGEFIELD Name ", Start=5, End=22)
    )
    result = merge.insert_merge_field(
        merge.MergeFieldInput(
            document=native.document.FullName, position=4, name="Name"
        )
    )
    native.document.Range.assert_called_once_with(4, 4)
    native.document.MailMerge.Fields.Add.assert_called_once_with(
        native.document.Range.return_value, "Name"
    )
    wrong.MailMerge.Fields.Add.assert_not_called()
    assert result["document"] == native.document.FullName


def test_field_rejects_invalid_position_and_missing_column(
    native: MergeHarness,
) -> None:
    with pytest.raises(ValueError, match="within"):
        merge.insert_merge_field(
            merge.MergeFieldInput(position=21, name="Name")
        )
    native.document.Range.assert_not_called()
    with pytest.raises(ValueError, match="absent"):
        merge.insert_merge_field(
            merge.MergeFieldInput(position=3, name="Unknown")
        )
    native.document.MailMerge.Fields.Add.assert_not_called()


@pytest.mark.parametrize(
    "name", [" ", 'Name" \\* MERGEFORMAT', "line\nname", "{Name}"]
)
def test_field_rejects_field_code_injection(name: str) -> None:
    with pytest.raises(ValidationError):
        merge.MergeFieldInput(position=0, name=name)


@pytest.mark.parametrize(
    "values",
    [
        {"file_path": "source.odc"},
        {"file_path": "source.csv", "worksheet": "Sheet1"},
        {"file_path": "source.csv", "connection": "Provider=unsafe"},
    ],
)
def test_source_input_rejects_ambiguous_or_arbitrary_queries(
    values: dict[str, str],
) -> None:
    with pytest.raises(ValidationError):
        merge.MergeSourceInput.model_validate(values)


@pytest.mark.parametrize(
    "path",
    ["https://host/source.csv", r"\\host\share\source.csv", "source.csv"],
)
def test_connect_rejects_nonlocal_paths_before_native_call(
    native: MergeHarness, path: str
) -> None:
    with pytest.raises(ValueError, match="local|absolute"):
        merge.connect_merge_source(merge.MergeSourceInput(file_path=path))
    native.document.MailMerge.OpenDataSource.assert_not_called()


def test_connect_does_not_replace_an_attached_source(
    native: MergeHarness,
) -> None:
    with pytest.raises(ValueError, match="disconnect"):
        merge.connect_merge_source(
            merge.MergeSourceInput(file_path=native.source.Name)
        )
    native.document.MailMerge.OpenDataSource.assert_not_called()


def test_connect_csv_attaches_read_only_and_restores_alerts(
    native: MergeHarness,
) -> None:
    native.document.MailMerge.State = 0

    def open_source(**options: object) -> None:
        assert native.application.DisplayAlerts == 0
        assert native.document.MailMerge.MainDocumentType == 0
        native.document.MailMerge.State = 2

    native.document.MailMerge.OpenDataSource.side_effect = open_source
    result = merge.connect_merge_source(
        merge.MergeSourceInput(file_path=native.source.Name)
    )
    native.document.MailMerge.OpenDataSource.assert_called_once_with(
        Name=native.source.Name,
        ConfirmConversions=False,
        ReadOnly=True,
        LinkToSource=True,
        AddToRecentFiles=False,
        OpenExclusive=False,
    )
    assert result["source"]["record_count"] == 2
    assert native.application.DisplayAlerts == -1


def test_connect_xlsx_is_rejected_before_accessing_wps(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    session = Mock()
    monkeypatch.setattr(merge, "get_session", session)
    with pytest.raises(
        ValidationError,
        match="Export the desired worksheet to CSV before connecting; "
        "native XLSX selection cannot run unattended",
    ):
        merge.connect_merge_source(
            merge.MergeSourceInput(file_path="source.xlsx")
        )
    session.assert_not_called()
    assert (
        "worksheet"
        not in merge.MergeSourceInput.model_json_schema()["properties"]
    )


def test_connect_failure_restores_alerts_without_replaying(
    native: MergeHarness,
) -> None:
    native.document.MailMerge.State = 0
    native.document.MailMerge.OpenDataSource.side_effect = RuntimeError(
        "native failure"
    )
    with pytest.raises(RuntimeError, match="native failure"):
        merge.connect_merge_source(
            merge.MergeSourceInput(file_path=native.source.Name)
        )
    assert native.application.DisplayAlerts == -1
    assert native.document.MailMerge.OpenDataSource.call_count == 1


def test_connect_rejects_silent_failure(native: MergeHarness) -> None:
    native.document.MailMerge.State = 0
    with pytest.raises(WpsError, match="did not attach"):
        merge.connect_merge_source(
            merge.MergeSourceInput(file_path=native.source.Name)
        )


def test_execute_uses_new_document_difference_and_restores_state(
    native: MergeHarness,
) -> None:
    result_document = Mock(Name="Letters1", FullName="Letters1", Saved=False)
    unrelated = Mock(Name="Unrelated", FullName="Unrelated")
    native.documents.append(unrelated)
    native.session.target_name = unrelated.FullName

    def execute(*, Pause: bool) -> None:  # noqa: N803 - native COM keyword
        assert Pause is False
        assert CALL_DOCUMENT.get() == native.document.FullName
        assert native.document.MailMerge.Destination == 0
        assert native.source.FirstRecord == 1
        assert native.source.LastRecord == -16
        native.documents.append(result_document)

    native.document.MailMerge.Execute.side_effect = execute
    token = CALL_DOCUMENT.set("previous-call-target")
    try:
        result = merge.execute_merge(
            merge.MergeExecuteInput(document=native.document.FullName)
        )
        assert CALL_DOCUMENT.get() == "previous-call-target"
    finally:
        CALL_DOCUMENT.reset(token)
    assert result["full_name"] == "Letters1"
    assert result["source_document"] == native.document.FullName
    assert result["last_record"] == 2
    assert native.session.target_name == unrelated.FullName
    assert native.document.MailMerge.Destination == 2
    assert native.document.MailMerge.SuppressBlankLines is False
    assert native.source.FirstRecord == native.source.LastRecord == 2
    assert native.application.DisplayAlerts == -1
    assert native.application.ActiveWindow is native.previous_window
    native.previous_window.Activate.assert_called_once_with()
    native.application.ActiveDocument.Activate.assert_not_called()


def test_execute_failure_restores_state_and_never_retries(
    native: MergeHarness,
) -> None:
    native.document.MailMerge.Execute.side_effect = RuntimeError(
        "merge failed"
    )
    with pytest.raises(RuntimeError, match="merge failed"):
        merge.execute_merge(merge.MergeExecuteInput())
    assert native.document.MailMerge.Destination == 2
    assert native.source.FirstRecord == native.source.LastRecord == 2
    assert native.application.DisplayAlerts == -1
    assert native.application.ActiveWindow is native.previous_window
    assert native.document.MailMerge.Execute.call_count == 1


@pytest.mark.parametrize("created_count", [0, 2])
def test_execute_refuses_to_guess_ambiguous_output(
    native: MergeHarness, created_count: int
) -> None:
    native.document.MailMerge.Execute.side_effect = lambda **options: (
        native.documents.extend(
            [
                Mock(FullName=f"Letters{index}")
                for index in range(created_count)
            ]
        )
    )
    with pytest.raises(WpsError, match="Expected one new"):
        merge.execute_merge(merge.MergeExecuteInput())
    assert native.document.MailMerge.Destination == 2


@pytest.mark.parametrize(
    "count,first,last", [(0, 1, None), (2, 3, None), (2, 1, 3)]
)
def test_execute_bounds_records_before_mutation(
    native: MergeHarness, count: int, first: int, last: int | None
) -> None:
    native.source.RecordCount = count
    with pytest.raises(ValueError, match="Record interval"):
        merge.execute_merge(
            merge.MergeExecuteInput(first_record=first, last_record=last)
        )
    native.document.MailMerge.Execute.assert_not_called()
    assert native.document.MailMerge.Destination == 2


def test_execute_rejects_nonlocal_existing_source(
    native: MergeHarness,
) -> None:
    native.source.Name = "https://example.com/data.csv"
    with pytest.raises(ValueError, match="local"):
        merge.execute_merge(merge.MergeExecuteInput())
    native.document.MailMerge.Execute.assert_not_called()


def test_execute_rejects_email_template(native: MergeHarness) -> None:
    native.document.MailMerge.MainDocumentType = 4
    with pytest.raises(ValueError, match="form-letter"):
        merge.execute_merge(merge.MergeExecuteInput())
    native.document.MailMerge.Execute.assert_not_called()


def test_execute_unknown_record_count_passes_native_last_record(
    native: MergeHarness,
) -> None:
    native.source.RecordCount = -1
    native.document.MailMerge.Execute.side_effect = lambda **options: (
        native.documents.append(
            Mock(Name="Letters1", FullName="Letters1", Saved=False)
        )
    )
    result = merge.execute_merge(merge.MergeExecuteInput())
    assert result["last_record"] is None


def test_disconnect_closes_only_the_data_source(native: MergeHarness) -> None:
    native.source.Close.side_effect = lambda: setattr(
        native.document.MailMerge, "State", 1
    )
    assert (
        merge.disconnect_merge_source(DocumentInput())["disconnected"] is True
    )
    native.source.Close.assert_called_once_with()
    native.document.Close.assert_not_called()
    assert native.document.MailMerge.MainDocumentType == 0
    assert (
        merge.disconnect_merge_source(DocumentInput())["disconnected"] is False
    )
    assert native.source.Close.call_count == 1


def test_disconnect_rejects_silent_failure(native: MergeHarness) -> None:
    with pytest.raises(WpsError, match="still reports"):
        merge.disconnect_merge_source(DocumentInput())


def test_missing_native_support_is_reported_without_fallback(
    native: MergeHarness,
) -> None:
    native.document.MailMerge.Execute.side_effect = NotImplementedError(
        "host stub"
    )
    with pytest.raises(WpsError, match="does not implement"):
        merge.execute_merge(merge.MergeExecuteInput())
    assert native.document.MailMerge.Execute.call_count == 1
    assert native.application.DisplayAlerts == -1


def test_execute_input_rejects_reversed_records() -> None:
    with pytest.raises(ValidationError, match="last_record"):
        merge.MergeExecuteInput(first_record=2, last_record=1)


def test_unnamed_csv_source_connect_inspect_and_execute(
    native: MergeHarness,
) -> None:
    path = Path(native.source.Name)
    native.source.Name = ""
    native.source.ConnectString = (
        "Provider=Microsoft.ACE.OLEDB.12.0;\t\tData Source="
        + str(path.parent)
        + ";\t\tExtended Properties='text; HDR = Yes; FMT = Delimited'"
    )
    native.source.QueryString = "SELECT * FROM `records`"
    native.document.MailMerge.State = 0
    native.document.MailMerge.OpenDataSource.side_effect = lambda **options: (
        setattr(native.document.MailMerge, "State", 2)
    )
    connected = merge.connect_merge_source(
        merge.MergeSourceInput(file_path=str(path))
    )
    assert connected["source"]["file_path"] == str(path)
    inspected = merge.inspect_merge(DocumentInput())
    assert inspected["source"]["file_path"] == str(path)
    assert "Provider" not in str(inspected)
    native.document.MailMerge.Execute.side_effect = lambda **options: (
        native.documents.append(
            Mock(Name="Letters1", FullName="Letters1", Saved=False)
        )
    )
    assert (
        merge.execute_merge(merge.MergeExecuteInput())["full_name"]
        == "Letters1"
    )
    assert native.document.MailMerge.Execute.call_count == 1


def test_unnamed_source_accepts_quoted_local_csv(tmp_path: Path) -> None:
    path = tmp_path / "source;quoted.csv"
    path.touch()
    source = SimpleNamespace(
        Name="",
        ConnectString=(
            f'Provider=Microsoft.ACE.OLEDB.12.0; Data Source="{path}";'
        ),
    )
    assert merge.source_path(source) == path


@pytest.mark.parametrize(
    "query",
    [
        "SELECT * FROM [records.csv]",
        'select * from "records";',
        "SELECT * FROM records",
    ],
)
def test_unnamed_csv_accepts_only_a_single_table(
    tmp_path: Path, query: str
) -> None:
    path = tmp_path / "records.csv"
    path.touch()
    source = SimpleNamespace(
        Name="", ConnectString=f"Data Source={tmp_path};", QueryString=query
    )
    assert merge.source_path(source) == path


@pytest.mark.parametrize(
    "query",
    [
        "SELECT Name FROM `records`",
        "SELECT * FROM `records` WHERE 1=1",
        "SELECT * FROM `records`; DELETE FROM records",
        "SELECT * FROM `records` JOIN Other ON 1=1",
        "SELECT * FROM `../records`",
        r"SELECT * FROM `..\records`",
        r"SELECT * FROM `C:\records`",
        r"SELECT * FROM `\\server\records`",
        "SELECT * FROM `https://example.com/records`",
        "SELECT * FROM `records` IN 'remote'",
        "SELECT * FROM ``",
        "SELECT * FROM `records\n`",
        "SELECT * FROM [records].[other]",
        "SELECT * FROM `records` -- comment",
        "SELECT * FROM `missing`",
    ],
)
def test_unnamed_csv_rejects_unsafe_query_before_execute(
    native: MergeHarness, tmp_path: Path, query: str
) -> None:
    native.source.Name = ""
    native.source.ConnectString = f"Data Source={tmp_path};"
    native.source.QueryString = query
    with pytest.raises(ValueError):
        merge.execute_merge(merge.MergeExecuteInput())
    native.document.MailMerge.Execute.assert_not_called()
    assert native.document.MailMerge.Destination == 2


@pytest.mark.parametrize(
    "connection",
    [
        "",
        "Provider=Microsoft.ACE.OLEDB.12.0;",
        "Data Source=https://example.com/records.csv;",
        r"Data Source=\\server\share\records.csv;",
        "Data Source=relative.csv;",
        "Data Source='unterminated;",
        'Data Source="unterminated;',
        "Data Source='{directory}'; Data Source='{directory}';",
        "Data Source={directory}; trailing garbage",
        "Extended Properties='text; Data Source={directory}';",
    ],
)
def test_unnamed_source_rejects_invalid_connection(
    native: MergeHarness, tmp_path: Path, connection: str
) -> None:
    native.source.Name = ""
    native.source.ConnectString = connection.replace(
        "{directory}", str(tmp_path)
    )
    native.source.QueryString = "SELECT * FROM `records`"
    with pytest.raises(ValueError):
        merge.execute_merge(merge.MergeExecuteInput())
    native.document.MailMerge.Execute.assert_not_called()


def test_unnamed_source_cannot_hide_unsafe_source_inside_quoted_property(
    tmp_path: Path,
) -> None:
    source = SimpleNamespace(
        Name="",
        ConnectString=(
            f"Extended Properties='text; Data Source={tmp_path}'; "
            "Data Source=https://example.com/remote.csv;"
        ),
        QueryString="SELECT * FROM `records`",
    )
    with pytest.raises(ValueError, match="local"):
        merge.source_path(source)


@pytest.mark.parametrize("extension", [".xlsx", ".accdb"])
@pytest.mark.parametrize("named", [True, False])
def test_execute_rejects_non_csv_sources(
    native: MergeHarness, tmp_path: Path, extension: str, named: bool
) -> None:
    path = tmp_path / ("database" + extension)
    path.touch()
    native.source.Name = str(path) if named else ""
    native.source.ConnectString = f"Data Source={path};"
    with pytest.raises(ValueError, match="only local"):
        merge.execute_merge(merge.MergeExecuteInput())
    native.document.MailMerge.Execute.assert_not_called()


def test_odbc_text_driver_resolves_default_directory(
    native: MergeHarness,
) -> None:
    path = Path(native.source.Name)
    password = "wps-test-password-must-not-leak"
    native.source.Name = ""
    native.source.ConnectString = (
        "DRIVER={Microsoft Text Driver (*.txt; *.csv)};\t\t"
        f"DefaultDir={path.parent};UID=Admin;PWD={password};"
        "Extensions=txt,csv,tab"
    )
    native.source.QueryString = "SELECT * FROM `records`"
    details = merge.source_details(native.source)
    assert details["file_path"] == str(path)
    assert details["record_count"] == 2
    assert password not in str(details)
    assert "DRIVER" not in str(details)


@pytest.mark.parametrize("property_name", ["DefaultDir", "Data Source"])
def test_braced_source_preserves_semicolons_and_escaped_closing_brace(
    tmp_path: Path, property_name: str
) -> None:
    directory = tmp_path / "source;with}brace"
    directory.mkdir()
    path = directory / "records.csv"
    path.touch()
    escaped = str(directory).replace("}", "}}")
    source = SimpleNamespace(
        Name="",
        ConnectString=(
            "DRIVER={Text;Driver}}Version};"
            f"{property_name}={{{escaped}}};PWD={{password;with}}}}brace}};"
        ),
        QueryString="SELECT * FROM `records`",
    )
    assert merge.source_path(source) == path


@pytest.mark.parametrize(
    "connection",
    [
        "DRIVER={Text;Driver;DefaultDir={directory};",
        "DRIVER={Text}garbage;DefaultDir={directory};",
        "DefaultDir={{directory};",
        "DefaultDir={directory};DEFAULTDIR={directory};",
        "DefaultDir={directory};Data Source={directory};",
        "DefaultDir={https://example.com/data};",
        r"DefaultDir={\\server\share};",
        "DefaultDir={relative};",
        "DRIVER={Text; DefaultDir={directory}};",
        "PWD={secret;DefaultDir={directory}};",
        " =unused;DefaultDir={directory};",
    ],
)
def test_odbc_connection_rejects_injection_and_ambiguous_paths(
    native: MergeHarness, tmp_path: Path, connection: str
) -> None:
    native.source.Name = ""
    native.source.ConnectString = connection.replace(
        "{directory}", str(tmp_path)
    )
    native.source.QueryString = "SELECT * FROM `records`"
    with pytest.raises(ValueError) as caught:
        merge.execute_merge(merge.MergeExecuteInput())
    assert "secret" not in str(caught.value)
    assert "DRIVER=" not in str(caught.value)
    native.document.MailMerge.Execute.assert_not_called()


@pytest.mark.parametrize(
    "query",
    [
        "SELECT Name FROM `records`",
        "SELECT * FROM `records`; DELETE FROM records",
        "SELECT * FROM `records` WHERE 1=1",
        "SELECT * FROM `../records`",
        r"SELECT * FROM `..\records`",
        "SELECT * FROM `https://example.com/records`",
        "SELECT * FROM `records` -- comment",
    ],
)
def test_odbc_query_keeps_single_local_csv_table_boundary(
    native: MergeHarness, tmp_path: Path, query: str
) -> None:
    native.source.Name = ""
    native.source.ConnectString = f"DRIVER={{Text}};DefaultDir={tmp_path};"
    native.source.QueryString = query
    with pytest.raises(ValueError):
        merge.execute_merge(merge.MergeExecuteInput())
    native.document.MailMerge.Execute.assert_not_called()


def test_odbc_default_directory_cannot_be_a_file(native: MergeHarness) -> None:
    native.source.ConnectString = f"DefaultDir={native.source.Name};"
    native.source.Name = ""
    native.source.QueryString = "SELECT * FROM `records`"
    with pytest.raises(ValueError, match="local source directory"):
        merge.source_path(native.source)
