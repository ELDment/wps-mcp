"""Preserve existing file bytes and never replay failed native writes."""

from collections.abc import Mapping, Set
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
from zipfile import ZipFile

import pytest
from pydantic import ValidationError

from src.common import files, paths
from src.documents import models, service
from src.native.components import SPECS


@pytest.fixture
def session(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> SimpleNamespace:
    document = Mock()
    document.Name = "source.docx"
    document.FullName = str(tmp_path / document.Name)
    document.Path = str(tmp_path)
    document.Saved = False
    application = SimpleNamespace(DisplayAlerts=7)
    current = SimpleNamespace(
        component="writer",
        spec=SPECS["writer"],
        target_name=document.FullName,
        get_document=Mock(return_value=document),
        get_documents=Mock(return_value=[document]),
        get_application=Mock(return_value=application),
    )
    monkeypatch.setattr(service, "get_session", lambda component: current)
    return current


def test_save_requires_explicit_overwrite(
    session: SimpleNamespace, tmp_path: Path
) -> None:
    path = tmp_path / "existing.docx"
    path.write_bytes(b"original")
    with pytest.raises(ValueError, match="overwrite"):
        service.save_document(
            models.SaveInput(component="writer", file_path=str(path))
        )
    assert path.read_bytes() == b"original"
    session.get_document().SaveAs.assert_not_called()


def test_save_refuses_file_created_after_destination_check(
    monkeypatch: pytest.MonkeyPatch, session: SimpleNamespace, tmp_path: Path
) -> None:
    path = tmp_path / "race.docx"
    document = session.get_document()
    document.Saved = True
    original_target = session.target_name
    destination = service.destination

    def race_after_check(
        file_path: str,
        extensions: Mapping[str, int | None] | Set[str],
        overwrite: bool = False,
    ) -> Path:
        checked = destination(file_path, extensions, overwrite)
        checked.write_bytes(b"concurrent document")
        return checked

    monkeypatch.setattr(service, "destination", race_after_check)
    with pytest.raises(FileExistsError):
        service.save_document(
            models.SaveInput(component="writer", file_path=str(path))
        )
    assert path.read_bytes() == b"concurrent document"
    assert document.Saved is True
    assert session.target_name == original_target
    assert session.get_application().DisplayAlerts == 7
    document.SaveAs.assert_not_called()
    assert not list(tmp_path.glob(".*.recovery-*"))


def test_failed_new_save_removes_only_untouched_reservation(
    session: SimpleNamespace, tmp_path: Path
) -> None:
    path = tmp_path / "new.docx"
    document = session.get_document()

    def fail_save(file_path: str, file_format: int) -> None:
        assert Path(file_path).read_bytes() == b""
        raise RuntimeError("native failure before writing")

    document.SaveAs.side_effect = fail_save
    with pytest.raises(RuntimeError, match="native state may have changed"):
        service.save_document(
            models.SaveInput(component="writer", file_path=str(path))
        )
    assert not path.exists()
    assert not list(tmp_path.glob(".*.recovery-*"))
    document.SaveAs.assert_called_once()
    assert session.get_application().DisplayAlerts == 7


def test_failed_new_save_keeps_partial_output(
    session: SimpleNamespace, tmp_path: Path
) -> None:
    path = tmp_path / "new.docx"
    document = session.get_document()

    def fail_save(file_path: str, file_format: int) -> None:
        Path(file_path).write_bytes(b"partial document")
        raise RuntimeError("native failure after writing")

    document.SaveAs.side_effect = fail_save
    with pytest.raises(RuntimeError, match="native state may have changed"):
        service.save_document(
            models.SaveInput(component="writer", file_path=str(path))
        )
    assert path.read_bytes() == b"partial document"
    assert not list(tmp_path.glob(".*.recovery-*"))
    document.SaveAs.assert_called_once()


def test_failed_new_save_does_not_remove_replacement_empty_file(
    session: SimpleNamespace, tmp_path: Path
) -> None:
    path = tmp_path / "new.docx"
    replacement = tmp_path / "other.docx"
    replacement.touch()
    replacement_identity = replacement.stat().st_ino
    document = session.get_document()

    def fail_save(file_path: str, file_format: int) -> None:
        replacement.replace(file_path)
        raise RuntimeError("another process replaced the reservation")

    document.SaveAs.side_effect = fail_save
    with pytest.raises(RuntimeError, match="native state may have changed"):
        service.save_document(
            models.SaveInput(component="writer", file_path=str(path))
        )
    assert path.is_file()
    assert path.stat().st_ino == replacement_identity
    assert path.read_bytes() == b""


def test_same_path_save_does_not_require_exclusive_creation(
    session: SimpleNamespace,
) -> None:
    document = session.get_document()
    path = Path(document.FullName)
    path.write_bytes(b"original document")

    def save() -> None:
        with ZipFile(path, "w") as archive:
            archive.writestr("word/document.xml", "<document />")
        document.Saved = True

    document.Save.side_effect = save
    service.save_document(models.SaveInput(component="writer"))
    document.Save.assert_called_once_with()
    document.SaveAs.assert_not_called()
    assert document.Saved is True


def test_save_rejects_another_open_document(
    session: SimpleNamespace, tmp_path: Path
) -> None:
    path = tmp_path / "other.docx"
    other = SimpleNamespace(FullName=str(path))
    session.get_documents.return_value.append(other)
    with pytest.raises(ValueError, match="another open document"):
        service.save_document(
            models.SaveInput(
                component="writer", file_path=str(path), overwrite=True
            )
        )
    session.get_document().SaveAs.assert_not_called()


def test_save_requires_a_saved_path(session: SimpleNamespace) -> None:
    document = session.get_document()
    document.FullName = "Document1"
    document.Path = ""
    with pytest.raises(ValueError, match="supply file_path"):
        service.save_document(models.SaveInput(component="writer"))
    document.Save.assert_not_called()


def test_save_failure_restores_bytes_alerts_and_dirty_state(
    session: SimpleNamespace, tmp_path: Path
) -> None:
    path = tmp_path / "existing.docx"
    path.write_bytes(b"original")
    document = session.get_document()

    def fail_save(file_path: str, file_format: int) -> None:
        Path(file_path).write_bytes(b"partly written")
        document.FullName = file_path
        document.Saved = True
        raise RuntimeError("busy after mutation")

    document.SaveAs.side_effect = fail_save
    with pytest.raises(RuntimeError, match="native state may have changed"):
        service.save_document(
            models.SaveInput(
                component="writer", file_path=str(path), overwrite=True
            )
        )
    assert path.read_bytes() == b"original"
    assert document.Saved is False
    assert session.target_name == str(path)
    assert session.get_application().DisplayAlerts == 7
    assert document.SaveAs.call_count == 1
    assert not list(tmp_path.glob(".*.recovery-*"))


def test_invalid_saved_payload_restores_original(
    session: SimpleNamespace, tmp_path: Path
) -> None:
    path = tmp_path / "existing.docx"
    path.write_bytes(b"original")
    document = session.get_document()

    def wrong_payload(file_path: str, file_format: int) -> None:
        with ZipFile(file_path, "w") as archive:
            archive.writestr("xl/workbook.xml", "<workbook />")
        document.FullName = file_path
        document.Saved = True

    document.SaveAs.side_effect = wrong_payload
    with pytest.raises(
        RuntimeError, match="Existing destination bytes were preserved"
    ):
        service.save_document(
            models.SaveInput(
                component="writer", file_path=str(path), overwrite=True
            )
        )
    assert path.read_bytes() == b"original"
    assert document.SaveAs.call_count == 1


def test_failed_restore_retains_recovery_copy(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    path = tmp_path / "existing.docx"
    path.write_bytes(b"original")
    monkeypatch.setattr(
        files.os, "replace", Mock(side_effect=PermissionError("locked"))
    )
    with pytest.raises(RuntimeError, match="Original bytes are preserved at"):
        with files.protect_existing_file(path):
            path.write_bytes(b"damaged")
            raise RuntimeError("native failure")
    recovery = list(tmp_path.glob(".*.recovery-*"))
    assert len(recovery) == 1
    assert recovery[0].read_bytes() == b"original"


def test_save_success_updates_target(
    session: SimpleNamespace, tmp_path: Path
) -> None:
    path = tmp_path / "new.docx"
    document = session.get_document()

    def save(file_path: str, file_format: int) -> None:
        assert file_format == 12
        with ZipFile(file_path, "w") as archive:
            archive.writestr("word/document.xml", "<document />")
        document.FullName = file_path
        document.Name = Path(file_path).name
        document.Saved = True

    document.SaveAs.side_effect = save
    result = service.save_document(
        models.SaveInput(component="writer", file_path=str(path))
    )
    assert result == {
        "name": "new.docx",
        "full_name": str(path),
        "saved": True,
    }
    assert session.target_name == str(path)
    assert session.get_application().DisplayAlerts == 7


def test_close_unsaved_requires_explicit_choice(
    session: SimpleNamespace,
) -> None:
    with pytest.raises(ValueError, match="unsaved changes"):
        service.close_document(models.CloseInput(component="writer"))
    session.get_document().Close.assert_not_called()


def test_close_modes_are_mutually_exclusive() -> None:
    with pytest.raises(ValidationError, match="mutually exclusive"):
        models.CloseInput(
            component="wpp", save_changes=True, discard_changes=True
        )


def test_failed_wpp_close_restores_dirty_state(
    session: SimpleNamespace,
) -> None:
    session.component = "wpp"
    document = session.get_document()
    original_name = session.target_name
    document.Close.side_effect = RuntimeError("close failed")
    with pytest.raises(RuntimeError, match="close failed"):
        service.close_document(
            models.CloseInput(component="wpp", discard_changes=True)
        )
    assert document.Saved is False
    assert session.target_name == original_name
    document.Close.assert_called_once_with()


def test_successful_close_clears_target(session: SimpleNamespace) -> None:
    expected = session.target_name
    result = service.close_document(
        models.CloseInput(component="writer", discard_changes=True)
    )
    assert result == {"closed": expected}
    assert session.target_name is None
    session.get_document().Close.assert_called_once_with(False)


def test_invalid_pdf_export_leaves_existing_file(
    session: SimpleNamespace, tmp_path: Path
) -> None:
    path = tmp_path / "existing.pdf"
    path.write_bytes(b"original")
    document = session.get_document()

    def export(file_path: str, file_format: int) -> None:
        Path(file_path).write_bytes(b"not a PDF")

    document.ExportAsFixedFormat.side_effect = export
    with pytest.raises(RuntimeError, match="PDF"):
        service.export_pdf(
            models.ExportInput(
                component="writer", file_path=str(path), overwrite=True
            )
        )
    assert path.read_bytes() == b"original"
    document.ExportAsFixedFormat.assert_called_once()


def test_pdf_export_refuses_file_created_during_native_export(
    session: SimpleNamespace, tmp_path: Path
) -> None:
    path = tmp_path / "race.pdf"
    document = session.get_document()

    def export(file_path: str, file_format: int) -> None:
        Path(file_path).write_bytes(b"%PDF-1.7\nexported document")
        path.write_bytes(b"concurrent document")

    document.ExportAsFixedFormat.side_effect = export
    with pytest.raises(FileExistsError):
        service.export_pdf(
            models.ExportInput(component="writer", file_path=str(path))
        )
    assert path.read_bytes() == b"concurrent document"
    assert not list(tmp_path.glob("wps-export-*"))
    document.ExportAsFixedFormat.assert_called_once()


def test_pdf_export_publishes_new_file(
    session: SimpleNamespace, tmp_path: Path
) -> None:
    path = tmp_path / "new.pdf"
    document = session.get_document()
    content = b"%PDF-1.7\nexported document"

    def export(file_path: str, file_format: int) -> None:
        Path(file_path).write_bytes(content)

    document.ExportAsFixedFormat.side_effect = export
    result = service.export_pdf(
        models.ExportInput(component="writer", file_path=str(path))
    )
    assert path.read_bytes() == content
    assert result == {"file_path": str(path), "bytes": len(content)}
    assert not list(tmp_path.glob("wps-export-*"))


@pytest.mark.parametrize(
    "path",
    [
        "relative.docx",
        "https://example.com/file.docx",
        "\\\\server\\file.docx",
    ],
)
def test_local_paths_reject_remote_and_relative(path: str) -> None:
    with pytest.raises(ValueError):
        paths.local_path(path)
