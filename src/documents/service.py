"""Perform document lifecycle operations independently of MCP registration."""

import logging
import sys
from pathlib import Path
from tempfile import TemporaryDirectory

from src.common.files import (
    protect_existing_file,
    publish_file,
    verify_output,
    verify_saved_file,
)
from src.common.paths import destination, local_path
from src.common.rendering import render_pdf_page
from src.documents.models import (
    CloseInput,
    ComponentDocumentInput,
    ComponentInput,
    ConnectInput,
    CreateInput,
    ExportInput,
    OpenInput,
    PreviewInput,
    SaveInput,
)
from src.models import InputModel, JsonObject, NativeObject
from src.native.components import SPECS
from src.native.runtime import get_session
from src.native.selection import document_key
from src.native.session import WpsSession
from src.native.state import temporary_attributes
from src.presentation.formats import check_presentation_format

logger = logging.getLogger(__name__)


def describe_document(document: NativeObject) -> JsonObject:
    """Detach identity and save status from the apartment-owned document."""
    return {
        "name": str(document.Name),
        "full_name": str(document.FullName),
        "saved": bool(document.Saved),
    }


def connect(params: ConnectInput) -> JsonObject:
    """Connect to a native WPS component, launching it when needed."""
    session = get_session(params.component)
    application = session.connect(visible=params.visible)
    return {
        "component": params.component,
        "prog_id": session.spec.prog_id,
        "name": str(application.Name),
        "version": str(application.Version),
        "document_count": int(
            getattr(application, session.spec.collection).Count
        ),
    }


def list_documents(params: ComponentInput) -> JsonObject:
    """List the full_name identifiers of documents open in one component."""
    return {
        "component": params.component,
        "documents": [
            describe_document(item)
            for item in get_session(params.component).get_documents()
        ],
    }


def create_document(params: CreateInput) -> JsonObject:
    """Create a Writer document, ET workbook or WPP presentation."""
    template = (
        local_path(params.template_path, exists=True)
        if params.template_path
        else None
    )
    session = get_session(params.component)
    application = session.connect(visible=params.visible)
    collection = getattr(application, session.spec.collection)
    if params.component == "writer":
        document = (
            collection.Add(Template=str(template))
            if template
            else collection.Add()
        )
    elif params.component == "et":
        document = (
            collection.Add(str(template)) if template else collection.Add()
        )
    else:
        document = (
            collection.Open(str(template), False, True, params.visible)
            if template
            else collection.Add()
        )
    return describe_document(document)


def open_document(params: OpenInput) -> JsonObject:
    """Open a local document; reject a duplicate already open in WPS."""
    path = local_path(params.file_path, exists=True)
    session = get_session(params.component)
    if path.suffix.lower() not in session.spec.formats:
        raise ValueError(
            f"File type {path.suffix} is unsupported by {params.component}"
        )
    application = session.connect(visible=params.visible)
    collection = getattr(application, session.spec.collection)
    for document in session.get_documents():
        if document_key(document.FullName) == document_key(path):
            raise ValueError("The document is already open; use its full_name")
    if params.component == "writer":
        document = collection.Open(
            str(path), ReadOnly=params.read_only, AddToRecentFiles=False
        )
    elif params.component == "et":
        document = collection.Open(
            str(path), UpdateLinks=0, ReadOnly=params.read_only
        )
    else:
        document = collection.Open(
            str(path), ReadOnly=params.read_only, WithWindow=params.visible
        )
    return describe_document(document)


def save_document(params: SaveInput) -> JsonObject:
    """Save once through WPS, preserving existing destination bytes on failure.

    Supply file_path for unsaved documents. Replacing a different existing file
    requires overwrite. A failed native operation may already have changed WPS.
    """
    session = get_session(params.component)
    document = session.get_document(params.document)
    old_name = str(document.FullName)
    old_key = document_key(old_name)
    if params.file_path is None and not str(getattr(document, "Path", "")):
        raise ValueError("This document has no saved path; supply file_path")
    wanted = params.file_path or old_name
    path = local_path(wanted)
    same_path = document_key(path) == old_key
    if params.component == "wpp":
        check_presentation_format(document, str(path))
    for opened in session.get_documents():
        if (
            document_key(opened.FullName) == document_key(path)
            and not same_path
        ):
            raise ValueError(
                "Output belongs to another open document; "
                "close it before overwriting"
            )
    path = destination(
        str(path), session.spec.formats, params.overwrite or same_path
    )
    selected = (
        session.target_name is not None
        and document_key(session.target_name) == old_key
    )
    previous_stamp = (
        (path.stat().st_mtime_ns, path.stat().st_size)
        if path.is_file()
        else None
    )
    native_started = False
    try:
        with (
            protect_existing_file(
                path, overwrite=params.overwrite or same_path
            ),
            temporary_attributes(
                session.get_application(),
                DisplayAlerts=1 if params.component == "wpp" else 0,
            ),
        ):
            native_started = True
            if same_path:
                document.Save()
            else:
                file_format = session.spec.formats[path.suffix.lower()]
                if file_format is None:
                    document.SaveAs(str(path))
                else:
                    document.SaveAs(str(path), file_format)
            verify_saved_file(path)
            if not same_path and previous_stamp == (
                path.stat().st_mtime_ns,
                path.stat().st_size,
            ):
                raise RuntimeError(
                    "WPS did not update the existing output file"
                )
            if not bool(document.Saved):
                raise RuntimeError("WPS did not report the document as saved")
            if document_key(document.FullName) != document_key(path):
                raise RuntimeError(
                    "WPS did not adopt the requested destination"
                )
    except Exception:
        if native_started:
            try:
                document.Saved = False
            except Exception:
                logger.warning(
                    "Could not mark the failed save as unsaved", exc_info=True
                )
        raise
    finally:
        if selected and native_started:
            try:
                session.target_name = str(document.FullName)
            except Exception:
                logger.warning(
                    "Could not refresh the selected document after saving",
                    exc_info=True,
                )
    result = describe_document(document)
    if path.suffix.lower() in {".wps", ".et"}:
        result["warnings"] = [
            "WPS controls the stored encoding; "
            ".wps/.et may contain OOXML data."
        ]
    return result


def close_document(params: CloseInput) -> JsonObject:
    """Close one document with an explicit save or discard choice."""
    session = get_session(params.component)
    document = session.get_document(params.document)
    identity = describe_document(document)
    if params.save_changes:
        save_document(
            SaveInput(
                component=params.component, document=identity["full_name"]
            )
        )
    elif not identity["saved"] and not params.discard_changes:
        raise ValueError(
            "Document has unsaved changes; save first or set discard_changes"
        )
    if params.component == "wpp":
        marked_saved = params.discard_changes and not identity["saved"]
        if marked_saved:
            document.Saved = True
        try:
            document.Close()
        except Exception:
            if marked_saved:
                try:
                    document.Saved = False
                except Exception:
                    logger.warning(
                        "Could not restore unsaved state after failed close",
                        exc_info=True,
                    )
            raise
    else:
        document.Close(False)
    if session.target_name and document_key(
        session.target_name
    ) == document_key(identity["full_name"]):
        session.target_name = None
    return {"closed": identity["full_name"]}


def get_document_info(params: ComponentDocumentInput) -> JsonObject:
    """Read document identity, save state and component-specific structure."""
    document = get_session(params.component).get_document(params.document)
    info = describe_document(document)
    if params.component == "writer":
        info.update(
            paragraphs=int(document.Paragraphs.Count),
            tables=int(document.Tables.Count),
            characters=int(document.Content.End),
            pages=int(document.ComputeStatistics(2)),
        )
    elif params.component == "et":
        info["worksheets"] = [
            str(document.Worksheets.Item(index).Name)
            for index in range(1, document.Worksheets.Count + 1)
        ]
    else:
        info.update(
            slides=int(document.Slides.Count),
            width=float(document.PageSetup.SlideWidth),
            height=float(document.PageSetup.SlideHeight),
        )
    return info


def export_pdf_to(
    session: WpsSession, document: NativeObject, file_path: Path
) -> JsonObject:
    """Export through native WPS and verify the PDF output signature."""
    if session.component == "writer":
        document.ExportAsFixedFormat(str(file_path), 17)
    elif session.component == "et":
        document.ExportAsFixedFormat(0, str(file_path))
    else:
        document.ExportAsFixedFormat(str(file_path), 2, PrintRange=None)
    result = verify_output(file_path)
    with file_path.open("rb") as stream:
        if stream.read(5) != b"%PDF-":
            raise RuntimeError("WPS did not produce a PDF file")
    return result


def export_pdf(params: ExportInput) -> JsonObject:
    """Export a PDF through WPS; publish only a validated, completed file."""
    session = get_session(params.component)
    document = session.get_document(params.document)
    path = destination(params.file_path, {".pdf"}, params.overwrite)
    with TemporaryDirectory(
        prefix="wps-export-", dir=path.parent
    ) as directory:
        staged = Path(directory) / "export.pdf"
        export_pdf_to(session, document, staged)
        publish_file(staged, path, overwrite=params.overwrite)
    return verify_output(path)


def preview_document(params: PreviewInput) -> bytes:
    """Return a PNG of a WPS-rendered page; ET pages follow print layout."""
    session = get_session(params.component)
    document = session.get_document(params.document)
    with TemporaryDirectory(prefix="wps-preview-") as directory:
        path = Path(directory) / "preview.pdf"
        export_pdf_to(session, document, path)
        return render_pdf_page(path, params.page_index, params.width)


def detect_components(params: InputModel) -> JsonObject:
    """Inspect WPS registration without starting applications."""
    result = []
    for component, spec in SPECS.items():
        registered = False
        if sys.platform == "win32":
            import winreg

            try:
                with winreg.OpenKey(
                    winreg.HKEY_CLASSES_ROOT, spec.prog_id + "\\CLSID"
                ):
                    registered = True
            except FileNotFoundError:
                pass
        result.append(
            {
                "component": component,
                "prog_id": spec.prog_id,
                "registered": registered,
                "formats": list(spec.formats),
                "native_execution_verified": False,
            }
        )
    return {
        "platform": sys.platform,
        "components": result,
        "note": (
            "Registration does not verify support for every feature "
            "in this WPS build."
        ),
    }
