"""Perform document lifecycle operations independently of MCP registration."""

import logging

from wps_mcp.common.files import protect_existing_file, verify_saved_file
from wps_mcp.common.paths import destination, local_path
from wps_mcp.documents.formats import check_presentation_format
from wps_mcp.documents.models import (
    CloseInput,
    ComponentDocumentInput,
    ComponentInput,
    ConnectInput,
    CreateInput,
    OpenInput,
    SaveInput,
)
from wps_mcp.documents.presentation import (
    SLIDE_SIZE_PRESETS,
    presentation_details,
    presentation_summary,
)
from wps_mcp.errors import native_errors
from wps_mcp.models import InputModel, JsonObject, NativeObject
from wps_mcp.native.discovery import detect_components as inspect_components
from wps_mcp.native.runtime import get_session
from wps_mcp.native.selection import document_key
from wps_mcp.native.session import LocalSession
from wps_mcp.native.state import temporary_attributes

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
        "visible": bool(application.Visible),
        "document_count": int(
            getattr(application, session.spec.collection).Count
        ),
    }


def list_documents(params: ComponentInput) -> JsonObject:
    """List the full_name identifiers of documents open in one component."""
    documents = []
    for index, item in enumerate(
        get_session(params.component).get_documents(), 1
    ):
        info = {"index": index, **describe_document(item)}
        if params.component == "wpp":
            info.update(presentation_summary(item))
        documents.append(info)
    return {
        "component": params.component,
        "documents": documents,
        "count": len(documents),
    }


def _created_document(
    session: LocalSession,
    document: NativeObject,
    *,
    activate: bool,
    with_window: bool = True,
) -> JsonObject:
    """Describe a new target and optionally activate its editor window."""
    result = describe_document(document)
    result["activated"] = activate
    if activate:
        session.target_name = str(document.FullName)
    if session.component == "wpp":
        result.update(presentation_summary(document))
        result["index"] = next(
            (
                index
                for index, item in enumerate(session.get_documents(), 1)
                if document_key(item.FullName)
                == document_key(document.FullName)
            ),
            None,
        )
        if activate and with_window:
            try:
                document.Windows(1).Activate()
            except native_errors() + (AttributeError,) as error:
                result["warnings"] = [f"Window activation failed: {error}"]
    return result


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
        dimensions = (
            SLIDE_SIZE_PRESETS.get(params.preset) if params.preset else None
        )
        if dimensions is None and params.slide_width is not None:
            assert params.slide_height is not None
            dimensions = params.slide_width, params.slide_height
        if dimensions is not None:
            document.PageSetup.SlideWidth, document.PageSetup.SlideHeight = (
                dimensions
            )
    result = _created_document(session, document, activate=params.activate)
    if params.component == "wpp":
        try:
            result["template_name"] = document.TemplateName
        except native_errors() + (AttributeError,):
            result["template_name"] = None
    return result


def open_document(params: OpenInput) -> JsonObject:
    """Open a local document; reject a duplicate already open in WPS."""
    path = local_path(params.file_path, exists=True)
    session = get_session(params.component)
    if path.suffix.lower() not in session.spec.formats:
        raise ValueError(
            f"File type {path.suffix} is unsupported by {params.component}"
        )
    visible = (
        None
        if params.component == "wpp" and not params.with_window
        else params.visible
    )
    application = session.connect(visible=visible)
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
            str(path), ReadOnly=params.read_only, WithWindow=params.with_window
        )
    return _created_document(
        session,
        document,
        activate=params.activate,
        with_window=params.with_window,
    )


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

    def verify_closed() -> None:
        if any(
            document_key(opened.FullName)
            == document_key(identity["full_name"])
            for opened in session.get_documents()
        ):
            raise RuntimeError(
                "WPS did not close the document; dismiss its dialog "
                "and inspect the document before repeating the operation"
            )

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
            verify_closed()
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
        verify_closed()
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
        info.update(presentation_details(document))
    return info


def detect_components(params: InputModel) -> JsonObject:
    """Inspect platform prerequisites without starting applications."""
    return inspect_components()
