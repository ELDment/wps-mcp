"""Export and preview documents with validated staged output."""

from contextlib import closing
from pathlib import Path

import pypdfium2 as pdfium

from wps_mcp.common.files import publish_file, verify_output
from wps_mcp.common.paths import destination
from wps_mcp.common.rendering import RENDER_LOCK, render_pdf_page
from wps_mcp.documents.models import ExportInput, PreviewInput
from wps_mcp.models import JsonObject, NativeObject
from wps_mcp.native.files import temporary_directory
from wps_mcp.native.runtime import get_session
from wps_mcp.native.session import LocalSession


def export_pdf_to(
    session: LocalSession,
    document: NativeObject,
    file_path: Path,
) -> JsonObject:
    """Export through native WPS and verify the PDF output signature."""
    if session.component == "writer":
        document.ExportAsFixedFormat(str(file_path), 17)
    elif session.component == "et":
        document.ExportAsFixedFormat(0, str(file_path))
    else:
        document.ExportAsFixedFormat(
            str(file_path),
            2,
            PrintHiddenSlides=-1,
            PrintRange=None,
            RangeType=1,
        )
    result = verify_output(file_path)
    with file_path.open("rb") as stream:
        if stream.read(5) != b"%PDF-":
            raise RuntimeError("WPS did not produce a PDF file")
    return result


def export_pdf(params: ExportInput) -> JsonObject:
    """Publish a verified PDF without modifying document print ranges."""
    path = destination(params.file_path, {".pdf"}, params.overwrite)
    session = get_session(params.component)
    document = session.get_document(params.document)
    total = int(document.Slides.Count) if params.component == "wpp" else None
    start = params.slide_range_start or 1
    end = params.slide_range_end or total
    if total is not None and (end is None or not 1 <= start <= end <= total):
        raise ValueError(
            f"Slide range must satisfy 1 <= start <= end <= {total}"
        )
    with temporary_directory(
        prefix="wps-export-", dir=path.parent
    ) as directory:
        rendered = Path(directory) / "document.pdf"
        export_pdf_to(session, document, rendered)
        if total is not None:
            assert end is not None
            with RENDER_LOCK:
                with closing(pdfium.PdfDocument(str(rendered))) as source:
                    if len(source) != total:
                        raise RuntimeError(
                            f"WPS exported {len(source)} pages; "
                            f"expected {total}"
                        )
                    if start != 1 or end != total:
                        selected = Path(directory) / "selected.pdf"
                        with closing(pdfium.PdfDocument.new()) as output:
                            output.import_pages(
                                source, pages=list(range(start - 1, end))
                            )
                            output.save(str(selected))
                        rendered = selected
                with closing(pdfium.PdfDocument(str(rendered))) as verified:
                    if len(verified) != end - start + 1:
                        raise RuntimeError(
                            "The exported PDF has an unexpected page count"
                        )
        publish_file(rendered, path, overwrite=params.overwrite)
    result = verify_output(path)
    if total is not None:
        result.update(
            slide_range_start=start, slide_range_end=end, total_slides=total
        )
    return result


def preview_document(params: PreviewInput) -> bytes:
    """Return a PNG of a WPS-rendered page; ET follows print layout."""
    session = get_session(params.component)
    document = session.get_document(params.document)
    with temporary_directory(prefix="wps-preview-") as directory:
        path = Path(directory) / "preview.pdf"
        export_pdf_to(session, document, path)
        return render_pdf_page(path, params.page_index, params.width)
