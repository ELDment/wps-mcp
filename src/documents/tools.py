"""Register component-neutral document services in the tool catalog."""

from src.documents import service
from src.protocol.definitions import (
    READ_ANNOTATIONS,
    WRITE_ANNOTATIONS,
    tool,
)

tool(
    "wps_components",
    operation="connect",
    component="auto",
    annotations=WRITE_ANNOTATIONS,
)(service.connect)

tool(
    "wps_documents",
    operation="list",
    component="auto",
    annotations=READ_ANNOTATIONS,
)(service.list_documents)

tool(
    "wps_documents",
    operation="create",
    component="auto",
    annotations=WRITE_ANNOTATIONS,
)(service.create_document)

tool(
    "wps_documents",
    operation="open",
    component="auto",
    annotations=WRITE_ANNOTATIONS,
)(service.open_document)

tool(
    "wps_documents",
    operation="save",
    component="auto",
    annotations=WRITE_ANNOTATIONS,
)(service.save_document)

tool(
    "wps_documents",
    operation="close",
    component="auto",
    annotations=WRITE_ANNOTATIONS,
)(service.close_document)

tool(
    "wps_documents",
    operation="info",
    component="auto",
    annotations=READ_ANNOTATIONS,
)(service.get_document_info)

tool(
    "wps_documents",
    operation="export_pdf",
    component="auto",
    annotations=WRITE_ANNOTATIONS,
)(service.export_pdf)

tool(
    "wps_documents",
    operation="preview",
    component="auto",
    annotations=READ_ANNOTATIONS,
)(service.preview_document)

tool(
    "wps_components",
    operation="detect",
    component=None,
    annotations=READ_ANNOTATIONS,
)(service.detect_components)
