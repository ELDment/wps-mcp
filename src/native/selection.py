"""Resolve explicit document identities without foreground heuristics."""

from collections.abc import Sequence
from os import path

from src.errors import DocumentNotFoundError
from src.models import NativeObject


def document_key(value: str | object) -> str:
    """Normalize document identities using the host filesystem rules."""
    return path.normcase(path.normpath(str(value)))


def select_document(
    documents: Sequence[NativeObject], wanted: str | int | None = None
) -> NativeObject:
    """Select by full path, unique name or one-based collection index."""
    if not documents:
        raise DocumentNotFoundError(
            "No document is open; create or open one first"
        )
    if wanted is None:
        if len(documents) == 1:
            return documents[0]
        raise DocumentNotFoundError(
            "Several documents are open; specify a full path or unique name"
        )
    if isinstance(wanted, int) and not isinstance(wanted, bool):
        if 1 <= wanted <= len(documents):
            return documents[wanted - 1]
        raise DocumentNotFoundError("Document index is out of range")
    if not isinstance(wanted, str) or not wanted.strip():
        raise DocumentNotFoundError(
            "document must be a non-empty name or full path"
        )
    key = document_key(wanted)
    exact = [doc for doc in documents if document_key(doc.FullName) == key]
    if len(exact) == 1:
        return exact[0]
    matches = (
        []
        if path.dirname(wanted)
        else [
            doc
            for doc in documents
            if key
            in {
                document_key(doc.Name),
                document_key(path.splitext(doc.Name)[0]),
            }
        ]
    )
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1 or len(exact) > 1:
        raise DocumentNotFoundError(
            f"Ambiguous document name: {wanted}; specify its full path"
        )
    raise DocumentNotFoundError(f"Requested document is not open: {wanted}")
