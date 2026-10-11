"""Strict inputs for component and document lifecycle operations."""

from typing import Self

from pydantic import Field, model_validator

from src.models import Component, DocumentInput, InputModel


class ComponentInput(InputModel):
    """Select the Writer, ET or WPP component."""

    component: Component


class ConnectInput(ComponentInput):
    """Choose a component and its requested application visibility."""

    visible: bool = Field(
        default=True,
        description=(
            "Show the application. Only Writer and ET support false; "
            "do not attempt to hide WPP."
        ),
    )

    @model_validator(mode="after")
    def validate_visibility(self) -> Self:
        """Reject unsupported WPP visibility before native execution."""
        if self.component == "wpp" and self.visible is False:
            raise ValueError("WPP cannot be hidden; use visible=true")
        return self


class ComponentDocumentInput(DocumentInput):
    """Identify an open document within a WPS component."""

    component: Component


class CreateInput(ConnectInput):
    """Create a document, optionally from a local template."""

    template_path: str | None = Field(default=None, min_length=1)


class OpenInput(ConnectInput):
    """Open an existing local document with an explicit read-only mode."""

    file_path: str = Field(min_length=1)
    read_only: bool = False


class SaveInput(ComponentDocumentInput):
    """Save a document with an explicit destination overwrite policy."""

    file_path: str | None = Field(default=None, min_length=1)
    overwrite: bool = False


class CloseInput(ComponentDocumentInput):
    """Close a document with an explicit save or discard decision."""

    save_changes: bool = False
    discard_changes: bool = False

    @model_validator(mode="after")
    def validate_close_mode(self) -> Self:
        """Reject mutually exclusive save and discard requests."""
        if self.save_changes and self.discard_changes:
            raise ValueError(
                "save_changes and discard_changes are mutually exclusive"
            )
        return self


class ExportInput(ComponentDocumentInput):
    """Select a PDF destination and its overwrite policy."""

    file_path: str = Field(min_length=1)
    overwrite: bool = False


class PreviewInput(ComponentDocumentInput):
    """Select a one-based printed page and bounded image width."""

    page_index: int = Field(default=1, ge=1)
    width: int = Field(default=1280, ge=320, le=4096)
