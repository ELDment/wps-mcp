"""Strict inputs for component and document lifecycle operations."""

from typing import Literal, Self

from pydantic import Field, model_validator

from wps_mcp.models import Component, DocumentInput, InputModel
from wps_mcp.native.capabilities import validate_visibility


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
        validate_visibility(self.component, self.visible)
        return self


class ComponentDocumentInput(DocumentInput):
    """Identify an open document within a WPS component."""

    component: Component


class CreateInput(ConnectInput):
    """Create a document, optionally from a local template."""

    template_path: str | None = Field(default=None, min_length=1)
    activate: bool = True
    slide_width: float | None = Field(default=None, gt=0)
    slide_height: float | None = Field(default=None, gt=0)
    preset: Literal["16:9", "4:3"] | None = None

    @model_validator(mode="after")
    def validate_dimensions(self) -> Self:
        """Accept slide dimensions only for blank presentations."""
        custom = self.slide_width is not None or self.slide_height is not None
        if self.component != "wpp" and (custom or self.preset is not None):
            raise ValueError("Slide dimensions require component='wpp'")
        if self.template_path is not None and (custom or self.preset):
            raise ValueError("Template dimensions cannot be overridden")
        if self.preset is not None and custom:
            raise ValueError("Choose preset or custom slide dimensions")
        if (self.slide_width is None) != (self.slide_height is None):
            raise ValueError(
                "slide_width and slide_height must be supplied together"
            )
        return self


class OpenInput(ConnectInput):
    """Open an existing local document with an explicit read-only mode."""

    file_path: str = Field(min_length=1)
    read_only: bool = False
    with_window: bool = True
    activate: bool = True

    @model_validator(mode="after")
    def validate_window(self) -> Self:
        """Keep document windows distinct from application visibility."""
        if self.component != "wpp" and not self.with_window:
            raise ValueError("with_window=false requires component='wpp'")
        return self


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

    slide_range_start: int | None = Field(default=None, ge=1)
    slide_range_end: int | None = Field(default=None, ge=1)

    @model_validator(mode="after")
    def validate_range(self) -> Self:
        """Require a paired ascending range for presentation exports."""
        if (self.slide_range_start is None) != (self.slide_range_end is None):
            raise ValueError(
                "Provide both slide_range_start and slide_range_end"
            )
        if self.slide_range_start is not None:
            if self.component != "wpp":
                raise ValueError("Slide ranges require component='wpp'")
            assert self.slide_range_end is not None
            if self.slide_range_start > self.slide_range_end:
                raise ValueError(
                    "slide_range_start must not exceed slide_range_end"
                )
        return self


class PreviewInput(ComponentDocumentInput):
    """Select a one-based printed page and bounded image width."""

    page_index: int = Field(default=1, ge=1)
    width: int = Field(default=1280, ge=320, le=4096)
