"""Check numeric dates at the Writer and Presentation output boundaries."""

from datetime import UTC, date, datetime
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from wps_mcp.common.dates import date_text
from wps_mcp.presentation import annotations
from wps_mcp.writer import review


@pytest.mark.parametrize(
    ("serial", "expected"),
    [
        (0, "1899-12-30 00:00:00"),
        (0.0, "1899-12-30 00:00:00"),
        (0.25, "1899-12-30 06:00:00"),
        (-0.25, "1899-12-30 06:00:00"),
        (1.25, "1899-12-31 06:00:00"),
        (-1.25, "1899-12-29 06:00:00"),
        (-2.75, "1899-12-28 18:00:00"),
        (-1.0, "1899-12-29 00:00:00"),
        (2, "1900-01-01 00:00:00"),
        (60.5, "1900-02-28 12:00:00"),
        (61.5, "1900-03-01 12:00:00"),
        (46319.5, "2026-10-24 12:00:00"),
    ],
)
def test_ole_date_uses_absolute_fractional_time(
    serial: int | float, expected: str
) -> None:
    assert date_text(serial) == expected


@pytest.mark.parametrize(
    "value",
    [
        date(2026, 10, 24),
        datetime(2026, 10, 24, 12, 34, 56, 123456),
        datetime(2026, 10, 24, 12, 34, 56, tzinfo=UTC),
        "2026-10-24 12:00:00",
        "unrecognized date",
        None,
        True,
        SimpleNamespace(value="unknown"),
    ],
)
def test_existing_date_and_unknown_values_keep_string_contract(
    value: object,
) -> None:
    assert date_text(value) == str(value)


@pytest.mark.parametrize("value", [float("nan"), float("inf"), -float("inf")])
def test_ole_date_rejects_nonfinite_values(value: float) -> None:
    with pytest.raises(ValueError, match="finite"):
        date_text(value)


@pytest.mark.parametrize("value", [-1e20, 1e20, 10**400])
def test_ole_date_rejects_values_outside_datetime_range(
    value: int | float,
) -> None:
    with pytest.raises(ValueError, match="supported date range"):
        date_text(value)


def test_writer_revision_uses_numeric_date_conversion(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    revision = SimpleNamespace(
        Author="Reviewer",
        Date=46319.5,
        Type=1,
        Range=SimpleNamespace(Start=0, End=2, Text="text"),
    )
    session = Mock()
    session.get_document.return_value.Revisions = SimpleNamespace(
        Count=1, Item=lambda index: revision
    )
    monkeypatch.setattr(review, "get_session", lambda component: session)
    result = review.list_revisions(review.ReviewListInput())
    assert result["revisions"][0]["date"] == "2026-10-24 12:00:00"


def test_writer_comment_uses_numeric_date_conversion() -> None:
    comment = SimpleNamespace(
        Index=1,
        Range=SimpleNamespace(Text="comment"),
        Author="Reviewer",
        Initial="R",
        Date=-1.25,
        Scope=SimpleNamespace(Start=0, End=2, Text="text"),
    )
    assert review.comment_details(comment)["date"] == "1899-12-29 06:00:00"


@pytest.mark.parametrize(
    ("value", "expected"),
    [(46319.5, "2026-10-24 12:00:00"), (float("nan"), None)],
)
def test_presentation_comment_converts_or_reports_unavailable_date(
    monkeypatch: pytest.MonkeyPatch, value: float, expected: str | None
) -> None:
    comment = SimpleNamespace(
        DateTime=value,
        Author="Reviewer",
        AuthorInitials="R",
        Text="comment",
        Left=10.0,
        Top=20.0,
    )
    collection = Mock(Count=1, return_value=comment)
    session = Mock()
    session.get_document.return_value.Slides.return_value.Comments = collection
    monkeypatch.setattr(annotations, "get_session", lambda component: session)
    result = annotations.list_comments(
        annotations.ListCommentsInput(slide_index=1)
    )
    assert result["comments"][0]["date_time"] == expected
