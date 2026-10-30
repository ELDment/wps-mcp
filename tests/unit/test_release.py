"""Verify dated release allocation independently of network and GitHub."""

from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from unittest.mock import Mock

import pytest

from scripts.ci import prepare_release
from scripts.ci.prepare_release import release_version


@pytest.mark.parametrize(
    ("releases", "expected"),
    [
        ([], "2026.10.26"),
        (["2026.10.26"], "2026.10.26.post1"),
        (
            ["2026.10.26.post2", "2026.10.26", "2026.10.26.post9"],
            "2026.10.26.post10",
        ),
        (
            [
                "2026.10.25.post99",
                "2026.10.26.dev12",
                "2026.10.26.post1.dev13",
                "2026.10.260",
                "2026.10.26.post3.invalid",
            ],
            "2026.10.26",
        ),
    ],
)
def test_stable_release_allocation(releases: list[str], expected: str) -> None:
    assert release_version(date(2026, 10, 26), releases) == expected


@pytest.mark.parametrize(
    ("releases", "expected"),
    [
        ([], "2026.1.2.dev42"),
        (["2026.1.2"], "2026.1.2.post1.dev42"),
        (["2026.1.2.post3"], "2026.1.2.post4.dev42"),
    ],
)
def test_feature_release_tracks_run_number(
    releases: list[str], expected: str
) -> None:
    assert (
        release_version(date(2026, 1, 2), releases, development_run=42)
        == expected
    )


@pytest.mark.parametrize("run", [0, -1])
def test_feature_release_rejects_invalid_run_number(run: int) -> None:
    with pytest.raises(ValueError, match="positive"):
        release_version(date(2026, 10, 26), [], development_run=run)


@pytest.mark.parametrize(
    ("ref", "expected"),
    [
        ("refs/heads/main", "2026.10.26.post1"),
        ("refs/heads/feat/example", "2026.10.26.post1.dev42"),
        ("refs/heads/fix/example", "2026.10.26.post1.dev42"),
    ],
)
def test_release_changes_metadata_without_rewriting_package_source(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, ref: str, expected: str
) -> None:
    monkeypatch.chdir(tmp_path)
    project = tmp_path / "pyproject.toml"
    lock = tmp_path / "uv.lock"
    summary = tmp_path / "summary.md"
    source = tmp_path / "src" / "wps_mcp" / "__init__.py"
    source.parent.mkdir(parents=True)
    source.write_bytes(b"from importlib.metadata import version\r\n")
    project.write_bytes(b'[project]\nname = "wps-mcp"\nversion = "1.2.0"\n')
    lock.write_bytes(b"version = 1\n")
    run = Mock()
    monkeypatch.setattr(prepare_release.subprocess, "run", run)
    monkeypatch.setattr(
        prepare_release,
        "urlopen",
        Mock(return_value=BytesIO(b'{"releases":{"2026.10.26":[]}}')),
    )
    monkeypatch.setattr(
        prepare_release,
        "datetime",
        Mock(now=Mock(return_value=datetime(2026, 10, 26))),
    )
    monkeypatch.setenv("GITHUB_REF", ref)
    monkeypatch.setenv("GITHUB_RUN_NUMBER", "42")
    monkeypatch.setenv("GITHUB_STEP_SUMMARY", str(summary))

    prepare_release.main()

    run.assert_called_once_with(
        ["uv", "version", expected, "--no-sync"], check=True
    )
    assert expected in summary.read_text(encoding="utf-8")
    assert source.read_bytes() == b"from importlib.metadata import version\r\n"
    for path in (project, lock):
        assert b"\n" not in path.read_bytes().replace(b"\r\n", b"")
