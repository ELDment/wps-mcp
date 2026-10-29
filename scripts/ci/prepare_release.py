"""Prepare dated package versions for GitHub Actions releases."""

import json
import os
import re
import subprocess
import tomllib
from collections.abc import Iterable
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import urlopen


def release_version(
    release_date: date,
    releases: Iterable[str],
    *,
    development_run: int | None = None,
) -> str:
    """Allocate the next dated release, optionally as a development build."""
    if development_run is not None and development_run < 1:
        raise ValueError("Development run must be positive")
    base_version = (
        f"{release_date.year}.{release_date.month}.{release_date.day}"
    )
    release_pattern = re.compile(rf"{re.escape(base_version)}(?:\.post(\d+))?")
    revisions = [
        int(match[1] or 0)
        for version in releases
        if (match := release_pattern.fullmatch(version))
    ]
    version = base_version
    if revisions:
        version += f".post{max(revisions) + 1}"
    if development_run is not None:
        version += f".dev{development_run}"
    return version


def main() -> None:
    """Update the sole project version and its lockfile before building."""
    project_path = Path("pyproject.toml")
    project = tomllib.loads(project_path.read_text(encoding="utf-8"))
    package_name = project["project"]["name"]
    try:
        with urlopen(
            f"https://pypi.org/pypi/{package_name}/json", timeout=30
        ) as response:
            releases = json.load(response)["releases"]
    except HTTPError as error:
        if error.code != 404:
            raise
        releases = {}
    version = release_version(
        datetime.now(timezone(timedelta(hours=8))).date(),
        releases,
        development_run=(
            int(os.environ["GITHUB_RUN_NUMBER"])
            if os.environ["GITHUB_REF"].startswith(
                ("refs/heads/feat/", "refs/heads/fix/")
            )
            else None
        ),
    )

    # Re-locking may need metadata beyond the current Python's cache.
    subprocess.run(
        ["uv", "version", version, "--no-sync"],
        check=True,
    )
    for path in (project_path, Path("uv.lock")):
        content = path.read_text(encoding="utf-8")
        path.write_text(content, encoding="utf-8", newline="\r\n")

    print(f"Prepared {package_name} {version}")
    with open(
        os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8"
    ) as summary:
        summary.write(f"Prepared **{package_name} {version}** for PyPI.\n")


if __name__ == "__main__":
    main()
