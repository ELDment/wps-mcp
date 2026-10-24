"""Prepare dated package versions for GitHub Actions releases."""

import json
import os
import re
import subprocess
import tomllib
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import urlopen


def main() -> None:
    """Update release metadata from the repository root before building."""
    project_path = Path("pyproject.toml")
    project = tomllib.loads(project_path.read_text(encoding="utf-8"))
    package_name = project["project"]["name"]
    release_date = datetime.now(timezone(timedelta(hours=8)))
    base_version = (
        f"{release_date.year}.{release_date.month}.{release_date.day}"
    )

    try:
        with urlopen(
            f"https://pypi.org/pypi/{package_name}/json", timeout=30
        ) as response:
            releases = json.load(response)["releases"]
    except HTTPError as error:
        if error.code != 404:
            raise
        releases = {}

    release_pattern = re.compile(rf"{re.escape(base_version)}(?:\.post(\d+))?")
    revisions = [
        int(match[1] or 0)
        for version in releases
        if (match := release_pattern.fullmatch(version))
    ]
    release_version = base_version
    if revisions:
        release_version += f".post{max(revisions) + 1}"

    if os.environ["GITHUB_REF"].startswith("refs/heads/feat/"):
        release_version += f".dev{os.environ['GITHUB_RUN_NUMBER']}"

    source_path = Path("src/__init__.py")
    source, replacements = re.subn(
        r'^__version__ = "[^"]+"$',
        f'__version__ = "{release_version}"',
        source_path.read_text(encoding="utf-8"),
        flags=re.MULTILINE,
    )
    if replacements != 1:
        raise SystemExit("Expected exactly one src.__version__ assignment")

    # Re-locking may need metadata beyond the current Python's cache.
    subprocess.run(
        ["uv", "version", release_version, "--no-sync"],
        check=True,
    )
    source_path.write_text(source, encoding="utf-8", newline="\r\n")
    for path in (project_path, Path("uv.lock")):
        content = path.read_text(encoding="utf-8")
        path.write_text(content, encoding="utf-8", newline="\r\n")

    print(f"Prepared {package_name} {release_version}")
    with open(
        os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8"
    ) as summary:
        summary.write(
            f"Prepared **{package_name} {release_version}** for PyPI.\n"
        )


if __name__ == "__main__":
    main()
