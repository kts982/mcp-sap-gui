#!/usr/bin/env python3
"""Extract one release's section from CHANGELOG.md for the GitHub release body."""

from __future__ import annotations

import argparse
import re
import sys
import urllib.parse
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CHANGELOG = ROOT / "CHANGELOG.md"
DEFAULT_REPO = "kts982/mcp-sap-gui"

LINK_RE = re.compile(r"(?<!!)(\[[^\]]+\]\()([^)]+)(\))")


def release_section(text: str, version: str) -> str | None:
    """Text under ``## <version> — <date>`` up to the next ``## `` heading."""
    heading = re.compile(rf"^## {re.escape(version)}(\s|$)")
    lines = text.splitlines()
    for start, line in enumerate(lines):
        if heading.match(line):
            break
    else:
        return None

    body: list[str] = []
    for line in lines[start + 1 :]:
        if line.startswith("## "):
            break
        body.append(line)
    return "\n".join(body).strip()


def absolute_links(text: str, repo: str, ref: str) -> str:
    """Point repo-relative links at the tagged tree.

    A release page keeps the href as written, so ``docs/X.md`` would resolve
    under ``/releases/tag/`` and 404.
    """
    base = f"https://github.com/{repo}/blob/{ref}/"

    def rewrite(match: re.Match[str]) -> str:
        target = match.group(2).strip()
        if target.startswith(("#", "/")) or urllib.parse.urlparse(target).scheme:
            return match.group(0)
        return f"{match.group(1)}{base}{target}{match.group(3)}"

    return LINK_RE.sub(rewrite, text)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("version", help="Release version without the v, e.g. 0.5.2.")
    parser.add_argument("--output", type=Path, help="Write here instead of to stdout.")
    parser.add_argument("--repo", default=DEFAULT_REPO, help="owner/name for absolute links.")
    parser.add_argument("--ref", help="Tag the links point at (default: v<version>).")
    args = parser.parse_args()

    section = release_section(CHANGELOG.read_text(encoding="utf-8"), args.version)
    if section is None:
        raise SystemExit(
            f"CHANGELOG.md has no '## {args.version}' section; "
            f"add one before tagging v{args.version}."
        )
    if not section:
        raise SystemExit(f"CHANGELOG.md section '## {args.version}' is empty.")

    # No trailing newline: the release action joins this body and GitHub's
    # generated notes with one blank line.
    notes = absolute_links(section, args.repo, args.ref or f"v{args.version}")
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(notes, encoding="utf-8", newline="\n")
    else:
        sys.stdout.buffer.write(notes.encode("utf-8"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
