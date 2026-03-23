#!/usr/bin/env python3
"""Lightweight Markdown link checker for README and docs."""

from __future__ import annotations

import argparse
import re
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MARKDOWN_FILES = [ROOT / "README.md", *sorted((ROOT / "docs").rglob("*.md"))]

LINK_RE = re.compile(r"(?<!!)\[[^\]]+\]\(([^)]+)\)")
HEADER_RE = re.compile(r"^\s{0,3}#{1,6}\s+(.+?)\s*$")
SKIP_EXTERNAL_PREFIXES = ("mailto:",)


def normalize_anchor(text: str) -> str:
    text = re.sub(r"`([^`]*)`", r"\1", text)
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
    text = text.strip().lower()
    text = re.sub(r"[^a-z0-9 _-]", "", text)
    text = text.replace(" ", "-")
    text = re.sub(r"-+", "-", text)
    return text.strip("-")


def anchors_for_markdown(path: Path) -> set[str]:
    anchors: set[str] = set()
    for line in path.read_text(encoding="utf-8").splitlines():
        match = HEADER_RE.match(line)
        if match:
            anchors.add(normalize_anchor(match.group(1)))
    return anchors


def clean_link_target(raw: str) -> str:
    target = raw.strip()
    if target.startswith("<") and target.endswith(">"):
        target = target[1:-1].strip()
    if " " in target and not target.startswith("http"):
        target = target.split(" ", 1)[0]
    return target


def iter_links(path: Path) -> list[str]:
    text = path.read_text(encoding="utf-8")
    return [clean_link_target(match.group(1)) for match in LINK_RE.finditer(text)]


def check_internal_links() -> list[str]:
    errors: list[str] = []
    anchor_cache: dict[Path, set[str]] = {}

    for md_file in MARKDOWN_FILES:
        for link in iter_links(md_file):
            if not link or link.startswith("#"):
                continue
            if urllib.parse.urlparse(link).scheme:
                continue

            target, _, anchor = link.partition("#")
            resolved = (md_file.parent / urllib.parse.unquote(target)).resolve()

            if not resolved.exists():
                errors.append(f"{md_file.relative_to(ROOT)} -> missing target: {link}")
                continue

            if anchor and resolved.suffix.lower() == ".md":
                if resolved not in anchor_cache:
                    anchor_cache[resolved] = anchors_for_markdown(resolved)
                if normalize_anchor(anchor) not in anchor_cache[resolved]:
                    errors.append(f"{md_file.relative_to(ROOT)} -> missing anchor: {link}")

    return errors


def check_external_links() -> list[str]:
    errors: list[str] = []
    seen: set[str] = set()

    for md_file in MARKDOWN_FILES:
        for link in iter_links(md_file):
            if not link:
                continue
            parsed = urllib.parse.urlparse(link)
            if parsed.scheme not in {"http", "https"}:
                continue
            if link in seen or link.startswith(SKIP_EXTERNAL_PREFIXES):
                continue
            seen.add(link)

            request = urllib.request.Request(
                link,
                headers={"User-Agent": "mcp-sap-gui-doc-check/1.0"},
                method="HEAD",
            )
            try:
                with urllib.request.urlopen(request, timeout=20) as response:
                    if response.status >= 400:
                        errors.append(f"external link failed: {link} ({response.status})")
            except urllib.error.HTTPError as exc:
                if exc.code in {401, 403, 405, 429}:
                    continue
                errors.append(f"external link failed: {link} ({exc.code})")
            except Exception as exc:  # pragma: no cover - network-specific
                errors.append(f"external link failed: {link} ({exc})")

    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description="Check Markdown links in README and docs.")
    parser.add_argument(
        "--external",
        action="store_true",
        help="Also check external HTTP/HTTPS links.",
    )
    args = parser.parse_args()

    errors = check_internal_links()
    if args.external:
        errors.extend(check_external_links())

    if errors:
        for error in errors:
            print(error)
        return 1

    mode = "internal + external" if args.external else "internal"
    print(f"Documentation link check passed ({mode}).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
