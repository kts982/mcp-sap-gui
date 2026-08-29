"""Builders for the ``sap_preview`` approval/preview surface.

Pure functions: no COM, no MCP context, no I/O.  ``build_preview_text``
produces the authored summary every host receives; ``build_preview_card``
produces the Prefab UI card for hosts that negotiated the MCP Apps UI
extension.

prefab-ui is imported lazily inside ``build_preview_card`` so this module
(and the tool that uses it) import and work without the optional ``apps``
extra installed.
"""

from __future__ import annotations

import importlib.util
import re
from typing import Any, Mapping

__all__ = [
    "SCREENSHOT_INCLUDED",
    "SCREENSHOT_OMITTED",
    "SCREENSHOT_UNAVAILABLE",
    "build_preview_card",
    "build_preview_text",
    "prefab_available",
]

_CLOSING_LINE = "Preview only — nothing has been written or saved."
_FOOTER_LINE = "Saving still requires confirmation."
_TITLE = "SAP SCREEN PREVIEW"

SCREENSHOT_INCLUDED = "included"
SCREENSHOT_OMITTED = "omitted"
SCREENSHOT_UNAVAILABLE = "unavailable"

_SCREENSHOT_LINES = {
    SCREENSHOT_INCLUDED: "Screenshot: included with this preview.",
    SCREENSHOT_OMITTED: "Screenshot: omitted (include_screenshot=false).",
    SCREENSHOT_UNAVAILABLE: "Screenshot: unavailable (capture failed).",
}

# Mirrors controller._is_sensitive_field_id's token set. Keys are supplied by
# the agent (a label or an element ID), so match the same tokens on the key.
_SENSITIVE_TOKENS = ("PWD", "BCODE", "PASSWORD")
_MASKED = "***"

# Screen text is attacker-influenced (an SAP field can contain newlines that
# would forge extra summary lines) and unbounded in length.
_MAX_VALUE_CHARS = 200
_MAX_FIELD_ROWS = 50
_WHITESPACE_RE = re.compile(r"\s+")


def prefab_available() -> bool:
    """Report whether the optional ``apps`` extra (prefab-ui) is importable."""
    return importlib.util.find_spec("prefab_ui") is not None


# ---------------------------------------------------------------------------
# Shared extraction (text and card must not drift apart)
# ---------------------------------------------------------------------------

def _text(value: Any) -> str:
    """Flatten a value to a single-line string (newlines cannot forge lines)."""
    if value is None:
        return ""
    return _WHITESPACE_RE.sub(" ", str(value)).strip()


def _clip(value: str) -> str:
    if len(value) <= _MAX_VALUE_CHARS:
        return value
    return value[: _MAX_VALUE_CHARS - 3] + "..."


def _get(mapping: Mapping[str, Any] | None, key: str) -> str:
    if not isinstance(mapping, Mapping):
        return ""
    return _text(mapping.get(key))


def _screen_error(screen: Mapping[str, Any] | None) -> str:
    """Return the screen-read error, if the controller reported one."""
    return _clip(_get(screen, "error"))


def _heading(screen: Mapping[str, Any] | None) -> str:
    transaction = _clip(_get(screen, "transaction"))
    title = _clip(_get(screen, "title"))
    if transaction and title:
        return f"{transaction} — {title}"
    return transaction or title or "Current SAP screen"


def _context_line(
    screen: Mapping[str, Any] | None,
    session: Mapping[str, Any] | None,
) -> str:
    parts: list[str] = []
    system = _get(session, "system_name")
    client = _get(session, "client")
    if system:
        parts.append(f"system {system}/{client}" if client else f"system {system}")
    user = _get(session, "user")
    if user:
        parts.append(f"user {user}")
    program = _get(screen, "program")
    if program:
        parts.append(f"program {program}")
    window = _get(screen, "active_window")
    if window:
        parts.append(f"window {window}" + (" (popup)" if window != "wnd[0]" else ""))
    return _clip(" | ".join(parts))


def _status_line(screen: Mapping[str, Any] | None) -> str:
    message = _clip(_get(screen, "message"))
    if not message:
        return ""
    message_type = _get(screen, "message_type")
    if message_type:
        return f"Status bar: [{message_type}] {message}"
    return f"Status bar: {message}"


def _is_sensitive(name: str) -> bool:
    upper = name.upper()
    return any(token in upper for token in _SENSITIVE_TOKENS)


def _normalize_fields(
    pending_fields: Mapping[str, Any] | None,
) -> tuple[list[tuple[str, str]], int]:
    """Return (rows, overflow_count), sanitized, masked, clipped and capped."""
    if not pending_fields:
        return [], 0

    rows: list[tuple[str, str]] = []
    for name, value in pending_fields.items():
        label = _clip(_text(name)) or _MASKED
        if _is_sensitive(str(name)):
            rows.append((label, _MASKED))
        else:
            rows.append((label, _clip(_text(value))))

    overflow = max(0, len(rows) - _MAX_FIELD_ROWS)
    return rows[:_MAX_FIELD_ROWS], overflow


def _badge(rows: list[tuple[str, str]]) -> tuple[str, str]:
    """Return (label, prefab variant) for the read/write context badge."""
    if rows:
        return "WRITE", "destructive"
    return "READ", "secondary"


def _overflow_line(overflow: int) -> str:
    return f"... and {overflow} more field{'s' if overflow != 1 else ''}"


# ---------------------------------------------------------------------------
# Text (every host)
# ---------------------------------------------------------------------------

def build_preview_text(
    *,
    screen: Mapping[str, Any] | None = None,
    session: Mapping[str, Any] | None = None,
    note: str = "",
    pending_fields: Mapping[str, Any] | None = None,
    screenshot: str = SCREENSHOT_OMITTED,
) -> str:
    """Render the authored summary shown to every host and to the model.

    Must stand alone: hosts without the MCP Apps UI extension see only this,
    so it mirrors every element of the card except the screenshot itself.
    """
    rows, overflow = _normalize_fields(pending_fields)
    lines: list[str] = []

    error = _screen_error(screen)
    if error:
        lines.append(f"WARNING: screen could not be read: {error}")
    lines.extend([_TITLE, _heading(screen)])

    context = _context_line(screen, session)
    if context:
        lines.append(context)
    status = _status_line(screen)
    if status:
        lines.append(status)
    note_text = _clip(_text(note))
    if note_text:
        lines.append(f"Note: {note_text}")

    lines.append("")
    if rows:
        lines.append(f"Values about to be written ({len(rows) + overflow}):")
        width = max(len(name) for name, _ in rows)
        lines.extend(f"  {name.ljust(width)} : {value}" for name, value in rows)
        if overflow:
            lines.append(f"  {_overflow_line(overflow)}")
    else:
        lines.append("No pending field values were supplied.")

    lines.append("")
    lines.append(_SCREENSHOT_LINES.get(screenshot, _SCREENSHOT_LINES[SCREENSHOT_OMITTED]))
    lines.append(_FOOTER_LINE)
    lines.append(_CLOSING_LINE)
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Card (MCP Apps UI hosts only)
# ---------------------------------------------------------------------------

def build_preview_card(
    *,
    screen: Mapping[str, Any] | None = None,
    session: Mapping[str, Any] | None = None,
    note: str = "",
    pending_fields: Mapping[str, Any] | None = None,
    image_data_uri: str | None = None,
):
    """Build the Prefab card for the current SAP screen.

    Returns a ``PrefabApp`` suitable for ``ToolResult(structured_content=...)``.
    Raises ``ImportError`` when the optional ``apps`` extra is missing — call
    ``prefab_available()`` first.

    SAP-derived strings are routed only through Text/TableCell/Heading/Badge;
    Markdown/Code/Svg/Embed are raw-HTML-capable by design and must never
    carry screen content.
    """
    from prefab_ui.app import PrefabApp
    from prefab_ui.components import (
        Badge,
        Card,
        CardContent,
        CardHeader,
        CardTitle,
        Column,
        Heading,
        Image,
        Row,
        Separator,
        Table,
        TableBody,
        TableCell,
        TableHead,
        TableHeader,
        TableRow,
        Text,
    )

    rows, overflow = _normalize_fields(pending_fields)
    badge_label, badge_variant = _badge(rows)
    heading = _heading(screen)
    error = _screen_error(screen)
    context = _context_line(screen, session)
    status = _status_line(screen)
    note_text = _clip(_text(note))
    muted = "text-xs text-muted-foreground"

    with PrefabApp() as app:
        with Column(gap=4, css_class="p-6"):
            with Row(gap=3, align="center"):
                Heading(heading, level=2)
                Badge(badge_label, variant=badge_variant)
                if error:
                    Badge("SCREEN READ FAILED", variant="warning")
            if error:
                Text(f"WARNING: screen could not be read: {error}")
            if context:
                Text(context, css_class=muted)
            if status:
                Text(status)
            if note_text:
                Text(note_text)
            if image_data_uri:
                Image(src=image_data_uri, alt=f"Current SAP screen — {heading}")
            if rows:
                Separator()
                with Card():
                    with CardHeader():
                        CardTitle(
                            f"Values about to be written ({len(rows) + overflow})"
                        )
                    with CardContent():
                        with Table():
                            with TableHeader():
                                with TableRow():
                                    TableHead("Field")
                                    TableHead("Value")
                            with TableBody():
                                for name, value in rows:
                                    with TableRow():
                                        TableCell(name)
                                        TableCell(value, css_class="font-mono")
                        if overflow:
                            Text(_overflow_line(overflow), css_class=muted)
            Separator()
            Text(f"{_CLOSING_LINE} {_FOOTER_LINE}", css_class=muted)
    return app
