"""Confirmation points: blocking user confirmation before write operations.

A *confirmation point* is a named category of write operations that requires a
blocking ``ctx.elicit`` round-trip to the user before the server executes it.
Points are resolved per MCP session (``sap_set_confirmation_points``) on top of
an immutable server floor (``--confirm``).  ``save`` is always on and can never
be removed.

``ConfirmationMiddleware`` is the single chokepoint.  It is registered AFTER
``AuditMiddleware`` so that a blocked call still produces an audit line
(``add_middleware`` order is outermost-first), and it also covers the
``--code-mode`` sandbox because ``FastMCP.call_tool`` re-enters the middleware
chain.

Two consequences of running as middleware, both handled here:

* The gate runs BEFORE the tool body, i.e. before ``_check_write()`` and
  ``_enforce_transaction_policy()``.  Those checks are re-run here so a call
  that was going to be rejected is rejected without prompting the user.
* Blocking must ``raise`` (not return a bare ``ToolResult``): a content-only
  result fails fastmcp's outputSchema validation for ``-> dict`` tools.

The middleware never touches COM; every SAP operation still happens inside the
tool body on the single-threaded COM executor.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from typing import Any, Iterable, Mapping

import mcp.types
from fastmcp.exceptions import ToolError
from fastmcp.server.middleware import CallNext, Middleware, MiddlewareContext
from fastmcp.tools.tool import ToolResult
from mcp.shared.exceptions import McpError

from .audit import logger as audit_logger
from .preview import _normalize_fields, _overflow_line

# ---------------------------------------------------------------------------
# Point taxonomy
# ---------------------------------------------------------------------------

#: Always on, never settable, never removable.  Covers the existing F11/Save
#: elicitation in ``sap_send_key`` (which stays in the tool body, untouched)
#: plus the Save toolbar button handled by this middleware.
SAVE_POINT = "save"

#: Points a session or the ``--confirm`` floor may activate.
SETTABLE_POINTS: tuple[str, ...] = (
    "batch_fields",
    "field_writes",
    "transactions",
    "all_writes",
)

ALL_POINTS: tuple[str, ...] = (SAVE_POINT, *SETTABLE_POINTS)

#: Static tool -> point map.  ``all_writes`` is resolved from the tool's
#: ``write`` tag instead, so a newly added write tool is covered automatically.
POINT_BY_TOOL: dict[str, str] = {
    "sap_execute_transaction": "transactions",
    "sap_set_batch_fields": "batch_fields",
    "sap_set_field": "field_writes",
    "sap_modify_cell": "field_writes",
    "sap_set_textedit": "field_writes",
    "sap_select_checkbox": "field_writes",
    "sap_select_radio_button": "field_writes",
    "sap_select_combobox_entry": "field_writes",
}

#: The SAP-standard Save button on the system toolbar.  Pressing it reaches the
#: same code path as F11, so it must hit the ``save`` point too — otherwise
#: ``sap_press_button`` is a complete bypass of the save gate.  Menu paths
#: (System -> Save) and ALV/app-toolbar Save buttons remain a documented gap:
#: their IDs are positional and locale-dependent, so there is no reliable
#: static ID to match.
#: Matched by suffix, so a popup's own tbar[0]/btn[11] (e.g. wnd[1]) also hits
#: the gate. Intentional: over-prompting on a dialog errs on the safe side.
_SAVE_BUTTON_SUFFIX = "tbar[0]/btn[11]"

#: Human one-liners for the elicitation prompt.
_TOOL_ACTIONS: dict[str, str] = {
    "sap_execute_transaction": "start an SAP transaction",
    "sap_set_batch_fields": "write several fields at once",
    "sap_set_field": "write a value into a screen field",
    "sap_modify_cell": "write a value into a table/grid cell",
    "sap_set_textedit": "replace the contents of a text editor",
    "sap_select_checkbox": "change a checkbox",
    "sap_select_radio_button": "select a radio button",
    "sap_select_combobox_entry": "change a dropdown selection",
    "sap_press_button": "press the SAP Save button (system toolbar)",
}


def is_save_button(button_id: Any) -> bool:
    """Return True when *button_id* addresses the Save toolbar button."""
    if not isinstance(button_id, str):
        return False
    normalized = button_id.strip().replace("\\", "/").lower()
    return normalized.endswith(_SAVE_BUTTON_SUFFIX)


def effective_points(
    floor: Iterable[str] | None,
    session_points: Iterable[str] | None,
) -> set[str]:
    """Return the points active for a session: floor + session + ``save``."""
    return {SAVE_POINT} | set(floor or ()) | set(session_points or ())


def points_with_provenance(
    floor: Iterable[str] | None,
    session_points: Iterable[str] | None,
) -> list[dict[str, str]]:
    """Return the effective points, each tagged with where it came from.

    ``server-floor`` wins over ``session`` because a floor point stays active
    whether or not the session also asked for it.
    """
    floor_set = set(floor or ())
    rows: list[dict[str, str]] = []
    for point in sorted(effective_points(floor_set, session_points)):
        if point == SAVE_POINT:
            source = "always-on"
        elif point in floor_set:
            source = "server-floor"
        else:
            source = "session"
        rows.append({"point": point, "source": source})
    return rows


def is_effective_write(tool_name: str, args: Mapping[str, Any], tags: Iterable[str]) -> bool:
    """Return True when this *call* actually mutates (arg-aware, not tag-only).

    ``sap_get_tree_node_children`` is ``write``-tagged but only mutates when
    ``expand=True`` — its own ``_check_write()`` is conditional in the same way,
    so gating the read form would both prompt and (in read-only mode) reject a
    call the server otherwise allows.
    """
    if "write" not in set(tags):
        return False
    if tool_name == "sap_get_tree_node_children":
        return bool(args.get("expand", False))
    return True


def point_for_call(
    tool_name: str,
    args: Mapping[str, Any],
    active_points: Iterable[str],
    tags: Iterable[str],
) -> str | None:
    """Resolve the confirmation point that gates this call, or ``None``.

    ``sap_send_key`` is deliberately absent from the ``save`` branch: its F11 /
    Save elicitation lives in the tool body and stays there.
    """
    active = set(active_points)

    if tool_name == "sap_press_button" and is_save_button(args.get("button_id")):
        return SAVE_POINT  # always on

    mapped = POINT_BY_TOOL.get(tool_name)
    if mapped is not None and mapped in active:
        return mapped

    if "all_writes" in active and is_effective_write(tool_name, args, tags):
        return "all_writes"

    return None


# ---------------------------------------------------------------------------
# Prompt rendering
# ---------------------------------------------------------------------------

def _prompt_fields(tool_name: str, args: Mapping[str, Any]) -> Mapping[str, Any]:
    """Map tool arguments onto the ``field -> value`` lines shown to the user."""
    if tool_name == "sap_set_batch_fields":
        fields = args.get("fields")
        return fields if isinstance(fields, Mapping) else {}
    if tool_name == "sap_set_field":
        return {args.get("field_id", "?"): args.get("value", "")}
    if tool_name == "sap_set_textedit":
        return {args.get("textedit_id", "?"): args.get("text", "")}
    if tool_name == "sap_modify_cell":
        cell = f"{args.get('grid_id', '?')}[{args.get('row', '?')}].{args.get('column', '?')}"
        return {cell: args.get("value", "")}
    if tool_name == "sap_select_checkbox":
        return {args.get("checkbox_id", "?"): args.get("selected", True)}
    if tool_name == "sap_select_radio_button":
        return {args.get("radio_id", "?"): "selected"}
    if tool_name == "sap_select_combobox_entry":
        return {args.get("combobox_id", "?"): args.get("key_or_value", "")}
    if tool_name == "sap_execute_transaction":
        return {"transaction": args.get("tcode", "")}
    if tool_name == "sap_press_button":
        return {"button": args.get("button_id", "")}
    # all_writes catch-all: show the raw arguments.  _normalize_fields masks
    # password-like keys and clips long values, so this is safe to render.
    return {k: v for k, v in args.items() if k != "ctx"}


#: An elicitation is a modal prompt, not a report: past ~15 rows the user is
#: scrolling a dialog to find the approve button.  The full list is
#: sap_preview's job; preview.py keeps its own (larger) cap for that.
_MAX_PROMPT_FIELD_ROWS = 15


def build_prompt(tool_name: str, args: Mapping[str, Any], point: str) -> str:
    """Render the elicitation message for a gated call."""
    action = _TOOL_ACTIONS.get(tool_name, "perform a write operation in SAP")
    lines = [
        f"Confirmation point '{point}' is active.",
        f"The agent wants to {action} via {tool_name}.",
    ]

    rows, overflow = _normalize_fields(_prompt_fields(tool_name, args))
    if len(rows) > _MAX_PROMPT_FIELD_ROWS:
        overflow += len(rows) - _MAX_PROMPT_FIELD_ROWS
        rows = rows[:_MAX_PROMPT_FIELD_ROWS]
    if rows:
        lines.append("")
        lines.extend(f"  {label} = {value}" for label, value in rows)
        if overflow:
            lines.append(f"  {_overflow_line(overflow)}")
        lines.append("")

    lines.append("Do you want to proceed?")
    return "\n".join(lines)


def build_relax_prompt(points: Iterable[str]) -> str:
    """Render the elicitation message for disabling confirmation points."""
    names = ", ".join(sorted(points))
    return (
        f"The agent asks to disable confirmation for: {names}. "
        "Those operations would then run without asking you first. "
        "Allow?"
    )


# ---------------------------------------------------------------------------
# Audit
# ---------------------------------------------------------------------------

def log_confirmation_event(point: str, tool: str, outcome: str) -> None:
    """Emit a confirmation outcome to the audit log.

    AuditMiddleware never inspects results, so it cannot tell an accepted
    confirmation from a declined one — this is the only record of the outcome.
    """
    audit_logger.info(
        json.dumps({
            "event": "confirmation",
            "ts": datetime.now(timezone.utc).isoformat(),
            "point": point,
            "tool": tool,
            "outcome": outcome,
        }, default=str),
    )


# ---------------------------------------------------------------------------
# Middleware
# ---------------------------------------------------------------------------

async def _tool_tags(ctx: Any, tool_name: str) -> set[str]:
    """Look up a tool's tags through the live server.

    Only consulted when ``all_writes`` is active.  A failed lookup (unknown
    tool, transformed catalog, no server on ctx) returns ``{"write"}`` —
    fail closed: a spurious prompt beats a silent write.
    """
    try:
        tool = await ctx.fastmcp.get_tool(tool_name)
    except Exception:
        return {"write"}
    return set(getattr(tool, "tags", None) or ())


class ConfirmationMiddleware(Middleware):
    """Gate write tools behind a blocking user confirmation.

    Register AFTER ``AuditMiddleware`` so blocked calls still get an audit line.
    """

    async def on_call_tool(
        self,
        context: MiddlewareContext[mcp.types.CallToolRequestParams],
        call_next: CallNext[mcp.types.CallToolRequestParams, ToolResult],
    ) -> ToolResult:
        # Deferred import: server.py imports this module at module scope, so
        # the reverse edge can only be resolved per call.
        from . import server

        ctx = context.fastmcp_context
        if ctx is None:
            return await call_next(context)

        tool_name = context.message.name
        args = context.message.arguments or {}

        active = server.active_confirmation_points(ctx)
        tags: set[str] = set()
        if "all_writes" in active:
            # Only needed for the tag-driven point; skip the lookup otherwise.
            tags = await _tool_tags(ctx, tool_name)

        point = point_for_call(tool_name, args, active, tags)
        if point is None:
            return await call_next(context)

        # Ordering fix: this middleware runs upstream of the tool body, so a
        # call that read-only mode or the transaction policy would reject must
        # be rejected here — never prompt for a call that cannot run.
        server.precheck_before_confirmation(tool_name, args)

        try:
            result = await ctx.elicit(
                message=build_prompt(tool_name, args, point),
                response_type=bool,
            )
        except McpError as exc:
            # Fail closed, matching the existing save gate.
            log_confirmation_event(point, tool_name, "unsupported_client")
            raise ToolError(
                f"{tool_name} is gated by confirmation point '{point}' but the "
                f"client does not support elicitation: {exc}"
            ) from exc

        if result.action == "accept" and getattr(result, "data", None) is True:
            log_confirmation_event(point, tool_name, "accepted")
            return await call_next(context)

        log_confirmation_event(point, tool_name, "declined")
        raise ToolError(
            f"{tool_name} declined by user at confirmation point '{point}' "
            f"({result.action}). Nothing was executed."
        )
