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
import logging
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Callable, Iterable, Mapping

import mcp.types
from fastmcp.exceptions import ToolError
from fastmcp.server.middleware import CallNext, Middleware, MiddlewareContext
from fastmcp.tools.tool import ToolResult
from mcp.shared.exceptions import McpError

from .audit import logger as audit_logger
from .preview import _normalize_fields, _overflow_line

logger = logging.getLogger(__name__)

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
    "attach",
)

#: Not a yes/no gate: ``sap_connect_existing`` asks in its tool body which
#: session to attach to (a choice dialog), so no tool maps to it here.
ATTACH_POINT = "attach"

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
#: Popups: sap_handle_popup and sap_press_button (on a button outside wnd[0])
#: ask in their tool bodies before a press that may save (a Save button, or
#: Yes/OK on a popup about saving) and before any answer but Cancel to a popup
#: question the tool cannot read, using the sap_send_key Save prompt.
#: sap_send_key always targets wnd[0]: live on SAP GUI 8.10, Enter there left
#: an open confirmation popup untouched.
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

#: Nobody reads a dialog and answers it this fast: an answer quicker than
#: this came from the client, which showed no dialog. Live on 2026-10-19 the
#: Claude Code desktop app answered "decline" in 7-13 ms, in auto and in
#: manual mode, while the Claude Code CLI showed the dialog.
NO_DIALOG_MS = 250.0


async def timed_elicit(ctx: Any, **kwargs: Any) -> tuple[Any, float]:
    """``ctx.elicit(**kwargs)`` and how long the answer took, in ms."""
    start = time.perf_counter()
    result = await ctx.elicit(**kwargs)
    return result, (time.perf_counter() - start) * 1000


def no_dialog_note(elapsed_ms: float) -> str:
    """A sentence for a decline that came back too fast for a person, else ''.

    The decline still counts (fail closed); the note keeps the error from
    claiming that the user declined, and tells the agent what to do.
    """
    if elapsed_ms >= NO_DIALOG_MS:
        return ""
    return (
        f" The client answered in {elapsed_ms:.0f} ms without showing a "
        "dialog, so no person saw this request. It cannot be approved in "
        "this client: tell the user, and use a client that shows MCP "
        "confirmation dialogs."
    )


def fast_yes_note(elapsed_ms: float) -> str:
    """The sentence for a yes that came back too fast for a person.

    No person can read a dialog, tick its box and press Accept this fast, so
    the yes came from the client (or something answering for it) and is
    refused: a gate exists so that a person approves.
    """
    return (
        f" The client answered yes in {elapsed_ms:.0f} ms without showing a "
        "dialog, so no person approved this, and it counts as declined. Tell "
        "the user, and use a client that shows MCP confirmation dialogs."
    )


def log_confirmation_event(point: str, tool: str, outcome: str,
                           elapsed_ms: float | None = None,
                           client_answer: str | None = None) -> None:
    """Emit a confirmation outcome to the audit log.

    AuditMiddleware records only whether the call failed, not which point
    prompted or what the user answered — this is the only record of that.
    ``elapsed_ms`` is how long the answer took; ``no_dialog`` marks one that
    came back faster than a person can answer (``NO_DIALOG_MS``), and
    ``client_answer`` what the client said when that answer was overruled.
    """
    line: dict[str, Any] = {
        "event": "confirmation",
        "ts": datetime.now(timezone.utc).isoformat(),
        "point": point,
        "tool": tool,
        "outcome": outcome,
    }
    if elapsed_ms is not None:
        line["elapsed_ms"] = round(elapsed_ms, 1)
        if elapsed_ms < NO_DIALOG_MS:
            line["no_dialog"] = True
    if client_answer is not None:
        line["client_answer"] = client_answer
    audit_logger.info(json.dumps(line, default=str))


@dataclass
class Answer:
    """What ask_user found out.

    ``yes`` is true only when a person approved (or picked a valid option).
    ``action`` is the client's action, ``decline`` for an overruled fast yes,
    or ``invalid`` for an accepted answer that is not one of the options.
    ``note`` is empty when a person answered; otherwise it explains why no
    person did, for the caller to append to its error.
    """

    yes: bool
    action: str
    data: Any
    elapsed_ms: float
    note: str


def _is_true(data: Any) -> bool:
    return data is True


async def ask_user(
    ctx: Any,
    *,
    points: str | Iterable[str],
    tool: str,
    message: str,
    response_type: Any = bool,
    response_title: str | None = None,
    is_yes: Callable[[Any], bool] = _is_true,
    is_valid: Callable[[Any], bool] | None = None,
) -> Answer:
    """Ask the user one question; the same rules for every gate.

    - The answer is timed (``timed_elicit``). An answer faster than
      ``NO_DIALOG_MS`` came from the client, not a person: a decline gets
      ``no_dialog_note``, and a yes is refused with ``fast_yes_note``.
    - ``is_yes(data)`` decides what an accepted answer means (a ticked box
      by default). ``is_valid(data)``, when given, separates a wrong answer
      (``action: invalid``) from a no.
    - Every outcome is written to the audit log, one line per point.
    - ``McpError`` (a client without elicitation) is not caught: each gate
      fails closed with its own message.
    """
    point_list = [points] if isinstance(points, str) else list(points)
    kwargs: dict[str, Any] = {"message": message, "response_type": response_type}
    if response_title is not None:
        kwargs["response_title"] = response_title
    result, elapsed = await timed_elicit(ctx, **kwargs)
    data = getattr(result, "data", None)
    accepted = result.action == "accept"

    def log(outcome: str, client_answer: str | None = None) -> None:
        for point in point_list:
            log_confirmation_event(point, tool, outcome, elapsed, client_answer)

    if accepted and is_valid is not None and not is_valid(data):
        log("invalid_answer")
        return Answer(False, "invalid", data, elapsed, "")
    if accepted and is_yes(data):
        if elapsed < NO_DIALOG_MS:
            log("declined", client_answer="accept")
            return Answer(False, "decline", data, elapsed, fast_yes_note(elapsed))
        log("accepted")
        return Answer(True, "accept", data, elapsed, "")
    log("declined")
    return Answer(False, result.action, data, elapsed, no_dialog_note(elapsed))


def log_attach_event(session: Mapping[str, Any], chosen_by: str) -> None:
    """Record which SAP session the agent was attached to, and how it was
    chosen: ``named`` (the call named the system), ``dialog`` (the user
    picked it) or ``index`` (connection and session index)."""
    audit_logger.info(
        json.dumps({
            "event": "attach",
            "ts": datetime.now(timezone.utc).isoformat(),
            "system": session.get("system"),
            "client": session.get("client"),
            "user": session.get("user"),
            "session": session.get("id"),
            "chosen_by": chosen_by,
        }, default=str),
    )


# ---------------------------------------------------------------------------
# Middleware
# ---------------------------------------------------------------------------

def _mark_confirmed(result: ToolResult) -> ToolResult:
    """Add ``confirmed_by_user: true`` to an accepted call's result.

    The prompt goes to the user, not the agent: without the flag the agent
    cannot tell an approved write from an ungated one.  Only a plain dict
    result (one JSON text block) is rewritten; anything else passes as is.
    """
    data = result.structured_content
    if (
        result.is_error
        or not isinstance(data, dict)
        or len(result.content) != 1
        or not isinstance(result.content[0], mcp.types.TextContent)
    ):
        return result
    marked = {**data, "confirmed_by_user": True}
    return ToolResult(content=marked, structured_content=marked, meta=result.meta)


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

    ``active_points`` and ``precheck`` are injected by the RUNNING server
    module at registration.  The middleware must never import the server
    module by name: under ``python -m mcp_sap_gui.server`` the running module
    is ``__main__``, and a ``from . import server`` here loads a SECOND copy
    of server.py with its own empty globals (``_session_mgr = None``, default
    config) — the gate then reads confirmation points from the wrong module
    and silently never fires.  Found live; pinned by tests/test_stdio_gate.py.
    """

    def __init__(self, *, active_points, precheck) -> None:
        self._active_points = active_points
        self._precheck = precheck

    async def on_call_tool(
        self,
        context: MiddlewareContext[mcp.types.CallToolRequestParams],
        call_next: CallNext[mcp.types.CallToolRequestParams, ToolResult],
    ) -> ToolResult:
        ctx = context.fastmcp_context
        if ctx is None:
            return await call_next(context)

        tool_name = context.message.name
        args = context.message.arguments or {}

        active = self._active_points(ctx)
        tags: set[str] = set()
        if "all_writes" in active:
            # Only needed for the tag-driven point; skip the lookup otherwise.
            tags = await _tool_tags(ctx, tool_name)

        point = point_for_call(tool_name, args, active, tags)
        if point is None:
            return await call_next(context)
        logger.debug("confirmation gate: tool=%s point=%s", tool_name, point)

        # Ordering fix: this middleware runs upstream of the tool body, so a
        # call that read-only mode or the transaction policy would reject must
        # be rejected here — never prompt for a call that cannot run.
        self._precheck(tool_name, args, ctx)

        try:
            answer = await ask_user(
                ctx,
                points=point,
                tool=tool_name,
                message=build_prompt(tool_name, args, point),
                # The CLI shows the answer as a checkbox: without a title it
                # reads "Value". Only a ticked box with Accept proceeds.
                response_title="Yes, run it",
            )
        except McpError as exc:
            # Fail closed, matching the existing save gate.
            log_confirmation_event(point, tool_name, "unsupported_client")
            raise ToolError(
                f"{tool_name} is gated by confirmation point '{point}' but the "
                f"client does not support elicitation: {exc}"
            ) from exc

        if answer.yes:
            return _mark_confirmed(await call_next(context))
        raise ToolError(
            f"{tool_name} declined by user at confirmation point '{point}' "
            f"({answer.action}). Nothing was executed.{answer.note}"
        )
