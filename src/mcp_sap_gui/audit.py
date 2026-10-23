"""
Audit middleware for structured logging of MCP tool calls.

Logs every tool invocation with arguments, timing, and outcome.
Designed to feed downstream analysis (playbook generation, compliance).
"""

import json
import logging
import time
from typing import Any

import mcp.types
from fastmcp.server.middleware import CallNext, Middleware, MiddlewareContext
from fastmcp.tools.tool import ToolResult

from .strict import result_error

# Patterns in field IDs or argument keys that indicate secrets
_SECRET_PATTERNS = frozenset({"password", "pwd", "bcode", "secret", "token"})

logger = logging.getLogger("mcp_sap_gui.audit")


def _is_secret_key(key: Any) -> bool:
    key_lower = str(key).lower()
    return any(p in key_lower for p in _SECRET_PATTERNS)


# Arguments that carry what is written into the target another argument names
# (sap_set_field's value, sap_set_textedit's text, sap_modify_cell's value).
_WRITTEN_VALUE_KEYS = frozenset({"value", "text"})


def _mask_secrets(args: dict[str, Any]) -> dict[str, Any]:
    """Return a copy of args with secret-looking values masked.

    - A key that looks secret (``password``, ``pwd``, ``bcode``, ...) has its
      value masked.
    - When another argument names a secret-looking target (sap_set_field's
      ``field_id`` "wnd[0]/usr/pwdRSYST-BCODE", sap_modify_cell's ``column``
      "PASSWORD"), the value written into it (``value`` / ``text``) is
      masked. The target itself stays visible: the audit shows which field
      was written, not what.
    - Recurses one level into dict-valued arguments: tool arguments that carry
      field maps (``sap_set_batch_fields``, ``sap_preview.pending_fields``)
      hide their sensitive keys one level down, and MCP arguments never nest
      deeper.
    """
    secret_target = any(
        isinstance(value, str) and _is_secret_key(value)
        for key, value in args.items()
        if key not in _WRITTEN_VALUE_KEYS
    )
    masked = {}
    for key, value in args.items():
        if _is_secret_key(key):
            masked[key] = "***"
        elif secret_target and key in _WRITTEN_VALUE_KEYS:
            masked[key] = "***"
        elif isinstance(value, dict):
            masked[key] = {
                k: ("***" if _is_secret_key(k) else v) for k, v in value.items()
            }
        else:
            masked[key] = value
    return masked


class AuditMiddleware(Middleware):
    """Logs every tool call with structured JSON for audit and analysis.

    Emits one JSON log line per tool call containing:
    - tool_name, arguments (secrets masked), timestamp
    - duration_ms, status (ok/error), error details

    ``status`` is ``error`` when the tool raised, and also when it returned a
    result that ``strict.result_error`` counts as failed (``error_msg`` then
    carries the reason). Exceptions keep their ``error`` (type) and
    ``error_msg`` keys.
    """

    async def on_call_tool(
        self,
        context: MiddlewareContext[mcp.types.CallToolRequestParams],
        call_next: CallNext[mcp.types.CallToolRequestParams, ToolResult],
    ) -> ToolResult:
        tool_name = context.message.name
        raw_args = context.message.arguments or {}
        safe_args = _mask_secrets(raw_args)
        ts = context.timestamp.isoformat()

        start = time.perf_counter()
        try:
            result = await call_next(context)
            duration_ms = round((time.perf_counter() - start) * 1000, 1)

            # Most SAP failures come back as a normal result with an
            # ``error`` key, not as an exception: the status follows the
            # result, in strict mode and without it.
            error_msg = result_error(
                tool_name, getattr(result, "structured_content", None),
            )
            failed = error_msg is not None or (
                isinstance(result, ToolResult) and result.is_error
            )
            line: dict[str, Any] = {
                "event": "tool_call",
                "ts": ts,
                "tool": tool_name,
                "args": safe_args,
                "status": "error" if failed else "ok",
                "duration_ms": duration_ms,
            }
            if error_msg is not None:
                line["error_msg"] = error_msg
            (logger.warning if failed else logger.info)(
                json.dumps(line, default=str),
            )
            return result

        except Exception as exc:
            duration_ms = round((time.perf_counter() - start) * 1000, 1)

            logger.warning(
                json.dumps({
                    "event": "tool_call",
                    "ts": ts,
                    "tool": tool_name,
                    "args": safe_args,
                    "status": "error",
                    "duration_ms": duration_ms,
                    "error": type(exc).__name__,
                    "error_msg": str(exc),
                }, default=str),
            )
            raise
