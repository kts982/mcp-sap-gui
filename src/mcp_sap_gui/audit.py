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

# Patterns in field IDs or argument keys that indicate secrets
_SECRET_PATTERNS = frozenset({"password", "pwd", "bcode", "secret", "token"})

logger = logging.getLogger("mcp_sap_gui.audit")


def _is_secret_key(key: Any) -> bool:
    key_lower = str(key).lower()
    return any(p in key_lower for p in _SECRET_PATTERNS)


def _mask_secrets(args: dict[str, Any]) -> dict[str, Any]:
    """Return a copy of args with secret-looking values masked.

    Recurses one level into dict-valued arguments: tool arguments that carry
    field maps (``sap_set_batch_fields``, ``sap_preview.pending_fields``) hide
    their sensitive keys one level down, and MCP arguments never nest deeper.
    """
    masked = {}
    for key, value in args.items():
        if _is_secret_key(key):
            masked[key] = "***"
        elif isinstance(value, dict):
            masked[key] = {
                k: ("***" if _is_secret_key(k) else v) for k, v in value.items()
            }
        elif isinstance(value, str) and any(
            p in value.lower() for p in _SECRET_PATTERNS
        ):
            # Field IDs referencing password fields pass through unchanged
            masked[key] = value
        else:
            masked[key] = value
    return masked


class AuditMiddleware(Middleware):
    """Logs every tool call with structured JSON for audit and analysis.

    Emits one JSON log line per tool call containing:
    - tool_name, arguments (secrets masked), timestamp
    - duration_ms, status (ok/error), error details
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

            logger.info(
                json.dumps({
                    "event": "tool_call",
                    "ts": ts,
                    "tool": tool_name,
                    "args": safe_args,
                    "status": "ok",
                    "duration_ms": duration_ms,
                }, default=str),
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
