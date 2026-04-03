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


def _mask_secrets(args: dict[str, Any]) -> dict[str, Any]:
    """Return a copy of args with secret-looking values masked."""
    masked = {}
    for key, value in args.items():
        key_lower = key.lower()
        if any(p in key_lower for p in _SECRET_PATTERNS):
            masked[key] = "***"
        elif isinstance(value, str) and any(p in value.lower() for p in _SECRET_PATTERNS):
            # Mask values that reference password fields (e.g., field_id="wnd[0]/usr/pwdRSYST-BCODE")
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
