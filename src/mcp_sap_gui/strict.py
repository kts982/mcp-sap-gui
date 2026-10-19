"""Opt-in strict error mode: failed tool results become MCP errors.

Most SAP failures do not raise. The controller catches the COM exception and
returns the tool's normal dict with an ``error`` key (``_error_result``), so
the client receives an ordinary result with ``isError: false``. A client that
checks only ``isError`` then treats a refused write as done.

``result_error`` is the one definition of a failed result. ``AuditMiddleware``
uses it in every mode, so the audit status follows the result.
``StrictErrorMiddleware`` uses it only when the server runs with
``--strict-errors`` (or ``MCP_SAP_GUI_STRICT_ERRORS=1``): the result is then
returned with ``is_error=True`` and its structured content unchanged, so the
client still gets ``available_keys``, ``hint`` and ``screen``.
"""

from __future__ import annotations

from typing import Any, Callable, Mapping

import mcp.types
from fastmcp.server.middleware import CallNext, Middleware, MiddlewareContext
from fastmcp.tools.tool import ToolResult

#: Environment variable that turns strict mode on, for launches (uvx, an MCP
#: client's server config) where adding a CLI flag is awkward.
STRICT_ERRORS_ENV = "MCP_SAP_GUI_STRICT_ERRORS"


def result_error(tool_name: str, structured: Any) -> str | None:
    """Return why a normally returned tool result is a failure, or ``None``.

    - Any tool: a truthy top-level ``error`` key. Nested errors (a ``screen``
      carrying the status-bar text of an action that did happen, one row of
      a table read) are not failures of the call.
    - ``sap_set_batch_fields``: a field that was not set (``failed > 0``;
      skipped read-only fields do not count), or a ``validation`` that
      carries an ``error`` (the Enter that ``validate`` asked for failed).
    """
    if not isinstance(structured, Mapping):
        return None
    error = structured.get("error")
    if error:
        return str(error)
    if tool_name == "sap_set_batch_fields":
        failed = structured.get("failed")
        if isinstance(failed, int) and failed > 0:
            return f"{failed} field(s) were not set"
        validation = structured.get("validation")
        if isinstance(validation, Mapping) and validation.get("error"):
            return f"validation Enter failed: {validation['error']}"
    return None


def _env_flag(value: str | None) -> bool:
    return (value or "").strip().lower() in {"1", "true", "yes", "on"}


def strict_errors_from_env(environ: Mapping[str, str]) -> bool:
    """Read ``MCP_SAP_GUI_STRICT_ERRORS`` (1/true/yes/on)."""
    return _env_flag(environ.get(STRICT_ERRORS_ENV))


class StrictErrorMiddleware(Middleware):
    """Return a failed result as an MCP error (``isError: true``).

    Register right after ``AuditMiddleware`` and before
    ``ConfirmationMiddleware``: audit then sees what it returns, and the
    gate's ``confirmed_by_user`` (which it adds only to a result that is not
    an error) is already in the result this middleware flags.

    ``enabled`` is injected by the running server module, for the same reason
    ``ConfirmationMiddleware`` gets its helpers injected: ``main()`` rebinds
    the server's ``config`` after this middleware is registered, and a lookup
    by module name could read a second copy of the module.

    The result is returned, not raised. ``ToolResult(is_error=True)`` keeps
    ``structured_content`` (fastmcp sends it as a ``CallToolResult``), where a
    ``ToolError`` would flatten it into one line of text.
    """

    def __init__(self, *, enabled: Callable[[], bool]) -> None:
        self._enabled = enabled

    async def on_call_tool(
        self,
        context: MiddlewareContext[mcp.types.CallToolRequestParams],
        call_next: CallNext[mcp.types.CallToolRequestParams, ToolResult],
    ) -> ToolResult:
        result = await call_next(context)
        if not isinstance(result, ToolResult) or result.is_error:
            return result
        if not self._enabled():
            return result
        if result_error(context.message.name, result.structured_content) is None:
            return result
        return result.model_copy(update={"is_error": True})
