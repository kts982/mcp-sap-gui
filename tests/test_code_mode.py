"""Tests for the experimental --code-mode transform.

Mock variants of the spike safety gates (see .dev/CODE_MODE_SPIKE.local.md):
T1 session binding through the sandbox boundary, T2 audit trail of inner
calls, T3 policy-profile visibility inside execute. Live gates run against
a real SAP system separately.
"""

import json
import logging
from unittest.mock import MagicMock, patch

import pytest

pytest.importorskip("pydantic_monty", reason="code-mode extra not installed")

from fastmcp import Client, Context, FastMCP  # noqa: E402

import mcp_sap_gui.server as _server_mod  # noqa: E402
from mcp_sap_gui.audit import AuditMiddleware  # noqa: E402


@pytest.fixture
def mock_win32com():
    """Mock COM so the real server's lifespan can start without SAP."""
    mock = MagicMock()
    with patch.dict("sys.modules", {"win32com": mock, "win32com.client": mock.client}):
        yield mock

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_code_mode_server():
    """Small FastMCP server with the production CodeMode transform attached.

    Uses dummy tools instead of SAP tools so no COM mocking is needed; the
    transform instance is built by the same factory the real server uses.
    """
    srv = FastMCP("code-mode-test")
    sessions: dict[str, object] = {}

    @srv.tool(tags={"read"})
    async def capture_session(label: str, ctx: Context) -> dict:
        """Record the ctx.session object observed for this call."""
        sessions[label] = ctx.session
        return {"label": label}

    @srv.tool(tags={"read"})
    async def read_rows(start_row: int = 0) -> dict:
        """Paginated dummy rows (5 per page, 12 total)."""
        rows = list(range(start_row, min(start_row + 5, 12)))
        return {"rows": rows, "total_rows": 12}

    @srv.tool(tags={"write"})
    async def write_thing(value: str) -> str:
        """Write-tagged tool for policy visibility tests."""
        return f"wrote {value}"

    srv.add_transform(_server_mod._build_code_mode_transform())
    return srv, sessions


# ===========================================================================
# Catalog shape
# ===========================================================================


class TestCodeModeCatalog:
    async def test_catalog_replaced_with_meta_tools(self):
        srv, _ = _make_code_mode_server()
        async with Client(srv) as client:
            names = sorted(t.name for t in await client.list_tools())
        assert names == ["execute", "get_schema", "search", "tags"]

    async def test_execute_description_has_sap_guardrails(self):
        srv, _ = _make_code_mode_server()
        async with Client(srv) as client:
            tools = {t.name: t for t in await client.list_tools()}
        desc = tools["execute"].description
        assert "sap_get_popup_window" in desc
        assert "start_row" in desc
        assert "sap_screenshot" in desc

    async def test_search_finds_tool_by_description(self):
        srv, _ = _make_code_mode_server()
        async with Client(srv) as client:
            result = await client.call_tool(
                "search", {"query": "paginated rows table"}
            )
        assert "read_rows" in str(result.content)


# ===========================================================================
# T1 — Session binding through the sandbox boundary (the NO-GO gate)
# ===========================================================================


class TestSessionBinding:
    async def test_sandbox_tool_sees_same_ctx_session_as_direct_call(self):
        """_ctrl(ctx) keys SAP bindings by id(ctx.session); a sandbox-invoked
        tool must observe the originating client's session object."""
        srv, sessions = _make_code_mode_server()
        async with Client(srv) as client:
            await client.call_tool("capture_session", {"label": "direct"})
            await client.call_tool(
                "execute",
                {
                    "code": (
                        'r = await call_tool("capture_session",'
                        ' {"label": "sandbox"})\n'
                        "return r"
                    )
                },
            )
        assert sessions["direct"] is not None
        assert sessions["direct"] is sessions["sandbox"]


# ===========================================================================
# T2 — Audit trail covers execute plus every inner call
# ===========================================================================


class TestAuditTrail:
    async def test_inner_calls_are_audited(self, caplog):
        srv, _ = _make_code_mode_server()
        srv.add_middleware(AuditMiddleware())
        script = (
            'a = await call_tool("read_rows", {"start_row": 0})\n'
            'b = await call_tool("read_rows", {"start_row": 5})\n'
            'return {"total": a["total_rows"], "n": len(a["rows"]) + len(b["rows"])}'
        )
        with caplog.at_level(logging.INFO, logger="mcp_sap_gui.audit"):
            async with Client(srv) as client:
                result = await client.call_tool("execute", {"code": script})

        assert result.data == {"total": 12, "n": 10}
        entries = [json.loads(r.message) for r in caplog.records]
        tools_logged = [e["tool"] for e in entries if e["event"] == "tool_call"]
        assert tools_logged.count("read_rows") == 2
        assert "execute" in tools_logged
        inner = [e for e in entries if e["tool"] == "read_rows"]
        assert inner[0]["args"] == {"start_row": 0}
        assert all(e["status"] == "ok" for e in entries)


# ===========================================================================
# T3 — Policy-profile visibility holds inside execute
# ===========================================================================


class TestPolicyVisibilityInSandbox:
    async def test_disabled_write_tool_unreachable_from_sandbox(self):
        """Disabling write-tagged tools (exploration profile) must also gate
        call_tool() inside the sandbox, not just the visible catalog."""
        srv, _ = _make_code_mode_server()
        srv.disable(tags={"write"})
        async with Client(srv) as client:
            with pytest.raises(Exception) as excinfo:
                await client.call_tool(
                    "execute",
                    {
                        "code": (
                            'r = await call_tool("write_thing",'
                            ' {"value": "x"})\n'
                            "return r"
                        )
                    },
                )
        assert "write_thing" in str(excinfo.value)

    async def test_read_tool_still_reachable_after_write_disable(self):
        srv, _ = _make_code_mode_server()
        srv.disable(tags={"write"})
        async with Client(srv) as client:
            result = await client.call_tool(
                "execute",
                {
                    "code": (
                        'r = await call_tool("read_rows", {"start_row": 0})\n'
                        'return r["total_rows"]'
                    )
                },
            )
        # Scalar returns surface as text content only (no structured data).
        assert result.content[0].text == "12"


class TestPolicyProfileToolWithCodeMode:
    async def test_profile_switch_preserves_meta_tools(self, mock_win32com):
        """Regression: sap_set_policy_profile previously used
        disable_components(match_all=True), which removed the untagged
        code-mode meta-tools for the session — 'execute' disappeared and
        never came back. Subtractive semantics must leave them alone."""
        transform = _server_mod._build_code_mode_transform()
        _server_mod.mcp.add_transform(transform)
        try:
            async with Client(_server_mod.mcp) as client:
                r = await client.call_tool(
                    "execute",
                    {
                        "code": (
                            'r = await call_tool("sap_set_policy_profile",'
                            ' {"profile": "exploration"})\n'
                            "return r"
                        )
                    },
                )
                assert r.data["profile"] == "exploration"
                names = {t.name for t in await client.list_tools()}
                assert "execute" in names, "meta-tool lost after profile switch"
                r2 = await client.call_tool(
                    "execute",
                    {
                        "code": (
                            'r = await call_tool("sap_set_policy_profile",'
                            ' {"profile": "full"})\n'
                            "return r"
                        )
                    },
                )
                assert r2.data["profile"] == "full"
        finally:
            # private attr, but the test must not leak the transform into
            # the shared module-level server used by other test files
            _server_mod.mcp._transforms.remove(transform)


# ===========================================================================
# Sandbox behavior notes encoded as tests
# ===========================================================================


class TestSandboxBehavior:
    async def test_pagination_loop_aggregates_in_sandbox(self):
        """The archetype flow: N paginated reads collapse into one execute."""
        srv, _ = _make_code_mode_server()
        script = (
            "all_rows = []\n"
            "start = 0\n"
            "while True:\n"
            '    page = await call_tool("read_rows", {"start_row": start})\n'
            '    all_rows.extend(page["rows"])\n'
            '    if len(all_rows) >= page["total_rows"]:\n'
            "        break\n"
            "    start += 5\n"
            "return {'count': len(all_rows), 'last': all_rows[-1]}"
        )
        async with Client(srv) as client:
            result = await client.call_tool("execute", {"code": script})
        assert result.data == {"count": 12, "last": 11}

    async def test_hidden_tools_remain_directly_callable(self):
        """CodeMode replaces the catalog listing, not reachability: clients
        that know a tool name (e.g. sap_screenshot) can still call it
        directly. Policy enforcement must therefore stay tag/in-tool based."""
        srv, _ = _make_code_mode_server()
        async with Client(srv) as client:
            result = await client.call_tool("read_rows", {"start_row": 10})
        assert result.data["rows"] == [10, 11]
