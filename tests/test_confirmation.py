"""Tests for confirmation points: the middleware gate, the tool, and the CLI floor.

Client-driven tests use fastmcp's in-memory ``Client`` with an elicitation
handler so accept/decline round-trips exercise the real middleware chain
(audit -> confirmation -> tool), which direct tool calls bypass.
"""

import json
import logging
from contextlib import contextmanager
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from fastmcp import Client
from fastmcp.client.elicitation import ElicitResult

import mcp_sap_gui.confirmation as _confirm_mod
import mcp_sap_gui.server as _server_mod
from mcp_sap_gui.session_manager import SessionManager

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


@pytest.fixture
def srv():
    """Configure the server module with a fresh SessionManager and config."""
    original = _server_mod.config
    _server_mod._session_mgr = SessionManager()
    _server_mod.config = _server_mod.ServerConfig()
    yield _server_mod
    _server_mod.config = original


def _make_mock_ctx():
    """Create a mock MCP Context for direct tool calls in tests."""
    return MagicMock()


def _make_elicit_ctx(action="accept", data=True, side_effect=None):
    """Mock Context whose elicit() returns a duck-typed elicitation result."""
    ctx = _make_mock_ctx()
    result = MagicMock()
    result.action = action
    result.data = data
    ctx.elicit = AsyncMock(return_value=result, side_effect=side_effect)
    return ctx


class _Elicitor:
    """Scripted client-side elicitation handler that records every prompt."""

    def __init__(self, action="accept", value=True):
        self.action = action
        self.value = value
        self.messages: list[str] = []

    @property
    def handler(self):
        async def _handler(message, response_type, params, context):
            self.messages.append(message)
            if self.action == "accept":
                return ElicitResult(action="accept", content=response_type(value=self.value))
            return ElicitResult(action=self.action)

        return _handler


async def _fake_com(fn):
    """Replacement for server._com: run the callable inline, no COM thread."""
    return fn()


def _patched_controller(**returns):
    """Patch _ctrl/_com so tool bodies hit a MagicMock instead of SAP GUI."""
    controller = MagicMock(Busy=False)
    for name, value in returns.items():
        getattr(controller, name).return_value = value
    ctx = patch.multiple(
        _server_mod,
        _ctrl=MagicMock(return_value=controller),
        _com=_fake_com,
    )
    return controller, ctx


def _text(result):
    return " ".join(getattr(c, "text", "") for c in result.content)


@contextmanager
def _cli(srv, argv):
    """Run server.main() hermetically: no transport, no developer .env."""
    with patch("sys.argv", argv), \
         patch.object(srv, "load_dotenv", MagicMock()), \
         patch.object(srv.mcp, "run", MagicMock()):
        yield


# ===========================================================================
# Point resolution (pure functions)
# ===========================================================================


class TestPointResolution:
    """Static tool->point map, arg-aware refinements, save-button closure."""

    def test_save_button_recognised(self):
        assert _confirm_mod.is_save_button("wnd[0]/tbar[0]/btn[11]")
        assert _confirm_mod.is_save_button("/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[11]")
        assert not _confirm_mod.is_save_button("wnd[0]/tbar[1]/btn[8]")
        assert not _confirm_mod.is_save_button(None)

    def test_save_button_gated_without_any_session_points(self):
        point = _confirm_mod.point_for_call(
            "sap_press_button", {"button_id": "wnd[0]/tbar[0]/btn[11]"},
            {"save"}, {"write"},
        )
        assert point == "save"

    def test_other_buttons_not_gated(self):
        assert _confirm_mod.point_for_call(
            "sap_press_button", {"button_id": "wnd[0]/tbar[1]/btn[8]"},
            {"save"}, {"write"},
        ) is None

    def test_send_key_is_not_mapped_to_save(self):
        """The F11 gate stays in the sap_send_key body — no double prompt."""
        assert _confirm_mod.point_for_call(
            "sap_send_key", {"key": "F11"}, {"save"}, {"write"},
        ) is None

    def test_field_writes_only_when_active(self):
        args = {"field_id": "wnd[0]/usr/txtX", "value": "1"}
        assert _confirm_mod.point_for_call("sap_set_field", args, {"save"}, {"write"}) is None
        assert _confirm_mod.point_for_call(
            "sap_set_field", args, {"save", "field_writes"}, {"write"},
        ) == "field_writes"

    def test_all_writes_covers_unmapped_write_tools(self):
        assert _confirm_mod.point_for_call(
            "sap_press_button", {"button_id": "wnd[0]/tbar[1]/btn[8]"},
            {"save", "all_writes"}, {"write"},
        ) == "all_writes"

    def test_all_writes_never_gates_read_tools(self):
        assert _confirm_mod.point_for_call(
            "sap_read_table", {"table_id": "x"}, {"save", "all_writes"}, {"read"},
        ) is None

    def test_tree_children_gated_only_when_expanding(self):
        """_check_write() is conditional for this tool; the gate mirrors it."""
        active = {"save", "all_writes"}
        assert _confirm_mod.point_for_call(
            "sap_get_tree_node_children", {"expand": False}, active, {"write"},
        ) is None
        assert _confirm_mod.point_for_call(
            "sap_get_tree_node_children", {"expand": True}, active, {"write"},
        ) == "all_writes"

    def test_prompt_includes_field_lines_and_masks_secrets(self):
        prompt = _confirm_mod.build_prompt(
            "sap_set_batch_fields",
            {"fields": {
                "wnd[0]/usr/txtMATNR": "MAT-001",
                "wnd[0]/usr/pwdRSYST-BCODE": "hunter2",
            }},
            "batch_fields",
        )
        assert "sap_set_batch_fields" in prompt
        assert "batch_fields" in prompt
        assert "MAT-001" in prompt
        assert "hunter2" not in prompt

    def test_prompt_caps_field_rows(self):
        """A modal prompt is not a report: the overflow line carries the rest."""
        prompt = _confirm_mod.build_prompt(
            "sap_set_batch_fields",
            {"fields": {f"wnd[0]/usr/txtF{i}": str(i) for i in range(60)}},
            "batch_fields",
        )
        rows = [ln for ln in prompt.splitlines() if ln.startswith("  wnd[0]")]
        assert len(rows) == 15
        assert "... and 45 more fields" in prompt

    def test_effective_points_always_include_save(self):
        assert _confirm_mod.effective_points([], []) == {"save"}
        assert _confirm_mod.effective_points(["transactions"], ["field_writes"]) == {
            "save", "transactions", "field_writes",
        }

    def test_provenance_sources(self):
        rows = _confirm_mod.points_with_provenance(["transactions"], ["field_writes"])
        by_point = {r["point"]: r["source"] for r in rows}
        assert by_point == {
            "save": "always-on",
            "transactions": "server-floor",
            "field_writes": "session",
        }

    def test_floor_wins_over_session_provenance(self):
        rows = _confirm_mod.points_with_provenance(["transactions"], ["transactions"])
        by_point = {r["point"]: r["source"] for r in rows}
        assert by_point["transactions"] == "server-floor"


# ===========================================================================
# sap_set_confirmation_points — point set management
# ===========================================================================


class TestConfirmationPointTool:
    """Add silently, elicit on effective removal, floor and save immutable."""

    async def test_adding_points_is_silent(self, srv):
        ctx = _make_elicit_ctx()
        result = await srv.sap_set_confirmation_points(["field_writes"], ctx)

        ctx.elicit.assert_not_called()
        assert result["active"] == ["field_writes", "save"]

    async def test_provenance_in_response(self, srv):
        ctx = _make_elicit_ctx()
        result = await srv.sap_set_confirmation_points(["transactions"], ctx)

        by_point = {r["point"]: r["source"] for r in result["points"]}
        assert by_point == {"save": "always-on", "transactions": "session"}

    async def test_removal_elicits_and_accept_removes(self, srv):
        ctx = _make_elicit_ctx()
        await srv.sap_set_confirmation_points(["field_writes"], ctx)
        ctx.elicit.reset_mock()

        result = await srv.sap_set_confirmation_points([], ctx)

        ctx.elicit.assert_called_once()
        assert "field_writes" in ctx.elicit.call_args.kwargs["message"]
        assert result["active"] == ["save"]
        assert result["removal_declined"] == []

    async def test_removal_declined_keeps_point(self, srv):
        ctx = _make_elicit_ctx()
        await srv.sap_set_confirmation_points(["field_writes"], ctx)
        ctx.elicit = AsyncMock(return_value=MagicMock(action="decline", data=None))

        result = await srv.sap_set_confirmation_points([], ctx)

        assert result["removal_declined"] == ["field_writes"]
        assert "field_writes" in result["active"]
        assert srv._session_mgr.get_confirmation_points(
            srv._session_key(ctx)
        ) == {"field_writes"}

    async def test_removal_cancelled_keeps_point(self, srv):
        ctx = _make_elicit_ctx()
        await srv.sap_set_confirmation_points(["transactions"], ctx)
        ctx.elicit = AsyncMock(return_value=MagicMock(action="cancel", data=None))

        result = await srv.sap_set_confirmation_points([], ctx)

        assert "transactions" in result["active"]

    async def test_removal_accept_false_keeps_point(self, srv):
        ctx = _make_elicit_ctx()
        await srv.sap_set_confirmation_points(["transactions"], ctx)
        ctx.elicit = AsyncMock(return_value=MagicMock(action="accept", data=False))

        result = await srv.sap_set_confirmation_points([], ctx)

        assert "transactions" in result["active"]

    async def test_floor_point_is_not_removable_and_does_not_elicit(self, srv):
        srv.config = srv.ServerConfig(confirmation_floor=["field_writes"])
        ctx = _make_elicit_ctx()

        result = await srv.sap_set_confirmation_points([], ctx)

        ctx.elicit.assert_not_called()
        by_point = {r["point"]: r["source"] for r in result["points"]}
        assert by_point["field_writes"] == "server-floor"

    async def test_save_point_is_not_settable(self, srv):
        ctx = _make_elicit_ctx()
        with pytest.raises(ValueError, match="Unknown confirmation point"):
            await srv.sap_set_confirmation_points(["save"], ctx)

    async def test_save_point_always_listed(self, srv):
        ctx = _make_elicit_ctx()
        result = await srv.sap_set_confirmation_points([], ctx)
        assert result["active"] == ["save"]
        assert result["always_on"] == ["save"]

    async def test_relax_fails_closed_without_elicitation(self, srv):
        from mcp.shared.exceptions import McpError
        from mcp.types import INVALID_REQUEST, ErrorData

        ctx = _make_elicit_ctx()
        await srv.sap_set_confirmation_points(["field_writes"], ctx)
        ctx.elicit = AsyncMock(side_effect=McpError(
            ErrorData(code=INVALID_REQUEST, message="Elicitation not supported")
        ))

        with pytest.raises(ValueError, match="does not support elicitation"):
            await srv.sap_set_confirmation_points([], ctx)

        assert srv._session_mgr.get_confirmation_points(
            srv._session_key(ctx)
        ) == {"field_writes"}

    async def test_per_session_isolation(self, srv):
        ctx_a = _make_elicit_ctx()
        ctx_b = _make_elicit_ctx()

        await srv.sap_set_confirmation_points(["field_writes"], ctx_a)
        await srv.sap_set_confirmation_points(["transactions"], ctx_b)

        assert srv.active_confirmation_points(ctx_a) == {"save", "field_writes"}
        assert srv.active_confirmation_points(ctx_b) == {"save", "transactions"}

    async def test_release_does_not_clear_session_points(self, srv):
        """Point lifetime is the MCP session, not the SAP binding.

        Clearing on release would make sap_disconnect — or even a no-op
        disconnect with no binding at all — a one-call gate bypass.
        """
        ctx = _make_elicit_ctx()
        await srv.sap_set_confirmation_points(["field_writes"], ctx)

        srv._session_mgr.release(srv._session_key(ctx))

        assert srv.active_confirmation_points(ctx) == {"save", "field_writes"}

    async def test_shutdown_clears_session_points(self, srv):
        ctx = _make_elicit_ctx()
        await srv.sap_set_confirmation_points(["field_writes"], ctx)

        srv._session_mgr.shutdown()

        assert srv.active_confirmation_points(ctx) == {"save"}

    async def test_tool_is_read_tagged(self, srv):
        """Adding safety must be possible in every profile."""
        tools = await srv.mcp.list_tools()
        tool = next(t for t in tools if t.name == "sap_set_confirmation_points")
        assert "read" in tool.tags


# ===========================================================================
# Gating through the real middleware chain
# ===========================================================================


class TestGating:
    """Each point category gates a representative tool call end to end."""

    async def test_no_points_means_no_prompt(self, srv):
        controller, patcher = _patched_controller(set_field={"status": "ok"})
        elicitor = _Elicitor()
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_set_field",
                    {"field_id": "wnd[0]/usr/txtX", "value": "1"},
                )
        assert elicitor.messages == []
        controller.set_field.assert_called_once()

    async def test_field_writes_accept_proceeds(self, srv):
        controller, patcher = _patched_controller(set_field={"status": "ok"})
        elicitor = _Elicitor("accept")
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["field_writes"]}
                )
                await client.call_tool(
                    "sap_set_field",
                    {"field_id": "wnd[0]/usr/txtMATNR", "value": "MAT-001"},
                )

        assert len(elicitor.messages) == 1
        assert "field_writes" in elicitor.messages[0]
        assert "MAT-001" in elicitor.messages[0]
        controller.set_field.assert_called_once()

    async def test_field_writes_decline_blocks(self, srv):
        controller, patcher = _patched_controller(set_field={"status": "ok"})
        elicitor = _Elicitor("decline")
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["field_writes"]}
                )
                result = await client.call_tool(
                    "sap_set_field",
                    {"field_id": "wnd[0]/usr/txtMATNR", "value": "MAT-001"},
                    raise_on_error=False,
                )

        assert result.is_error
        assert "declined by user at confirmation point 'field_writes'" in _text(result)
        controller.set_field.assert_not_called()

    async def test_field_writes_cancel_blocks(self, srv):
        controller, patcher = _patched_controller(set_field={"status": "ok"})
        elicitor = _Elicitor("cancel")
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["field_writes"]}
                )
                result = await client.call_tool(
                    "sap_set_field",
                    {"field_id": "wnd[0]/usr/txtMATNR", "value": "MAT-001"},
                    raise_on_error=False,
                )

        assert result.is_error
        controller.set_field.assert_not_called()

    async def test_transactions_point_gates_execute_transaction(self, srv):
        controller, patcher = _patched_controller(
            execute_transaction={"screen": {"transaction": "MM03"}}
        )
        elicitor = _Elicitor("decline")
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["transactions"]}
                )
                result = await client.call_tool(
                    "sap_execute_transaction", {"tcode": "MM03"}, raise_on_error=False
                )

        assert result.is_error
        assert "MM03" in elicitor.messages[0]
        controller.execute_transaction.assert_not_called()

    async def test_batch_fields_point_gates_batch_write(self, srv):
        controller, patcher = _patched_controller(set_batch_fields={"set": 2})
        elicitor = _Elicitor("accept")
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["batch_fields"]}
                )
                await client.call_tool(
                    "sap_set_batch_fields",
                    {"fields": {"wnd[0]/usr/txtA": "1", "wnd[0]/usr/txtB": "2"}},
                )

        assert "wnd[0]/usr/txtA = 1" in elicitor.messages[0]
        controller.set_batch_fields.assert_called_once()

    async def test_save_button_is_gated_without_configuration(self, srv):
        """The Save toolbar button bypassed the F11 gate before this point."""
        controller, patcher = _patched_controller(press_button={"status": "ok"})
        elicitor = _Elicitor("decline")
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                result = await client.call_tool(
                    "sap_press_button", {"button_id": "wnd[0]/tbar[0]/btn[11]"},
                    raise_on_error=False,
                )

        assert result.is_error
        assert "confirmation point 'save'" in _text(result)
        controller.press_button.assert_not_called()

    async def test_ordinary_button_is_not_gated(self, srv):
        controller, patcher = _patched_controller(press_button={"status": "ok"})
        elicitor = _Elicitor("decline")
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_press_button", {"button_id": "wnd[0]/tbar[1]/btn[8]"}
                )

        assert elicitor.messages == []
        controller.press_button.assert_called_once()

    async def test_all_writes_gates_sampled_write_tools(self, srv):
        controller, patcher = _patched_controller(
            press_button={"status": "ok"},
            select_tab={"status": "ok"},
            set_textedit={"status": "ok"},
        )
        elicitor = _Elicitor("accept")
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["all_writes"]}
                )
                await client.call_tool(
                    "sap_press_button", {"button_id": "wnd[0]/tbar[1]/btn[8]"}
                )
                await client.call_tool(
                    "sap_select_tab", {"tab_id": "wnd[0]/usr/tabsT/tabpX"}
                )
                await client.call_tool(
                    "sap_set_textedit",
                    {"textedit_id": "wnd[0]/usr/cntlED/shell", "text": "note"},
                )

        assert len(elicitor.messages) == 3
        assert all("all_writes" in m for m in elicitor.messages)
        # Accepting must still run the tool, not just clear the prompt.
        controller.press_button.assert_called_once()
        controller.select_tab.assert_called_once()
        controller.set_textedit.assert_called_once()

    async def test_all_writes_never_gates_read_tools(self, srv):
        controller, patcher = _patched_controller(
            read_field={"value": "X"},
            get_screen_info={"active_window": "wnd[0]"},
        )
        elicitor = _Elicitor("decline")
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["all_writes"]}
                )
                await client.call_tool("sap_read_field", {"field_id": "wnd[0]/usr/txtX"})
                await client.call_tool("sap_get_screen_info", {})

        assert elicitor.messages == []
        controller.read_field.assert_called_once()

    async def test_send_key_is_gated_once_by_the_body_gate(self, srv):
        """No point maps sap_send_key, so F11 keeps exactly its own gate."""
        controller, patcher = _patched_controller(send_vkey={"status": "ok"})
        elicitor = _Elicitor("accept")
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["field_writes"]}
                )
                await client.call_tool("sap_send_key", {"key": "F11"})
                await client.call_tool("sap_send_key", {"key": "Enter"})

        assert len(elicitor.messages) == 1
        assert "triggers Save (F11)" in elicitor.messages[0]
        assert controller.send_vkey.call_count == 2

    async def test_all_writes_prompts_twice_for_the_save_key(self, srv):
        """Documented consequence: all_writes covers every write-tagged tool and
        the sap_send_key body gate is unchanged, so F11 asks twice."""
        controller, patcher = _patched_controller(send_vkey={"status": "ok"})
        elicitor = _Elicitor("accept")
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["all_writes"]}
                )
                await client.call_tool("sap_send_key", {"key": "F11"})

        assert len(elicitor.messages) == 2
        assert "all_writes" in elicitor.messages[0]
        assert "triggers Save (F11)" in elicitor.messages[1]
        controller.send_vkey.assert_called_once()

    async def test_points_survive_sap_disconnect(self, srv):
        """Regression: releasing the SAP binding must not disarm the gate.

        sap_disconnect is read-tagged and ungated, so if release() dropped the
        points it would be a one-call bypass — even when nothing was connected
        and the disconnect was a no-op.
        """
        controller, patcher = _patched_controller(set_field={"status": "ok"})
        elicitor = _Elicitor("decline")
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["field_writes"]}
                )
                await client.call_tool("sap_disconnect", {})
                result = await client.call_tool(
                    "sap_set_field", {"field_id": "wnd[0]/usr/txtX", "value": "1"},
                    raise_on_error=False,
                )

        assert len(elicitor.messages) == 1
        assert result.is_error
        controller.set_field.assert_not_called()

    async def test_tag_lookup_is_skipped_when_no_points_are_active(self, srv):
        """Fast path: the catalog lookup only happens for all_writes."""
        controller, patcher = _patched_controller(set_field={"status": "ok"})
        elicitor = _Elicitor("decline")
        exploding = AsyncMock(side_effect=AssertionError("tags must not be read"))
        with patcher, patch.object(_confirm_mod, "_tool_tags", exploding):
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_set_field", {"field_id": "wnd[0]/usr/txtX", "value": "1"}
                )

        exploding.assert_not_called()
        assert elicitor.messages == []
        controller.set_field.assert_called_once()

    async def test_setting_points_is_itself_never_gated(self, srv):
        """sap_set_confirmation_points is read-tagged: adding safety never prompts."""
        _, patcher = _patched_controller()
        elicitor = _Elicitor("decline")
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["all_writes"]}
                )
                result = await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["all_writes"]}
                )

        assert elicitor.messages == []
        assert "all_writes" in result.data["active"]


class TestTagLookupFailsClosed:
    """A broken tag lookup must prompt, not wave the call through."""

    async def test_tool_tags_returns_write_on_lookup_failure(self):
        ctx = MagicMock()
        ctx.fastmcp.get_tool = AsyncMock(side_effect=RuntimeError("catalog gone"))

        assert await _confirm_mod._tool_tags(ctx, "sap_read_field") == {"write"}

    async def test_all_writes_still_prompts_when_the_lookup_raises(self, srv):
        """Driven at the middleware directly: patching mcp.get_tool through a
        Client would break tool dispatch itself."""
        from datetime import datetime, timezone

        from fastmcp.exceptions import ToolError
        from fastmcp.server.middleware import MiddlewareContext

        ctx = MagicMock()
        ctx.session = object()
        ctx.fastmcp.get_tool = AsyncMock(side_effect=RuntimeError("catalog gone"))
        ctx.elicit = AsyncMock(return_value=MagicMock(action="decline", data=None))
        srv._session_mgr.set_confirmation_points(id(ctx.session), {"all_writes"})

        params = MagicMock()
        params.name = "sap_some_unlisted_tool"
        params.arguments = {"value": "x"}
        mw_ctx = MiddlewareContext(
            message=params,
            fastmcp_context=ctx,
            timestamp=datetime.now(timezone.utc),
            method="tools/call",
        )
        ran = []

        async def _next(_ctx):
            ran.append(True)
            return MagicMock()

        with pytest.raises(ToolError, match="declined by user"):
            await _confirm_mod.ConfirmationMiddleware().on_call_tool(mw_ctx, _next)

        ctx.elicit.assert_awaited_once()
        assert ran == []


class TestFailClosed:
    """Clients without elicitation support: match the save gate (fail closed)."""

    async def test_gated_call_fails_without_elicitation_handler(self, srv):
        controller, patcher = _patched_controller(set_field={"status": "ok"})
        with patcher:
            async with Client(srv.mcp) as client:  # no elicitation_handler
                await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["field_writes"]}
                )
                result = await client.call_tool(
                    "sap_set_field", {"field_id": "wnd[0]/usr/txtX", "value": "1"},
                    raise_on_error=False,
                )

        assert result.is_error
        assert "does not support elicitation" in _text(result)
        controller.set_field.assert_not_called()

    async def test_save_gate_still_fails_closed(self, srv):
        """Pins the pre-existing sap_send_key behaviour this feature must match."""
        controller, patcher = _patched_controller(send_vkey={"status": "ok"})
        with patcher:
            async with Client(srv.mcp) as client:
                result = await client.call_tool(
                    "sap_send_key", {"key": "F11"}, raise_on_error=False
                )

        assert result.is_error
        assert "does not support elicitation" in _text(result)
        controller.send_vkey.assert_not_called()


class TestPolicyOrdering:
    """Read-only, the blocklist and the OK-code guard reject before any prompt.

    Rejections must also be byte-identical with and without an active point,
    or the error text tells the agent which points are armed.
    """

    async def _call(self, srv, points, tool, args, controller_returns):
        controller, patcher = _patched_controller(**controller_returns)
        elicitor = _Elicitor("accept")
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                if points:
                    await client.call_tool(
                        "sap_set_confirmation_points", {"points": points}
                    )
                result = await client.call_tool(tool, args, raise_on_error=False)
        return result, elicitor, controller

    async def test_read_only_wins_without_eliciting(self, srv):
        srv.config = srv.ServerConfig(read_only=True)
        args = {"field_id": "wnd[0]/usr/txtX", "value": "1"}

        ungated, _, _ = await self._call(
            srv, [], "sap_set_field", args, {"set_field": {"status": "ok"}},
        )
        gated, elicitor, controller = await self._call(
            srv, ["field_writes"], "sap_set_field", args, {"set_field": {"status": "ok"}},
        )

        assert gated.is_error and ungated.is_error
        assert "read-only mode" in _text(gated)
        # Byte-identical: the gate must not be observable through the error.
        assert _text(gated) == _text(ungated)
        assert elicitor.messages == []
        controller.set_field.assert_not_called()

    async def test_blocked_transaction_wins_without_eliciting(self, srv):
        args = {"tcode": "SE16N"}

        ungated, _, _ = await self._call(
            srv, [], "sap_execute_transaction", args, {"execute_transaction": {"ok": 1}},
        )
        gated, elicitor, controller = await self._call(
            srv, ["transactions"], "sap_execute_transaction", args,
            {"execute_transaction": {"ok": 1}},
        )

        assert gated.is_error and ungated.is_error
        assert "blocked by security policy" in _text(gated)
        assert _text(gated) == _text(ungated)
        assert elicitor.messages == []
        # The blocked t-code never reaches the approval dialog.
        assert "SE16N" not in "".join(elicitor.messages)
        controller.execute_transaction.assert_not_called()

    async def test_okcode_bypass_wins_without_eliciting(self, srv):
        """A blocked t-code smuggled through the OK-code field rejects first."""
        args = {"field_id": "wnd[0]/tbar[0]/okcd", "value": "/nSE16N"}

        ungated, _, _ = await self._call(
            srv, [], "sap_set_field", args, {"set_field": {"status": "ok"}},
        )
        gated, elicitor, controller = await self._call(
            srv, ["field_writes"], "sap_set_field", args, {"set_field": {"status": "ok"}},
        )

        assert gated.is_error and ungated.is_error
        assert "command field" in _text(gated)
        assert _text(gated) == _text(ungated)
        assert elicitor.messages == []
        controller.set_field.assert_not_called()

    async def test_okcode_bypass_in_batch_fields_wins_without_eliciting(self, srv):
        args = {"fields": {
            "wnd[0]/usr/txtMATNR": "MAT-001",
            "wnd[0]/tbar[0]/okcd": "SE38",
        }}

        ungated, _, _ = await self._call(
            srv, [], "sap_set_batch_fields", args, {"set_batch_fields": {"set": 0}},
        )
        gated, elicitor, controller = await self._call(
            srv, ["batch_fields"], "sap_set_batch_fields", args,
            {"set_batch_fields": {"set": 0}},
        )

        assert gated.is_error and ungated.is_error
        assert "command field" in _text(gated)
        assert _text(gated) == _text(ungated)
        assert elicitor.messages == []
        controller.set_batch_fields.assert_not_called()


# ===========================================================================
# Audit
# ===========================================================================


def _confirmation_records(caplog):
    records = []
    for record in caplog.records:
        if record.name != "mcp_sap_gui.audit":
            continue
        payload = json.loads(record.message)
        if payload.get("event") == "confirmation":
            records.append(payload)
    return records


class TestConfirmationAudit:
    """The audit log must distinguish accepted from declined confirmations."""

    async def test_accepted_confirmation_is_logged(self, srv, caplog):
        controller, patcher = _patched_controller(set_field={"status": "ok"})
        elicitor = _Elicitor("accept")
        with caplog.at_level(logging.INFO, logger="mcp_sap_gui.audit"), patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["field_writes"]}
                )
                await client.call_tool(
                    "sap_set_field", {"field_id": "wnd[0]/usr/txtX", "value": "1"}
                )

        events = _confirmation_records(caplog)
        assert events == [{
            "event": "confirmation",
            "ts": events[0]["ts"],
            "point": "field_writes",
            "tool": "sap_set_field",
            "outcome": "accepted",
        }]

    async def test_declined_confirmation_is_logged(self, srv, caplog):
        _, patcher = _patched_controller(set_field={"status": "ok"})
        elicitor = _Elicitor("decline")
        with caplog.at_level(logging.INFO, logger="mcp_sap_gui.audit"), patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["field_writes"]}
                )
                await client.call_tool(
                    "sap_set_field", {"field_id": "wnd[0]/usr/txtX", "value": "1"},
                    raise_on_error=False,
                )

        events = _confirmation_records(caplog)
        assert [e["outcome"] for e in events] == ["declined"]
        assert events[0]["tool"] == "sap_set_field"

    async def test_unsupported_client_is_logged(self, srv, caplog):
        _, patcher = _patched_controller(set_field={"status": "ok"})
        with caplog.at_level(logging.INFO, logger="mcp_sap_gui.audit"), patcher:
            async with Client(srv.mcp) as client:
                await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["field_writes"]}
                )
                await client.call_tool(
                    "sap_set_field", {"field_id": "wnd[0]/usr/txtX", "value": "1"},
                    raise_on_error=False,
                )

        events = _confirmation_records(caplog)
        assert [e["outcome"] for e in events] == ["unsupported_client"]

    async def test_relaxation_outcome_is_logged(self, srv, caplog):
        ctx = _make_elicit_ctx()
        await srv.sap_set_confirmation_points(["field_writes"], ctx)
        ctx.elicit = AsyncMock(return_value=MagicMock(action="decline", data=None))

        with caplog.at_level(logging.INFO, logger="mcp_sap_gui.audit"):
            await srv.sap_set_confirmation_points([], ctx)

        events = _confirmation_records(caplog)
        assert events[0]["tool"] == "sap_set_confirmation_points"
        assert events[0]["outcome"] == "declined"
        assert events[0]["point"] == "field_writes"

    async def test_blocked_call_still_produces_a_tool_call_audit_line(self, srv, caplog):
        """Audit must stay outside the gate, else blocked calls vanish."""
        _, patcher = _patched_controller(set_field={"status": "ok"})
        elicitor = _Elicitor("decline")
        with caplog.at_level(logging.INFO, logger="mcp_sap_gui.audit"), patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool(
                    "sap_set_confirmation_points", {"points": ["field_writes"]}
                )
                await client.call_tool(
                    "sap_set_field", {"field_id": "wnd[0]/usr/txtX", "value": "1"},
                    raise_on_error=False,
                )

        tool_calls = [
            json.loads(r.message) for r in caplog.records
            if r.name == "mcp_sap_gui.audit"
        ]
        errored = [
            e for e in tool_calls
            if e.get("event") == "tool_call" and e.get("tool") == "sap_set_field"
        ]
        assert errored and errored[0]["status"] == "error"


# ===========================================================================
# CLI floor
# ===========================================================================


class TestConfirmationCLIFloor:
    """--confirm is an immutable floor, modelled on --read-only."""

    def test_cli_populates_the_floor(self, srv):
        argv = ["mcp-sap-gui", "--confirm", "field_writes", "transactions"]
        with _cli(srv, argv):
            srv.main()
        assert srv.config.confirmation_floor == ["field_writes", "transactions"]

    def test_cli_rejects_unknown_points(self, srv):
        argv = ["mcp-sap-gui", "--confirm", "everything"]
        with _cli(srv, argv), pytest.raises(SystemExit):
            srv.main()

    def test_config_rejects_unknown_floor_points(self, srv):
        with pytest.raises(ValueError, match="Unknown confirmation point"):
            srv.ServerConfig(confirmation_floor=["nope"])

    async def test_floor_applies_to_a_fresh_session(self, srv):
        srv.config = srv.ServerConfig(confirmation_floor=["transactions"])
        controller, patcher = _patched_controller(execute_transaction={"ok": True})
        elicitor = _Elicitor("decline")
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                result = await client.call_tool(
                    "sap_execute_transaction", {"tcode": "MM03"}, raise_on_error=False
                )

        assert result.is_error
        assert "confirmation point 'transactions'" in _text(result)
        controller.execute_transaction.assert_not_called()

    async def test_floor_survives_a_removal_attempt(self, srv):
        srv.config = srv.ServerConfig(confirmation_floor=["transactions"])
        controller, patcher = _patched_controller(execute_transaction={"ok": True})
        elicitor = _Elicitor("accept")
        with patcher:
            async with Client(srv.mcp, elicitation_handler=elicitor.handler) as client:
                await client.call_tool("sap_set_confirmation_points", {"points": []})
                await client.call_tool(
                    "sap_execute_transaction", {"tcode": "MM03"}, raise_on_error=False
                )

        # One prompt only: the gate. Removing a floor point never prompts.
        assert len(elicitor.messages) == 1
        assert "Confirmation point 'transactions' is active" in elicitor.messages[0]
        controller.execute_transaction.assert_called_once()

