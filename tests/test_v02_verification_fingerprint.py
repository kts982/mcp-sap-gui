"""v0.2 verification fingerprint tests.

This file is the regression mirror of the live SAP verification run on
2026-04-07 (see ``.dev/LIVE_VERIFICATION_0.2.local.md``). Each test asserts
the *response shape contract* for one of the 13 live scenarios.

The point is **not** to duplicate the existing per-feature test classes
(``TestWorkflowGuideTool``, ``TestSetBatchFields``, ``TestGetPopupWindow``,
``TestHandlePopup``, ``TestSaveConfirmation``, ``TestAuditMiddleware``,
``TestElementIdValidation``, ``TestTransactionBlocking``,
``TestOkCodeBypassPrevention``, ``TestNormalizeTransaction``,
``TestTransactionGuideTool``). Those tests cover the *internals* of each
feature in depth.

This file covers the *external contract*: a single canary per scenario
that fails loudly if v0.2 surface drift breaks what we live-verified.
Break a test here and you have changed something the live SAP run
depended on.

Coverage matrix (T-IDs match the live verification checklist):

- T02 — workflow guide payload keys
- T03 — transaction guide aliasing → /SCWM/MON
- T04 — element ID normalization (short and full path forms)
- T05 — element ID validation rejects malformed IDs
- T06 — transaction policy allows display transactions
- T07 — transaction policy blocks restricted transactions
- T08 — okcd bypass blocked at the field-write layer
- T09 — batch field validation result shape
- T10 — popup classification fields
- T11 — sap_handle_popup auto refuses risky popups
- T12 — F11/Save elicitation cancellation shape
- T13 — audit middleware writes structured JSON

T01 is covered indirectly: every test in this file imports the server
module and exercises the AuditMiddleware-wrapped tools.
"""

from unittest.mock import AsyncMock, MagicMock, patch

import pytest

import mcp_sap_gui.server as srv_mod
from mcp_sap_gui.session_manager import SessionManager

# ===========================================================================
# Helpers
# ===========================================================================

def _make_mock_ctx():
    return MagicMock()


def _make_elicit_result(action: str, data=None):
    """Build an elicitation result that matches what fastmcp returns."""
    result = MagicMock()
    result.action = action
    result.data = data
    return result


@pytest.fixture
def srv():
    """Fresh server module with a clean SessionManager + default config."""
    srv_mod._session_mgr = SessionManager()
    srv_mod.config = srv_mod.ServerConfig()
    yield srv_mod
    srv_mod._session_mgr.shutdown()


# ===========================================================================
# T02 — workflow guide
# ===========================================================================

class TestT02WorkflowGuideContract:
    """Live scenario T02: sap_get_workflow_guide returns structured payload."""

    @pytest.mark.asyncio
    async def test_search_help_payload_has_required_keys(self, srv):
        result = await srv.sap_get_workflow_guide(
            workflow="search_help",
            target="wnd[0]/usr/ctxtFIELD",
        )
        # The exact key set the live run relied on
        assert set(result.keys()) >= {
            "workflow",
            "target_parameter",
            "target",
            "guide",
        }
        assert result["workflow"] == "search_help"
        assert result["target"] == "wnd[0]/usr/ctxtFIELD"
        assert isinstance(result["guide"], str) and result["guide"]

    @pytest.mark.asyncio
    async def test_table_export_payload_has_required_keys(self, srv):
        result = await srv.sap_get_workflow_guide(
            workflow="table_export",
            target="wnd[0]/usr/cntlGRID/shellcont/shell",
        )
        assert set(result.keys()) >= {
            "workflow",
            "target_parameter",
            "target",
            "guide",
        }
        assert result["workflow"] == "table_export"
        assert result["target_parameter"] == "table_id"


# ===========================================================================
# T03 — transaction guide aliasing
# ===========================================================================

class TestT03TransactionGuideContract:
    """Live scenario T03: aliases resolve to canonical /SCWM/MON."""

    @pytest.mark.asyncio
    @pytest.mark.parametrize(
        "alias",
        [
            "/SCWM/MON",
            "SCWM/MON",
            "warehouse monitor",
            "ewm warehouse monitor",
            "  /SCWM/MON  ",
        ],
    )
    async def test_alias_resolves_to_canonical(self, srv, alias):
        result = await srv.sap_get_transaction_guide(transaction=alias)
        assert result["transaction"] == "/SCWM/MON"
        assert result["mode"] == "read-first"
        assert isinstance(result["guide"], str) and result["guide"]

    @pytest.mark.asyncio
    async def test_payload_includes_task_field(self, srv):
        """The live run passed task='find open warehouse tasks' and got it back."""
        result = await srv.sap_get_transaction_guide(
            transaction="warehouse monitor",
            task="find open warehouse tasks",
        )
        assert result["task"] == "find open warehouse tasks"


# ===========================================================================
# T04 + T05 — ID normalization and validation
# ===========================================================================

class TestT04T05ElementIdContract:
    """Live scenarios T04 (positive) and T05 (negative)."""

    def _controller(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_t04_full_session_path_normalizes_to_short_form(self):
        controller = self._controller()
        normalized = controller._normalize_element_id(
            "/app/con[0]/ses[0]/wnd[0]/tbar[0]/okcd"
        )
        assert normalized == "wnd[0]/tbar[0]/okcd"

    def test_t04_short_form_passes_through_unchanged(self):
        controller = self._controller()
        normalized = controller._normalize_element_id("wnd[0]/tbar[0]/okcd")
        assert normalized == "wnd[0]/tbar[0]/okcd"

    def test_t05_malformed_field_id_rejected_before_findbyid(self):
        controller = self._controller()
        result = controller.read_field("bad-field-id")
        assert "Invalid SAP element ID" in result["error"]
        controller._session.findById.assert_not_called()

    @pytest.mark.asyncio
    async def test_t05_screen_elements_rejects_malformed_container(self, srv):
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        mock_ctrl.get_screen_elements.side_effect = ValueError(
            "Invalid SAP element ID: 'wnd[0]/bad'"
        )
        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            with pytest.raises(ValueError, match="Invalid SAP element ID"):
                await srv.sap_get_screen_elements(ctx, "wnd[0]/bad")


# ===========================================================================
# T06 + T07 — transaction policy
# ===========================================================================

class TestT06T07TransactionPolicyContract:
    """Live scenarios T06 (allowed) and T07 (blocked)."""

    def test_t06_display_transaction_is_allowed(self, srv):
        # MM03 must remain allowed for the live agent flow
        assert srv._is_transaction_blocked("MM03") is False

    @pytest.mark.parametrize("blocked", ["SE80", "SU01", "SE16N", "PFCG"])
    def test_t07_restricted_transactions_are_blocked(self, srv, blocked):
        assert srv._is_transaction_blocked(blocked) is True

    @pytest.mark.asyncio
    async def test_t07_execute_transaction_raises_with_policy_message(self, srv):
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="blocked by security policy"):
            await srv.sap_execute_transaction("SE80", ctx)


# ===========================================================================
# T08 — okcd bypass
# ===========================================================================

class TestT08OkCodeBypassContract:
    """Live scenario T08: okcd bypass blocked before COM."""

    @pytest.mark.asyncio
    async def test_set_field_okcd_blocks_se80(self, srv):
        ctx = _make_mock_ctx()
        with pytest.raises(
            ValueError,
            match="blocked by security policy.*command field",
        ):
            await srv.sap_set_field(
                "wnd[0]/tbar[0]/okcd", "/nSE80", ctx,
            )

    @pytest.mark.asyncio
    async def test_set_field_okcd_blocks_full_path_variant(self, srv):
        """Full /app/con[0]/ses[0]/... form is also covered."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="blocked by security policy"):
            await srv.sap_set_field(
                "/app/con[0]/ses[0]/wnd[0]/tbar[0]/okcd",
                "/nSU01",
                ctx,
            )


# ===========================================================================
# T09 — batch field validation
# ===========================================================================

class TestT09BatchFieldValidationContract:
    """Live scenario T09: validate=true returns full validation payload."""

    @pytest.mark.asyncio
    async def test_validation_payload_shape_matches_live_run(self, srv):
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        # Mirror exactly what the live SAP run returned for T09
        mock_ctrl.set_batch_fields.return_value = {
            "total": 1,
            "succeeded": 1,
            "failed": 0,
            "skipped": 0,
            "results": {"wnd[0]/usr/ctxtRMMG1-MATNR": "success"},
            "validation": {
                "performed": True,
                "message": (
                    "The material TESTMAT_VERIFY does not exist or is not activated"
                ),
                "message_type": "E",
                "screen": {
                    "active_window": "wnd[0]",
                    "transaction": "MM03",
                    "program": "SAPLMGMM",
                    "screen_number": 60,
                    "title": "HA9(1)/200 Display Material (Initial Screen)",
                    "message": (
                        "The material TESTMAT_VERIFY does not exist or is not activated"
                    ),
                    "message_type": "E",
                    "message_id": "M3                  ",
                    "message_number": "305",
                },
                "highlighted_fields": ["wnd[0]/usr/ctxtRMMG1-MATNR"],
            },
        }

        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            result = await srv.sap_set_batch_fields(
                {"wnd[0]/usr/ctxtRMMG1-MATNR": "TESTMAT_VERIFY"},
                ctx,
                validate=True,
                skip_readonly=True,
            )

        # The exact set of top-level keys the live run depended on
        assert set(result.keys()) >= {
            "total", "succeeded", "failed", "skipped", "results", "validation",
        }
        validation = result["validation"]
        assert validation["performed"] is True
        assert validation["message_type"] == "E"
        assert "screen" in validation
        assert "highlighted_fields" in validation
        # The screen sub-dict is what unblocks the agent's next decision
        assert "active_window" in validation["screen"]
        assert "transaction" in validation["screen"]
        assert "message_type" in validation["screen"]


# ===========================================================================
# T10 — popup classification
# ===========================================================================

class TestT10PopupClassificationContract:
    """Live scenario T10: popup payload includes classification keys."""

    def test_popup_input_required_payload_keys(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)

        # Minimal popup with one editable field — matches the F4 search
        # help dialog the live run encountered (Restrict Value Range)
        mock_popup = MagicMock()
        mock_popup.Text = "HA9(1)/200 Restrict Value Range"

        mock_usr = MagicMock()
        field = MagicMock()
        field.Type = "GuiTextField"
        field.Id = "wnd[1]/usr/txtMAXRECORDS"
        field.Name = "MAXRECORDS"
        field.Text = "500"
        field.Changeable = True
        field.Children = MagicMock()
        field.Children.Count = 0

        mock_usr.Children.Count = 1
        mock_usr.Children.side_effect = lambda i: field

        def find_by_id(elem_id):
            if elem_id == "wnd[1]":
                return mock_popup
            if elem_id == "wnd[1]/usr":
                return mock_usr
            if elem_id == "wnd[1]/sbar":
                raise Exception("no sbar")
            if elem_id.startswith("wnd[1]/tbar"):
                raise Exception("no toolbar")
            raise Exception("not found")

        controller._session.findById.side_effect = find_by_id

        result = controller.get_popup_window()

        # The exact contract the live verification reported
        assert result["popup_exists"] is True
        assert "classification" in result
        assert "recommended_action" in result
        assert "has_inputs" in result
        assert "interactive_elements" in result
        assert result["classification"] == "input_required"
        assert result["has_inputs"] is True
        assert len(result["interactive_elements"]) >= 1


# ===========================================================================
# T11 — safe popup auto handling
# ===========================================================================

class TestT11HandlePopupAutoContract:
    """Live scenario T11: action='auto' refuses to act on input_required popups."""

    def _controller_with_input_popup(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)

        mock_popup = MagicMock()
        mock_popup.Text = "Restrict Value Range"

        mock_usr = MagicMock()
        field = MagicMock()
        field.Type = "GuiTextField"
        field.Id = "wnd[1]/usr/txtMAXRECORDS"
        field.Name = "MAXRECORDS"
        field.Text = "500"
        field.Changeable = True
        field.Children = MagicMock()
        field.Children.Count = 0
        mock_usr.Children.Count = 1
        mock_usr.Children.side_effect = lambda i: field

        def find_by_id(elem_id):
            if elem_id == "wnd[1]":
                return mock_popup
            if elem_id == "wnd[1]/usr":
                return mock_usr
            if elem_id.startswith("wnd[1]/tbar"):
                raise Exception("no toolbar")
            if elem_id == "wnd[1]/sbar":
                raise Exception("no sbar")
            raise Exception("not found")

        controller._session.findById.side_effect = find_by_id
        return controller, mock_popup

    def test_auto_decides_read_for_input_required_popup(self):
        controller, popup = self._controller_with_input_popup()

        result = controller.handle_popup("auto")

        assert result["action_requested"] == "auto"
        assert result["auto_decision"] == "read"
        assert result["action"] == "read"
        # The contract: auto MUST NOT press anything on a risky popup
        popup.sendVKey.assert_not_called()


# ===========================================================================
# T12 — F11/Save elicitation
# ===========================================================================

class TestT12SaveElicitationContract:
    """Live scenario T12: declining F11 returns a structured cancellation."""

    @pytest.mark.asyncio
    async def test_decline_returns_cancellation_payload(self, srv):
        ctx = _make_mock_ctx()
        ctx.elicit = AsyncMock(return_value=_make_elicit_result("decline"))
        mock_ctrl = MagicMock()

        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            result = await srv.sap_send_key("F11", ctx)

        # The exact payload the live run received and depended on
        assert result["status"] == "cancelled"
        assert result["action"] == "decline"
        assert result["key"] == "F11"
        assert "reason" in result
        # Critical: the controller's send_vkey was never called
        mock_ctrl.send_vkey.assert_not_called()

    @pytest.mark.asyncio
    async def test_save_alias_also_elicits(self, srv):
        """The 'Save' string alias must trigger the same elicitation flow."""
        ctx = _make_mock_ctx()
        ctx.elicit = AsyncMock(return_value=_make_elicit_result("decline"))
        mock_ctrl = MagicMock()

        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            result = await srv.sap_send_key("Save", ctx)

        ctx.elicit.assert_called_once()
        assert result["status"] == "cancelled"
        mock_ctrl.send_vkey.assert_not_called()


# ===========================================================================
# T13 — audit middleware
# ===========================================================================

class TestT13AuditContract:
    """Live scenario T13: every tool call writes a structured JSON line."""

    def test_middleware_is_registered_on_server(self, srv):
        from mcp_sap_gui.audit import AuditMiddleware
        assert any(
            isinstance(m, AuditMiddleware) for m in srv.mcp.middleware
        )

    def test_middleware_emits_required_fields(self, srv, caplog):
        """The JSONL line shape must match what the live run wrote."""
        import asyncio
        import json
        import logging
        from datetime import datetime, timezone

        from fastmcp.server.middleware import MiddlewareContext

        from mcp_sap_gui.audit import AuditMiddleware

        with caplog.at_level(logging.INFO, logger="mcp_sap_gui.audit"):
            mw = AuditMiddleware()
            params = MagicMock()
            params.name = "sap_read_field"
            params.arguments = {"field_id": "wnd[0]/tbar[0]/okcd"}
            ctx = MiddlewareContext(
                message=params,
                timestamp=datetime.now(timezone.utc),
                method="tools/call",
            )

            async def fake_next(ctx):
                result = MagicMock()
                result.content = []
                return result

            asyncio.new_event_loop().run_until_complete(
                mw.on_call_tool(ctx, fake_next)
            )

        records = [r for r in caplog.records if r.name == "mcp_sap_gui.audit"]
        assert len(records) == 1
        log = json.loads(records[0].message)
        # The exact field set the live run's JSONL contained
        assert set(log.keys()) >= {
            "event", "ts", "tool", "args", "status", "duration_ms",
        }
        assert log["event"] == "tool_call"
        assert log["tool"] == "sap_read_field"
        assert log["status"] == "ok"

    def test_middleware_does_not_log_passwords(self, srv, caplog):
        """No live JSONL line may contain a password substring."""
        import asyncio
        import json
        import logging
        from datetime import datetime, timezone

        from fastmcp.server.middleware import MiddlewareContext

        from mcp_sap_gui.audit import AuditMiddleware

        with caplog.at_level(logging.INFO, logger="mcp_sap_gui.audit"):
            mw = AuditMiddleware()
            params = MagicMock()
            params.name = "sap_connect"
            params.arguments = {
                "system_description": "HA9",
                "password": "do-not-log-me",
                "user": "KTSIO",
            }
            ctx = MiddlewareContext(
                message=params,
                timestamp=datetime.now(timezone.utc),
                method="tools/call",
            )

            async def fake_next(ctx):
                result = MagicMock()
                result.content = []
                return result

            asyncio.new_event_loop().run_until_complete(
                mw.on_call_tool(ctx, fake_next)
            )

        records = [r for r in caplog.records if r.name == "mcp_sap_gui.audit"]
        assert len(records) == 1
        # The full serialized JSONL line must not contain the password
        assert "do-not-log-me" not in records[0].message
        log = json.loads(records[0].message)
        assert log["args"]["password"] == "***"
        assert log["args"]["user"] == "KTSIO"
