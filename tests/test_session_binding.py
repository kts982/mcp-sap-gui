"""Per-chat state across calls, through a real in-memory client.

Most tool tests patch server._ctrl, so none of them would notice the binding
or the confirmation points getting lost between two calls. That is how the
FastMCP 4 port first went wrong: a new ServerSession per request changed the
key on every call, so points set in one call no longer gated the next. These
tests leave _ctrl alone and go through Client(server.mcp), so the key, the
SessionManager and the gate are the real ones.
"""

import pytest
from fastmcp import Client
from fastmcp.client.elicitation import ElicitResult

from mcp_sap_gui import server
from mcp_sap_gui.models import SessionInfo


@pytest.fixture
def default_config(monkeypatch):
    monkeypatch.setattr(server, "config", server.ServerConfig())


def _bindings():
    return list(server._session_mgr._sessions.values())


async def test_second_call_reaches_the_controller_of_the_first(default_config):
    async with Client(server.mcp) as client:
        first = await client.call_tool("sap_get_session_info", {}, raise_on_error=False)
        assert first.is_error  # nothing bound to SAP yet
        [managed] = _bindings()
        managed.controller.get_session_info = lambda: SessionInfo(
            system_name="DEV", system_number="00", client="100", user="TESTUSER",
            language="EN", transaction="SESSION_MANAGER", program="SAPLSMTR_NAVIGATION",
            screen_number=100, session_number=1,
        )

        second = await client.call_tool("sap_get_session_info", {})

        assert second.structured_content["system_name"] == "DEV"
        assert [m.controller for m in _bindings()] == [managed.controller]


async def test_points_set_in_one_call_gate_the_next(default_config):
    asked = []

    async def decline(message, response_type, params, context):
        asked.append(message)
        return ElicitResult(action="decline")

    async with Client(server.mcp, elicitation_handler=decline) as client:
        await client.call_tool("sap_set_confirmation_points", {"points": ["field_writes"]})
        result = await client.call_tool(
            "sap_set_field", {"field_id": "wnd[0]/usr/txtX", "value": "1"},
            raise_on_error=False,
        )

    assert result.is_error
    assert "declined" in result.content[0].text
    assert len(asked) == 1
