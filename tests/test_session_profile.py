"""A session profile holds until the session changes it, through a real client.

sap_set_policy_profile hides tools with FastMCP's session visibility, which
FastMCP keeps in session state with a 24-hour TTL. When it lapsed, the hidden
write tools were listed and callable again. SessionProfileMiddleware refuses
them at call time from the SessionManager's copy, which does not expire.
"""

import asyncio

import pytest
from fastmcp import Client
from fastmcp.server.context import Context

from mcp_sap_gui import server

SET_FIELD = {"field_id": "wnd[0]/usr/txtX", "value": "1"}
REFUSED = "outside this session's 'exploration' profile"


@pytest.fixture(autouse=True)
def default_config(monkeypatch):
    monkeypatch.setattr(server, "config", server.ServerConfig())


@pytest.fixture
def short_state_ttl(monkeypatch):
    # Stands in for FastMCP's 24 hours: the visibility rules expire after 1 s.
    monkeypatch.setattr(Context, "_STATE_TTL_SECONDS", 1)


async def _error_text(client, name, args):
    result = await client.call_tool(name, args, raise_on_error=False)
    assert result.is_error
    return result.content[0].text


async def test_exploration_refuses_a_write_tool():
    async with Client(server.mcp) as client:
        await client.call_tool("sap_set_policy_profile", {"profile": "exploration"})
        # While FastMCP's hiding holds, it refuses first; ours is the backstop.
        text = await _error_text(client, "sap_set_field", SET_FIELD)
        assert REFUSED in text or "Unknown tool" in text
        # A read tool passes the profile and reaches its body (no SAP here).
        text = await _error_text(client, "sap_get_session_info", {})
        assert "profile" not in text


async def test_profile_still_refuses_after_the_state_ttl(short_state_ttl):
    async with Client(server.mcp) as client:
        await client.call_tool("sap_set_policy_profile", {"profile": "exploration"})
        await asyncio.sleep(1.5)
        listed = {tool.name for tool in await client.list_tools()}
        assert "sap_set_field" in listed  # FastMCP's hiding has lapsed ...
        assert REFUSED in await _error_text(client, "sap_set_field", SET_FIELD)


async def test_switching_back_to_full_lets_writes_through():
    async with Client(server.mcp) as client:
        await client.call_tool("sap_set_policy_profile", {"profile": "exploration"})
        await client.call_tool("sap_set_policy_profile", {"profile": "full"})
        text = await _error_text(client, "sap_set_field", SET_FIELD)
        assert "profile" not in text  # reached the tool body: not connected


async def test_a_profile_belongs_to_the_session_that_set_it():
    async with Client(server.mcp) as restricted, Client(server.mcp) as other:
        await restricted.call_tool("sap_set_policy_profile", {"profile": "exploration"})
        text = await _error_text(other, "sap_set_field", SET_FIELD)
        assert "profile" not in text


async def test_the_floor_caps_what_is_recorded(monkeypatch):
    monkeypatch.setattr(server, "config", server.ServerConfig(profile="exploration"))
    async with Client(server.mcp) as client:
        result = await client.call_tool("sap_set_policy_profile", {"profile": "full"})
        assert result.data["profile"] == "exploration"
