"""Server side of session ownership: closing is the destructive tier."""

import asyncio
from types import SimpleNamespace

import pytest

from mcp_sap_gui import server
from mcp_sap_gui.session_manager import ManagedSession, SessionManager


class FakeController:
    def __init__(self, owned: int) -> None:
        self.is_connected = True
        self._owns_session = owned > 0
        self._owned_sessions = [object()] * owned
        self.scripting_read_only = False
        self.calls: list[bool] = []

    def disconnect(self, close_owned: bool = True) -> list:
        self.calls.append(close_owned)
        return ["/app/con[1]/ses[0]"] if close_owned and self._owned_sessions else []


@pytest.fixture
def bound(monkeypatch):
    """A server with one MCP session bound to an owned SAP session."""
    manager = SessionManager()
    monkeypatch.setattr(server, "_session_mgr", manager)
    monkeypatch.setattr(server, "config", server.ServerConfig())
    ctx = SimpleNamespace(session=object())
    controller = FakeController(owned=1)
    manager._sessions[server._session_key(ctx)] = ManagedSession(controller=controller)
    yield SimpleNamespace(manager=manager, ctx=ctx, controller=controller)
    manager.shutdown()


def _tool(name: str):
    return asyncio.run(server.mcp.get_tool(name))


def test_close_tool_is_the_destructive_tier():
    tool = _tool("sap_close_session")
    assert tool.tags == {"destructive"}
    assert tool.annotations.destructiveHint is True
    assert tool.annotations.readOnlyHint is False


def test_disconnect_stays_visible_everywhere_but_is_not_read_only():
    tool = _tool("sap_disconnect")
    assert tool.tags == {"read"}
    assert tool.annotations.readOnlyHint is False


def test_connect_is_not_read_only():
    tool = _tool("sap_connect")
    assert tool.tags == {"read"}
    assert tool.annotations.readOnlyHint is False
    assert tool.annotations.destructiveHint is False


def test_full_profile_disconnect_closes_as_before(bound):
    result = asyncio.run(server.sap_disconnect(bound.ctx))
    assert result["action"] == "closed"
    assert bound.controller.calls == [True]


@pytest.mark.parametrize("floor", ["exploration", "operator"])
def test_floor_below_full_disconnect_only_detaches(bound, monkeypatch, floor):
    monkeypatch.setattr(server, "config", server.ServerConfig(profile=floor))
    result = asyncio.run(server.sap_disconnect(bound.ctx))
    assert result["action"] == "detached"
    assert result["owned_sessions_left_open"] == 1
    assert "full" in result["note"]
    assert bound.controller.calls == [False]


def test_session_profile_below_full_disconnect_only_detaches(bound):
    bound.manager.set_profile(server._session_key(bound.ctx), "operator")
    asyncio.run(server.sap_disconnect(bound.ctx))
    assert bound.controller.calls == [False]


def test_close_session_refuses_below_full(bound):
    bound.manager.set_profile(server._session_key(bound.ctx), "exploration")
    with pytest.raises(ValueError, match="full"):
        asyncio.run(server.sap_close_session(bound.ctx))
    assert bound.controller.calls == []


def test_close_session_closes_at_full(bound):
    result = asyncio.run(server.sap_close_session(bound.ctx))
    assert result["closed_sessions"] == ["/app/con[1]/ses[0]"]


def test_exit_closes_at_full_by_default(bound):
    assert server._may_close_on_exit(server._session_key(bound.ctx)) is True


def test_exit_detaches_with_detach_on_exit(bound, monkeypatch):
    monkeypatch.setattr(server, "config", server.ServerConfig(detach_on_exit=True))
    assert server._may_close_on_exit(server._session_key(bound.ctx)) is False


def test_exit_detaches_below_full(bound, monkeypatch):
    monkeypatch.setattr(server, "config", server.ServerConfig(profile="operator"))
    bound.manager.release_all(may_close=server._may_close_on_exit)
    assert bound.controller.calls == [False]
