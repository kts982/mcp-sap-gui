"""A handshake-era session that ends takes its binding and points with it.

The hook rides on the MCP SDK's connection exit stack; on 2026-07-28 that
stack closes after every request, so nothing may be registered there.
"""

import json
import time
from contextlib import AsyncExitStack
from types import SimpleNamespace

import pytest
from starlette.testclient import TestClient

from mcp_sap_gui import server
from mcp_sap_gui.session_manager import SessionManager


def _ctx(protocol_version, exit_stack):
    return SimpleNamespace(
        request_context=SimpleNamespace(protocol_version=protocol_version),
        session=SimpleNamespace(_connection=SimpleNamespace(exit_stack=exit_stack)),
    )


@pytest.fixture
def manager(monkeypatch):
    mgr = SessionManager()
    monkeypatch.setattr(server, "_session_mgr", mgr)
    yield mgr
    mgr.shutdown()


async def test_closing_the_stack_forgets_binding_and_points(manager):
    stack = AsyncExitStack()
    ctx = _ctx("2025-11-25", stack)
    manager.get_or_create("s1")
    manager.set_confirmation_points("s1", {"field_writes"})
    server._forget_when_session_ends(ctx, "s1")
    server._forget_when_session_ends(ctx, "s1")  # registered once
    assert len(stack._exit_callbacks) == 1

    await stack.aclose()

    assert "s1" not in manager._sessions
    assert manager.get_confirmation_points("s1") == set()
    assert "s1" not in server._watched_session_keys


async def test_nothing_is_registered_on_2026_07_28(manager):
    stack = AsyncExitStack()
    server._forget_when_session_ends(_ctx("2026-07-28", stack), "process")
    assert len(stack._exit_callbacks) == 0


async def test_without_the_private_connection_nothing_is_registered(manager):
    ctx = SimpleNamespace(request_context=SimpleNamespace(protocol_version="2025-11-25"),
                          session=SimpleNamespace())
    server._forget_when_session_ends(ctx, "s2")  # no error, nothing watched
    assert "s2" not in server._watched_session_keys


HEADERS = {"Content-Type": "application/json",
           "Accept": "application/json, text/event-stream"}


def _rpc(client, method, params=None, session=None, id_=1):
    message = {"jsonrpc": "2.0", "method": method, "params": params or {}}
    if id_ is not None:
        message["id"] = id_
    headers = dict(HEADERS, **({"mcp-session-id": session} if session else {}))
    return client.post("/mcp", content=json.dumps(message), headers=headers)


@pytest.mark.parametrize("ending", ["client ends it", "idle expiry"])
def test_http_session_that_ends_drops_its_binding_and_points(ending):
    # 1 s stands in for the SDK's 30-minute idle timeout.
    app = server.mcp.http_app(session_idle_timeout=1 if ending == "idle expiry" else None)
    with TestClient(app, base_url="http://127.0.0.1:8000") as client:
        init = _rpc(client, "initialize", {
            "protocolVersion": "2025-06-18", "capabilities": {},
            "clientInfo": {"name": "test", "version": "0"}})
        session = init.headers["mcp-session-id"]
        _rpc(client, "notifications/initialized", session=session, id_=None)
        _rpc(client, "tools/call", {"name": "sap_set_confirmation_points",
                                    "arguments": {"points": ["field_writes"]}},
             session=session, id_=2)
        _rpc(client, "tools/call", {"name": "sap_get_session_info", "arguments": {}},
             session=session, id_=3)
        mgr = server._session_mgr
        assert session in mgr._sessions
        assert mgr.get_confirmation_points(session) == {"field_writes"}

        if ending == "client ends it":
            client.delete("/mcp", headers={"mcp-session-id": session})
        deadline = time.monotonic() + 5
        while session in mgr._sessions and time.monotonic() < deadline:
            time.sleep(0.05)

        assert session not in mgr._sessions
        assert mgr.get_confirmation_points(session) == set()
