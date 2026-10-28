"""The SAP system, client and user on every tool_call audit line."""

import json
import logging
from datetime import datetime, timezone
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest
from fastmcp import Client
from fastmcp.server.middleware import MiddlewareContext
from mcp.types import Implementation

from mcp_sap_gui import server
from mcp_sap_gui.audit import AuditMiddleware
from mcp_sap_gui.sap_controller import SAPGUIController

DEV = {"system": "DEV", "client": "100", "user": "TESTUSER"}


def _info(system="DEV", client="100", user="TESTUSER"):
    return SimpleNamespace(
        SystemName=system, SystemNumber="00", Client=client, User=user,
        Language="EN", Transaction="SESSION_MANAGER", Program="SAPLSMTR_NAVIGATION",
        ScreenNumber=100, SessionNumber=1, ScriptingModeReadOnly=False,
    )


def _audit_lines(caplog):
    return [json.loads(r.message) for r in caplog.records
            if r.name == "mcp_sap_gui.audit"
            and json.loads(r.message).get("event") == "tool_call"]


async def _run(mw, call_next, tool="sap_read_field"):
    params = MagicMock()
    params.name = tool
    params.arguments = {}
    context = MiddlewareContext(
        message=params, timestamp=datetime.now(timezone.utc), method="tools/call",
    )
    return await mw.on_call_tool(context, call_next)


async def _ok(_context):
    return MagicMock(content=[], structured_content=None, is_error=False)


# ---------------------------------------------------------------------------
# The middleware
# ---------------------------------------------------------------------------


async def test_line_names_the_bound_system(caplog):
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")
    await _run(AuditMiddleware(identity=lambda ctx: DEV), _ok)
    assert _audit_lines(caplog)[0]["sap"] == DEV


async def test_no_sap_key_while_nothing_is_bound(caplog):
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")
    await _run(AuditMiddleware(identity=lambda ctx: {}), _ok)
    await _run(AuditMiddleware(), _ok)
    assert all("sap" not in line for line in _audit_lines(caplog))


async def test_attach_logs_the_system_it_attached(caplog):
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")
    bound = {}

    async def attach(_context):
        bound.update(DEV)
        return await _ok(_context)

    await _run(AuditMiddleware(identity=lambda ctx: dict(bound)), attach,
               tool="sap_connect_existing")
    assert _audit_lines(caplog)[0]["sap"] == DEV


async def test_disconnect_logs_the_system_it_left(caplog):
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")
    bound = dict(DEV)

    async def disconnect(_context):
        bound.clear()
        return await _ok(_context)

    await _run(AuditMiddleware(identity=lambda ctx: dict(bound)), disconnect,
               tool="sap_disconnect")
    assert _audit_lines(caplog)[0]["sap"] == DEV


async def test_failed_call_keeps_the_system(caplog):
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")

    async def fail(_context):
        raise ValueError("Element not found")

    with pytest.raises(ValueError):
        await _run(AuditMiddleware(identity=lambda ctx: DEV), fail)
    line = _audit_lines(caplog)[0]
    assert line["status"] == "error" and line["sap"] == DEV


async def test_identity_that_raises_never_fails_the_call(caplog):
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")

    def broken(ctx):
        raise RuntimeError("no binding")

    result = await _run(AuditMiddleware(identity=broken), _ok)
    assert result.is_error is False
    assert "sap" not in _audit_lines(caplog)[0]


# ---------------------------------------------------------------------------
# The controller's cache
# ---------------------------------------------------------------------------


def _bound_controller(info):
    controller = SAPGUIController()
    controller._application = SimpleNamespace(MajorVersion=8100, Patchlevel=0)
    controller._session = SimpleNamespace(Info=info, Busy=False)
    return controller


def test_bind_reads_system_client_user():
    controller = _bound_controller(_info())
    controller._read_session_traits()
    assert controller.sap_identity == DEV


def test_logon_screen_has_no_user_until_session_info_refreshes_it():
    info = _info(user="")
    controller = _bound_controller(info)
    controller._read_session_traits()
    assert controller.sap_identity == {"system": "DEV", "client": "100"}

    info.User = "TESTUSER"  # the user logged on in the sap_connect window
    controller.get_session_info()
    assert controller.sap_identity == DEV


def test_disconnect_forgets_it():
    controller = _bound_controller(_info())
    controller._read_session_traits()
    controller.disconnect()
    assert controller.sap_identity == {}


# ---------------------------------------------------------------------------
# Through the server
# ---------------------------------------------------------------------------


async def test_server_writes_the_bound_system_into_each_line(monkeypatch, caplog):
    monkeypatch.setattr(server, "config", server.ServerConfig())
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")

    async with Client(server.mcp) as client:
        await client.call_tool("sap_get_session_info", {}, raise_on_error=False)
        [managed] = server._session_mgr._sessions.values()
        managed.controller._session = SimpleNamespace(Info=_info(), Busy=False)
        await client.call_tool("sap_get_session_info", {})
        await client.call_tool("sap_get_session_info", {})

    first, second, third = _audit_lines(caplog)
    assert "sap" not in first  # nothing bound yet
    assert second["sap"] == DEV  # get_session_info cached it during the call
    assert third["sap"] == DEV


# ---------------------------------------------------------------------------
# The MCP client
# ---------------------------------------------------------------------------


async def test_line_names_the_mcp_client(monkeypatch, caplog):
    monkeypatch.setattr(server, "config", server.ServerConfig())
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")
    info = Implementation(name="test-client", version="1.2.3")

    async with Client(server.mcp, client_info=info) as client:
        await client.call_tool("sap_get_session_info", {}, raise_on_error=False)

    [line] = _audit_lines(caplog)
    assert line["mcp_client"] == "test-client 1.2.3"
    assert line["protocol"]


async def test_no_mcp_client_key_without_a_session(caplog):
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")
    await _run(AuditMiddleware(), _ok)
    assert "mcp_client" not in _audit_lines(caplog)[0]
