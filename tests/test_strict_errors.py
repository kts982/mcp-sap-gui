"""--strict-errors, the audit status, and the Save key contract.

Every call goes through the real server module and its middleware chain
(fastmcp's in-memory client), with a fake controller in place of SAP GUI.
"""

import asyncio
import json
import logging

import pytest
from fastmcp import Client
from fastmcp.client.elicitation import ElicitResult

from mcp_sap_gui import server
from mcp_sap_gui.strict import result_error, strict_errors_from_env

FIELD = "wnd[0]/usr/txtTHOLT-KTEXT"
REFUSED = {"field_id": FIELD, "error": "Could not set field"}


class FakeController:
    """Stands in for SAPGUIController; returns canned results."""

    def __init__(self):
        self.set_field_result = REFUSED
        self.batch_result = {"total": 1, "succeeded": 1, "failed": 0,
                             "skipped": 0, "results": {}}
        self.sent = []

    def set_field(self, field_id, value):
        return dict(self.set_field_result)

    def set_batch_fields(self, fields, **kwargs):
        return dict(self.batch_result)

    def read_field(self, field_id):
        return {"field_id": field_id, "value": "x", "type": "GuiTextField"}

    def check_vkey(self, vkey, window="wnd[0]"):
        return None

    def send_vkey(self, vkey, window="wnd[0]"):
        self.sent.append(vkey)
        return {"vkey": vkey, "screen": {"active_window": "wnd[0]"}}


@pytest.fixture
def fake(monkeypatch):
    controller = FakeController()
    monkeypatch.setattr(server, "_ctrl", lambda ctx: controller)
    monkeypatch.setattr(server, "config", server.ServerConfig())
    return controller


def strict(monkeypatch, on=True):
    monkeypatch.setattr(server, "config", server.ServerConfig(strict_errors=on))


def call(tool, args, elicitation_handler=None):
    async def run():
        async with Client(server.mcp,
                          elicitation_handler=elicitation_handler) as client:
            return await client.call_tool(tool, args, raise_on_error=False)
    return asyncio.run(run())


def audit_lines(caplog, event="tool_call"):
    lines = []
    for record in caplog.records:
        if record.name != "mcp_sap_gui.audit":
            continue
        try:
            line = json.loads(record.getMessage())
        except ValueError:
            continue
        if line.get("event") == event:
            lines.append(line)
    return lines


# --- result_error ----------------------------------------------------------

@pytest.mark.parametrize("tool, structured, expected", [
    ("sap_set_field", {"error": "Could not set field"}, "Could not set field"),
    ("sap_set_field", {"status": "success", "error": ""}, None),
    ("sap_set_field", {"status": "success"}, None),
    ("sap_read_field", None, None),
    # A screen's status-bar text is not the call's failure.
    ("sap_press_button", {"screen": {"error": "x"}}, None),
    ("sap_set_batch_fields", {"failed": 2, "succeeded": 1},
     "2 field(s) were not set"),
    ("sap_set_batch_fields", {"failed": 0, "skipped": 3}, None),
    ("sap_set_batch_fields",
     {"failed": 0, "validation": {"performed": False, "error": "boom"}},
     "validation Enter failed: boom"),
    ("sap_set_batch_fields",
     {"failed": 0, "validation": {"performed": False,
                                  "reason": "no fields were set"}}, None),
    # failed counts only for the batch tool.
    ("sap_read_table", {"failed": 2}, None),
])
def test_result_error(tool, structured, expected):
    assert result_error(tool, structured) == expected


@pytest.mark.parametrize("value, expected", [
    (None, False), ("", False), ("0", False), ("no", False),
    ("1", True), ("true", True), ("YES", True), (" on ", True),
])
def test_env_flag(value, expected):
    env = {} if value is None else {"MCP_SAP_GUI_STRICT_ERRORS": value}
    assert strict_errors_from_env(env) is expected


def test_strict_is_off_by_default():
    assert server.ServerConfig().strict_errors is False


# --- the middleware ----------------------------------------------------------

def test_default_mode_returns_failure_as_normal_result(fake):
    result = call("sap_set_field", {"field_id": FIELD, "value": "x"})

    assert result.is_error is False
    assert result.structured_content == REFUSED


def test_strict_mode_returns_failure_as_mcp_error(fake, monkeypatch):
    strict(monkeypatch)

    result = call("sap_set_field", {"field_id": FIELD, "value": "x"})

    assert result.is_error is True
    # The whole payload survives, not only a line of text.
    assert result.structured_content == REFUSED
    assert json.loads(result.content[0].text) == REFUSED


def test_strict_mode_leaves_success_alone(fake, monkeypatch):
    strict(monkeypatch)
    fake.set_field_result = {"field_id": FIELD, "value": "x",
                             "status": "success"}

    result = call("sap_set_field", {"field_id": FIELD, "value": "x"})

    assert result.is_error is False
    assert result.structured_content["status"] == "success"


def test_strict_mode_flags_partial_batch(fake, monkeypatch):
    strict(monkeypatch)
    fake.batch_result = {"total": 2, "succeeded": 1, "failed": 1, "skipped": 0,
                         "results": {FIELD: "error: Could not set field"}}

    result = call("sap_set_batch_fields", {"fields": {FIELD: "x", "b": "y"}})

    assert result.is_error is True
    assert result.structured_content["failed"] == 1


def test_strict_mode_off_again(fake, monkeypatch):
    strict(monkeypatch, on=False)
    assert call("sap_set_field", {"field_id": FIELD, "value": "x"}).is_error is False


# --- audit status ------------------------------------------------------------

@pytest.mark.parametrize("strict_on", [False, True])
def test_audit_records_failed_result_as_error(fake, monkeypatch, caplog,
                                             strict_on):
    strict(monkeypatch, on=strict_on)
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")

    call("sap_set_field", {"field_id": FIELD, "value": "x"})

    (line,) = audit_lines(caplog)
    assert line["tool"] == "sap_set_field"
    assert line["status"] == "error"
    assert line["error_msg"] == "Could not set field"


def test_audit_records_success_as_ok(fake, caplog):
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")

    call("sap_read_field", {"field_id": FIELD})

    (line,) = audit_lines(caplog)
    assert line["status"] == "ok"
    assert "error_msg" not in line


def test_audit_still_records_raised_errors(fake, monkeypatch, caplog):
    monkeypatch.setattr(server, "config", server.ServerConfig(read_only=True))
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")

    result = call("sap_set_field", {"field_id": FIELD, "value": "x"})

    assert result.is_error is True
    (line,) = audit_lines(caplog)
    assert line["status"] == "error"
    assert line["error"] == "ToolError"


# --- sap_send_key Save -------------------------------------------------------

def test_declined_save_carries_error(fake, caplog):
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")

    async def handler(message, response_type, params, context):
        return ElicitResult(action="decline")

    result = call("sap_send_key", {"key": "Save"}, handler)

    sc = result.structured_content
    assert fake.sent == []
    # Existing keys unchanged; error added.
    assert sc["status"] == "cancelled"
    assert sc["action"] == "decline"
    assert sc["reason"] == "User did not confirm save (decline)"
    assert "declined by user" in sc["error"]
    (line,) = audit_lines(caplog)
    assert line["status"] == "error"


def test_declined_save_in_strict_mode_is_an_mcp_error(fake, monkeypatch):
    strict(monkeypatch)

    async def handler(message, response_type, params, context):
        return ElicitResult(action="cancel")

    result = call("sap_send_key", {"key": "F11"}, handler)

    assert result.is_error is True
    assert result.structured_content["status"] == "cancelled"


def test_accepted_save_is_sent_and_audited_ok(fake, monkeypatch, caplog):
    strict(monkeypatch)
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")

    async def handler(message, response_type, params, context):
        return True

    result = call("sap_send_key", {"key": "F11"}, handler)

    assert fake.sent == [11]
    assert result.is_error is False
    assert result.structured_content["confirmed_by_user"] is True
    (line,) = audit_lines(caplog)
    assert line["status"] == "ok"


def test_declined_popup_answer_carries_error():
    class DeclineCtx:
        async def elicit(self, message, response_type=None, **_):
            return ElicitResult(action="decline")

    plan = {"popup": {"window_id": "wnd[1]", "title": "Exit Editor",
                      "texts": ["Save changes?"]},
            "button": {"text": "Yes"}}

    result = asyncio.run(server._confirm_popup_save(DeclineCtx(), plan))

    assert result["status"] == "cancelled"
    assert result["action"] == "none"
    assert "declined by user" in result["error"]


def test_confirmed_write_that_failed_keeps_confirmed_by_user(fake, monkeypatch):
    # Strict mode sits outside the confirmation gate: the gate marks the
    # result before strict mode flags it, so both survive.
    monkeypatch.setattr(server, "config", server.ServerConfig(
        strict_errors=True, confirmation_floor=["field_writes"]))

    async def handler(message, response_type, params, context):
        return True

    result = call("sap_set_field", {"field_id": FIELD, "value": "x"}, handler)

    assert result.is_error is True
    assert result.structured_content["error"] == "Could not set field"
    assert result.structured_content["confirmed_by_user"] is True
