"""A dialog answered faster than a person can answer one.

Live on 2026-10-19 the Claude Code desktop app answered every elicitation
with "decline" in 7-13 ms and showed nothing. The decline still counts (fail
closed), but the error says that no person saw the request, and the audit
line has elapsed_ms and no_dialog. conftest sets the threshold to 0 for every
other test; these tests use the real one.
"""

import asyncio
import json
import logging
from types import SimpleNamespace
from typing import Any, List

import pytest
from fastmcp import Client
from fastmcp.client.elicitation import ElicitResult

from mcp_sap_gui import confirmation, server
from mcp_sap_gui.confirmation import no_dialog_note

REAL_THRESHOLD = 250.0


@pytest.fixture
def real_threshold(monkeypatch):
    monkeypatch.setattr(confirmation, "NO_DIALOG_MS", REAL_THRESHOLD)


class FakeCtx:
    """Answers at once, like the desktop app did, or after *delay* seconds."""

    def __init__(self, action: str = "decline", delay: float = 0.0,
                 data: Any = None) -> None:
        self.action = action
        self.delay = delay
        self.data = data
        self.prompts: List[str] = []

    async def elicit(self, message: str, response_type: Any = None, **_: Any):
        self.prompts.append(message)
        if self.delay:
            await asyncio.sleep(self.delay)
        return SimpleNamespace(action=self.action, data=self.data)


def _confirmations(caplog):
    out = []
    for record in caplog.records:
        try:
            data = json.loads(record.getMessage())
        except ValueError:
            continue
        if data.get("event") == "confirmation":
            out.append(data)
    return out


def test_note_only_below_the_threshold(real_threshold):
    assert no_dialog_note(REAL_THRESHOLD) == ""
    assert no_dialog_note(5000.0) == ""
    note = no_dialog_note(8.4)
    assert "answered in 8 ms without showing a dialog" in note
    assert "no person saw this request" in note


def test_fast_save_decline_says_no_person_saw_it(real_threshold, caplog):
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")

    declined = asyncio.run(server._confirm_save(FakeCtx("decline"), "F11"))

    assert declined["status"] == "cancelled"
    assert declined["error"].startswith(
        "Save declined by user (decline). Nothing was sent. The client answered in"
    )
    assert "without showing a dialog" in declined["error"]
    [event] = _confirmations(caplog)
    assert event["outcome"] == "declined"
    assert event["no_dialog"] is True
    assert event["elapsed_ms"] < REAL_THRESHOLD


def test_slow_save_decline_keeps_the_old_wording(monkeypatch, caplog):
    # A person who takes their time: a 30 ms threshold and a 50 ms answer.
    monkeypatch.setattr(confirmation, "NO_DIALOG_MS", 30.0)
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")

    declined = asyncio.run(
        server._confirm_save(FakeCtx("decline", delay=0.05), "F11"))

    assert declined["error"] == "Save declined by user (decline). Nothing was sent."
    [event] = _confirmations(caplog)
    assert "no_dialog" not in event
    assert event["elapsed_ms"] >= 30.0


def test_fast_popup_decline_says_so(real_threshold):
    plan = {"popup": {"title": "Exit Editor", "texts": ["Save changes?"]},
            "button": {"text": "Yes"}}

    declined = asyncio.run(server._confirm_popup_save(FakeCtx("decline"), plan))

    assert declined["error"].startswith(
        "Popup answer declined by user (decline). Nothing was pressed. "
        "The client answered in")


def test_fast_decline_at_a_confirmation_point(real_threshold, monkeypatch):
    class Controller:
        def set_field(self, field_id, value):
            return {"status": "ok"}

    monkeypatch.setattr(server, "_ctrl", lambda ctx: Controller())
    monkeypatch.setattr(server, "config", server.ServerConfig(
        confirmation_floor=["field_writes"]))

    async def decline_now(message, response_type, params, context):
        return ElicitResult(action="decline")

    async def run():
        async with Client(server.mcp, elicitation_handler=decline_now) as client:
            return await client.call_tool(
                "sap_set_field", {"field_id": "wnd[0]/usr/txtX", "value": "1"},
                raise_on_error=False)

    result = asyncio.run(run())

    assert result.is_error
    text = result.content[0].text
    assert "declined by user at confirmation point 'field_writes'" in text
    assert "without showing a dialog" in text


def test_fast_attach_decline_says_so(real_threshold, monkeypatch):
    candidates = [{"index": 0, "id": "/app/con[0]/ses[0]", "system": "DEV",
                   "client": "100", "user": "TESTUSER", "transaction": "SE16",
                   "session_number": 1, "title": "DEV(1)/100 Data Browser",
                   "connection_index": 0}]

    with pytest.raises(ValueError) as exc:
        asyncio.run(server._pick_session(FakeCtx("decline"), candidates, "Which?"))

    assert str(exc.value).startswith(
        "Attach declined by user (decline). Nothing was attached. "
        "The client answered in")


# --- the yes/no prompts carry a label ----------------------------------------

def test_yes_no_prompts_are_labelled(monkeypatch):
    # The Claude Code CLI shows a boolean answer as a checkbox; without a
    # title it reads "Value". Only a ticked box with Accept proceeds.
    class Ctx:
        def __init__(self):
            self.titles = []

        async def elicit(self, message, response_type=None, response_title=None):
            self.titles.append(response_title)
            return SimpleNamespace(action="decline", data=None)

    ctx = Ctx()
    asyncio.run(server._confirm_save(ctx, "F11"))
    assert ctx.titles == ["Yes, save"]

    class Controller:
        def set_field(self, field_id, value):
            return {"status": "ok"}

    monkeypatch.setattr(server, "_ctrl", lambda c: Controller())
    monkeypatch.setattr(server, "config", server.ServerConfig(
        confirmation_floor=["field_writes"]))
    seen = []

    async def decline(message, response_type, params, context):
        seen.append(params.requestedSchema["properties"]["value"])
        return ElicitResult(action="decline")

    async def run():
        async with Client(server.mcp, elicitation_handler=decline) as client:
            await client.call_tool(
                "sap_set_field", {"field_id": "wnd[0]/usr/txtX", "value": "1"},
                raise_on_error=False)

    asyncio.run(run())
    assert seen == [{"type": "boolean", "title": "Yes, run it"}]


# --- a yes that no person gave is refused ----------------------------------

def test_fast_yes_to_a_save_is_refused(real_threshold, caplog):
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")

    declined = asyncio.run(
        server._confirm_save(FakeCtx("accept", data=True), "F11"))

    assert declined is not None, "a yes in 0 ms must not save"
    assert declined["status"] == "cancelled"
    assert "answered yes in" in declined["error"]
    assert "no person approved this" in declined["error"]
    [event] = _confirmations(caplog)
    assert event["outcome"] == "declined"
    assert event["client_answer"] == "accept"
    assert event["no_dialog"] is True


def test_slow_yes_to_a_save_goes_through(monkeypatch, caplog):
    monkeypatch.setattr(confirmation, "NO_DIALOG_MS", 30.0)
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")

    accepted = asyncio.run(server._confirm_save(
        FakeCtx("accept", delay=0.05, data=True), "F11"))

    assert accepted is None
    [event] = _confirmations(caplog)
    assert event["outcome"] == "accepted"
    assert "client_answer" not in event


def test_fast_yes_at_a_confirmation_point_runs_nothing(real_threshold, monkeypatch):
    calls = []

    class Controller:
        def set_field(self, field_id, value):
            calls.append((field_id, value))
            return {"status": "ok"}

    monkeypatch.setattr(server, "_ctrl", lambda ctx: Controller())
    monkeypatch.setattr(server, "config", server.ServerConfig(
        confirmation_floor=["field_writes"]))

    async def accept_now(message, response_type, params, context):
        return ElicitResult(action="accept", content=response_type(value=True))

    async def run():
        async with Client(server.mcp, elicitation_handler=accept_now) as client:
            return await client.call_tool(
                "sap_set_field", {"field_id": "wnd[0]/usr/txtX", "value": "1"},
                raise_on_error=False)

    result = asyncio.run(run())

    assert result.is_error
    assert "answered yes in" in result.content[0].text
    assert calls == []


def test_fast_yes_to_removing_a_point_keeps_it(real_threshold, monkeypatch):
    from mcp_sap_gui.session_manager import SessionManager

    mgr = SessionManager()
    monkeypatch.setattr(server, "_session_mgr", mgr)
    monkeypatch.setattr(server, "config", server.ServerConfig())
    ctx = SimpleNamespace(session_id="chat-1", session=object(),
                          elicit=FakeCtx("accept", data=True).elicit)
    mgr.set_confirmation_points("chat-1", {"field_writes"})

    result = asyncio.run(server.sap_set_confirmation_points([], ctx))

    assert result["removal_declined"] == ["field_writes"]
    assert "field_writes" in result["active"]
    assert "answered yes in" in result["message"]


def test_fast_pick_in_the_attach_picker_is_refused(real_threshold):
    candidates = [{"index": 0, "id": "/app/con[0]/ses[0]", "system": "DEV",
                   "client": "100", "user": "TESTUSER", "transaction": "SE16",
                   "session_number": 1, "title": "DEV(1)/100 Data Browser",
                   "connection_index": 0}]

    with pytest.raises(ValueError) as exc:
        asyncio.run(server._pick_session(
            FakeCtx("accept", data="0/0"), candidates, "Which?"))

    assert "answered yes in" in str(exc.value)
    assert "Nothing was attached" in str(exc.value)
