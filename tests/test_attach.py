"""sap_connect_existing: attach by system, and the opt-in 'attach' point.

No SAP GUI: a fake controller lists sessions and records which one is
attached, and a fake context answers the session-picker dialog.
"""

import asyncio
import json
import logging
from types import SimpleNamespace
from typing import Any, List

import pytest
from mcp.shared.exceptions import McpError
from mcp.types import ErrorData

from mcp_sap_gui import server
from mcp_sap_gui.confirmation import SETTABLE_POINTS
from mcp_sap_gui.models import SessionInfo


def _session(index, system, client, user, tcode, number, title):
    return {"index": index, "id": f"/app/con[0]/ses[{index}]", "user": user,
            "transaction": tcode, "system": system, "client": client,
            "session_number": number, "title": title}


def _connections(*sessions, disabled=False, index=0):
    return {"index": index, "id": f"/app/con[{index}]", "description": "",
            "session_count": len(sessions), "sessions": list(sessions),
            "host": "saplogon", "rot_entry": "SAPGUI",
            "scripting_disabled": disabled}


DEV_1 = _session(0, "DEV", "100", "TESTUSER", "SE16", 1,
                 "DEV(1)/100 Data Browser: Initial Screen")
DEV_2 = _session(1, "DEV", "100", "TESTUSER", "SPRO", 2,
                 "DEV(2)/100 Customizing: Execute Project")
QAS_1 = _session(0, "QAS", "200", "TESTUSER", "SESSION_MANAGER", 1,
                 "QAS(1)/200 SAP Easy Access")


class FakeController:
    def __init__(self, connections):
        self._connections = connections
        # What the positions hold when attaching, if the sessions changed
        # after the listing (None: unchanged).
        self.after = None
        self.attached: List[tuple] = []
        self.disconnected = False

    def list_connections(self):
        return self._connections

    def disconnect(self):
        self.disconnected = True
        self._session = None

    def connect_to_existing_session(self, connection_index=0, session_index=0):
        self.attached.append((connection_index, session_index))
        current = self.after or self._connections
        sess = current[connection_index]["sessions"][session_index]
        # Like the real controller: the bound session object with its Id.
        self._session = SimpleNamespace(Id=sess["id"])
        return SessionInfo(
            system_name=sess["system"], system_number="00",
            client=sess["client"], user=sess["user"], language="EN",
            transaction=sess["transaction"], program="SAPMSYST",
            screen_number=100, session_number=sess["session_number"],
        )


class FakeCtx:
    def __init__(self, answer: str = "accept", pick: Any = 0,
                 unsupported: bool = False) -> None:
        self.prompts: List[dict] = []
        self._answer = answer
        self._pick = pick
        self._unsupported = unsupported

    async def elicit(self, message, response_type=None, response_title=None):
        # Titled options: {"conn/sess": {"title": label}}; the user sees the
        # titles and the answer is the key.
        titles = [option["title"] for option in response_type.values()]
        self.prompts.append({"message": message, "choices": titles,
                             "keys": list(response_type),
                             "title": response_title})
        if self._unsupported:
            raise McpError(ErrorData(code=-32601, message="Method not found"))
        data = (list(response_type)[self._pick] if isinstance(self._pick, int)
                else self._pick)
        return SimpleNamespace(action=self._answer, data=data)


@pytest.fixture
def attach(monkeypatch):
    def setup(connections, point=False):
        controller = FakeController(connections)

        async def com(fn):
            return fn()

        monkeypatch.setattr(server, "_com", com)
        monkeypatch.setattr(server, "_ctrl", lambda ctx: controller)
        monkeypatch.setattr(server, "_session_mgr", None)
        monkeypatch.setattr(server.config, "confirmation_floor",
                            ["attach"] if point else [])
        return controller
    return setup


def _run(coro):
    return asyncio.run(coro)


def _attach_events(caplog):
    events = []
    for record in caplog.records:
        try:
            data = json.loads(record.getMessage())
        except ValueError:
            continue
        if data.get("event") in ("attach", "confirmation"):
            events.append(data)
    return events


def test_attach_is_a_settable_point():
    assert "attach" in SETTABLE_POINTS


# --- without the point -------------------------------------------------------

def test_without_point_or_system_attaches_by_index_as_before(attach, caplog):
    controller = attach([_connections(DEV_1, DEV_2)])
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")
    ctx = FakeCtx()

    result = _run(server.sap_connect_existing(ctx, connection_index=0,
                                              session_index=1))

    assert controller.attached == [(0, 1)]
    assert result["chosen_by"] == "index"
    assert ctx.prompts == []
    events = _attach_events(caplog)
    assert events[-1]["event"] == "attach"
    assert events[-1]["chosen_by"] == "index"


def test_system_with_one_open_session_attaches_it(attach):
    controller = attach([_connections(DEV_1), _connections(QAS_1, index=1)])

    result = _run(server.sap_connect_existing(FakeCtx(), system="qas"))

    assert controller.attached == [(1, 0)]
    assert result["chosen_by"] == "named"
    assert result["system_name"] == "QAS"


def test_system_and_client(attach):
    controller = attach([_connections(DEV_1), _connections(QAS_1, index=1)])

    _run(server.sap_connect_existing(FakeCtx(), system="QAS/200"))

    assert controller.attached == [(1, 0)]


def test_system_with_several_sessions_lists_them_without_the_point(attach):
    controller = attach([_connections(DEV_1, DEV_2)])

    with pytest.raises(ValueError, match="2 sessions are open on DEV") as exc:
        _run(server.sap_connect_existing(FakeCtx(), system="DEV"))

    assert "connection_index=0, session_index=1" in str(exc.value)
    assert "Customizing: Execute Project" in str(exc.value)
    assert controller.attached == []


def test_system_with_no_session_lists_what_is_open(attach):
    controller = attach([_connections(DEV_1)])

    with pytest.raises(ValueError, match="No open session on PRD") as exc:
        _run(server.sap_connect_existing(FakeCtx(), system="PRD"))

    assert "DEV (1) / 100 / TESTUSER" in str(exc.value)
    assert controller.attached == []


def test_a_disabled_connection_is_not_a_candidate(attach):
    controller = attach([_connections(DEV_1, disabled=True),
                         _connections(QAS_1, index=1)])

    with pytest.raises(ValueError, match="No open session on DEV"):
        _run(server.sap_connect_existing(FakeCtx(), system="DEV"))
    assert controller.attached == []


@pytest.mark.parametrize("bad", ["", "DE", "DEV/1", "DEV/0010", "D EV"])
def test_system_must_look_like_a_system_id(attach, bad):
    attach([_connections(DEV_1)])
    with pytest.raises(ValueError, match="expected a system ID"):
        _run(server.sap_connect_existing(FakeCtx(), system=bad))


# --- with the point ----------------------------------------------------------

def test_point_shows_every_session_even_one(attach, caplog):
    controller = attach([_connections(DEV_1)], point=True)
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")
    ctx = FakeCtx(pick=0)

    result = _run(server.sap_connect_existing(ctx))

    assert len(ctx.prompts) == 1
    assert ctx.prompts[0]["message"] == "Attach the agent to which SAP session?"
    assert ctx.prompts[0]["choices"] == [
        "DEV (1) / 100 / TESTUSER - SE16 - Data Browser: Initial Screen",
    ]
    assert ctx.prompts[0]["title"] == "SAP session"
    assert controller.attached == [(0, 0)]
    assert result["chosen_by"] == "dialog"
    events = _attach_events(caplog)
    assert [e.get("outcome") or e["event"] for e in events] == ["accepted", "attach"]
    assert events[-1]["chosen_by"] == "dialog"
    assert events[-1]["session"] == "/app/con[0]/ses[0]"


def test_point_attaches_the_session_the_user_picked(attach):
    controller = attach([_connections(DEV_1, DEV_2), _connections(QAS_1, index=1)],
                        point=True)
    ctx = FakeCtx(pick=2)

    _run(server.sap_connect_existing(ctx, connection_index=0, session_index=0))

    # The agent's indexes are not the user's choice: the dialog decides.
    assert len(ctx.prompts[0]["choices"]) == 3
    assert controller.attached == [(1, 0)]


def test_point_skips_the_dialog_when_the_system_names_one_session(attach):
    controller = attach([_connections(DEV_1), _connections(QAS_1, index=1)],
                        point=True)
    ctx = FakeCtx()

    result = _run(server.sap_connect_existing(ctx, system="DEV"))

    assert ctx.prompts == []
    assert controller.attached == [(0, 0)]
    assert result["chosen_by"] == "named"


def test_point_asks_among_the_sessions_of_a_named_system(attach):
    controller = attach([_connections(DEV_1, DEV_2), _connections(QAS_1, index=1)],
                        point=True)
    ctx = FakeCtx(pick=1)

    _run(server.sap_connect_existing(ctx, system="DEV"))

    assert ctx.prompts[0]["message"].startswith("Several sessions are open on DEV")
    assert len(ctx.prompts[0]["choices"]) == 2
    assert controller.attached == [(0, 1)]


def test_declined_dialog_attaches_nothing(attach, caplog):
    controller = attach([_connections(DEV_1)], point=True)
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")

    with pytest.raises(ValueError, match=r"Attach declined by user \(decline\)"):
        _run(server.sap_connect_existing(FakeCtx(answer="decline", pick=None)))

    assert controller.attached == []
    assert [e.get("outcome") for e in _attach_events(caplog)] == ["declined"]


def test_client_without_dialogs_fails_closed(attach, caplog):
    controller = attach([_connections(DEV_1)], point=True)
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")

    with pytest.raises(ValueError, match=r"system='SID'"):
        _run(server.sap_connect_existing(FakeCtx(unsupported=True)))

    assert controller.attached == []
    assert [e.get("outcome") for e in _attach_events(caplog)] == [
        "unsupported_client"]


def test_same_label_on_two_connections_is_told_apart(attach):
    other = dict(DEV_1)
    attach([_connections(DEV_1), _connections(other, index=1)], point=True)
    ctx = FakeCtx(pick=1)

    _run(server.sap_connect_existing(ctx))

    choices = ctx.prompts[0]["choices"]
    assert len(set(choices)) == 2
    assert choices[1].endswith("[connection 1]")


def _real_client_attach(monkeypatch, handler):
    from fastmcp import Client

    controller = FakeController([_connections(DEV_1, DEV_2)])
    monkeypatch.setattr(server, "_ctrl", lambda ctx: controller)
    monkeypatch.setattr(server, "config",
                        server.ServerConfig(confirmation_floor=["attach"]))

    async def run():
        async with Client(server.mcp, elicitation_handler=handler) as client:
            return await client.call_tool("sap_connect_existing", {},
                                          raise_on_error=False)

    return controller, asyncio.run(run())


def test_dialog_over_a_real_client(monkeypatch):
    # The options go over the wire as titled values (oneOf const/title): the
    # client shows the titles and answers with the const, the session's
    # "connection/session" pair (fastmcp in-memory client, real middleware).
    from fastmcp.client.elicitation import ElicitResult
    seen = []

    async def pick_second(message, response_type, params, context):
        seen.append(params.requestedSchema)
        choice = params.requestedSchema["properties"]["value"]["oneOf"][1]["const"]
        return ElicitResult(action="accept", content=response_type(value=choice))

    controller, result = _real_client_attach(monkeypatch, pick_second)

    assert not result.is_error
    assert controller.attached == [(0, 1)]
    assert result.structured_content["chosen_by"] == "dialog"
    value = seen[0]["properties"]["value"]
    assert value["title"] == "SAP session"
    assert value["oneOf"] == [
        {"const": "0/0",
         "title": "DEV (1) / 100 / TESTUSER - SE16 - Data Browser: Initial Screen"},
        {"const": "0/1",
         "title": "DEV (2) / 100 / TESTUSER - SPRO - Customizing: Execute Project"},
    ]


def test_a_typed_answer_that_is_no_option_attaches_nothing(monkeypatch, caplog):
    # Live in the Claude Code CLI: the picker came up as a text box and the
    # answer was "test?". fastmcp rejects it; the tool says so plainly.
    from fastmcp.client.elicitation import ElicitResult
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")

    async def type_something(message, response_type, params, context):
        return ElicitResult(action="accept", content={"value": "test?"})

    controller, result = _real_client_attach(monkeypatch, type_something)

    assert result.is_error
    assert "was not one of the listed sessions. Nothing was attached." in (
        result.content[0].text)
    assert controller.attached == []
    assert "invalid_answer" in [e.get("outcome") for e in _attach_events(caplog)]


def test_no_session_to_attach(attach):
    attach([_connections(DEV_1, disabled=True)], point=True)
    with pytest.raises(ValueError, match="No SAP session is open"):
        _run(server.sap_connect_existing(FakeCtx()))


def test_a_session_list_that_changed_meanwhile_attaches_nothing(attach):
    # The user picks DEV (2) at position 0/1; while the picker was open,
    # sessions changed and position 0/1 now holds QAS (1).
    controller = attach([_connections(DEV_1, DEV_2)], point=True)
    controller.after = [_connections(DEV_1, QAS_1)]

    with pytest.raises(ValueError, match="sessions changed") as exc:
        _run(server.sap_connect_existing(FakeCtx(pick=1)))

    assert "DEV (2)" in str(exc.value) and "QAS (1)" in str(exc.value)
    assert "Nothing was attached" in str(exc.value)
    assert controller.disconnected is True


def test_no_call_runs_on_a_changed_session_before_the_check(attach, monkeypatch):
    # Another tool call (one sent in parallel) can run between any two COM
    # calls. The session the check refuses must never be bound at such a point.
    controller = attach([_connections(DEV_1, DEV_2)], point=True)
    controller.after = [_connections(DEV_1, QAS_1)]
    bound_between_calls = []

    async def com(fn):
        result = fn()
        bound_between_calls.append(getattr(controller, "_session", None))
        return result

    monkeypatch.setattr(server, "_com", com)

    with pytest.raises(ValueError, match="sessions changed"):
        _run(server.sap_connect_existing(FakeCtx(pick=1)))

    # After the listing, and after attaching, checking and detaching.
    assert bound_between_calls == [None, None]


def test_the_audit_names_the_session_actually_attached(attach, caplog):
    attach([_connections(DEV_1, DEV_2)], point=True)
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")

    _run(server.sap_connect_existing(FakeCtx(pick=1)))

    event = _attach_events(caplog)[-1]
    assert event["event"] == "attach"
    assert (event["system"], event["client"], event["user"]) == (
        "DEV", "100", "TESTUSER")
    assert event["chosen_by"] == "dialog"
