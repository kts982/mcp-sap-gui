"""Questions asked in a tool body, on MCP 2026-07-28.

There the body runs again when the answer arrives, so it reads SAP again
before acting. An answer counts only if what the question showed still
holds: the screen for F11, the popup and its button for a popup answer, the
session list for the attach picker, the points for a removal. These tests
change SAP between the question round and the answer round.
"""

import time
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import pytest
from fastmcp import Client
from fastmcp.client.elicitation import ElicitResult
from mcp_types import ElicitResult as WireElicitResult
from mcp_types import InputRequiredResult

from mcp_sap_gui import confirmation, server
from tests.conftest import MCP_TEST_ERA
from tests.test_attach import DEV_1, DEV_2, FakeController, _connections

pytestmark = pytest.mark.skipif(
    MCP_TEST_ERA == "legacy", reason="MCP 2026-07-28 only",
)

YES = WireElicitResult(action="accept", content={"value": True})
NO = WireElicitResult(action="decline")


async def _accept(message, response_type, params, context):
    return ElicitResult(action="accept", content=response_type(value=True))


async def _inline(fn):
    return fn()


@pytest.fixture
def default_config(monkeypatch):
    monkeypatch.setattr(server, "config", server.ServerConfig())


async def _rounds(session, tool, args, answer, between=lambda: None):
    """Ask, run *between* (SAP changes), answer; return both results."""
    first = await session.call_tool(tool, args, allow_input_required=True)
    assert isinstance(first, InputRequiredResult)
    [gate] = list(first.input_requests)
    between()
    second = await session.call_tool(
        tool, args, input_responses={gate: answer},
        request_state=first.request_state, allow_input_required=True,
    )
    return first, second


# ---------------------------------------------------------------------------
# F11 / Save
# ---------------------------------------------------------------------------


def _save_controller(titles):
    controller = MagicMock(Busy=False)
    controller.check_vkey.return_value = None
    controller.send_vkey.return_value = {"status": "ok"}
    screens = iter(titles)
    controller.get_screen_info.side_effect = lambda include_popup=True: {
        "active_window": "wnd[0]", "transaction": "SU3", "program": "SAPLSUU5",
        "screen_number": 100, "title": next(screens),
    }
    return controller


async def test_save_is_sent_on_the_screen_it_was_asked_on(default_config):
    controller = _save_controller(["Maintain User Profile"] * 2)
    with patch.multiple(server, _ctrl=MagicMock(return_value=controller), _com=_inline):
        async with Client(server.mcp, elicitation_handler=_accept) as client:
            first, done = await _rounds(client.session, "sap_send_key", {"key": "F11"}, YES)
    assert list(first.input_requests) == ["save"]
    assert done.structured_content["confirmed_by_user"] is True
    controller.send_vkey.assert_called_once()


async def test_save_is_not_sent_after_the_screen_changed(default_config):
    controller = _save_controller(["Maintain User Profile", "Change Material"])
    with patch.multiple(server, _ctrl=MagicMock(return_value=controller), _com=_inline):
        async with Client(server.mcp, elicitation_handler=_accept) as client:
            _, done = await _rounds(client.session, "sap_send_key", {"key": "F11"}, YES)
    assert done.structured_content["status"] == "cancelled"
    assert "changed before the answer arrived" in done.structured_content["error"]
    controller.send_vkey.assert_not_called()


def _slow_answer_round(controller):
    """check_vkey takes 300 ms when the answer round reads SAP again."""
    calls = []

    def check_vkey(*args, **kwargs):
        calls.append(1)
        if len(calls) > 1:
            time.sleep(0.3)

    controller.check_vkey.side_effect = check_vkey


async def test_a_yes_is_timed_without_the_servers_own_work(default_config, monkeypatch):
    # The client answers at once; only the server's SAP reads in the answer
    # round take time. That must not make the yes look like a person's.
    monkeypatch.setattr(confirmation, "NO_DIALOG_MS", 250.0)
    controller = _save_controller(["Maintain User Profile"] * 2)
    _slow_answer_round(controller)
    with patch.multiple(server, _ctrl=MagicMock(return_value=controller), _com=_inline):
        async with Client(server.mcp, elicitation_handler=_accept) as client:
            _, done = await _rounds(client.session, "sap_send_key", {"key": "F11"}, YES)
    assert done.structured_content["status"] == "cancelled"
    assert "answered yes in" in done.structured_content["error"]
    controller.send_vkey.assert_not_called()


async def test_a_person_who_takes_their_time_still_saves(default_config, monkeypatch):
    monkeypatch.setattr(confirmation, "NO_DIALOG_MS", 250.0)
    controller = _save_controller(["Maintain User Profile"] * 2)
    with patch.multiple(server, _ctrl=MagicMock(return_value=controller), _com=_inline):
        async with Client(server.mcp, elicitation_handler=_accept) as client:
            _, done = await _rounds(client.session, "sap_send_key", {"key": "F11"}, YES,
                                    between=lambda: time.sleep(0.3))
    assert done.structured_content["confirmed_by_user"] is True
    controller.send_vkey.assert_called_once()


async def test_declined_save_over_mrtr(default_config):
    controller = _save_controller(["Maintain User Profile"] * 2)
    with patch.multiple(server, _ctrl=MagicMock(return_value=controller), _com=_inline):
        async with Client(server.mcp, elicitation_handler=_accept) as client:
            _, done = await _rounds(client.session, "sap_send_key", {"key": "F11"}, NO)
    assert done.structured_content["status"] == "cancelled"
    assert done.structured_content["action"] == "decline"
    controller.send_vkey.assert_not_called()


# ---------------------------------------------------------------------------
# A popup answer that may save
# ---------------------------------------------------------------------------


def _popup_plan(texts):
    return {
        "save_risk": True,
        "popup": {"window_id": "wnd[1]", "title": "Exit Editor", "texts": texts},
        "button": {"id": "wnd[1]/usr/btnSPOP-OPTION1", "text": "Yes", "tooltip": ""},
    }


async def test_popup_answer_counts_only_for_the_popup_shown(default_config):
    controller = MagicMock(Busy=False)
    plans = iter([_popup_plan(["Save changes?"]), _popup_plan(["Delete the object?"])])
    controller.plan_popup_action.side_effect = lambda action, button_text: next(plans)
    with patch.multiple(server, _ctrl=MagicMock(return_value=controller), _com=_inline):
        async with Client(server.mcp, elicitation_handler=_accept) as client:
            _, done = await _rounds(
                client.session, "sap_handle_popup", {"action": "confirm"}, YES,
            )
    assert done.structured_content["status"] == "cancelled"
    assert "changed before the answer arrived" in done.structured_content["error"]
    controller.apply_popup_plan.assert_not_called()


async def test_popup_answer_for_the_same_popup_is_pressed(default_config):
    controller = MagicMock(Busy=False)
    controller.plan_popup_action.side_effect = (
        lambda action, button_text: _popup_plan(["Save changes?"]))
    controller.apply_popup_plan.return_value = {"action": "pressed"}
    with patch.multiple(server, _ctrl=MagicMock(return_value=controller), _com=_inline):
        async with Client(server.mcp, elicitation_handler=_accept) as client:
            _, done = await _rounds(
                client.session, "sap_handle_popup", {"action": "confirm"}, YES,
            )
    assert done.structured_content["confirmed_by_user"] is True
    controller.apply_popup_plan.assert_called_once()


async def test_an_approval_from_an_earlier_round_counts_once(monkeypatch):
    # all_writes asks first, then the body asks about the save. Sending the
    # second round's sealed state again must not press another popup.
    monkeypatch.setattr(server, "config", server.ServerConfig(confirmation_floor=["all_writes"]))
    controller = MagicMock(Busy=False)
    plans = iter([_popup_plan(["Save changes?"])] * 2
                 + [dict(_popup_plan(["Continue?"]), save_risk=False)])
    controller.plan_popup_action.side_effect = lambda action, button_text: next(plans)
    controller.apply_popup_plan.return_value = {"action": "pressed"}
    args = {"action": "confirm"}
    with patch.multiple(server, _ctrl=MagicMock(return_value=controller), _com=_inline):
        async with Client(server.mcp, elicitation_handler=_accept) as client:
            session = client.session
            first = await session.call_tool("sap_handle_popup", args,
                                            allow_input_required=True)
            assert list(first.input_requests) == ["point:all_writes"]
            second = await session.call_tool(
                "sap_handle_popup", args, input_responses={"point:all_writes": YES},
                request_state=first.request_state, allow_input_required=True)
            assert list(second.input_requests) == ["save"]
            done = await session.call_tool(
                "sap_handle_popup", args, input_responses={"save": YES},
                request_state=second.request_state, allow_input_required=True)
            assert done.structured_content["confirmed_by_user"] is True

            replay = await session.call_tool(
                "sap_handle_popup", args, request_state=second.request_state,
                allow_input_required=True)
    assert replay.is_error
    assert "already used" in replay.content[0].text
    controller.apply_popup_plan.assert_called_once()


# ---------------------------------------------------------------------------
# The attach picker
# ---------------------------------------------------------------------------


async def test_picker_answer_counts_only_for_the_sessions_shown(monkeypatch):
    monkeypatch.setattr(server, "config", server.ServerConfig(confirmation_floor=["attach"]))
    fake = FakeController([_connections(DEV_1, DEV_2)])
    with patch.object(server, "_ctrl", return_value=fake), \
            patch.object(server, "_com", _inline):
        async with Client(server.mcp, elicitation_handler=_accept) as client:
            pick = WireElicitResult(action="accept", content={"value": "0/1"})

            def another_session_opened():
                third = dict(DEV_2, index=2, id="/app/con[0]/ses[2]", session_number=3,
                             title="DEV(3)/100 SAP Easy Access")
                fake._connections = [_connections(DEV_1, DEV_2, third)]

            _, done = await _rounds(client.session, "sap_connect_existing", {},
                                    pick, another_session_opened)
    assert done.is_error
    assert "Nothing was attached" in done.content[0].text
    assert fake.attached == []


async def test_picker_attaches_the_session_picked(monkeypatch):
    monkeypatch.setattr(server, "config", server.ServerConfig(confirmation_floor=["attach"]))
    fake = FakeController([_connections(DEV_1, DEV_2)])
    with patch.object(server, "_ctrl", return_value=fake), \
            patch.object(server, "_com", _inline):
        async with Client(server.mcp, elicitation_handler=_accept) as client:
            pick = WireElicitResult(action="accept", content={"value": "0/1"})
            _, done = await _rounds(client.session, "sap_connect_existing", {}, pick)
    assert not done.is_error
    assert done.structured_content["chosen_by"] == "dialog"
    assert fake.attached == [(0, 1)]


# ---------------------------------------------------------------------------
# Removing a confirmation point
# ---------------------------------------------------------------------------


async def test_removal_over_mrtr(default_config):
    async with Client(server.mcp, elicitation_handler=_accept) as client:
        await client.call_tool("sap_set_confirmation_points", {"points": ["field_writes"]})
        _, declined = await _rounds(client.session, "sap_set_confirmation_points",
                                    {"points": []}, NO)
        assert declined.structured_content["removal_declined"] == ["field_writes"]
        _, accepted = await _rounds(client.session, "sap_set_confirmation_points",
                                    {"points": []}, YES)
        assert accepted.structured_content["active"] == ["save"]


async def test_save_from_a_script_is_refused(default_config):
    pytest.importorskip("pydantic_monty", reason="code-mode extra not installed")
    controller = _save_controller(["Maintain User Profile"] * 2)
    transform = server._build_code_mode_transform()
    server.mcp.add_transform(transform)
    try:
        with patch.multiple(server, _ctrl=MagicMock(return_value=controller), _com=_inline):
            async with Client(server.mcp, elicitation_handler=_accept) as client:
                result = await client.call_tool(
                    "execute", {"code": 'return await call_tool("sap_send_key", {"key": "F11"})'},
                    raise_on_error=False,
                )
    finally:
        server.mcp._transforms.remove(transform)
    text = " ".join(getattr(c, "text", "") for c in result.content)
    assert "called from a script" in text
    controller.send_vkey.assert_not_called()


def test_fixture_session_ids_are_distinct():
    # Guard for the picker tests: the third session must not reuse an Id.
    assert SimpleNamespace(**DEV_1).id != SimpleNamespace(**DEV_2).id
