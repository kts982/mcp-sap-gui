"""``--client-approves save``: a reviewed script approves its own saves.

Since 0.5.3 a yes faster than a person can give one is refused, which left
no way to save from a script. With the launch option the question is still
asked and a no still counts, but every yes to a question for the named
points counts as the launching program's: the result says
``confirmed_by_launch`` (never ``confirmed_by_user``) and the audit line
``approved_by: launch``. Command line only, stdio only, not with
``--code-mode``; never for the attach picker or for turning a point off.
"""

from __future__ import annotations

import asyncio
import json
import logging
import sys
from types import SimpleNamespace
from typing import Any
from unittest.mock import MagicMock, patch

import pytest
from fastmcp import Client
from fastmcp.client.elicitation import ElicitResult
from fastmcp.client.transports import StdioTransport
from mcp_types import ElicitResult as WireElicitResult
from mcp_types import InputRequiredResult
from mcp_types.version import MODERN_PROTOCOL_VERSIONS

from mcp_sap_gui import confirmation, server
from tests.conftest import MCP_TEST_ERA
from tests.test_popup_save_gate import SAVE_BUTTON, SAVE_FIRST, FakeController, _open

REAL_THRESHOLD = 250.0
YES = WireElicitResult(action="accept", content={"value": True})
modern_only = pytest.mark.skipif(MCP_TEST_ERA == "legacy", reason="MCP 2026-07-28 only")


@pytest.fixture
def real_threshold(monkeypatch):
    monkeypatch.setattr(confirmation, "NO_DIALOG_MS", REAL_THRESHOLD)


def _use_config(monkeypatch, **kwargs) -> server.ServerConfig:
    cfg = server.ServerConfig(**kwargs)
    monkeypatch.setattr(server, "config", cfg)
    return cfg


def _save_controller():
    """A controller whose screen stays the same: a save goes out."""
    controller = MagicMock(Busy=False)
    controller.check_vkey.return_value = None
    controller.send_vkey.return_value = {"status": "ok"}
    controller.press_button.return_value = {"status": "ok"}
    controller.save_target.return_value = {
        "active_window": "wnd[0]", "transaction": "SU3", "program": "SAPLSUU5",
        "screen_number": "100", "title": "Maintain User Profile", "system": "DEV",
        "client": "100", "user": "TESTUSER", "session": "/app/con[0]/ses[0]",
    }
    return controller


async def _inline(fn):
    return fn()


def _patched(controller):
    return patch.multiple(server, _ctrl=MagicMock(return_value=controller), _com=_inline)


async def _accept_now(message, response_type, params, context):
    return ElicitResult(action="accept", content=response_type(value=True))


async def _decline_now(message, response_type, params, context):
    return ElicitResult(action="decline")


async def _accept_slowly(message, response_type, params, context):
    await asyncio.sleep(0.05)
    return ElicitResult(action="accept", content=response_type(value=True))


async def _call(tool, args, handler):
    async with Client(server.mcp, elicitation_handler=handler) as client:
        return await client.call_tool(tool, args, raise_on_error=False)


def _confirmations(caplog):
    lines = []
    for record in caplog.records:
        if record.name != "mcp_sap_gui.audit":
            continue
        line = json.loads(record.getMessage())
        if line.get("event") == "confirmation" and line["outcome"] != "asked":
            lines.append(line)
    return lines


# ---------------------------------------------------------------------------
# A fast yes: refused without the option, the launcher's with it
# ---------------------------------------------------------------------------


async def test_fast_yes_to_a_save_is_refused_without_the_option(real_threshold, monkeypatch):
    _use_config(monkeypatch)
    controller = _save_controller()
    with _patched(controller):
        result = await _call("sap_send_key", {"key": "F11"}, _accept_now)
    sc = result.structured_content
    assert sc["status"] == "cancelled"
    assert "no person approved this" in sc["error"]
    controller.send_vkey.assert_not_called()


async def test_fast_yes_to_a_save_counts_as_the_launchers(real_threshold, monkeypatch, caplog):
    _use_config(monkeypatch, client_approves=["save"])
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")
    controller = _save_controller()
    with _patched(controller):
        result = await _call("sap_send_key", {"key": "F11"}, _accept_now)
    sc = result.structured_content
    assert sc["confirmed_by_launch"] is True
    assert "confirmed_by_user" not in sc
    controller.send_vkey.assert_called_once()
    [line] = _confirmations(caplog)
    assert (line["point"], line["outcome"], line["approved_by"]) == (
        "save", "accepted", "launch")
    # How long the answer took, and that no person could have given it,
    # stay on the line.
    assert line["elapsed_ms"] < REAL_THRESHOLD
    assert line["no_dialog"] is True
    assert "client_answer" not in line


async def test_save_button_under_the_option(real_threshold, monkeypatch):
    _use_config(monkeypatch, client_approves=["save"])
    controller = _save_controller()
    with _patched(controller):
        result = await _call("sap_press_button", {"button_id": SAVE_BUTTON}, _accept_now)
    assert result.structured_content["confirmed_by_launch"] is True
    assert "confirmed_by_user" not in result.structured_content
    controller.press_button.assert_called_once_with(SAVE_BUTTON)


async def test_slow_yes_under_the_option_is_the_launchers_too(monkeypatch, caplog):
    # A yes slow enough for a person: with the option on, the server cannot
    # tell a person from the script, so it still reads as the launcher's.
    monkeypatch.setattr(confirmation, "NO_DIALOG_MS", 30.0)
    _use_config(monkeypatch, client_approves=["save"])
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")
    controller = _save_controller()
    with _patched(controller):
        result = await _call("sap_send_key", {"key": "F11"}, _accept_slowly)
    sc = result.structured_content
    assert sc["confirmed_by_launch"] is True
    assert "confirmed_by_user" not in sc
    [line] = _confirmations(caplog)
    assert line["approved_by"] == "launch"
    assert line["elapsed_ms"] >= 30.0
    assert "no_dialog" not in line


async def test_slow_yes_without_the_option_is_a_persons(monkeypatch, caplog):
    monkeypatch.setattr(confirmation, "NO_DIALOG_MS", 30.0)
    _use_config(monkeypatch)
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")
    with _patched(_save_controller()):
        result = await _call("sap_send_key", {"key": "F11"}, _accept_slowly)
    assert result.structured_content["confirmed_by_user"] is True
    assert "confirmed_by_launch" not in result.structured_content
    [line] = _confirmations(caplog)
    assert line["approved_by"] == "user"


@pytest.mark.parametrize("threshold", [REAL_THRESHOLD, 0.0])
async def test_a_no_still_counts_under_the_option(monkeypatch, caplog, threshold):
    monkeypatch.setattr(confirmation, "NO_DIALOG_MS", threshold)
    _use_config(monkeypatch, client_approves=["save"])
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")
    controller = _save_controller()
    with _patched(controller):
        result = await _call("sap_send_key", {"key": "F11"}, _decline_now)
    sc = result.structured_content
    assert (sc["status"], sc["action"]) == ("cancelled", "decline")
    assert "confirmed_by_launch" not in sc
    controller.send_vkey.assert_not_called()
    [line] = _confirmations(caplog)
    assert line["outcome"] == "declined"
    assert "approved_by" not in line


async def test_a_confirmation_point_the_option_does_not_name_still_needs_a_person(
        real_threshold, monkeypatch):
    _use_config(monkeypatch, client_approves=["save"], confirmation_floor=["field_writes"])
    controller = MagicMock(Busy=False)
    controller.set_field.return_value = {"status": "ok"}
    with _patched(controller):
        result = await _call(
            "sap_set_field", {"field_id": "wnd[0]/usr/txtX", "value": "1"}, _accept_now)
    assert result.is_error
    assert "answered yes in" in result.content[0].text
    controller.set_field.assert_not_called()


# ---------------------------------------------------------------------------
# Popups that may save
# ---------------------------------------------------------------------------


class _FastCtx:
    """A handshake-era context whose client answers yes at once."""

    def __init__(self) -> None:
        self.prompts: list[str] = []

    async def elicit(self, message: str, response_type: Any = None, **_: Any):
        self.prompts.append(message)
        return SimpleNamespace(action="accept", data=True)


@pytest.fixture
def popup_tools(monkeypatch):
    controller = FakeController()
    monkeypatch.setattr(server, "_com", _inline)
    monkeypatch.setattr(server, "_ctrl", lambda ctx: controller)
    monkeypatch.setattr(server, "_session_mgr", None)
    return controller


def test_popup_answer_under_the_option(real_threshold, monkeypatch, popup_tools):
    _use_config(monkeypatch, client_approves=["save"])
    _open(popup_tools, SAVE_FIRST)
    ctx = _FastCtx()
    result = asyncio.run(server.sap_handle_popup(ctx, action="confirm"))
    assert len(ctx.prompts) == 1
    assert result["confirmed_by_launch"] is True
    assert "confirmed_by_user" not in result
    assert popup_tools.session.log == ["press wnd[1]/usr/btnSPOP-OPTION1"]


def test_popup_button_press_under_the_option(real_threshold, monkeypatch, popup_tools):
    _use_config(monkeypatch, client_approves=["save"])
    _open(popup_tools, SAVE_FIRST)
    result = asyncio.run(server.sap_press_button("wnd[1]/usr/btnSPOP-OPTION1", _FastCtx()))
    assert result["confirmed_by_launch"] is True
    assert "confirmed_by_user" not in result


def test_popup_answer_without_the_option_is_refused(real_threshold, monkeypatch, popup_tools):
    _use_config(monkeypatch)
    _open(popup_tools, SAVE_FIRST)
    result = asyncio.run(server.sap_handle_popup(_FastCtx(), action="confirm"))
    assert result["status"] == "cancelled"
    assert popup_tools.session.log == []


# ---------------------------------------------------------------------------
# What the option never answers
# ---------------------------------------------------------------------------


def _wide_option(monkeypatch, **kwargs) -> None:
    """The option as if it named every point, past the command line's
    choices: the attach picker and turning a point off must still need a
    person, whatever the option names."""
    cfg = _use_config(monkeypatch, **kwargs)
    cfg.client_approves = list(confirmation.ALL_POINTS)


def test_the_option_does_not_answer_the_attach_picker(real_threshold, monkeypatch):
    _wide_option(monkeypatch)
    candidates = [{"index": 0, "id": "/app/con[0]/ses[0]", "system": "DEV",
                   "client": "100", "user": "TESTUSER", "transaction": "SE16",
                   "session_number": 1, "title": "DEV(1)/100 Data Browser",
                   "connection_index": 0}]

    class PickAtOnce:
        async def elicit(self, message, response_type=None, **_):
            return SimpleNamespace(action="accept", data="0/0")

    with pytest.raises(ValueError) as exc:
        asyncio.run(server._pick_session(PickAtOnce(), candidates, "Which?"))
    assert "answered yes in" in str(exc.value)
    assert "Nothing was attached" in str(exc.value)


def test_the_option_does_not_answer_turning_a_point_off(real_threshold, monkeypatch):
    from mcp_sap_gui.session_manager import SessionManager

    _wide_option(monkeypatch)
    mgr = SessionManager()
    monkeypatch.setattr(server, "_session_mgr", mgr)
    mgr.set_confirmation_points("chat-1", {"field_writes"})
    ctx = SimpleNamespace(session_id="chat-1", session=object(), elicit=_FastCtx().elicit)

    result = asyncio.run(server.sap_set_confirmation_points([], ctx))

    assert result["removal_declined"] == ["field_writes"]
    assert "field_writes" in result["active"]
    assert "answered yes in" in result["message"]


def test_a_question_for_several_points_needs_them_all(real_threshold):
    """A question that covers several points passes only if the option
    names every one of them."""
    def ask(launch_points):
        return asyncio.run(confirmation.ask_user(
            _FastCtx(), points=["save", "field_writes"], tool="t", message="m",
            launch_points=launch_points,
        ))

    some = ask(["save"])
    assert not some.yes and "answered yes in" in some.note
    every = ask(["save", "field_writes"])
    assert every.yes and every.approved_by == "launch"


async def test_no_tool_sets_or_clears_the_option(monkeypatch):
    cfg = _use_config(monkeypatch, client_approves=["save"])
    with _patched(MagicMock(Busy=False)):
        async with Client(server.mcp, elicitation_handler=_accept_now) as client:
            await client.call_tool("sap_set_confirmation_points", {"points": ["field_writes"]})
            await client.call_tool("sap_set_confirmation_points", {"points": []})
            tools = await client.list_tools()
    assert server.config is cfg
    assert cfg.client_approves == ["save"]
    # No tool takes it as an argument either.
    for tool in tools:
        assert not [name for name in tool.input_schema.get("properties", {})
                    if "approve" in name], tool.name


async def test_the_agent_is_never_told_about_the_option():
    """An agent that hits a refused save in a chat client must not suggest
    the option: no instruction, tool description or refusal names it."""
    texts = [server.mcp.instructions or "", confirmation.NO_DIALOGS_NOTE,
             confirmation.fast_yes_note(5.0), confirmation.no_dialog_note(5.0)]
    async with Client(server.mcp) as client:
        texts += [tool.description or "" for tool in await client.list_tools()]
    for text in texts:
        assert "client-approves" not in text
        assert "client_approves" not in text
        assert "confirmed_by_launch" not in text


def test_unknown_point_is_refused():
    with pytest.raises(ValueError, match="--client-approves"):
        server.ServerConfig(client_approves=["field_writes"])


# ---------------------------------------------------------------------------
# MCP 2026-07-28: the input-required rounds
# ---------------------------------------------------------------------------


@modern_only
async def test_fast_yes_through_the_input_required_round(real_threshold, monkeypatch):
    _use_config(monkeypatch, client_approves=["save"])
    controller = _save_controller()
    with _patched(controller):
        async with Client(server.mcp, elicitation_handler=_accept_now) as client:
            first = await client.session.call_tool(
                "sap_send_key", {"key": "F11"}, allow_input_required=True)
            assert isinstance(first, InputRequiredResult)
            assert list(first.input_requests) == ["save"]
            done = await client.session.call_tool(
                "sap_send_key", {"key": "F11"}, input_responses={"save": YES},
                request_state=first.request_state, allow_input_required=True,
            )
    assert done.structured_content["confirmed_by_launch"] is True
    assert "confirmed_by_user" not in done.structured_content
    controller.send_vkey.assert_called_once()


@modern_only
async def test_an_approval_carried_into_a_second_round_keeps_the_launch_marker(
        real_threshold, monkeypatch):
    """all_writes asks first, then the save. The first approval is carried
    into the save's round, and must still read as the launcher's there.
    No question follows a save in one call today, so the option is widened
    here past the command line's choices to name all_writes as well."""
    cfg = _use_config(monkeypatch, client_approves=["save"],
                      confirmation_floor=["all_writes"])
    cfg.client_approves = ["all_writes", "save"]
    controller = _save_controller()
    args = {"button_id": SAVE_BUTTON}
    with _patched(controller):
        async with Client(server.mcp, elicitation_handler=_accept_now) as client:
            first = await client.session.call_tool(
                "sap_press_button", args, allow_input_required=True)
            assert list(first.input_requests) == ["point:all_writes"]
            second = await client.session.call_tool(
                "sap_press_button", args, input_responses={"point:all_writes": YES},
                request_state=first.request_state, allow_input_required=True,
            )
            assert isinstance(second, InputRequiredResult)
            assert list(second.input_requests) == ["save"]
            done = await client.session.call_tool(
                "sap_press_button", args, input_responses={"save": YES},
                request_state=second.request_state, allow_input_required=True,
            )
    sc = done.structured_content
    assert sc["confirmed_by_launch"] is True
    assert "confirmed_by_user" not in sc
    controller.press_button.assert_called_once_with(SAVE_BUTTON)


class _ModernCtx:
    """Just enough of a 2026-07-28 context for ask_user: one round."""

    def __init__(self, request_state=None, input_responses=None) -> None:
        self.request_context = SimpleNamespace(
            protocol_version=next(iter(MODERN_PROTOCOL_VERSIONS)))
        self.session = SimpleNamespace(
            client_capabilities=SimpleNamespace(elicitation=object()))
        self.request_state = request_state
        self.input_responses = input_responses
        self._state: dict[str, Any] = {}

    async def get_state(self, key):
        return self._state.get(key)

    async def set_state(self, key, value, serializable=True):
        self._state[key] = value


def test_a_save_approval_carried_into_the_next_round_keeps_its_marker(real_threshold):
    """A save approved in one round, when another question follows it in
    the same call, counts in the next round as the launcher's again: the
    round state records who approved it."""
    def save(ctx):
        return asyncio.run(confirmation.ask_user(
            ctx, points="save", tool="t", message="Save?", gate="save",
            can_ask=True, launch_points=["save"]))

    asked = save(_ModernCtx())
    assert isinstance(asked, confirmation.Ask)

    answer_round = _ModernCtx(asked.result.request_state, {"save": YES})
    approved = save(answer_round)
    assert (approved.yes, approved.approved_by) == (True, "launch")
    # Another question in the same round: the client calls again.
    follow_up = asyncio.run(confirmation.ask_user(
        answer_round, points="field_writes", tool="t", message="Write?",
        gate="point:field_writes", can_ask=True))
    assert isinstance(follow_up, confirmation.Ask)

    carried = save(_ModernCtx(follow_up.result.request_state,
                              {"point:field_writes": YES}))
    assert (carried.yes, carried.approved_by) == (True, "launch")


# ---------------------------------------------------------------------------
# The command line
# ---------------------------------------------------------------------------


def _main(monkeypatch, *argv):
    """server.main() without a transport; config is restored afterwards."""
    monkeypatch.setattr(server, "config", server.config)
    with patch("sys.argv", ["mcp-sap-gui", *argv]), \
         patch.object(server, "load_dotenv", MagicMock()), \
         patch.object(server.mcp, "run", MagicMock()) as run:
        server.main()
    return run


def test_startup_warning_names_the_option_and_its_points(monkeypatch, caplog):
    caplog.set_level(logging.WARNING, logger="mcp_sap_gui.server")
    _main(monkeypatch, "--client-approves", "save")
    assert server.config.client_approves == ["save"]
    [warning] = [r for r in caplog.records
                 if r.levelno == logging.WARNING and "--client-approves" in r.getMessage()]
    assert "--client-approves save" in warning.getMessage()
    assert "not a person's" in warning.getMessage()


def test_no_warning_without_the_option(monkeypatch, caplog):
    caplog.set_level(logging.WARNING, logger="mcp_sap_gui.server")
    _main(monkeypatch)
    assert server.config.client_approves == []
    assert not [r for r in caplog.records if "--client-approves" in r.getMessage()]


@pytest.mark.parametrize("argv, refusal", [
    (["--transport", "http"], "stdio only"),
    (["--code-mode"], "--code-mode"),
])
def test_refused_over_http_and_with_code_mode(monkeypatch, capsys, argv, refusal):
    with pytest.raises(SystemExit) as exc:
        _main(monkeypatch, "--client-approves", "save", *argv)
    assert exc.value.code == 2
    assert refusal in capsys.readouterr().err


def test_only_save_can_be_named(monkeypatch, capsys):
    with pytest.raises(SystemExit):
        _main(monkeypatch, "--client-approves", "field_writes")
    assert "invalid choice" in capsys.readouterr().err


async def test_the_option_starts_over_stdio_with_its_warning(tmp_path):
    """A real stdio server process: the option is accepted there, and the
    WARNING goes to stderr at start-up."""
    log = tmp_path / "stderr.log"
    transport = StdioTransport(
        command=sys.executable,
        args=["-m", "mcp_sap_gui.server", "--client-approves", "save"],
        log_file=log,
    )
    async with Client(transport) as client:
        assert await client.list_tools()
    stderr = log.read_text(encoding="utf-8", errors="replace")
    assert "WARNING" in stderr and "--client-approves save" in stderr
