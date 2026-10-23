"""Confirmation points on MCP 2026-07-28 (multi round-trip requests).

There a gated call returns its question (InputRequiredResult) and the client
calls again with the answer and the sealed request_state. These tests drive
the rounds by hand through the client session, to check what a client could
try: an answer to a question nobody asked, the same approval twice, the
approval with other arguments, and a gated call from a code-mode script.
"""

import json
import logging
from unittest.mock import MagicMock, patch

import pytest
from fastmcp import Client
from fastmcp.client.elicitation import ElicitResult
from fastmcp.exceptions import McpError
from mcp_types import ElicitResult as WireElicitResult
from mcp_types import InputRequiredResult

from mcp_sap_gui import server
from tests.conftest import MCP_TEST_ERA

pytestmark = pytest.mark.skipif(
    MCP_TEST_ERA == "legacy", reason="MCP 2026-07-28 only",
)

GATE = "point:field_writes"
ARGS = {"field_id": "wnd[0]/usr/txtX", "value": "1"}
YES = WireElicitResult(action="accept", content={"value": True})


async def _accept(message, response_type, params, context):
    return ElicitResult(action="accept", content=response_type(value=True))


@pytest.fixture
def gated(monkeypatch):
    """Server with field_writes as a floor point and a mock controller."""
    monkeypatch.setattr(server, "config", server.ServerConfig(
        confirmation_floor=["field_writes"],
    ))
    controller = MagicMock(Busy=False)
    controller.set_field.return_value = {"status": "ok"}

    async def inline(fn):
        return fn()

    with patch.multiple(server, _ctrl=MagicMock(return_value=controller), _com=inline):
        yield controller


async def _ask(session, args=ARGS) -> InputRequiredResult:
    first = await session.call_tool("sap_set_field", args, allow_input_required=True)
    assert isinstance(first, InputRequiredResult)
    assert list(first.input_requests) == [GATE]
    return first


async def test_answer_to_the_sealed_question_runs_the_call(gated):
    async with Client(server.mcp, elicitation_handler=_accept) as client:
        first = await _ask(client.session)
        done = await client.session.call_tool(
            "sap_set_field", ARGS, input_responses={GATE: YES},
            request_state=first.request_state, allow_input_required=True,
        )
    assert not done.is_error
    assert done.structured_content["confirmed_by_user"] is True
    gated.set_field.assert_called_once()


async def test_answer_without_a_sealed_question_asks_again(gated):
    async with Client(server.mcp, elicitation_handler=_accept) as client:
        forged = await client.session.call_tool(
            "sap_set_field", ARGS, input_responses={GATE: YES},
            allow_input_required=True,
        )
    assert isinstance(forged, InputRequiredResult)
    gated.set_field.assert_not_called()


async def test_an_approval_counts_once(gated):
    async with Client(server.mcp, elicitation_handler=_accept) as client:
        first = await _ask(client.session)
        again = {"input_responses": {GATE: YES}, "request_state": first.request_state,
                 "allow_input_required": True}
        await client.session.call_tool("sap_set_field", ARGS, **again)
        replay = await client.session.call_tool("sap_set_field", ARGS, **again)
    assert replay.is_error
    assert "already used" in replay.content[0].text
    gated.set_field.assert_called_once()


async def test_an_approval_does_not_carry_to_other_arguments(gated):
    async with Client(server.mcp, elicitation_handler=_accept) as client:
        first = await _ask(client.session)
        with pytest.raises(McpError, match="requestState"):
            await client.session.call_tool(
                "sap_set_field", {**ARGS, "value": "2"}, input_responses={GATE: YES},
                request_state=first.request_state, allow_input_required=True,
            )
    gated.set_field.assert_not_called()


async def test_question_round_is_audited(gated, caplog):
    caplog.set_level(logging.INFO, logger="mcp_sap_gui.audit")
    async with Client(server.mcp, elicitation_handler=_accept) as client:
        await _ask(client.session)
    lines = [json.loads(r.message) for r in caplog.records if r.name == "mcp_sap_gui.audit"]
    assert {"event": "tool_call", "status": "input_required"}.items() <= next(
        line for line in lines if line["event"] == "tool_call").items()
    asked = [line for line in lines if line["event"] == "confirmation"]
    assert [(a["point"], a["outcome"]) for a in asked] == [("field_writes", "asked")]
    assert "elapsed_ms" not in asked[0]


async def test_the_dialog_round_trip_through_the_client(gated):
    """fastmcp's Client answers the question with its elicitation handler
    and calls again by itself: one dialog, one write."""
    shown = []

    async def handler(message, response_type, params, context):
        shown.append(message)
        return ElicitResult(action="accept", content=response_type(value=True))

    async with Client(server.mcp, elicitation_handler=handler) as client:
        result = await client.call_tool("sap_set_field", ARGS)
    assert len(shown) == 1 and "sap_set_field" in shown[0]
    assert result.structured_content["confirmed_by_user"] is True
    gated.set_field.assert_called_once()


async def test_a_gated_call_from_a_script_is_refused_not_asked(gated):
    """On 2026-07-28 a question from inside code mode's execute would make the
    client re-run the whole script, earlier writes included. The gate refuses
    instead, and says to call the tool directly."""
    pytest.importorskip("pydantic_monty", reason="code-mode extra not installed")
    shown = []

    async def handler(message, response_type, params, context):
        shown.append(message)
        return ElicitResult(action="accept", content=response_type(value=True))

    transform = server._build_code_mode_transform()
    server.mcp.add_transform(transform)
    try:
        async with Client(server.mcp, elicitation_handler=handler) as client:
            script = f'return await call_tool("sap_set_field", {ARGS!r})'
            result = await client.call_tool("execute", {"code": script},
                                            raise_on_error=False)
    finally:
        server.mcp._transforms.remove(transform)  # do not leak into other tests

    assert shown == []
    gated.set_field.assert_not_called()
    text = " ".join(getattr(c, "text", "") for c in result.content)
    assert "called from a script" in text
