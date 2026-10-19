"""The save gate on popups (sap_handle_popup) and the Save key's order of
checks (sap_send_key). Fakes only: no SAP GUI, no COM."""

from __future__ import annotations

import asyncio
from types import SimpleNamespace
from typing import Any, Dict, List

import pytest

from mcp_sap_gui import server
from mcp_sap_gui.sap_controller import SAPGUIController


class FakeElement:
    """A COM element that records what was done to it."""

    def __init__(self, element_id: str, log: List[str], text: str = "",
                 allowed_vkeys: set[int] | None = None) -> None:
        self.Id = element_id
        self.Text = text
        self._log = log
        self._allowed = allowed_vkeys

    def press(self) -> None:
        self._log.append(f"press {self.Id}")

    def sendVKey(self, vkey: int) -> None:  # noqa: N802 - COM name
        self._log.append(f"vkey {vkey} {self.Id}")

    def IsVKeyAllowed(self, vkey: int) -> bool:  # noqa: N802 - COM name
        if self._allowed is None:
            raise AttributeError("IsVKeyAllowed")
        return vkey in self._allowed


class FakeSession:
    """findById over a dict of FakeElements; a missing ID raises like COM."""

    def __init__(self) -> None:
        self.log: List[str] = []
        self.elements: Dict[str, FakeElement] = {}
        self.Busy = False

    def add(self, element_id: str, **kwargs: Any) -> FakeElement:
        element = FakeElement(element_id, self.log, **kwargs)
        self.elements[element_id] = element
        return element

    def findById(self, element_id: str) -> FakeElement:  # noqa: N802 - COM name
        if element_id not in self.elements:
            raise Exception(f"element not found: {element_id}")
        return self.elements[element_id]


class FakeController(SAPGUIController):
    """SAPGUIController without pywin32: popup reads come from ``popups``."""

    def __init__(self) -> None:  # no super(): skips the pywin32 check
        self._session = FakeSession()
        self.popups: List[Dict[str, Any]] = []
        self.scripting_read_only = False

    @property
    def session(self) -> FakeSession:
        return self._session

    def _require_session(self) -> None:
        return None

    def get_popup_window(self) -> Dict[str, Any]:
        if not self.popups:
            return {"popup_exists": False}
        popup = dict(self.popups.pop(0))
        popup.setdefault("popup_exists", True)
        return self._classify_popup(popup)

    def get_screen_info(self, include_popup: bool = True) -> Dict[str, Any]:
        return {"active_window": "wnd[0]"}


@pytest.fixture
def controller() -> FakeController:
    return FakeController()

SAVE_FIRST = {
    "window_id": "wnd[1]",
    "title": "Exit Editing",
    "texts": ["Save the changes first?"],
    "buttons": [
        {"id": "wnd[1]/usr/btnSPOP-OPTION1", "text": "Yes", "tooltip": ""},
        {"id": "wnd[1]/usr/btnSPOP-OPTION2", "text": "No", "tooltip": ""},
        {"id": "wnd[1]/usr/btnSPOP-OPTION_CAN", "text": "Cancel", "tooltip": ""},
    ],
}

SAVE_FIRST_DE = {
    "window_id": "wnd[1]",
    "title": "Bearbeiten beenden",
    "texts": ["Sollen die Daten vorher gesichert werden? Sichern"],
    "buttons": [
        {"id": "wnd[1]/usr/btnSPOP-OPTION1", "text": "Ja", "tooltip": ""},
        {"id": "wnd[1]/usr/btnSPOP-OPTION2", "text": "Nein", "tooltip": ""},
    ],
}

CONTINUE = {
    "window_id": "wnd[1]",
    "title": "Information",
    "texts": ["Material 4711 is locked by user SMITH"],
    "buttons": [
        {"id": "wnd[1]/tbar[0]/btn[0]", "text": "", "tooltip": "Continue (Enter)"},
    ],
}

SAVED_NOTICE = {
    "window_id": "wnd[1]",
    "title": "Information",
    "texts": ["Data was saved"],
    "buttons": [
        {"id": "wnd[1]/tbar[0]/btn[0]", "text": "", "tooltip": "Continue (Enter)"},
    ],
}

NO_BUTTONS = {
    "window_id": "wnd[1]",
    "title": "Question",
    "texts": ["Do you want to proceed?"],
}


def _open(controller, popup: Dict[str, Any]) -> None:
    """Show *popup* for one read and register its window and buttons."""
    controller.popups.append(popup)
    controller.session.add("wnd[0]")
    controller.session.add(popup["window_id"], text=popup.get("title", ""))
    for btn in popup.get("buttons", []):
        controller.session.add(btn["id"])


# ---------------------------------------------------------------------------
# Controller: plan_popup_action / apply_popup_plan
# ---------------------------------------------------------------------------

def test_confirm_on_save_prompt_is_a_save_risk(controller):
    _open(controller, SAVE_FIRST)
    plan = controller.plan_popup_action("confirm")
    assert plan["save_risk"] is True
    assert plan["button"]["text"] == "Yes"
    assert controller.session.log == []  # planning presses nothing


def test_save_risk_from_german_text(controller):
    _open(controller, SAVE_FIRST_DE)
    plan = controller.plan_popup_action("press", "Ja")
    assert plan["save_risk"] is True


def test_pressing_no_on_save_prompt_is_not_gated(controller):
    _open(controller, SAVE_FIRST)
    plan = controller.plan_popup_action("press", "No")
    assert plan["save_risk"] is False
    result = controller.apply_popup_plan(plan)
    assert result["action"] == "pressed"
    assert controller.session.log == ["press wnd[1]/usr/btnSPOP-OPTION2"]


def test_cancel_on_save_prompt_is_not_gated(controller):
    _open(controller, SAVE_FIRST)
    plan = controller.plan_popup_action("cancel")
    assert plan["save_risk"] is False


def test_save_button_is_a_save_risk(controller):
    popup = {
        "window_id": "wnd[1]",
        "title": "Changes",
        "texts": ["Leave the editor?"],
        "buttons": [
            {"id": "wnd[1]/usr/btnSAVE", "text": "Save", "tooltip": ""},
            {"id": "wnd[1]/usr/btnBACK", "text": "Back", "tooltip": ""},
        ],
    }
    _open(controller, popup)
    assert controller.plan_popup_action("press", "Save")["save_risk"] is True


def test_non_save_popup_unchanged(controller):
    _open(controller, CONTINUE)
    plan = controller.plan_popup_action("confirm")
    assert plan["save_risk"] is False
    result = controller.apply_popup_plan(plan)
    assert result["action"] == "confirmed"
    assert controller.session.log == ["press wnd[1]/tbar[0]/btn[0]"]


def test_single_button_saved_notice_is_not_gated(controller):
    _open(controller, SAVED_NOTICE)
    assert controller.plan_popup_action("confirm")["save_risk"] is False


def test_enter_fallback_still_sent_on_non_save_popup(controller):
    _open(controller, NO_BUTTONS)
    result = controller.handle_popup("confirm")
    assert result["button_pressed"] == "Enter (fallback)"
    assert controller.session.log == ["vkey 0 wnd[1]"]


def test_no_enter_fallback_on_save_prompt(controller):
    popup = dict(SAVE_FIRST_DE)
    # Ja / Nein match no confirm pattern, so confirm would fall back to Enter.
    _open(controller, popup)
    result = controller.handle_popup("confirm")
    assert result["action"] == "error"
    assert "Enter was not sent" in result["error"]
    assert controller.session.log == []


def test_auto_never_confirms_a_save(controller):
    popup = {
        "window_id": "wnd[1]",
        "title": "Information",
        "texts": ["Changes were made"],
        "buttons": [{"id": "wnd[1]/usr/btnSAVE", "text": "Save", "tooltip": ""}],
    }
    _open(controller, popup)
    plan = controller.plan_popup_action("auto")
    assert plan["popup"]["auto_decision"] == "read"
    assert plan["done"] is True
    assert controller.apply_popup_plan(plan)["action"] == "read"
    assert controller.session.log == []


def test_auto_still_confirms_a_plain_notice(controller):
    _open(controller, CONTINUE)
    result = controller.handle_popup("auto")
    assert result["action"] == "confirmed"


def test_recheck_refuses_a_changed_popup(controller):
    _open(controller, SAVE_FIRST)
    plan = controller.plan_popup_action("confirm")
    controller.session.add("wnd[2]", text="Something else")  # a new popup on top
    result = controller.apply_popup_plan(plan, recheck=True)
    assert result["action"] == "error"
    assert "changed" in result["error"]
    assert controller.session.log == []


def test_recheck_passes_on_the_same_popup(controller):
    _open(controller, SAVE_FIRST)
    plan = controller.plan_popup_action("confirm")
    result = controller.apply_popup_plan(plan, recheck=True)
    assert result["action"] == "confirmed"
    assert controller.session.log == ["press wnd[1]/usr/btnSPOP-OPTION1"]


# ---------------------------------------------------------------------------
# Controller: check_vkey
# ---------------------------------------------------------------------------

def test_check_vkey_refuses_inactive_key(controller):
    controller.session.add("wnd[0]", allowed_vkeys={0, 3})
    refusal = controller.check_vkey(11)
    assert refusal is not None
    assert refusal["error"] == "F11 is not active on this screen"
    assert refusal["available_keys"] == ["Enter", "F3"]
    assert controller.session.log == []


def test_check_vkey_allows_active_key(controller):
    controller.session.add("wnd[0]", allowed_vkeys={11})
    assert controller.check_vkey(11) is None
    assert controller.session.log == []


def test_send_vkey_refusal_unchanged(controller):
    controller.session.add("wnd[0]", allowed_vkeys={0})
    result = controller.send_vkey(11)
    assert result["error"] == "F11 is not active on this screen"
    assert controller.session.log == []


# ---------------------------------------------------------------------------
# Server tools
# ---------------------------------------------------------------------------

class FakeCtx:
    def __init__(self, answer: str = "accept", data: Any = True) -> None:
        self.prompts: List[str] = []
        self._answer = answer
        self._data = data

    async def elicit(self, message: str, response_type: Any = None, **_: Any):
        self.prompts.append(message)
        return SimpleNamespace(action=self._answer, data=self._data)


@pytest.fixture
def tools(controller, monkeypatch):
    async def com(fn):
        return fn()

    monkeypatch.setattr(server, "_com", com)
    monkeypatch.setattr(server, "_ctrl", lambda ctx: controller)
    monkeypatch.setattr(server, "_session_mgr", None)
    return controller


def _run(coro):
    return asyncio.run(coro)


def test_handle_popup_asks_before_saving(tools):
    _open(tools, SAVE_FIRST)
    ctx = FakeCtx("accept", True)
    result = _run(server.sap_handle_popup(ctx, action="confirm"))
    assert len(ctx.prompts) == 1
    assert "'Yes'" in ctx.prompts[0] and "Exit Editing" in ctx.prompts[0]
    assert result["action"] == "confirmed"
    assert tools.session.log == ["press wnd[1]/usr/btnSPOP-OPTION1"]


def test_handle_popup_decline_presses_nothing(tools):
    _open(tools, SAVE_FIRST)
    ctx = FakeCtx("decline", None)
    result = _run(server.sap_handle_popup(ctx, action="confirm"))
    assert result["status"] == "cancelled"
    assert result["action"] == "none"
    assert result["buttons"]  # still there for the next step
    assert tools.session.log == []


def test_handle_popup_non_save_does_not_ask(tools):
    _open(tools, CONTINUE)
    ctx = FakeCtx()
    result = _run(server.sap_handle_popup(ctx, action="confirm"))
    assert ctx.prompts == []
    assert result["action"] == "confirmed"


def test_send_key_refused_save_is_not_asked(tools):
    tools.session.add("wnd[0]", allowed_vkeys={0})
    ctx = FakeCtx()
    result = _run(server.sap_send_key("F11", ctx))
    assert ctx.prompts == []
    assert result["error"] == "F11 is not active on this screen"
    assert tools.session.log == []


def test_send_key_active_save_is_asked_then_sent(tools):
    tools.session.add("wnd[0]", allowed_vkeys={11})
    ctx = FakeCtx()
    result = _run(server.sap_send_key("Save", ctx))
    assert len(ctx.prompts) == 1
    assert "error" not in result
    assert tools.session.log == ["vkey 11 wnd[0]"]


def test_send_key_declined_save_unchanged(tools):
    tools.session.add("wnd[0]", allowed_vkeys={11})
    ctx = FakeCtx("decline", None)
    result = _run(server.sap_send_key("F11", ctx))
    assert result == {
        "status": "cancelled",
        "action": "decline",
        "key": "F11",
        "reason": "User did not confirm save (decline)",
        "error": "Save declined by user (decline). Nothing was sent.",
    }
    assert tools.session.log == []


def test_handle_popup_save_without_elicitation_fails_closed(tools):
    from mcp.shared.exceptions import McpError
    from mcp.types import ErrorData

    class NoElicitCtx:
        async def elicit(self, message: str, response_type: Any = None, **_: Any):
            raise McpError(ErrorData(code=-32601, message="elicitation not supported"))

    _open(tools, SAVE_FIRST)
    with pytest.raises(ValueError, match="does not support elicitation"):
        _run(server.sap_handle_popup(NoElicitCtx(), action="confirm"))
    assert tools.session.log == []


# ---------------------------------------------------------------------------
# Questions the tool cannot read (SAPLSPO1 0500 with the Edge control)
# ---------------------------------------------------------------------------

# As read live on A4H, SAP GUI 8.10: the question is in an HTML viewer the
# Edge control keeps from scripts, so the read has buttons and a title only.
UNREAD = {
    "window_id": "wnd[1]",
    "title": "A4H(1)/001 Exit Editor",
    "question_readable": False,
    "buttons": [
        {"id": "wnd[1]/usr/btnBUTTON_1", "text": "Yes", "tooltip": "Yes"},
        {"id": "wnd[1]/usr/btnBUTTON_2", "text": "No", "tooltip": "No"},
        {"id": "wnd[1]/usr/btnBUTTON_3", "text": "Cancel", "tooltip": "Cancel"},
    ],
}

UNREAD_DE = {
    "window_id": "wnd[1]",
    "title": "A4H(1)/001 Editor beenden",
    "question_readable": False,
    "buttons": [
        {"id": "wnd[1]/usr/btnBUTTON_1", "text": "Ja", "tooltip": "Ja"},
        {"id": "wnd[1]/usr/btnBUTTON_2", "text": "Nein", "tooltip": "Nein"},
        {"id": "wnd[1]/usr/btnBUTTON_3", "text": "Abbrechen", "tooltip": "Abbrechen"},
    ],
}

UNREAD_NOTICE = {
    "window_id": "wnd[1]",
    "title": "Note",
    "question_readable": False,
    "buttons": [
        {"id": "wnd[1]/tbar[0]/btn[0]", "text": "", "tooltip": "Continue (Enter)"},
    ],
}


def test_confirm_on_unread_question_asks(controller):
    _open(controller, UNREAD)
    plan = controller.plan_popup_action("confirm")
    assert plan["save_risk"] is True
    assert plan["unread_question"] is True
    assert plan["button"]["text"] == "Yes"
    assert controller.session.log == []


@pytest.mark.parametrize("popup, answer, asks", [
    (UNREAD, "Yes", True),
    (UNREAD, "No", True),        # unread, No may drop the changes
    (UNREAD, "Cancel", False),
    (UNREAD_DE, "Ja", True),     # live: German answers matched no word
    (UNREAD_DE, "Nein", True),
    (UNREAD_DE, "Abbrechen", False),
])
def test_any_answer_but_cancel_to_unread_question_asks(controller, popup, answer, asks):
    _open(controller, popup)
    plan = controller.plan_popup_action("press", answer)
    assert plan["save_risk"] is asks
    assert controller.session.log == []


def test_no_enter_on_unread_question_without_a_yes_button(controller):
    """Live: confirm on Ja / Nein / Abbrechen fell back to Enter."""
    _open(controller, UNREAD_DE)
    result = controller.handle_popup("confirm")
    assert result["action"] == "error"
    assert "question cannot be read" in result["error"]
    assert "Enter was not sent" in result["error"]
    assert controller.session.log == []


def test_one_button_unread_popup_is_not_asked(controller):
    _open(controller, UNREAD_NOTICE)
    assert controller.plan_popup_action("confirm")["save_risk"] is False


@pytest.mark.parametrize("button, is_cancel", [
    ({"text": "Cancel   (F12)", "tooltip": ""}, True),
    ({"text": "", "tooltip": "Cancel   (F12)"}, True),
    ({"text": "Abbrechen", "tooltip": ""}, True),
    ({"text": "No", "tooltip": "No"}, False),
    ({"text": "Nein", "tooltip": ""}, False),
    ({"text": "Execute in Background", "tooltip": ""}, False),
])
def test_cancel_button_is_matched_as_a_word(controller, button, is_cancel):
    assert controller._is_cancel_button(button) is is_cancel


# ---- sap_press_button on a popup button ----

@pytest.mark.parametrize("button_id", [
    "wnd[1]/usr/btnBUTTON_1",
    "/app/con[0]/ses[0]/wnd[1]/usr/btnBUTTON_2",
])
def test_press_button_plan_for_an_unread_answer(controller, button_id):
    _open(controller, UNREAD)
    plan = controller.plan_popup_button(button_id)
    assert plan is not None
    assert plan["unread_question"] is True
    assert plan["button"]["id"] == button_id.split("ses[0]/")[-1]


def test_press_button_plan_for_a_readable_save_prompt(controller):
    _open(controller, SAVE_FIRST)
    plan = controller.plan_popup_button("wnd[1]/usr/btnSPOP-OPTION1")
    assert plan is not None and plan["unread_question"] is False


@pytest.mark.parametrize("popup, button_id", [
    (UNREAD, "wnd[1]/usr/btnBUTTON_3"),        # Cancel
    (CONTINUE, "wnd[1]/tbar[0]/btn[0]"),       # nothing to decide
    (UNREAD, "wnd[1]/usr/btnELSEWHERE"),       # not one of the popup's buttons
])
def test_press_button_plan_none(controller, popup, button_id):
    _open(controller, popup)
    assert controller.plan_popup_button(button_id) is None


def test_press_button_on_main_window_reads_no_popup(controller):
    _open(controller, UNREAD)
    assert controller.plan_popup_button("wnd[0]/tbar[1]/btn[8]") is None
    assert len(controller.popups) == 1  # the popup was not read


# ---- server tools ----

def test_handle_popup_asks_before_answering_unread_question(tools):
    _open(tools, UNREAD)
    ctx = FakeCtx("accept", True)
    result = _run(server.sap_handle_popup(ctx, action="confirm"))
    assert len(ctx.prompts) == 1
    assert "cannot read" in ctx.prompts[0]
    assert "'Yes'" in ctx.prompts[0] and "Exit Editor" in ctx.prompts[0]
    assert result["confirmed_by_user"] is True
    assert tools.session.log == ["press wnd[1]/usr/btnBUTTON_1"]


def test_press_button_asks_before_unread_answer(tools):
    _open(tools, UNREAD)
    ctx = FakeCtx("accept", True)
    result = _run(server.sap_press_button("wnd[1]/usr/btnBUTTON_1", ctx))
    assert len(ctx.prompts) == 1 and "cannot read" in ctx.prompts[0]
    assert result["status"] == "pressed"
    assert result["confirmed_by_user"] is True
    assert tools.session.log == ["press wnd[1]/usr/btnBUTTON_1"]


def test_press_button_decline_presses_nothing(tools):
    _open(tools, UNREAD)
    ctx = FakeCtx("decline", None)
    result = _run(server.sap_press_button("wnd[1]/usr/btnBUTTON_2", ctx))
    assert result["status"] == "cancelled"
    assert result["reason"] == "User did not confirm the answer (decline)"
    assert result["buttons"]  # still there for the next step
    assert tools.session.log == []


def test_press_button_cancel_on_unread_popup_does_not_ask(tools):
    _open(tools, UNREAD)
    ctx = FakeCtx()
    result = _run(server.sap_press_button("wnd[1]/usr/btnBUTTON_3", ctx))
    assert ctx.prompts == []
    assert "confirmed_by_user" not in result
    assert tools.session.log == ["press wnd[1]/usr/btnBUTTON_3"]


def test_press_button_refuses_when_the_popup_changed(tools):
    _open(tools, UNREAD)

    class NewPopupCtx(FakeCtx):
        async def elicit(self, message, response_type=None, **_):
            tools.session.add("wnd[2]", text="Something else")  # opened meanwhile
            return await super().elicit(message, response_type)

    result = _run(server.sap_press_button("wnd[1]/usr/btnBUTTON_1", NewPopupCtx()))
    assert "changed" in result["error"]
    assert tools.session.log == []


def test_popup_confirmation_is_audited_under_its_tool(tools, caplog):
    import json
    import logging

    _open(tools, UNREAD)
    with caplog.at_level(logging.INFO, logger="mcp_sap_gui.audit"):
        _run(server.sap_press_button("wnd[1]/usr/btnBUTTON_1", FakeCtx("decline", None)))
    events = [json.loads(r.getMessage()) for r in caplog.records
              if '"confirmation"' in r.getMessage()]
    assert [(e["tool"], e["outcome"]) for e in events] == [("sap_press_button", "declined")]


# ---------------------------------------------------------------------------
# cancel never answers "No"; a fallback key that changed nothing says so
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("popup, pressed", [
    # Live on "Exit Editor": cancel matched "No" first and left without saving.
    (UNREAD, "press wnd[1]/usr/btnBUTTON_3"),
    (SAVE_FIRST, "press wnd[1]/usr/btnSPOP-OPTION_CAN"),
    (UNREAD_DE, "press wnd[1]/usr/btnBUTTON_3"),
])
def test_cancel_presses_cancel_not_no(controller, popup, pressed):
    _open(controller, popup)
    result = controller.handle_popup("cancel")
    assert result["action"] == "canceled"
    assert controller.session.log == [pressed]


def test_cancel_without_a_cancel_button_sends_f12(controller):
    _open(controller, SAVE_FIRST_DE)  # Ja / Nein only
    result = controller.handle_popup("cancel")
    assert result["button_pressed"] == "F12 (fallback)"
    assert controller.session.log == ["vkey 12 wnd[1]"]


def test_cancel_refuses_when_f12_is_not_active_either(controller):
    _open(controller, SAVE_FIRST_DE)
    controller.session.add("wnd[1]", allowed_vkeys={0})  # F12 refused
    result = controller.handle_popup("cancel")
    assert result["action"] == "error"
    assert "no Cancel button and F12 is not active" in result["error"]
    assert controller.session.log == []


def test_fallback_key_that_changed_nothing_says_so(controller):
    _open(controller, NO_BUTTONS)
    controller.popups.append(NO_BUTTONS)  # still there after Enter
    result = controller.handle_popup("confirm")
    assert result["action"] == "no_effect"
    assert "Enter changed nothing" in result["note"]
    assert controller.session.log == ["vkey 0 wnd[1]"]


def test_fallback_key_with_a_new_message_is_not_no_effect(controller):
    _open(controller, NO_BUTTONS)
    controller.popups.append({**NO_BUTTONS, "message": "Enter a value"})
    result = controller.handle_popup("confirm")
    assert result["action"] == "confirmed"
