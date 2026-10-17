"""Session commands (/nex, /nend, /i) and the command-field guard.

No SAP GUI needed: the policy and the guard are pure functions, and the one
tool test replaces the controller and the COM thread with fakes.
"""

import asyncio
import json

import pytest

from mcp_sap_gui import server
from mcp_sap_gui.policy import PRESETS, SESSION_COMMANDS, load_policy
from mcp_sap_gui.server import (
    ServerConfig,
    _check_okcode_bypass,
    _enforce_transaction_policy,
    _normalize_transaction_code,
    precheck_before_confirmation,
)

OKCD = "wnd[0]/tbar[0]/okcd"


def _use_policy(monkeypatch, tmp_path=None, **file_data):
    """Install the effective policy the way main() builds it."""
    policy_file = None
    if file_data:
        policy_file = tmp_path / "policy.json"
        policy_file.write_text(json.dumps(file_data), encoding="utf-8")
    policy = load_policy(
        _normalize_transaction_code,
        policy_file=str(policy_file) if policy_file else None,
    )
    monkeypatch.setattr(server, "config", ServerConfig(
        blocked_transactions=policy.block,
        allowed_exceptions=policy.allow,
        policy_preset=policy.preset,
    ))


@pytest.fixture
def default_policy(monkeypatch):
    _use_policy(monkeypatch)


# --- the normalizer turns session commands into code-shaped values ----------

@pytest.mark.parametrize("raw, canonical", [
    ("/nex", "EX"), ("/NEX", "EX"), ("/nend", "END"), ("/i", "/I"),
])
def test_normalizer_gives_session_command_forms(raw, canonical):
    assert _normalize_transaction_code(raw) == canonical


# --- presets -----------------------------------------------------------------

@pytest.mark.parametrize("preset", sorted(PRESETS))
@pytest.mark.parametrize("command", sorted(SESSION_COMMANDS))
def test_every_preset_blocks_session_commands(preset, command):
    policy = load_policy(_normalize_transaction_code, preset=preset)
    assert not policy.permits(command)


def test_debugger_command_is_not_blocked_by_default():
    # /h is deliberately left out: it starts the debugger, it ends nothing.
    policy = load_policy(_normalize_transaction_code)
    assert policy.permits(_normalize_transaction_code("/h"))


def test_policy_file_can_allow_a_session_command(monkeypatch, tmp_path):
    _use_policy(monkeypatch, tmp_path, allow=["EX"])
    assert _enforce_transaction_policy("/nex") == "EX"
    with pytest.raises(ValueError, match="session command"):
        _enforce_transaction_policy("/nend")


# --- both routes refuse them --------------------------------------------------

@pytest.mark.parametrize("raw", ["/nex", "/nend", "/i", "EX", "/n END"])
def test_execute_transaction_route_refuses(default_policy, raw):
    with pytest.raises(ValueError, match="is a session command, not a transaction"):
        _enforce_transaction_policy(raw)


@pytest.mark.parametrize("raw", ["/nex", "/nend", "/i"])
def test_command_field_route_refuses(default_policy, raw):
    with pytest.raises(ValueError, match=r"session command.*attempted via command field"):
        _check_okcode_bypass(OKCD, raw)


def test_ordinary_transactions_unchanged(default_policy):
    assert _enforce_transaction_policy("/nMM03") == "MM03"
    assert _enforce_transaction_policy("/n") == "/N"
    with pytest.raises(ValueError, match="Transaction SU01 is blocked"):
        _enforce_transaction_policy("/nSU01")


# --- command field: fail closed only for transaction prefixes ---------------

@pytest.mark.parametrize("value", [
    "%PC", "%SC", "&ALL", "&F03", "BACK", "=SAVE", "/n", "/o", "", "  ",
    "/*MM03 RMMG1-MATNR=4711",  # unparseable as a whole, allowed code
    "*MM03",
    "/$SYNC",  # no transaction prefix: not this change's business
])
def test_command_field_values_that_pass(default_policy, value):
    _check_okcode_bypass(OKCD, value)


@pytest.mark.parametrize("value, message", [
    ("/*SU01 USR02-BNAME=X", "Transaction SU01 is blocked"),
    ("*SU01", "Transaction SU01 is blocked"),
    ("= /nSU01 X", "Transaction SU01 is blocked"),
    ("/n&ALL", "not a transaction code"),
    ("/*$X", "not a transaction code"),
])
def test_command_field_values_that_fail_closed(default_policy, value, message):
    with pytest.raises(ValueError, match=message):
        _check_okcode_bypass(OKCD, value)


def test_other_fields_are_not_checked(default_policy):
    _check_okcode_bypass("wnd[0]/usr/txtRMMG1-MATNR", "/nex")
    _check_okcode_bypass("wnd[0]/usr/cntlEDITOR/shellcont/shell", "*SU01 note")


# --- sap_set_textedit on the command field ------------------------------------

class _FakeController:
    def __init__(self):
        self.calls = []

    def set_textedit(self, textedit_id, text):
        self.calls.append((textedit_id, text))
        return {"status": "ok"}


@pytest.fixture
def fake_controller(monkeypatch):
    controller = _FakeController()

    async def run_inline(fn):
        return fn()

    monkeypatch.setattr(server, "_ctrl", lambda ctx: controller)
    monkeypatch.setattr(server, "_com", run_inline)
    return controller


def _set_textedit(textedit_id, text):
    tool = getattr(server.sap_set_textedit, "fn", server.sap_set_textedit)
    return asyncio.run(tool(textedit_id, text, None))


def test_set_textedit_refuses_blocked_transaction_in_command_field(
        default_policy, fake_controller):
    with pytest.raises(ValueError, match="SU01 is blocked.*command field"):
        _set_textedit(OKCD, "/nSU01")
    with pytest.raises(ValueError, match="session command"):
        _set_textedit(OKCD, "/nex")
    assert fake_controller.calls == []


def test_set_textedit_still_writes_text_editors(default_policy, fake_controller):
    editor = "wnd[0]/usr/cntlEDITOR/shellcont/shell"
    assert _set_textedit(editor, "/nSU01 is mentioned in this note") == {"status": "ok"}
    assert fake_controller.calls == [(editor, "/nSU01 is mentioned in this note")]


def test_confirmation_precheck_covers_set_textedit(default_policy):
    with pytest.raises(server.ToolError, match="sap_set_textedit.*SU01 is blocked"):
        precheck_before_confirmation(
            "sap_set_textedit", {"textedit_id": OKCD, "text": "/nSU01"})
    precheck_before_confirmation(
        "sap_set_textedit", {"textedit_id": OKCD, "text": "%PC"})
