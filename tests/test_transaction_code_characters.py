"""Transaction codes with characters other than letters, digits and '_'.

SAP's own codes use most of the keyboard: F-02, F.13, BRF+, FF$3, MC=1. The
policy has to recognise them, or a code typed into the command field without
a prefix is taken for a function code and not checked at all.

No SAP GUI needed: the policy and the guard are pure functions.
"""

import asyncio
import json

import pytest

from mcp_sap_gui import server
from mcp_sap_gui.policy import PolicyError, TransactionPolicy, load_policy
from mcp_sap_gui.server import (
    ServerConfig,
    _check_okcode_bypass,
    _enforce_transaction_policy,
    _normalize_transaction_code,
    precheck_before_confirmation,
)

OKCD = "wnd[0]/tbar[0]/okcd"

# Codes from TSTC of an S/4HANA system, one for each character found there.
REAL_CODES = [
    "F-02", "F.13", "BRF+", "FF$3", "J3G&", "MC=1", "MCS%", "MC1!", "FF:1",
    "MC$<", "MC$>", "J3G#", "J3G(", "MCB)", "OMO,", "MC?0", "F9C[",
    "MCM-", "J3G.",  # the character can be the last one
]


def _load(tmp_path, **file_data):
    policy_file = tmp_path / "policy.json"
    policy_file.write_text(json.dumps(file_data), encoding="utf-8")
    return load_policy(_normalize_transaction_code, policy_file=str(policy_file))


def _use_policy(monkeypatch, tmp_path, **file_data):
    """Install the effective policy the way main() builds it."""
    policy = _load(tmp_path, **file_data)
    monkeypatch.setattr(server, "config", ServerConfig(
        blocked_transactions=policy.block,
        allowed_exceptions=policy.allow,
        policy_preset=policy.preset,
    ))


@pytest.fixture
def fi_blocked(monkeypatch, tmp_path):
    _use_policy(monkeypatch, tmp_path, preset="default", block=["F-*", "F.*"])


@pytest.fixture
def strict(monkeypatch, tmp_path):
    _use_policy(monkeypatch, tmp_path, preset="strict", allow=["VA03"])


# --- the normalizer ------------------------------------------------------------

@pytest.mark.parametrize("code", REAL_CODES)
def test_real_codes_are_transaction_codes(code):
    assert _normalize_transaction_code(code) == code
    assert _normalize_transaction_code(f"/n{code.lower()}") == code


@pytest.mark.parametrize("value", [
    "%PC", "&ALL", "&F03", "$X", "-F02", ".F13", "+",  # start with another character
    "F 02", "F;02", "F*2",                              # a space, ';' or '*' is in no code
])
def test_function_codes_and_broken_values_are_not_codes(value):
    assert _normalize_transaction_code(value, strict=False) is None


def test_the_error_for_an_invalid_code_shows_such_a_code():
    with pytest.raises(ValueError, match="'F-02'"):
        _normalize_transaction_code("F 02")


# --- the command field ---------------------------------------------------------

@pytest.mark.parametrize("value, code", [
    ("F-02", "F-02"),       # typed bare: on SAP Easy Access this starts F-02
    ("f-02", "F-02"),
    ("  F-02 ", "F-02"),
    ("=F-02", "F-02"),
    ("F.13", "F.13"),
    ("/nF-02", "F-02"),
    ("/oF-28", "F-28"),
    ("*F-02", "F-02"),
    ("/*F-02 BKPF-BUKRS=1000", "F-02"),
])
def test_a_blocked_code_is_refused_in_the_command_field(fi_blocked, value, code):
    with pytest.raises(ValueError, match=rf"Transaction {code} is blocked.*command field"):
        _check_okcode_bypass(OKCD, value)


@pytest.mark.parametrize("value", ["FB01", "MC=1", "%PC", "&ALL", "P+", "P--"])
def test_other_values_pass_a_deny_list(fi_blocked, value):
    _check_okcode_bypass(OKCD, value)


@pytest.mark.parametrize("value", ["F-02", "F.13", "BRF+", "P+", "P--", "BACK"])
def test_strict_refuses_every_value_that_can_be_a_code(strict, value):
    with pytest.raises(ValueError, match="is blocked by security policy"):
        _check_okcode_bypass(OKCD, value)


@pytest.mark.parametrize("value", ["%PC", "%SC", "&ALL", "&F03", "VA03"])
def test_strict_leaves_function_codes_and_allowed_codes(strict, value):
    _check_okcode_bypass(OKCD, value)


def test_strict_can_allow_such_codes(monkeypatch, tmp_path):
    _use_policy(monkeypatch, tmp_path, preset="strict", allow=["F-02", "F.1?", "P+"])

    for value in ("F-02", "/nF-02", "F.13", "P+"):
        _check_okcode_bypass(OKCD, value)
    with pytest.raises(ValueError, match="Transaction F-28 is blocked"):
        _check_okcode_bypass(OKCD, "F-28")


def test_the_confirmation_precheck_refuses_it_too(fi_blocked):
    with pytest.raises(server.ToolError, match="sap_set_field.*F-02 is blocked"):
        precheck_before_confirmation("sap_set_field", {"field_id": OKCD, "value": "F-02"})
    with pytest.raises(server.ToolError, match="sap_set_batch_fields.*F.13 is blocked"):
        precheck_before_confirmation("sap_set_batch_fields", {"fields": {OKCD: "F.13"}})


# --- sap_execute_transaction -----------------------------------------------------

def test_execute_transaction_takes_such_a_code(monkeypatch, tmp_path):
    _use_policy(monkeypatch, tmp_path, preset="default")

    assert _enforce_transaction_policy("F-02") == "F-02"
    assert _enforce_transaction_policy("/nf.13") == "F.13"


def test_execute_transaction_refuses_a_blocked_one(fi_blocked):
    with pytest.raises(ValueError, match="Transaction F-02 is blocked by security policy"):
        _enforce_transaction_policy("F-02")


def test_the_tool_hands_the_code_to_the_controller(monkeypatch, tmp_path):
    _use_policy(monkeypatch, tmp_path, preset="default")
    started = []

    class Controller:
        def execute_transaction(self, tcode):
            started.append(tcode)
            return {"transaction": tcode}

    async def run_inline(fn):
        return fn()

    monkeypatch.setattr(server, "_ctrl", lambda ctx: Controller())
    monkeypatch.setattr(server, "_com", run_inline)
    tool = getattr(server.sap_execute_transaction, "fn", server.sap_execute_transaction)

    assert asyncio.run(tool("F-02", None)) == {"transaction": "F-02"}
    assert started == ["F-02"]


# --- the policy file -------------------------------------------------------------

@pytest.mark.parametrize("entry", ["F-02", "f-02", "F-*", "F.1?", "MC?0", "F9C[", "*+"])
def test_the_policy_file_takes_such_entries(tmp_path, entry):
    policy = _load(tmp_path, preset="default", block=[entry])

    assert entry.upper() in policy.block


@pytest.mark.parametrize("entry", ["F 02", "F;02", "F-* X", "%PC"])
def test_the_policy_file_refuses_what_is_no_code_and_no_pattern(tmp_path, entry):
    with pytest.raises(PolicyError):
        _load(tmp_path, preset="default", block=[entry])


@pytest.mark.parametrize("entry", ["S[UE]01", "S[U]*", "F9C]"])
def test_the_policy_file_refuses_a_set_of_characters(tmp_path, entry):
    # As a literal it would block nothing the author meant, without a word.
    with pytest.raises(PolicyError, match="not supported"):
        _load(tmp_path, preset="default", block=[entry])


def test_patterns_match_these_codes(tmp_path):
    policy = _load(tmp_path, preset="default", block=["F-*", "MC?0", "*+"])

    assert not policy.permits("F-02")
    assert not policy.permits("MC?0")
    assert not policy.permits("MC10")   # '?' stays the wildcard for one character
    assert not policy.permits("BRF+")
    assert policy.permits("F.13")
    assert policy.permits("FB01")


def test_a_bracket_in_an_entry_stands_for_itself(tmp_path):
    policy = _load(tmp_path, preset="default", block=["F9C["])

    assert not policy.permits("F9C[")
    assert policy.permits("F9C")

    # Matching itself never reads a set, also for a policy built in code.
    built = TransactionPolicy(block=["A[BC]*"])
    assert not built.permits("A[BC]1")
    assert built.permits("AB1")
