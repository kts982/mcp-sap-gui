"""Transaction policy: presets, glob patterns, policy file, fail-closed startup.

The policy is a guardrail against agent mistakes, not a security boundary (the
SAP user's authorizations are). These tests pin the rules it promises:
``allow`` wins over ``block``, ``strict`` starts nothing that is not allowed,
and a policy that cannot be read stops the server.
"""

import json
import subprocess
import sys
from pathlib import Path

import pytest

import mcp_sap_gui.server as srv_mod
from mcp_sap_gui.policy import (
    PRESETS,
    PolicyError,
    TransactionPolicy,
    default_policy_path,
    load_policy,
    normalize_pattern,
    resolve_policy_file,
)
from mcp_sap_gui.session_manager import SessionManager

normalize_code = srv_mod._normalize_transaction_code


@pytest.fixture
def srv():
    srv_mod._session_mgr = SessionManager()
    srv_mod.config = srv_mod.ServerConfig()
    yield srv_mod
    srv_mod._session_mgr.shutdown()
    srv_mod.config = srv_mod.ServerConfig()


def _write(tmp_path, data) -> str:
    path = tmp_path / "policy.json"
    path.write_text(data if isinstance(data, str) else json.dumps(data), encoding="utf-8")
    return str(path)


# ===========================================================================
# Rules
# ===========================================================================

class TestRules:
    def test_allow_wins_over_block(self):
        policy = TransactionPolicy(block=["S*"], allow=["SM30", "SM34"])

        assert policy.permits("SM30") is True
        assert policy.permits("SM34") is True
        assert policy.permits("SU01") is False
        assert policy.permits("MM03") is True

    def test_globs_match_namespaced_transactions(self):
        policy = TransactionPolicy(block=["/SCWM/*"], allow=["/SCWM/MON"])

        assert policy.permits("/SCWM/MON") is True
        assert policy.permits("/SCWM/RFUI") is False

    def test_question_mark_matches_one_character(self):
        policy = TransactionPolicy(block=["SM5?"])

        assert policy.permits("SM59") is False
        assert policy.permits("SM5") is True

    def test_a_plain_code_is_not_a_prefix_match(self):
        policy = TransactionPolicy(block=["SE16"])

        assert policy.permits("SE16") is False
        assert policy.permits("SE16N") is True


class TestNormalizePattern:
    def test_plain_codes_keep_the_prefix_stripping(self):
        assert normalize_pattern(" /nse80 ", normalize_code) == "SE80"

    def test_patterns_are_uppercased(self):
        assert normalize_pattern(" su* ", normalize_code) == "SU*"
        assert normalize_pattern("/scwm/*", normalize_code) == "/SCWM/*"

    @pytest.mark.parametrize("entry", ["", "   ", None, 5, "SU*;DROP", "S[U]*"])
    def test_garbage_is_rejected(self, entry):
        with pytest.raises(PolicyError):
            normalize_pattern(entry, normalize_code)


# ===========================================================================
# Presets
# ===========================================================================

class TestPresets:
    def test_default_blocks_basis_security_and_development(self):
        policy = load_policy(normalize_code)

        for tcode in ("SU01", "PFCG", "SE38", "SA38", "SE80", "SE37", "STMS",
                      "SM59", "STRUST", "SICF", "SM01", "SM18", "SM19"):
            assert policy.permits(tcode) is False, tcode
        # The customizing and display work this server exists for stays open.
        for tcode in ("SM30", "SM34", "SPRO", "SE16", "SE11", "MM03", "/SCWM/MON"):
            assert policy.permits(tcode) is True, tcode

    def test_abap_dev_opens_report_execution_and_the_repository(self):
        policy = load_policy(normalize_code, preset="abap-dev")

        for tcode in ("SA38", "SE11", "SE80"):
            assert policy.permits(tcode) is True, tcode
        for tcode in ("SE38", "SU01", "STMS", "SM59"):
            assert policy.permits(tcode) is False, tcode

    def test_strict_starts_nothing_by_itself(self):
        policy = load_policy(normalize_code, preset="strict")

        for tcode in ("MM03", "SM30", "SU01", "/SCWM/MON"):
            assert policy.permits(tcode) is False, tcode

    def test_unknown_preset_is_an_error(self):
        with pytest.raises(PolicyError, match="Unknown policy preset"):
            load_policy(normalize_code, preset="relaxed")

    def test_presets_do_not_share_list_objects(self):
        assert PRESETS["default"]["block"] is not PRESETS["abap-dev"]["block"]


# ===========================================================================
# Policy file
# ===========================================================================

class TestPolicyFile:
    def test_file_adds_to_the_preset(self, tmp_path):
        path = _write(tmp_path, {"block": ["zhr*"], "allow": ["sm59"]})

        policy = load_policy(normalize_code, policy_file=path)

        assert policy.permits("ZHR_PAYROLL") is False
        assert policy.permits("SM59") is True       # unblocked by the file
        assert policy.permits("SU01") is False      # preset still applies
        assert policy.source == path
        assert len(policy.file_sha256) == 64

    def test_strict_with_an_allowlist(self, tmp_path):
        path = _write(tmp_path, {"preset": "strict", "allow": ["MM03", "/SCWM/*"]})

        policy = load_policy(normalize_code, policy_file=path)

        assert policy.permits("MM03") is True
        assert policy.permits("/SCWM/MON") is True
        assert policy.permits("VA01") is False

    def test_command_line_preset_wins_over_the_file(self, tmp_path):
        path = _write(tmp_path, {"preset": "abap-dev"})

        policy = load_policy(normalize_code, preset="default", policy_file=path)

        assert policy.preset == "default"
        assert policy.permits("SA38") is False

    def test_utf8_bom_is_tolerated(self, tmp_path):
        path = tmp_path / "policy.json"
        path.write_bytes(b"\xef\xbb\xbf" + json.dumps({"allow": ["SM59"]}).encode())

        assert load_policy(normalize_code, policy_file=str(path)).permits("SM59")

    @pytest.mark.parametrize("content, message", [
        ("{not json", "Cannot read policy file"),
        ("[]", "must contain a JSON object"),
        ({"blok": ["SU01"]}, "Unknown keys"),
        ({"block": "SU01"}, "must be a list"),
        ({"block": ["SU*;x"]}, "Invalid transaction pattern"),
        ({"preset": "relaxed"}, "Unknown policy preset"),
    ])
    def test_invalid_files_are_rejected(self, tmp_path, content, message):
        with pytest.raises(PolicyError, match=message):
            load_policy(normalize_code, policy_file=_write(tmp_path, content))


class TestShippedExamples:
    """The files in examples/ are what users copy: they must stay loadable."""

    EXAMPLES = sorted(
        (Path(__file__).parent.parent / "examples").glob("policy*.json")
    )

    def test_examples_exist(self):
        assert [p.name for p in self.EXAMPLES] == [
            "policy.example.json", "policy.strict.example.json",
        ]

    @pytest.mark.parametrize("path", EXAMPLES, ids=lambda p: p.name)
    def test_example_loads(self, path):
        policy = load_policy(normalize_code, policy_file=str(path))

        assert policy.preset in PRESETS

    def test_default_example_does_what_the_readme_says(self):
        policy = load_policy(normalize_code, policy_file=str(self.EXAMPLES[0]))

        assert policy.permits("ST22") is True          # allowed on top of default
        assert policy.permits("SE16") is False         # blocked on top of default
        assert policy.permits("ZHR_PAYROLL") is False
        assert policy.permits("SU01") is False         # the preset still applies
        assert policy.permits("SM30") is True

    def test_strict_example_is_an_allowlist(self):
        policy = load_policy(normalize_code, policy_file=str(self.EXAMPLES[1]))

        assert policy.permits("MM03") is True
        assert policy.permits("/SCWM/MON") is True
        assert policy.permits("VA01") is False
        assert policy.permits("SE16") is False


class TestResolvePolicyFile:
    def test_explicit_path_must_exist(self, tmp_path, monkeypatch):
        monkeypatch.delenv("MCP_SAP_GUI_POLICY_FILE", raising=False)
        with pytest.raises(PolicyError, match="not found"):
            resolve_policy_file(str(tmp_path / "missing.json"))

    def test_env_var_is_used_when_no_flag_is_given(self, tmp_path, monkeypatch):
        path = _write(tmp_path, {})
        monkeypatch.setenv("MCP_SAP_GUI_POLICY_FILE", path)

        assert resolve_policy_file(None) == path

    def test_flag_wins_over_env_var(self, tmp_path, monkeypatch):
        flagged = _write(tmp_path, {})
        monkeypatch.setenv("MCP_SAP_GUI_POLICY_FILE", str(tmp_path / "other.json"))

        assert resolve_policy_file(flagged) == flagged

    def test_per_user_default_is_picked_up_when_present(self, tmp_path, monkeypatch):
        monkeypatch.delenv("MCP_SAP_GUI_POLICY_FILE", raising=False)
        monkeypatch.setenv("APPDATA", str(tmp_path))
        assert resolve_policy_file(None) is None

        target = tmp_path / "mcp-sap-gui" / "policy.json"
        target.parent.mkdir()
        target.write_text("{}", encoding="utf-8")

        assert default_policy_path() == str(target)
        assert resolve_policy_file(None) == str(target)


# ===========================================================================
# Server integration
# ===========================================================================

class TestServerIntegration:
    def test_config_accepts_glob_patterns(self, srv):
        srv.config = srv.ServerConfig(
            blocked_transactions=["su*", " /nse80 "], allowed_exceptions=["su3"],
        )

        assert srv.config.blocked_transactions == ["SU*", "SE80"]
        assert srv._is_transaction_blocked("SU01") is True
        assert srv._is_transaction_blocked("/nSU3") is False
        assert srv._is_transaction_blocked("SE80") is True

    def test_legacy_allowlist_still_restricts_on_top(self, srv):
        srv.config = srv.ServerConfig(
            allowed_transactions=["MM03", "SA38"], allowed_exceptions=["SA38"],
        )

        assert srv._is_transaction_blocked("MM03") is False
        assert srv._is_transaction_blocked("SA38") is False
        assert srv._is_transaction_blocked("VA03") is True

    def test_blocked_error_tells_the_user_what_to_do(self, srv):
        with pytest.raises(ValueError) as excinfo:
            srv._enforce_transaction_policy("SA38")

        text = str(excinfo.value)
        assert "Transaction SA38 is blocked by security policy" in text
        assert "preset 'default'" in text
        assert "--policy-file" in text and "--policy-preset" in text
        assert "tell the user" in text

    def test_okcode_bypass_uses_the_same_policy(self, srv):
        srv.config = srv.ServerConfig(blocked_transactions=["*"], allowed_exceptions=["MM03"])

        srv._check_okcode_bypass("wnd[0]/tbar[0]/okcd", "/nMM03")
        with pytest.raises(ValueError, match="blocked by security policy"):
            srv._check_okcode_bypass("wnd[0]/tbar[0]/okcd", "/nVA01")

    def test_bare_n_is_never_blocked_even_under_strict(self, srv):
        srv.config = srv.ServerConfig(blocked_transactions=["*"])

        assert srv._enforce_transaction_policy("/n") == "/N"

    def test_no_tool_can_change_the_policy(self, srv):
        """The policy is startup-only: an agent must not be able to widen it."""
        import asyncio
        names = {tool.name for tool in asyncio.run(srv.mcp.list_tools())}

        assert not [n for n in names if "policy_file" in n or "transaction_policy" in n]


class TestStartupFailsClosed:
    def _run(self, *args):
        return subprocess.run(
            [sys.executable, "-m", "mcp_sap_gui.server", *args],
            capture_output=True, text=True, timeout=60,
        )

    def test_missing_policy_file_stops_the_server(self, tmp_path):
        result = self._run("--policy-file", str(tmp_path / "missing.json"))

        assert result.returncode == 2
        assert "transaction policy" in result.stderr
        assert "not found" in result.stderr

    def test_invalid_policy_file_stops_the_server(self, tmp_path):
        result = self._run("--policy-file", _write(tmp_path, {"block": "SU01"}))

        assert result.returncode == 2
        assert "must be a list" in result.stderr
