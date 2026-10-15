"""Tests for the GetObjectTree switch, its logging and its trace file."""

import ast
import json
import logging
import os
import sys
from pathlib import Path

import pytest

from mcp_sap_gui import controller as controller_module
from mcp_sap_gui.controller import OBJECT_TREE_SITES, object_tree_disabled_sites
from mcp_sap_gui.sap_controller import SAPGUIController

SRC = Path(controller_module.__file__).resolve().parent

_RPC_S_SERVER_UNAVAILABLE = -2147023174
_DISP_E_UNKNOWNNAME = -2147352570

_TREE = {"children": [{"properties": {"Id": "/app/con[0]/ses[0]/wnd[1]"},
                       "children": []}]}


class FakeComError(Exception):
    """Shaped like pywintypes.com_error: hresult attribute and args[0]."""

    def __init__(self, hresult, text="failed"):
        super().__init__(hresult, text, None, None)
        self.hresult = hresult


class FakeSession:
    def __init__(self, result=None, error=None, on_call=None):
        self.calls = []
        self._result = json.dumps(_TREE) if result is None else result
        self._error = error
        self._on_call = on_call

    def GetObjectTree(self, root_id, props):  # noqa: N802 - COM name
        self.calls.append((root_id, list(props)))
        if self._on_call:
            self._on_call()
        if self._error:
            raise self._error
        return self._result


def _controller(session):
    # No __init__: it imports pywin32, which these tests do not need.
    ctrl = SAPGUIController.__new__(SAPGUIController)
    ctrl._session = session
    return ctrl


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch):
    monkeypatch.delenv("MCP_SAP_GUI_DISABLE_OBJECT_TREE", raising=False)
    monkeypatch.delenv("MCP_SAP_GUI_TRACE_FILE", raising=False)


class TestDisabledSites:
    def test_unset_and_empty_disable_nothing(self):
        assert object_tree_disabled_sites() == frozenset()
        assert object_tree_disabled_sites("") == frozenset()
        assert object_tree_disabled_sites("  ") == frozenset()

    @pytest.mark.parametrize("value", ["1", "true", "TRUE", "yes", "on", "all", " All "])
    def test_all_values_disable_every_site(self, value):
        assert object_tree_disabled_sites(value) == frozenset(OBJECT_TREE_SITES)

    def test_comma_list_is_stripped_and_lower_cased(self):
        assert object_tree_disabled_sites(" Table , popup ,") == {"table", "popup"}

    def test_unknown_site_raises_and_names_it(self):
        with pytest.raises(ValueError, match="tabel"):
            object_tree_disabled_sites("tabel,popup")

    def test_reads_the_environment(self, monkeypatch):
        monkeypatch.setenv("MCP_SAP_GUI_DISABLE_OBJECT_TREE", "list")
        assert object_tree_disabled_sites() == {"list"}


class TestObjectTreeSwitch:
    def test_unset_calls_get_object_tree(self):
        session = FakeSession()
        tree = _controller(session)._object_tree("wnd[1]", ["Id"], site="popup")
        assert tree == _TREE["children"][0]
        assert session.calls == [("wnd[1]", ["Id"])]

    def test_selected_is_refused_before_the_call(self):
        # Asking for Selected crashes SAP GUI 8.00 (issue #22).
        session = FakeSession()
        with pytest.raises(ValueError, match="Selected"):
            _controller(session)._object_tree("wnd[0]/usr", ["Id", "Selected"], site="list")
        assert session.calls == []

    def test_disabled_site_skips_the_call(self, monkeypatch):
        monkeypatch.setenv("MCP_SAP_GUI_DISABLE_OBJECT_TREE", "table")
        session = FakeSession()
        ctrl = _controller(session)
        assert ctrl._object_tree("wnd[0]/usr/tblT", ["Id"], site="table") is None
        assert session.calls == []
        assert ctrl._object_tree("wnd[1]", ["Id"], site="popup") is not None
        assert session.calls == [("wnd[1]", ["Id"])]

    def test_all_skips_every_site(self, monkeypatch):
        monkeypatch.setenv("MCP_SAP_GUI_DISABLE_OBJECT_TREE", "1")
        session = FakeSession()
        ctrl = _controller(session)
        for site in OBJECT_TREE_SITES:
            assert ctrl._object_tree("wnd[0]/usr", ["Id"], site=site) is None
        assert session.calls == []

    def test_site_is_required_and_checked(self):
        ctrl = _controller(FakeSession())
        with pytest.raises(TypeError):
            ctrl._object_tree("wnd[1]", ["Id"])
        with pytest.raises(ValueError):
            ctrl._object_tree("wnd[1]", ["Id"], site="discovery")

    def test_table_read_falls_back_when_table_site_is_off(self, monkeypatch):
        monkeypatch.setenv("MCP_SAP_GUI_DISABLE_OBJECT_TREE", "table")
        session = FakeSession()
        assert _controller(session)._table_control_values("wnd[0]/usr/tblT") is None
        assert session.calls == []


class TestFailureLogging:
    def test_other_failure_is_a_warning_with_hresult(self, caplog):
        session = FakeSession(error=FakeComError(_RPC_S_SERVER_UNAVAILABLE))
        with caplog.at_level(logging.DEBUG, logger="mcp_sap_gui.object_tree"):
            result = _controller(session)._object_tree("wnd[0]/usr/tblT", ["Id"],
                                                       site="table")
        assert result is None
        warnings = [r for r in caplog.records if r.levelno == logging.WARNING]
        assert len(warnings) == 1
        assert warnings[0].name == "mcp_sap_gui.object_tree"
        message = warnings[0].getMessage()
        assert f"hresult={_RPC_S_SERVER_UNAVAILABLE}" in message
        assert "site=table" in message and "root=wnd[0]/usr/tblT" in message

    @pytest.mark.parametrize("error", [
        AttributeError("GetObjectTree"),
        FakeComError(_DISP_E_UNKNOWNNAME),
        Exception(-2147352573, "Member not found.", None, None),
    ])
    def test_not_supported_is_debug_only(self, caplog, error):
        session = FakeSession(error=error)
        with caplog.at_level(logging.DEBUG, logger="mcp_sap_gui.object_tree"):
            assert _controller(session)._object_tree("wnd[1]", ["Id"], site="popup") is None
        assert not [r for r in caplog.records if r.levelno >= logging.WARNING]
        assert any("not available" in r.getMessage() for r in caplog.records)

    def test_success_logs_nothing_above_debug(self, caplog):
        with caplog.at_level(logging.DEBUG, logger="mcp_sap_gui.object_tree"):
            _controller(FakeSession())._object_tree("wnd[1]", ["Id"], site="popup")
        levels = [r.levelno for r in caplog.records if r.name == "mcp_sap_gui.object_tree"]
        assert levels == [logging.DEBUG, logging.DEBUG]


class TestTraceFile:
    def _lines(self, path):
        return path.read_text(encoding="utf-8").splitlines()

    def test_no_file_without_the_variable(self, tmp_path):
        _controller(FakeSession())._object_tree("wnd[1]", ["Id"], site="popup")
        assert list(tmp_path.iterdir()) == []

    def test_start_is_on_disk_before_the_call(self, tmp_path, monkeypatch):
        trace = tmp_path / "trace.txt"
        monkeypatch.setenv("MCP_SAP_GUI_TRACE_FILE", str(trace))
        seen = []
        session = FakeSession(on_call=lambda: seen.extend(self._lines(trace)))
        _controller(session)._object_tree("wnd[0]/usr/tblT", ["Id", "Text"], site="table")
        assert len(seen) == 1
        fields = seen[0].split()
        assert fields[1:5] == [str(os.getpid()), "start", "table", "wnd[0]/usr/tblT"]
        assert fields[5] == "props=2"

    def test_done_line_after_success(self, tmp_path, monkeypatch):
        trace = tmp_path / "trace.txt"
        monkeypatch.setenv("MCP_SAP_GUI_TRACE_FILE", str(trace))
        raw = json.dumps(_TREE)
        _controller(FakeSession(result=raw))._object_tree("wnd[1]", ["Id"], site="popup")
        start, done = self._lines(trace)
        assert start.split()[2] == "start"
        assert done.split()[2:5] == ["done", "popup", "wnd[1]"]
        assert done.endswith(f"bytes={len(raw)}")

    def test_failed_line_carries_error_and_hresult(self, tmp_path, monkeypatch):
        trace = tmp_path / "trace.txt"
        monkeypatch.setenv("MCP_SAP_GUI_TRACE_FILE", str(trace))
        session = FakeSession(error=FakeComError(_RPC_S_SERVER_UNAVAILABLE))
        _controller(session)._object_tree("wnd[0]/usr/tblT", ["Id"], site="table")
        start, failed = self._lines(trace)
        assert failed.split()[2:5] == ["failed", "table", "wnd[0]/usr/tblT"]
        assert "error=FakeComError" in failed
        assert f"hresult={_RPC_S_SERVER_UNAVAILABLE}" in failed

    def test_disabled_site_writes_nothing(self, tmp_path, monkeypatch):
        trace = tmp_path / "trace.txt"
        monkeypatch.setenv("MCP_SAP_GUI_TRACE_FILE", str(trace))
        monkeypatch.setenv("MCP_SAP_GUI_DISABLE_OBJECT_TREE", "all")
        _controller(FakeSession())._object_tree("wnd[1]", ["Id"], site="popup")
        assert not trace.exists()

    def test_unwritable_trace_file_does_not_break_the_read(self, tmp_path, monkeypatch):
        monkeypatch.setenv("MCP_SAP_GUI_TRACE_FILE", str(tmp_path))  # a directory
        tree = _controller(FakeSession())._object_tree("wnd[1]", ["Id"], site="popup")
        assert tree == _TREE["children"][0]


def _object_tree_calls():
    for path in sorted(SRC.glob("*.py")):
        module = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(module):
            if (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
                    and node.func.attr == "_object_tree"):
                yield path.name, node


def test_every_call_site_names_a_known_site():
    sites = []
    for name, call in _object_tree_calls():
        site = [k.value for k in call.keywords if k.arg == "site"]
        assert len(site) == 1, f"{name}:{call.lineno} passes no site="
        assert isinstance(site[0], ast.Constant), f"{name}:{call.lineno} site is not a literal"
        assert site[0].value in OBJECT_TREE_SITES, f"{name}:{call.lineno} {site[0].value!r}"
        sites.append(site[0].value)
    assert sorted(sites) == sorted(OBJECT_TREE_SITES)


class TestServerStartup:
    @pytest.fixture
    def server(self):
        pytest.importorskip("fastmcp")
        from mcp_sap_gui import server
        return server

    def test_settings_unset(self, server):
        assert server._object_tree_settings() == {"disabled_sites": [], "trace_file": None}

    def test_settings_report_sites_and_trace(self, server, tmp_path, monkeypatch):
        trace = tmp_path / "trace.txt"
        monkeypatch.setenv("MCP_SAP_GUI_DISABLE_OBJECT_TREE", "popup,table")
        monkeypatch.setenv("MCP_SAP_GUI_TRACE_FILE", str(trace))
        assert server._object_tree_settings() == {
            "disabled_sites": ["popup", "table"], "trace_file": str(trace)}

    def test_unwritable_trace_file_raises(self, server, tmp_path, monkeypatch):
        monkeypatch.setenv("MCP_SAP_GUI_TRACE_FILE", str(tmp_path))
        with pytest.raises(ValueError, match="MCP_SAP_GUI_TRACE_FILE"):
            server._object_tree_settings()

    def _isolate_main(self, server, tmp_path, monkeypatch, argv):
        monkeypatch.setenv("APPDATA", str(tmp_path))  # no per-user policy.json
        monkeypatch.delenv("MCP_SAP_GUI_POLICY_FILE", raising=False)
        monkeypatch.setattr(server, "load_dotenv", lambda *a, **k: None)
        monkeypatch.setattr(sys, "argv", ["mcp-sap-gui", *argv])
        ran = []
        monkeypatch.setattr(server.mcp, "run", lambda *a, **k: ran.append(True))
        return ran

    def test_main_refuses_an_unknown_site(self, server, tmp_path, monkeypatch, capsys):
        monkeypatch.setenv("MCP_SAP_GUI_DISABLE_OBJECT_TREE", "tabel")
        ran = self._isolate_main(server, tmp_path, monkeypatch, [])
        with pytest.raises(SystemExit) as exc_info:
            server.main()
        assert exc_info.value.code == 2
        assert "tabel" in capsys.readouterr().err
        assert ran == []

    def test_main_writes_the_audit_event(self, server, tmp_path, monkeypatch):
        audit = tmp_path / "audit.jsonl"
        monkeypatch.setenv("MCP_SAP_GUI_DISABLE_OBJECT_TREE", "table")
        ran = self._isolate_main(server, tmp_path, monkeypatch, ["--audit-log", str(audit)])
        audit_logger = logging.getLogger("mcp_sap_gui.audit")
        handlers = list(audit_logger.handlers)
        try:
            server.main()
        finally:
            for handler in audit_logger.handlers[:]:
                if handler not in handlers:
                    audit_logger.removeHandler(handler)
                    handler.close()
        assert ran == [True]
        events = [json.loads(line) for line in audit.read_text(encoding="utf-8").splitlines()]
        assert {"event": "object_tree", "disabled_sites": ["table"],
                "trace_file": None} in events

    def test_main_writes_no_event_when_unset(self, server, tmp_path, monkeypatch):
        audit = tmp_path / "audit.jsonl"
        self._isolate_main(server, tmp_path, monkeypatch, ["--audit-log", str(audit)])
        audit_logger = logging.getLogger("mcp_sap_gui.audit")
        handlers = list(audit_logger.handlers)
        try:
            server.main()
        finally:
            for handler in audit_logger.handlers[:]:
                if handler not in handlers:
                    audit_logger.removeHandler(handler)
                    handler.close()
        events = [json.loads(line) for line in audit.read_text(encoding="utf-8").splitlines()]
        assert [e["event"] for e in events] == ["policy"]
