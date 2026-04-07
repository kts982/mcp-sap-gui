"""Tests for SessionManager — per-MCP-client SAP session bindings.

These tests cover the gap identified by issue #9: SessionManager has no
direct unit coverage even though it owns the COM executor and the per-
client routing that the v0.2 architecture relies on.

Acceptance criteria from #9:
- Session-manager behavior is covered by automated tests
- Per-client routing and detach/owned-session cleanup are exercised
"""

from concurrent.futures import ThreadPoolExecutor
from unittest.mock import MagicMock, patch

from mcp_sap_gui.session_manager import ManagedSession, SessionManager


def _patch_controller():
    """Patch SAPGUIController so __init__ skips the pywin32 dependency check.

    SessionManager creates a real SAPGUIController per binding. The conftest
    fixture already mocks ``win32com``/``pythoncom`` at the sys.modules level,
    so the dependency check passes — but we still want to control the
    controller's behaviour without touching real COM.
    """
    return patch(
        "mcp_sap_gui.session_manager.SAPGUIController",
        autospec=True,
    )


def _make_controller_double(*, owns: bool = False, connected: bool = True):
    """Build a stand-in controller for tests that don't need full mocking."""
    ctrl = MagicMock()
    ctrl._owns_session = owns
    # is_connected is a property on the real class; on a MagicMock attribute
    # access works the same way.
    ctrl.is_connected = connected
    return ctrl


# ===========================================================================
# Construction / lifecycle
# ===========================================================================

class TestSessionManagerLifecycle:
    """Construction, executor ownership, and shutdown."""

    def test_starts_empty(self):
        mgr = SessionManager()
        try:
            assert mgr.active_count == 0
            assert mgr.list_sessions() == []
        finally:
            mgr.shutdown()

    def test_executor_is_single_threaded(self):
        """COM is apartment-threaded — the executor must serialise calls."""
        mgr = SessionManager()
        try:
            assert isinstance(mgr.executor, ThreadPoolExecutor)
            # Implementation detail but load-bearing: max_workers=1
            assert mgr.executor._max_workers == 1
        finally:
            mgr.shutdown()

    def test_executor_is_shared_across_sessions(self):
        """All managed sessions share the same COM thread."""
        mgr = SessionManager()
        try:
            with _patch_controller():
                mgr.get_or_create(1)
                mgr.get_or_create(2)
            # The executor reference must not change when new sessions arrive
            assert mgr.executor is mgr._executor
        finally:
            mgr.shutdown()


# ===========================================================================
# get_or_create
# ===========================================================================

class TestGetOrCreate:
    """Per-client routing via session_key."""

    def test_creates_new_binding_for_unknown_key(self):
        mgr = SessionManager()
        try:
            with _patch_controller() as ctrl_cls:
                managed = mgr.get_or_create(42)
                assert isinstance(managed, ManagedSession)
                assert mgr.active_count == 1
                ctrl_cls.assert_called_once_with()
        finally:
            mgr.shutdown()

    def test_returns_existing_binding_for_known_key(self):
        """Two calls with the same key return the same controller — no leak."""
        mgr = SessionManager()
        try:
            with _patch_controller() as ctrl_cls:
                first = mgr.get_or_create(42)
                second = mgr.get_or_create(42)
                assert first is second
                assert mgr.active_count == 1
                ctrl_cls.assert_called_once()  # NOT twice
        finally:
            mgr.shutdown()

    def test_different_keys_get_isolated_controllers(self):
        """Per-MCP-client isolation — keys must not cross-talk."""
        mgr = SessionManager()
        try:
            with _patch_controller() as ctrl_cls:
                # autospec returns the same MagicMock from every call by
                # default; force distinct controllers per construction so we
                # can prove the bindings really are isolated.
                ctrl_cls.side_effect = lambda: _make_controller_double()
                a = mgr.get_or_create(1)
                b = mgr.get_or_create(2)
                assert a is not b
                assert a.controller is not b.controller
                assert mgr.active_count == 2
        finally:
            mgr.shutdown()

    def test_get_or_create_touches_last_used(self):
        """Each get_or_create call refreshes the binding's last_used timestamp."""
        mgr = SessionManager()
        try:
            with _patch_controller():
                managed = mgr.get_or_create(1)
                first_seen = managed.last_used
                # Force a touch by re-fetching
                managed2 = mgr.get_or_create(1)
                assert managed is managed2
                assert managed.last_used >= first_seen
        finally:
            mgr.shutdown()


# ===========================================================================
# release / release_all — owned vs attached cleanup
# ===========================================================================

class TestRelease:
    """Cleanup behaviour for owned vs attached sessions."""

    def test_release_unknown_key_is_noop(self):
        mgr = SessionManager()
        try:
            result = mgr.release(999)
            assert result["released"] is False
            assert "no active session" in result["reason"]
        finally:
            mgr.shutdown()

    def test_release_owned_session_reports_closed(self):
        """Sessions opened via sap_connect are *closed* on release."""
        mgr = SessionManager()
        try:
            with _patch_controller() as ctrl_cls:
                ctrl_cls.return_value = _make_controller_double(
                    owns=True, connected=True,
                )
                mgr.get_or_create(1)
                result = mgr.release(1)
            assert result["released"] is True
            assert result["was_connected"] is True
            assert result["action"] == "closed"
            assert mgr.active_count == 0
        finally:
            mgr.shutdown()

    def test_release_attached_session_reports_detached(self):
        """Sessions attached via sap_connect_existing are *detached*, not closed."""
        mgr = SessionManager()
        try:
            with _patch_controller() as ctrl_cls:
                ctrl_cls.return_value = _make_controller_double(
                    owns=False, connected=True,
                )
                mgr.get_or_create(1)
                result = mgr.release(1)
            assert result["released"] is True
            assert result["was_connected"] is True
            assert result["action"] == "detached"
            assert mgr.active_count == 0
        finally:
            mgr.shutdown()

    def test_release_disconnected_session_reports_none(self):
        """A binding that was never connected reports action='none'."""
        mgr = SessionManager()
        try:
            with _patch_controller() as ctrl_cls:
                ctrl_cls.return_value = _make_controller_double(
                    owns=False, connected=False,
                )
                mgr.get_or_create(1)
                result = mgr.release(1)
            assert result["released"] is True
            assert result["was_connected"] is False
            assert result["action"] == "none"
        finally:
            mgr.shutdown()

    def test_release_calls_controller_disconnect(self):
        """release() must invoke disconnect() on the controller."""
        mgr = SessionManager()
        try:
            with _patch_controller() as ctrl_cls:
                ctrl = _make_controller_double(owns=True, connected=True)
                ctrl_cls.return_value = ctrl
                mgr.get_or_create(1)
                mgr.release(1)
                ctrl.disconnect.assert_called_once()
        finally:
            mgr.shutdown()

    def test_release_only_removes_targeted_key(self):
        """Releasing one binding must not affect others — isolation property."""
        mgr = SessionManager()
        try:
            with _patch_controller() as ctrl_cls:
                ctrl_cls.side_effect = lambda: _make_controller_double(
                    owns=True, connected=True,
                )
                mgr.get_or_create(1)
                mgr.get_or_create(2)
                mgr.release(1)
            assert mgr.active_count == 1
            keys = {entry["session_key"] for entry in mgr.list_sessions()}
            assert keys == {2}
        finally:
            mgr.shutdown()


class TestReleaseAll:
    """Bulk cleanup at server shutdown."""

    def test_release_all_clears_every_binding(self):
        mgr = SessionManager()
        try:
            with _patch_controller() as ctrl_cls:
                ctrl_cls.side_effect = lambda: _make_controller_double(
                    owns=True, connected=True,
                )
                mgr.get_or_create(1)
                mgr.get_or_create(2)
                mgr.get_or_create(3)
                assert mgr.active_count == 3
                mgr.release_all()
            assert mgr.active_count == 0
            assert mgr.list_sessions() == []
        finally:
            mgr.shutdown()

    def test_release_all_swallows_individual_errors(self):
        """One bad binding must not block cleanup of the others."""
        mgr = SessionManager()
        try:
            with _patch_controller() as ctrl_cls:
                bad_ctrl = _make_controller_double(owns=True, connected=True)
                bad_ctrl.disconnect.side_effect = RuntimeError("COM dead")
                good_ctrl = _make_controller_double(owns=True, connected=True)
                ctrl_cls.side_effect = [bad_ctrl, good_ctrl]

                mgr.get_or_create(1)
                mgr.get_or_create(2)
                mgr.release_all()
            assert mgr.active_count == 0
        finally:
            mgr.shutdown()


# ===========================================================================
# list_sessions
# ===========================================================================

class TestListSessions:
    """Metadata reporting for active bindings."""

    def test_list_sessions_returns_one_entry_per_binding(self):
        mgr = SessionManager()
        try:
            with _patch_controller() as ctrl_cls:
                ctrl_cls.side_effect = lambda: _make_controller_double(
                    owns=True, connected=True,
                )
                mgr.get_or_create(10)
                mgr.get_or_create(20)
                entries = mgr.list_sessions()
            assert len(entries) == 2
            keys = {e["session_key"] for e in entries}
            assert keys == {10, 20}
        finally:
            mgr.shutdown()

    def test_list_sessions_reports_owned_vs_attached(self):
        """list_sessions surfaces the owns_session flag for each binding."""
        mgr = SessionManager()
        try:
            owned = _make_controller_double(owns=True, connected=True)
            attached = _make_controller_double(owns=False, connected=True)
            with _patch_controller() as ctrl_cls:
                ctrl_cls.side_effect = [owned, attached]
                mgr.get_or_create(1)
                mgr.get_or_create(2)
                entries = {
                    e["session_key"]: e for e in mgr.list_sessions()
                }
            assert entries[1]["owns_session"] is True
            assert entries[2]["owns_session"] is False
            assert entries[1]["connected"] is True
            assert entries[2]["connected"] is True
        finally:
            mgr.shutdown()

    def test_list_sessions_reports_age_and_idle_seconds(self):
        """Each entry must include age_seconds and idle_seconds."""
        mgr = SessionManager()
        try:
            with _patch_controller() as ctrl_cls:
                ctrl_cls.return_value = _make_controller_double(
                    owns=True, connected=True,
                )
                mgr.get_or_create(1)
                entry = mgr.list_sessions()[0]
            assert "age_seconds" in entry
            assert "idle_seconds" in entry
            assert isinstance(entry["age_seconds"], (int, float))
            assert isinstance(entry["idle_seconds"], (int, float))
            assert entry["age_seconds"] >= 0
            assert entry["idle_seconds"] >= 0
        finally:
            mgr.shutdown()
