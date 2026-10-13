"""Session manager for per-MCP-client SAP GUI session bindings.

Each MCP client (identified by its ServerSession object identity) gets
its own SAPGUIController so that concurrent clients don't cross-talk.

For stdio transport (single client), there is effectively one entry.
For streamable HTTP, each client gets an independent binding.
"""

import logging
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional

from .sap_controller import SAPGUIController

logger = logging.getLogger(__name__)


@dataclass
class ManagedSession:
    """A per-MCP-client SAP session binding."""

    controller: SAPGUIController
    created_at: float = field(default_factory=time.monotonic)
    last_used: float = field(default_factory=time.monotonic)

    def touch(self) -> None:
        """Update last-used timestamp."""
        self.last_used = time.monotonic()


class SessionManager:
    """Registry of per-MCP-client SAP GUI session bindings.

    Owns the single-threaded ``ThreadPoolExecutor`` that serialises all
    COM calls (SAP GUI COM is apartment-threaded).
    """

    def __init__(self) -> None:
        self._sessions: Dict[int, ManagedSession] = {}
        # Session-added confirmation points, keyed like _sessions but stored
        # separately: the confirmation middleware consults them on *every*
        # tool call, and get_or_create would eagerly build a SAPGUIController
        # for calls that never touch SAP.
        self._confirmation_points: Dict[int, set] = {}
        # Session-level policy profile (sap_set_policy_profile), same lifetime
        # as the confirmation points. Absent means the server floor applies.
        self._profiles: Dict[int, str] = {}
        self._executor = ThreadPoolExecutor(max_workers=1)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    @property
    def executor(self) -> ThreadPoolExecutor:
        """The shared COM thread pool (max_workers=1)."""
        return self._executor

    @property
    def active_count(self) -> int:
        return len(self._sessions)

    def get_or_create(self, session_key: int) -> ManagedSession:
        """Get or lazily create a managed session for *session_key*."""
        if session_key not in self._sessions:
            controller = SAPGUIController()
            self._sessions[session_key] = ManagedSession(controller=controller)
            logger.info(
                "Created SAP session binding (key=%s, total=%d)",
                session_key,
                len(self._sessions),
            )
        managed = self._sessions[session_key]
        managed.touch()
        return managed

    def get_existing(self, session_key: int) -> Optional[ManagedSession]:
        """The binding for *session_key*, or None: never creates one."""
        return self._sessions.get(session_key)

    def get_confirmation_points(self, session_key: int) -> set:
        """Return the confirmation points this session added.

        Read-only lookup: never creates a session binding, so the confirmation
        middleware can call it on every tool call.  The server floor and the
        always-on ``save`` point are merged in by the caller.
        """
        return set(self._confirmation_points.get(session_key, ()))

    def set_confirmation_points(self, session_key: int, points: Any) -> set:
        """Replace the confirmation points this session added."""
        stored = set(points)
        self._confirmation_points[session_key] = stored
        return set(stored)

    def get_profile(self, session_key: int) -> Optional[str]:
        """The profile this session switched to, or None for the floor."""
        return self._profiles.get(session_key)

    def set_profile(self, session_key: int, profile: str) -> None:
        """Record the effective profile of this session."""
        self._profiles[session_key] = profile

    def release(self, session_key: int, close_owned: bool = True) -> Dict[str, Any]:
        """Release and clean up a managed session.

        Owned sessions (opened via ``sap_connect``, and /o windows opened from
        them) are closed when ``close_owned`` is True and left open otherwise.
        Attached sessions (via ``sap_connect_existing``) are always detached.

        Confirmation points are deliberately NOT dropped here: their lifetime
        is the MCP session, not the SAP binding.  Clearing them on release
        would let ``sap_disconnect`` (or a no-op disconnect) silently disarm
        every gate the user asked for.

        Safe to call from the COM executor thread.
        """
        managed = self._sessions.pop(session_key, None)
        if managed is None:
            return {"released": False, "reason": "no active session binding"}

        was_connected = managed.controller.is_connected
        owned_open = len(managed.controller._owned_sessions)
        closed = managed.controller.disconnect(close_owned=close_owned)

        if closed:
            action = "closed"
        else:
            action = "detached" if was_connected else "none"
        logger.info(
            "Released SAP session binding (key=%s, action=%s, remaining=%d)",
            session_key,
            action,
            len(self._sessions),
        )
        result: Dict[str, Any] = {
            "released": True,
            "was_connected": was_connected,
            "action": action,
        }
        if closed:
            result["closed_sessions"] = closed
        if not close_owned and owned_open:
            result["owned_sessions_left_open"] = owned_open
        return result

    def release_all(
        self, may_close: Optional[Callable[[int], bool]] = None,
    ) -> None:
        """Release every session.  Called on server shutdown.

        ``may_close(session_key)`` decides whether that binding's owned
        sessions are closed; without it they are (the previous behaviour).
        """
        for key in list(self._sessions):
            try:
                self.release(
                    key, close_owned=True if may_close is None else may_close(key),
                )
            except Exception:
                logger.warning("Error releasing session %s", key, exc_info=True)

    def list_sessions(self) -> List[Dict[str, Any]]:
        """Return metadata for every active managed session."""
        result: List[Dict[str, Any]] = []
        now = time.monotonic()
        for key, ms in self._sessions.items():
            result.append({
                "session_key": key,
                "connected": ms.controller.is_connected,
                "owns_session": ms.controller._owns_session,
                "age_seconds": round(now - ms.created_at, 1),
                "idle_seconds": round(now - ms.last_used, 1),
            })
        return result

    def shutdown(self) -> None:
        """Shut down the COM executor.  Call *after* ``release_all``.

        The only place confirmation points are cleared: they outlive SAP
        connect/disconnect cycles and die with the server.
        """
        self._confirmation_points.clear()
        self._profiles.clear()
        self._executor.shutdown(wait=False)
