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
from typing import Any, Dict, List

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

    def release(self, session_key: int) -> Dict[str, Any]:
        """Release and clean up a managed session.

        Owned sessions (opened via ``sap_connect``) are closed.
        Attached sessions (via ``sap_connect_existing``) are detached.

        Safe to call from the COM executor thread.
        """
        managed = self._sessions.pop(session_key, None)
        if managed is None:
            return {"released": False, "reason": "no active session binding"}

        was_connected = managed.controller.is_connected
        owns = managed.controller._owns_session
        managed.controller.disconnect()

        action = ("closed" if owns else "detached") if was_connected else "none"
        logger.info(
            "Released SAP session binding (key=%s, action=%s, remaining=%d)",
            session_key,
            action,
            len(self._sessions),
        )
        return {
            "released": True,
            "was_connected": was_connected,
            "action": action,
        }

    def release_all(self) -> None:
        """Release every session.  Called on server shutdown."""
        for key in list(self._sessions):
            try:
                self.release(key)
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
        """Shut down the COM executor.  Call *after* ``release_all``."""
        self._executor.shutdown(wait=False)
