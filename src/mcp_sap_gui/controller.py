"""
SAP GUI Controller Base - connection, navigation, and screen info.

This module provides the core controller class with connection management,
transaction execution, and screen information retrieval.
"""

import logging
from typing import Any, Dict, List, Optional

from .models import (
    SAPGUIError,
    SAPGUINotAvailableError,
    SAPGUINotConnectedError,
    SessionInfo,
    VKey,
    _strip_tcode_prefix,
)

logger = logging.getLogger(__name__)


class SAPGUIControllerBase:
    """
    Base controller for SAP GUI Scripting API via COM automation.

    Provides connection management, transaction navigation, and screen
    information retrieval.  Extended by mixins for fields, tables, trees,
    and discovery operations.
    """

    def __init__(self):
        """Initialize the SAP GUI controller."""
        self._win32com = None
        self._sap_gui_auto = None
        self._application = None
        self._connection = None
        self._session = None
        self._owns_session = False
        self._check_dependencies()

    def _check_dependencies(self):
        """Check if required dependencies are available."""
        try:
            import win32com.client
            self._win32com = win32com.client
        except ImportError:
            raise SAPGUINotAvailableError(
                "pywin32 is required but not installed. "
                "Install with: pip install pywin32"
            )

    @property
    def is_connected(self) -> bool:
        """Check if connected to an SAP system."""
        return self._session is not None

    def _ensure_com_initialized(self):
        """Ensure COM is initialized on the current thread (needed for thread pool workers)."""
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass  # Already initialized or not needed

    def _get_sap_gui(self):
        """Get the SAP GUI automation object."""
        if self._sap_gui_auto is None:
            self._ensure_com_initialized()
            try:
                self._sap_gui_auto = self._win32com.GetObject("SAPGUI")
            except Exception as e:
                raise SAPGUINotAvailableError(
                    f"Cannot connect to SAP GUI. Is SAP Logon Pad running? Error: {e}"
                )
        return self._sap_gui_auto

    def _get_application(self):
        """Get the SAP GUI scripting engine."""
        if self._application is None:
            sap_gui = self._get_sap_gui()
            # Use property access (no parentheses) - works more reliably
            self._application = sap_gui.GetScriptingEngine
            if self._application is None:
                raise SAPGUINotAvailableError(
                    "Could not get SAP GUI Scripting Engine. "
                    "Is SAP GUI Scripting enabled?"
                )
        return self._application

    def _normalize_window_id(self, window_id: Any) -> str:
        """Normalize SAP window IDs to the short form used by findById()."""
        if not isinstance(window_id, str):
            return "wnd[0]"
        if "wnd[" in window_id:
            return "wnd[" + window_id.split("wnd[", 1)[1]
        return window_id

    def _require_session(self):
        """Ensure we have an active session that is not busy."""
        if not self.is_connected:
            raise SAPGUINotConnectedError(
                "Not connected to SAP. Call connect() first."
            )
        try:
            if self._session.Busy:
                raise SAPGUIError(
                    "SAP session is busy processing a previous request. "
                    "Wait for it to complete before sending another command."
                )
        except AttributeError:
            pass  # Busy property not available on this version

    # =========================================================================
    # Connection Management
    # =========================================================================

    def connect(self, system_description: str,
                client: str = None,
                user: str = None,
                password: str = None,
                language: str = None) -> SessionInfo:
        """
        Connect to an SAP system.

        Opens a connection exactly like double-clicking in SAP Logon Pad.
        Optionally fills login credentials.

        Args:
            system_description: Exact system name as shown in SAP Logon
            client: SAP client number (optional)
            user: SAP username (optional)
            password: SAP password (optional)
            language: Login language (optional, e.g., "EN")

        Returns:
            SessionInfo with connection details

        Raises:
            SAPGUIError: If connection fails
        """
        try:
            app = self._get_application()

            logger.info("Opening connection to: %s", system_description)
            self._connection = app.OpenConnection(system_description, True)
            self._owns_session = True

            if self._connection is None:
                raise SAPGUIError(f"Failed to open connection to '{system_description}'")

            self._session = self._connection.Children(0)

            if self._session is None:
                raise SAPGUIError("No session available on the connection")

            # Fill login credentials if provided
            if client:
                self._safe_set_field("wnd[0]/usr/txtRSYST-MANDT", str(client))
            if user:
                self._safe_set_field("wnd[0]/usr/txtRSYST-BNAME", user)
            if password:
                self._safe_set_field("wnd[0]/usr/pwdRSYST-BCODE", password)
            if language:
                self._safe_set_field("wnd[0]/usr/txtRSYST-LANGU", language)

            # Press Enter to login if credentials were provided
            has_credentials = password is not None
            if has_credentials:
                self.send_vkey(VKey.ENTER)

            logger.info("Connected successfully to %s as %s",
                         system_description, user or "(existing credentials)")
            return self.get_session_info()

        except SAPGUIError:
            raise
        except Exception as e:
            raise SAPGUIError(f"Connection failed: {e}")

    def connect_to_existing_session(self, connection_index: int = 0,
                                     session_index: int = 0) -> SessionInfo:
        """
        Connect to an already open SAP session.

        Args:
            connection_index: Index of the connection (0 = first)
            session_index: Index of the session within the connection (0 = first)

        Returns:
            SessionInfo with session details
        """
        try:
            app = self._get_application()

            if app.Children.Count == 0:
                raise SAPGUIError("No SAP connections found")

            if connection_index >= app.Children.Count:
                raise SAPGUIError(
                    f"Connection index {connection_index} out of range. "
                    f"Available: 0-{app.Children.Count - 1}"
                )

            self._connection = app.Children(connection_index)

            if session_index >= self._connection.Children.Count:
                raise SAPGUIError(
                    f"Session index {session_index} out of range. "
                    f"Available: 0-{self._connection.Children.Count - 1}"
                )

            self._session = self._connection.Children(session_index)
            self._owns_session = False

            logger.info(f"Connected to existing session {connection_index}/{session_index}")
            return self.get_session_info()

        except SAPGUIError:
            raise
        except Exception as e:
            raise SAPGUIError(f"Failed to connect to existing session: {e}")

    def disconnect(self):
        """Detach from the current SAP session.

        Sessions opened by this controller are closed on disconnect. Sessions
        attached via connect_to_existing_session() are left open.
        """
        if self._session and self._owns_session and self._connection:
            try:
                self._connection.CloseSession(self._session.Id)
            except Exception:
                pass
        self._session = None
        self._connection = None
        self._owns_session = False
        logger.info("Disconnected")

    def _find_session_by_id(self, session_id: str):
        """Locate a session by ID and return its connection plus session object."""
        app = self._get_application()
        for connection_index in range(app.Children.Count):
            connection = app.Children(connection_index)
            for session_index in range(connection.Children.Count):
                session = connection.Children(session_index)
                if getattr(session, "Id", None) == session_id:
                    return connection, session
        return None, None

    def _rebind_after_new_session(
        self, previous_session_id: str, previous_session_count: int | None,
    ) -> None:
        """Rebind to the newest session after a /o command opens a new window."""
        try:
            active_session = self._get_application().ActiveSession
            active_session_id = getattr(active_session, "Id", None)
            if active_session_id and active_session_id != previous_session_id:
                connection, session = self._find_session_by_id(active_session_id)
                if session is not None:
                    self._connection = connection
                    self._session = session
                    return
        except Exception:
            pass

        if self._connection is None or previous_session_count is None:
            return

        try:
            current_session_count = self._connection.Children.Count
            if current_session_count > previous_session_count:
                self._session = self._connection.Children(current_session_count - 1)
        except Exception:
            pass

    def get_session_info(self) -> SessionInfo:
        """Get information about the current session."""
        self._require_session()

        info = self._session.Info
        return SessionInfo(
            system_name=info.SystemName,
            system_number=str(info.SystemNumber),
            client=info.Client,
            user=info.User,
            language=info.Language,
            transaction=info.Transaction,
            program=info.Program,
            screen_number=info.ScreenNumber,
            session_number=info.SessionNumber,
        )

    def list_connections(self) -> List[Dict[str, Any]]:
        """List all open SAP connections and sessions."""
        app = self._get_application()

        connections = []
        for i in range(app.Children.Count):
            conn = app.Children(i)
            sessions = []

            # Get connection description (try multiple properties)
            conn_desc = ""
            try:
                conn_desc = conn.Description
            except Exception:
                try:
                    conn_desc = conn.ConnectionString
                except Exception:
                    conn_desc = f"Connection {i}"

            for j in range(conn.Children.Count):
                try:
                    sess = conn.Children(j)
                    info = sess.Info
                    sessions.append({
                        "index": j,
                        "id": sess.Id,
                        "user": info.User,
                        "transaction": info.Transaction,
                        "system": info.SystemName,
                        "client": info.Client,
                    })
                except Exception as e:
                    sessions.append({
                        "index": j,
                        "error": str(e),
                    })

            connections.append({
                "index": i,
                "id": getattr(conn, 'Id', f"conn_{i}"),
                "description": conn_desc,
                "session_count": conn.Children.Count,
                "sessions": sessions,
            })

        return connections

    # =========================================================================
    # Transaction & Navigation
    # =========================================================================

    def execute_transaction(self, tcode: str) -> Dict[str, Any]:
        """
        Execute a transaction code.

        Uses Session.StartTransaction() when possible (cleaner API),
        falls back to okcd + sendVKey for /o (new window) prefix.

        Args:
            tcode: Transaction code (e.g., "MM03", "VA01", "SE80")
                   Can include /n prefix for new session or /o for new window

        Returns:
            Dict with transaction and screen info
        """
        self._require_session()

        # Ensure proper format
        if not tcode.startswith("/"):
            tcode = f"/n{tcode}"

        logger.info(f"Executing transaction: {tcode}")

        # /o and /* prefixes must use okcd approach (not StartTransaction)
        upper = tcode.upper()
        opens_new_session = upper.startswith("/O")
        previous_session_id = getattr(self._session, "Id", "")
        previous_session_count = None
        if opens_new_session and self._connection is not None:
            try:
                previous_session_count = self._connection.Children.Count
            except Exception:
                previous_session_count = None

        if opens_new_session or upper.startswith("/*"):
            self._session.findById("wnd[0]/tbar[0]/okcd").text = tcode
            self._session.findById("wnd[0]").sendVKey(VKey.ENTER)
            if opens_new_session and previous_session_id:
                self._rebind_after_new_session(
                    previous_session_id, previous_session_count,
                )
        else:
            # Use StartTransaction for /n prefix (preferred API)
            plain_tcode = tcode.removeprefix("/n").removeprefix("/N")
            try:
                self._session.StartTransaction(plain_tcode)
            except Exception:
                # Fallback to okcd approach
                self._session.findById("wnd[0]/tbar[0]/okcd").text = tcode
                self._session.findById("wnd[0]").sendVKey(VKey.ENTER)

        return {
            "transaction": _strip_tcode_prefix(tcode),
            "screen": self.get_screen_info(),
        }

    def send_vkey(self, vkey: int, window: str = "wnd[0]") -> Dict[str, Any]:
        """
        Send a virtual key to SAP.

        Args:
            vkey: Virtual key code (use VKey enum)
            window: Target window ID (default: main window)

        Returns:
            Dict with screen info after key press
        """
        self._require_session()

        logger.debug(f"Sending VKey {vkey} to {window}")
        self._session.findById(window).sendVKey(vkey)

        return {"vkey": vkey, "screen": self.get_screen_info()}

    def press_enter(self) -> Dict[str, Any]:
        """Press Enter key."""
        return self.send_vkey(VKey.ENTER)

    def press_back(self) -> Dict[str, Any]:
        """Press Back (F3)."""
        return self.send_vkey(VKey.F3)

    def press_cancel(self) -> Dict[str, Any]:
        """Press Cancel (F12/ESC)."""
        return self.send_vkey(VKey.F12)

    def press_save(self) -> Dict[str, Any]:
        """Press Save (Ctrl+S/F11)."""
        return self.send_vkey(VKey.F11)

    def press_execute(self) -> Dict[str, Any]:
        """Press Execute (F8)."""
        return self.send_vkey(VKey.F8)

    def get_screen_info(self) -> Dict[str, Any]:
        """Get information about the current screen.

        Reads from ``session.ActiveWindow`` so the title always reflects
        what the user actually sees.  When ActiveWindow is a popup
        (wnd[1], wnd[2], ...) the response ``active_window`` field tells
        the caller which window is in focus -- no separate tool call
        needed.

        ``session.Info`` (transaction, program, screen_number) already
        reflects the active screen regardless of which window is focused.
        """
        self._require_session()

        try:
            info = self._session.Info

            # Determine the active window -- use ActiveWindow when
            # available, fall back to wnd[0].
            active_wnd_id = "wnd[0]"
            try:
                active = self._session.ActiveWindow
                if active is not None:
                    active_wnd_id = self._normalize_window_id(active.Id)
            except Exception:
                pass

            # Read title from the active window
            try:
                title = self._session.findById(active_wnd_id).Text
            except Exception:
                title = ""

            # Status bar -- try active window first, fall back to wnd[0]
            status = self._get_status_bar_info(active_wnd_id)

            result: Dict[str, Any] = {
                "active_window": active_wnd_id,
                "transaction": info.Transaction,
                "program": info.Program,
                "screen_number": info.ScreenNumber,
                "title": title,
                "message": status.get("text"),
                "message_type": status.get("message_type", ""),
                "message_id": status.get("message_id", ""),
                "message_number": status.get("message_number", ""),
            }

            return result
        except Exception as e:
            return {"error": str(e)}

    def _get_status_bar_info(
        self, window_id: str = "wnd[0]",
    ) -> Dict[str, Any]:
        """Get structured information from the status bar.

        Tries the given *window_id*'s status bar first (e.g. a popup's
        ``wnd[1]/sbar``).  Falls back to ``wnd[0]/sbar`` when the
        active window doesn't have one.

        Returns a dict with:
        - text: The full message text
        - message_type: S=Success, W=Warning, E=Error, A=Abort, I=Info
        - message_id: SAP message class (e.g., 'MM', 'SD')
        - message_number: Three-digit message number
        - message_parameters: List of up to 4 message parameters
        """
        sbar = None
        for wnd in dict.fromkeys([window_id, "wnd[0]"]):
            try:
                sbar = self._session.findById(f"{wnd}/sbar")
                break
            except Exception:
                continue
        if sbar is None:
            return {"text": None}
        try:
            info: Dict[str, Any] = {"text": sbar.Text}
            for attr, key in [
                ("MessageType", "message_type"),
                ("MessageId", "message_id"),
                ("MessageNumber", "message_number"),
            ]:
                try:
                    info[key] = getattr(sbar, attr)
                except Exception:
                    info[key] = ""
            # Message parameters (up to 4)
            params = []
            for attr in ["MessageParameter", "MessageParameter1",
                         "MessageParameter2", "MessageParameter3"]:
                try:
                    val = getattr(sbar, attr, None)
                    if val is not None:
                        params.append(str(val))
                except Exception:
                    pass
            if params:
                info["message_parameters"] = params
            return info
        except Exception:
            return {"text": None}

    def _get_status_bar_message(self) -> Optional[str]:
        """Get the message text from the status bar (legacy helper)."""
        return self._get_status_bar_info().get("text")

    def _safe_set_field(self, field_id: str, value: str) -> bool:
        """Set field value, returning False on error instead of raising."""
        try:
            self._session.findById(field_id).text = value
            return True
        except Exception:
            return False
