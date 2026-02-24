"""
SAP GUI Controller - COM automation wrapper for SAP GUI Scripting API.

This module provides a Python interface to SAP GUI for Windows via COM automation.
It wraps the SAP GUI Scripting API to enable programmatic interaction with SAP.

Requirements:
- Windows OS
- SAP GUI for Windows installed
- SAP GUI Scripting enabled (transaction RZ11, parameter sapgui/user_scripting)
- pywin32 package

Reference:
- SAP GUI Scripting API: https://help.sap.com/docs/sap_gui_for_windows
- SAP Note 480149 - SAP GUI Scripting Security
"""

from typing import Optional, Dict, Any, List
from dataclasses import dataclass
from enum import IntEnum
import logging

logger = logging.getLogger(__name__)

# SAP GUI Virtual Keys
class VKey(IntEnum):
    """SAP GUI virtual key codes."""
    ENTER = 0
    F1 = 1   # Help
    F2 = 2
    F3 = 3   # Back
    F4 = 4   # Dropdown/Search help
    F5 = 5   # Refresh
    F6 = 6
    F7 = 7
    F8 = 8   # Execute
    F9 = 9
    F10 = 10
    F11 = 11  # Save
    F12 = 12  # Cancel
    SHIFT_F1 = 13
    SHIFT_F2 = 14
    SHIFT_F3 = 15  # Back (same as F3)
    SHIFT_F4 = 16
    SHIFT_F5 = 17
    SHIFT_F6 = 18
    SHIFT_F7 = 19
    SHIFT_F8 = 20
    SHIFT_F9 = 21
    CTRL_S = 11     # Save (same as F11)
    CTRL_F = 32     # Find
    CTRL_G = 33     # Continue search
    CTRL_P = 34     # Print
    ESC = 12        # Cancel (same as F12)


@dataclass
class SessionInfo:
    """Information about the current SAP session."""
    system_name: str
    system_number: str
    client: str
    user: str
    language: str
    transaction: str
    program: str
    screen_number: int
    session_number: int


@dataclass
class ScreenElement:
    """Information about a screen element."""
    id: str
    type: str
    name: str
    text: str
    changeable: bool
    visible: bool


class SAPGUIError(Exception):
    """Exception raised for SAP GUI errors."""
    pass


class SAPGUINotAvailableError(SAPGUIError):
    """Exception raised when SAP GUI is not available."""
    pass


class SAPGUINotConnectedError(SAPGUIError):
    """Exception raised when not connected to SAP."""
    pass


# GetToolbarButtonType() returns strings per SAP GUI Scripting API v8.00,
# but some SAP GUI versions may return numeric values. Handle both.
_TOOLBAR_BUTTON_TYPES = {
    0: "Button", 1: "ButtonAndMenu", 2: "Menu",
    3: "Separator", 4: "CheckBox", 5: "Group",
    "Button": "Button", "ButtonAndMenu": "ButtonAndMenu",
    "Menu": "Menu", "Separator": "Separator",
    "CheckBox": "CheckBox", "Group": "Group",
}


class SAPGUIController:
    """
    Controller for SAP GUI Scripting API via COM automation.

    This class provides methods to interact with SAP GUI programmatically,
    including connecting to systems, navigating transactions, reading/writing
    fields, and extracting data from tables.

    Example usage:
        controller = SAPGUIController()
        controller.connect("DEV - Development System")
        controller.execute_transaction("MM03")
        controller.set_field("wnd[0]/usr/ctxtRMMG1-MATNR", "MAT-001")
        controller.send_vkey(VKey.ENTER)
        description = controller.read_field("wnd[0]/usr/txtMAKT-MAKTX")
    """

    def __init__(self):
        """Initialize the SAP GUI controller."""
        self._win32com = None
        self._sap_gui_auto = None
        self._application = None
        self._connection = None
        self._session = None
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

            logger.info(f"Connected to existing session {connection_index}/{session_index}")
            return self.get_session_info()

        except SAPGUIError:
            raise
        except Exception as e:
            raise SAPGUIError(f"Failed to connect to existing session: {e}")

    def disconnect(self):
        """Close the current SAP session."""
        if self._session:
            try:
                self._connection.CloseSession(self._session.Id)
            except Exception:
                pass
        self._session = None
        self._connection = None
        logger.info("Disconnected")

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

        # /o prefix opens a new window - must use okcd approach
        if tcode.upper().startswith("/O"):
            self._session.findById("wnd[0]/tbar[0]/okcd").text = tcode
            self._session.findById("wnd[0]").sendVKey(VKey.ENTER)
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
            "transaction": tcode.removeprefix("/n").removeprefix("/o"),
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
        """Get information about the current screen."""
        self._require_session()

        try:
            window = self._session.findById("wnd[0]")
            info = self._session.Info
            status = self._get_status_bar_info()

            return {
                "transaction": info.Transaction,
                "program": info.Program,
                "screen_number": info.ScreenNumber,
                "title": window.Text,
                "message": status.get("text"),
                "message_type": status.get("message_type", ""),
                "message_id": status.get("message_id", ""),
                "message_number": status.get("message_number", ""),
            }
        except Exception as e:
            return {"error": str(e)}

    def _get_status_bar_info(self) -> Dict[str, Any]:
        """Get structured information from the status bar.

        Returns a dict with:
        - text: The full message text
        - message_type: S=Success, W=Warning, E=Error, A=Abort, I=Info
        - message_id: SAP message class (e.g., 'MM', 'SD')
        - message_number: Three-digit message number
        - message_parameters: List of up to 4 message parameters
        """
        try:
            sbar = self._session.findById("wnd[0]/sbar")
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

    # =========================================================================
    # Field Operations
    # =========================================================================

    def read_field(self, field_id: str) -> Dict[str, Any]:
        """
        Read a field value from the screen.

        Returns the field value plus metadata like required, max_length,
        numerical flag, and associated labels (when available from
        GuiTextField / GuiCTextField).

        Args:
            field_id: SAP GUI element ID (e.g., "wnd[0]/usr/txtMATNR")

        Returns:
            Dict with field value and properties
        """
        self._require_session()

        try:
            element = self._session.findById(field_id)

            result: Dict[str, Any] = {
                "field_id": field_id,
                "value": getattr(element, 'Text', ''),
                "type": element.Type,
                "name": getattr(element, 'Name', ''),
                "changeable": getattr(element, 'Changeable', None),
            }

            # Extended metadata for text fields
            for attr, key in [
                ("Required", "required"),
                ("MaxLength", "max_length"),
                ("Numerical", "numerical"),
                ("Highlighted", "highlighted"),
            ]:
                try:
                    val = getattr(element, attr, None)
                    if val is not None:
                        result[key] = val
                except Exception:
                    pass

            # Associated labels (GuiTextField / GuiCTextField)
            for attr, key in [
                ("LeftLabel", "left_label"),
                ("RightLabel", "right_label"),
            ]:
                try:
                    label = getattr(element, attr, None)
                    if label is not None:
                        result[key] = getattr(label, 'Text', str(label))
                except Exception:
                    pass

            return result
        except Exception as e:
            return {"field_id": field_id, "error": str(e)}

    def set_field(self, field_id: str, value: str) -> Dict[str, Any]:
        """
        Set a field value on the screen.

        Args:
            field_id: SAP GUI element ID
            value: Value to set

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            element = self._session.findById(field_id)
            element.text = value

            logger.debug(f"Set {field_id} = {value}")
            return {
                "field_id": field_id,
                "value": value,
                "status": "success",
            }
        except Exception as e:
            return {"field_id": field_id, "error": str(e)}

    def _safe_set_field(self, field_id: str, value: str) -> bool:
        """Set field value, returning False on error instead of raising."""
        try:
            self._session.findById(field_id).text = value
            return True
        except Exception:
            return False

    def press_button(self, button_id: str) -> Dict[str, Any]:
        """
        Press a button on the screen.

        Args:
            button_id: SAP GUI button ID (e.g., "wnd[0]/tbar[1]/btn[8]")

        Returns:
            Dict with screen info after button press
        """
        self._require_session()

        try:
            self._session.findById(button_id).press()

            logger.debug(f"Pressed button: {button_id}")
            return {
                "button_id": button_id,
                "status": "pressed",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {"button_id": button_id, "error": str(e)}

    def select_menu(self, menu_id: str) -> Dict[str, Any]:
        """
        Select a menu item from the menu bar or a submenu.

        Args:
            menu_id: SAP GUI menu ID (e.g., 'wnd[0]/mbar/menu[3]/menu[0]')

        Returns:
            Dict with screen info after menu selection
        """
        self._require_session()

        try:
            self._session.findById(menu_id).Select()

            return {
                "menu_id": menu_id,
                "status": "selected",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {"menu_id": menu_id, "error": str(e)}

    def select_checkbox(self, checkbox_id: str, selected: bool = True) -> Dict[str, Any]:
        """
        Select or deselect a checkbox.

        Args:
            checkbox_id: SAP GUI checkbox ID
            selected: True to select, False to deselect
        """
        self._require_session()

        try:
            element = self._session.findById(checkbox_id)
            element.selected = selected

            return {
                "checkbox_id": checkbox_id,
                "selected": selected,
                "status": "success",
            }
        except Exception as e:
            return {"checkbox_id": checkbox_id, "error": str(e)}

    def select_radio_button(self, radio_id: str) -> Dict[str, Any]:
        """
        Select a radio button.

        Args:
            radio_id: SAP GUI radio button ID (e.g., 'wnd[0]/usr/radOPT1')

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            element = self._session.findById(radio_id)
            element.Select()

            return {
                "radio_id": radio_id,
                "status": "success",
            }
        except Exception as e:
            return {"radio_id": radio_id, "error": str(e)}

    def select_combobox_entry(self, combobox_id: str, key_or_value: str) -> Dict[str, Any]:
        """
        Select an entry in a combobox/dropdown.

        First tries to set the key directly. If that fails, searches the
        Entries collection by value text.

        Args:
            combobox_id: SAP GUI combobox ID (e.g., 'wnd[0]/usr/cmbLANGU')
            key_or_value: Key or display value text of the entry to select

        Returns:
            Dict with result status and selected key/value
        """
        self._require_session()

        try:
            combobox = self._session.findById(combobox_id)

            # Try setting key directly first
            try:
                combobox.Key = key_or_value
                return {
                    "combobox_id": combobox_id,
                    "key": key_or_value,
                    "status": "success",
                }
            except Exception:
                pass

            # Fallback: search Entries by value text
            entries = combobox.Entries
            for i in range(entries.Count):
                entry = entries(i)
                if entry.Value == key_or_value or entry.Key == key_or_value:
                    combobox.Key = entry.Key
                    return {
                        "combobox_id": combobox_id,
                        "key": entry.Key,
                        "value": entry.Value,
                        "status": "success",
                    }

            return {
                "combobox_id": combobox_id,
                "error": f"Entry '{key_or_value}' not found in combobox",
            }
        except Exception as e:
            return {"combobox_id": combobox_id, "error": str(e)}

    def select_tab(self, tab_id: str) -> Dict[str, Any]:
        """
        Select a tab in a tab strip.

        Args:
            tab_id: SAP GUI tab ID (e.g., 'wnd[0]/usr/tabsTAB/tabpTAB1')

        Returns:
            Dict with result status and screen info after selection
        """
        self._require_session()

        try:
            tab = self._session.findById(tab_id)
            tab.Select()

            return {
                "tab_id": tab_id,
                "status": "success",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {"tab_id": tab_id, "error": str(e)}

    def get_combobox_entries(self, combobox_id: str) -> Dict[str, Any]:
        """
        List all entries in a combobox/dropdown.

        Returns the available key-value pairs so the caller knows which
        values are valid before attempting to set one.

        Args:
            combobox_id: SAP GUI combobox ID

        Returns:
            Dict with list of entries (key, value pairs)
        """
        self._require_session()

        try:
            combo = self._session.findById(combobox_id)
            entries = []
            for i in range(combo.Entries.Count):
                entry = combo.Entries(i)
                entries.append({
                    "key": entry.Key,
                    "value": entry.Value,
                })
            return {
                "combobox_id": combobox_id,
                "current_key": getattr(combo, 'Key', ''),
                "entry_count": len(entries),
                "entries": entries,
            }
        except Exception as e:
            return {"combobox_id": combobox_id, "error": str(e)}

    def set_batch_fields(self, fields: Dict[str, str]) -> Dict[str, Any]:
        """
        Set multiple field values at once.

        More efficient than calling set_field repeatedly — all values are
        set before a single round-trip return.

        Args:
            fields: Dict mapping field_id → value

        Returns:
            Dict with per-field results
        """
        self._require_session()

        results = {}
        for field_id, value in fields.items():
            try:
                self._session.findById(field_id).text = value
                results[field_id] = "success"
            except Exception as e:
                results[field_id] = f"error: {e}"

        succeeded = sum(1 for v in results.values() if v == "success")
        return {
            "total": len(fields),
            "succeeded": succeeded,
            "failed": len(fields) - succeeded,
            "results": results,
        }

    def read_textedit(self, textedit_id: str) -> Dict[str, Any]:
        """
        Read the content of a multiline text editor (GuiTextedit).

        Args:
            textedit_id: SAP GUI textedit ID

        Returns:
            Dict with full text, line count, and individual lines
        """
        self._require_session()

        try:
            textedit = self._session.findById(textedit_id)
            line_count = textedit.LineCount
            lines = []
            for i in range(line_count):
                try:
                    lines.append(textedit.GetLineText(i))
                except Exception:
                    lines.append("")

            return {
                "textedit_id": textedit_id,
                "line_count": line_count,
                "text": "\n".join(lines),
                "lines": lines,
                "changeable": getattr(textedit, 'Changeable', None),
            }
        except Exception as e:
            return {"textedit_id": textedit_id, "error": str(e)}

    def set_textedit(self, textedit_id: str, text: str) -> Dict[str, Any]:
        """
        Set the content of a multiline text editor (GuiTextedit).

        Attempts to set via the Text property first, then falls back
        to SetUnprotectedTextPart for protected editors.

        Args:
            textedit_id: SAP GUI textedit ID
            text: Text content to set

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            textedit = self._session.findById(textedit_id)
            try:
                textedit.Text = text
            except Exception:
                # Fallback: set only the unprotected part
                textedit.SetUnprotectedTextPart(text)

            return {
                "textedit_id": textedit_id,
                "status": "success",
            }
        except Exception as e:
            return {"textedit_id": textedit_id, "error": str(e)}

    def set_focus(self, element_id: str) -> Dict[str, Any]:
        """
        Set focus to any screen element.

        Args:
            element_id: SAP GUI element ID

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            self._session.findById(element_id).SetFocus()
            return {"element_id": element_id, "status": "success"}
        except Exception as e:
            return {"element_id": element_id, "error": str(e)}

    # =========================================================================
    # Table/Grid Operations
    # =========================================================================

    # ---- GuiTableControl helpers ----

    def _get_table_control_columns(self, table) -> list:
        """Get column metadata from a GuiTableControl's Columns collection.

        Each column entry has: index, name, title, tooltip.

        Note: The SAP GUI Scripting API docs state that GuiTableColumn members
        in the Columns collection "do not support properties like id or name".
        Column names are therefore extracted from the first row's cell Name
        property, with Title as fallback.
        """
        columns = []
        col_count = table.Columns.Count

        # Get column names from first row's cells (safer than col.Name
        # which is documented as unsupported on GuiTableColumn)
        cell_names = []
        for i in range(col_count):
            name = None
            try:
                cell = table.GetCell(0, i)
                name = getattr(cell, 'Name', None)
            except Exception:
                pass
            cell_names.append(name)

        # Get column titles and tooltips from the Columns collection
        for i in range(col_count):
            info = {"index": i}
            try:
                col = table.Columns(i)
                try:
                    info["title"] = col.Title
                except Exception:
                    info["title"] = ""
                try:
                    info["tooltip"] = col.Tooltip
                except Exception:
                    info["tooltip"] = ""
            except Exception:
                info["title"] = ""
                info["tooltip"] = ""

            info["name"] = cell_names[i] or info.get("title") or f"col_{i}"
            columns.append(info)
        return columns

    def _read_cell_value(self, cell):
        """Read a cell value, handling different element types."""
        try:
            cell_type = getattr(cell, 'Type', '')
            if cell_type == "GuiCheckBox":
                return bool(cell.Selected)
            elif cell_type == "GuiComboBox":
                return getattr(cell, 'Key', cell.Text)
            else:
                return cell.Text
        except Exception:
            return None

    def _scroll_table_control_to_row(self, table, abs_row: int) -> int:
        """Scroll a GuiTableControl so *abs_row* is visible.

        Returns the **visible-row offset** to pass to ``GetCell()``.
        ``GetCell`` uses visible-row indexing (0 = first visible row),
        so callers must use the returned offset, not the original
        *abs_row*.
        """
        visible = table.VisibleRowCount
        scrollbar = table.VerticalScrollbar
        current_top = scrollbar.Position

        if current_top <= abs_row < current_top + visible:
            return abs_row - current_top  # already visible

        new_pos = max(scrollbar.Minimum, min(abs_row, scrollbar.Maximum))
        scrollbar.Position = new_pos
        return abs_row - new_pos

    def _resolve_table_control_column(self, table, column) -> int:
        """Resolve a column name/index to a numeric column index for GuiTableControl."""
        if isinstance(column, int):
            return column
        if isinstance(column, str) and column.isdigit():
            return int(column)

        # Search by cell Name (from first row) and column Title
        for i in range(table.Columns.Count):
            # Try cell Name first
            try:
                cell = table.GetCell(0, i)
                if getattr(cell, 'Name', None) == column:
                    return i
            except Exception:
                pass
            # Try column Title
            try:
                col = table.Columns(i)
                if col.Title == column:
                    return i
            except Exception:
                pass

        raise ValueError(f"Column '{column}' not found in table")

    # ---- Table reading ----

    def read_table(self, table_id: str, max_rows: int = 100) -> Dict[str, Any]:
        """
        Read data from an ALV grid or table control.

        Automatically detects GuiGridView (ALV) vs GuiTableControl and
        uses the appropriate API for each.

        Args:
            table_id: SAP GUI table/grid ID
            max_rows: Maximum rows to read (default 100)

        Returns:
            Dict with table data and metadata
        """
        self._require_session()

        try:
            table = self._session.findById(table_id)
            if getattr(table, 'Type', '') == "GuiTableControl":
                return self._read_table_control(table, table_id, max_rows)
            return self._read_alv_grid(table, table_id, max_rows)
        except Exception as e:
            return {"table_id": table_id, "error": str(e)}

    def _read_alv_grid(self, grid, table_id: str, max_rows: int) -> Dict[str, Any]:
        """Read data from an ALV grid (GuiGridView)."""
        columns = []
        for i in range(grid.ColumnCount):
            columns.append(grid.ColumnOrder(i))

        column_info = []
        for col in columns:
            info = {"name": col}
            try:
                info["tooltip"] = grid.GetColumnTooltip(col)
            except Exception:
                info["tooltip"] = ""
            try:
                info["title"] = grid.GetDisplayedColumnTitle(col)
            except Exception:
                info["title"] = col
            column_info.append(info)

        row_count = min(grid.RowCount, max_rows)
        data = []
        for row in range(row_count):
            row_data = {}
            for col in columns:
                try:
                    row_data[col] = grid.GetCellValue(row, col)
                except Exception:
                    row_data[col] = None
            data.append(row_data)

        return {
            "table_id": table_id,
            "table_type": "GuiGridView",
            "total_rows": grid.RowCount,
            "rows_returned": len(data),
            "columns": columns,
            "column_info": column_info,
            "data": data,
        }

    def _read_table_control(self, table, table_id: str, max_rows: int) -> Dict[str, Any]:
        """Read visible rows from a GuiTableControl.

        Reads from the current scroll position without changing it.  This
        preserves the user's navigated position (e.g. after "Position..." in
        SPRO/SM30) and avoids crashing SAP GUI's COM server — programmatic
        scrollbar.Position changes can destabilize certain table views.

        The read is capped at VisibleRowCount.  GuiTableControl.RowCount often
        includes padding rows (empty rows that fill the visible area beyond
        the actual data), so reading stops early when an all-empty row is
        encountered.
        """
        columns_info = self._get_table_control_columns(table)
        column_names = [c["name"] for c in columns_info]
        col_count = len(columns_info)

        total_rows = table.RowCount
        visible_rows = table.VisibleRowCount

        data = []
        if total_rows > 0 and col_count > 0:
            scrollbar = table.VerticalScrollbar
            start_position = scrollbar.Position
            rows_to_read = min(visible_rows, max_rows)

            for vis_idx in range(rows_to_read):
                row_data = {}
                all_empty = True
                for col_idx in range(col_count):
                    try:
                        cell = table.GetCell(vis_idx, col_idx)
                        value = self._read_cell_value(cell)
                    except Exception:
                        value = None
                    row_data[column_names[col_idx]] = value
                    if value is not None and value != "":
                        all_empty = False
                if all_empty:
                    break
                data.append(row_data)
        else:
            start_position = 0

        return {
            "table_id": table_id,
            "table_type": "GuiTableControl",
            "table_field_name": getattr(table, 'TableFieldName', ''),
            "total_rows": total_rows,
            "first_visible_row": start_position,
            "visible_rows": visible_rows,
            "rows_returned": len(data),
            "columns": column_names,
            "column_info": columns_info,
            "data": data,
        }

    def get_alv_toolbar(self, grid_id: str) -> Dict[str, Any]:
        """
        Get all toolbar buttons from an ALV grid.

        Args:
            grid_id: SAP GUI grid ID

        Returns:
            Dict with list of toolbar buttons (id, text, type)
        """
        self._require_session()

        try:
            grid = self._session.findById(grid_id)
            button_count = grid.ToolbarButtonCount

            buttons = []
            for i in range(button_count):
                btn_id = grid.GetToolbarButtonId(i)
                btn_text = grid.GetToolbarButtonText(i)
                btn_type = grid.GetToolbarButtonType(i)

                btn_info = {
                    "index": i,
                    "id": btn_id,
                    "text": btn_text,
                    "type": _TOOLBAR_BUTTON_TYPES.get(btn_type, str(btn_type)),
                }

                # Add tooltip if available
                try:
                    btn_info["tooltip"] = grid.GetToolbarButtonTooltip(i)
                except Exception:
                    btn_info["tooltip"] = ""

                # Add enabled state if available
                try:
                    btn_info["enabled"] = bool(grid.GetToolbarButtonEnabled(i))
                except Exception:
                    btn_info["enabled"] = True

                buttons.append(btn_info)

            return {
                "grid_id": grid_id,
                "button_count": button_count,
                "buttons": buttons,
            }
        except Exception as e:
            return {"grid_id": grid_id, "error": str(e)}

    def press_alv_toolbar_button(self, grid_id: str, button_id: str) -> Dict[str, Any]:
        """
        Press a toolbar button on an ALV grid.

        Automatically detects Menu/ButtonAndMenu types and uses
        PressToolbarContextButton instead of PressToolbarButton.
        For menu types, this opens the context menu — use
        select_alv_context_menu_item() to pick an item.

        Args:
            grid_id: SAP GUI grid ID
            button_id: Toolbar button ID (from get_alv_toolbar)

        Returns:
            Dict with result status, screen info, and menu items if a menu was opened
        """
        self._require_session()

        try:
            grid = self._session.findById(grid_id)

            # Try PressToolbarContextButton first (works for Menu/ButtonAndMenu)
            # Fall back to PressToolbarButton for regular buttons
            menu_opened = False
            try:
                grid.PressToolbarContextButton(button_id)
                menu_opened = True
            except Exception:
                # Not a menu button, use regular press
                grid.PressToolbarButton(button_id)

            if menu_opened:
                return {
                    "grid_id": grid_id,
                    "button_id": button_id,
                    "status": "menu_opened",
                    "hint": "Use select_alv_context_menu_item with SelectToolbarMenuItem to pick an item",
                    "screen": self.get_screen_info(),
                }
            else:
                return {
                    "grid_id": grid_id,
                    "button_id": button_id,
                    "status": "pressed",
                    "screen": self.get_screen_info(),
                }
        except Exception as e:
            return {"grid_id": grid_id, "button_id": button_id, "error": str(e)}

    def select_alv_context_menu_item(self, grid_id: str, menu_item_id: str,
                                      toolbar_button_id: str = None) -> Dict[str, Any]:
        """
        Select an item from an ALV context menu.

        If toolbar_button_id is provided, opens the context menu first and then
        immediately selects the item — all in a single call to avoid timing issues.

        Args:
            grid_id: SAP GUI grid ID
            menu_item_id: Function code or visible text of the menu item
            toolbar_button_id: Optional toolbar button to open the menu first

        Returns:
            Dict with result status and screen info
        """
        self._require_session()

        try:
            grid = self._session.findById(grid_id)

            # Open context menu first if toolbar button specified
            if toolbar_button_id:
                grid.PressToolbarContextButton(toolbar_button_id)

            # Use SelectToolbarMenuItem (for toolbar context menus opened
            # via PressToolbarContextButton). Falls back to
            # SelectContextMenuItemByText for visible-text matching.
            if ' ' in menu_item_id:
                # Looks like visible text, use SelectContextMenuItemByText
                grid.SelectContextMenuItemByText(menu_item_id)
            else:
                # Looks like a function code — use SelectToolbarMenuItem
                # (the correct API for toolbar menus), with fallback
                try:
                    grid.SelectToolbarMenuItem(menu_item_id)
                except Exception:
                    grid.SelectContextMenuItemByText(menu_item_id)

            return {
                "grid_id": grid_id,
                "menu_item_id": menu_item_id,
                "status": "selected",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {"grid_id": grid_id, "menu_item_id": menu_item_id, "error": str(e)}

    def select_table_row(self, table_id: str, row: int) -> Dict[str, Any]:
        """Select a row in a table/grid.

        For GuiTableControl, ``row`` is an **absolute** row index.  The
        table is scrolled if necessary to make the row visible, then
        ``GetAbsoluteRow(row).Selected`` is set.

        Per the SAP GUI Scripting API, ``GetAbsoluteRow`` uses absolute
        indexing (independent of scroll position), unlike ``Rows()``
        which resets after scrolling.  Confirmed via SAP GUI script
        recording which produces ``getAbsoluteRow(N).selected = true``.
        """
        self._require_session()

        try:
            table = self._session.findById(table_id)

            if getattr(table, 'Type', '') == "GuiTableControl":
                self._scroll_table_control_to_row(table, row)
                # GetAbsoluteRow: absolute indexing, not affected by scroll
                table.GetAbsoluteRow(row).Selected = True
            else:
                table.selectedRows = str(row)

            return {
                "table_id": table_id,
                "selected_row": row,
                "status": "success",
            }
        except Exception as e:
            return {"table_id": table_id, "error": str(e)}

    def double_click_table_cell(self, table_id: str, row: int, column: str) -> Dict[str, Any]:
        """Double-click a cell in a table/grid.

        For GuiTableControl, ``row`` is an **absolute** row index.  The
        table is scrolled if necessary, the row is selected via
        ``GetAbsoluteRow``, focus is set on the target cell via
        ``GetCell`` (visible-row indexing), and F2 is sent to the
        owning window.
        """
        self._require_session()

        try:
            table = self._session.findById(table_id)

            if getattr(table, 'Type', '') == "GuiTableControl":
                vis_row = self._scroll_table_control_to_row(table, row)
                col_idx = self._resolve_table_control_column(table, column)
                table.GetAbsoluteRow(row).Selected = True
                cell = table.GetCell(vis_row, col_idx)
                cell.SetFocus()
                # Send F2 to the window that owns this table
                wnd_id = table_id.split("/usr")[0]
                self._session.findById(wnd_id).sendVKey(VKey.F2)
            else:
                table.DoubleClick(row, column)

            return {
                "table_id": table_id,
                "row": row,
                "column": column,
                "status": "double_clicked",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {"table_id": table_id, "error": str(e)}

    def modify_cell(self, grid_id: str, row: int, column: str, value: str) -> Dict[str, Any]:
        """
        Modify the value of a cell in an ALV grid or table control.

        Args:
            grid_id: SAP GUI grid/table ID
            row: Row index (0-based)
            column: Column name (ALV) or column name/index (table control)
            value: New cell value

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            table = self._session.findById(grid_id)

            if getattr(table, 'Type', '') == "GuiTableControl":
                vis_row = self._scroll_table_control_to_row(table, row)
                col_idx = self._resolve_table_control_column(table, column)
                cell = table.GetCell(vis_row, col_idx)
                cell.Text = value
            else:
                table.ModifyCell(row, column, value)

            return {
                "grid_id": grid_id,
                "row": row,
                "column": column,
                "value": value,
                "status": "success",
            }
        except Exception as e:
            return {"grid_id": grid_id, "row": row, "column": column, "error": str(e)}

    def set_current_cell(self, grid_id: str, row: int, column: str) -> Dict[str, Any]:
        """
        Set the current (focused) cell in an ALV grid or table control.

        Args:
            grid_id: SAP GUI grid/table ID
            row: Row index (0-based)
            column: Column name (ALV) or column name/index (table control)

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            table = self._session.findById(grid_id)

            if getattr(table, 'Type', '') == "GuiTableControl":
                vis_row = self._scroll_table_control_to_row(table, row)
                col_idx = self._resolve_table_control_column(table, column)
                cell = table.GetCell(vis_row, col_idx)
                cell.SetFocus()
            else:
                table.SetCurrentCell(row, column)

            return {
                "grid_id": grid_id,
                "row": row,
                "column": column,
                "status": "success",
            }
        except Exception as e:
            return {"grid_id": grid_id, "row": row, "column": column, "error": str(e)}

    def get_column_info(self, grid_id: str) -> Dict[str, Any]:
        """
        Get detailed column information from an ALV grid or table control.

        Returns column names, displayed titles, tooltips for each column.

        Args:
            grid_id: SAP GUI grid/table ID

        Returns:
            Dict with column details
        """
        self._require_session()

        try:
            table = self._session.findById(grid_id)

            if getattr(table, 'Type', '') == "GuiTableControl":
                columns = self._get_table_control_columns(table)
                return {
                    "grid_id": grid_id,
                    "table_type": "GuiTableControl",
                    "column_count": len(columns),
                    "columns": columns,
                }

            # ALV grid
            columns = []
            for i in range(table.ColumnCount):
                col_name = table.ColumnOrder(i)
                col_info = {"name": col_name, "index": i}

                try:
                    col_info["title"] = table.GetDisplayedColumnTitle(col_name)
                except Exception:
                    col_info["title"] = col_name

                try:
                    col_info["tooltip"] = table.GetColumnTooltip(col_name)
                except Exception:
                    col_info["tooltip"] = ""

                columns.append(col_info)

            return {
                "grid_id": grid_id,
                "column_count": len(columns),
                "columns": columns,
            }
        except Exception as e:
            return {"grid_id": grid_id, "error": str(e)}

    # ---- TableControl-specific operations ----

    def scroll_table_control(self, table_id: str, position: int) -> Dict[str, Any]:
        """
        Scroll a GuiTableControl to a specific row position.

        Since read_table does not scroll (it reads only visible rows),
        use this tool to navigate to a different section of the table
        before reading.

        Args:
            table_id: SAP GUI table control ID
            position: Absolute row position to scroll to

        Returns:
            Dict with new scroll position and visible data summary
        """
        self._require_session()

        try:
            table = self._session.findById(table_id)
            if getattr(table, 'Type', '') != "GuiTableControl":
                return {"table_id": table_id, "error": "Not a GuiTableControl. Use ALV grid methods instead."}

            scrollbar = table.VerticalScrollbar
            new_pos = max(scrollbar.Minimum, min(position, scrollbar.Maximum))
            scrollbar.Position = new_pos

            return {
                "table_id": table_id,
                "status": "success",
                "position": new_pos,
                "visible_rows": table.VisibleRowCount,
                "total_rows": table.RowCount,
                "scroll_max": scrollbar.Maximum,
            }
        except Exception as e:
            return {"table_id": table_id, "error": str(e)}

    def get_table_control_row_info(self, table_id: str,
                                    rows: Optional[List[int]] = None) -> Dict[str, Any]:
        """
        Get row metadata from a GuiTableControl.

        Returns whether each row is selectable and currently selected,
        using GetAbsoluteRow() for absolute indexing.

        Args:
            table_id: SAP GUI table control ID
            rows: List of absolute row indices to query.
                  If None, queries all currently visible rows.

        Returns:
            Dict with row info list
        """
        self._require_session()

        try:
            table = self._session.findById(table_id)
            if getattr(table, 'Type', '') != "GuiTableControl":
                return {"table_id": table_id, "error": "Not a GuiTableControl"}

            if rows is None:
                start = table.VerticalScrollbar.Position
                rows = list(range(start, start + table.VisibleRowCount))

            row_info = []
            for r in rows:
                info: Dict[str, Any] = {"row": r}
                try:
                    abs_row = table.GetAbsoluteRow(r)
                    info["selectable"] = getattr(abs_row, 'Selectable', True)
                    info["selected"] = getattr(abs_row, 'Selected', False)
                except Exception as e:
                    info["error"] = str(e)
                row_info.append(info)

            return {
                "table_id": table_id,
                "row_count": len(row_info),
                "rows": row_info,
            }
        except Exception as e:
            return {"table_id": table_id, "error": str(e)}

    def select_all_table_control_columns(self, table_id: str,
                                          select: bool = True) -> Dict[str, Any]:
        """
        Select or deselect all columns in a GuiTableControl.

        Args:
            table_id: SAP GUI table control ID
            select: True to select all, False to deselect all

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            table = self._session.findById(table_id)
            if getattr(table, 'Type', '') != "GuiTableControl":
                return {"table_id": table_id, "error": "Not a GuiTableControl"}

            if select:
                table.SelectAllColumns()
            else:
                table.DeselectAllColumns()

            return {
                "table_id": table_id,
                "status": "all_selected" if select else "all_deselected",
            }
        except Exception as e:
            return {"table_id": table_id, "error": str(e)}

    # ---- ALV-specific operations ----

    def get_cell_info(self, grid_id: str, row: int,
                      column: str) -> Dict[str, Any]:
        """
        Get detailed cell metadata from an ALV grid (GuiGridView).

        Returns whether the cell is editable, its color/style, and tooltip.

        Args:
            grid_id: SAP GUI grid ID (ALV)
            row: Row index (0-based)
            column: Column name

        Returns:
            Dict with cell properties
        """
        self._require_session()

        try:
            grid = self._session.findById(grid_id)
            info: Dict[str, Any] = {
                "grid_id": grid_id,
                "row": row,
                "column": column,
                "value": grid.GetCellValue(row, column),
            }

            for method, key in [
                ("GetCellChangeable", "changeable"),
                ("GetCellColor", "color"),
                ("GetCellTooltip", "tooltip"),
                ("GetCellStyle", "style"),
                ("GetCellMaxLength", "max_length"),
            ]:
                try:
                    info[key] = getattr(grid, method)(row, column)
                except Exception:
                    pass

            return info
        except Exception as e:
            return {"grid_id": grid_id, "row": row, "column": column, "error": str(e)}

    def press_column_header(self, grid_id: str,
                             column: str) -> Dict[str, Any]:
        """
        Click a column header in an ALV grid (triggers sort/filter).

        Args:
            grid_id: SAP GUI grid ID (ALV)
            column: Column name

        Returns:
            Dict with result status and screen info
        """
        self._require_session()

        try:
            grid = self._session.findById(grid_id)
            grid.PressColumnHeader(column)
            return {
                "grid_id": grid_id,
                "column": column,
                "status": "pressed",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {"grid_id": grid_id, "column": column, "error": str(e)}

    def select_all_rows(self, grid_id: str) -> Dict[str, Any]:
        """
        Select all rows in an ALV grid.

        Args:
            grid_id: SAP GUI grid ID (ALV)

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            grid = self._session.findById(grid_id)
            grid.SelectAll()
            return {"grid_id": grid_id, "status": "all_selected"}
        except Exception as e:
            return {"grid_id": grid_id, "error": str(e)}

    # ---- Operations for both table types ----

    def get_current_cell(self, table_id: str) -> Dict[str, Any]:
        """
        Get the currently focused cell in an ALV grid or table control.

        Args:
            table_id: SAP GUI grid/table ID

        Returns:
            Dict with current row and column
        """
        self._require_session()

        try:
            table = self._session.findById(table_id)

            if getattr(table, 'Type', '') == "GuiTableControl":
                return {
                    "table_id": table_id,
                    "table_type": "GuiTableControl",
                    "current_row": getattr(table, 'CurrentRow', -1),
                    "current_col": getattr(table, 'CurrentCol', -1),
                }
            else:
                return {
                    "table_id": table_id,
                    "table_type": "GuiGridView",
                    "current_row": getattr(table, 'CurrentCellRow', -1),
                    "current_column": getattr(table, 'CurrentCellColumn', ''),
                }
        except Exception as e:
            return {"table_id": table_id, "error": str(e)}

    # =========================================================================
    # Tree Control Operations
    # =========================================================================

    def _get_tree_column_info(self, tree) -> tuple:
        """
        Get column names and titles from a tree control.

        Uses the official SAP GUI Scripting API methods:
        - GetColumnNames() → GuiCollection of internal column names
        - GetColumnHeaders() → GuiCollection of display titles
        - GetColumnTitleFromName(name) → title for a specific column
        - ColumnOrder property → column sequence (Column Tree only)

        Tree types (GetTreeType): 0=Simple, 1=List, 2=Column

        Returns:
            Tuple of (column_names: list, column_titles: list)
        """
        column_names = []
        column_titles = []

        # Detect tree type
        tree_type_num = -1
        try:
            tree_type_num = tree.GetTreeType()
            logger.debug("Tree type number: %s", tree_type_num)
        except Exception as e:
            logger.debug("GetTreeType failed: %s", e)

        # Simple trees (type 0) have no columns
        if tree_type_num == 0:
            return column_names, column_titles

        # Strategy 1: GetColumnNames() — returns a GuiCollection (works for List & Column trees)
        try:
            names_col = tree.GetColumnNames()
            if hasattr(names_col, 'Count'):
                for i in range(names_col.Count):
                    column_names.append(str(names_col(i)))
            elif hasattr(names_col, 'Length'):
                for i in range(names_col.Length):
                    column_names.append(str(names_col(i)))
            elif hasattr(names_col, '__iter__'):
                column_names = [str(n) for n in names_col]
            logger.debug("Got column names via GetColumnNames: %s", column_names)
        except Exception as e:
            logger.debug("GetColumnNames failed: %s", e)

        # Strategy 2: ColumnOrder property (Column Tree type 2 only)
        if not column_names and tree_type_num == 2:
            try:
                col_order = tree.ColumnOrder
                if hasattr(col_order, 'Count'):
                    for i in range(col_order.Count):
                        column_names.append(str(col_order(i)))
                elif hasattr(col_order, '__iter__'):
                    column_names = [str(n) for n in col_order]
                logger.debug("Got column names via ColumnOrder: %s", column_names)
            except Exception as e:
                logger.debug("ColumnOrder failed: %s", e)

        # Get column titles
        if column_names:
            # Try GetColumnTitleFromName for each column
            for name in column_names:
                try:
                    column_titles.append(tree.GetColumnTitleFromName(name))
                except Exception:
                    column_titles.append(name)

            # If all titles are empty, try GetColumnHeaders as fallback
            if all(not t for t in column_titles):
                try:
                    headers_col = tree.GetColumnHeaders()
                    fallback_titles = []
                    if hasattr(headers_col, 'Count'):
                        for i in range(headers_col.Count):
                            fallback_titles.append(str(headers_col(i)))
                    if fallback_titles:
                        column_titles = fallback_titles
                        logger.debug("Got titles via GetColumnHeaders: %s", column_titles)
                except Exception as e:
                    logger.debug("GetColumnHeaders fallback failed: %s", e)

        return column_names, column_titles

    def read_tree(self, tree_id: str, max_nodes: int = 200) -> Dict[str, Any]:
        """
        Read data from a tree control (SAP.TableTreeControl, SAP.ColumnTreeControl, etc.).

        Args:
            tree_id: SAP GUI tree ID (e.g., "wnd[0]/usr/shell/shellcont[0]/shell")
            max_nodes: Maximum number of nodes to read (default 200)

        Returns:
            Dict with tree structure, columns, and node data
        """
        self._require_session()

        try:
            tree = self._session.findById(tree_id)

            # Detect tree type (0=Simple, 1=List, 2=Column)
            tree_type = ""
            tree_type_num = -1
            try:
                tree_type_num = tree.GetTreeType()
                tree_type_names = {0: "Simple", 1: "List", 2: "Column"}
                tree_type = tree_type_names.get(tree_type_num, f"Unknown({tree_type_num})")
            except Exception:
                # Fallback to SubType / Text for display
                try:
                    tree_type = tree.SubType if hasattr(tree, 'SubType') else ""
                except Exception:
                    pass
                if not tree_type:
                    try:
                        tree_type = tree.Text
                    except Exception:
                        pass
            logger.debug("Tree type: %s (num=%s)", tree_type, tree_type_num)

            # Get hierarchy title if available (List/Column trees)
            hierarchy_title = ""
            try:
                hierarchy_title = tree.GetHierarchyTitle()
            except Exception:
                pass

            # Get column info using fallback chain
            column_names, column_titles = self._get_tree_column_info(tree)

            # Get all visible node keys
            node_keys = []
            try:
                keys = tree.GetAllNodeKeys()
                if hasattr(keys, 'Count'):
                    for i in range(min(keys.Count, max_nodes)):
                        node_keys.append(keys(i))
                elif hasattr(keys, '__iter__'):
                    for i, key in enumerate(keys):
                        if i >= max_nodes:
                            break
                        node_keys.append(key)
            except Exception as e:
                return {"tree_id": tree_id, "error": f"Cannot read node keys: {e}"}

            # Read each node
            nodes = []
            for key in node_keys:
                node = {"key": key}

                # Node text (works for SimpleTree; empty for TableTree/ColumnTree)
                try:
                    node["text"] = tree.GetNodeTextByKey(key)
                except Exception:
                    node["text"] = ""

                # Parent key - API documents GetParent(), fall back to GetParentNodeKey()
                try:
                    node["parent_key"] = tree.GetParent(key)
                except Exception:
                    try:
                        node["parent_key"] = tree.GetParentNodeKey(key)
                    except Exception:
                        node["parent_key"] = None

                # Children count
                try:
                    node["children_count"] = tree.GetNodeChildrenCount(key)
                except Exception:
                    node["children_count"] = 0

                # Folder state
                try:
                    node["is_folder"] = tree.IsFolderExpandable(key)
                except Exception:
                    node["is_folder"] = False

                try:
                    node["is_expanded"] = tree.IsFolderExpanded(key)
                except Exception:
                    node["is_expanded"] = False

                # Hierarchy level
                try:
                    node["hierarchy_level"] = tree.GetHierarchyLevel(key)
                except Exception:
                    node["hierarchy_level"] = None

                # Column values (for TableTree / ColumnTree)
                if column_names:
                    col_values = {}
                    for col_name in column_names:
                        try:
                            col_values[col_name] = tree.GetItemText(key, col_name)
                        except Exception:
                            col_values[col_name] = None
                    node["columns"] = col_values

                    # If node text is empty, use first non-empty column value
                    if not node["text"]:
                        for col_name in column_names:
                            val = col_values.get(col_name)
                            if val:
                                node["text"] = val
                                break

                nodes.append(node)

            return {
                "tree_id": tree_id,
                "tree_type": tree_type,
                "hierarchy_title": hierarchy_title,
                "total_nodes": len(node_keys),
                "column_titles": column_titles,
                "column_names": column_names,
                "nodes": nodes,
            }

        except Exception as e:
            return {"tree_id": tree_id, "error": str(e)}

    def expand_tree_node(self, tree_id: str, node_key: str) -> Dict[str, Any]:
        """
        Expand a folder node in a tree control.

        Args:
            tree_id: SAP GUI tree ID
            node_key: The key of the node to expand

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            tree = self._session.findById(tree_id)
            tree.ExpandNode(node_key)

            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "status": "expanded",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {"tree_id": tree_id, "node_key": node_key, "error": str(e)}

    def collapse_tree_node(self, tree_id: str, node_key: str) -> Dict[str, Any]:
        """
        Collapse a folder node in a tree control.

        Args:
            tree_id: SAP GUI tree ID
            node_key: The key of the node to collapse

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            tree = self._session.findById(tree_id)
            tree.CollapseNode(node_key)

            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "status": "collapsed",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {"tree_id": tree_id, "node_key": node_key, "error": str(e)}

    def select_tree_node(self, tree_id: str, node_key: str) -> Dict[str, Any]:
        """
        Select a node in a tree control.

        Args:
            tree_id: SAP GUI tree ID
            node_key: The key of the node to select

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            tree = self._session.findById(tree_id)
            tree.SelectNode(node_key)

            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "status": "selected",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {"tree_id": tree_id, "node_key": node_key, "error": str(e)}

    def double_click_tree_node(self, tree_id: str, node_key: str) -> Dict[str, Any]:
        """
        Double-click a node in a tree control (often opens details).

        Args:
            tree_id: SAP GUI tree ID
            node_key: The key of the node to double-click

        Returns:
            Dict with result status and screen info
        """
        self._require_session()

        try:
            tree = self._session.findById(tree_id)
            tree.DoubleClickNode(node_key)

            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "status": "double_clicked",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {"tree_id": tree_id, "node_key": node_key, "error": str(e)}

    def double_click_tree_item(self, tree_id: str, node_key: str,
                               item_name: str) -> Dict[str, Any]:
        """
        Double-click a specific item (column cell) in a tree node.

        Unlike DoubleClickNode which clicks the node itself, this clicks
        on a specific column cell within the node row.

        Args:
            tree_id: SAP GUI tree ID
            node_key: The key of the node
            item_name: Column name / item name to double-click

        Returns:
            Dict with result status and screen info
        """
        self._require_session()

        try:
            tree = self._session.findById(tree_id)
            tree.DoubleClickItem(node_key, item_name)

            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "item_name": item_name,
                "status": "double_clicked",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {
                "tree_id": tree_id, "node_key": node_key,
                "item_name": item_name, "error": str(e),
            }

    def click_tree_link(self, tree_id: str, node_key: str,
                        item_name: str) -> Dict[str, Any]:
        """
        Click a hyperlink in a tree node item.

        Args:
            tree_id: SAP GUI tree ID
            node_key: The key of the node
            item_name: Column name / item name containing the link

        Returns:
            Dict with result status and screen info
        """
        self._require_session()

        try:
            tree = self._session.findById(tree_id)
            tree.ClickLink(node_key, item_name)

            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "item_name": item_name,
                "status": "clicked",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {
                "tree_id": tree_id, "node_key": node_key,
                "item_name": item_name, "error": str(e),
            }

    def find_tree_node_by_path(self, tree_id: str, path: str) -> Dict[str, Any]:
        """
        Find a node key by its path in the tree hierarchy.

        The path is a backslash-separated string of child indices,
        e.g. "2\\1\\2" means: 2nd child of root, then 1st child, then 2nd child.

        Args:
            tree_id: SAP GUI tree ID
            path: Path string (e.g., "2\\1\\2")

        Returns:
            Dict with the found node key
        """
        self._require_session()

        try:
            tree = self._session.findById(tree_id)
            node_key = tree.FindNodeKeyByPath(path)

            return {
                "tree_id": tree_id,
                "path": path,
                "node_key": node_key,
                "status": "found",
            }
        except Exception as e:
            return {"tree_id": tree_id, "path": path, "error": str(e)}

    # =========================================================================
    # Screen Element Discovery
    # =========================================================================

    def get_screen_elements(self, container_id: str = "wnd[0]/usr",
                            max_depth: int = 3) -> List[ScreenElement]:
        """
        Enumerate all elements on the current screen.

        Useful for discovering field IDs when automating a new transaction.

        Args:
            container_id: Starting container (default: main user area)
            max_depth: Maximum recursion depth

        Returns:
            List of ScreenElement objects
        """
        self._require_session()

        try:
            container = self._session.findById(container_id)
            elements = self._enumerate_elements(container, max_depth)
            return elements
        except Exception as e:
            logger.error(f"Failed to enumerate elements: {e}")
            return []

    def _enumerate_elements(self, container, max_depth: int,
                            current_depth: int = 0) -> List[ScreenElement]:
        """Recursively enumerate screen elements."""
        elements = []

        if current_depth >= max_depth:
            return elements

        try:
            for i in range(container.Children.Count):
                child = container.Children(i)

                element = ScreenElement(
                    id=child.Id,
                    type=child.Type,
                    name=getattr(child, 'Name', ''),
                    text=str(getattr(child, 'Text', ''))[:200],
                    changeable=getattr(child, 'Changeable', False),
                    visible=getattr(child, 'Visible', True),
                )
                elements.append(element)

                # Recurse into containers
                if hasattr(child, 'Children') and child.Children.Count > 0:
                    child_elements = self._enumerate_elements(
                        child, max_depth, current_depth + 1
                    )
                    elements.extend(child_elements)

        except Exception as e:
            logger.debug(f"Error enumerating at depth {current_depth}: {e}")

        return elements

    # =========================================================================
    # Screenshot & Visual
    # =========================================================================

    def _find_topmost_window(self) -> str:
        """Find the topmost SAP GUI window (highest wnd index that exists).

        Popups appear as wnd[1], wnd[2], etc. This returns the topmost
        window so screenshots and screen reads capture what the user sees.

        Tries Session.ActiveWindow first (faster), falls back to loop.
        """
        try:
            active = self._session.ActiveWindow
            if active is not None:
                return active.Id
        except Exception:
            pass

        topmost = "wnd[0]"
        for i in range(1, 10):
            try:
                self._session.findById(f"wnd[{i}]")
                topmost = f"wnd[{i}]"
            except Exception:
                break
        return topmost

    def take_screenshot(self, filepath: str = None) -> Dict[str, Any]:
        """
        Take a screenshot of the current SAP window.

        Args:
            filepath: Optional file path. If not provided, returns base64 data.

        Returns:
            Dict with filepath or base64 encoded image data
        """
        self._require_session()

        import tempfile
        import os

        try:
            if filepath is None:
                filepath = os.path.join(tempfile.gettempdir(), "sap_screenshot.png")
                return_base64 = True
            else:
                return_base64 = False

            # Find the topmost window (popups are wnd[1], wnd[2], etc.)
            window_id = self._find_topmost_window()
            window = self._session.findById(window_id)
            window.HardCopy(filepath, "PNG")

            # Optimize image size with Pillow if available
            self._optimize_screenshot(filepath)

            if return_base64:
                import base64
                with open(filepath, "rb") as f:
                    data = base64.b64encode(f.read()).decode()
                os.remove(filepath)
                return {
                    "format": "png",
                    "encoding": "base64",
                    "window": window_id,
                    "data": data,
                }
            else:
                return {
                    "format": "png",
                    "filepath": filepath,
                    "window": window_id,
                }

        except Exception as e:
            return {"error": str(e)}

    def _optimize_screenshot(self, filepath: str) -> None:
        """
        Optimize screenshot file size using Pillow if available.

        Resizes large images and applies PNG optimization to significantly
        reduce file size (typically 70-90% reduction).
        """
        try:
            from PIL import Image
        except ImportError:
            logger.debug("Pillow not installed, skipping screenshot optimization")
            return

        try:
            img = Image.open(filepath)

            # Downscale if image is very large (> 1920px wide)
            max_width = 1920
            if img.width > max_width:
                ratio = max_width / img.width
                new_size = (max_width, int(img.height * ratio))
                img = img.resize(new_size, Image.LANCZOS)

            # Convert RGBA to RGB if no transparency (smaller file)
            if img.mode == "RGBA":
                background = Image.new("RGB", img.size, (255, 255, 255))
                background.paste(img, mask=img.split()[3])
                img = background

            # Save with optimization
            img.save(filepath, "PNG", optimize=True)
        except Exception as e:
            logger.debug(f"Screenshot optimization failed (using original): {e}")
