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
        """Ensure we have an active session."""
        if not self.is_connected:
            raise SAPGUINotConnectedError(
                "Not connected to SAP. Call connect() first."
            )

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

            logger.info(f"Opening connection to: {system_description}")
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
            if password:
                self.send_vkey(VKey.ENTER)

            logger.info("Connected successfully")
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

        self._session.findById("wnd[0]/tbar[0]/okcd").text = tcode
        self._session.findById("wnd[0]").sendVKey(VKey.ENTER)

        return {
            "transaction": tcode.lstrip("/n").lstrip("/o"),
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

            return {
                "transaction": info.Transaction,
                "program": info.Program,
                "screen_number": info.ScreenNumber,
                "title": window.Text,
                "message": self._get_status_bar_message(),
            }
        except Exception as e:
            return {"error": str(e)}

    def _get_status_bar_message(self) -> Optional[str]:
        """Get the message from the status bar."""
        try:
            statusbar = self._session.findById("wnd[0]/sbar")
            return statusbar.Text
        except Exception:
            return None

    # =========================================================================
    # Field Operations
    # =========================================================================

    def read_field(self, field_id: str) -> Dict[str, Any]:
        """
        Read a field value from the screen.

        Args:
            field_id: SAP GUI element ID (e.g., "wnd[0]/usr/txtMATNR")

        Returns:
            Dict with field value and properties
        """
        self._require_session()

        try:
            element = self._session.findById(field_id)

            return {
                "field_id": field_id,
                "value": getattr(element, 'Text', ''),
                "type": element.Type,
                "name": getattr(element, 'Name', ''),
                "changeable": getattr(element, 'Changeable', None),
            }
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

    # =========================================================================
    # Table/Grid Operations
    # =========================================================================

    def read_table(self, table_id: str, max_rows: int = 100) -> Dict[str, Any]:
        """
        Read data from an ALV grid or table control.

        Args:
            table_id: SAP GUI table/grid ID
            max_rows: Maximum rows to read (default 100)

        Returns:
            Dict with table data and metadata
        """
        self._require_session()

        try:
            grid = self._session.findById(table_id)

            # Get column information
            columns = []
            for i in range(grid.ColumnCount):
                columns.append(grid.ColumnOrder(i))

            # Read data
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
                "total_rows": grid.RowCount,
                "rows_returned": len(data),
                "columns": columns,
                "data": data,
            }
        except Exception as e:
            return {"table_id": table_id, "error": str(e)}

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

                # Skip separators (type 3)
                type_names = {
                    0: "Button",
                    1: "ButtonAndMenu",
                    2: "Menu",
                    3: "Separator",
                    4: "CheckBox",
                }

                buttons.append({
                    "index": i,
                    "id": btn_id,
                    "text": btn_text,
                    "type": type_names.get(btn_type, str(btn_type)),
                })

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
            menu_items = []
            try:
                grid.PressToolbarContextButton(button_id)
                menu_opened = True

                # Try to read the context menu items
                try:
                    ctx_menu = grid.ContextMenu
                    for i in range(ctx_menu.Count):
                        item = ctx_menu.Item(i)
                        menu_items.append({
                            "id": item.FunctionCode,
                            "text": item.Text,
                        })
                except Exception:
                    pass

            except Exception:
                # Not a menu button, use regular press
                grid.PressToolbarButton(button_id)

            if menu_opened:
                return {
                    "grid_id": grid_id,
                    "button_id": button_id,
                    "status": "menu_opened",
                    "menu_items": menu_items,
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

            # Detect if menu_item_id looks like a function code or visible text
            # Function codes are typically UPPERCASE_WITH_UNDERSCORES
            # Visible text contains spaces and mixed case
            if ' ' in menu_item_id:
                # Looks like visible text, use SelectContextMenuItemByText
                grid.SelectContextMenuItemByText(menu_item_id)
            else:
                # Looks like a function code, try by code first then by text
                try:
                    grid.SelectContextMenuItem(menu_item_id)
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
        """Select a row in a table/grid."""
        self._require_session()

        try:
            grid = self._session.findById(table_id)
            grid.selectedRows = str(row)

            return {
                "table_id": table_id,
                "selected_row": row,
                "status": "success",
            }
        except Exception as e:
            return {"table_id": table_id, "error": str(e)}

    def double_click_table_cell(self, table_id: str, row: int, column: str) -> Dict[str, Any]:
        """Double-click a cell in a table/grid."""
        self._require_session()

        try:
            grid = self._session.findById(table_id)
            grid.doubleClickCell(row, column)

            return {
                "table_id": table_id,
                "row": row,
                "column": column,
                "status": "double_clicked",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {"table_id": table_id, "error": str(e)}

    # =========================================================================
    # Tree Control Operations
    # =========================================================================

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

            # Get column names if available (TableTree / ColumnTree)
            columns = []
            try:
                col_count = tree.ColumnCount
                for i in range(col_count):
                    columns.append(tree.GetColumnTitleByNumber(i))
            except Exception:
                pass

            column_names = []
            try:
                col_count = tree.ColumnCount
                for i in range(col_count):
                    column_names.append(tree.GetColumnNameByNumber(i))
            except Exception:
                pass

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

                # Node text
                try:
                    node["text"] = tree.GetNodeTextByKey(key)
                except Exception:
                    node["text"] = ""

                # Parent key
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

                # Column values (for TableTree / ColumnTree)
                if column_names:
                    col_values = {}
                    for col_name in column_names:
                        try:
                            col_values[col_name] = tree.GetItemText(key, col_name)
                        except Exception:
                            col_values[col_name] = None
                    node["columns"] = col_values

                nodes.append(node)

            return {
                "tree_id": tree_id,
                "total_nodes": len(node_keys),
                "column_titles": columns,
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

            window = self._session.findById("wnd[0]")
            window.HardCopy(filepath, "PNG")

            if return_base64:
                import base64
                with open(filepath, "rb") as f:
                    data = base64.b64encode(f.read()).decode()
                os.remove(filepath)
                return {
                    "format": "png",
                    "encoding": "base64",
                    "data": data,
                }
            else:
                return {
                    "format": "png",
                    "filepath": filepath,
                }

        except Exception as e:
            return {"error": str(e)}
