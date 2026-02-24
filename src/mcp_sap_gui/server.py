"""
MCP Server for SAP GUI Scripting.

This module implements a Model Context Protocol (MCP) server that exposes
SAP GUI automation capabilities to AI assistants like Claude.

Uses FastMCP from the official MCP Python SDK for decorator-based tool
definitions with automatic JSON schema generation from type hints.

Security Note:
    This server provides powerful automation capabilities. Production deployments
    should implement:
    - Transaction whitelisting
    - Read-only modes
    - Audit logging
    - Rate limiting
    - User confirmation for write operations
"""

import asyncio
import base64
import logging
from concurrent.futures import ThreadPoolExecutor
from typing import Optional, List, Literal
from dataclasses import dataclass, field

from mcp.server.fastmcp import FastMCP, Image

from .sap_controller import (
    SAPGUIController,
    SAPGUIError,
    SAPGUINotAvailableError,
    SAPGUINotConnectedError,
    VKey,
    SessionInfo,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

@dataclass
class ServerConfig:
    """Configuration for the MCP SAP GUI server."""

    # Security settings
    read_only: bool = False
    allowed_transactions: Optional[List[str]] = None  # None = all allowed
    blocked_transactions: List[str] = field(default_factory=lambda: [
        "SU01", "SU10", "SU01D",  # User administration
        "PFCG", "SU53",           # Role administration
        "SM21", "ST22",           # System logs
        "SE16N",                  # Table maintenance
    ])
    require_confirmation_for_writes: bool = True

    # Behavior settings
    auto_connect_existing: bool = True
    default_language: str = "EN"
    max_table_rows: int = 500


# ---------------------------------------------------------------------------
# Module-level state
# ---------------------------------------------------------------------------

mcp = FastMCP("mcp-sap-gui")
controller: Optional[SAPGUIController] = None
config = ServerConfig()
_executor: Optional[ThreadPoolExecutor] = None


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

async def _com(fn):
    """Run a synchronous COM operation in the dedicated thread."""
    global _executor
    if _executor is None:
        _executor = ThreadPoolExecutor(max_workers=1)
    return await asyncio.get_running_loop().run_in_executor(_executor, fn)


def _check_write():
    """Raise if server is in read-only mode."""
    if config.read_only:
        raise ValueError("Write operations disabled in read-only mode")


def _is_transaction_blocked(tcode: str) -> bool:
    """Check if a transaction is blocked by security policy."""
    tcode_upper = tcode.upper().removeprefix("/N").removeprefix("/O")

    # Check blocklist
    if tcode_upper in config.blocked_transactions:
        return True

    # Check allowlist if configured
    if config.allowed_transactions is not None:
        return tcode_upper not in config.allowed_transactions

    return False


_KEY_MAP = {
    "Enter": VKey.ENTER,
    "F1": VKey.F1, "F2": VKey.F2, "F3": VKey.F3, "Back": VKey.F3,
    "F4": VKey.F4, "F5": VKey.F5, "Refresh": VKey.F5,
    "F6": VKey.F6, "F7": VKey.F7, "F8": VKey.F8, "Execute": VKey.F8,
    "F9": VKey.F9, "F10": VKey.F10, "F11": VKey.F11, "Save": VKey.F11,
    "F12": VKey.F12, "Cancel": VKey.F12,
    # Shift+F keys (common in SAP: Shift+F1=WhereUsed, Shift+F4=CloseAll, etc.)
    "Shift+F1": VKey.SHIFT_F1, "Shift+F2": VKey.SHIFT_F2,
    "Shift+F3": VKey.SHIFT_F3, "Shift+F4": VKey.SHIFT_F4,
    "Shift+F5": VKey.SHIFT_F5, "Shift+F6": VKey.SHIFT_F6,
    "Shift+F7": VKey.SHIFT_F7, "Shift+F8": VKey.SHIFT_F8,
    "Shift+F9": VKey.SHIFT_F9,
    # Ctrl combinations
    "Ctrl+F": VKey.CTRL_F, "Ctrl+G": VKey.CTRL_G, "Ctrl+P": VKey.CTRL_P,
}


def _parse_key(key: str) -> int:
    """Parse key name to VKey code."""
    if key not in _KEY_MAP:
        raise ValueError(
            f"Unknown key: '{key}'. Valid keys: {', '.join(_KEY_MAP.keys())}"
        )
    return _KEY_MAP[key]


def _to_dict(obj):
    """Convert dataclass to dict, pass dicts through unchanged."""
    return obj.__dict__ if hasattr(obj, '__dict__') and not isinstance(obj, dict) else obj


# ===========================================================================
# Connection tools
# ===========================================================================

@mcp.tool()
async def sap_connect(
    system_description: str,
    client: str | None = None,
    user: str | None = None,
    password: str | None = None,
    language: str | None = None,
) -> dict:
    """Connect to an SAP system by its name in SAP Logon Pad. Optionally provide credentials for automatic login."""
    kwargs: dict[str, str] = {"system_description": system_description}
    for key, val in [("client", client), ("user", user),
                     ("password", password), ("language", language)]:
        if val is not None:
            kwargs[key] = val
    return _to_dict(await _com(lambda: controller.connect(**kwargs)))


@mcp.tool()
async def sap_connect_existing(
    connection_index: int = 0,
    session_index: int = 0,
) -> dict:
    """Connect to an already open SAP session. Use this when SAP is already logged in."""
    return _to_dict(await _com(
        lambda: controller.connect_to_existing_session(connection_index, session_index)
    ))


@mcp.tool()
async def sap_list_connections() -> dict:
    """List all open SAP connections and sessions"""
    return await _com(controller.list_connections)


@mcp.tool()
async def sap_get_session_info() -> dict:
    """Get information about the current SAP session (system, client, user, transaction, screen)"""
    return _to_dict(await _com(controller.get_session_info))


# ===========================================================================
# Navigation tools
# ===========================================================================

@mcp.tool()
async def sap_execute_transaction(tcode: str) -> dict:
    """Execute an SAP transaction code (e.g., MM03, VA01, SE80)"""
    _check_write()
    if _is_transaction_blocked(tcode):
        return {"error": f"Transaction {tcode} is blocked by security policy"}
    return await _com(lambda: controller.execute_transaction(tcode))


@mcp.tool()
async def sap_send_key(
    key: Literal[
        "Enter", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8",
        "F9", "F10", "F11", "F12", "Back", "Save", "Cancel",
        "Execute", "Refresh",
        "Shift+F1", "Shift+F2", "Shift+F3", "Shift+F4", "Shift+F5",
        "Shift+F6", "Shift+F7", "Shift+F8", "Shift+F9",
        "Ctrl+F", "Ctrl+G", "Ctrl+P",
    ],
) -> dict:
    """Send a keyboard key. Common keys: Enter, F1 (Help), F3 (Back), F4 (Search help), F5 (Refresh), F8 (Execute), F11 (Save), F12 (Cancel). Also supports Shift+F1..F9 and Ctrl+F (Find), Ctrl+G (Continue search), Ctrl+P (Print)."""
    _check_write()
    vkey = _parse_key(key)
    return await _com(lambda: controller.send_vkey(vkey))


@mcp.tool()
async def sap_get_screen_info() -> dict:
    """Get information about the current SAP screen (transaction, program, screen number, title, status message).

    Reads from ``session.ActiveWindow`` so the response always reflects
    what the user sees.  The ``active_window`` field tells you which
    window is in focus (e.g. ``wnd[0]`` for the main screen, ``wnd[1]``
    for a popup).  The ``title`` comes from the active window.

    Every action tool (press_button, send_key, select_menu, etc.) returns
    this same screen info, so you always know when a popup appears.
    Use sap_get_popup_window for full popup content (texts, buttons)."""
    return await _com(controller.get_screen_info)


# ===========================================================================
# Field tools
# ===========================================================================

@mcp.tool()
async def sap_read_field(field_id: str) -> dict:
    """Read the value of a field on the current SAP screen"""
    return await _com(lambda: controller.read_field(field_id))


@mcp.tool()
async def sap_set_field(field_id: str, value: str) -> dict:
    """Set a value in a field on the current SAP screen"""
    _check_write()
    return await _com(lambda: controller.set_field(field_id, value))


@mcp.tool()
async def sap_press_button(button_id: str) -> dict:
    """Press a button on the current SAP screen"""
    _check_write()
    return await _com(lambda: controller.press_button(button_id))


@mcp.tool()
async def sap_select_menu(menu_id: str) -> dict:
    """Select a menu item from the menu bar (e.g., 'wnd[0]/mbar/menu[1]/menu[0]'). Use sap_get_screen_elements on 'wnd[0]/mbar' to discover available menus."""
    _check_write()
    return await _com(lambda: controller.select_menu(menu_id))


@mcp.tool()
async def sap_select_checkbox(checkbox_id: str, selected: bool = True) -> dict:
    """Select or deselect a checkbox on the current SAP screen"""
    _check_write()
    return await _com(lambda: controller.select_checkbox(checkbox_id, selected))


@mcp.tool()
async def sap_select_radio_button(radio_id: str) -> dict:
    """Select a radio button on the current SAP screen"""
    _check_write()
    return await _com(lambda: controller.select_radio_button(radio_id))


@mcp.tool()
async def sap_select_combobox_entry(combobox_id: str, key_or_value: str) -> dict:
    """Select an entry in a combobox/dropdown by its key or display value text"""
    _check_write()
    return await _com(lambda: controller.select_combobox_entry(combobox_id, key_or_value))


@mcp.tool()
async def sap_select_tab(tab_id: str) -> dict:
    """Select a tab in a tab strip control"""
    _check_write()
    return await _com(lambda: controller.select_tab(tab_id))


@mcp.tool()
async def sap_get_combobox_entries(combobox_id: str) -> dict:
    """List all entries in a combobox/dropdown. Returns key-value pairs so you know which values are valid."""
    return await _com(lambda: controller.get_combobox_entries(combobox_id))


@mcp.tool()
async def sap_set_batch_fields(fields: dict) -> dict:
    """Set multiple field values at once (dict of field_id → value). More efficient than repeated sap_set_field calls."""
    _check_write()
    return await _com(lambda: controller.set_batch_fields(fields))


@mcp.tool()
async def sap_read_textedit(textedit_id: str) -> dict:
    """Read the content of a multiline text editor (GuiTextedit). Returns full text, line count, and individual lines."""
    return await _com(lambda: controller.read_textedit(textedit_id))


@mcp.tool()
async def sap_set_textedit(textedit_id: str, text: str) -> dict:
    """Set the content of a multiline text editor (GuiTextedit)."""
    _check_write()
    return await _com(lambda: controller.set_textedit(textedit_id, text))


@mcp.tool()
async def sap_set_focus(element_id: str) -> dict:
    """Set focus to any screen element by its ID."""
    _check_write()
    return await _com(lambda: controller.set_focus(element_id))


# ===========================================================================
# Table tools
# ===========================================================================

@mcp.tool()
async def sap_read_table(table_id: str, max_rows: int = 100) -> dict:
    """Read data from an ALV grid or table on the current screen.

    Auto-detects the table type. The response includes a 'table_type' field
    ('GuiGridView' for ALV or 'GuiTableControl') so you know which
    type-specific tools to use next (e.g., sap_get_alv_toolbar for ALV,
    sap_scroll_table_control for TableControl)."""
    capped = min(max_rows, config.max_table_rows)
    return await _com(lambda: controller.read_table(table_id, capped))


@mcp.tool()
async def sap_get_alv_toolbar(grid_id: str) -> dict:
    """Get all toolbar buttons from an ALV grid. Returns button IDs, texts, and types. Use this to discover available actions."""
    return await _com(lambda: controller.get_alv_toolbar(grid_id))


@mcp.tool()
async def sap_press_alv_toolbar_button(grid_id: str, button_id: str) -> dict:
    """Press a toolbar button on an ALV grid (e.g., sort, filter, export, custom actions). Use sap_get_alv_toolbar to find button IDs."""
    _check_write()
    return await _com(lambda: controller.press_alv_toolbar_button(grid_id, button_id))


@mcp.tool()
async def sap_select_alv_context_menu_item(
    grid_id: str,
    menu_item_id: str,
    toolbar_button_id: str | None = None,
) -> dict:
    """Select an item from an opened ALV context menu. First use sap_press_alv_toolbar_button on a Menu button to open the menu, then use this to select an item. Alternatively, pass toolbar_button_id to open the menu and select the item in one atomic call (recommended)."""
    _check_write()
    return await _com(
        lambda: controller.select_alv_context_menu_item(
            grid_id, menu_item_id, toolbar_button_id
        )
    )


@mcp.tool()
async def sap_select_table_row(table_id: str, row: int) -> dict:
    """Select a row in a table/grid"""
    _check_write()
    return await _com(lambda: controller.select_table_row(table_id, row))


@mcp.tool()
async def sap_double_click_cell(table_id: str, row: int, column: str) -> dict:
    """Double-click a cell in a table/grid (often opens details)"""
    _check_write()
    return await _com(
        lambda: controller.double_click_table_cell(table_id, row, column)
    )


@mcp.tool()
async def sap_modify_cell(grid_id: str, row: int, column: str, value: str) -> dict:
    """Modify the value of a cell in an ALV grid or table control (e.g., for editable grids)"""
    _check_write()
    return await _com(lambda: controller.modify_cell(grid_id, row, column, value))


@mcp.tool()
async def sap_set_current_cell(grid_id: str, row: int, column: str) -> dict:
    """Set the current (focused) cell in an ALV grid or table control"""
    _check_write()
    return await _com(lambda: controller.set_current_cell(grid_id, row, column))


@mcp.tool()
async def sap_get_column_info(grid_id: str) -> dict:
    """Get detailed column information from an ALV grid or table control (names, titles, tooltips)"""
    return await _com(lambda: controller.get_column_info(grid_id))


@mcp.tool()
async def sap_get_current_cell(table_id: str) -> dict:
    """Get the currently focused cell position in an ALV grid or table control."""
    return await _com(lambda: controller.get_current_cell(table_id))


# ---- TableControl-specific tools ----

@mcp.tool()
async def sap_scroll_table_control(table_id: str, position: int) -> dict:
    """Scroll a GuiTableControl to a specific row position. Use before sap_read_table to navigate to different sections. Does NOT work on ALV grids (they handle scrolling internally)."""
    _check_write()
    return await _com(lambda: controller.scroll_table_control(table_id, position))


@mcp.tool()
async def sap_get_table_control_row_info(
    table_id: str,
    rows: list[int] | None = None,
) -> dict:
    """Get row metadata (selectable, selected) from a GuiTableControl. If rows is omitted, queries all visible rows. Does NOT work on ALV grids."""
    return await _com(lambda: controller.get_table_control_row_info(table_id, rows))


@mcp.tool()
async def sap_select_all_table_control_columns(
    table_id: str,
    select: bool = True,
) -> dict:
    """Select or deselect all columns in a GuiTableControl. Does NOT work on ALV grids."""
    _check_write()
    return await _com(lambda: controller.select_all_table_control_columns(table_id, select))


# ---- ALV-specific tools ----

@mcp.tool()
async def sap_get_cell_info(grid_id: str, row: int, column: str) -> dict:
    """Get detailed cell metadata from an ALV grid: value, changeable, color, tooltip, style, max_length. Does NOT work on GuiTableControl."""
    return await _com(lambda: controller.get_cell_info(grid_id, row, column))


@mcp.tool()
async def sap_press_column_header(grid_id: str, column: str) -> dict:
    """Click a column header in an ALV grid (triggers sort). Does NOT work on GuiTableControl."""
    _check_write()
    return await _com(lambda: controller.press_column_header(grid_id, column))


@mcp.tool()
async def sap_select_all_rows(grid_id: str) -> dict:
    """Select all rows in an ALV grid. Does NOT work on GuiTableControl."""
    _check_write()
    return await _com(lambda: controller.select_all_rows(grid_id))


@mcp.tool()
async def sap_select_multiple_rows(table_id: str, rows: list[int]) -> dict:
    """Select multiple rows at once in an ALV grid or table control. Pass a list of row indices (e.g., [0, 2, 5])."""
    _check_write()
    return await _com(lambda: controller.select_multiple_rows(table_id, rows))


# ---- Popup & dialog tools ----

@mcp.tool()
async def sap_get_popup_window() -> dict:
    """Check if a popup/modal dialog is open (wnd[1], wnd[2], etc.). Returns the popup's title, text content, and available buttons so you know how to respond. Returns {popup_exists: false} if no popup."""
    return await _com(controller.get_popup_window)


# ---- Toolbar discovery ----

@mcp.tool()
async def sap_get_toolbar_buttons(window_id: str = "wnd[0]") -> dict:
    """List all buttons on the system toolbar (tbar[0]) and application toolbar (tbar[1]). Returns button IDs, text, tooltip, and enabled state. Useful for discovering available actions on a screen. This is for standard SAP toolbars, NOT ALV toolbars (use sap_get_alv_toolbar for ALV)."""
    return await _com(lambda: controller.get_toolbar_buttons(window_id))


# ---- Shell content ----

@mcp.tool()
async def sap_read_shell_content(shell_id: str) -> dict:
    """Read content from a GuiShell subtype (e.g., HTMLViewer). Extracts HTML, URL, or text depending on the shell type. Use sap_get_screen_elements first to find shell element IDs."""
    return await _com(lambda: controller.read_shell_content(shell_id))


# ===========================================================================
# Tree tools
# ===========================================================================

@mcp.tool()
async def sap_read_tree(tree_id: str, max_nodes: int = 200) -> dict:
    """Read data from a tree control (SAP.TableTreeControl, SAP.ColumnTreeControl, etc.). Returns node hierarchy with texts and column values."""
    capped = min(max_nodes, config.max_table_rows)
    return await _com(lambda: controller.read_tree(tree_id, capped))


@mcp.tool()
async def sap_expand_tree_node(tree_id: str, node_key: str) -> dict:
    """Expand a folder node in a tree control to reveal its children"""
    _check_write()
    return await _com(lambda: controller.expand_tree_node(tree_id, node_key))


@mcp.tool()
async def sap_collapse_tree_node(tree_id: str, node_key: str) -> dict:
    """Collapse a folder node in a tree control"""
    _check_write()
    return await _com(lambda: controller.collapse_tree_node(tree_id, node_key))


@mcp.tool()
async def sap_select_tree_node(tree_id: str, node_key: str) -> dict:
    """Select a node in a tree control"""
    _check_write()
    return await _com(lambda: controller.select_tree_node(tree_id, node_key))


@mcp.tool()
async def sap_double_click_tree_node(tree_id: str, node_key: str) -> dict:
    """Double-click a node in a tree control (often opens details or drills down)"""
    _check_write()
    return await _com(lambda: controller.double_click_tree_node(tree_id, node_key))


@mcp.tool()
async def sap_double_click_tree_item(tree_id: str, node_key: str, item_name: str) -> dict:
    """Double-click a specific item (column cell) in a tree node row"""
    _check_write()
    return await _com(
        lambda: controller.double_click_tree_item(tree_id, node_key, item_name)
    )


@mcp.tool()
async def sap_click_tree_link(tree_id: str, node_key: str, item_name: str) -> dict:
    """Click a hyperlink in a tree node item"""
    _check_write()
    return await _com(
        lambda: controller.click_tree_link(tree_id, node_key, item_name)
    )


@mcp.tool()
async def sap_find_tree_node_by_path(tree_id: str, path: str) -> dict:
    """Find a tree node key by its path (e.g., '2\\1\\2' = 2nd child of root, then 1st child, then 2nd child)"""
    return await _com(lambda: controller.find_tree_node_by_path(tree_id, path))


# ===========================================================================
# Discovery tools
# ===========================================================================

@mcp.tool()
async def sap_get_screen_elements(container_id: str = "wnd[0]/usr") -> dict:
    """Discover all elements on the current SAP screen. Useful for finding field IDs when working with a new screen."""
    elements = await _com(
        lambda: controller.get_screen_elements(container_id)
    )
    return {
        "element_count": len(elements),
        "elements": [e.__dict__ for e in elements],
    }


@mcp.tool()
async def sap_screenshot() -> Image:
    """Take a screenshot of the current SAP window"""
    result = await _com(controller.take_screenshot)
    if "error" in result:
        raise ValueError(result["error"])
    return Image(data=base64.b64decode(result["data"]), format="png")


# ===========================================================================
# Main entry point
# ===========================================================================

def main():
    """Main entry point."""
    import argparse
    global controller, config

    parser = argparse.ArgumentParser(description="MCP Server for SAP GUI")
    parser.add_argument("--read-only", action="store_true",
                        help="Run in read-only mode (no write operations)")
    parser.add_argument("--allowed-transactions", nargs="*",
                        help="Whitelist of allowed transaction codes")
    parser.add_argument("--debug", action="store_true",
                        help="Enable debug logging")
    args = parser.parse_args()

    if args.debug:
        logging.basicConfig(level=logging.DEBUG)
    else:
        logging.basicConfig(level=logging.INFO)

    config = ServerConfig(
        read_only=args.read_only,
        allowed_transactions=args.allowed_transactions,
    )
    controller = SAPGUIController()

    mcp.run()


if __name__ == "__main__":
    main()
