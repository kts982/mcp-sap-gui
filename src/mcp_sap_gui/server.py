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
import time
from dataclasses import dataclass, field
from typing import List, Literal, Optional

from fastmcp import Context, FastMCP
from fastmcp.server.lifespan import lifespan
from fastmcp.utilities.types import Image

from .sap_controller import VKey
from .session_manager import SessionManager

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

    # Behavior settings
    auto_connect_existing: bool = True
    default_language: str = "EN"
    max_table_rows: int = 500

    def __post_init__(self):
        self.blocked_transactions = [t.upper() for t in self.blocked_transactions]
        if self.allowed_transactions is not None:
            self.allowed_transactions = [t.upper() for t in self.allowed_transactions]


# ---------------------------------------------------------------------------
# Tool annotation presets
# ---------------------------------------------------------------------------

_READ_ONLY = {"readOnlyHint": True, "destructiveHint": False}
_WRITE = {"readOnlyHint": False, "destructiveHint": False}
_DESTRUCTIVE = {"readOnlyHint": False, "destructiveHint": True}


# ---------------------------------------------------------------------------
# Lifespan
# ---------------------------------------------------------------------------

@lifespan
async def _lifespan(server: FastMCP):
    """Create the session manager on startup, clean up on shutdown."""
    global _session_mgr
    _session_mgr = SessionManager()
    try:
        yield {"session_mgr": _session_mgr}
    finally:
        if _session_mgr is not None:
            try:
                _session_mgr.executor.submit(
                    _session_mgr.release_all
                ).result(timeout=10)
            except Exception:
                pass
            _session_mgr.shutdown()
            _session_mgr = None


# ---------------------------------------------------------------------------
# MCP Instructions (injected into every client's system prompt)
# ---------------------------------------------------------------------------

_INSTRUCTIONS = """\
SAP GUI automation server. Controls SAP GUI for Windows via COM scripting.

## Getting Started

1. If SAP is already open and logged in, use `sap_connect_existing` (most common).
2. If SAP is not open, use `sap_connect` with the system name from SAP Logon Pad.
3. After connecting, use `sap_get_session_info` to see the current transaction/screen.

## Screen Discovery (use this workflow on every new screen)

1. `sap_get_screen_elements` — discovers all element IDs on the current screen. \
Use `type_filter` (e.g. "GuiTextField,GuiCTextField") and `changeable_only=true` \
to reduce output on complex screens.
2. `sap_get_toolbar_buttons` — lists toolbar button IDs (system bar + app bar).
3. `sap_read_field` — reads a specific field's value and labels.

**CRITICAL: Never guess or hallucinate element IDs.** Always discover them first \
with the tools above. SAP field IDs vary across systems and customizations.

## Popup Handling

Every action tool returns `screen.active_window`. If it is `wnd[1]` or higher, \
a popup appeared. Call `sap_get_popup_window` to read popup text and buttons. \
Always check for popups after actions that may trigger them (button presses, \
Enter key, transaction execution, double-clicks).

## Tables: ALV Grid vs TableControl

`sap_read_table` auto-detects the type. The response includes `table_type`:
- **GuiGridView (ALV)**: Standard for reports. All rows accessible via `start_row`. \
Use `sap_get_alv_toolbar` for toolbar actions.
- **GuiTableControl**: Standard for customizing (SPRO, SM30). Visible rows only.

**Pagination**: Use `start_row` parameter to paginate through large tables. \
Check `total_rows` in the response to know when you've read everything.

**Schema first**: Use `columns_only=true` to see column names before reading data. \
Then use `columns="COL_A,COL_B"` to fetch only what you need.

**Table maintenance views (SM30-style)**: These use GuiTableControl. If you need \
to jump to a specific entry, look for a "Position..." button in the toolbar \
(`sap_get_toolbar_buttons`) — it opens a search dialog. Don't manually scroll.

## SPRO / Customizing Tree Navigation

SPRO uses a deep tree (1000+ nodes). **Do NOT use `sap_read_tree`** — it's too slow.

1. `sap_search_tree_nodes` — find nodes by text (e.g. "Storage Type"). \
**Note**: only searches already-expanded nodes. If you don't find a match, \
expand parent nodes first with `sap_get_tree_node_children(expand=true)`, \
then search again.
2. `sap_get_tree_node_children` with `expand=true` — step through the hierarchy.
3. **To execute an activity**: use `sap_click_tree_link(tree_id, node_key, "2")` \
(column "2" is the execute icon). Do NOT use `sap_double_click_tree_node` — that \
opens documentation, not the activity.
4. After clicking, check for popups (selection screens often appear as dialogs).

## Key Reference

- **Enter**: Confirm / continue
- **F3 / Back**: Go back one screen
- **F4**: Open search help / dropdown (set focus on field first with `sap_set_focus`)
- **F5 / Refresh**: Context-dependent — in table maintenance views this means \
"New Entries", NOT refresh. Check toolbar button tooltips first.
- **F8 / Execute**: Run report / execute selection
- **F11 / Save**: Save (Ctrl+S equivalent)
- **F12 / Cancel**: Cancel / close dialog

## Common Mistakes to Avoid

- **Guessing IDs**: Always use `sap_get_screen_elements` or `sap_get_toolbar_buttons` first.
- **Ignoring popups**: Check `active_window` in every action response.
- **F5 in table maintenance**: F5 means "New Entries", not refresh.
- **double_click_tree_node in SPRO**: Opens docs, not activities. Use `click_tree_link`.
- **Scrolling TableControls manually**: Use `start_row` in `sap_read_table` or \
the "Position..." button instead.
- **Reading huge trees**: Use `search_tree_nodes` + `get_tree_node_children`, \
not `read_tree` on large trees like SPRO.
"""


# ---------------------------------------------------------------------------
# MCP Resource: detailed SAP GUI reference guide
# ---------------------------------------------------------------------------

_SAP_GUI_GUIDE = """\
# SAP GUI Navigation Reference for AI Agents

Detailed reference guide for automating SAP GUI via MCP tools.

## Element Types and ID Patterns

| Type | Description | Typical ID Pattern |
|------|-------------|-------------------|
| GuiTextField | Text input | `wnd[0]/usr/txtFIELD_NAME` |
| GuiCTextField | Text with search help (F4) | `wnd[0]/usr/ctxtFIELD_NAME` |
| GuiPasswordField | Password input | `wnd[0]/usr/pwdFIELD_NAME` |
| GuiButton | Pushbutton | `wnd[0]/tbar[1]/btn[8]` or `wnd[0]/usr/btnBUTTON_NAME` |
| GuiCheckBox | Checkbox | `wnd[0]/usr/chkFIELD_NAME` |
| GuiRadioButton | Radio button | `wnd[0]/usr/radFIELD_NAME` |
| GuiComboBox | Dropdown | `wnd[0]/usr/cmbFIELD_NAME` |
| GuiTab | Tab in strip | `wnd[0]/usr/tabsTABSTRIP/tabpTAB01` |
| GuiMenu | Menu bar item | `wnd[0]/mbar/menu[0]/menu[1]` |
| GuiGridView | ALV Grid | `wnd[0]/usr/cntlGRID/shellcont/shell` |
| GuiTableControl | Classic table | `wnd[0]/usr/tblSAPLBD41TCTRL_V_TBDLS` |
| GuiTree | Tree control | `wnd[0]/usr/cntlTREE/shellcont/shell` |
| GuiStatusbar | Status bar | `wnd[0]/sbar` |
| GuiOkCodeField | Command/OK-code field | `wnd[0]/tbar[0]/okcd` |

## ID Naming Conventions

- `wnd[0]` = main window, `wnd[1]` = first popup, `wnd[2]` = second popup
- `usr` = user area (screen body)
- `tbar[0]` = system toolbar (top), `tbar[1]` = application toolbar
- `mbar` = menu bar
- `btn[N]` = toolbar button by position
- `shell/shellcont/shell` = container for ALV grids and tree controls
- `txt` prefix = GuiTextField, `ctxt` = GuiCTextField, `chk` = GuiCheckBox
- `rad` = GuiRadioButton, `cmb` = GuiComboBox, `pwd` = GuiPasswordField

## Transaction Code Conventions

- Standard: `MM03` (display material), `VA03` (display sales order)
- With /n prefix: `/nMM03` — closes current transaction, opens new one
- With /o prefix: `/oMM03` — opens in new session (new window)
- SCWM transactions need /n: `/n/SCWM/MON`, `/n/SCWM/PRDO`

## Table Type Comparison

### GuiGridView (ALV)
- Used in: Reports, list displays, transaction results
- ID contains: `shellcont/shell`
- All rows accessible directly (internal scrolling)
- Has built-in toolbar: sort, filter, export, layout
- Use `sap_get_alv_toolbar` for toolbar button discovery
- Use `sap_press_alv_toolbar_button` to execute toolbar actions
- Pagination: `start_row` parameter in `sap_read_table`

### GuiTableControl
- Used in: SPRO customizing, SM30 table maintenance, config screens
- ID starts with: `tbl`
- Only visible rows accessible at a time
- Scrolling via `VerticalScrollbar.Position`
- Pagination: `start_row` in `sap_read_table` (handles scrolling internally)
- "Position..." button for jumping to specific entries
- F5 = "New Entries" (NOT refresh!)

## Reading Status Bar Messages

The status bar (`wnd[0]/sbar`) shows messages after actions:
- **S** (Success): Operation completed successfully
- **E** (Error): Operation failed — read the message for details
- **W** (Warning): Warning that may need acknowledgment
- **I** (Info): Informational message

Use `sap_read_field("wnd[0]/sbar")` to read the status bar after actions.

## SPRO Navigation Pattern (Step-by-Step)

SPRO (transaction SPRO) is SAP's customizing tree with 1000+ nodes.

1. Execute transaction: `sap_execute_transaction("SPRO")`
2. Click "SAP Reference IMG" button (check toolbar with `sap_get_toolbar_buttons`)
3. Search for your target: `sap_search_tree_nodes(tree_id, "your search text")`
   - **Limitation**: only searches already-expanded/loaded nodes
   - If no results, expand parent nodes with `sap_get_tree_node_children(expand=true)` first
4. Navigate to the node using `sap_get_tree_node_children` with `expand=true`
5. Execute the activity: `sap_click_tree_link(tree_id, node_key, "2")`
   - Column "2" is the execute/activity icon in SPRO trees
   - Do NOT use `double_click_tree_node` — that opens documentation
6. Handle the resulting screen (often a popup selection screen or table maintenance view)

## Table Maintenance Views (SM30-style)

These appear when you execute SPRO activities or use SM30:

1. Usually shows a selection screen first (popup on wnd[1])
2. After execution, shows a GuiTableControl with configuration entries
3. Toolbar has: "New Entries", "Position...", "Delete", "Save"
4. To find specific entries, use "Position..." button (opens search dialog)
5. To read all entries, use `sap_read_table` with `start_row` pagination
6. Check `total_rows` in response to know total entry count

## Web Dynpro Screens

Some newer SAP screens use Web Dynpro technology:
- `sap_get_screen_elements` may return 0 elements on wnd[0]/usr
- Use `sap_screenshot` as fallback to see what's displayed
- Element IDs follow different patterns (deeper nesting)
- Try increasing `max_depth` in `sap_get_screen_elements`

## Handling Splitter Layouts

Some transactions use split-screen layouts:
- Elements are nested inside `shellcont[0]`, `shellcont[1]`, etc.
- Use `sap_get_screen_elements` with `max_depth=3` or higher
- Example: Warehouse Monitor `/SCWM/MON` uses splitter with tree in `shellcont[0]`
"""


# ---------------------------------------------------------------------------
# Module-level state
# ---------------------------------------------------------------------------

mcp = FastMCP("mcp-sap-gui", instructions=_INSTRUCTIONS, lifespan=_lifespan)
_session_mgr: Optional[SessionManager] = None
config = ServerConfig()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# Known transient COM error codes that are safe to retry
_TRANSIENT_COM_HRESULTS = {
    -2147417851,  # RPC_E_CALL_REJECTED (0x80010001) — "Call was rejected by callee"
    -2147418111,  # RPC_E_DISCONNECTED (0x80010108) — object disconnected
    -2147417848,  # RPC_E_RETRY (0x80010109) — "The callee is not available"
}

_COM_MAX_RETRIES = 3
_COM_BASE_DELAY = 0.3  # seconds


def _is_transient_com_error(exc: Exception) -> bool:
    """Check if a COM exception is a known transient error safe to retry."""
    hresult = getattr(exc, "hresult", None)
    if hresult is not None and hresult in _TRANSIENT_COM_HRESULTS:
        return True
    # pywintypes.com_error stores hresult in args[0]
    if exc.args and isinstance(exc.args[0], int) and exc.args[0] in _TRANSIENT_COM_HRESULTS:
        return True
    return False


def _com_with_retry(fn):
    """Run a COM operation with retry for transient errors.

    Runs synchronously in the COM thread. Retries up to _COM_MAX_RETRIES
    times with exponential backoff for known transient COM errors.
    """
    last_exc = None
    for attempt in range(_COM_MAX_RETRIES + 1):
        try:
            return fn()
        except Exception as e:
            if attempt < _COM_MAX_RETRIES and _is_transient_com_error(e):
                delay = _COM_BASE_DELAY * (2 ** attempt)
                logger.warning(
                    "Transient COM error (attempt %d/%d), retrying in %.1fs: %s",
                    attempt + 1, _COM_MAX_RETRIES + 1, delay, e,
                )
                time.sleep(delay)
                last_exc = e
            else:
                raise
    raise last_exc  # pragma: no cover — unreachable, loop always raises


async def _com(fn):
    """Run a synchronous COM operation in the dedicated thread.

    Automatically retries known transient COM errors (e.g. RPC_E_CALL_REJECTED)
    up to 3 times with exponential backoff.
    """
    if _session_mgr is None:
        raise RuntimeError("Server not initialised – lifespan has not started")
    return await asyncio.get_running_loop().run_in_executor(
        _session_mgr.executor, lambda: _com_with_retry(fn)
    )


def _session_key(ctx: Context) -> int:
    """Derive a stable key for the current MCP client session."""
    return id(ctx.session)


def _ctrl(ctx: Context):
    """Get the SAPGUIController bound to the current MCP session."""
    return _session_mgr.get_or_create(_session_key(ctx)).controller


def _check_write():
    """Raise if server is in read-only mode."""
    if config.read_only:
        raise ValueError("Write operations disabled in read-only mode")


def _is_okcode_field(field_id: str) -> bool:
    """Return True when a field ID targets an SAP command field."""
    normalized = field_id.replace("\\", "/").strip().lower()
    return normalized.endswith("/okcd")


def _check_okcode_bypass(field_id: str, value: str):
    """Block attempts to execute blocked transactions via the OK-code field."""
    if _is_okcode_field(field_id) and _is_transaction_blocked(value):
        raise ValueError(
            f"Transaction {value} is blocked by security policy "
            f"(attempted via OK-code field)"
        )


def _is_transaction_blocked(tcode: str) -> bool:
    """Check if a transaction is blocked by security policy."""
    tcode_upper = (
        tcode.strip().upper()
        .removeprefix("/N").removeprefix("/O").removeprefix("/*")
    )

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

@mcp.tool(annotations=_READ_ONLY)
async def sap_connect(
    system_description: str,
    ctx: Context,
    client: str | None = None,
    user: str | None = None,
    language: str | None = None,
) -> dict:
    """Connect to an SAP system by its name in SAP Logon Pad.

    For security, this tool does not accept a password parameter.
    If SAP is already open and logged in, use sap_connect_existing instead."""
    c = _ctrl(ctx)
    kwargs: dict[str, str] = {"system_description": system_description}
    for key, val in [("client", client), ("user", user), ("language", language)]:
        if val is not None:
            kwargs[key] = val
    return _to_dict(await _com(lambda: c.connect(**kwargs)))


@mcp.tool(annotations=_READ_ONLY)
async def sap_connect_existing(
    ctx: Context,
    connection_index: int = 0,
    session_index: int = 0,
) -> dict:
    """Connect to an already open SAP session. Use this when SAP is already logged in.

    This is the most common starting point. connection_index=0 and session_index=0
    connect to the first open session. Use sap_list_connections to see all sessions."""
    c = _ctrl(ctx)
    return _to_dict(await _com(
        lambda: c.connect_to_existing_session(connection_index, session_index)
    ))


@mcp.tool(annotations=_READ_ONLY)
async def sap_list_connections(ctx: Context) -> dict:
    """List all open SAP connections and sessions"""
    c = _ctrl(ctx)
    connections = await _com(c.list_connections)
    return {"connections": connections}


@mcp.tool(annotations=_READ_ONLY)
async def sap_get_session_info(ctx: Context) -> dict:
    """Get information about the current SAP session (system, client, user, transaction, screen)"""
    c = _ctrl(ctx)
    return _to_dict(await _com(c.get_session_info))


# ===========================================================================
# Navigation tools
# ===========================================================================

@mcp.tool(annotations=_DESTRUCTIVE)
async def sap_execute_transaction(tcode: str, ctx: Context) -> dict:
    """Execute an SAP transaction code (e.g., MM03, VA01, SE80).

    Navigates to the transaction's initial screen. Always check the
    screen info in the response to understand what screen you landed on.
    Some transactions require /n prefix for SCWM (e.g., /n/SCWM/MON)."""
    _check_write()
    if _is_transaction_blocked(tcode):
        raise ValueError(f"Transaction {tcode} is blocked by security policy")
    c = _ctrl(ctx)
    return await _com(lambda: c.execute_transaction(tcode))


@mcp.tool(annotations=_WRITE)
async def sap_send_key(
    key: Literal[
        "Enter", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8",
        "F9", "F10", "F11", "F12", "Back", "Save", "Cancel",
        "Execute", "Refresh",
        "Shift+F1", "Shift+F2", "Shift+F3", "Shift+F4", "Shift+F5",
        "Shift+F6", "Shift+F7", "Shift+F8", "Shift+F9",
        "Ctrl+F", "Ctrl+G", "Ctrl+P",
    ],
    ctx: Context,
) -> dict:
    """Send a keyboard key.

    Common keys: Enter, F1 (Help), F3 (Back), F4 (Search help),
    F5 (Refresh), F8 (Execute), F11 (Save), F12 (Cancel).
    Also supports Shift+F1..F9 and Ctrl+F, Ctrl+G, Ctrl+P."""
    _check_write()
    c = _ctrl(ctx)
    vkey = _parse_key(key)
    return await _com(lambda: c.send_vkey(vkey))


@mcp.tool(annotations=_READ_ONLY)
async def sap_get_screen_info(ctx: Context) -> dict:
    """Get current SAP screen info (transaction, program, screen number, title, status).

    Reads from ``session.ActiveWindow`` so the response always reflects
    what the user sees.  The ``active_window`` field tells you which
    window is in focus (e.g. ``wnd[0]`` for the main screen, ``wnd[1]``
    for a popup).  The ``title`` comes from the active window.

    Every action tool (press_button, send_key, select_menu, etc.) returns
    this same screen info, so you always know when a popup appears.
    Use sap_get_popup_window for full popup content (texts, buttons)."""
    c = _ctrl(ctx)
    return await _com(c.get_screen_info)


# ===========================================================================
# Field tools
# ===========================================================================

@mcp.tool(annotations=_READ_ONLY)
async def sap_read_field(field_id: str, ctx: Context) -> dict:
    """Read the value of a field on the current SAP screen.

    Returns value, type, changeable status, and labels (left/right).
    Use sap_get_screen_elements to discover field IDs on unknown screens."""
    c = _ctrl(ctx)
    return await _com(lambda: c.read_field(field_id))


@mcp.tool(annotations=_WRITE)
async def sap_set_field(field_id: str, value: str, ctx: Context) -> dict:
    """Set a value in a field on the current SAP screen.

    For setting multiple fields at once, use sap_set_batch_fields instead.
    After setting a field, you may need to press Enter to trigger validation."""
    _check_write()
    _check_okcode_bypass(field_id, value)
    c = _ctrl(ctx)
    return await _com(lambda: c.set_field(field_id, value))


@mcp.tool(annotations=_WRITE)
async def sap_press_button(button_id: str, ctx: Context) -> dict:
    """Press a button on the current SAP screen.

    Returns screen info after the press so you can detect navigation or popups.
    Use sap_get_toolbar_buttons to discover toolbar button IDs.
    Use sap_get_screen_elements to find on-screen button IDs."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.press_button(button_id))


@mcp.tool(annotations=_WRITE)
async def sap_select_menu(menu_id: str, ctx: Context) -> dict:
    """Select a menu item from the menu bar or a submenu.

    Example: 'wnd[0]/mbar/menu[1]/menu[0]'.
    Use sap_get_screen_elements on 'wnd[0]/mbar' to discover menu structure.
    Returns screen info after selection so you can detect navigation."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.select_menu(menu_id))


@mcp.tool(annotations=_WRITE)
async def sap_select_checkbox(checkbox_id: str, ctx: Context, selected: bool = True) -> dict:
    """Select or deselect a checkbox on the current SAP screen.

    Set selected=false to uncheck. Use sap_get_screen_elements with
    type_filter='GuiCheckBox' to find checkbox IDs."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.select_checkbox(checkbox_id, selected))


@mcp.tool(annotations=_WRITE)
async def sap_select_radio_button(radio_id: str, ctx: Context) -> dict:
    """Select a radio button on the current SAP screen.

    Use sap_get_screen_elements with type_filter='GuiRadioButton'
    to find radio button IDs on the current screen."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.select_radio_button(radio_id))


@mcp.tool(annotations=_WRITE)
async def sap_select_combobox_entry(combobox_id: str, key_or_value: str, ctx: Context) -> dict:
    """Select an entry in a combobox/dropdown by its key or display value text.

    Accepts either the technical key or the visible display text.
    Use sap_get_combobox_entries first to see all valid options."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.select_combobox_entry(combobox_id, key_or_value))


@mcp.tool(annotations=_WRITE)
async def sap_select_tab(tab_id: str, ctx: Context) -> dict:
    """Select a tab in a tab strip control.

    Returns screen info after selection (tab content changes).
    Tab IDs typically look like 'wnd[0]/usr/tabsTABSTRIP/tabpTAB01'."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.select_tab(tab_id))


@mcp.tool(annotations=_READ_ONLY)
async def sap_get_combobox_entries(combobox_id: str, ctx: Context) -> dict:
    """List all entries in a combobox/dropdown.

    Returns key-value pairs so you know which values are valid."""
    c = _ctrl(ctx)
    return await _com(lambda: c.get_combobox_entries(combobox_id))


@mcp.tool(annotations=_WRITE)
async def sap_set_batch_fields(fields: dict, ctx: Context) -> dict:
    """Set multiple field values at once (dict of field_id to value).

    More efficient than repeated sap_set_field calls."""
    _check_write()
    for fid, val in fields.items():
        _check_okcode_bypass(fid, str(val))
    c = _ctrl(ctx)
    return await _com(lambda: c.set_batch_fields(fields))


@mcp.tool(annotations=_READ_ONLY)
async def sap_read_textedit(textedit_id: str, ctx: Context, max_lines: int = 0) -> dict:
    """Read the content of a multiline text editor (GuiTextedit).

    Returns full text and line count. Use max_lines to cap output for
    large text editors (0 = all lines)."""
    c = _ctrl(ctx)
    return await _com(lambda: c.read_textedit(textedit_id, max_lines))


@mcp.tool(annotations=_WRITE)
async def sap_set_textedit(textedit_id: str, text: str, ctx: Context) -> dict:
    """Set the content of a multiline text editor (GuiTextedit)."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.set_textedit(textedit_id, text))


@mcp.tool(annotations=_WRITE)
async def sap_set_focus(element_id: str, ctx: Context) -> dict:
    """Set focus to any screen element by its ID.

    Some SAP actions require focus on a specific element before they work
    (e.g., F4 search help on a field). Use this to set focus before
    sending keys with sap_send_key."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.set_focus(element_id))


# ===========================================================================
# Table tools
# ===========================================================================

@mcp.tool(annotations=_READ_ONLY)
async def sap_read_table(
    table_id: str,
    ctx: Context,
    max_rows: int = 100,
    columns: str = "",
    columns_only: bool = False,
    start_row: int = 0,
) -> dict:
    """Read data from an ALV grid or table on the current screen.

    Auto-detects the table type. The response includes a 'table_type' field
    ('GuiGridView' for ALV or 'GuiTableControl') so you know which
    type-specific tools to use next (e.g., sap_get_alv_toolbar for ALV,
    sap_scroll_table_control for TableControl).

    Use columns_only=true for schema discovery (returns column metadata
    only, no data). Use columns to fetch only specific columns (CSV).
    Use start_row to paginate through large tables."""
    c = _ctrl(ctx)
    capped = min(max_rows, config.max_table_rows)
    return await _com(lambda: c.read_table(
        table_id, capped, columns=columns,
        columns_only=columns_only, start_row=start_row,
    ))


@mcp.tool(annotations=_READ_ONLY)
async def sap_get_alv_toolbar(grid_id: str, ctx: Context) -> dict:
    """Get all toolbar buttons from an ALV grid.

    Returns button IDs, texts, and types. Use this to discover available
    actions (sort, filter, export, etc.) before pressing them with
    sap_press_alv_toolbar_button. Only works on GuiGridView (ALV),
    not GuiTableControl."""
    c = _ctrl(ctx)
    return await _com(lambda: c.get_alv_toolbar(grid_id))


@mcp.tool(annotations=_WRITE)
async def sap_press_alv_toolbar_button(grid_id: str, button_id: str, ctx: Context) -> dict:
    """Press a toolbar button on an ALV grid (e.g., sort, filter, export).

    Use sap_get_alv_toolbar to find button IDs."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.press_alv_toolbar_button(grid_id, button_id))


@mcp.tool(annotations=_WRITE)
async def sap_select_alv_context_menu_item(
    grid_id: str,
    menu_item_id: str,
    ctx: Context,
    toolbar_button_id: str | None = None,
    select_by: Literal["auto", "id", "text", "position"] = "auto",
) -> dict:
    """Select an item from an opened ALV context menu.

    First use sap_press_alv_toolbar_button on a Menu button to open it,
    then use this to select an item. Alternatively, pass toolbar_button_id
    to open the menu and select in one atomic call (recommended).

    `menu_item_id` can be a technical function code (for example `@M00006`),
    visible menu text (for example `Confirm WT in Foreground`), or a position
    descriptor. You can pass human-readable menu text directly.
    SAP GUI does not expose a way to enumerate GuiGridView context menu items.

    `select_by` controls how selection is performed:
    - `auto` (default): heuristic based on whether `menu_item_id` has spaces
    - `id`: function code
    - `text`: visible menu text
    - `position`: position descriptor
    """
    _check_write()
    c = _ctrl(ctx)
    return await _com(
        lambda: c.select_alv_context_menu_item(
            grid_id, menu_item_id, toolbar_button_id, select_by
        )
    )


@mcp.tool(annotations=_WRITE)
async def sap_select_table_row(table_id: str, row: int, ctx: Context) -> dict:
    """Select a row in a table/grid.

    Works on both ALV grids and table controls. Row index is zero-based.
    For ALV: uses absolute row index. For TableControl: scrolls to make
    the row visible first if needed."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.select_table_row(table_id, row))


@mcp.tool(annotations=_WRITE)
async def sap_double_click_cell(table_id: str, row: int, column: str, ctx: Context) -> dict:
    """Double-click a cell in a table/grid (often opens details or drills down).

    Row is zero-based. Column is the column name (from sap_read_table
    or sap_get_column_info). Works on both ALV and TableControl."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(
        lambda: c.double_click_table_cell(table_id, row, column)
    )


@mcp.tool(annotations=_WRITE)
async def sap_modify_cell(grid_id: str, row: int, column: str, value: str, ctx: Context) -> dict:
    """Modify the value of a cell in an ALV grid or table control.

    Only works on editable cells. Use sap_get_cell_info to check if
    a cell is changeable before attempting to modify it."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.modify_cell(grid_id, row, column, value))


@mcp.tool(annotations=_WRITE)
async def sap_set_current_cell(grid_id: str, row: int, column: str, ctx: Context) -> dict:
    """Set the current (focused) cell in an ALV grid or table control.

    Useful before pressing toolbar buttons that act on the current cell."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.set_current_cell(grid_id, row, column))


@mcp.tool(annotations=_READ_ONLY)
async def sap_get_column_info(grid_id: str, ctx: Context) -> dict:
    """Get detailed column info from an ALV grid or table control.

    Returns column names, titles, widths, and visibility. Useful for
    understanding table structure. For a lighter alternative, use
    sap_read_table with columns_only=true."""
    c = _ctrl(ctx)
    return await _com(lambda: c.get_column_info(grid_id))


@mcp.tool(annotations=_READ_ONLY)
async def sap_get_current_cell(table_id: str, ctx: Context) -> dict:
    """Get the currently focused cell position in an ALV grid or table control."""
    c = _ctrl(ctx)
    return await _com(lambda: c.get_current_cell(table_id))


# ---- TableControl-specific tools ----

@mcp.tool(annotations=_WRITE)
async def sap_scroll_table_control(table_id: str, position: int, ctx: Context) -> dict:
    """Scroll a GuiTableControl to a specific row position.

    Does NOT work on ALV grids (they handle scrolling internally).
    For reading data at a specific offset, prefer sap_read_table with
    start_row parameter — it handles scrolling automatically."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.scroll_table_control(table_id, position))


@mcp.tool(annotations=_READ_ONLY)
async def sap_get_table_control_row_info(
    table_id: str,
    ctx: Context,
    rows: list[int] | None = None,
) -> dict:
    """Get row metadata (selectable, selected) from a GuiTableControl.

    If rows is omitted, queries all visible rows.
    Does NOT work on ALV grids."""
    c = _ctrl(ctx)
    return await _com(lambda: c.get_table_control_row_info(table_id, rows))


@mcp.tool(annotations=_WRITE)
async def sap_select_all_table_control_columns(
    table_id: str,
    ctx: Context,
    select: bool = True,
) -> dict:
    """Select or deselect all columns in a GuiTableControl. Does NOT work on ALV grids."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.select_all_table_control_columns(table_id, select))


# ---- ALV-specific tools ----

@mcp.tool(annotations=_READ_ONLY)
async def sap_get_cell_info(grid_id: str, row: int, column: str, ctx: Context) -> dict:
    """Get detailed cell metadata from an ALV grid.

    Returns value, changeable, color, tooltip, style, max_length.
    Does NOT work on GuiTableControl."""
    c = _ctrl(ctx)
    return await _com(lambda: c.get_cell_info(grid_id, row, column))


@mcp.tool(annotations=_WRITE)
async def sap_press_column_header(grid_id: str, column: str, ctx: Context) -> dict:
    """Click a column header in an ALV grid (triggers sort). Does NOT work on GuiTableControl."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.press_column_header(grid_id, column))


@mcp.tool(annotations=_WRITE)
async def sap_select_all_rows(grid_id: str, ctx: Context) -> dict:
    """Select all rows in an ALV grid. Does NOT work on GuiTableControl."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.select_all_rows(grid_id))


@mcp.tool(annotations=_WRITE)
async def sap_select_multiple_rows(table_id: str, rows: list[int], ctx: Context) -> dict:
    """Select multiple rows at once in an ALV grid or table control.

    Pass a list of row indices (e.g., [0, 2, 5])."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.select_multiple_rows(table_id, rows))


# ---- Popup & dialog tools ----

@mcp.tool(annotations=_READ_ONLY)
async def sap_get_popup_window(ctx: Context) -> dict:
    """Check if a popup/modal dialog is open (wnd[1], wnd[2], etc.).

    Returns the popup's title, text content, and available buttons so
    you know how to respond. Returns {popup_exists: false} if no popup."""
    c = _ctrl(ctx)
    return await _com(c.get_popup_window)


@mcp.tool(annotations=_WRITE)
async def sap_handle_popup(
    ctx: Context,
    action: Literal["read", "confirm", "cancel", "press"] = "read",
    button_text: str = "",
) -> dict:
    """Read and optionally act on the current popup/modal dialog.

    Use when active_window shows wnd[1] or higher. Combines popup
    inspection and response in a single call.

    Actions:
    - read: return popup content without acting (default)
    - confirm: press OK/Yes/Continue/Enter on the popup
    - cancel: press Cancel/No/F12 on the popup
    - press: press a specific button by its text or tooltip"""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.handle_popup(action, button_text))


# ---- Toolbar discovery ----

@mcp.tool(annotations=_READ_ONLY)
async def sap_get_toolbar_buttons(ctx: Context, window_id: str = "wnd[0]") -> dict:
    """List all buttons on the system toolbar (tbar[0]) and app toolbar (tbar[1]).

    Returns button IDs, text, tooltip, and enabled state. This is for
    standard SAP toolbars, NOT ALV (use sap_get_alv_toolbar for ALV)."""
    c = _ctrl(ctx)
    return await _com(lambda: c.get_toolbar_buttons(window_id))


# ---- Shell content ----

@mcp.tool(annotations=_READ_ONLY)
async def sap_read_shell_content(shell_id: str, ctx: Context) -> dict:
    """Read content from a GuiShell subtype (e.g., HTMLViewer).

    Extracts HTML, URL, or text depending on the shell type.
    Use sap_get_screen_elements first to find shell element IDs."""
    c = _ctrl(ctx)
    return await _com(lambda: c.read_shell_content(shell_id))


# ===========================================================================
# Tree tools
# ===========================================================================

@mcp.tool(annotations=_READ_ONLY)
async def sap_read_tree(tree_id: str, ctx: Context, max_nodes: int = 200) -> dict:
    """Read data from a tree control (TableTreeControl, ColumnTreeControl, etc.).

    Returns node hierarchy with texts and column values. For large trees
    (e.g., SPRO with 1000+ nodes), use sap_get_tree_node_children for
    step-by-step navigation instead. Use sap_search_tree_nodes to find
    specific nodes by text."""
    c = _ctrl(ctx)
    capped = min(max_nodes, config.max_table_rows)
    return await _com(lambda: c.read_tree(tree_id, capped))


@mcp.tool(annotations=_WRITE)
async def sap_expand_tree_node(tree_id: str, node_key: str, ctx: Context) -> dict:
    """Expand a folder node in a tree control to reveal its children.

    After expanding, use sap_get_tree_node_children or sap_read_tree
    to see the newly visible child nodes."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.expand_tree_node(tree_id, node_key))


@mcp.tool(annotations=_WRITE)
async def sap_collapse_tree_node(tree_id: str, node_key: str, ctx: Context) -> dict:
    """Collapse a folder node in a tree control"""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.collapse_tree_node(tree_id, node_key))


@mcp.tool(annotations=_WRITE)
async def sap_select_tree_node(tree_id: str, node_key: str, ctx: Context) -> dict:
    """Select a node in a tree control.

    Highlights the node without opening it. For SPRO-style trees,
    use sap_click_tree_link on the execute icon column instead."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.select_tree_node(tree_id, node_key))


@mcp.tool(annotations=_WRITE)
async def sap_double_click_tree_node(tree_id: str, node_key: str, ctx: Context) -> dict:
    """Double-click a node in a tree control (often opens details or drills down).

    In SPRO/customizing trees, this may open documentation (hypertext)
    rather than executing the activity. Use sap_click_tree_link on the
    execute column (typically column '2') for SPRO activities."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.double_click_tree_node(tree_id, node_key))


@mcp.tool(annotations=_WRITE)
async def sap_double_click_tree_item(
    tree_id: str, node_key: str, item_name: str, ctx: Context,
) -> dict:
    """Double-click a specific item (column cell) in a tree node row.

    item_name is the column name (e.g., 'Column1', 'Column2').
    Use sap_read_tree to discover column names for the tree."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(
        lambda: c.double_click_tree_item(tree_id, node_key, item_name)
    )


@mcp.tool(annotations=_WRITE)
async def sap_click_tree_link(tree_id: str, node_key: str, item_name: str, ctx: Context) -> dict:
    """Click a hyperlink in a tree node item.

    For SPRO/customizing trees, click on the execute icon column
    (typically item_name='2') to run an activity. Use sap_read_tree
    to see which columns have link-type items."""
    _check_write()
    c = _ctrl(ctx)
    return await _com(
        lambda: c.click_tree_link(tree_id, node_key, item_name)
    )


@mcp.tool(annotations=_READ_ONLY)
async def sap_find_tree_node_by_path(tree_id: str, path: str, ctx: Context) -> dict:
    """Find a tree node key by its path.

    E.g., '2\\1\\2' = 2nd child of root, then 1st child, then 2nd."""
    c = _ctrl(ctx)
    return await _com(lambda: c.find_tree_node_by_path(tree_id, path))


@mcp.tool(annotations=_READ_ONLY)
async def sap_search_tree_nodes(tree_id: str, search_text: str, ctx: Context,
                                column: str = "", max_results: int = 20) -> dict:
    """Search for tree nodes by text. Returns matches with full ancestor paths.

    Useful for finding nodes in deep trees where the same label appears
    in multiple branches. Case-insensitive substring match.
    Optionally pass column to search in a specific column instead of node text.

    **Important limitation**: Only searches nodes that are already loaded
    (expanded) in the tree. Collapsed subtrees are not searched. If you
    don't find what you expect, expand parent nodes first using
    sap_get_tree_node_children with expand=true, then search again."""
    c = _ctrl(ctx)
    capped = min(max_results, config.max_table_rows)
    return await _com(lambda: c.search_tree_nodes(
        tree_id, search_text, column, capped
    ))


@mcp.tool(annotations=_WRITE)
async def sap_get_tree_node_children(tree_id: str, ctx: Context, node_key: str = "",
                                     expand: bool = False) -> dict:
    """Get direct children of a tree node. Much faster than read_tree for
    step-by-step navigation of deep trees (e.g., SPRO).

    Omit node_key or pass empty string for root-level nodes.
    Set expand=true to expand the node first (requires write permission)."""
    if expand:
        _check_write()
    c = _ctrl(ctx)
    return await _com(lambda: c.get_tree_node_children(
        tree_id, node_key, expand
    ))


# ===========================================================================
# Discovery tools
# ===========================================================================

@mcp.tool(annotations=_READ_ONLY)
async def sap_get_screen_elements(
    ctx: Context,
    container_id: str = "wnd[0]/usr",
    max_depth: int = 2,
    type_filter: str = "",
    changeable_only: bool = False,
) -> dict:
    """Discover all elements on the current SAP screen.

    Useful for finding field IDs when working with a new screen.

    Use type_filter and changeable_only to reduce response size on
    complex screens (e.g. type_filter="GuiTextField,GuiCTextField"
    to find only input fields).

    max_depth controls how deep to recurse into containers (default 2).
    Use max_depth=1 for a quick overview, max_depth=3+ for deeply nested
    layouts (splitter containers, tab strips with sub-containers).

    Pass container_id='wnd[0]/mbar' to discover the menu bar structure."""
    c = _ctrl(ctx)
    elements = await _com(
        lambda: c.get_screen_elements(
            container_id, max_depth=max_depth,
            type_filter=type_filter,
            changeable_only=changeable_only,
        )
    )
    return {
        "element_count": len(elements),
        "elements": [e.__dict__ for e in elements],
    }


@mcp.tool(annotations=_READ_ONLY)
async def sap_screenshot(ctx: Context) -> Image:
    """Take a screenshot of the current SAP window.

    Use as a fallback when structured tools (sap_get_screen_elements,
    sap_read_field, sap_read_table) return empty or confusing results,
    e.g., on Web Dynpro screens where the element tree is non-standard."""
    c = _ctrl(ctx)
    result = await _com(c.take_screenshot)
    if "error" in result:
        raise ValueError(result["error"])
    return Image(data=base64.b64decode(result["data"]), format="png")


# ===========================================================================
# Resources
# ===========================================================================

@mcp.tool(annotations=_READ_ONLY)
async def sap_disconnect(ctx: Context) -> dict:
    """Disconnect from the current SAP session and release the binding.

    Sessions opened by this MCP session (via sap_connect) are closed.
    Sessions that were attached (via sap_connect_existing) are detached
    but left open so the user can continue working manually."""
    key = _session_key(ctx)
    return await _com(lambda: _session_mgr.release(key))


@mcp.resource("docs://sap-gui-guide")
def sap_gui_guide() -> str:
    """Detailed SAP GUI navigation reference for AI agents.

    Covers element types, ID patterns, transaction conventions,
    table type comparison, SPRO navigation, and troubleshooting."""
    return _SAP_GUI_GUIDE


# ===========================================================================
# Main entry point
# ===========================================================================

def main():
    """Main entry point."""
    import argparse
    global config

    parser = argparse.ArgumentParser(description="MCP Server for SAP GUI")
    parser.add_argument("--read-only", action="store_true",
                        help="Run in read-only mode (no write operations)")
    parser.add_argument("--allowed-transactions", nargs="*",
                        help="Whitelist of allowed transaction codes")
    parser.add_argument("--transport", choices=["stdio", "http"],
                        default="stdio",
                        help="Transport mode (default: stdio)")
    parser.add_argument("--host", default="127.0.0.1",
                        help="HTTP bind address (default: 127.0.0.1)")
    parser.add_argument("--port", type=int, default=8000,
                        help="HTTP port (default: 8000)")
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
    # SessionManager is created in _lifespan, controllers are per-session.

    if args.transport == "http":
        logger.info("Starting HTTP transport on %s:%d", args.host, args.port)
        mcp.run(transport="http", host=args.host, port=args.port)
    else:
        mcp.run()


if __name__ == "__main__":
    main()
