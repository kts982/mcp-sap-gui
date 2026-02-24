# CLAUDE.md

Development context for Claude Code when working on the MCP SAP GUI Server.

## Project Overview

An MCP (Model Context Protocol) server that exposes SAP GUI Scripting capabilities to AI assistants. Uses Windows COM automation via `pywin32` to control SAP GUI for Windows.

## Architecture

```
┌─────────────────┐     MCP Protocol      ┌──────────────────┐
│  Claude/AI      │◄────────────────────►│  MCP Server      │
│  Assistant      │    (stdio/JSON-RPC)   │  (server.py)     │
└─────────────────┘                       └────────┬─────────┘
                                                   │
                                          Python calls
                                                   │
                                          ┌────────▼─────────┐
                                          │  SAP Controller  │
                                          │  (sap_controller)│
                                          └────────┬─────────┘
                                                   │
                                          COM/IDispatch
                                                   │
                                          ┌────────▼─────────┐
                                          │  SAP GUI         │
                                          │  (sapfewse.ocx)  │
                                          └────────┬─────────┘
                                                   │
                                          ┌────────▼─────────┐
                                          │  SAP System      │
                                          └──────────────────┘
```

## Key Files

| File | Purpose |
|------|---------|
| `src/mcp_sap_gui/server.py` | MCP server implementation, tool definitions, request routing |
| `src/mcp_sap_gui/sap_controller.py` | SAP GUI COM wrapper, all SAP interaction logic |
| `src/mcp_sap_gui/__init__.py` | Package initialization |

## SAP GUI COM Object Model

The SAP GUI Scripting API exposes this hierarchy:

```
SAPGUI (ROT object)
└── ScriptingEngine (GuiApplication)
    └── Children (Connections)
        └── Connection (GuiConnection)
            └── Children (Sessions)
                └── Session (GuiSession)
                    ├── Info (session metadata)
                    ├── findById() → Screen elements
                    └── Children → Windows
                        └── Window (GuiFrameWindow)
                            └── Children → UI elements
```

## Common SAP GUI Element Types

| Type | Description | Example ID |
|------|-------------|------------|
| `GuiTextField` | Text input field | `wnd[0]/usr/txtMATNR` |
| `GuiCTextField` | Text field with search help | `wnd[0]/usr/ctxtRMMG1-MATNR` |
| `GuiPasswordField` | Password field | `wnd[0]/usr/pwdRSYST-BCODE` |
| `GuiButton` | Pushbutton | `wnd[0]/tbar[1]/btn[8]` |
| `GuiCheckBox` | Checkbox | `wnd[0]/usr/chkFLAG` |
| `GuiRadioButton` | Radio button | `wnd[0]/usr/radOPT1` |
| `GuiComboBox` | Dropdown list | `wnd[0]/usr/cmbLANGU` |
| `GuiTab` | Tab strip tab | `wnd[0]/usr/tabsTAB/tabpTAB1` |
| `GuiMenu` | Menu bar item | `wnd[0]/mbar/menu[0]/menu[1]` |
| `GuiGridView` | ALV Grid | `wnd[0]/usr/cntlGRID1/shellcont/shell` |
| `GuiTableControl` | Classic table (customizing) | `wnd[0]/usr/tblSAPLBD41TCTRL_V_TBDLS` |
| `GuiTree` | Tree control | `wnd[0]/usr/cntlTREE/shellcont/shell` |
| `GuiStatusbar` | Status bar | `wnd[0]/sbar` |
| `GuiOkCodeField` | Command field | `wnd[0]/tbar[0]/okcd` |

## Virtual Keys (VKey)

Common virtual key codes for `sendVKey()`:

| Key | VKey | Use |
|-----|------|-----|
| Enter | 0 | Confirm/Continue |
| F1 | 1 | Help |
| F3 | 3 | Back |
| F4 | 4 | Dropdown/Search help |
| F5 | 5 | Refresh |
| F8 | 8 | Execute |
| F11 | 11 | Save (Ctrl+S) |
| F12 | 12 | Cancel (ESC) |

## Development Commands

```bash
# Install dependencies
uv sync --extra dev --extra screenshots

# Run server
uv run python -m mcp_sap_gui.server

# Run with debug logging
uv run python -m mcp_sap_gui.server --debug

# Run tests
uv run pytest tests/

# Type checking
uv run mypy src/mcp_sap_gui/

# Linting
uv run ruff check src/
```

## Testing Strategy

**Unit tests** - Mock the COM interface
```python
from unittest.mock import MagicMock
controller = SAPGUIController()
controller._win32com = MagicMock()
```

**Integration tests** - Require actual SAP system
- Use a dedicated test SAP system
- Test with display-only transactions (MM03, VA03)
- Never test on production

## Security Implementation

### Transaction Blocking

Default blocklist in `ServerConfig`:
```python
blocked_transactions = [
    "SU01", "SU10", "SU01D",  # User admin
    "PFCG", "SU53",           # Roles
    "SM21", "ST22",           # Logs
    "SE16N",                  # Table access
]
```

### Read-Only Mode

When `config.read_only = True`, all write tools raise `ValueError`. Each write tool calls `_check_write()` at the start. The tools are always visible but return an error if invoked in read-only mode.

**Write tools** (blocked in read-only): `sap_execute_transaction`, `sap_send_key`, `sap_set_field`, `sap_press_button`, `sap_select_menu`, `sap_select_checkbox`, `sap_select_radio_button`, `sap_select_combobox_entry`, `sap_select_tab`, `sap_set_batch_fields`, `sap_set_textedit`, `sap_set_focus`, `sap_select_table_row`, `sap_double_click_cell`, `sap_modify_cell`, `sap_set_current_cell`, `sap_press_alv_toolbar_button`, `sap_select_alv_context_menu_item`, `sap_scroll_table_control`, `sap_select_all_table_control_columns`, `sap_press_column_header`, `sap_select_all_rows`, `sap_expand_tree_node`, `sap_collapse_tree_node`, `sap_select_tree_node`, `sap_double_click_tree_node`, `sap_double_click_tree_item`, `sap_click_tree_link`

**Read tools** (always allowed): `sap_connect`, `sap_connect_existing`, `sap_list_connections`, `sap_get_session_info`, `sap_get_screen_info`, `sap_read_field`, `sap_get_combobox_entries`, `sap_read_textedit`, `sap_read_table`, `sap_get_alv_toolbar`, `sap_get_column_info`, `sap_get_current_cell`, `sap_get_table_control_row_info`, `sap_get_cell_info`, `sap_read_tree`, `sap_find_tree_node_by_path`, `sap_get_screen_elements`, `sap_screenshot`

### Adding New Blocked Transactions

Edit `ServerConfig.blocked_transactions` in `server.py`.

## Adding New Tools

The server uses **FastMCP** (`@mcp.tool()` decorators). Adding a new tool is two steps:

1. Add the controller method in `sap_controller.py`:
```python
def new_method(self, param1: str) -> Dict[str, Any]:
    self._require_session()
    # Implementation using self._session.findById(...)
    return {"result": "..."}
```

2. Add the MCP tool in `server.py` (schema auto-generated from type hints):
```python
@mcp.tool()
async def sap_new_tool(param1: str) -> dict:
    """Description shown to Claude."""
    _check_write()  # Add this for write operations
    return await _com(lambda: controller.new_method(param1))
```

3. Update tests:
   - Add read-only test in `test_server.py` if it's a write operation
   - Add the tool name to the `test_all_tools_registered` set
   - Add controller unit tests in `test_sap_controller.py`

## Common Patterns

### Finding Elements on Unknown Screens

```python
# Get all elements
elements = controller.get_screen_elements()
for e in elements:
    if e.changeable:  # Input fields
        print(f"{e.id}: {e.type} = {e.text}")
```

### Handling Popups

```python
# Check for popup window
try:
    popup = session.findById("wnd[1]")
    # Handle popup
    popup.findById("wnd[1]/usr/btnBUTTON_1").press()  # Yes
except:
    pass  # No popup
```

### Reading Status Messages

```python
def get_status():
    sbar = session.findById("wnd[0]/sbar")
    return {
        "message": sbar.Text,
        "type": sbar.MessageType,  # S=Success, E=Error, W=Warning, I=Info
    }
```

## GuiGridView (ALV) vs GuiTableControl

Both are supported by `sap_read_table` and other table tools — auto-detected via `Type` property.

| | GuiGridView (ALV) | GuiTableControl |
|---|---|---|
| **Type prefix** | `shell` (inside `shellcont`) | `tbl` |
| **Example ID** | `wnd[0]/usr/cntlGRID/shellcont/shell` | `wnd[0]/usr/tblSAPLBD41TCTRL_V_TBDLS` |
| **Common in** | Reports, list displays | SPRO/customizing, SM30 table maintenance |
| **Cell access** | `GetCellValue(row, colName)` | `GetCell(visRow, colIdx)` → `.Text` |
| **Row selection** | `selectedRows = "row"` | `GetAbsoluteRow(absRow).Selected = True` |
| **Column names** | `ColumnOrder(i)` collection | From `GetCell(0, i).Name` (NOT from Columns collection) |
| **Scrolling** | Internal (all rows accessible) | Manual via `VerticalScrollbar.Position` |
| **Reading** | All rows accessible directly | Visible rows only (no scroll during read) |
| **Toolbar** | Built-in ALV toolbar | Standard `tbar[1]` buttons + menu bar |

**Important GuiTableControl indexing** (from SAP GUI Scripting API docs):
- `GetCell(row, col)` uses **visible-row** indexing (0 = first visible row). Row must be visible or an exception is raised.
- `GetAbsoluteRow(row)` uses **absolute** indexing (independent of scroll position). Has a `Selected` property for row selection.
- `Rows(idx)` uses visible-row indexing (resets after scroll).
- Rapid `VerticalScrollbar.Position` changes can crash the COM server. Read operations avoid scrolling; interaction operations perform a single scroll to make the target row visible.

**Important**: GuiTableColumn objects "do not support properties like id or name" per the API docs. Accessing `col.Name` can crash SAP GUI. Column names must be read from cell objects instead.

## Known Limitations

1. **Windows only** - SAP GUI is Windows-only software
2. **Single-threaded COM** - Must use apartment threading
3. **Field IDs vary** - Custom SAP systems have different field IDs
4. **Session focus** - Some operations require window focus
5. **No background execution** - GUI must be visible

## Debugging Tips

1. **Enable SAP GUI Scripting Tracker**:
   - SAP GUI → Options → Scripting → Check "Notify when..."
   - Shows all scripting calls in real-time

2. **Record scripts in SAP**:
   - Transaction `/o/SAPGUI/RECORD`
   - Records VBScript that shows exact element IDs

3. **Print element tree**:
```python
def print_tree(elem, depth=0):
    print("  " * depth + f"{elem.Id} ({elem.Type})")
    for i in range(elem.Children.Count):
        print_tree(elem.Children(i), depth + 1)
```

## References

- [SAP GUI Scripting API Reference](https://help.sap.com/docs/sap_gui_for_windows)
- [SAP Note 480149](https://launchpad.support.sap.com/#/notes/480149) - Scripting Security
- [SAP Note 587202](https://launchpad.support.sap.com/#/notes/587202) - Scripting Setup
- [MCP Specification](https://modelcontextprotocol.io/docs)
