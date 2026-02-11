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
| `GuiGridView` | ALV Grid | `wnd[0]/usr/cntlGRID1/shellcont/shell` |
| `GuiTableControl` | Classic table | `wnd[0]/usr/tblSAPMV45A` |
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
# Run server
python -m mcp_sap_gui.server

# Run with debug logging
python -m mcp_sap_gui.server --debug

# Run tests
pytest tests/

# Type checking
mypy src/mcp_sap_gui/

# Linting
ruff check src/
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

When `config.read_only = True`:
- `sap_set_field` → returns error
- `sap_press_button` → returns error
- `sap_select_table_row` → returns error
- Only read/navigation operations allowed

### Adding New Blocked Transactions

Edit `ServerConfig.blocked_transactions` in `server.py`.

## Adding New Tools

1. Add tool definition in `list_tools()`:
```python
Tool(
    name="sap_new_tool",
    description="Description for Claude",
    inputSchema={
        "type": "object",
        "properties": {
            "param1": {"type": "string", "description": "..."}
        },
        "required": ["param1"]
    }
)
```

2. Add handler in `_handle_tool()`:
```python
elif name == "sap_new_tool":
    return await loop.run_in_executor(
        None, lambda: self.controller.new_method(arguments["param1"])
    )
```

3. Implement method in `sap_controller.py`:
```python
def new_method(self, param1: str) -> Dict[str, Any]:
    self._require_session()
    # Implementation
    return {"result": "..."}
```

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
