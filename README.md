# MCP SAP GUI Server

An [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server that enables AI assistants like Claude to interact with SAP GUI for Windows through the SAP GUI Scripting API.

## What This Does

This server allows Claude to:
- Connect to SAP systems (like double-clicking in SAP Logon Pad)
- Execute transactions (MM03, VA01, SE80, etc.)
- Read and write screen fields, checkboxes, and buttons
- Navigate through SAP screens
- Extract data from ALV grids and tables
- Read and interact with tree controls (TableTree, ColumnTree)
- Use ALV toolbar buttons and context menus (e.g., monitor methods)
- Take screenshots of SAP windows
- Discover screen elements for automation

## Example Conversation

```
User: "What's the description for material MAT-001 in our DEV system?"

Claude: [connects to DEV system]
        [executes MM03]
        [enters material number]
        [reads description field]

"The material MAT-001 is described as 'High-Grade Steel Plate 10mm'
in the DEV system."
```

## Requirements

- **Windows** (SAP GUI only runs on Windows)
- **SAP GUI for Windows** installed
- **SAP Logon Pad** running (for COM connections)
- **SAP GUI Scripting enabled** on your SAP systems
- **Python 3.10+**
- **[uv](https://docs.astral.sh/uv/)** (recommended Python package manager)

### Enabling SAP GUI Scripting

SAP GUI Scripting must be enabled both client-side and server-side:

**Client-side** (SAP GUI Options):
1. Open SAP GUI → Options → Accessibility & Scripting → Scripting
2. Check "Enable scripting"
3. Optionally uncheck "Notify when a script..." for smoother automation

**Server-side** (SAP System):
- Transaction `RZ11` → Parameter `sapgui/user_scripting` → Set to `TRUE`
- Requires SAP Basis administrator access

## Installation

```bash
# Clone the repository
git clone ssh://git@git.tsioumpris.de:2222/Kostas/mcp-sap-gui.git
cd mcp-sap-gui

# Install uv (if not already installed)
pip install uv

# Install all dependencies (creates .venv automatically)
uv sync

# With screenshot optimization (recommended - reduces screenshot size by 70-90%)
uv sync --extra screenshots

# With dev dependencies (for testing, linting, type checking)
uv sync --extra dev --extra screenshots
```

## Usage

### Running the MCP Server Directly

```bash
# Standard mode
uv run python -m mcp_sap_gui.server

# Read-only mode (safer for exploration)
uv run python -m mcp_sap_gui.server --read-only

# With transaction whitelist
uv run python -m mcp_sap_gui.server --allowed-transactions MM03 VA03 ME23N

# Debug mode
uv run python -m mcp_sap_gui.server --debug
```

### MCP Setup

This server communicates over **stdio** (stdin/stdout JSON-RPC), which is the standard MCP transport. You don't need to configure ports or URLs — the MCP client starts the server process and talks to it directly.

Below are setup examples for the most common MCP clients.

---

#### Option 1: Claude Code (Recommended for development)

The repository includes a `.mcp.json` at the project root. When you open this project in **Claude Code**, the MCP server is automatically discovered — no manual configuration needed.

To use it:

```bash
cd mcp-sap-gui
claude
```

Claude Code will detect `.mcp.json` and start the SAP GUI MCP server automatically.

If you want to configure it globally for Claude Code (available in any project), add it to your user settings at `~/.claude/.mcp.json`:

```json
{
  "mcpServers": {
    "sap-gui": {
      "type": "stdio",
      "command": "uv",
      "args": ["run", "python", "-m", "mcp_sap_gui.server"],
      "cwd": "D:\\mcp-sap-gui"
    }
  }
}
```

---

#### Option 2: Claude Desktop

Add to your Claude Desktop config file:

- **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`

**Standard mode:**

```json
{
  "mcpServers": {
    "sap-gui": {
      "command": "uv",
      "args": ["run", "python", "-m", "mcp_sap_gui.server"],
      "cwd": "D:\\mcp-sap-gui"
    }
  }
}
```

**Read-only mode** (recommended when exploring/querying data):

```json
{
  "mcpServers": {
    "sap-gui": {
      "command": "uv",
      "args": ["run", "python", "-m", "mcp_sap_gui.server", "--read-only"],
      "cwd": "D:\\mcp-sap-gui"
    }
  }
}
```

**With transaction whitelist** (only allow specific transactions):

```json
{
  "mcpServers": {
    "sap-gui": {
      "command": "uv",
      "args": [
        "run", "python", "-m", "mcp_sap_gui.server",
        "--allowed-transactions", "MM03", "VA03", "ME23N"
      ],
      "cwd": "D:\\mcp-sap-gui"
    }
  }
}
```

After editing the config, restart Claude Desktop for changes to take effect.

---

#### Option 3: Any MCP-compatible client

The server uses stdio transport. Point any MCP client at:

```
Command:   uv run python -m mcp_sap_gui.server
Arguments: [--read-only] [--debug] [--allowed-transactions T1 T2 ...]
Transport: stdio
```

Ensure the working directory is set to the project root (where `pyproject.toml` lives).

---

### Verifying the Setup

Once configured, you can verify the MCP server is working by asking Claude:

```
"List all available SAP GUI tools"
```

Claude should respond with the full list of `sap_*` tools. If SAP GUI is running, try:

```
"Connect to my open SAP session and tell me what system I'm on"
```

## Available Tools

### Connection Tools
| Tool | Description |
|------|-------------|
| `sap_connect` | Connect to SAP system by name (with optional credentials) |
| `sap_connect_existing` | Attach to an already open SAP session |
| `sap_list_connections` | List all open SAP connections/sessions |
| `sap_get_session_info` | Get current session info (system, user, transaction) |

### Navigation Tools
| Tool | Description |
|------|-------------|
| `sap_execute_transaction` | Execute a transaction code (MM03, VA01, etc.) |
| `sap_send_key` | Send keyboard keys (Enter, F1-F12, Back, Save, etc.) |
| `sap_get_screen_info` | Get current screen information |

### Field Tools
| Tool | Description |
|------|-------------|
| `sap_read_field` | Read a field value |
| `sap_set_field` | Set a field value |
| `sap_press_button` | Press a button |
| `sap_select_checkbox` | Select or deselect a checkbox |

### Table/Grid Tools
| Tool | Description |
|------|-------------|
| `sap_read_table` | Read data from ALV grid/table |
| `sap_select_table_row` | Select a table row |
| `sap_double_click_cell` | Double-click a table cell |
| `sap_get_alv_toolbar` | List all toolbar buttons on an ALV grid |
| `sap_press_alv_toolbar_button` | Press an ALV toolbar button (auto-detects menu types) |
| `sap_select_alv_context_menu_item` | Select from ALV context menu (supports atomic open+select) |

### Tree Tools
| Tool | Description |
|------|-------------|
| `sap_read_tree` | Read tree control nodes, hierarchy, and column values |
| `sap_expand_tree_node` | Expand a tree folder node |
| `sap_collapse_tree_node` | Collapse a tree folder node |
| `sap_select_tree_node` | Select a tree node |
| `sap_double_click_tree_node` | Double-click a tree node (drill down) |

### Discovery Tools
| Tool | Description |
|------|-------------|
| `sap_get_screen_elements` | List all elements on current screen |
| `sap_screenshot` | Capture screenshot of SAP window (auto-detects popups) |

## Security Considerations

This server provides powerful automation capabilities. **Use responsibly.**

### Built-in Safeguards

1. **Transaction Blocklist** - Sensitive transactions blocked by default:
   - `SU01`, `SU10`, `SU01D` (User administration)
   - `PFCG` (Role administration)
   - `SE16N` (Direct table access)

2. **Read-Only Mode** - `--read-only` flag disables all mutating operations (field writes, button presses, transaction execution, key sends, tree/table interactions)

3. **Transaction Whitelist** - `--allowed-transactions` limits to specific t-codes

### Recommendations for Production Use

- **Never expose to untrusted users**
- **Use read-only mode** for exploration/queries
- **Implement transaction whitelists** for automation
- **Enable audit logging** on SAP side
- **Use dedicated service accounts** with minimal authorizations
- **Run on isolated systems** (test/sandbox, not production)

### SAP Licensing

Consult your SAP licensing agreement regarding:
- Automated access and scripting
- Indirect access considerations
- Named vs. concurrent user licensing

## Example Workflows

### Display Material Master

```python
# Claude would execute these tools:
sap_connect("DEV - Development System")
sap_execute_transaction("MM03")
sap_set_field("wnd[0]/usr/ctxtRMMG1-MATNR", "MAT-001")
sap_send_key("Enter")
# Select views...
sap_send_key("Enter")
description = sap_read_field("wnd[0]/usr/txtMAKT-MAKTX")
```

### Extract Purchase Order List

```python
sap_execute_transaction("ME2M")
sap_set_field("wnd[0]/usr/ctxtEN_LIFNR-LOW", "1000")  # Vendor
sap_send_key("Execute")  # F8
data = sap_read_table("wnd[0]/usr/cntlGRID1/shellcont/shell", max_rows=50)
```

### Navigate Sales Order

```python
sap_execute_transaction("VA03")
sap_set_field("wnd[0]/usr/ctxtVBAK-VBELN", "12345")
sap_send_key("Enter")
# Read header data
customer = sap_read_field("wnd[0]/usr/subSUBSCREEN.../txtVBAK-KUNNR")
# Navigate to items
sap_press_button("wnd[0]/usr/tabsTAXI_TABSTRIP.../tabpT\\01")
items = sap_read_table("wnd[0]/usr/.../cntlGRID1/shellcont/shell")
```

## Project Structure

```
mcp-sap-gui/
├── src/
│   └── mcp_sap_gui/
│       ├── __init__.py
│       ├── server.py         # MCP server implementation
│       └── sap_controller.py # SAP GUI COM wrapper
├── tests/
│   ├── test_sap_controller.py  # Controller unit tests
│   └── test_server.py          # Server security & routing tests
├── examples/
│   └── ...
├── .mcp.json                  # MCP server config (auto-detected by Claude Code)
├── pyproject.toml
├── uv.lock                    # Dependency lock file (managed by uv)
├── README.md
└── CLAUDE.md                  # Development context for Claude Code
```

## Troubleshooting

### "Cannot connect to SAP GUI"
- Ensure SAP Logon Pad is running
- Check that SAP GUI Scripting is enabled in SAP GUI options

### "Scripting disabled" error
- Enable scripting server-side: `RZ11` → `sapgui/user_scripting` = `TRUE`
- Requires SAP Basis administrator

### "Element not found"
- Use `sap_get_screen_elements()` to discover available field IDs
- Field IDs vary between SAP systems due to customization

### COM errors on startup
- Ensure dependencies are installed: `uv sync`
- Run `uv run python -m win32com.client.makepy` if COM registration issues occur

## Development

```bash
# Install all dependencies (dev + screenshots)
uv sync --extra dev --extra screenshots

# Run tests
uv run pytest

# Type checking
uv run mypy src/

# Linting
uv run ruff check src/
```

## Related Projects

- [ZSAPConnect Manager](../sapgui-manager) - SAP connection & credential manager
- [SAP GUI Scripting API Documentation](https://help.sap.com/docs/sap_gui_for_windows)

## License

MIT License - see LICENSE file for details.

## Disclaimer

This project is not affiliated with, endorsed by, or sponsored by SAP SE. SAP, SAP GUI, and other SAP products mentioned are trademarks of SAP SE.

Use of this software with SAP systems should comply with your SAP licensing agreement and your organization's security policies.
