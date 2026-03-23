# MCP SAP GUI Server

An [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server that enables AI assistants to interact with SAP GUI for Windows through the SAP GUI Scripting API.

It is client-agnostic: if your MCP client can launch a local `stdio` server, it can use this project. Examples in this README use Claude because the setup is easy to demonstrate, but the same server can be used from Codex, GitHub Copilot, Gemini CLI, and similar MCP-capable tools.

Current release target: `0.1.0` alpha for local Windows use over MCP `stdio`.

## Status

- GitHub workflows are included for `CI`, `Docs`, and tag-based `Release`.
- Forgejo workflows are included for `CI` and `Docs`.
- Live GitHub badges can be enabled once the public GitHub mirror slug is known.

<!--
Replace <owner>/<repo> after publishing the GitHub mirror:

[![CI](https://github.com/<owner>/<repo>/actions/workflows/ci.yml/badge.svg)](https://github.com/<owner>/<repo>/actions/workflows/ci.yml)
[![Docs](https://github.com/<owner>/<repo>/actions/workflows/docs.yml/badge.svg)](https://github.com/<owner>/<repo>/actions/workflows/docs.yml)
[![Release](https://img.shields.io/github/v/release/<owner>/<repo>)](https://github.com/<owner>/<repo>/releases)
-->

## What This Does

This server allows AI assistants to:
- Connect to SAP systems (like double-clicking in SAP Logon Pad)
- Execute transactions (MM03, VA01, SE80, etc.)
- Read and write screen fields, checkboxes, radio buttons, comboboxes, and tabs
- Select menu items from the menu bar (Table View, Edit, Selection, etc.)
- Navigate through SAP screens using keyboard keys and buttons
- Extract data from ALV grids (GuiGridView) and classic table controls (GuiTableControl)
- Interact with ALV toolbar buttons and context menus
- Read and interact with tree controls (TableTree, ColumnTree, SimpleTree)
- Take screenshots of SAP windows
- Discover screen elements for automation

## Example Conversation

```
User: "What's the description for material MAT-001 in system D01?"

Assistant: [connects to D01]
           [executes MM03]
           [enters material number]
           [reads description field]

"The material MAT-001 is described as 'High-Grade Steel Plate 10mm'
in system D01."
```

## Quick Start

1. Install dependencies:

```bash
uv sync --extra screenshots
```

2. Start SAP Logon Pad and open an SAP GUI session, or at least have SAP Logon running.

3. Configure your MCP client to launch this server:

```text
Command:   uv
Arguments: run python -m mcp_sap_gui.server
Transport: stdio
Working directory: <path-to-mcp-sap-gui>
```

4. Try one of these prompts:

```text
Connect to my open SAP session and tell me what system I'm on
Show me the current screen info
List all editable fields on this screen
Read the first 20 rows of the visible table
```

## Requirements

- **Windows** (SAP GUI only runs on Windows)
- **SAP GUI for Windows** installed
- **SAP Logon Pad** running (for COM connections)
- **SAP GUI Scripting enabled** on your SAP systems
- **Python 3.10+**
- **[uv](https://docs.astral.sh/uv/)** (recommended Python package manager)

## Supported Scope

Supported in `0.1.0`:
- SAP GUI for Windows via the SAP GUI Scripting COM API
- Local MCP `stdio` transport
- Interactive use from MCP-compatible clients
- Read and write SAP GUI automation within the permissions of the logged-in SAP user

Not part of `0.1.0`:
- Streamable HTTP / remote server deployment
- SAP GUI for Java or SAP GUI for HTML
- Browser-based Fiori automation
- Unattended multi-user production orchestration

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
git clone <repository-url>
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

For any client, the core launch configuration is the same:

```text
Command:   uv
Arguments: run python -m mcp_sap_gui.server
Transport: stdio
Working directory: <path-to-mcp-sap-gui>
```

### Client Setup Links

- **Claude Code / Claude Desktop**: setup examples are included below
- **Codex**: configure an MCP server in Codex and point it at the command above. Official MCP docs: https://developers.openai.com/learn/docs-mcp
- **GitHub Copilot**: configure a local MCP server in Copilot Chat / agent mode. Official docs: https://docs.github.com/en/copilot/how-tos/provide-context/use-mcp
- **Gemini CLI**: add the server under `mcpServers` in your Gemini CLI settings. Official docs: https://github.com/google-gemini/gemini-cli/blob/main/docs/tools/mcp-server.md

Below are full examples for the most common local SAP GUI setup paths.

For a client-by-client setup guide, see **[docs/CLIENTS.md](docs/CLIENTS.md)**.

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
      "cwd": "<path-to-mcp-sap-gui>"
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
      "cwd": "<path-to-mcp-sap-gui>"
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
      "cwd": "<path-to-mcp-sap-gui>"
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
      "cwd": "<path-to-mcp-sap-gui>"
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

Then try:

```
"Show me the current screen info"
"List all editable fields on this screen"
"Read the first 20 rows from the visible table"
```

## Built-in Agent Guidance

The server includes built-in navigation knowledge that helps any MCP client (Claude Code, Copilot, Cursor, Cline, etc.) use SAP GUI effectively:

- **MCP Instructions** — Injected into every client's system prompt during initialization. Covers screen discovery workflow, popup handling, table pagination, SPRO tree navigation, key reference, and common mistakes to avoid.
- **`docs://sap-gui-guide` Resource** — Detailed reference guide available on-demand via `resources/read`. Covers element types, ID naming conventions, transaction code formats, table type comparison, status bar messages, and step-by-step patterns for SPRO and table maintenance views.

These prevent common agent mistakes like guessing element IDs, ignoring popups, pressing F5 (="New Entries") when meaning to refresh, or using `double_click_tree_node` in SPRO (which opens docs instead of executing the activity).

## Available Tools

The server currently exposes **52 MCP tools**.

| Category | Count | What it covers |
|---|---:|---|
| Connection | 4 | Connect to SAP, attach to open sessions, inspect sessions |
| Navigation | 3 | Execute transactions, send keys, inspect current screen |
| Fields & UI | 13 | Read/write fields, buttons, tabs, comboboxes, textedit, focus |
| Tables & Grids | 17 | ALV grids, TableControls, row selection, column info, cell ops |
| Popup / Toolbar / Shell | 3 | Popup inspection, toolbar discovery, shell content |
| Trees | 10 | Read/search/expand/select/click SAP tree controls |
| Discovery | 2 | Screen element discovery and screenshots |

The most important patterns:
- `sap_get_screen_elements` to discover IDs instead of guessing
- `sap_read_table` to start with any SAP table/grid
- `sap_get_popup_window` when `active_window` reports a popup
- `sap_read_tree` plus search/expand helpers for SPRO-style navigation

For the full tool catalog, grouped by category with short descriptions, see **[docs/TOOLS.md](docs/TOOLS.md)**.

## Security Considerations

This server provides powerful automation capabilities. **Use responsibly.**

### Built-in Safeguards

1. **Transaction Blocklist** - Sensitive transactions blocked by default:
   - `SU01`, `SU10`, `SU01D` (User administration)
   - `PFCG` (Role administration)
   - `SE16N` (Direct table access)
   - Case-insensitive matching; handles `/n`, `/o`, `/*` prefixes and whitespace

2. **OK-Code Bypass Prevention** - Setting the OK-code field (`tbar[0]/okcd`) to a blocked transaction is also blocked, preventing circumvention of the transaction blocklist

3. **Read-Only Mode** - `--read-only` flag disables all mutating operations (field writes, button presses, transaction execution, key sends, tree/table interactions)

4. **Transaction Whitelist** - `--allowed-transactions` limits to specific t-codes

5. **MCP Tool Annotations** - All 52 tools are annotated with `readOnlyHint`/`destructiveHint` per the MCP spec, so clients can display appropriate UI hints

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
sap_connect("D01 - Development System")
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

### Filter Customizing Table (GuiTableControl)

```python
# In SPRO or SM30 table maintenance view
# Use Selection -> By Contents to filter
sap_select_menu("wnd[0]/mbar/menu[3]/menu[0]")    # Selection > By Contents
# Select the field to filter on, enter value
sap_select_table_row("wnd[1]/usr/tblSAPLSVIXTCTRL_SEL_FLDS", 0)
sap_send_key("Enter")
sap_set_field("wnd[1]/usr/.../txtQUERY_TAB-BUFFER[3,0]", "EXTSYS001")
sap_send_key("Execute")
# Read the filtered table
data = sap_read_table("wnd[0]/usr/tblSAPLBD41TCTRL_V_TBDLS")
```

## Project Structure

```
mcp-sap-gui/
├── docs/
│   ├── CLIENTS.md             # Client-specific MCP setup notes
│   ├── OVERVIEW.md            # Product overview and roadmap direction
│   └── TOOLS.md               # Full MCP tool catalog
├── scripts/
│   └── check_docs.py          # Markdown link checker used by docs workflows
├── src/
│   └── mcp_sap_gui/
│       ├── __init__.py        # Package exports
│       ├── server.py          # MCP server implementation (tool definitions)
│       ├── sap_controller.py  # Facade class (composes all mixins)
│       ├── models.py          # VKey enum, SessionInfo, exceptions
│       ├── controller.py      # Base controller (connection, navigation, screen info)
│       ├── fields.py          # FieldsMixin (read/write fields, buttons, combos)
│       ├── tables.py          # TablesMixin (ALV grid + TableControl operations)
│       ├── trees.py           # TreesMixin (tree read, expand, select, click)
│       └── discovery.py       # DiscoveryMixin (popups, toolbars, screenshots)
├── tests/
│   ├── test_sap_controller.py # Controller unit tests
│   └── test_server.py         # Server security & routing tests
├── examples/
│   └── basic_usage.py         # Direct controller example
├── .mcp.json                  # MCP server config (auto-detected by Claude Code)
├── CONTRIBUTING.md            # Contribution guidelines for public changes
├── LICENSE                    # MIT license
├── pyproject.toml
├── uv.lock                    # Dependency lock file (managed by uv)
└── README.md
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

### "No SAP tools appear in my MCP client"
- Confirm the client is launching the server from the project root
- Restart the MCP client after changing its MCP configuration
- Run `uv sync` first so the environment and dependencies exist

### "The tool is available, but the action is blocked"
- Check whether the server is running with `--read-only`
- Check whether the transaction is blocked by the default blocklist
- Check whether you started the server with `--allowed-transactions`

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

## Related

- [SAP GUI Scripting API Documentation](https://help.sap.com/docs/sap_gui_for_windows)
- [MCP Specification](https://modelcontextprotocol.io/docs)
- [Contributing Guide](CONTRIBUTING.md)
- [Client Setup Guide](docs/CLIENTS.md)
- [Tool Catalog](docs/TOOLS.md)
- [Project Overview](docs/OVERVIEW.md)

## License

MIT. See [LICENSE](LICENSE).

## Disclaimer

This project is not affiliated with, endorsed by, or sponsored by SAP SE. SAP, SAP GUI, and other SAP products mentioned are trademarks of SAP SE.

Use of this software with SAP systems should comply with your SAP licensing agreement and your organization's security policies.
