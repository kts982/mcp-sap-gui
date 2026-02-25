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

**Write tools** (blocked in read-only): `sap_execute_transaction`, `sap_send_key`, `sap_set_field`, `sap_press_button`, `sap_select_menu`, `sap_select_checkbox`, `sap_select_radio_button`, `sap_select_combobox_entry`, `sap_select_tab`, `sap_set_batch_fields`, `sap_set_textedit`, `sap_set_focus`, `sap_select_table_row`, `sap_double_click_cell`, `sap_modify_cell`, `sap_set_current_cell`, `sap_press_alv_toolbar_button`, `sap_select_alv_context_menu_item`, `sap_scroll_table_control`, `sap_select_all_table_control_columns`, `sap_press_column_header`, `sap_select_all_rows`, `sap_select_multiple_rows`, `sap_expand_tree_node`, `sap_collapse_tree_node`, `sap_select_tree_node`, `sap_double_click_tree_node`, `sap_double_click_tree_item`, `sap_click_tree_link`, `sap_get_tree_node_children` (only when expand=True)

**Read tools** (always allowed): `sap_connect`, `sap_connect_existing`, `sap_list_connections`, `sap_get_session_info`, `sap_get_screen_info`, `sap_read_field`, `sap_get_combobox_entries`, `sap_read_textedit`, `sap_read_table`, `sap_get_alv_toolbar`, `sap_get_column_info`, `sap_get_current_cell`, `sap_get_table_control_row_info`, `sap_get_cell_info`, `sap_get_popup_window`, `sap_get_toolbar_buttons`, `sap_read_shell_content`, `sap_read_tree`, `sap_find_tree_node_by_path`, `sap_search_tree_nodes`, `sap_get_screen_elements`, `sap_screenshot`

### Adding New Blocked Transactions

Edit `ServerConfig.blocked_transactions` in `server.py`.

## MCP Instructions & Resource

The server provides SAP navigation knowledge to all MCP clients via two mechanisms:

### `_INSTRUCTIONS` (in `server.py`)

Passed to `FastMCP(instructions=...)`. Injected into every client's system prompt during MCP `initialize`. Keep this **concise** — agents read it every session. Covers:
- Getting started / connect flow
- Screen discovery workflow (always discover, never guess IDs)
- Popup handling (`active_window` check)
- Table types, pagination, Position button
- SPRO tree navigation (`search_tree_nodes` + `click_tree_link`)
- Key reference with SAP-specific gotchas (F5 = "New Entries" in table maintenance)
- Common mistakes to avoid

### `_SAP_GUI_GUIDE` (in `server.py`)

Exposed as `@mcp.resource("docs://sap-gui-guide")`. Detailed reference for clients that support resources. Contains element type tables, ID naming conventions, full SPRO step-by-step, table maintenance patterns, Web Dynpro fallback, etc.

### Maintenance

When adding new tools or discovering new navigation patterns:
1. Update `_INSTRUCTIONS` if the pattern is critical (agents should always know it)
2. Update `_SAP_GUI_GUIDE` for detailed reference information
3. Update tests in `TestInstructionsAndResource` if new key sections are added

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
# Get only input fields (much smaller response on complex screens)
elements = controller.get_screen_elements(changeable_only=True)
# Or filter by type
elements = controller.get_screen_elements(type_filter="GuiTextField,GuiCTextField")
```

### Handling Popups

`get_screen_info()` returns `active_window` which is `"wnd[1]"` (or higher) when a popup is open. Every action tool includes this in its response, so popups are detected automatically — no separate check needed.

```python
# After any action, check the screen response:
result = controller.press_button("wnd[0]/tbar[1]/btn[8]")
if result["screen"]["active_window"] != "wnd[0]":
    # A popup appeared — use get_popup_window() for full details
    popup = controller.get_popup_window()
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

### Response Size Optimization

Three tools support optional filtering to reduce context token usage:

- **`sap_get_screen_elements`**: `type_filter` (CSV, e.g. `"GuiTextField,GuiCTextField"`) and `changeable_only=True` to return only editable input fields instead of all 200+ elements on complex screens.
- **`sap_read_table`**: `columns_only=True` for schema discovery (no data), `columns` (CSV) to fetch specific columns, `start_row` to paginate. Typical workflow: `columns_only=True` → inspect schema → `columns="COL_A,COL_B"` with `start_row` pagination.
- **`sap_read_textedit`**: `max_lines` to cap output for large text editors (0 = all).

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


# Multi-CLI Orchestration Guide

You (Claude Code) are the **orchestrator**. You have access to two additional AI CLI tools — `codex` (OpenAI) and `gemini` (Google Gemini CLI). Use them as specialized workers when their strengths match the task. **You remain the decision-maker, architect, and integrator.**

All three CLIs run on subscription plans — there is **no cost concern**. Always optimize for **quality and speed**, never for token savings. Let each tool use its best available model. Feed full context when it helps.

---

## Core Principle: File-Based Context Passing

Never pipe large context through CLI arguments. Instead:

1. Write a focused task spec to `.tasks/<n>.task.md`
2. Invoke the CLI with that file
3. Capture output to `.tasks/<n>.result.md`
4. Read the result back and integrate it

```bash
# Setup (run once per project)
mkdir -p .tasks
echo ".tasks/" >> .gitignore
```

---

## Routing: When to Delegate vs Handle Directly

Route by **task shape**, not by language or domain.

### KEEP in Claude Code (yourself)

- **Architecture & design** — system design, API contracts, module boundaries
- **Multi-file refactoring** — changes that ripple across the project
- **Tasks needing project context you already hold** — if you'd have to
  dump half the project into a task spec, just do it yourself
- **Ambiguous or underspecified tasks** — you can ask the user to clarify;
  workers can't
- **Orchestration itself** — decomposition, integration, final review
- **Domain-specific work** where you have strong knowledge the others lack
  (e.g., SAP/ABAP/EWM/SD, MCP protocol internals)
- **Complex debugging** requiring holistic understanding of state and flow
- **Security-sensitive code** — auth, crypto, access control

### DELEGATE to Codex CLI (`codex`)

Codex excels at **focused, well-scoped code generation** with clear specs.

- Single function, single file, or single module generation
- Unit/integration test generation from interfaces or examples
- Implementing a function/method from a clear signature + docstring
- Code translation between languages (e.g., Python → Go, JS → TS)
- Regex, SQL queries, shell scripts from plain-English specs
- Boilerplate and scaffolding (REST handlers, CLI arg parsing, CRUD, etc.)
- Algorithm implementation from a description
- Generating types/schemas/protobuf from specs or examples
- Quick, well-defined bug fixes with clear repro steps

**Invocation:**
```bash
# Use exec subcommand for non-interactive mode
codex exec --full-auto \
  "$(cat .tasks/codex-job.task.md)" > .tasks/codex-job.result.md 2>&1
```

### DELEGATE to Gemini CLI (`gemini`)

Gemini's 1M token context window makes it the best choice for **large-context tasks**. Don't hesitate to feed it entire files or even whole directories.

- Analyzing large codebases, modules, or log files
- Reviewing large diffs or pull requests
- Documentation generation from large or complex code
- Summarizing lengthy logs, traces, stack dumps, or CI output
- Cross-referencing many files to answer architectural questions
- Generating comprehensive test plans or migration checklists
- Processing/analyzing large data files (CSV, JSON, XML, ABAP transports)
- "Find all places in this codebase where X happens"

**Invocation:**
```bash
# Use -p for non-interactive mode; feed file content inline or via stdin
gemini -p "$(cat .tasks/gemini-job.task.md)" \
  --sandbox false > .tasks/gemini-job.result.md 2>&1
# For large file context, prepend file contents to the prompt or use
# --include-directories for whole directories
```

---

## Decision Flowchart

```
New task arrives
│
├─ Can I do this faster than writing a spec? ──────→ DO IT YOURSELF
│
├─ Is it ambiguous / needs user clarification? ────→ DO IT YOURSELF (ask user)
│
├─ Does it need project context I already hold? ───→ DO IT YOURSELF
│
├─ Is it a focused code-gen task with clear spec?
│  ├─ Single file / function / module? ────────────→ CODEX
│  └─ Needs reading many files first? ────────────→ GEMINI (analyze) then
│                                                    CODEX (generate) or YOURSELF
│
├─ Does it involve analyzing/reviewing large input?
│  ├─ Large files, logs, diffs, traces? ───────────→ GEMINI
│  └─ Large codebase sweep / search? ─────────────→ GEMINI
│
├─ Are there independent subtasks? ────────────────→ PARALLELIZE (see below)
│
└─ Default ────────────────────────────────────────→ DO IT YOURSELF
```

---

## Parallel Execution

When subtasks are independent, run them concurrently:

```bash
(codex exec --full-auto \
  "$(cat .tasks/codex-api.task.md)" > .tasks/codex-api.result.md 2>&1) &
PID1=$!

(gemini -p "$(cat .tasks/gemini-review.task.md)" \
  --sandbox false > .tasks/gemini-review.result.md 2>&1) &
PID2=$!

wait $PID1 $PID2
echo "Both tasks complete."
```

You can also run **two Codex tasks** or **two Gemini tasks** in parallel — they're separate processes.

---

## Task Spec Template

When writing task specs for delegation, use this structure. Make each spec **self-contained** — the worker has zero project context.

```markdown
## Task
[One sentence: what to produce]

## Context
[Only what the worker needs. Include relevant type definitions, interfaces,
function signatures, or data structures inline. Don't reference files the
worker can't see — paste the relevant parts.]

## Requirements
1. [Concrete, testable requirement]
2. [Another one]
3. ...

## Constraints
- Language / framework / version
- Patterns to follow or avoid
- Dependencies available
- Style conventions (naming, error handling, etc.)

## Output Format
[Exactly what to produce: a complete file, a function, a diff, a report, etc.]

## Examples (optional but helpful)
[Example input/output, example usage, or a similar function to mimic]
```

**Rules for good task specs:**
- Be specific and self-contained
- Include relevant types/interfaces/signatures inline
- Specify language, version, and conventions
- Include examples when the spec alone might be ambiguous
- For Go: mention error handling style, module path
- For Python: mention Python version, type hints expectation, async or sync
- For ABAP: include relevant data dictionary definitions and naming conventions

---

## Quality Gates

**Always review delegated output before using it.** Run through:

- [ ] **Correctness** — does it solve the task? Does it compile/run?
- [ ] **Consistency** — does it match existing project style and patterns?
- [ ] **Completeness** — all requirements met? Edge cases handled?
- [ ] **Integration** — does it fit cleanly into the codebase?

If a result is poor, **refine the spec and retry once**. If it fails again, do it yourself.

---

## Anti-Patterns

1. **Over-delegating** — don't delegate a 10-line function; just write it
2. **Under-specifying** — a vague task spec produces vague results
3. **Context dumping** — don't paste the entire project into a task spec;
   include only what's relevant
4. **Trust without review** — always verify delegated output before integration
5. **Sequential when parallel works** — if tasks are independent, parallelize
6. **Re-delegating repeated failures** — if a worker fails twice, do it yourself
7. **Delegating what you already know** — if the answer is in your context, use it
8. **Forcing models** — let each CLI choose its best model for the task
