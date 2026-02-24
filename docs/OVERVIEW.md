# MCP SAP GUI — AI-Assisted SAP Automation

## What Is This?

An open bridge between **AI assistants (Claude)** and **SAP GUI for Windows**. It allows Claude to see, read, and interact with SAP screens exactly like a human consultant would — but faster, and without manual repetition.

Built on the **Model Context Protocol (MCP)**, the open standard by Anthropic for connecting AI models to external tools and data sources.

## The Problem

SAP consulting work involves enormous amounts of repetitive, screen-by-screen interaction:

- **Customizing**: Navigating dozens of SPRO nodes, setting values, saving, moving to the next
- **Testing**: Executing the same transaction flow across 15 test cases, capturing screenshots each time
- **Documentation**: Reading configurations screen by screen, copying into Word documents
- **Knowledge transfer**: Explaining what's configured where, producing handover docs

A senior consultant's time is spent 60–70% on navigation and data entry, and 30–40% on actual decision-making.

## The Solution

Claude connects to SAP GUI through 50 specialized tools and can:

| Capability | What It Means |
|-----------|---------------|
| **Read any screen** | Fields (with metadata), tables, trees, status messages — structured data, not screenshots |
| **Navigate freely** | Execute transactions, press buttons, use menus, send function keys (F1-F12, Shift+F, Ctrl+) |
| **Handle popups** | Detect modal dialogs, read their content and buttons, respond appropriately |
| **Work with ALV grids** | Read data, get toolbar buttons, sort by column, inspect cell metadata, select rows |
| **Work with TableControls** | Read data, scroll through pages, get row metadata, select rows |
| **Work with trees** | Navigate SPRO-style tree structures, expand/collapse, drill into nodes |
| **Interact with all UI elements** | Text fields, checkboxes, radio buttons, dropdowns, tabs, text editors, comboboxes |
| **Discover screen layout** | Enumerate all elements, read toolbar buttons, identify input fields |
| **Read shell content** | Extract HTML from embedded viewers and other GuiShell subtypes |
| **Capture evidence** | Take screenshots at any point for documentation |
| **Respect security** | Read-only mode, transaction blocklists, transaction whitelists |

## How It Works

```
┌─────────────────────┐                    ┌──────────────────────┐
│                     │   Natural          │                      │
│   Consultant        │   Language         │   Claude (AI)        │
│   (you)             │ ──────────────►    │                      │
│                     │                    │   Understands SAP    │
│   "Configure        │                    │   context, decides   │
│    putaway          │                    │   which screens to   │
│    strategies       │                    │   navigate, which    │
│    for WH 1710"     │                    │   fields to set      │
│                     │                    │                      │
└─────────────────────┘                    └──────────┬───────────┘
                                                      │
                                              MCP Protocol
                                            (50 SAP tools)
                                                      │
                                           ┌──────────▼───────────┐
                                           │                      │
                                           │   MCP SAP GUI Server │
                                           │                      │
                                           │   Translates AI      │
                                           │   requests into      │
                                           │   SAP GUI Scripting  │
                                           │   API calls          │
                                           │                      │
                                           └──────────┬───────────┘
                                                      │
                                              COM Automation
                                                      │
                                           ┌──────────▼───────────┐
                                           │                      │
                                           │   SAP GUI            │
                                           │   for Windows        │
                                           │                      │
                                           │   Real SAP session   │
                                           │   visible on screen  │
                                           └──────────────────────┘
```

The consultant sees SAP GUI moving in real time — fields being filled, buttons pressed, screens navigating. It's transparent and auditable. The consultant stays in control and can intervene at any point.

## Use Cases

### 1. Customizing from Specification Documents

A consultant has a Word document specifying putaway strategy configuration for EWM. Instead of manually navigating 30 SPRO screens:

> "Read the document putaway_strategy.docx and configure EWM putaway strategies in warehouse 1710."

Claude reads the document, identifies the required transactions, navigates to each configuration screen, enters the values, and confirms. The consultant watches and verifies.

### 2. Configuration Analysis and Comparison

Before a go-live, the team needs to verify that DEV, QAS, and PRD systems have consistent customizing:

> "Connect to DEV and read all putaway strategy configurations for warehouse 1710. Then I'll connect you to QAS and we'll compare."

Claude extracts structured data from tables and trees, making it easy to compare configurations across systems.

### 3. Test Scenario Execution

A test manager has 20 test cases in a spreadsheet. Each requires executing a transaction, entering specific data, and capturing the result:

> "Execute transaction /SCWM/PRDO, create a production order with these parameters, and capture a screenshot of the result."

Claude executes each step, captures screenshots, and notes document numbers — evidence that would take a tester 15 minutes per case.

### 4. Knowledge Transfer Documentation

A consultant is leaving a project and needs to document what's configured:

> "Go through SPRO for EWM warehouse 1710 and document all customizing entries for putaway, stock removal, and wave management."

Claude navigates systematically through SPRO nodes, reads each table, and produces a structured summary.

### 5. Quick Data Lookups

For everyday consulting questions that would otherwise require opening SAP and navigating manually:

> "What logical systems are configured in the system? What's the material description for MAT-001?"

Claude answers in seconds by running the right transaction and reading the result.

## Live Demo: What It Looks Like

### Example 1: Reading a Table

```
User:  "Read the logical systems table in SPRO"

Claude: [connects to SAP session]
        [navigates to SPRO → SAP Customizing → table maintenance view]
        [reads table data]

        "Found 13 logical systems:

         | Logical System | Name                          |
         |---------------|-------------------------------|
         | DEVCLNT100    | DEV Client 100                |
         | QASCLNT200    | QAS Client 200                |
         | OUTRIDER      | External Distribution System  |
         | ..."
```

### Example 2: Filtering and Exploring

```
User:  "Filter the table to show only OUTRIDER"

Claude: [opens Selection → By Contents menu]
        [selects Logical System field]
        [enters OUTRIDER as filter value]
        [executes]

        "Filtered results: 1 entry found.
         OUTRIDER — External Distribution System"
```

### Example 3: Screen Discovery

```
User:  "What fields and buttons are on this screen?"

Claude: [reads all screen elements]

        "Current screen: Display Logical Systems
         Fields: Logical System (ctxtVIEW_FIELD), Name (txtVIEW_TEXT)
         Toolbar: Save, Back, Exit, Cancel, New Entries, Delete
         Menu: Table View, Edit, Goto, Selection, Utilities, System, Help"
```

## Security Model

The server is designed with production safety in mind:

| Safeguard | Description |
|-----------|-------------|
| **Read-only mode** | `--read-only` flag blocks all write operations at the server level |
| **Transaction blocklist** | Sensitive transactions (SU01, PFCG, SE16N) blocked by default |
| **Transaction whitelist** | Optionally restrict to only specific transactions (e.g., MM03, VA03) |
| **Visible execution** | All actions happen in the visible SAP GUI — nothing runs in the background |
| **Consultant oversight** | The consultant watches the screen and can interrupt at any time |
| **SAP authorization** | Standard SAP authorization checks still apply — the AI has only the permissions of the logged-in user |

## Technology Stack

| Component | Technology |
|-----------|-----------|
| AI Model | Claude (Anthropic) — via Claude Code or Claude Desktop |
| Protocol | MCP (Model Context Protocol) — open standard |
| SAP Integration | SAP GUI Scripting API — official, supported by SAP |
| Transport | COM automation via pywin32 on Windows |
| Language | Python 3.10+ |

No SAP system modifications required. No ABAP development. No RFC connections. Uses the same scripting API that SAP's own recording tools use.

## Roadmap

### Phase 1: MCP Server (Current)

The foundation — 50 tools covering SAP GUI interactions including:
- Full field, table, and tree support
- Both ALV grid and TableControl parity
- Popup detection and handling
- Toolbar discovery
- Shell content reading (HTMLViewer, etc.)
- Extended keyboard support (Shift+F, Ctrl+ combos)

Works with Claude Code (CLI) and Claude Desktop today.

**Status: Complete and tested (146 unit tests) against live SAP systems.**

### Phase 1.5: Slash Commands (Current)

Built-in workflow commands for common SAP tasks:
- `/sap-explore` — Auto-discover and summarize the current SAP screen
- `/sap-table-dump` — Read all rows from a table, handling pagination for both ALV and TableControl
- `/sap-status` — Quick session status check with screenshot

These are Claude Code custom commands that orchestrate multiple MCP tools into useful workflows.

**Status: Available in `.claude/commands/`.**

### Phase 2: Real-World Consulting Use

Use the MCP server in actual consulting projects to:
- Validate the tool set covers real scenarios
- Build a library of effective prompts and patterns
- Identify gaps and edge cases
- Measure time savings vs. manual work

**Status: Ready to start.**

### Phase 3: Automation Agent

A Python script that orchestrates Claude programmatically:
- Reads specification documents (Word, PDF) and executes them unattended
- Runs batches of test scenarios from spreadsheets
- Captures screenshots and document numbers into structured reports
- Built on Claude Code SDK — uses the same Claude Max subscription, no separate API costs

A proof-of-concept exists at `examples/sap_agent.py`.

**Status: Proof of concept ready. To be developed after Phase 2 validates the approach.**

### Phase 4: Full Integration

- Web interface to upload specs and review results
- Multi-system orchestration (run same config across DEV/QAS/PRD)
- Audit trail and compliance logging
- Integration with project documentation tools

**Status: Future.**

## Requirements

- Windows (SAP GUI is Windows-only)
- SAP GUI for Windows with scripting enabled
- Python 3.10+ with [uv](https://docs.astral.sh/uv/) package manager
- Claude Code or Claude Desktop with MCP support

No special SAP licenses beyond what the consultant already uses. The AI operates through the same GUI session, using the same user's authorizations.

## Getting Started

```bash
# Clone and install
git clone <repository-url>
cd mcp-sap-gui
uv sync --extra screenshots

# Add to Claude Code (automatic via .mcp.json)
claude

# Or add to Claude Desktop (see README.md for config)
```

Then simply ask Claude:

> "Connect to my SAP session and tell me what system I'm on."
