# MCP SAP GUI Server

<!-- mcp-name: io.github.kts982/mcp-sap-gui -->

[![CI](https://github.com/kts982/mcp-sap-gui/actions/workflows/ci.yml/badge.svg)](https://github.com/kts982/mcp-sap-gui/actions/workflows/ci.yml)
[![Docs](https://github.com/kts982/mcp-sap-gui/actions/workflows/docs.yml/badge.svg)](https://github.com/kts982/mcp-sap-gui/actions/workflows/docs.yml)
[![Release](https://img.shields.io/github/v/release/kts982/mcp-sap-gui)](https://github.com/kts982/mcp-sap-gui/releases)

An [MCP](https://modelcontextprotocol.io/) server that lets AI assistants work in **SAP GUI for Windows**: ask in plain language, and the assistant reads screens and tables, fills in fields, runs transactions and handles popups, through SAP's own Scripting API. Saving always waits for your yes.

It works with any MCP client that can start a local server: Claude Code, GitHub Copilot, Codex, Google Antigravity and others.

<img src="https://raw.githubusercontent.com/kts982/mcp-sap-gui/master/docs/images/sap-preview-card.png" alt="sap_preview card in VS Code Copilot Chat: the SAP screen and the values about to be written" width="640">

## What it does

- **Looks things up:** display transactions, fields, ALV grids, table controls, classic report lists and trees, read in full however long they are.
- **Fills in and changes:** fields, checkboxes, dropdowns, whole forms, table cells, F4 value help, table maintenance and IMG activities.
- **Navigates:** transactions, buttons, function keys, menus, tabs, SPRO and other trees.
- **Handles popups**, recognising questions, warnings, F4 lists and SAP messages in any logon language.
- **Shows before it saves:** a preview of the screen and the values about to be written, and screenshots on request.

More, with example prompts: [What it can do](https://github.com/kts982/mcp-sap-gui/blob/master/docs/OVERVIEW.md).

## Requirements

- Windows with SAP GUI for Windows (8.10 verified, 8.00 works; [SAP GUI versions](https://github.com/kts982/mcp-sap-gui/blob/master/docs/SAP_GUI_VERSIONS.md)), and SAP Logon running.
- SAP GUI Scripting allowed in SAP GUI and on the SAP system: [Preparing SAP](https://github.com/kts982/mcp-sap-gui/blob/master/docs/SAP_SETUP.md).
- [uv](https://docs.astral.sh/uv/), which runs the server from PyPI.
- An MCP client. To approve saves, one that shows the server's confirmation questions (listed below).

## Quick start

1. Install uv: `winget install --id=astral-sh.uv -e`
2. Add the server to your MCP client:

   ```text
   Command:   uvx
   Arguments: mcp-sap-gui[screenshots]
   Transport: stdio
   ```

   In the Claude Code CLI, for example: `claude mcp add sap-gui --scope user -- uvx "mcp-sap-gui[screenshots]"`. Every client: [Clients](https://github.com/kts982/mcp-sap-gui/blob/master/docs/CLIENTS.md).
3. Log on to SAP in SAP Logon, then ask: *"Connect to my open SAP session and tell me what system I'm on."*

Step by step, including updating to a new release: [Getting started](https://github.com/kts982/mcp-sap-gui/blob/master/docs/GETTING_STARTED.md).

## Safety

- **Saves wait for your yes**, and you can make other changes wait too, for example every field write: [Confirmations](https://github.com/kts982/mcp-sap-gui/blob/master/docs/CONFIRMATIONS.md).
- **A transaction policy** decides which transactions the assistant may start; read-only mode and profiles limit it further.
- **An audit log** records every call, question and answer.
- **Everything runs as the logged-on SAP user**, under that user's authorizations, which remain the real boundary.

Details for admins, Basis and security teams: [Security guide](https://github.com/kts982/mcp-sap-gui/blob/master/docs/SECURITY_GUIDE.md).

## Which clients show the confirmation questions

Tested on 2026-10-28: the **Claude Code CLI**, the **GitHub Copilot CLI**, **VS Code with GitHub Copilot** and the **Codex CLI** show them. The **Claude desktop app** does not (its Code tab declines without showing anything, its Chat and Cowork tabs do not support them), so a save is refused there: use it for reading. The full table: [Clients](https://github.com/kts982/mcp-sap-gui/blob/master/docs/CLIENTS.md#which-clients-show-the-confirmation-questions).

## Documentation

| Page | For |
|---|---|
| [Getting started](https://github.com/kts982/mcp-sap-gui/blob/master/docs/GETTING_STARTED.md) | Installing it and the first questions |
| [What it can do](https://github.com/kts982/mcp-sap-gui/blob/master/docs/OVERVIEW.md) | What to ask, and what is covered |
| [Clients](https://github.com/kts982/mcp-sap-gui/blob/master/docs/CLIENTS.md) | Setup per MCP client, and which ones show the confirmation questions |
| [Confirmations](https://github.com/kts982/mcp-sap-gui/blob/master/docs/CONFIRMATIONS.md) | What asks, and what each answer does |
| [Troubleshooting](https://github.com/kts982/mcp-sap-gui/blob/master/docs/TROUBLESHOOTING.md) | When something does not work |
| [Preparing SAP](https://github.com/kts982/mcp-sap-gui/blob/master/docs/SAP_SETUP.md) | Enabling scripting, also for named users or read-only; SAP Business Client |
| [Configuration](https://github.com/kts982/mcp-sap-gui/blob/master/docs/CONFIGURATION.md) | Every option and environment variable; profiles, HTTP, code mode |
| [Security guide](https://github.com/kts982/mcp-sap-gui/blob/master/docs/SECURITY_GUIDE.md) | Safeguards, the transaction policy, the audit log, a recommended setup |
| [SAP GUI versions](https://github.com/kts982/mcp-sap-gui/blob/master/docs/SAP_GUI_VERSIONS.md) | Which SAP GUI releases work, and how |
| [Tools](https://github.com/kts982/mcp-sap-gui/blob/master/docs/TOOLS.md) | Every tool the assistant can call |
| [CHANGELOG](https://github.com/kts982/mcp-sap-gui/blob/master/CHANGELOG.md) | What changed in each release |

## What's new

**0.6.0** supports MCP 2026-07-28, the newest protocol version, next to the earlier ones, and the confirmation questions work on both. It needs Python 3.11 or later, retires `sap_set_policy_profile`, names the MCP client in the audit log, and comes with reorganised documentation. Every change: [CHANGELOG](https://github.com/kts982/mcp-sap-gui/blob/master/CHANGELOG.md).

## Contributing

Bug reports and pull requests are welcome: [CONTRIBUTING](https://github.com/kts982/mcp-sap-gui/blob/master/CONTRIBUTING.md). Security issues: [SECURITY](https://github.com/kts982/mcp-sap-gui/blob/master/SECURITY.md).

## License

MIT. See [LICENSE](https://github.com/kts982/mcp-sap-gui/blob/master/LICENSE).

## Disclaimer

This project is not affiliated with, endorsed by, or sponsored by SAP SE. SAP, SAP GUI, and other SAP products mentioned are trademarks of SAP SE.

Use of this software with SAP systems should comply with your SAP licensing agreement and your organization's security policies.
