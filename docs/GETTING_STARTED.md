# Getting started

This page takes you from nothing to a first answer from SAP in your AI assistant. It takes about ten minutes once SAP GUI Scripting is allowed.

## What you need

- **Windows** with **SAP GUI for Windows** installed, and SAP Logon running (or a session in SAP Business Client, experimental).
- **SAP GUI Scripting allowed**, in your SAP GUI options and on the SAP system. The second one is up to your Basis team: [Preparing SAP](SAP_SETUP.md) has what to ask for.
- **An AI assistant that supports MCP servers**, ideally one that shows the server's confirmation questions, so that you can approve a save. The Claude Code CLI, GitHub Copilot (CLI and VS Code) and the Codex CLI do; the Claude desktop app does not. Details per client: [Clients](CLIENTS.md).

## 1. Install uv

The server is published on PyPI, and `uvx` (part of [uv](https://docs.astral.sh/uv/)) downloads and runs it, with nothing to clone or install by hand. In PowerShell:

```powershell
winget install --id=astral-sh.uv -e
```

Without winget, use uv's own installer: `powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"`. Open a new terminal afterwards so `uvx` is on the path.

## 2. Add the server to your assistant

Every client needs the same three things:

```text
Command:   uvx
Arguments: mcp-sap-gui[screenshots]
Transport: stdio
```

The `[screenshots]` part makes screenshots 70-90% smaller. Where each client keeps this, with a ready-made snippet for each: [Clients](CLIENTS.md). For example, in the Claude Code CLI:

```powershell
claude mcp add sap-gui --scope user -- uvx "mcp-sap-gui[screenshots]"
```

Restart the client (or reload its MCP servers) after adding the server.

## 3. Log on to SAP

Open SAP Logon and log on as usual. The server attaches to the session that is already open, so your password never passes through the assistant. (`sap_connect` can open an SAP Logon entry and take credentials from a local `.env` file, but logging on yourself is the safer way.)

## 4. Ask something

```text
Connect to my open SAP session and tell me what system I'm on
Show me the current screen info
List all editable fields on this screen
Read the first 20 rows of the visible table
```

If several SAP sessions are open, name the system: *"connect to my DEV session"*.

More ideas of what to ask: [What it can do](OVERVIEW.md).

## When the assistant wants to change something

Saving always asks you first: the server puts a question in front of you (*"Yes, save"*), and nothing is saved until you answer yes. You can make the assistant ask before other kinds of changes too, for example before every field it fills: say *"ask me before every field write"*. How the questions look and what each answer does: [Confirmations](CONFIRMATIONS.md).

## Updating to a new release

`uvx` keeps the version it downloaded first, so a new release is not picked up by itself. Run this once, then restart your assistant:

```powershell
uvx "mcp-sap-gui[screenshots]@latest" --help
```

To stay on one release instead, put the version in the client's arguments: `mcp-sap-gui[screenshots]==<version>`. What changed in each release: [CHANGELOG](../CHANGELOG.md).

## If something does not work

[Troubleshooting](TROUBLESHOOTING.md) covers the usual causes: SAP Logon not running, scripting not allowed, the tools not showing up in the client, an action that is blocked.
