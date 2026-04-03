# MCP Client Setup

This server is client-agnostic. If your MCP client can launch a local `stdio` server, it can use `mcp-sap-gui`.

## Shared Launch Definition

Use this launch definition as the baseline in any client:

```text
Command:   uv
Arguments: run python -m mcp_sap_gui.server
Transport: stdio
Working directory: <path-to-mcp-sap-gui>
```

Optional flags:

```text
--read-only
--debug
--allowed-transactions MM03 VA03 ME23N
```

## Claude Code

If you open the repository in Claude Code, the root `.mcp.json` is auto-discovered.

```bash
cd mcp-sap-gui
claude
```

For a global user-level setup, add this to `~/.claude/.mcp.json`:

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

## Claude Desktop

Add this to your Claude Desktop config:

- Windows: `%APPDATA%\Claude\claude_desktop_config.json`
- macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "sap-gui": {
      "command": "uv",
      "args": [
        "run", "--directory", "<path-to-mcp-sap-gui>",
        "python", "-m", "mcp_sap_gui.server"
      ],
      "cwd": "<path-to-mcp-sap-gui>"
    }
  }
}
```

The `--directory` flag is required so `uv` finds the project's virtual
environment regardless of where Claude Desktop launches the process.

Read-only example:

```json
{
  "mcpServers": {
    "sap-gui": {
      "command": "uv",
      "args": [
        "run", "--directory", "<path-to-mcp-sap-gui>",
        "python", "-m", "mcp_sap_gui.server", "--read-only"
      ],
      "cwd": "<path-to-mcp-sap-gui>"
    }
  }
}
```

## GitHub Copilot

For VS Code / Copilot Chat, create `.vscode/mcp.json` in the repository root:

```json
{
  "servers": {
    "sap-gui": {
      "type": "stdio",
      "command": "uv",
      "args": ["run", "python", "-m", "mcp_sap_gui.server"],
      "cwd": "<path-to-mcp-sap-gui>"
    }
  }
}
```

Then:

1. Open Copilot Chat.
2. Switch to `Agent` mode.
3. Start or enable the MCP server from the tools / MCP UI.

If you prefer user-level configuration instead of per-repo config, use the equivalent Copilot MCP settings flow for your editor and reuse the same launch definition.

## Gemini CLI

Add the server to your Gemini CLI `settings.json` under `mcpServers`:

```json
{
  "mcpServers": {
    "sap-gui": {
      "command": "uv",
      "args": ["run", "python", "-m", "mcp_sap_gui.server"],
      "cwd": "<path-to-mcp-sap-gui>",
      "trust": false
    }
  }
}
```

If you want read-only mode:

```json
{
  "mcpServers": {
    "sap-gui": {
      "command": "uv",
      "args": ["run", "python", "-m", "mcp_sap_gui.server", "--read-only"],
      "cwd": "<path-to-mcp-sap-gui>",
      "trust": false
    }
  }
}
```

## Codex

Codex supports MCP in both the CLI and IDE flows. OpenAI’s MCP docs currently show:

- MCP server management via the Codex CLI
- shared MCP configuration between CLI and IDE
- manual configuration in `~/.codex/config.toml`

Use the shared launch definition from the top of this file when configuring this server in Codex, and prefer the official Codex MCP documentation for the exact current setup flow.

## Any Other MCP Client

If the client supports a local `stdio` MCP server, configure:

- command: `uv`
- args: `run python -m mcp_sap_gui.server`
- cwd: your project root

Then verify by asking:

```text
Connect to my open SAP session and tell me what system I'm on
Show me the current screen info
List all editable fields on this screen
```

## Common Setup Mistakes

- `cwd` does not point to the project root where `pyproject.toml` lives
- dependencies were not installed with `uv sync`
- SAP Logon Pad is not running
- SAP GUI scripting is disabled in SAP GUI options or on the SAP server
- the client needs a restart after MCP config changes
- the server was started with `--read-only` or an allowlist that blocks your flow

## Official References

- OpenAI Docs MCP / Codex MCP docs: https://developers.openai.com/learn/docs-mcp
- GitHub Copilot MCP docs: https://docs.github.com/en/copilot/how-tos/provide-context/use-mcp
- Gemini CLI MCP docs: https://github.com/google-gemini/gemini-cli/blob/main/docs/tools/mcp-server.md
