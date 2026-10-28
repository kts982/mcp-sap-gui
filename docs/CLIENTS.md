# Clients

The server works with any MCP client that can start a local `stdio` server. What differs between clients is whether they show the server's **confirmation questions**: a save, and every action at a confirmation point you switched on, waits for your answer in the client ([Confirmations](CONFIRMATIONS.md)).

## Which clients show the confirmation questions

Tested on 2026-10-28:

| Client | Version tested | MCP protocol | Confirmation questions |
|---|---|---|---|
| Claude Code CLI | 2.1.312 | 2026-07-28 | Shown: a checkbox ("Yes, save") and Accept |
| GitHub Copilot CLI | 1.0.105 | 2026-07-28 | Shown: Yes / No |
| VS Code with GitHub Copilot | 1.143.1 | 2025-11-25 | Shown: True / False and Submit |
| Codex CLI | 0.165.0 | 2025-06-18 | Shown: True / False |
| Claude desktop app, Code tab | 2.1.308 (bundled Claude Code) | 2026-07-28 | Not shown: the app declines at once, without showing anything |
| Claude desktop app, Chat and Cowork | | 2025-11-25 | Not supported |

In a client that does not show them, everything that reads works, but a save, a popup answer that may save and any call at a confirmation point are refused, and the refusal says why. Nothing is ever done without an answer. In such a client, start the server with `--read-only` so the assistant does not try. Clients not listed here (Gemini CLI, Cursor and others) are untested: try a save on a test system and see whether the question appears.

## The launch definition

Every client needs the same three things. `uvx` downloads the released package from PyPI on first launch and runs it in its own environment:

```text
Command:   uvx
Arguments: mcp-sap-gui[screenshots]
Transport: stdio
```

Options go after the package name, for example `mcp-sap-gui[screenshots] --read-only` or `mcp-sap-gui[screenshots] --confirm field_writes --audit-log C:\sap\audit.jsonl`. Every option: [Configuration](CONFIGURATION.md).

Restart the client, or reload its MCP servers, after changing the configuration.

## Claude Code CLI

For every project:

```powershell
claude mcp add sap-gui --scope user -- uvx "mcp-sap-gui[screenshots]"
```

For one project only, commit a `.mcp.json` in its root (Claude Code asks once whether to trust it):

```json
{
  "mcpServers": {
    "sap-gui": {
      "type": "stdio",
      "command": "uvx",
      "args": ["mcp-sap-gui[screenshots]"]
    }
  }
}
```

`/mcp` in Claude Code shows whether the server is connected.

## GitHub Copilot CLI

```powershell
copilot mcp add sap-gui -- uvx "mcp-sap-gui[screenshots]"
```

That writes your user configuration (`~/.copilot/mcp-config.json`). Copilot CLI also reads a `.mcp.json` (the same format as above) or `.github/mcp.json` in a folder you trust. `/mcp` lists the servers.

Copilot CLI puts a server's instructions into its system prompt only for servers on its allowlist. Our instructions tell the assistant how to discover SAP screens and how to treat a declined question; to include them, start Copilot CLI with `--allow-all-mcp-server-instructions`.

## VS Code with GitHub Copilot

Create `.vscode/mcp.json` in your workspace, or run **MCP: Add Server** from the Command Palette to add it for every workspace:

```json
{
  "servers": {
    "sap-gui": {
      "type": "stdio",
      "command": "uvx",
      "args": ["mcp-sap-gui[screenshots]"]
    }
  }
}
```

Then open Copilot Chat, switch to Agent mode, and start the server from the tools list. VS Code also shows `sap_preview` as a rendered card with the screenshot and the values about to be written (install with `mcp-sap-gui[screenshots,apps]`; see [Configuration](CONFIGURATION.md#preview-cards)).

## Codex CLI

```powershell
codex mcp add sap-gui -- uvx "mcp-sap-gui[screenshots]"
```

That writes `~/.codex/config.toml`. Add a longer tool timeout there by hand: Codex waits for a confirmation question inside the tool call, and its default of 60 seconds can cut you off while you read the question.

```toml
[mcp_servers.sap-gui]
command = "uvx"
args = ["mcp-sap-gui[screenshots]"]
tool_timeout_sec = 300
```

For one project only, put the same section in `.codex/config.toml` in that folder; Codex reads it once you trust the folder.

## Claude desktop app

The app has three tabs, and none of them shows the confirmation questions (see the table above), so use it for reading, with `--read-only`.

- **Chat and Cowork** read `%APPDATA%\Claude\claude_desktop_config.json`:

  ```json
  {
    "mcpServers": {
      "sap-gui": {
        "command": "uvx",
        "args": ["mcp-sap-gui[screenshots]", "--read-only"]
      }
    }
  }
  ```

  Quit the app from the tray icon and start it again after editing the file.
- **Code** is Claude Code: it uses the same configuration as the Claude Code CLI above.

## Gemini CLI

Add the server under `mcpServers` in Gemini CLI's `settings.json`:

```json
{
  "mcpServers": {
    "sap-gui": {
      "command": "uvx",
      "args": ["mcp-sap-gui[screenshots]"],
      "trust": false
    }
  }
}
```

Not tested yet with the confirmation questions.

## Any other client

Give it the launch definition above. Then check the setup by asking:

```text
Connect to my open SAP session and tell me what system I'm on
```

Running the server from a source checkout instead of PyPI: [CONTRIBUTING](../CONTRIBUTING.md#running-from-a-source-checkout).

## Official client documentation

- Claude Code: https://code.claude.com/docs/en/mcp
- GitHub Copilot: https://docs.github.com/en/copilot/how-tos/provide-context/use-mcp
- Codex: https://developers.openai.com/learn/docs-mcp
- Gemini CLI: https://github.com/google-gemini/gemini-cli/blob/main/docs/tools/mcp-server.md
