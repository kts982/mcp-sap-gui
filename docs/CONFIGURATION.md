# Configuration

Options go after the package name in the client's server configuration, for example:

```json
{
  "mcpServers": {
    "sap-gui": {
      "command": "uvx",
      "args": ["mcp-sap-gui[screenshots]", "--confirm", "field_writes", "--audit-log", "C:\\sap\\audit.jsonl"],
      "env": { "MCP_SAP_GUI_STRICT_ERRORS": "1" }
    }
  }
}
```

`uvx "mcp-sap-gui[screenshots]" --help` lists them all. The server reads them once at startup: to change one, change the configuration and restart the client.

## Options

| Option | What it does |
|---|---|
| `--read-only` | Refuses every tool that can change something: field writes, buttons, transactions, keys, table and tree actions. Reading works |
| `--profile exploration\|operator\|full` | Which tools the assistant sees at all (default `full`). See [Profiles](#profiles) |
| `--confirm POINT ...` | Confirmation points for every chat: `transactions`, `batch_fields`, `field_writes`, `all_writes`, `attach`. A chat can add points but never remove these. See [Confirmations](CONFIRMATIONS.md) |
| `--policy-preset default\|abap-dev\|strict` | Which transactions the assistant may start (default: the preset named in the policy file, else `default`) |
| `--policy-file FILE` | A JSON transaction policy with your own blocks and allows |
| `--allowed-transactions T ...` | Older allowlist; it restricts on top of the policy |
| `--audit-log FILE` | Writes every tool call, question and answer as JSON lines. See [Audit log](SECURITY_GUIDE.md#audit-log) |
| `--strict-errors` | Returns a failed result as an MCP error. See [Strict errors](#strict-errors) |
| `--detach-on-exit` | On shutdown, leaves every SAP session open, also the ones the server opened with `sap_connect` |
| `--transport stdio\|http` | `stdio` (default) or streamable HTTP. See [HTTP transport](#http-transport) |
| `--host`, `--port` | HTTP bind address and port (default `127.0.0.1:8000`) |
| `--allowed-host HOST` | HTTP: another name clients use for this machine in the `Host` header (repeatable) |
| `--allowed-origin ORIGIN` | HTTP: a browser origin allowed to call the server (repeatable) |
| `--code-mode` | Experimental. See [Code mode](#code-mode-experimental) |
| `--debug` | Debug logging (to stderr) |

The transaction policy, the presets and the policy file are described in the [Security guide](SECURITY_GUIDE.md#transaction-policy).

## Environment variables

Set them in the client's server configuration (`env`, as in the example above).

| Variable | What it does |
|---|---|
| `MCP_SAP_GUI_POLICY_FILE` | The policy file, when `--policy-file` is not given. Without either, the server reads `%APPDATA%\mcp-sap-gui\policy.json` if it exists |
| `MCP_SAP_GUI_STRICT_ERRORS=1` | The same as `--strict-errors` |
| `MCP_SAP_GUI_DISABLE_OBJECT_TREE`, `MCP_SAP_GUI_TRACE_FILE` | Switches for a SAP GUI build that misbehaves on one-call reads: [SAP GUI versions](SAP_GUI_VERSIONS.md#switches-for-a-misbehaving-sap-gui-build) |
| `PREFAB_BUNDLED_RENDERER=1` | Builds the preview card without loading its renderer from a CDN: [Preview cards](#preview-cards) |
| `SAP_USER`, `SAP_PASSWORD`, `SAP_CLIENT`, `SAP_LANGUAGE` | Logon data for `sap_connect`. Passwords are never a tool parameter and never reach the client or the audit log. Logging on in SAP Logon yourself and attaching with `sap_connect_existing` is the safer way, since these values sit in a configuration file in plain text. In a source checkout a `.env` file in the repository root works too |

## Extras

The package has optional parts, chosen in square brackets: `mcp-sap-gui[screenshots,apps]`.

| Extra | What it adds |
|---|---|
| `screenshots` | Screenshots 70-90% smaller (Pillow). Recommended |
| `apps` | The rendered `sap_preview` card in clients that support MCP Apps |
| `code-mode` | The sandbox for `--code-mode` |

## Profiles

Every tool is tagged `read` or `write`, and the profile decides which ones the assistant sees:

| Profile | Tools | What the assistant can do |
|---|---|---|
| `exploration` | `read` | Look only: inspect screens, read tables, take screenshots. It cannot navigate or change anything |
| `operator` | `read`, `write` | Normal SAP work: start transactions, fill fields, press buttons. The transaction policy and the confirmation questions still apply |
| `full` | all | All tools. The default |

The profile applies to every chat the server serves. To be asked before writes in one chat only, use confirmation points.

## Strict errors

Most SAP failures come back as a normal result with an `error` key, so a client that only checks the MCP error flag treats a refused write as done. With `--strict-errors` (or `MCP_SAP_GUI_STRICT_ERRORS=1`) such a result, and a `sap_set_batch_fields` call with a field that was not set, is returned as an MCP error, with the same content. Off by default. Inside a `--code-mode` script `call_tool` still returns the result, so scripts check `error` themselves.

## HTTP transport

stdio is the default and has no network surface: the client starts the server and talks to it directly. With `--transport http` the server listens on `127.0.0.1:8000`:

```powershell
uvx "mcp-sap-gui[screenshots]" --transport http
uvx "mcp-sap-gui[screenshots]" --transport http --host 0.0.0.0 --port 9000 --allowed-host sap-box
```

- A request addressed to a host name other than localhost, 127.0.0.1, the bound address or an `--allowed-host` is refused, and so is a browser request from another origin unless it is listed with `--allowed-origin`. That stops a web page from reaching the server through DNS rebinding. MCP clients send no `Origin` header and are not affected.
- Over HTTP the server needs a client on the session-based MCP protocol (2025-11-25 or earlier, for example VS Code). A request on MCP 2026-07-28 carries no client identity, so the server could not keep one SAP session per client, and refuses it with an error that says so. stdio works with every protocol version.
- An HTTP session that ends (the client closes it, or it is idle for 30 minutes) releases its SAP session and its confirmation points, and closes the SAP sessions the server opened for it, unless the server runs with `--detach-on-exit`.

The HTTP transport has no authentication of its own: never make it reachable from other machines you do not control.

## Preview cards

`sap_preview` always returns a written summary and the screenshot, so it works in every client. Clients that support MCP Apps (for example VS Code with Copilot) show a card with the screenshot and a table of the values about to be written instead. Install the `apps` extra for that: `uvx "mcp-sap-gui[screenshots,apps]"`.

By default the card loads its renderer from the jsDelivr CDN, so on a PC without internet access the card silently does not render. `PREFAB_BUNDLED_RENDERER=1` in the server's environment makes the renderer self-contained (about 6.3 MB inline, no third-party domains).

## Code mode (experimental)

`--code-mode` (with the `code-mode` extra) replaces the tool list with four meta-tools (`search`, `get_schema`, `tags`, `execute`): the assistant writes a short script that chains SAP calls in a sandbox. It is faster and cheaper on long chained flows such as full table dumps, and slower on quick questions. The confirmation questions and the transaction policy apply inside the sandbox too; on MCP 2026-07-28 a script cannot ask a question, so an action that needs one is refused there and has to be called directly.

```powershell
uvx "mcp-sap-gui[screenshots,code-mode]" --code-mode
```
