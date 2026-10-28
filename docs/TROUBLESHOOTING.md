# Troubleshooting

## The SAP tools do not appear in the client

- Restart the client after changing its MCP configuration, or reload its MCP servers. Most clients list the servers and their state: `/mcp` in Claude Code, Copilot CLI and Codex, the MCP view in VS Code.
- Check that `uvx` can start the server: run `uvx "mcp-sap-gui[screenshots]" --help` in a terminal. If `uvx` is not found, open a new terminal after installing uv. On the first launch `uvx` downloads the package from PyPI: behind a proxy, set `HTTPS_PROXY`; with an internal package mirror, point uv at it with `UV_DEFAULT_INDEX`.
- A project-level configuration (`.mcp.json`, `.vscode/mcp.json`, `.codex/config.toml`) is only read in a folder the client trusts.

## "Cannot connect to SAP GUI"

- SAP Logon, SAP Business Client or another SAP GUI session has to be running, and you logged on.
- Scripting has to be enabled in SAP GUI: Options → Accessibility & Scripting → Scripting → "Enable scripting".

## "Scripting disabled"

`sap_list_connections` reports such a system with `scripting_disabled: true` and no sessions. The SAP system does not allow scripting: your Basis team sets `sapgui/user_scripting` in `RZ11`. If they will not enable it for everyone, they can enable it for named users only, or read-only: [Preparing SAP](SAP_SETUP.md#where-scripting-is-restricted).

## A save or another change is refused

- **"does not support elicitation"** or **"answered ... without showing a dialog"**: your client cannot show the server's confirmation questions, so nothing that needs one can be approved there. Use a client that shows them: [Clients](CLIENTS.md).
- **"declined by user"**: you, or the client, answered no. Nothing was done.
- **"changed before the answer arrived"**: the SAP screen or popup changed while the question was open. Ask again.
- **The client gives up while the question is still open** (Codex, after 60 seconds by default): raise the client's tool timeout, see [Clients](CLIENTS.md#codex-cli).

## The action is blocked

- The server may run with `--read-only`.
- The transaction policy may block the transaction: the error names the active preset. Allow it in a policy file, or pick another preset, then restart the server; the policy cannot be changed from inside a chat. See [Transaction policy](SECURITY_GUIDE.md#transaction-policy).
- The server may run with `--allowed-transactions`, which allows only the transactions listed.

## "Element not found"

Element IDs differ between SAP systems, releases and customizing. The assistant should discover them on the screen (`sap_get_screen_elements`) instead of guessing; ask it to look at the screen again.

## SAP Logon exits during a read

The SAP Logon crash on SAP GUI 8.00 in [issue #22](https://github.com/kts982/mcp-sap-gui/issues/22) is fixed since 0.5.1. On another SAP GUI build that fails on one-call reads, switch them off per call site with `MCP_SAP_GUI_DISABLE_OBJECT_TREE` and find the call with `MCP_SAP_GUI_TRACE_FILE`: [Switches for a misbehaving SAP GUI build](SAP_GUI_VERSIONS.md#switches-for-a-misbehaving-sap-gui-build).

## A new release is not picked up

`uvx` keeps the version it downloaded first. Run `uvx "mcp-sap-gui[screenshots]@latest" --help` once, then restart the client: [Getting started](GETTING_STARTED.md#updating-to-a-new-release).
