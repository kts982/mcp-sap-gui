# Security guide

The server lets an AI assistant work in SAP GUI as the logged-on user. Everything it does runs under that user's SAP authorizations and shows in SAP's logs as that user's activity. **The SAP user's authorizations are the security boundary.** The safeguards below keep an assistant from going places by mistake and keep you in charge of what is written. They do not replace that boundary.

To report a vulnerability, see [SECURITY.md](../SECURITY.md).

## Safeguards at a glance

| Safeguard | What it does | Set with |
|---|---|---|
| [Transaction policy](#transaction-policy) | Which transactions the assistant may start | `--policy-preset`, `--policy-file` |
| [Confirmations](CONFIRMATIONS.md) | Saves always wait for your yes; confirmation points add more | `--confirm`, or in the chat |
| [Read-only mode](CONFIGURATION.md#options) | Refuses every tool that can change something | `--read-only` |
| [Profiles](CONFIGURATION.md#profiles) | Which tools the assistant sees at all | `--profile` |
| [Audit log](#audit-log) | Every call, question and answer as JSON lines | `--audit-log` |
| Credentials | Passwords are never a tool parameter and never reach the client, its history or the audit log | logon in SAP Logon, or `SAP_*` variables |
| Element ID checks | IDs the assistant passes are validated before they reach SAP GUI; full session paths are shortened to `wnd[0]/...` | always on |
| [HTTP checks](CONFIGURATION.md#http-transport) | Over HTTP, requests for another host name or from another web origin are refused (DNS rebinding) | `--allowed-host`, `--allowed-origin` |

Every tool is tagged `read` or `write` and carries the MCP hints `readOnlyHint` and `destructiveHint`, so a client can show what a call may do before it runs.

## Transaction policy

The policy decides which transactions an assistant may start. What is dangerous depends on who runs the server against which system, so it is configurable.

- **`--policy-preset default`** (the default) blocks user and role administration (`SU01`, `SU10`, `SU01D`, `PFCG`, `SU53`), logs and dumps (`SM21`, `ST22`), development and program execution (`SE38`, `SA38`, `SE80`, `SE37`, `SE16N`), and system administration (`STMS`, `SCC4`, `RZ10`, `RZ11`, `SM36`, `SM49`, `SM69`, `SM59`, `STRUST`, `SICF`, `SM01`, `SM18`, `SM19`). Customizing and display work (`SM30`, `SM34`, `SPRO`, `SE16`, `SE11`) stays open.
- **`--policy-preset abap-dev`** additionally allows `SA38`, `SE11` and `SE80`, for developers testing their own programs.
- **`--policy-preset strict`** blocks everything the policy file does not allow: an allowlist.
- **`--policy-file policy.json`** adds to the preset: `{"preset": "abap-dev", "block": ["ZHR*"], "allow": ["SM59"]}`. Entries are transaction codes or patterns (`SU*`, `/SCWM/*`), and `allow` always wins over `block`. Without the option the server reads `$MCP_SAP_GUI_POLICY_FILE`, else `%APPDATA%\mcp-sap-gui\policy.json` if it exists. Keep the file outside the assistant's workspace, so it cannot edit it.
- Two examples ship in the repository's [`examples/`](../examples/) folder: `policy.example.json` starts from `default`, also blocks a customer namespace (`ZHR*`) and `SE16`, and allows `ST22` for dump analysis; `policy.strict.example.json` is an allowlist for a display-only assistant. To make one your own policy:

  ```powershell
  New-Item -ItemType Directory -Force "$env:APPDATA\mcp-sap-gui"
  Copy-Item examples\policy.example.json "$env:APPDATA\mcp-sap-gui\policy.json"
  ```

How the policy holds:

- It is read once at startup and written to the audit log with the file's SHA-256. No tool can change it, and a policy file that cannot be read stops the server instead of falling back to something more permissive.
- Matching ignores case and handles the `/n`, `/o` and `/*` prefixes and spaces.
- Writing a blocked transaction into the command field, or into a field that works as one (`txtOK_CODE` and similar), is blocked as well.
- The older `--allowed-transactions T1 T2 ...` restricts on top of the policy.

It is a guardrail against an assistant's mistakes, not a security boundary: a blocked transaction can often be reached another way, through a menu path or a screen you left open. The SAP user's authorizations are the boundary.

## Audit log

`--audit-log FILE` writes one JSON object per line. Arguments are recorded with secrets masked: an argument named like a password, token or secret, and a value written into a field whose ID or column looks like one (`pwdRSYST-BCODE`, `PASSWORD`). The field itself stays visible, so the log shows where a value went.

| `event` | When | Fields |
|---|---|---|
| `policy` | At startup | `preset`, `source`, `block`, `allow`, `file_sha256` (when a policy file was read) |
| `tool_call` | Every tool call | `ts`, `tool`, `args`, `status` (`ok`, `error`, or `input_required` when the call asks you a question on MCP 2026-07-28), `duration_ms`, `error` and `error_msg` on a failure, `sap` (the system, client and user the chat is attached to), `protocol` (the MCP protocol version), `mcp_client` (the client's name and version as it reports them) |
| `confirmation` | Every question and answer | `point`, `tool`, `outcome`, `elapsed_ms`, `no_dialog`, `client_answer`: see [Confirmations](CONFIRMATIONS.md#in-the-audit-log) |
| `attach` | The chat is attached to an SAP session | `system`, `client`, `user`, `session`, `chosen_by` (`named`, `dialog` or `index`) |
| `object_tree` | At startup, when one-call reads are switched off somewhere | the switched-off call sites and the trace file |

## Recommended setup

- **Test and trial systems first.** Run against production only when your organisation has decided to, with the controls below.
- **A dedicated SAP user with minimal authorizations** for the work at hand.
- **Scripting per user, not for everyone:** `sapgui/user_scripting_per_user` with the `S_SCR` authorization limits scripting to named users ([Preparing SAP](SAP_SETUP.md#where-scripting-is-restricted)).
- **`--read-only`** for looking things up, and an allowlist (`--policy-preset strict` with a policy file) for repeated automation.
- **`--audit-log`** on the server, together with SAP's own logging.
- **A client that shows the confirmation questions** ([Clients](CLIENTS.md)), so that saves need your yes.
- **Never make the HTTP transport reachable for users or machines you do not trust.** stdio, the default, has no network surface.

## For Basis and security teams

In short: the server is a local process on the user's own Windows PC. It attaches to SAP GUI sessions the user has already opened, through SAP's own Scripting API, and every action runs as that user, under that user's authorizations, and shows as that user's activity in SAP's logs. Scripting can be limited to named users (`S_SCR`) or to read-only on the SAP side ([Preparing SAP](SAP_SETUP.md#where-scripting-is-restricted)); on the server side, the transaction policy, read-only mode, the confirmation questions and the audit log above add control and a record of what the assistant did.
