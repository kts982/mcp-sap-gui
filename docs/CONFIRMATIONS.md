# Confirmations

Some actions wait for your answer before they reach SAP. The server asks the question itself, through your MCP client, and decides on your answer: what the assistant does or says cannot get past it.

## What asks

**A save asks.** This cannot be switched off. It covers the standard ways to save:

- `F11` / the Save key, and the Save button in the toolbar (`wnd[0]/tbar[0]/btn[11]`);
- on a popup: a Save button, Yes or OK on a popup about saving, and any answer except Cancel to a popup whose question the server cannot read (SAP shows some questions in a browser control that scripting cannot read: the question then tells you to check the SAP window).

**Confirmation points** add more. Switch them on in the chat (*"ask me before every field write"*), or for every chat with the `--confirm` option:

| Point | What asks |
|---|---|
| `transactions` | Starting a transaction (`sap_execute_transaction`) |
| `batch_fields` | Filling several fields at once (`sap_set_batch_fields`) |
| `field_writes` | Each single write: a field, a table cell, a text editor, a checkbox, a radio button, a dropdown |
| `all_writes` | Every tool that can change something, including ones that only select, scroll or navigate. Strict and noisy; a save then asks twice, once for the point and once for the save |
| `attach` | Attaching to an SAP session: you pick the session from a list of every open one (system, session number, client, user, transaction, window title), unless the assistant names a system with exactly one open session |

## What a question looks like

The server's text says what is about to happen and where (*"You are about to send 'F11' which triggers Save (F11) in SAP. [...] It goes to screen 'Maintain User Profile (SU3, SAPLSUU5, 100)' on system DEV/100."*), and the answer is labelled with what a yes does: *Yes, save*, *Yes, run it*, *Yes, press 'No'*. Each client draws it its own way: a checkbox and Accept (Claude Code CLI), Yes / No (GitHub Copilot CLI), True / False and Submit (VS Code, Codex). Only a yes goes ahead; anything else, including closing the question, is a no.

## After you answer

- **Yes:** the action runs, and its result carries `confirmed_by_user: true`, so the assistant, which never sees the question, knows that you approved it.
- **No:** nothing is done. The assistant gets an error that says you declined, and the server's rules tell it not to retry the call or reach the same result another way, but to tell you and wait for your next instruction.

Three more rules protect the answer:

- **An answer that comes back faster than a person can give one** (under 250 ms) came from the client, not from you. A no still counts, and the error says that no person saw the question; a yes is refused.
- **An answer counts only for what the question showed.** A save goes only to the screen the question named: right before it sends the save, the server reads the screen, system, client, user and SAP session again, and if any of them changed while the question was open (you moved on in SAP GUI, a popup opened), nothing is saved and the result says `action: changed`. On MCP 2026-07-28 the server also reads SAP again when your answer arrives, for a popup and the list of sessions too. The assistant then reads the screen again and asks again if still needed.
- **An answer counts once.** It cannot be replayed to repeat the action.

## Clients that cannot show the questions

The question has to reach you. Some clients never show it: the Claude desktop app declines at once without showing anything, and its Chat and Cowork tabs do not support the questions at all ([Clients](CLIENTS.md) has the list). There, a save and every action at a confirmation point are refused, and the refusal tells the assistant to tell you and to use a client that shows the questions. A confirmation point cannot be switched off there either, since switching it off asks too.

## Switching points on and off

- **In the chat:** ask the assistant, which calls `sap_set_confirmation_points`. Adding a point takes effect at once. Removing one asks you first, and if you say no, it stays.
- **For every chat:** start the server with `--confirm`, for example `--confirm field_writes transactions`. These points cannot be removed in a chat.
- Points belong to the chat (the MCP session), not to the SAP connection: they stay after `sap_disconnect` and a later reconnect.

## See it before it happens

Ask the assistant to *"show me the values first"*: `sap_preview` shows the current screen, the session and the values about to be written, without writing anything. Most clients get a written summary and a screenshot; VS Code shows a card. A preview is no approval: the save still asks.

## In the audit log

With `--audit-log`, every question and answer gets a `confirmation` line: the point, the tool, the outcome (`asked`, `accepted`, `declined`, `changed`, `replayed`, `unsolicited_answer`, `invalid_answer`, `unsupported_client`), and `elapsed_ms`, how long the answer took. An answer faster than a person also has `no_dialog: true`, and a refused yes `client_answer: "accept"`. An accepted answer says who approved it: `approved_by: "user"` for a person, `"launch"` under `--client-approves` ([Scripted clients](#scripted-clients)). The other audit lines: [Security guide](SECURITY_GUIDE.md#audit-log).

## Scripted clients

This section is for programs, not for chat clients. A reviewed script that a person starts, naming the system and the change, cannot save otherwise: it answers the save question itself, at once, and such a yes is refused as no person's. Started with `--client-approves save`, the server takes every yes to a save question as the approval of the program that started it:

- The question is still asked, and the script has to answer it. A no still counts.
- Every yes to a save counts as the launcher's, a slow one too: with the option on, the server cannot tell a person's yes from the program's. The result carries `confirmed_by_launch: true` instead of `confirmed_by_user: true`, and the audit line `approved_by: "launch"`.
- Only `save` can be named. Confirmation points, the session picker of `attach` and the question that turns a point off still need a person.
- Over stdio only: the server refuses to start with `--transport http` or with `--code-mode`. It logs a WARNING at start-up that names the option and its points.
- The option is set on the command line only; no tool can set or clear it.

Never put it into a chat client's configuration: there the client would approve every save without you.

Two things to know when writing the client with the MCP Python SDK's `mcp.ClientSession` (mcp 2.3):

- `initialize()` stays on the handshake protocol (2025-11-25), where the SDK calls your `elicitation_callback` inside `call_tool`. On 2026-07-28 (`discover()`), `ClientSession` does not run the `input_required` loop for you: call the tool with `allow_input_required=True`, answer each request through `dispatch_input_request`, and call again with `input_responses` and `request_state`.
- `stdio_client` passes only a short default environment to the server, so `MCP_SAP_GUI_*` settings go in `StdioServerParameters(env=...)`:

```python
StdioServerParameters(
    command="uvx",
    args=["mcp-sap-gui", "--client-approves", "save", "--audit-log", "audit.jsonl"],
    env={"MCP_SAP_GUI_STRICT_ERRORS": "1"},
)
```

## Known gaps

The server recognises a save by how it is sent, not by what it does in SAP. What a function does depends on the screen: the same code may save a document on one screen and only a selection variant on another, and a customer's own screen can use any code. So these do not ask:

- A save through a menu (for example Table View → Save in SM30) or through a Save button in an ALV or application toolbar: those element IDs differ by screen and logon language, so the server cannot recognise them as a save.
- An application button or function that posts or changes data (Execute, Activate, Post, an ALV "More Methods" entry), or a function code typed into the command field.

To be asked before every action that can change something, switch on `all_writes`. The SAP user's authorizations remain the boundary for what can change at all.

With the experimental `--code-mode`, a script cannot ask a question on MCP 2026-07-28: an action that needs one is refused inside the script, and the assistant has to call that tool directly.
