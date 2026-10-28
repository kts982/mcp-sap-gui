# Confirmations

Some actions wait for your answer before they reach SAP. The server asks the question itself, through your MCP client, and decides on your answer: what the assistant does or says cannot get past it.

## What asks

**Saving always asks.** This cannot be switched off:

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

The server's text says what is about to happen (*"You are about to send 'F11' which triggers Save (F11) in SAP"*), and the answer is labelled with what a yes does: *Yes, save*, *Yes, run it*, *Yes, press 'No'*. Each client draws it its own way: a checkbox and Accept (Claude Code CLI), Yes / No (GitHub Copilot CLI), True / False and Submit (VS Code, Codex). Only a yes goes ahead; anything else, including closing the question, is a no.

## After you answer

- **Yes:** the action runs, and its result carries `confirmed_by_user: true`, so the assistant, which never sees the question, knows that you approved it.
- **No:** nothing is done. The assistant gets an error that says you declined, and the server's rules tell it not to retry the call or reach the same result another way, but to tell you and wait for your next instruction.

Three more rules protect the answer:

- **An answer that comes back faster than a person can give one** (under 250 ms) came from the client, not from you. A no still counts, and the error says that no person saw the question; a yes is refused.
- **An answer counts only for what the question showed.** On MCP 2026-07-28 the server reads SAP again when your answer arrives. If the screen, the popup or the list of sessions changed in the meantime, the answer does not count and nothing is done; the assistant reads the screen again and asks again if still needed.
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

With `--audit-log`, every question and answer gets a `confirmation` line: the point, the tool, the outcome (`asked`, `accepted`, `declined`, `changed`, `replayed`, `unsolicited_answer`, `invalid_answer`, `unsupported_client`), and `elapsed_ms`, how long the answer took. An answer faster than a person also has `no_dialog: true`, and a refused yes `client_answer: "accept"`. The other audit lines: [Security guide](SECURITY_GUIDE.md#audit-log).

## Known gaps

- A save through a menu (for example Table View → Save in SM30) or through a Save button in an ALV or application toolbar does not ask: those element IDs differ by screen and logon language, so the server cannot recognise them as a save. The SAP user's authorizations still apply.
- With the experimental `--code-mode`, a script cannot ask a question on MCP 2026-07-28: an action that needs one is refused inside the script, and the assistant has to call that tool directly.
