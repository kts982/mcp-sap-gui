# Changelog

Every release and what changed for someone running the server. The [GitHub release pages](https://github.com/kts982/mcp-sap-gui/releases) also list the merged pull requests and their authors.

## 0.5.4 — 2026-10-27

A security release for the HTTP transport, a session profile that holds, and the server's own version in `serverInfo`. stdio, the default transport, was not affected by the security issue.

### Safety

- **The HTTP transport refuses requests a web page could forge** ([GHSA-rqrh-r823-7mqf](https://github.com/kts982/mcp-sap-gui/security/advisories/GHSA-rqrh-r823-7mqf)). With `--transport http`, any `Host` and `Origin` header was accepted, so a page opened in a browser could point its own domain at 127.0.0.1 (DNS rebinding), act as an MCP client and use the bound SAP session, answering the confirmation prompts itself. A request addressed to another host name now gets 421, and a browser request from another origin 403. `--allowed-host` and `--allowed-origin` add names and origins, for example with `--host 0.0.0.0`; such a bind without `--allowed-host` logs a warning. MCP clients send no `Origin` and are not affected. stdio, the default, has no network surface.
- **A session profile no longer lapses after a day** ([#37](https://github.com/kts982/mcp-sap-gui/pull/37)). `sap_set_policy_profile` hid tools through FastMCP's session state, which expires 24 hours after it was set: the hidden write tools were then listed and callable again, without notice. The server now keeps the profile itself and refuses a tool outside it on every call, also from a `--code-mode` script. The `--profile` floor and the confirmation points were not affected.

### Fixed

- **The server reported FastMCP's version as its own** ([#36](https://github.com/kts982/mcp-sap-gui/pull/36)). The `initialize` answer named the server `mcp-sap-gui` but carried the installed FastMCP's version (0.5.3 reported `3.4.11`). It now reports the package version.

## 0.5.3 — 2026-10-27

Session commands refused, only the server's own SAP sessions closed, and a choice of the session to attach to. Six of these changes were contributed by [@nelz0n](https://github.com/nelz0n); the session changes were checked live on SAP GUI 8.10 and 8.00.

### Safety

- **The audit log no longer writes a password typed into a field.** `sap_set_field` on a password field (e.g. `pwdRSYST-BCODE`) logged the value in clear text; only `sap_set_batch_fields` masked it. Now, when an argument names a secret-looking target (a field ID, a table column, a text editor), the value written into it is logged as `***`. The target stays visible, so the audit still shows which field was written.
- **Session commands are blocked, and three ways around the transaction policy are closed** ([#25](https://github.com/kts982/mcp-sap-gui/pull/25), by [@nelz0n](https://github.com/nelz0n)).
  - `/nex` (log off, closing every session of the connection without asking), `/nend` and `/i` (close a session; `/i3` closes session 3) passed the `default` and `abap-dev` presets, through `sap_execute_transaction` and through the command field. They are now refused with their own message; a policy file can still allow one (`"allow": ["EX"]`).
  - `sap_set_textedit` on the command field skipped the transaction check, so `/nSU01` got through. It is now checked like `sap_set_field`.
  - A command-field value that starts a transaction but does not parse as a whole (`/*SU01 USR02-BNAME=X`, which skips the first screen with a field filled) was not checked at all. The code it names is now checked, and one the policy cannot check is refused.
- **The server closes only the SAP sessions it opened** ([#26](https://github.com/kts982/mcp-sap-gui/pull/26), by [@nelz0n](https://github.com/nelz0n)).
  - After `sap_connect`, an `/o` command moved the binding to `GuiApplication.ActiveSession` (which can be a window the user works in, on any connection) or to the last session of the connection (which can be an existing one), and `sap_disconnect` or server shutdown then closed that session, while the one the server had opened stayed open. The server now keeps the sessions it opened, the `sap_connect` logon session and `/o` windows opened from it, and closes only those, newest first; one the user already closed is skipped.
  - `/o` binds only a session that was not on the bound connection before the command (it waits up to 2 seconds for it) and reports `new_session_bound`. An `/o` window opened from an attached session is attached too, so it is never closed.
  - `--detach-on-exit` leaves every session open on shutdown, including the ones the server opened.
  - `sap_connect` is no longer annotated read-only, and `sap_disconnect` is annotated destructive. Both keep the `read` tag, so every profile can still log on and detach.

### Added

- **Choose the SAP session to attach to.** `sap_connect_existing` takes `system` (`DEV`, or `DEV/100` with the client) for when the user named the system: one open session on it is attached, and several or none return an error that lists what is open. A new opt-in confirmation point `attach` (`--confirm attach`, or `sap_set_confirmation_points`) lets the user pick the session in a dialog, unless the call names a system with exactly one open session; a client without dialogs gets an error telling the agent to ask in the chat. `sap_list_connections` adds each session's `session_number` and window `title`, the response says how the session was chosen (`chosen_by`), and the audit log records an `attach` event. After the dialog or the name lookup, the server checks that the session it attached is the one chosen (sessions can open or close meanwhile) and attaches nothing if not; attaching and the check are one step, so no other tool call reaches the session before it is checked. Without the point and without `system`, attaching works as before.
- **The audit log names the SAP system of each call.** Every `tool_call` line carries `sap: {"system", "client", "user"}` of the session the chat is bound to, so a log that spans several systems can be split by system. An attach names the system it attached, a disconnect the one it left; while nothing is bound the key is left out. The values are read when a session is attached or opened, and `sap_get_session_info` refreshes them (a `sap_connect` window has no user until the logon).
- **Which radio button or checkbox is selected.** `sap_read_field`, `sap_get_popup_window` (`interactive_elements`) and `sap_get_screen_elements` report `selected` for radio buttons and checkboxes. It is read over COM for each of them, never through GetObjectTree, which crashes SAP GUI 8.00 when asked for it. A changeable radio button or checkbox that is selected in a popup is listed in `prefilled_inputs` with the notice, like a pre-filled field: a popup can open with an option already chosen (SCAL's holiday dialog: "with fixed date").
- **`--strict-errors`** ([#28](https://github.com/kts982/mcp-sap-gui/pull/28), by [@nelz0n](https://github.com/nelz0n)). Most SAP failures come back as a normal result with an `error` key, so a client that checks only `isError` treats a refused write as done. With `--strict-errors` (or `MCP_SAP_GUI_STRICT_ERRORS=1`) such a result, and a `sap_set_batch_fields` call with a field that was not set, is returned with `isError: true` and its structured content unchanged. Off by default. Inside a `--code-mode` script `call_tool` still returns the result, so scripts check `error` themselves.
- **Read results say when they are incomplete** ([#29](https://github.com/kts982/mcp-sap-gui/pull/29), by [@nelz0n](https://github.com/nelz0n)). A table-control read adds `next_start_row` and `has_more`, echoes `start_row`, and adds a `start_row_note` when the table could not be scrolled to the requested row. `sap_read_tree` adds `available_nodes` and `truncated`, `sap_search_tree_nodes` adds `searched_nodes` and `stopped_early`, and `sap_read_table`, `sap_read_tree` and `sap_search_tree_nodes` add `capped_to` when the server lowered the limit you asked for.

### Changed

- **`sap_set_field` refuses a value that does not fit before writing** ([#28](https://github.com/kts982/mcp-sap-gui/pull/28)). A value longer than the field (`MaxLength`) or a field that is not changeable gives `status: "rejected"` with the reason, and `max_length` / `value_length` for a value that is too long; the field keeps its content. Before, both came back as "Could not set field". `sap_set_batch_fields` explains a refused field the same way.
- **The audit log marks failed results as errors** ([#28](https://github.com/kts982/mcp-sap-gui/pull/28)). A result with an `error` key, or a batch with a field that was not set, is logged with `status: "error"` and an `error_msg`, at WARNING, in every mode. Before, only an exception was. Counts of errors in existing audit logs go up.
- **A declined save or popup answer carries `error`** ([#28](https://github.com/kts982/mcp-sap-gui/pull/28)): "Save declined by user (…). Nothing was sent." / "Popup answer declined by user (…). Nothing was pressed.", next to the unchanged `status: "cancelled"`.
- **An answer that no person gave is treated as one.** The Claude Code desktop app (tested 2026-10-19) answers every confirmation prompt with "decline" within milliseconds and shows nothing; the Claude Code CLI shows them. An answer faster than a person can give one (under 250 ms) came from the client: a "no" still counts, and the error adds that no dialog was shown and the request cannot be approved in that client; a "yes" is refused, because no person approved it, so a client or hook that answers for you cannot approve a save. Every `confirmation` audit line now has `elapsed_ms`, `no_dialog: true` for such an answer, and `client_answer: "accept"` for a refused yes. All prompts go through one helper with these rules.
- **`sap_set_batch_fields(validate=true)` no longer presses Enter when a field failed** ([#29](https://github.com/kts982/mcp-sap-gui/pull/29)): Enter on a half-filled screen checks only what was set. `validation` then says `performed: false` with the reason. `validate_partial=true` keeps the old behaviour.
- **`sap_get_popup_window` says where a popup's text is** ([#30](https://github.com/kts982/mcp-sap-gui/pull/30), by [@nelz0n](https://github.com/nelz0n)). Labels and display-only fields are in `texts`; the body of a classic list (an F4 hit list, a SAP note, a message long text) is in `list`, and its lines are never in `texts`. Every list popup now gets the same hint: read the lines with `sap_read_list(window_id=...)`, and in an F4 hit list pick a line by its ID, then Enter. The docstrings and [TOOLS.md](docs/TOOLS.md) say the same.

### Fixed

- **Table-control paging skipped rows** ([#29](https://github.com/kts982/mcp-sap-gui/pull/29)). One read returns at most the visible rows, but the guidance stepped `start_row` by `max_rows`, so every row between the visible ones and the next step was missed. Page with `start_row=next_start_row` while `has_more` is true. Checked live: an SM30 view with 137 entries reads every entry once.
- **A blank New Entries line with a dropdown read as an entry** ([#29](https://github.com/kts982/mcp-sap-gui/pull/29)). The dropdown's empty entry reads as a single space; a whitespace-only cell now counts as blank, so `empty_rows_from` marks where the data ends.
- **`validation.performed` was true when Enter was refused or failed** ([#29](https://github.com/kts982/mcp-sap-gui/pull/29)). It is now false, with the error.
- **One "allow scripting" dialog per SAP GUI process, not one per tool call** ([#32](https://github.com/kts982/mcp-sap-gui/pull/32), by [@nelz0n](https://github.com/nelz0n)). Since 0.5.0 every `sap_list_connections` and `sap_connect_existing` looked up the scripting engine of each SAP GUI process again, and a system with `sapgui/user_scripting_force_notification` set asks the user at each lookup, for each of its connections. The engine of each process is now kept and reused while it answers; one that stops answering (SAP Logon restarted) is looked up again, and a GUI server process that ended is forgotten.
- **`sap_connect` failed after SAP Logon was restarted, until the server restarted.** It kept the SAP Logon engine from its first use; after a restart that engine was dead and `OpenConnection` failed. It now uses the engine discovery keeps, checked before each use, so a new SAP Logon is picked up.

### Dependencies

- `pillow` 12.3.0 → 13.0.0 (the `screenshots` extra) for the advisories fixed in [Pillow 13.0.0](https://pillow.readthedocs.io/en/stable/releasenotes/13.0.0.html), among them [GHSA-78g4-xqwv-q3g8](https://github.com/advisories/GHSA-78g4-xqwv-q3g8) in the bundled OpenJPEG. They are all in decoding or encoding image files; the server opens only the PNG that SAP GUI writes for a screenshot. Pillow 13 needs Python 3.11 and there is no 12.x fix, so on Python 3.10 the `screenshots` extra now installs nothing: `sap_screenshot` still works, but the image is no longer downscaled or optimized.
- `certifi` 2026.6.17 → 2026.10.19 (through httpx) for [GHSA-6q48-6c7w-9gw2](https://github.com/advisories/GHSA-6q48-6c7w-9gw2): eight intermediate CAs under the bundled `emSign ECC Root CA - G3` had compromised keys, and the new bundle drops that root.
- `opentelemetry-api` 1.43.0 → 1.45.1 (through fastmcp) for [GHSA-xcq7-p3pf-p88h](https://github.com/advisories/GHSA-xcq7-p3pf-p88h) (CVE-2026-114603): a malformed W3C `tracestate` value in an incoming request could cost the server a lot of CPU time. Availability only.
- `fastmcp` 3.4.7 → 3.4.11. 3.4.10 backports a security fix (component authorization now applies after transforms, and a denied resource can no longer be served through a matching template; this server uses neither component authorization nor resource templates), and 3.4.11 requires the fixed `certifi`.
- `authlib` 1.7.2 → 1.9.1 (through fastmcp) for seven advisories published on 2026-10-27, among them [GHSA-6v38-4h7w-rj9h](https://github.com/advisories/GHSA-6v38-4h7w-rj9h) (a JWS with an empty signature passed validation) and [GHSA-8m4g-j4wp-wgcg](https://github.com/advisories/GHSA-8m4g-j4wp-wgcg) (an empty HMAC key let anyone sign a JWT). All seven are in JWT, JWS and OAuth handling; this server configures no authentication on either transport and does not load authlib.
- `dnspython` 2.8.0 → 2.9.0 on Python 3.11 and later (through pydantic's `email` extra, which fastmcp uses) for [GHSA-45c3-73f7-pv4m](https://github.com/advisories/GHSA-45c3-73f7-pv4m): parsing a crafted DNS answer could cost a lot of CPU time. dnspython parses DNS answers only for email-validator's deliverability check, which nothing here runs. 2.9.0 needs Python 3.11 and there is no 2.8.x fix, so Python 3.10 keeps 2.8.0 and the dependency audit skips this advisory.

### Python versions

- Python 3.10 reached end of life on 2026-10-01, and Pillow 13 and dnspython 2.9.0 already need 3.11. 0.5.x is the last line that supports Python 3.10; 0.6 needs Python 3.11 or later.

## 0.5.2 — 2026-10-18

Popup answers that ask first, and table writes on the right row. Three of these fixes were contributed by [@nelz0n](https://github.com/nelz0n); all were checked live on SAP GUI 8.10 and 8.00.

### Safety

- **A popup answer that may save asks first, as F11 does** ([#27](https://github.com/kts982/mcp-sap-gui/pull/27)). This covers `sap_handle_popup` (`confirm`, `press`) and `sap_press_button` on a popup button by ID. A popup is save-shaped when the button saves, or when its text is about saving and it offers more than one answer.
- **Popup questions the server cannot read.** SAP's confirmation popup (SAPLSPO1, e.g. "Save before exiting editor?", "Data may be lost") draws its question in an HTML viewer. With the Internet Explorer browser control the question is read into `texts`. With Edge (the default) scripting cannot read it: the popup reports `question_readable: false` with a `question_note`, is classified `confirmation`, and any answer but Cancel asks the user first.
- **`cancel` never presses No.** It presses a Cancel-type button (in the common logon languages) or F12. On a save prompt, No left the editor and dropped the changes.
- **`sap_send_key` checks the key before asking to save.** F11 / Save on a screen where it is not active is refused with the available keys, without a save prompt first.
- **Table-control writes after a scroll land on the right row.** Setting `VerticalScrollbar.Position` can put a different row on top than the one asked for (SE16's multiple selection: row 7 or 8 for 9, 14 for 15). `sap_modify_cell`, `sap_set_current_cell` and `sap_double_click_cell` counted visible rows from the requested position, so a write could go into another row and still report success. They now count from the row SAP put on top, and refuse with "Nothing was changed" when the row is still not visible. Row selection was not affected.

### Fixed

- **Element IDs with any SAP screen name** ([#24](https://github.com/kts982/mcp-sap-gui/pull/24)). Unnamed elements (`%#AUTOTEXT001`), `*` work areas (SE24), `=>` and `->` attributes (SLIN), field symbols `<FS>` and `+ & ' ! ? " ( ) {` were refused by the tools that take an element ID. A path segment is now a type prefix plus anything but whitespace and `/ [ ]`; the `wnd[n]` prefix and the top-level children are still checked.
- **A closed SAP window fails at once** ([#31](https://github.com/kts982/mcp-sap-gui/pull/31)). `RPC_E_DISCONNECTED` (0x80010108) and `RPC_S_SERVER_UNAVAILABLE` (0x800706BA) raise a "reconnect with `sap_connect_existing`" error at once, instead of three retries (about 2 s) and the raw COM error. `RPC_E_SERVERFAULT` (0x80010105) is SAP GUI's stale-reference error and is no longer retried: the retry re-ran the whole tool body, presses and keys included. Only `RPC_E_CALL_REJECTED` and `RPC_E_SERVERCALL_RETRYLATER` are retried, and the HRESULT names in the code are now right.
- **`sap_scroll_table_control` reports the row actually on top.** `position` is the row SAP put at the top; `requested_position` and a `note` are added when it differs from the one asked for.
- **`sap_read_table` with `start_row=0` on a table control starts at the first row.** It read from wherever the table was scrolled. Without `start_row` the read still starts at the current scroll position, e.g. after "Position...".
- **ALV toolbar: a plain button is pressed, not opened as a menu.** On SAP GUI 8.10 `PressToolbarContextButton` on a plain button runs it instead of failing, so Find opened its popup and was reported as `menu_opened`. The button's type on the toolbar now picks the press.
- **`sap_handle_popup`:** when the Enter or F12 fallback leaves the popup as it was, the result is `action: "no_effect"` instead of `confirmed`.

### Dependencies

- `mcp` 1.28.1 → 1.30.0 (through fastmcp) for [GHSA-84m7-p3x7-pcfv](https://github.com/advisories/GHSA-84m7-p3x7-pcfv) (CVE-2026-59951): the Streamable HTTP transport kept every session, so a client could exhaust the server's memory. This affects `--transport http` only. With 1.30.0, an HTTP session idle for 30 minutes is closed (a client that does not keep the event stream open then has to initialize a new session, which starts without the old session's SAP binding), and a server holds at most 10,000 sessions. stdio is unchanged.

### Release

- `SHA256SUMS.txt` no longer lists `dist/.gitignore`, which was never a release asset.

## 0.5.1 — 2026-10-15

A crash fix for SAP GUI 8.00, and screens that say more.

- **No more SAP Logon crashes on SAP GUI 8.00.** On 8.00, reading a table control or a classic list in one call could take SAP Logon down with every session in it ([issue #22](https://github.com/kts982/mcp-sap-gui/issues/22)). The cause was one property in that call: the server no longer asks for it and reads checkboxes one by one, still about 3x faster than reading cell by cell. Two environment variables, contributed by [@nelz0n](https://github.com/nelz0n) in [#23](https://github.com/kts982/mcp-sap-gui/pull/23), turn the one-call read off per call site and trace each call to a file, for any other SAP GUI build that misbehaves. Which SAP GUI releases SAP still supports, and what needs which: [SAP GUI Versions](docs/SAP_GUI_VERSIONS.md).
- **Screen info names sub-screens and the cursor.** Every action response lists the active window's sub-screens (`subscreens`, as `program:dynnr`) and the field the cursor is in (`cursor`), so an RF step, a tab page or a redisplay that only moves the cursor shows without a screenshot.
- **Table controls say where their data ends.** `sap_read_table` on a table control reports `empty_rows_from`, the first padding line: `total_rows` also counts the empty lines that pad an SM30 view.
- **Confirmed writes say so.** After you approve a save or another confirmation point, the result carries `confirmed_by_user: true`, so the agent can tell an approved write from an ungated one.

## 0.5.0 — 2026-10-06

The SAP GUI Scripting API, used to the full. This round went through the 8.10 Scripting API guide on a trial system and closed the gaps it found:

- **Much faster on large screens.** On SAP GUI 7.70 PL3 or later a whole screen, popup, list page or table control is read in one call (`GetObjectTree`): a 1,139-cell list page takes 0.2 s instead of 4.7 s. Older releases fall back to reading element by element.
- **Value help that agents get right.** F4 popups are recognised for what they are (hit list, restriction dialog, date picker, time picker) and never auto-cancelled, and a rewritten `search_help` guide covers each case. Dropdowns are set with `sap_set_field` and `sap_set_batch_fields` by key or text, and discovery shows their selected key instead of 200 padded characters.
- **Safer popups.** SAP's message popups are recognised in any logon language, and only an information message is ever auto-confirmed: a warning standing between you and a save is always left to you.
- **Keys that mean what they say.** `sap_send_key` refuses a key the screen does not accept and lists the ones it does, and `Ctrl+F` / `Ctrl+G` / `Ctrl+P` now send Find / Find next / Print (they sent Ctrl+F8 to Ctrl+F10 before).
- **More of SAP's own metadata.** Session info reports the SAP GUI release and whether the server only allows read-only scripting (write tools then refuse up front); ALV cells report their type, error or warning state, F4 help and hotspots; an ALV toolbar menu lists its items when it opens; a status message with a long text says so, and pressing the status bar opens it.
- **SAP Business Client (experimental).** Sessions of Business Client's embedded SAP GUI are discovered next to SAP Logon's. Details: [SAP Business Client](README.md#sap-business-client-experimental).

Changed behaviour when upgrading: discovery elements no longer carry `visible` (SAP GUI Scripting has no such property, so it was always true); F4 hit lists are classified `list` instead of `information`; a key the screen does not accept returns an error with `available_keys` instead of a COM error.

## 0.4.0 — 2026-09-20

Shaped by two days of real agent work on a live system. The lesson was that token cost per call hurts more than missing tools, and that a few screens could not be reached at all:

- **Lighter responses.** A table control is one discovery element instead of one per cell (an SM30 screen went from 231 elements to 3), element IDs come back in the short `wnd[0]/...` form, and batch fills and popup actions report what matters instead of echoing everything.
- **`sap_read_list`** reads a classic report list (`WRITE` output, F4 hit lists) as lines of text with list colours and paging — 43 lines instead of 492 label elements, and no screenshot needed.
- **View clusters work.** Docking containers (`wnd[0]/shellcont`) are valid IDs and discovery points to them, so the dialog structure of SM34 and most IMG activities can be navigated. A new `SM30` transaction guide covers table maintenance and view clusters end to end.
- **Your transaction policy.** Presets (`default`, `abap-dev`, `strict`) plus a policy file with glob patterns replace the hard-coded blocklist. Details: [Security Considerations](README.md#security-considerations).
- **Popups announce pre-filled values.** When an action opens a popup, its response already says what the popup is and which inputs hold values you never chose — such as a customizing request pre-filled with another project's request.
- **Screenshots to a file.** `sap_screenshot(save_path=...)` writes a full-resolution PNG for documents and never overwrites a file.

## 0.3.0 and earlier

See the [GitHub release pages](https://github.com/kts982/mcp-sap-gui/releases). 0.3.0 (2026-08-30) brought `sap_preview` approval cards and confirmation points: see [What's New](README.md#whats-new) in the README.
