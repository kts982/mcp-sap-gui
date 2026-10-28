# Preparing SAP

The server drives SAP GUI for Windows through the SAP GUI Scripting API. Scripting has to be allowed in two places: in SAP GUI on the user's PC, and on the SAP system. The second one needs your Basis team.

## Enabling SAP GUI Scripting

**In SAP GUI** (each user, on their PC):

1. SAP GUI → Options → Accessibility & Scripting → Scripting.
2. Tick "Enable scripting".
3. Optionally untick "Notify when a script attaches to SAP GUI" and "Notify when a script opens a connection", so SAP GUI does not ask each time the server attaches.

**On the SAP system** (Basis):

- Transaction `RZ11`, parameter `sapgui/user_scripting` = `TRUE`.
- A change in `RZ11` is not persistent: it reverts at the next restart of the application server, which suits a time-boxed trial on a development system. For a lasting setting, use the instance profile (`RZ10`).

`sap_list_connections` reports a system that does not allow scripting with `scripting_disabled: true`.

## Where scripting is restricted

Many organisations switch scripting off as a hardening default. It does not have to be all or nothing: SAP ships server-side parameters that let Basis enable it narrowly, for example for named users on a development system.

| Profile parameter (`RZ11` / `RZ10`) | Effect |
|---|---|
| `sapgui/user_scripting = TRUE` | Master switch; required for any scripting |
| `sapgui/user_scripting_per_user = TRUE` | Scripting works only for users holding authorization object `S_SCR` (class `BC_A`, activity 16); all other users stay blocked (SAP Note 983990) |
| `sapgui/user_scripting_set_readonly = TRUE` | Scripts may read screen state but cannot send anything that changes server state (SAP Note 692245) |
| `sapgui/user_scripting_force_notification = TRUE` | Users always see a notification when a script attaches; it cannot be switched off in the local SAP GUI options (SAP Note 3591984) |
| `sapgui/user_scripting_disable_recording = TRUE` | Blocks recording of new scripts; playback still works |

Useful facts when proposing this to a Basis or security team:

- Per-user and read-only modes can be combined since SAP GUI 7.40 PL17 / 7.50 PL4: the full API for `S_SCR` holders, read-only for everyone else (SAP Note 2565390).
- On a read-only session the server reports `scripting_read_only: true` in the session info and refuses every write tool up front with the reason, instead of letting each action fail inside SAP GUI. All read tools keep working.
- Scripted actions run under the SAP user's own authorizations and appear in the SAP logs as ordinary user activity. The user's authorizations are the real security boundary: see the [Security guide](SECURITY_GUIDE.md).
- Reference: the [SAP GUI Scripting Security Guide](https://help.sap.com/doc/97d2d0bc2ed248a4a85a0bec608704f8/800.13/en-US/sap_gui_scripting_sec_guide.pdf) (help.sap.com).

**Why the Scripting API is needed at all:** SAP GUI for Windows draws its screens on a canvas that exposes no usable structure to Windows UI Automation or other accessibility interfaces. The Scripting API is the only structured way to read and drive those screens. Commercial RPA products have the same dependency and fall back to screenshots and OCR when scripting is off; this project deliberately does not, because that approach is imprecise and brittle.

## SAP GUI versions

SAP GUI for Windows 8.10 is verified live and 8.00 works; 7.70 and older are out of SAP support and run best effort. On 7.70 Patch Level 3 or later the server reads a whole screen, popup, list page or table control in one call; older releases read element by element, which takes seconds instead of a fraction of a second on a large screen. SAP's support dates, what needs which release, and the switches for a misbehaving build: [SAP GUI versions](SAP_GUI_VERSIONS.md).

## SAP Business Client (experimental)

Classic transactions in SAP Business Client run in an embedded SAP GUI for Windows, and the server can attach to them.

- The server-side scripting parameter is still required. `sapgui/nwbc_scripting` gives read-only scripting (the same effect as `sapgui/user_scripting` plus `sapgui/user_scripting_set_readonly`): write tools refuse, and grids cannot be scrolled, so a table read stops at the rows on screen.
- `sap_list_connections` reports Business Client sessions with `host: "sapguiserver"`. The tabs of one Business Client window are sessions of one connection; a second system is a second connection.
- Business Client connections have no SAP Logon description; the list shows the system and client instead, for example `DEV (100)`.
- Pressing Back repeatedly closes the tab instead of landing on the SAP Easy Access menu.
- A Fiori launchpad connection may show a transaction as SAP GUI for HTML. Such a transaction is invisible to scripting.

## Licensing

Check your SAP licensing agreement for automated access and scripting, indirect access, and named versus concurrent users.
