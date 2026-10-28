# What it can do

`mcp-sap-gui` gives an AI assistant hands and eyes in SAP GUI for Windows. It works on real SAP GUI screens through SAP's Scripting API: no screenshots to guess from, no OCR, no clicking at screen coordinates.

## In SAP terms

- **Look things up:** run display transactions (`MM03`, `VA03`, `SE16`, `/SCWM/MON`), read fields, ALV grids, table controls, classic report lists (`WRITE` output, with line colours) and trees, page through tables of any length, read message long texts.
- **Fill in and change:** set fields, checkboxes, radio buttons and dropdowns, fill whole forms at once, edit table cells and text editors, use F4 value help (hit lists, restriction dialogs, date and time pickers), work in table maintenance (`SM30`, `SM34` view clusters) and IMG activities.
- **Navigate:** start transactions, press buttons and function keys, choose menu entries, switch tabs, expand and search trees such as SPRO.
- **Handle popups:** every action reports a popup it opened, recognised for what it is (a question, a warning, an F4 list, a date picker, an SAP message in any logon language). Only clearly harmless information popups are closed automatically.
- **Show and record:** take screenshots, inline or saved as a full-resolution file, and show you a preview of the screen and the values about to be written before anything is saved.

Saving always waits for your yes, and you can make other changes wait for it too: [Confirmations](CONFIRMATIONS.md).

## Things to ask

```text
What's the description of material MAT-001 in DEV?
Run ME2M for vendor 1000 and give me the open purchase orders as a table.
Read all rows of this ALV grid and sum the quantity per plant.
In SM30, show me the entries of V_T005 for country GR.
Find the IMG activity for defining countries in SPRO and open it.
Fill in this form with the values below, show me the screen first, then save.
What does the error in the status bar mean? Open its long text.
Take a screenshot of this screen and save it to C:\temp\shot.png.
Ask me before every field write from now on.
```

The assistant works out the steps: it reads the screen, finds the fields, and handles the popups. You can watch it in SAP GUI as it goes.

## How the assistant knows SAP GUI

The server teaches every client how to work with SAP GUI:

- **Instructions** that the client puts into the assistant's system prompt: discover element IDs on the screen instead of guessing them, check for a popup after each action, page through tables, find activities in SPRO without reading its whole tree, and what each function key means (F5 is "New Entries" in table maintenance, not refresh).
- **A reference guide** (`docs://sap-gui-guide`) the assistant can read when it needs more: element types and ID conventions, table types, status bar messages, step-by-step patterns for SPRO and table maintenance.
- **Workflow guides** for patterns that are easy to get wrong: F4 value help, exporting a whole table, SPRO navigation, the warehouse monitor and table maintenance.

The tool reference, for developers and the curious: [Tools](TOOLS.md).

## What it covers

Supported:

- SAP GUI for Windows, also when embedded in SAP Business Client (experimental): [Preparing SAP](SAP_SETUP.md)
- Any MCP client that can start a local server: [Clients](CLIENTS.md)
- Several clients at once, each with its own SAP session
- Everything the logged-on SAP user is authorized for, and nothing more

Not covered:

- SAP GUI for Java, SAP GUI for HTML, and Fiori or other browser apps
- Unattended or multi-user automation as a central service: the server runs on the user's own PC, for the user
