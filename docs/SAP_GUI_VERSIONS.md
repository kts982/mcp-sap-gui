# SAP GUI Versions

Which SAP GUI for Windows releases SAP still supports, which ones this server
is checked on, and what changes on an older release.

## SAP's support dates

| Release | Shipped by SAP | SAP support |
|---|---|---|
| 8.10 | 16 July 2026 | until 30 June 2029 |
| 8.00 | 27 January 2023 | full support until 12 January 2027, then restricted support until 31 July 2027 |
| 7.70 | | ended 9 April 2024 |
| 7.60 and older | | ended |

Restricted support means SAP tries to solve a problem as far as it can, but
an upgrade may be the only fix. SAP supports a SAP GUI release only on a
Windows version that Microsoft still supports.

Sources:

- [SAP GUI for Windows: Platform and Availability Matrix](https://www.sap.com/documents/2016/06/586c0458-767c-0010-82c7-eda71af511fa.html) (last changed 22 July 2026)
- [SAP GUI for Windows 7.70: New Features & Lifecycle Information](https://blogs.sap.com/2021/01/27/sap-gui-for-windows-7.70-new-features-lifecycle-information/) (SAP blog, with the end-of-support update)
- SAP Note 147519, the authoritative list of maintenance dates (needs an SAP for Me login)

## Releases this server is checked on

- **8.10**: live-verified on 8.10 64-bit PL0. Every feature is used.
- **8.00**: works. The SAP Logon crash on 8.00 described in [issue #22](https://github.com/kts982/mcp-sap-gui/issues/22) is fixed in 0.5.1; the fix was checked on 8.00 PL16.
- **7.70 and older**: out of SAP support, so best effort. The server falls back where a newer scripting member is missing (see below), but these releases are not tested.

## What needs a newer release

The server reads the SAP GUI release when it attaches (`sap_gui_version` in
`sap_get_session_info`) and falls back on its own when a member is missing.
Nothing has to be configured.

| Feature | Needs | On an older release |
|---|---|---|
| Whole screens, popups, list pages and table controls read in one call (`GetObjectTree`) | 7.70 PL3 | read element by element: same result, seconds instead of a fraction of a second on large screens |
| Which browser control an HTML viewer uses (`GetBrowserControlType`) | 7.70 | taken to be Internet Explorer, the only control before 7.70 |
| ALV dropdown cell: entry count and focused index (`GetCellListBoxCount`, `GetCellListBoxCurIndex`) | 8.00 | left out of `sap_get_cell_info` |
| ALV dropdown cell: focused entry's text (`GetCellListBoxCurValue`) | 8.10 | left out of `sap_get_cell_info` |
| Status bar message has a long text (`MessageHasLongText`) | 7.60 PL2 | `has_long_text` is left out |

## Switches for a misbehaving SAP GUI build

Two environment variables, set in the MCP client's server definition
(`"env": {...}`), help when a SAP GUI build fails on the one-call read:

- `MCP_SAP_GUI_DISABLE_OBJECT_TREE` turns `GetObjectTree` off. Use `all`, or
  a comma list of the call sites `table`, `popup`, `list`, `elements` and
  `subscreens`. A disabled site reads element by element, as on a release
  before 7.70 PL3. An unknown site name stops the server at start-up.
- `MCP_SAP_GUI_TRACE_FILE` names a file that gets a `start` line before each
  `GetObjectTree` call and a `done` or `failed` line after it, written to disk
  at once. If SAP Logon exits, the last `start` with nothing after it is the
  call that was running. Match its time against the Windows Application log
  (event 1000, `saplogon.exe`).

```json
{
  "mcpServers": {
    "sap-gui": {
      "command": "uvx",
      "args": ["mcp-sap-gui"],
      "env": {
        "MCP_SAP_GUI_DISABLE_OBJECT_TREE": "table",
        "MCP_SAP_GUI_TRACE_FILE": "C:\\Temp\\sap-object-tree.log"
      }
    }
  }
}
```

## Known issues

- **SAP GUI 8.00: SAP Logon exits during a table-control or list read**
  ([issue #22](https://github.com/kts982/mcp-sap-gui/issues/22)). On 8.00
  (seen on PL14 and PL16), a `GetObjectTree` call that asks for the
  `Selected` property fails, and SAP Logon often exits with every session in
  it. 0.5.0 asked for it when reading table controls (`sap_read_table`) and
  classic lists (`sap_read_list`). 0.5.1 never asks for it and reads
  checkboxes one by one instead. On 0.5.0 with SAP GUI 8.00, avoid those two
  reads or stay on 0.4.0.
