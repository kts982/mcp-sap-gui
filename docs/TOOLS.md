# Tool Catalog

`mcp-sap-gui` currently exposes **55 MCP tools**.

Two practical rules:

- Start with discovery instead of guessing IDs.
- Start table work with `sap_read_table`, then branch into ALV- or TableControl-specific tools as needed.

## Policy Profiles

Every tool is tagged `read` or `write`. Three profiles control which tools are visible:

| Profile | Tags | What the agent can do |
|---|---|---|
| `exploration` | `read` | Observe only: inspect screens, read tables, take screenshots. Cannot interact or navigate. |
| `operator` | `read`, `write` | Normal SAP work: navigate transactions, fill fields, press buttons. Transaction blocklist still applies. |
| `full` | `read`, `write` | All tools. Default. |

Set the default profile at startup with `--profile operator`, or switch per-session with `sap_set_policy_profile`.

## Connection

Preferred usage: use `sap_connect_existing` when the user is already logged in to SAP. Use `sap_connect` mainly for SSO flows or to open the SAP login screen before the user completes manual login.

| Tool | Description |
|------|-------------|
| `sap_connect` | Connect to an SAP system by SAP Logon entry name. Credentials resolved from `.env` file — passwords never appear in MCP |
| `sap_connect_existing` | Attach to an already open SAP session |
| `sap_list_connections` | List all currently open SAP connections and sessions |
| `sap_get_session_info` | Get current session metadata like system, client, user, transaction, and screen |
| `sap_disconnect` | Disconnect from the current SAP session (detaches attached sessions, closes owned sessions) |
| `sap_set_policy_profile` | Switch the active policy profile for this session (exploration, operator, full) |

## Navigation

| Tool | Description |
|------|-------------|
| `sap_execute_transaction` | Execute a transaction code such as `MM03`, `VA01`, or `SE80` |
| `sap_send_key` | Send SAP keys such as `Enter`, function keys, `Back`, or `Save` |
| `sap_get_screen_info` | Read current screen info including transaction, program, screen number, title, status, and active window |

## Fields And UI Elements

| Tool | Description |
|------|-------------|
| `sap_read_field` | Read a field value with metadata |
| `sap_set_field` | Set a field value |
| `sap_press_button` | Press a button |
| `sap_select_menu` | Select a menu item or submenu |
| `sap_select_checkbox` | Select or clear a checkbox |
| `sap_select_radio_button` | Select a radio button |
| `sap_select_combobox_entry` | Select a combobox entry by key or visible value |
| `sap_select_tab` | Select a tab strip tab |
| `sap_get_combobox_entries` | List combobox entries |
| `sap_set_batch_fields` | Set multiple fields in one call |
| `sap_read_textedit` | Read a multiline text editor |
| `sap_set_textedit` | Set a multiline text editor |
| `sap_set_focus` | Set focus to a screen element |

## Tables And Grids

`sap_read_table` auto-detects whether the control is:

- `GuiGridView` for ALV-style grids
- `GuiTableControl` for classic SAP table controls

### Shared Table Tools

| Tool | Description |
|------|-------------|
| `sap_read_table` | Read rows and columns from a table or grid |
| `sap_select_table_row` | Select a row |
| `sap_double_click_cell` | Double-click a cell |
| `sap_modify_cell` | Modify an editable cell |
| `sap_set_current_cell` | Set the focused cell |
| `sap_get_column_info` | Get column names, titles, and metadata |
| `sap_get_current_cell` | Get the current focused cell |
| `sap_select_multiple_rows` | Select multiple rows |

### ALV Grid Tools

| Tool | Description |
|------|-------------|
| `sap_get_alv_toolbar` | List ALV toolbar buttons |
| `sap_press_alv_toolbar_button` | Press an ALV toolbar button |
| `sap_select_alv_context_menu_item` | Select an ALV context menu item |
| `sap_get_cell_info` | Read detailed ALV cell metadata |
| `sap_press_column_header` | Click a column header, typically to sort |
| `sap_select_all_rows` | Select all ALV rows |

### TableControl Tools

| Tool | Description |
|------|-------------|
| `sap_scroll_table_control` | Scroll a classic table control to a row position |
| `sap_get_table_control_row_info` | Read row metadata from a classic table control |
| `sap_select_all_table_control_columns` | Select or clear all column headers in a classic table control |

## Popup, Toolbar, And Shell

| Tool | Description |
|------|-------------|
| `sap_get_popup_window` | Read popup title, text, and buttons |
| `sap_handle_popup` | Read and act on popups in one call (confirm, cancel, or press a specific button) |
| `sap_get_toolbar_buttons` | List standard SAP toolbar buttons |
| `sap_read_shell_content` | Read content from shell-based controls such as HTML viewers |

## Trees

| Tool | Description |
|------|-------------|
| `sap_read_tree` | Read nodes, hierarchy, and values from a tree |
| `sap_expand_tree_node` | Expand a folder node |
| `sap_collapse_tree_node` | Collapse a folder node |
| `sap_select_tree_node` | Select a tree node |
| `sap_double_click_tree_node` | Double-click a tree node |
| `sap_double_click_tree_item` | Double-click a specific item in a tree row |
| `sap_click_tree_link` | Click a hyperlink inside a tree row |
| `sap_find_tree_node_by_path` | Resolve a node key from a path |
| `sap_search_tree_nodes` | Search tree nodes by text |
| `sap_get_tree_node_children` | Get direct children of a node, optionally expanding first |

## Discovery

| Tool | Description |
|------|-------------|
| `sap_get_screen_elements` | Enumerate screen elements, optionally by container or filter |
| `sap_screenshot` | Capture a screenshot of the active SAP window |

## Recommended Usage Patterns

### New Or Unfamiliar Screen

1. `sap_get_screen_info`
2. `sap_get_screen_elements`
3. Then read or set fields based on discovered IDs

### Popup Handling

1. Check `active_window` in tool responses
2. If it is not `wnd[0]`, call `sap_get_popup_window`

### Table Handling

1. Start with `sap_read_table`
2. Check `table_type`
3. Use ALV-specific or TableControl-specific tools if needed

### SPRO Or Tree Navigation

1. `sap_read_tree`
2. `sap_search_tree_nodes`
3. `sap_expand_tree_node` or `sap_click_tree_link`

## Related

- [Client Setup Guide](CLIENTS.md)
- [Overview](OVERVIEW.md)
