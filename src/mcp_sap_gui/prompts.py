"""
MCP prompts for common SAP GUI workflows.

Prompts provide step-by-step guidance for multi-tool SAP patterns
that agents frequently get wrong. They cost zero tool count and
reduce agent errors by prescribing the correct tool sequence.

Register on the server with: register_prompts(mcp)
"""

from fastmcp import FastMCP


def register_prompts(mcp: FastMCP) -> None:
    """Register all SAP workflow prompts on the given server."""

    @mcp.prompt()
    def sap_search_help(field_id: str) -> str:
        """Open F4 search help on a field, browse results, and select a value.

        Use this when you need to pick a value from a dropdown or search help
        dialog (F4) for a specific field."""
        return f"""\
Follow these steps exactly to use F4 search help on field `{field_id}`:

## Step 1 — Set focus on the target field
```
sap_set_focus("{field_id}")
```

## Step 2 — Open search help
```
sap_send_key("F4")
```

## Step 3 — Verify the popup opened
Check that `active_window` is `wnd[1]` or higher in the response.
If it is still `wnd[0]`, the field may not support F4. Read the status bar:
```
sap_read_field("wnd[0]/sbar")
```

## Step 4 — Discover the popup structure
```
sap_get_screen_elements(container_id="wnd[1]/usr", type_filter="GuiGridView,GuiTableControl")
```
The popup typically contains a table with selectable values.
If there are filter fields at the top, fill them first with `sap_set_field`
and press Enter to narrow results.

## Step 5 — Read the results table
```
sap_read_table("<table_id_from_step_4>", columns_only=true)
```
Then read the data columns you need:
```
sap_read_table("<table_id>", columns="<relevant_columns>")
```

## Step 6 — Select a row
```
sap_select_table_row("<table_id>", <row_index>)
```
Or double-click to select and confirm in one step:
```
sap_double_click_cell("<table_id>", <row_index>, "<column_name>")
```

## Step 7 — Verify selection
The popup should close (active_window back to `wnd[0]`).
Confirm the field was filled:
```
sap_read_field("{field_id}")
```
If the popup is still open, press Enter to confirm:
```
sap_send_key("Enter")
```

## Common pitfalls
- **Forgetting sap_set_focus first**: F4 acts on the focused field. Without focus, it may open the wrong search help or do nothing.
- **Using sap_send_key("F4") without checking the popup**: Always verify `active_window` changed.
- **Trying to type in the field instead**: Some fields require F4 selection and reject typed values.
"""

    @mcp.prompt()
    def sap_table_export(table_id: str) -> str:
        """Read all rows from a large SAP table with proper pagination.

        Use this when you need to export or analyze the complete contents
        of an ALV grid or TableControl."""
        return f"""\
Follow these steps to read all data from table `{table_id}`:

## Step 1 — Discover the table schema
```
sap_read_table("{table_id}", columns_only=true)
```
This returns column names, titles, and total row count without reading data.
Note the `total_rows` value — you will need it for pagination.
Note the `table_type` — it affects pagination behavior.

## Step 2 — Identify useful columns
Review the column names and titles. Select only the columns you actually need
to minimize response size:
```
sap_read_table("{table_id}", columns="COL_A,COL_B,COL_C", max_rows=100)
```

## Step 3 — Paginate through all rows
Use `start_row` to read in batches. Continue until you have read all rows:
```
sap_read_table("{table_id}", columns="COL_A,COL_B,COL_C", start_row=0, max_rows=100)
sap_read_table("{table_id}", columns="COL_A,COL_B,COL_C", start_row=100, max_rows=100)
sap_read_table("{table_id}", columns="COL_A,COL_B,COL_C", start_row=200, max_rows=100)
```
Stop when `start_row >= total_rows` from step 1, or when `rows_returned` is 0.

## Step 4 — Verify completeness
Compare the total rows read against `total_rows` from step 1.
If they don't match, you may have missed rows — check your pagination math.

## Tips
- **Start with `columns_only=true`**: Don't read data until you know the schema.
- **Select specific columns**: Wide tables waste tokens. Only fetch what you need.
- **Batch size**: 100 rows per call is a good default. For very wide tables, use smaller batches.
- **GuiTableControl**: `total_rows` may report the visible capacity rather than the actual data rows. Rows with all-empty values indicate you've reached the end of real data.
- **Position button**: For SM30-style table maintenance, look for a "Position..." button in the toolbar (`sap_get_toolbar_buttons`) to jump to specific entries instead of paginating.
"""

    @mcp.prompt()
    def sap_spro_navigate(activity_name: str) -> str:
        """Navigate the SPRO customizing tree to find and execute an activity.

        Use this when you need to reach a specific customizing activity
        in SPRO (e.g., "Define Storage Types", "Maintain Number Ranges")."""
        return f"""\
Follow these steps to find and execute "{activity_name}" in SPRO:

## Step 1 — Open SPRO
```
sap_execute_transaction("SPRO")
```
Then click the "SAP Reference IMG" button. Find it with:
```
sap_get_toolbar_buttons()
```
Look for a button with tooltip containing "SAP Reference IMG" and press it:
```
sap_press_button("<button_id>")
```

## Step 2 — Find the tree control
```
sap_get_screen_elements(type_filter="GuiTree")
```
Note the tree ID (typically something like `wnd[0]/usr/shell/shellcont[0]/shell`).

## Step 3 — Search for the activity
```
sap_search_tree_nodes("<tree_id>", "{activity_name}")
```

**IMPORTANT**: `sap_search_tree_nodes` only searches **already-loaded** nodes.
SPRO starts with most nodes collapsed. If you get no results:

### Expand the path step by step
Get top-level nodes:
```
sap_get_tree_node_children("<tree_id>", expand=true)
```
Then expand relevant parent nodes one level at a time:
```
sap_get_tree_node_children("<tree_id>", node_key="<parent_key>", expand=true)
```
After expanding, search again:
```
sap_search_tree_nodes("<tree_id>", "{activity_name}")
```

## Step 4 — Execute the activity
Once you have the node key, use `sap_click_tree_link` with column `"2"`:
```
sap_click_tree_link("<tree_id>", "<node_key>", "2")
```

**CRITICAL**: Do NOT use `sap_double_click_tree_node` — in SPRO that opens
documentation (hypertext), not the activity. Always use `sap_click_tree_link`
on column `"2"` (the execute icon).

## Step 5 — Handle the result
After clicking, check for popups:
- A selection screen popup (`wnd[1]`) often appears asking for organizational
  parameters (company code, warehouse number, etc.)
- Fill the required fields and press Execute (F8) or Enter
- The customizing table/screen then appears on `wnd[0]`

Check `active_window` in the response. If it shows `wnd[1]`:
```
sap_get_popup_window()
```

## Common pitfalls
- **Using `sap_read_tree` on SPRO**: SPRO has 1000+ nodes. `read_tree` is far too slow. Always use `search_tree_nodes` + `get_tree_node_children`.
- **Double-clicking nodes**: Opens documentation, not activities. Use `click_tree_link`.
- **Search not finding nodes**: Nodes must be expanded/loaded first. Expand parent nodes, then search again.
- **Forgetting the selection popup**: Most SPRO activities show a parameter popup before the actual screen. Always check for `wnd[1]` after clicking.
"""
