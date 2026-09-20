"""
MCP prompts for common SAP GUI workflows.

Prompts provide step-by-step guidance for multi-tool SAP patterns
that agents frequently get wrong. They cost zero tool count and
reduce agent errors by prescribing the correct tool sequence.

Register on the server with: register_prompts(mcp)
"""

from typing import Literal

from fastmcp import FastMCP

# ---------------------------------------------------------------------------
# Supported workflow names
# ---------------------------------------------------------------------------

WORKFLOW_NAMES = ("search_help", "table_export", "spro_navigate")
WorkflowName = Literal["search_help", "table_export", "spro_navigate"]
WORKFLOW_TARGET_PARAMETERS = {
    "search_help": "field_id",
    "table_export": "table_id",
    "spro_navigate": "activity_name",
}
TRANSACTION_GUIDE_NAMES = ("/SCWM/MON", "SM30")
TransactionGuideName = Literal["/SCWM/MON", "SM30"]

# Alias → canonical transaction mapping (all keys lowercase)
_TRANSACTION_ALIASES: dict[str, str] = {
    "/scwm/mon": "/SCWM/MON",
    "scwm/mon": "/SCWM/MON",
    "warehouse monitor": "/SCWM/MON",
    "ewm warehouse monitor": "/SCWM/MON",
    # One guide covers both: a view cluster opens SM30-style views.
    "sm30": "SM30",
    "sm34": "SM30",
    "table maintenance": "SM30",
    "view maintenance": "SM30",
    "view cluster": "SM30",
    "customizing view": "SM30",
}


def normalize_transaction(raw: str) -> str:
    """Normalize a transaction name or human-friendly alias to its canonical form.

    Raises ``ValueError`` for unrecognised inputs.
    """
    canonical = _TRANSACTION_ALIASES.get(raw.strip().lower())
    if canonical is not None:
        return canonical
    raise ValueError(
        f"Unknown transaction {raw!r}. "
        f"Supported transactions: {', '.join(TRANSACTION_GUIDE_NAMES)}. "
        f"Also accepted: {', '.join(sorted(a for a in _TRANSACTION_ALIASES if a not in {t.lower() for t in TRANSACTION_GUIDE_NAMES}))}"
    )

# ---------------------------------------------------------------------------
# Guide text generators (shared by prompts and the workflow-guide tool)
# ---------------------------------------------------------------------------


def render_search_help_guide(field_id: str) -> str:
    """Return the step-by-step guide for F4 search help on *field_id*."""
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


def render_table_export_guide(table_id: str) -> str:
    """Return the step-by-step guide for paginated table export of *table_id*."""
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


def render_spro_navigate_guide(activity_name: str) -> str:
    """Return the step-by-step guide for navigating SPRO to *activity_name*."""
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


# ---------------------------------------------------------------------------
# Dispatch helper
# ---------------------------------------------------------------------------

_RENDERERS = {
    "search_help": render_search_help_guide,
    "table_export": render_table_export_guide,
    "spro_navigate": render_spro_navigate_guide,
}


def render_workflow_guide(workflow: WorkflowName, target: str) -> str:
    """Render the guide text for *workflow* with *target* argument.

    Raises ``ValueError`` for unknown workflow names.
    """
    renderer = _RENDERERS.get(workflow)
    if renderer is None:
        raise ValueError(
            f"Unknown workflow {workflow!r}. "
            f"Valid workflows: {', '.join(WORKFLOW_NAMES)}"
    )
    return renderer(target)


def render_scwm_mon_transaction_guide(task: str = "") -> str:
    """Return a generic, read-first guide for EWM Warehouse Monitor."""
    task_note = (
        f"Keep the user goal in mind while navigating: {task}.\n\n"
        if task.strip()
        else ""
    )
    return f"""\
Use this generic guide for **EWM Warehouse Monitor** (`/SCWM/MON`).

{task_note}## Usage mode
- Treat this as a **read-first / navigation-first** guide.
- Prefer display and inspection actions before opening follow-on documents.
- Do not hardcode field IDs, tree IDs, tab IDs, or ALV IDs. Discover them on the current system.

## Step 1 — Enter the transaction correctly
Use the SCWM form with `/n` prefix:
```
sap_execute_transaction("/n/SCWM/MON")
```

## Step 2 — Handle first-access selection screens or popups
- On first access, the monitor may ask for warehouse number, monitor profile, or similar entry criteria.
- If `active_window` shows a popup, read it first:
```
sap_get_popup_window()
```
- Discover editable fields before filling anything:
```
sap_get_screen_elements(container_id="wnd[1]/usr", changeable_only=true)
```
- Fill only the fields needed to enter the monitor, then confirm with Enter or Execute.

## Step 3 — Expect a splitter layout
- Warehouse Monitor commonly uses a split layout.
- The navigation tree is often in `shellcont[0]`.
- Results often appear in another shell container or ALV area on the right.
- If the initial element scan looks sparse, increase discovery depth:
```
sap_get_screen_elements(max_depth=4)
```

## Step 4 — Discover the tree and navigate by text
- Find the tree control instead of guessing its ID:
```
sap_get_screen_elements(type_filter="GuiTree")
```
- Search by business text first:
```
sap_search_tree_nodes("<tree_id>", "<business_text>")
```
- If nothing is found, expand parent nodes step by step:
```
sap_get_tree_node_children("<tree_id>", expand=true)
sap_get_tree_node_children("<tree_id>", node_key="<parent_key>", expand=true)
```

## Step 5 — Execute or display the selected monitor node
- Prefer the monitor's display/navigation action rather than aggressive double-clicking.
- When the tree row exposes a link/action column, use:
```
sap_click_tree_link("<tree_id>", "<node_key>", "<column>")
```
- If the tree behaves like a standard selectable tree, select the node first and then inspect the result area.

## Step 6 — Inspect the result area as ALV first
- After selecting a monitor node, discover whether the result area is a grid or table:
```
sap_get_screen_elements(type_filter="GuiGridView,GuiTableControl")
```
- Start by reading the schema:
```
sap_read_table("<table_id>", columns_only=true)
```
- Then read a small page of data:
```
sap_read_table("<table_id>", max_rows=50)
```

## Step 7 — Navigate to specific documents carefully
- Search or filter in the result set before opening anything.
- Prefer row selection and metadata reads before double-clicking into a document.
- If opening a document is necessary, verify whether the next screen is display-oriented or action-oriented.
- After every open/navigation action, check for:
  - popup dialogs
  - selection screens
  - status-bar warnings

## Common patterns to expect
- warehouse or monitor-selection step on first entry
- left-side tree navigation + right-side ALV results
- document lists that support drilldown into deliveries, warehouse tasks, handling units, or stock objects
- follow-on screens that may expose additional tabs, trees, or ALV grids

## Common pitfalls
- Forgetting the `/n` prefix for SCWM transactions
- Assuming the same monitor profile or warehouse-selection fields exist everywhere
- Guessing splitter container IDs without discovery
- Treating every double-click as safe display navigation
- Reading huge result sets before checking schema and filters
"""


def render_sm30_transaction_guide(task: str = "") -> str:
    """Return a read-first guide for view maintenance (SM30) and view clusters (SM34)."""
    task_note = (
        f"Keep the user goal in mind while navigating: {task}.\n\n"
        if task.strip()
        else ""
    )
    return f"""\
Use this guide for **table/view maintenance** (`SM30`) and **view clusters** (`SM34`). \
Most IMG (SPRO) activities open one of these two screens, so it applies there too.

{task_note}## Usage mode
- **Read first.** Open in display mode, understand the view, and only then switch to \
maintenance if the user asked for a change.
- Customizing changes are recorded in a transport request. Which request is the \
USER's decision, never yours.

## Step 1 — Open the view
```
sap_execute_transaction("SM30")     # single view
sap_execute_transaction("SM34")     # view cluster
```
- Discover the initial screen instead of guessing IDs:
```
sap_get_screen_elements(changeable_only=true)
sap_get_screen_elements(type_filter="GuiButton")
```
- Enter the view / view cluster name, then press **Display** (not Maintain) for a first look.
- A "Determine Work Area" popup may ask for a subset (e.g. a warehouse or company \
code). The response of the button press carries a `popup` digest; fill the fields it \
lists and confirm.

## Step 2 — Read the table
- The entries are a `GuiTableControl`. Discovery lists it as ONE element:
```
sap_get_screen_elements(type_filter="GuiTableControl")
sap_read_table("<table_id>", columns_only=true)
sap_read_table("<table_id>", max_rows=50)
```
- `total_rows` includes empty padding rows. The real count is in the position text \
under the table ("Entry 1 of 210").
- To jump to an entry use the **Position...** button (a popup asks for the key), \
not manual scrolling. For everything else page with `start_row`.
- An EMPTY view in display mode has no cells, so technical column names do not exist \
yet: columns are flagged `name_is_title`. Read the schema again once a row exists.

## Step 3 — View clusters: the dialog structure
- The tree on the left is a docking container BESIDE the user area. \
`sap_get_screen_elements` reports it as `docking_containers` (typically \
`wnd[0]/shellcont`), the tree itself is usually `wnd[0]/shellcont/shell`:
```
sap_read_tree("wnd[0]/shellcont/shell")
```
- Switch to a node's view with an ITEM double-click, using the tree's column name \
(from `column_names`, usually `Column1`):
```
sap_double_click_tree_item("wnd[0]/shellcont/shell", "<node_key>", "Column1")
```
- `sap_double_click_tree_node` reports success there but does NOT switch the view. \
Always verify: the window title changes to the node's name.
- A dependent node (a child in the tree) shows the entries OF the selected parent \
entry. Select the parent row first, then open the child:
```
sap_select_table_row("<parent_table_id>", 0)
sap_double_click_tree_item("wnd[0]/shellcont/shell", "<child_key>", "Column1")
```
- "No entries found that match the selection criteria" on a child view means that \
parent entry simply has none.

## Step 4 — Maintain entries (only when asked)
- Switch to change mode (Display <-> Change button), then **New Entries** is `F5` \
(in table maintenance F5 is NOT refresh).
- Take the cell IDs from the schema instead of building them by hand — it gives the \
right txt/ctxt/cmb prefix per column:
```
sap_read_table("<table_id>", columns_only=true)   # column_info[*].cell_id, {{row}} = visible row
sap_set_batch_fields({{"<cell_id row 0>": "...", "<cell_id row 1>": "..."}}, validate=true)
```
- Only VISIBLE rows can be filled. When the visible rows are used up, scroll with \
`sap_scroll_table_control` and read the schema again (row 0 is then the first visible row).
- `validate=true` presses Enter: check `message_type` (E = the row was rejected) before \
filling more rows. Offer `sap_preview` before saving a larger batch.

## Step 5 — Save
- `F11` / Save asks the user for confirmation first.
- A **"Prompt for Customizing request"** popup usually follows. SAP pre-fills the \
user's last request, which may belong to a different project. The save response shows \
it as `popup.prefilled_inputs` with a notice. Tell the user which request is pre-filled \
and ask; change the field first if it is wrong. Never confirm it blindly.
- Success is the status message "Data was saved" (`message_type` S).

## Leaving
- `F3` from a changed view asks whether to save. Read that popup; do not auto-confirm.
- `sap_execute_transaction("/n")` leaves without saving: unsaved entries are lost.

## Common pitfalls
- Looking for the dialog-structure tree under `wnd[0]/usr` (it is under `wnd[0]/shellcont`)
- Using `sap_double_click_tree_node` in a view cluster and assuming the view switched
- Opening a child node without a selected parent entry
- Treating `total_rows` as the number of entries
- Pressing F5 to "refresh" (it creates new entries)
- Accepting the pre-filled customizing request
- Drag-and-drop screens cannot be scripted at all — stop and tell the user
"""


_TRANSACTION_RENDERERS = {
    "/SCWM/MON": render_scwm_mon_transaction_guide,
    "SM30": render_sm30_transaction_guide,
}


def render_transaction_guide(transaction: TransactionGuideName, task: str = "") -> str:
    """Render a generic guide for a supported transaction."""
    renderer = _TRANSACTION_RENDERERS.get(transaction)
    if renderer is None:
        raise ValueError(
            f"Unknown transaction {transaction!r}. "
            f"Valid transactions: {', '.join(TRANSACTION_GUIDE_NAMES)}"
        )
    return renderer(task)


# ---------------------------------------------------------------------------
# MCP prompt registration
# ---------------------------------------------------------------------------


def register_prompts(mcp: FastMCP) -> None:
    """Register all SAP workflow prompts on the given server."""

    @mcp.prompt()
    def sap_search_help(field_id: str) -> str:
        """Open F4 search help on a field, browse results, and select a value.

        Use this when you need to pick a value from a dropdown or search help
        dialog (F4) for a specific field."""
        return render_search_help_guide(field_id)

    @mcp.prompt()
    def sap_table_export(table_id: str) -> str:
        """Read all rows from a large SAP table with proper pagination.

        Use this when you need to export or analyze the complete contents
        of an ALV grid or TableControl."""
        return render_table_export_guide(table_id)

    @mcp.prompt()
    def sap_spro_navigate(activity_name: str) -> str:
        """Navigate the SPRO customizing tree to find and execute an activity.

        Use this when you need to reach a specific customizing activity
        in SPRO (e.g., "Define Storage Types", "Maintain Number Ranges")."""
        return render_spro_navigate_guide(activity_name)
