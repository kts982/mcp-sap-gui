Read ALL rows from the table on the current SAP screen, handling pagination for both ALV grids and TableControls.

The user wants: $ARGUMENTS

Steps:
1. Connect to the existing SAP session using `sap_connect_existing`
2. Get screen info with `sap_get_screen_info` to confirm we're on the right screen
3. Discover screen elements with `sap_get_screen_elements` to find the table element ID
4. Identify the table — look for elements with type containing `GuiGridView` (ALV) or `GuiTableControl`
5. Read the table with `sap_read_table` using a generous `max_rows` (e.g., 500)
6. Check the `table_type` in the response:

**If GuiGridView (ALV):**
- ALV grids handle scrolling internally — one `sap_read_table` call gets all rows up to `max_rows`
- If `total_rows` > `rows_returned`, increase `max_rows` and re-read
- Get column info with `sap_get_column_info` for proper column headers

**If GuiTableControl:**
- TableControls only show visible rows at a time
- Check `total_rows`, `first_visible_row`, and `rows_returned` in the response
- If there are more rows beyond what's visible:
  a. Collect the current page of data
  b. Use `sap_scroll_table_control` to scroll to the next page (`position` = current `first_visible_row` + `rows_returned`)
  c. Read again with `sap_read_table`
  d. Repeat until you've collected all rows or reached `total_rows`
- Watch for padding rows (all-empty values) at the end of a page — stop collecting when you see them

After collecting all data:
- Present the data as a well-formatted markdown table
- Include column headers (use titles from `sap_get_column_info` if available)
- Report total row count and any truncation
- If the user provided specific filter criteria, highlight matching rows
