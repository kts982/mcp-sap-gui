Explore the current SAP screen and report what you find.

Steps:
1. Connect to the existing SAP session using `sap_connect_existing`
2. Get session info with `sap_get_session_info` to know where we are
3. Get screen info with `sap_get_screen_info` — check `active_window` in the response:
   - If `wnd[0]`: normal screen, continue with step 4
   - If `wnd[1]` or higher: a popup is open — use `sap_get_popup_window` for full details and report it first
4. Discover all screen elements with `sap_get_screen_elements` on the active window's user area
5. Get available toolbar buttons with `sap_get_toolbar_buttons`
7. If the screen contains a table or grid (look for element types containing `GuiGridView` or `GuiTableControl`):
   - Read the table data with `sap_read_table` (limit to 20 rows)
   - Note the `table_type` in the response
   - If ALV: also get the ALV toolbar buttons with `sap_get_alv_toolbar`
   - If TableControl: note the scroll position and total rows
8. If the screen contains a tree control, read it with `sap_read_tree` (limit to 50 nodes)

Present a clear summary:
- **Transaction**: current transaction code and screen number
- **Screen Title**: what this screen is for
- **Status**: any status bar messages
- **Input Fields**: list changeable fields with their current values and labels
- **Tables/Grids**: column names, row counts, and first few rows of data
- **Available Actions**: toolbar buttons, function key shortcuts
- **Suggested Next Steps**: what the user might want to do on this screen
