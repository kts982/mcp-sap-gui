Give a quick status report on the current SAP session.

Steps:
1. List all open SAP connections with `sap_list_connections`
2. If there's an active session, connect with `sap_connect_existing`
3. Get session info with `sap_get_session_info` (system, client, user, transaction)
4. Get screen info with `sap_get_screen_info` — check `active_window` to see if a popup is open
5. If `active_window` is `wnd[1]` or higher, use `sap_get_popup_window` for full popup details
6. Take a screenshot with `sap_screenshot` so the user can see the current state

Present a concise summary:
- **System**: system name, client, language
- **User**: logged-in user
- **Current Transaction**: t-code and screen number
- **Screen**: title and description
- **Status Message**: any current status bar message (and type: Success/Error/Warning/Info)
- **Popup**: whether a popup is open and what it says
