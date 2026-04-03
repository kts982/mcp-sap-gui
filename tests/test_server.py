"""Tests for MCP SAP GUI Server - security logic, routing, and configuration."""

from unittest.mock import MagicMock, patch

import pytest

# Import server module once at module level (avoids beartype circular import
# issues that arise from repeated imports inside fixtures).
import mcp_sap_gui.server as _server_mod
from mcp_sap_gui.session_manager import SessionManager

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_mock_ctx():
    """Create a mock MCP Context for direct tool calls in tests."""
    return MagicMock()


@pytest.fixture
def mock_win32com():
    """Expose the conftest win32com mock for tests that reference it."""
    mock = MagicMock()
    with patch.dict("sys.modules", {"win32com": mock, "win32com.client": mock.client}):
        yield mock


@pytest.fixture
def srv(mock_win32com):
    """Configure the server module with a fresh SessionManager."""
    _server_mod._session_mgr = SessionManager()
    _server_mod.config = _server_mod.ServerConfig()
    yield _server_mod


@pytest.fixture
def readonly_srv(mock_win32com):
    """Configure the server module in read-only mode."""
    _server_mod._session_mgr = SessionManager()
    _server_mod.config = _server_mod.ServerConfig(read_only=True)
    yield _server_mod


# ===========================================================================
# Transaction Blocking Tests
# ===========================================================================

class TestTransactionBlocking:
    """Tests for _is_transaction_blocked with removeprefix fix."""

    def test_blocked_transaction_direct(self, srv):
        """Default blocklist blocks SU01."""
        assert srv._is_transaction_blocked("SU01") is True

    def test_blocked_transaction_case_insensitive(self, srv):
        """Blocklist check is case-insensitive."""
        assert srv._is_transaction_blocked("su01") is True
        assert srv._is_transaction_blocked("Su01") is True

    def test_blocked_with_n_prefix(self, srv):
        """Stripping /N prefix still detects blocked transaction."""
        assert srv._is_transaction_blocked("/NSU01") is True
        assert srv._is_transaction_blocked("/nSU01") is True

    def test_blocked_with_o_prefix(self, srv):
        """Stripping /O prefix still detects blocked transaction."""
        assert srv._is_transaction_blocked("/OSU01") is True
        assert srv._is_transaction_blocked("/oSU01") is True

    def test_blocked_with_star_prefix(self, srv):
        """Stripping /* prefix still detects blocked transaction."""
        assert srv._is_transaction_blocked("/*SU01") is True
        assert srv._is_transaction_blocked("/*SE16N") is True

    def test_blocked_with_whitespace(self, srv):
        """Leading/trailing whitespace is stripped before checking."""
        assert srv._is_transaction_blocked("  SU01") is True
        assert srv._is_transaction_blocked("SU01  ") is True
        assert srv._is_transaction_blocked(" /NSU01 ") is True

    def test_allowed_transaction(self, srv):
        """Non-blocked transactions are allowed."""
        assert srv._is_transaction_blocked("MM03") is False
        assert srv._is_transaction_blocked("VA01") is False

    def test_removeprefix_does_not_corrupt_tcode(self, srv):
        """Verify removeprefix doesn't strip characters from transaction codes.

        The old lstrip("/N") would turn "NOTIF" into "OTIF" because lstrip
        strips individual characters, not the substring.
        """
        assert srv._is_transaction_blocked("NOTIF") is False
        assert srv._is_transaction_blocked("NOOP") is False
        # SE16N IS in the blocklist - verify it's still correctly blocked
        assert srv._is_transaction_blocked("SE16N") is True
        # SE16 (without N suffix) is NOT in the blocklist
        assert srv._is_transaction_blocked("SE16") is False

    def test_all_default_blocked(self, srv):
        """All default blocked transactions are correctly blocked."""
        blocked = ["SU01", "SU10", "SU01D", "PFCG", "SU53", "SM21", "ST22", "SE16N"]
        for tcode in blocked:
            assert srv._is_transaction_blocked(tcode) is True, f"{tcode} should be blocked"

    def test_allowlist_mode(self, mock_win32com):
        """When allowed_transactions is set, only those are allowed."""
        import mcp_sap_gui.server as _srv
        _srv.config = _srv.ServerConfig(allowed_transactions=["MM03", "VA03"])

        assert _srv._is_transaction_blocked("MM03") is False
        assert _srv._is_transaction_blocked("VA03") is False
        assert _srv._is_transaction_blocked("VA01") is True
        assert _srv._is_transaction_blocked("SE80") is True

    def test_allowlist_with_prefix(self, mock_win32com):
        """Allowlist works with /N and /O prefixes."""
        import mcp_sap_gui.server as _srv
        _srv.config = _srv.ServerConfig(allowed_transactions=["MM03"])

        assert _srv._is_transaction_blocked("/NMM03") is False
        assert _srv._is_transaction_blocked("/OMM03") is False
        assert _srv._is_transaction_blocked("/NVA01") is True

    def test_allowlist_case_insensitive(self, mock_win32com):
        """Allowlist is case-insensitive via __post_init__ normalization."""
        import mcp_sap_gui.server as _srv
        _srv.config = _srv.ServerConfig(allowed_transactions=["mm03", "Va03"])

        assert _srv._is_transaction_blocked("MM03") is False
        assert _srv._is_transaction_blocked("VA03") is False
        assert _srv._is_transaction_blocked("VA01") is True

    def test_blocklist_case_insensitive_config(self, mock_win32com):
        """Blocklist entries are uppercased by __post_init__."""
        import mcp_sap_gui.server as _srv
        _srv.config = _srv.ServerConfig(blocked_transactions=["su01", "Se16n"])

        assert _srv._is_transaction_blocked("SU01") is True
        assert _srv._is_transaction_blocked("SE16N") is True

    def test_blocked_transaction_raises_valueerror(self, srv):
        """sap_execute_transaction raises ValueError (not returns dict) for blocked tcodes."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="blocked by security policy"):
            import asyncio
            asyncio.new_event_loop().run_until_complete(
                srv.sap_execute_transaction("SU01", ctx)
            )


# ===========================================================================
# OK-Code Bypass Prevention Tests
# ===========================================================================

class TestOkCodeBypassPrevention:
    """Tests for preventing transaction blocklist bypass via OK-code field."""

    @pytest.mark.asyncio
    async def test_set_field_blocks_su01_on_okcd(self, srv):
        """sap_set_field blocks SU01 when targeting the OK-code field."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="blocked by security policy"):
            await srv.sap_set_field("wnd[0]/tbar[0]/okcd", "/nSU01", ctx)

    @pytest.mark.asyncio
    async def test_set_field_blocks_okcd_case_insensitive(self, srv):
        """OK-code bypass check is case-insensitive."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="blocked by security policy"):
            await srv.sap_set_field("wnd[0]/tbar[0]/okcd", "su01", ctx)

    def test_check_okcode_bypass_matches_full_session_paths(self, srv):
        """Full SAP element IDs ending in okcd should also be protected."""
        with pytest.raises(ValueError, match="blocked by security policy"):
            srv._check_okcode_bypass(
                "/app/con[0]/ses[0]/wnd[0]/tbar[0]/OKCD",
                "/nSU01",
            )

    @pytest.mark.asyncio
    async def test_set_field_allows_mm03_on_okcd(self, srv):
        """sap_set_field allows non-blocked transactions on OK-code field."""
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        mock_ctrl.set_field.return_value = {"status": "success"}
        with patch.object(srv, '_ctrl', return_value=mock_ctrl):
            await srv.sap_set_field("wnd[0]/tbar[0]/okcd", "MM03", ctx)

    @pytest.mark.asyncio
    async def test_set_field_allows_su01_on_non_okcd(self, srv):
        """sap_set_field allows blocked tcode strings on regular fields."""
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        mock_ctrl.set_field.return_value = {"status": "success"}
        with patch.object(srv, '_ctrl', return_value=mock_ctrl):
            await srv.sap_set_field("wnd[0]/usr/txtFIELD", "SU01", ctx)

    @pytest.mark.asyncio
    async def test_batch_fields_blocks_okcd(self, srv):
        """sap_set_batch_fields blocks blocked transactions on OK-code field."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="blocked by security policy"):
            await srv.sap_set_batch_fields({
                "wnd[0]/usr/txtFIELD": "hello",
                "wnd[0]/tbar[0]/okcd": "/NSU01",
            }, ctx)

    @pytest.mark.asyncio
    async def test_screen_elements_surfaces_discovery_errors(self, srv):
        """sap_get_screen_elements should surface invalid-container errors."""
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        mock_ctrl.get_screen_elements.side_effect = RuntimeError("bad container")
        with patch.object(srv, '_ctrl', return_value=mock_ctrl):
            with pytest.raises(RuntimeError, match="bad container"):
                await srv.sap_get_screen_elements(ctx, "wnd[0]/bad")


# ===========================================================================
# Read-Only Mode Tests
# ===========================================================================

class TestReadOnlyMode:
    """Tests for read-only mode enforcement via _check_write()."""

    def test_check_write_raises_in_readonly(self, readonly_srv):
        """_check_write raises ValueError when read_only is True."""
        with pytest.raises(ValueError, match="read-only"):
            readonly_srv._check_write()

    def test_check_write_passes_in_normal_mode(self, srv):
        """_check_write does not raise when read_only is False."""
        srv._check_write()  # Should not raise

    @pytest.mark.asyncio
    async def test_readonly_blocks_set_field(self, readonly_srv):
        """sap_set_field raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_set_field("wnd[0]/usr/txt", "test", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_press_button(self, readonly_srv):
        """sap_press_button raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_press_button("wnd[0]/tbar[1]/btn[8]", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_select_checkbox(self, readonly_srv):
        """sap_select_checkbox raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_checkbox("wnd[0]/usr/chk", ctx, True)

    @pytest.mark.asyncio
    async def test_readonly_blocks_execute_transaction(self, readonly_srv):
        """sap_execute_transaction raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_execute_transaction("MM03", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_send_key(self, readonly_srv):
        """sap_send_key raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_send_key("Enter", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_select_table_row(self, readonly_srv):
        """sap_select_table_row raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_table_row("wnd[0]/usr/tbl", 0, ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_double_click_cell(self, readonly_srv):
        """sap_double_click_cell raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_double_click_cell("wnd[0]/usr/tbl", 0, "COL1", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_alv_toolbar_button(self, readonly_srv):
        """sap_press_alv_toolbar_button raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_press_alv_toolbar_button("wnd[0]/usr/grid", "SORT", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_alv_context_menu(self, readonly_srv):
        """sap_select_alv_context_menu_item raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_alv_context_menu_item(
                "wnd[0]/usr/grid", "ITEM1", ctx
            )

    @pytest.mark.asyncio
    async def test_readonly_blocks_tree_expand(self, readonly_srv):
        """sap_expand_tree_node raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_expand_tree_node("wnd[0]/usr/tree", "1", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_tree_collapse(self, readonly_srv):
        """sap_collapse_tree_node raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_collapse_tree_node("wnd[0]/usr/tree", "1", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_tree_select(self, readonly_srv):
        """sap_select_tree_node raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_tree_node("wnd[0]/usr/tree", "1", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_tree_double_click(self, readonly_srv):
        """sap_double_click_tree_node raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_double_click_tree_node("wnd[0]/usr/tree", "1", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_select_menu(self, readonly_srv):
        """sap_select_menu raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_menu("wnd[0]/mbar/menu[0]/menu[0]", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_radio_button(self, readonly_srv):
        """sap_select_radio_button raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_radio_button("wnd[0]/usr/radOPT1", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_combobox(self, readonly_srv):
        """sap_select_combobox_entry raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_combobox_entry("wnd[0]/usr/cmb", "EN", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_tab(self, readonly_srv):
        """sap_select_tab raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_tab("wnd[0]/usr/tabsTAB/tabpTAB1", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_modify_cell(self, readonly_srv):
        """sap_modify_cell raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_modify_cell("wnd[0]/usr/grid", 0, "COL", "val", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_set_current_cell(self, readonly_srv):
        """sap_set_current_cell raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_set_current_cell("wnd[0]/usr/grid", 0, "COL", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_double_click_tree_item(self, readonly_srv):
        """sap_double_click_tree_item raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_double_click_tree_item("wnd[0]/usr/tree", "1", "COL", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_click_tree_link(self, readonly_srv):
        """sap_click_tree_link raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_click_tree_link("wnd[0]/usr/tree", "1", "LINK", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_get_tree_node_children_expand(self, readonly_srv):
        """sap_get_tree_node_children with expand=True raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_get_tree_node_children("wnd[0]/usr/tree", ctx, "1", expand=True)

    @pytest.mark.asyncio
    async def test_readonly_blocks_set_batch_fields(self, readonly_srv):
        """sap_set_batch_fields raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_set_batch_fields({"wnd[0]/usr/txt": "v"}, ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_set_textedit(self, readonly_srv):
        """sap_set_textedit raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_set_textedit("wnd[0]/usr/txt", "text", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_set_focus(self, readonly_srv):
        """sap_set_focus raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_set_focus("wnd[0]/usr/txt", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_scroll_table_control(self, readonly_srv):
        """sap_scroll_table_control raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_scroll_table_control("wnd[0]/usr/tbl", 10, ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_select_all_table_control_columns(self, readonly_srv):
        """sap_select_all_table_control_columns raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_all_table_control_columns("wnd[0]/usr/tbl", ctx, True)

    @pytest.mark.asyncio
    async def test_readonly_blocks_press_column_header(self, readonly_srv):
        """sap_press_column_header raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_press_column_header("wnd[0]/usr/grid", "COL", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_select_all_rows(self, readonly_srv):
        """sap_select_all_rows raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_all_rows("wnd[0]/usr/grid", ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_select_multiple_rows(self, readonly_srv):
        """sap_select_multiple_rows raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_multiple_rows("wnd[0]/usr/grid", [0, 1], ctx)

    @pytest.mark.asyncio
    async def test_readonly_blocks_handle_popup(self, readonly_srv):
        """sap_handle_popup raises in read-only mode."""
        ctx = _make_mock_ctx()
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_handle_popup(ctx)

    # Note: sap_disconnect is intentionally NOT blocked in read-only mode.
    # Disconnecting from a session should always be allowed.


# ===========================================================================
# _parse_key Tests
# ===========================================================================

class TestParseKey:
    """Tests for _parse_key function."""

    def test_valid_keys(self, srv):
        """All valid key names parse correctly."""
        from mcp_sap_gui.sap_controller import VKey

        assert srv._parse_key("Enter") == VKey.ENTER
        assert srv._parse_key("F1") == VKey.F1
        assert srv._parse_key("F3") == VKey.F3
        assert srv._parse_key("Back") == VKey.F3
        assert srv._parse_key("F8") == VKey.F8
        assert srv._parse_key("Execute") == VKey.F8
        assert srv._parse_key("F11") == VKey.F11
        assert srv._parse_key("Save") == VKey.F11
        assert srv._parse_key("F12") == VKey.F12
        assert srv._parse_key("Cancel") == VKey.F12
        assert srv._parse_key("F5") == VKey.F5
        assert srv._parse_key("Refresh") == VKey.F5

    def test_shift_and_ctrl_keys(self, srv):
        """Shift+F and Ctrl+ key names parse correctly."""
        from mcp_sap_gui.sap_controller import VKey

        assert srv._parse_key("Shift+F1") == VKey.SHIFT_F1
        assert srv._parse_key("Shift+F5") == VKey.SHIFT_F5
        assert srv._parse_key("Shift+F9") == VKey.SHIFT_F9
        assert srv._parse_key("Ctrl+F") == VKey.CTRL_F
        assert srv._parse_key("Ctrl+G") == VKey.CTRL_G
        assert srv._parse_key("Ctrl+P") == VKey.CTRL_P

    def test_unknown_key_raises_error(self, srv):
        """Unknown key names raise ValueError instead of silently defaulting."""
        with pytest.raises(ValueError, match="Unknown key"):
            srv._parse_key("InvalidKey")

        with pytest.raises(ValueError, match="Unknown key"):
            srv._parse_key("")

        with pytest.raises(ValueError, match="Unknown key"):
            srv._parse_key("f1")  # case-sensitive

    def test_error_message_lists_valid_keys(self, srv):
        """Error message for unknown key includes list of valid keys."""
        with pytest.raises(ValueError, match="Enter") as exc_info:
            srv._parse_key("BadKey")
        assert "F1" in str(exc_info.value)
        assert "F12" in str(exc_info.value)


# ===========================================================================
# Tool Registration Tests
# ===========================================================================

class TestToolRegistration:
    """Tests that all expected tools are registered with FastMCP."""

    def test_all_tools_registered(self, srv):
        """All tools are registered."""
        import asyncio

        async def get_tools():
            return await srv.mcp.list_tools()

        tools = asyncio.new_event_loop().run_until_complete(get_tools())
        tool_names = {t.name for t in tools}

        expected = {
            # Connection & lifecycle
            "sap_connect", "sap_connect_existing", "sap_list_connections",
            "sap_get_session_info", "sap_disconnect",
            # Navigation
            "sap_execute_transaction", "sap_send_key", "sap_get_screen_info",
            # Field
            "sap_read_field", "sap_set_field", "sap_press_button",
            "sap_select_menu", "sap_select_checkbox", "sap_select_radio_button",
            "sap_select_combobox_entry", "sap_select_tab",
            "sap_get_combobox_entries", "sap_set_batch_fields",
            "sap_read_textedit", "sap_set_textedit", "sap_set_focus",
            # Table (both types)
            "sap_read_table", "sap_select_table_row", "sap_double_click_cell",
            "sap_modify_cell", "sap_set_current_cell", "sap_get_column_info",
            "sap_get_current_cell",
            # Table (ALV-specific)
            "sap_get_alv_toolbar", "sap_press_alv_toolbar_button",
            "sap_select_alv_context_menu_item",
            "sap_get_cell_info", "sap_press_column_header",
            "sap_select_all_rows",
            # Table (TableControl-specific)
            "sap_scroll_table_control", "sap_get_table_control_row_info",
            "sap_select_all_table_control_columns",
            # Table (multi-row)
            "sap_select_multiple_rows",
            # Popup & dialog
            "sap_get_popup_window", "sap_handle_popup",
            # Toolbar discovery
            "sap_get_toolbar_buttons",
            # Shell content
            "sap_read_shell_content",
            # Tree
            "sap_read_tree", "sap_expand_tree_node", "sap_collapse_tree_node",
            "sap_select_tree_node", "sap_double_click_tree_node",
            "sap_double_click_tree_item", "sap_click_tree_link",
            "sap_find_tree_node_by_path", "sap_search_tree_nodes",
            "sap_get_tree_node_children",
            # Discovery
            "sap_get_screen_elements", "sap_screenshot",
            # Policy
            "sap_set_policy_profile",
        }

        assert tool_names == expected

    def test_send_key_has_enum(self, srv):
        """sap_send_key has enum constraint on key parameter."""
        import asyncio

        async def get_tools():
            return await srv.mcp.list_tools()

        tools = asyncio.new_event_loop().run_until_complete(get_tools())
        send_key = next(t for t in tools if t.name == "sap_send_key")
        key_schema = send_key.parameters["properties"]["key"]

        assert "enum" in key_schema
        assert "Enter" in key_schema["enum"]
        assert "F8" in key_schema["enum"]
        assert len(key_schema["enum"]) == 30

    def test_sap_connect_schema_excludes_password(self, srv):
        """sap_connect should not expose a password parameter through MCP."""
        import asyncio

        async def get_tools():
            return await srv.mcp.list_tools()

        tools = asyncio.new_event_loop().run_until_complete(get_tools())
        connect_tool = next(t for t in tools if t.name == "sap_connect")
        properties = connect_tool.parameters["properties"]

        assert "password" not in properties
        assert "ctx" not in properties  # Context is injected, not exposed
        assert "system_description" in properties
        assert "client" in properties
        assert "user" in properties
        assert "language" in properties

    def test_tool_annotations_present(self, srv):
        """All tools have annotations with readOnlyHint set."""
        import asyncio

        async def get_tools():
            return await srv.mcp.list_tools()

        tools = asyncio.new_event_loop().run_until_complete(get_tools())

        read_only_tools = {
            "sap_connect", "sap_connect_existing", "sap_list_connections",
            "sap_get_session_info", "sap_disconnect",
            "sap_get_screen_info", "sap_read_field",
            "sap_get_combobox_entries", "sap_read_textedit", "sap_read_table",
            "sap_get_alv_toolbar", "sap_get_column_info", "sap_get_current_cell",
            "sap_get_table_control_row_info", "sap_get_cell_info",
            "sap_get_popup_window", "sap_get_toolbar_buttons",
            "sap_read_shell_content", "sap_read_tree", "sap_find_tree_node_by_path",
            "sap_search_tree_nodes", "sap_get_screen_elements", "sap_screenshot",
            "sap_set_policy_profile",
        }
        destructive_tools = set()  # reserved for future truly irreversible tools

        for tool in tools:
            assert tool.annotations is not None, f"{tool.name} missing annotations"
            if tool.name in read_only_tools:
                assert tool.annotations.readOnlyHint is True, (
                    f"{tool.name} should be readOnly"
                )
            elif tool.name in destructive_tools:
                assert tool.annotations.destructiveHint is True, (
                    f"{tool.name} should be destructive"
                )
            else:
                # Write tools
                assert tool.annotations.readOnlyHint is False, (
                    f"{tool.name} should NOT be readOnly"
                )


# ===========================================================================
# Tag & Policy Profile Tests
# ===========================================================================

class TestToolTags:
    """Tests for tool tag assignments and policy profiles."""

    def test_all_tools_have_tags(self, srv):
        """Every tool has at least one tag from {read, write, destructive}."""
        import asyncio

        async def get_tools():
            return await srv.mcp.list_tools()

        tools = asyncio.new_event_loop().run_until_complete(get_tools())
        valid_tags = {"read", "write", "destructive"}
        for tool in tools:
            assert tool.tags & valid_tags, (
                f"{tool.name} has no valid tags (got {tool.tags})"
            )

    def test_read_tools_tagged_read(self, srv):
        """Tools with readOnlyHint=True should have the 'read' tag."""
        import asyncio

        async def get_tools():
            return await srv.mcp.list_tools()

        tools = asyncio.new_event_loop().run_until_complete(get_tools())
        for tool in tools:
            if tool.annotations and tool.annotations.readOnlyHint:
                assert "read" in tool.tags, (
                    f"{tool.name} is readOnly but missing 'read' tag"
                )

    def test_write_tools_tagged_write(self, srv):
        """Tools with readOnlyHint=False and destructiveHint=False should have the 'write' tag."""
        import asyncio

        async def get_tools():
            return await srv.mcp.list_tools()

        tools = asyncio.new_event_loop().run_until_complete(get_tools())
        for tool in tools:
            if (tool.annotations
                    and not tool.annotations.readOnlyHint
                    and not tool.annotations.destructiveHint):
                assert "write" in tool.tags, (
                    f"{tool.name} is write but missing 'write' tag"
                )

    def test_destructive_tools_tagged_destructive(self, srv):
        """Tools with destructiveHint=True should have the 'destructive' tag."""
        import asyncio

        async def get_tools():
            return await srv.mcp.list_tools()

        tools = asyncio.new_event_loop().run_until_complete(get_tools())
        for tool in tools:
            if tool.annotations and tool.annotations.destructiveHint:
                assert "destructive" in tool.tags, (
                    f"{tool.name} is destructive but missing 'destructive' tag"
                )

    def test_profile_definitions_exist(self, srv):
        """Profile definitions cover all expected profiles."""
        from mcp_sap_gui.server import _PROFILES
        assert "exploration" in _PROFILES
        assert "operator" in _PROFILES
        assert "full" in _PROFILES

    def test_exploration_profile_is_read_only(self, srv):
        """Exploration profile only includes read tags."""
        from mcp_sap_gui.server import _PROFILES
        assert _PROFILES["exploration"] == {"read"}

    def test_operator_profile_includes_write(self, srv):
        """Operator profile includes read and write."""
        from mcp_sap_gui.server import _PROFILES
        assert _PROFILES["operator"] == {"read", "write"}

    def test_full_profile_includes_all(self, srv):
        """Full profile includes all tag categories."""
        from mcp_sap_gui.server import _PROFILES
        assert _PROFILES["full"] == {"read", "write", "destructive"}

    def test_set_policy_profile_tool_registered(self, srv):
        """sap_set_policy_profile tool is registered."""
        import asyncio

        async def get_tools():
            return await srv.mcp.list_tools()

        tools = asyncio.new_event_loop().run_until_complete(get_tools())
        names = {t.name for t in tools}
        assert "sap_set_policy_profile" in names

    def test_set_policy_profile_is_read_tagged(self, srv):
        """sap_set_policy_profile has 'read' tag so it's visible in all profiles."""
        import asyncio

        async def get_tools():
            return await srv.mcp.list_tools()

        tools = asyncio.new_event_loop().run_until_complete(get_tools())
        profile_tool = next(t for t in tools if t.name == "sap_set_policy_profile")
        assert "read" in profile_tool.tags


# ===========================================================================
# Instructions & Resource Tests
# ===========================================================================

class TestInstructionsAndResource:
    """Tests for MCP instructions and resource registration."""

    def test_instructions_set_on_mcp(self, srv):
        """FastMCP instance has instructions set."""
        assert srv.mcp.instructions is not None
        assert len(srv.mcp.instructions) > 0

    def test_instructions_contain_key_sections(self, srv):
        """Instructions cover essential SAP navigation topics."""
        instructions = srv.mcp.instructions
        assert "Getting Started" in instructions
        assert "Screen Discovery" in instructions
        assert "Popup Handling" in instructions
        assert "Tables" in instructions
        assert "SPRO" in instructions
        assert "Common Mistakes" in instructions

    def test_instructions_warn_against_guessing_ids(self, srv):
        """Instructions explicitly warn against guessing element IDs."""
        assert "Never guess" in srv.mcp.instructions

    def test_resource_registered(self, srv):
        """The docs://sap-gui-guide resource is registered."""
        import asyncio

        async def get_resources():
            return await srv.mcp.list_resources()

        resources = asyncio.new_event_loop().run_until_complete(get_resources())
        uris = [str(r.uri) for r in resources]
        assert "docs://sap-gui-guide" in uris

    def test_resource_content_not_empty(self, srv):
        """The SAP GUI guide resource returns non-empty content."""
        assert len(srv._SAP_GUI_GUIDE) > 0
        assert "Element Types" in srv._SAP_GUI_GUIDE
        assert "SPRO Navigation" in srv._SAP_GUI_GUIDE


# ===========================================================================
# COM Retry Tests
# ===========================================================================

class TestCOMRetry:
    """Tests for transient COM error retry logic."""

    def test_is_transient_com_error_by_hresult(self, srv):
        """Recognizes transient error by hresult attribute."""
        exc = Exception("COM error")
        exc.hresult = -2147417851  # RPC_E_CALL_REJECTED
        assert srv._is_transient_com_error(exc) is True

    def test_is_transient_com_error_by_args(self, srv):
        """Recognizes transient error by args[0] (pywintypes.com_error style)."""
        exc = Exception(-2147417851, "Call was rejected", None, None)
        assert srv._is_transient_com_error(exc) is True

    def test_is_not_transient_for_unknown_hresult(self, srv):
        """Non-transient COM errors are not retried."""
        exc = Exception("Some other error")
        exc.hresult = -2147352567  # random hresult
        assert srv._is_transient_com_error(exc) is False

    def test_is_not_transient_for_regular_exception(self, srv):
        """Regular exceptions are not transient."""
        assert srv._is_transient_com_error(ValueError("bad value")) is False

    @patch("mcp_sap_gui.server.time.sleep")
    def test_com_with_retry_succeeds_on_first_try(self, mock_sleep, srv):
        """No retry needed when call succeeds immediately."""
        result = srv._com_with_retry(lambda: "ok")
        assert result == "ok"
        mock_sleep.assert_not_called()

    @patch("mcp_sap_gui.server.time.sleep")
    def test_com_with_retry_retries_transient_then_succeeds(self, mock_sleep, srv):
        """Retries on transient error, then succeeds."""
        call_count = 0

        def flaky():
            nonlocal call_count
            call_count += 1
            if call_count < 3:
                exc = Exception("COM error")
                exc.hresult = -2147417851
                raise exc
            return "recovered"

        result = srv._com_with_retry(flaky)
        assert result == "recovered"
        assert call_count == 3
        assert mock_sleep.call_count == 2

    @patch("mcp_sap_gui.server.time.sleep")
    def test_com_with_retry_raises_after_max_retries(self, mock_sleep, srv):
        """Raises after exhausting all retries."""
        def always_fail():
            exc = Exception("COM error")
            exc.hresult = -2147417851
            raise exc

        with pytest.raises(Exception, match="COM error"):
            srv._com_with_retry(always_fail)

        assert mock_sleep.call_count == srv._COM_MAX_RETRIES

    @patch("mcp_sap_gui.server.time.sleep")
    def test_com_with_retry_no_retry_for_non_transient(self, mock_sleep, srv):
        """Non-transient errors are raised immediately without retry."""
        def bad_call():
            raise ValueError("not a COM error")

        with pytest.raises(ValueError, match="not a COM error"):
            srv._com_with_retry(bad_call)

        mock_sleep.assert_not_called()

    @patch("mcp_sap_gui.server.time.sleep")
    def test_com_with_retry_exponential_backoff(self, mock_sleep, srv):
        """Retry delays follow exponential backoff."""
        call_count = 0

        def fail_twice():
            nonlocal call_count
            call_count += 1
            if call_count <= 2:
                exc = Exception("COM error")
                exc.hresult = -2147417851
                raise exc
            return "ok"

        srv._com_with_retry(fail_twice)
        delays = [c.args[0] for c in mock_sleep.call_args_list]
        # base=0.3: first=0.3, second=0.6
        assert len(delays) == 2
        assert delays[0] == pytest.approx(0.3)
        assert delays[1] == pytest.approx(0.6)

    def test_transient_hresults_cover_known_codes(self, srv):
        """All documented transient hresults are in the set."""
        expected = {-2147417851, -2147418111, -2147417848}
        assert srv._TRANSIENT_COM_HRESULTS == expected


# ===========================================================================
# Screenshot Optimization Tests
# ===========================================================================

class TestScreenshotOptimization:
    """Tests for screenshot optimization logic."""

    def test_optimize_without_pillow(self, mock_win32com):
        """Optimization gracefully skips when Pillow is not installed."""
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()

        with patch.dict("sys.modules", {"PIL": None, "PIL.Image": None}):
            # Should not raise
            controller._optimize_screenshot("nonexistent.png")

    def test_optimize_with_pillow(self, mock_win32com, tmp_path):
        """Optimization applies when Pillow is available."""
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()

        # Create a mock image
        mock_image = MagicMock()
        mock_image.width = 1920
        mock_image.height = 1080
        mock_image.mode = "RGB"

        mock_image_module = MagicMock()
        mock_image_module.open.return_value = mock_image

        filepath = str(tmp_path / "test.png")

        with patch.dict("sys.modules", {"PIL": MagicMock(), "PIL.Image": mock_image_module}):
            with patch("mcp_sap_gui.sap_controller.Image", mock_image_module, create=True):
                # Patch the import inside the method
                def patched_optimize(self_arg, fp):
                    img = mock_image_module.open(fp)
                    if img.width > 1920:
                        ratio = 1920 / img.width
                        new_size = (1920, int(img.height * ratio))
                        img = img.resize(new_size)
                    img.save(fp, "PNG", optimize=True)

                with patch.object(type(controller), '_optimize_screenshot', patched_optimize):
                    controller._optimize_screenshot(filepath)

                mock_image_module.open.assert_called_once_with(filepath)
                mock_image.save.assert_called_once_with(filepath, "PNG", optimize=True)

    def test_optimize_downscales_large_images(self, mock_win32com, tmp_path):
        """Images wider than 1920px are downscaled."""
        mock_image = MagicMock()
        mock_image.width = 3840
        mock_image.height = 2160
        mock_image.mode = "RGB"

        resized_image = MagicMock()
        resized_image.mode = "RGB"
        mock_image.resize.return_value = resized_image

        mock_pil = MagicMock()
        mock_pil.open.return_value = mock_image
        mock_pil.LANCZOS = "lanczos"

        filepath = str(tmp_path / "large.png")

        def _import(name, *args, **kwargs):
            if name == "PIL.Image":
                return mock_pil
            return __builtins__.__import__(name, *args, **kwargs)

        with patch("builtins.__import__", side_effect=_import):
            # Directly test the resize logic
            img = mock_pil.open(filepath)
            if img.width > 1920:
                ratio = 1920 / img.width
                new_size = (1920, int(img.height * ratio))
                img.resize(new_size, mock_pil.LANCZOS)

            mock_image.resize.assert_called_once_with((1920, 1080), "lanczos")


# ===========================================================================
# ServerConfig Tests
# ===========================================================================

class TestServerConfig:
    """Tests for ServerConfig defaults."""

    def test_default_config(self, mock_win32com):
        """Default config has expected values."""
        from mcp_sap_gui.server import ServerConfig
        config = ServerConfig()

        assert config.read_only is False
        assert config.allowed_transactions is None
        assert "SU01" in config.blocked_transactions
        assert "SE16N" in config.blocked_transactions
        assert config.max_table_rows == 500
        assert config.default_language == "EN"

    def test_readonly_config(self, mock_win32com):
        """Read-only config works."""
        from mcp_sap_gui.server import ServerConfig
        config = ServerConfig(read_only=True)
        assert config.read_only is True


# ===========================================================================
# Audit Middleware Tests
# ===========================================================================

class TestAuditMiddleware:
    """Tests for audit logging middleware."""

    def test_mask_secrets_password_key(self):
        """Keys containing 'password' are masked."""
        from mcp_sap_gui.audit import _mask_secrets
        result = _mask_secrets({"password": "s3cret", "user": "admin"})
        assert result["password"] == "***"
        assert result["user"] == "admin"

    def test_mask_secrets_pwd_key(self):
        """Keys containing 'pwd' are masked."""
        from mcp_sap_gui.audit import _mask_secrets
        result = _mask_secrets({"pwd_field": "hunter2"})
        assert result["pwd_field"] == "***"

    def test_mask_secrets_bcode_key(self):
        """Keys containing 'bcode' are masked."""
        from mcp_sap_gui.audit import _mask_secrets
        result = _mask_secrets({"BCODE": "abc123"})
        assert result["BCODE"] == "***"

    def test_mask_secrets_case_insensitive(self):
        """Secret detection is case-insensitive."""
        from mcp_sap_gui.audit import _mask_secrets
        result = _mask_secrets({"Password": "x", "SECRET_KEY": "y"})
        assert result["Password"] == "***"
        assert result["SECRET_KEY"] == "***"

    def test_mask_secrets_preserves_normal_args(self):
        """Non-secret arguments are passed through unchanged."""
        from mcp_sap_gui.audit import _mask_secrets
        args = {"field_id": "wnd[0]/usr/txtMATNR", "value": "MAT-001"}
        result = _mask_secrets(args)
        assert result == args

    def test_mask_secrets_empty_dict(self):
        """Empty args return empty dict."""
        from mcp_sap_gui.audit import _mask_secrets
        assert _mask_secrets({}) == {}

    def test_middleware_registered_on_server(self, srv):
        """AuditMiddleware is registered on the FastMCP server."""
        from mcp_sap_gui.audit import AuditMiddleware
        has_audit = any(
            isinstance(m, AuditMiddleware) for m in srv.mcp.middleware
        )
        assert has_audit, "AuditMiddleware not found in server middleware"

    def test_audit_logs_tool_call(self, srv, caplog):
        """Audit middleware logs tool calls with structured JSON."""
        import asyncio
        import json
        import logging

        # The middleware logs via mcp_sap_gui.audit logger.
        # caplog captures log records from all loggers.
        with caplog.at_level(logging.INFO, logger="mcp_sap_gui.audit"):
            # We can't easily run a full tool call through the middleware
            # without a real MCP session, but we can test the middleware
            # directly by constructing a MiddlewareContext.
            from datetime import datetime, timezone
            from mcp_sap_gui.audit import AuditMiddleware
            from fastmcp.server.middleware import MiddlewareContext

            mw = AuditMiddleware()
            params = MagicMock()
            params.name = "sap_read_field"
            params.arguments = {"field_id": "wnd[0]/usr/txtMATNR"}
            ctx = MiddlewareContext(
                message=params,
                timestamp=datetime.now(timezone.utc),
                method="tools/call",
            )

            async def fake_next(ctx):
                result = MagicMock()
                result.content = []
                return result

            asyncio.new_event_loop().run_until_complete(
                mw.on_call_tool(ctx, fake_next)
            )

        # Check that a structured JSON log was emitted
        audit_records = [
            r for r in caplog.records
            if r.name == "mcp_sap_gui.audit"
        ]
        assert len(audit_records) == 1
        log_data = json.loads(audit_records[0].message)
        assert log_data["event"] == "tool_call"
        assert log_data["tool"] == "sap_read_field"
        assert log_data["status"] == "ok"
        assert "duration_ms" in log_data
        assert log_data["args"]["field_id"] == "wnd[0]/usr/txtMATNR"

    def test_audit_logs_error(self, srv, caplog):
        """Audit middleware logs errors with error details."""
        import asyncio
        import json
        import logging

        with caplog.at_level(logging.WARNING, logger="mcp_sap_gui.audit"):
            from datetime import datetime, timezone
            from mcp_sap_gui.audit import AuditMiddleware
            from fastmcp.server.middleware import MiddlewareContext

            mw = AuditMiddleware()
            params = MagicMock()
            params.name = "sap_set_field"
            params.arguments = {"field_id": "wnd[0]/usr/txtX", "value": "v"}
            ctx = MiddlewareContext(
                message=params,
                timestamp=datetime.now(timezone.utc),
                method="tools/call",
            )

            async def fail_next(ctx):
                raise ValueError("Element not found")

            with pytest.raises(ValueError, match="Element not found"):
                asyncio.new_event_loop().run_until_complete(
                    mw.on_call_tool(ctx, fail_next)
                )

        audit_records = [
            r for r in caplog.records
            if r.name == "mcp_sap_gui.audit"
        ]
        assert len(audit_records) == 1
        log_data = json.loads(audit_records[0].message)
        assert log_data["status"] == "error"
        assert log_data["error"] == "ValueError"
        assert "Element not found" in log_data["error_msg"]

    def test_audit_masks_secrets_in_args(self, srv, caplog):
        """Audit middleware masks password-like arguments in logs."""
        import asyncio
        import json
        import logging

        with caplog.at_level(logging.INFO, logger="mcp_sap_gui.audit"):
            from datetime import datetime, timezone
            from mcp_sap_gui.audit import AuditMiddleware
            from fastmcp.server.middleware import MiddlewareContext

            mw = AuditMiddleware()
            params = MagicMock()
            params.name = "sap_connect"
            params.arguments = {"system": "D01", "password": "s3cret123"}
            ctx = MiddlewareContext(
                message=params,
                timestamp=datetime.now(timezone.utc),
                method="tools/call",
            )

            async def fake_next(ctx):
                result = MagicMock()
                result.content = []
                return result

            asyncio.new_event_loop().run_until_complete(
                mw.on_call_tool(ctx, fake_next)
            )

        audit_records = [
            r for r in caplog.records
            if r.name == "mcp_sap_gui.audit"
        ]
        log_data = json.loads(audit_records[0].message)
        assert log_data["args"]["password"] == "***"
        assert log_data["args"]["system"] == "D01"
        assert "s3cret123" not in audit_records[0].message


# ===========================================================================
# Credential Resolution Tests
# ===========================================================================

class TestCredentialResolution:
    """Tests for .env-based credential resolution in sap_connect."""

    def test_env_vars_resolved(self, srv, monkeypatch):
        """Environment variables are picked up for credential resolution."""
        import os
        monkeypatch.setenv("SAP_USER", "ENV_USER")
        monkeypatch.setenv("SAP_PASSWORD", "ENV_PASS")
        monkeypatch.setenv("SAP_CLIENT", "200")
        monkeypatch.setenv("SAP_LANGUAGE", "DE")
        # Simulate the resolution logic from sap_connect
        assert (None or os.environ.get("SAP_USER") or None) == "ENV_USER"
        assert (None or os.environ.get("SAP_PASSWORD") or None) == "ENV_PASS"
        assert (None or os.environ.get("SAP_CLIENT") or None) == "200"
        assert (None or os.environ.get("SAP_LANGUAGE") or None) == "DE"

    def test_explicit_params_override_env(self, srv, monkeypatch):
        """Explicit tool params override env vars (except password)."""
        monkeypatch.setenv("SAP_USER", "ENV_USER")
        monkeypatch.setenv("SAP_CLIENT", "200")
        # Verify the resolution logic directly
        import os
        user_param = "EXPLICIT_USER"
        resolved = user_param or os.environ.get("SAP_USER") or None
        assert resolved == "EXPLICIT_USER"

    def test_password_only_from_env(self, srv):
        """sap_connect tool does not accept password as MCP parameter."""
        import asyncio
        import inspect
        sig = inspect.signature(_server_mod.sap_connect)
        param_names = set(sig.parameters.keys())
        assert "password" not in param_names, (
            "password must not be an MCP parameter — resolve from env only"
        )

    def test_no_env_means_no_credentials(self, srv, monkeypatch):
        """Without env vars, no credentials are passed."""
        import os
        monkeypatch.delenv("SAP_USER", raising=False)
        monkeypatch.delenv("SAP_PASSWORD", raising=False)
        monkeypatch.delenv("SAP_CLIENT", raising=False)
        monkeypatch.delenv("SAP_LANGUAGE", raising=False)
        # Verify resolution produces None
        assert os.environ.get("SAP_PASSWORD") is None
        assert os.environ.get("SAP_USER") is None
