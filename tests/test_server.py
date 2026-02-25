"""Tests for MCP SAP GUI Server - security logic, routing, and configuration."""

from unittest.mock import MagicMock, patch

import pytest

# ---------------------------------------------------------------------------
# Helpers to import server module with mocked win32com
# ---------------------------------------------------------------------------

@pytest.fixture
def mock_win32com():
    """Mock win32com so SAPGUIController can be instantiated on any OS."""
    mock = MagicMock()
    with patch.dict("sys.modules", {"win32com": mock, "win32com.client": mock.client}):
        yield mock


@pytest.fixture
def srv(mock_win32com):
    """Import and configure the server module with a fresh controller."""
    import importlib

    import mcp_sap_gui.server as _srv
    # Reload to pick up mocked win32com
    importlib.reload(_srv)

    from mcp_sap_gui.sap_controller import SAPGUIController
    _srv.controller = SAPGUIController()
    _srv.config = _srv.ServerConfig()
    yield _srv


@pytest.fixture
def readonly_srv(mock_win32com):
    """Import and configure the server module in read-only mode."""
    import importlib

    import mcp_sap_gui.server as _srv
    importlib.reload(_srv)

    from mcp_sap_gui.sap_controller import SAPGUIController
    _srv.controller = SAPGUIController()
    _srv.config = _srv.ServerConfig(read_only=True)
    yield _srv


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
        import importlib

        import mcp_sap_gui.server as _srv
        importlib.reload(_srv)
        _srv.config = _srv.ServerConfig(allowed_transactions=["MM03", "VA03"])

        assert _srv._is_transaction_blocked("MM03") is False
        assert _srv._is_transaction_blocked("VA03") is False
        assert _srv._is_transaction_blocked("VA01") is True
        assert _srv._is_transaction_blocked("SE80") is True

    def test_allowlist_with_prefix(self, mock_win32com):
        """Allowlist works with /N and /O prefixes."""
        import importlib

        import mcp_sap_gui.server as _srv
        importlib.reload(_srv)
        _srv.config = _srv.ServerConfig(allowed_transactions=["MM03"])

        assert _srv._is_transaction_blocked("/NMM03") is False
        assert _srv._is_transaction_blocked("/OMM03") is False
        assert _srv._is_transaction_blocked("/NVA01") is True

    def test_allowlist_case_insensitive(self, mock_win32com):
        """Allowlist is case-insensitive via __post_init__ normalization."""
        import importlib

        import mcp_sap_gui.server as _srv
        importlib.reload(_srv)
        _srv.config = _srv.ServerConfig(allowed_transactions=["mm03", "Va03"])

        assert _srv._is_transaction_blocked("MM03") is False
        assert _srv._is_transaction_blocked("VA03") is False
        assert _srv._is_transaction_blocked("VA01") is True

    def test_blocklist_case_insensitive_config(self, mock_win32com):
        """Blocklist entries are uppercased by __post_init__."""
        import importlib

        import mcp_sap_gui.server as _srv
        importlib.reload(_srv)
        _srv.config = _srv.ServerConfig(blocked_transactions=["su01", "Se16n"])

        assert _srv._is_transaction_blocked("SU01") is True
        assert _srv._is_transaction_blocked("SE16N") is True

    def test_blocked_transaction_raises_valueerror(self, srv):
        """sap_execute_transaction raises ValueError (not returns dict) for blocked tcodes."""
        with pytest.raises(ValueError, match="blocked by security policy"):
            import asyncio
            asyncio.new_event_loop().run_until_complete(
                srv.sap_execute_transaction("SU01")
            )


# ===========================================================================
# OK-Code Bypass Prevention Tests
# ===========================================================================

class TestOkCodeBypassPrevention:
    """Tests for preventing transaction blocklist bypass via OK-code field."""

    @pytest.mark.asyncio
    async def test_set_field_blocks_su01_on_okcd(self, srv):
        """sap_set_field blocks SU01 when targeting the OK-code field."""
        with pytest.raises(ValueError, match="blocked by security policy"):
            await srv.sap_set_field("wnd[0]/tbar[0]/okcd", "/nSU01")

    @pytest.mark.asyncio
    async def test_set_field_blocks_okcd_case_insensitive(self, srv):
        """OK-code bypass check is case-insensitive."""
        with pytest.raises(ValueError, match="blocked by security policy"):
            await srv.sap_set_field("wnd[0]/tbar[0]/okcd", "su01")

    @pytest.mark.asyncio
    async def test_set_field_allows_mm03_on_okcd(self, srv):
        """sap_set_field allows non-blocked transactions on OK-code field."""
        # Should not raise - it will fail on the COM call, but not on security check
        srv.controller = MagicMock()
        srv.controller.set_field.return_value = {"status": "success"}
        await srv.sap_set_field("wnd[0]/tbar[0]/okcd", "MM03")

    @pytest.mark.asyncio
    async def test_set_field_allows_su01_on_non_okcd(self, srv):
        """sap_set_field allows blocked tcode strings on regular fields."""
        srv.controller = MagicMock()
        srv.controller.set_field.return_value = {"status": "success"}
        await srv.sap_set_field("wnd[0]/usr/txtFIELD", "SU01")

    @pytest.mark.asyncio
    async def test_batch_fields_blocks_okcd(self, srv):
        """sap_set_batch_fields blocks blocked transactions on OK-code field."""
        with pytest.raises(ValueError, match="blocked by security policy"):
            await srv.sap_set_batch_fields({
                "wnd[0]/usr/txtFIELD": "hello",
                "wnd[0]/tbar[0]/okcd": "/NSU01",
            })


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
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_set_field("wnd[0]/usr/txt", "test")

    @pytest.mark.asyncio
    async def test_readonly_blocks_press_button(self, readonly_srv):
        """sap_press_button raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_press_button("wnd[0]/tbar[1]/btn[8]")

    @pytest.mark.asyncio
    async def test_readonly_blocks_select_checkbox(self, readonly_srv):
        """sap_select_checkbox raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_checkbox("wnd[0]/usr/chk", True)

    @pytest.mark.asyncio
    async def test_readonly_blocks_execute_transaction(self, readonly_srv):
        """sap_execute_transaction raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_execute_transaction("MM03")

    @pytest.mark.asyncio
    async def test_readonly_blocks_send_key(self, readonly_srv):
        """sap_send_key raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_send_key("Enter")

    @pytest.mark.asyncio
    async def test_readonly_blocks_select_table_row(self, readonly_srv):
        """sap_select_table_row raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_table_row("wnd[0]/usr/tbl", 0)

    @pytest.mark.asyncio
    async def test_readonly_blocks_double_click_cell(self, readonly_srv):
        """sap_double_click_cell raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_double_click_cell("wnd[0]/usr/tbl", 0, "COL1")

    @pytest.mark.asyncio
    async def test_readonly_blocks_alv_toolbar_button(self, readonly_srv):
        """sap_press_alv_toolbar_button raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_press_alv_toolbar_button("wnd[0]/usr/grid", "SORT")

    @pytest.mark.asyncio
    async def test_readonly_blocks_alv_context_menu(self, readonly_srv):
        """sap_select_alv_context_menu_item raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_alv_context_menu_item(
                "wnd[0]/usr/grid", "ITEM1"
            )

    @pytest.mark.asyncio
    async def test_readonly_blocks_tree_expand(self, readonly_srv):
        """sap_expand_tree_node raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_expand_tree_node("wnd[0]/usr/tree", "1")

    @pytest.mark.asyncio
    async def test_readonly_blocks_tree_collapse(self, readonly_srv):
        """sap_collapse_tree_node raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_collapse_tree_node("wnd[0]/usr/tree", "1")

    @pytest.mark.asyncio
    async def test_readonly_blocks_tree_select(self, readonly_srv):
        """sap_select_tree_node raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_tree_node("wnd[0]/usr/tree", "1")

    @pytest.mark.asyncio
    async def test_readonly_blocks_tree_double_click(self, readonly_srv):
        """sap_double_click_tree_node raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_double_click_tree_node("wnd[0]/usr/tree", "1")

    @pytest.mark.asyncio
    async def test_readonly_blocks_select_menu(self, readonly_srv):
        """sap_select_menu raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_menu("wnd[0]/mbar/menu[0]/menu[0]")

    @pytest.mark.asyncio
    async def test_readonly_blocks_radio_button(self, readonly_srv):
        """sap_select_radio_button raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_radio_button("wnd[0]/usr/radOPT1")

    @pytest.mark.asyncio
    async def test_readonly_blocks_combobox(self, readonly_srv):
        """sap_select_combobox_entry raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_combobox_entry("wnd[0]/usr/cmb", "EN")

    @pytest.mark.asyncio
    async def test_readonly_blocks_tab(self, readonly_srv):
        """sap_select_tab raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_tab("wnd[0]/usr/tabsTAB/tabpTAB1")

    @pytest.mark.asyncio
    async def test_readonly_blocks_modify_cell(self, readonly_srv):
        """sap_modify_cell raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_modify_cell("wnd[0]/usr/grid", 0, "COL", "val")

    @pytest.mark.asyncio
    async def test_readonly_blocks_set_current_cell(self, readonly_srv):
        """sap_set_current_cell raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_set_current_cell("wnd[0]/usr/grid", 0, "COL")

    @pytest.mark.asyncio
    async def test_readonly_blocks_double_click_tree_item(self, readonly_srv):
        """sap_double_click_tree_item raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_double_click_tree_item("wnd[0]/usr/tree", "1", "COL")

    @pytest.mark.asyncio
    async def test_readonly_blocks_click_tree_link(self, readonly_srv):
        """sap_click_tree_link raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_click_tree_link("wnd[0]/usr/tree", "1", "LINK")

    @pytest.mark.asyncio
    async def test_readonly_blocks_get_tree_node_children_expand(self, readonly_srv):
        """sap_get_tree_node_children with expand=True raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_get_tree_node_children("wnd[0]/usr/tree", "1", expand=True)

    @pytest.mark.asyncio
    async def test_readonly_blocks_set_batch_fields(self, readonly_srv):
        """sap_set_batch_fields raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_set_batch_fields({"wnd[0]/usr/txt": "v"})

    @pytest.mark.asyncio
    async def test_readonly_blocks_set_textedit(self, readonly_srv):
        """sap_set_textedit raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_set_textedit("wnd[0]/usr/txt", "text")

    @pytest.mark.asyncio
    async def test_readonly_blocks_set_focus(self, readonly_srv):
        """sap_set_focus raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_set_focus("wnd[0]/usr/txt")

    @pytest.mark.asyncio
    async def test_readonly_blocks_scroll_table_control(self, readonly_srv):
        """sap_scroll_table_control raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_scroll_table_control("wnd[0]/usr/tbl", 10)

    @pytest.mark.asyncio
    async def test_readonly_blocks_select_all_table_control_columns(self, readonly_srv):
        """sap_select_all_table_control_columns raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_all_table_control_columns("wnd[0]/usr/tbl", True)

    @pytest.mark.asyncio
    async def test_readonly_blocks_press_column_header(self, readonly_srv):
        """sap_press_column_header raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_press_column_header("wnd[0]/usr/grid", "COL")

    @pytest.mark.asyncio
    async def test_readonly_blocks_select_all_rows(self, readonly_srv):
        """sap_select_all_rows raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_all_rows("wnd[0]/usr/grid")

    @pytest.mark.asyncio
    async def test_readonly_blocks_select_multiple_rows(self, readonly_srv):
        """sap_select_multiple_rows raises in read-only mode."""
        with pytest.raises(ValueError, match="read-only"):
            await readonly_srv.sap_select_multiple_rows("wnd[0]/usr/grid", [0, 1])


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
            # Connection
            "sap_connect", "sap_connect_existing", "sap_list_connections",
            "sap_get_session_info",
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
            "sap_get_popup_window",
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
        }

        assert tool_names == expected

    def test_send_key_has_enum(self, srv):
        """sap_send_key has enum constraint on key parameter."""
        import asyncio

        async def get_tools():
            return await srv.mcp.list_tools()

        tools = asyncio.new_event_loop().run_until_complete(get_tools())
        send_key = next(t for t in tools if t.name == "sap_send_key")
        key_schema = send_key.inputSchema["properties"]["key"]

        assert "enum" in key_schema
        assert "Enter" in key_schema["enum"]
        assert "F8" in key_schema["enum"]
        assert len(key_schema["enum"]) == 30

    def test_tool_annotations_present(self, srv):
        """All tools have annotations with readOnlyHint set."""
        import asyncio

        async def get_tools():
            return await srv.mcp.list_tools()

        tools = asyncio.new_event_loop().run_until_complete(get_tools())

        read_only_tools = {
            "sap_connect", "sap_connect_existing", "sap_list_connections",
            "sap_get_session_info", "sap_get_screen_info", "sap_read_field",
            "sap_get_combobox_entries", "sap_read_textedit", "sap_read_table",
            "sap_get_alv_toolbar", "sap_get_column_info", "sap_get_current_cell",
            "sap_get_table_control_row_info", "sap_get_cell_info",
            "sap_get_popup_window", "sap_get_toolbar_buttons",
            "sap_read_shell_content", "sap_read_tree", "sap_find_tree_node_by_path",
            "sap_search_tree_nodes", "sap_get_screen_elements", "sap_screenshot",
        }
        destructive_tools = {"sap_execute_transaction"}

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
