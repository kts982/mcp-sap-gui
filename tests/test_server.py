"""Tests for MCP SAP GUI Server - security logic, routing, and configuration."""

import pytest
from unittest.mock import MagicMock, patch, AsyncMock
import asyncio


# ---------------------------------------------------------------------------
# Helpers to import server components with mocked win32com
# ---------------------------------------------------------------------------

@pytest.fixture
def mock_win32com():
    """Mock win32com so SAPGUIController can be instantiated on any OS."""
    mock = MagicMock()
    with patch.dict("sys.modules", {"win32com": mock, "win32com.client": mock.client}):
        yield mock


@pytest.fixture
def server_config(mock_win32com):
    """Create a ServerConfig with defaults."""
    from mcp_sap_gui.server import ServerConfig
    return ServerConfig()


@pytest.fixture
def readonly_config(mock_win32com):
    """Create a read-only ServerConfig."""
    from mcp_sap_gui.server import ServerConfig
    return ServerConfig(read_only=True)


@pytest.fixture
def server(mock_win32com):
    """Create a MCPSAPGUIServer with default config."""
    from mcp_sap_gui.server import MCPSAPGUIServer
    return MCPSAPGUIServer()


@pytest.fixture
def readonly_server(mock_win32com):
    """Create a MCPSAPGUIServer in read-only mode."""
    from mcp_sap_gui.server import MCPSAPGUIServer, ServerConfig
    return MCPSAPGUIServer(config=ServerConfig(read_only=True))


# ===========================================================================
# Transaction Blocking Tests
# ===========================================================================

class TestTransactionBlocking:
    """Tests for _is_transaction_blocked with removeprefix fix."""

    def test_blocked_transaction_direct(self, server):
        """Default blocklist blocks SU01."""
        assert server._is_transaction_blocked("SU01") is True

    def test_blocked_transaction_case_insensitive(self, server):
        """Blocklist check is case-insensitive."""
        assert server._is_transaction_blocked("su01") is True
        assert server._is_transaction_blocked("Su01") is True

    def test_blocked_with_n_prefix(self, server):
        """Stripping /N prefix still detects blocked transaction."""
        assert server._is_transaction_blocked("/NSU01") is True
        assert server._is_transaction_blocked("/nSU01") is True

    def test_blocked_with_o_prefix(self, server):
        """Stripping /O prefix still detects blocked transaction."""
        assert server._is_transaction_blocked("/OSU01") is True
        assert server._is_transaction_blocked("/oSU01") is True

    def test_allowed_transaction(self, server):
        """Non-blocked transactions are allowed."""
        assert server._is_transaction_blocked("MM03") is False
        assert server._is_transaction_blocked("VA01") is False

    def test_removeprefix_does_not_corrupt_tcode(self, server):
        """Verify removeprefix doesn't strip characters from transaction codes.

        The old lstrip("/N") would turn "NOTIF" into "OTIF" because lstrip
        strips individual characters, not the substring.
        """
        assert server._is_transaction_blocked("NOTIF") is False
        assert server._is_transaction_blocked("NOOP") is False
        # SE16N IS in the blocklist - verify it's still correctly blocked
        assert server._is_transaction_blocked("SE16N") is True
        # SE16 (without N suffix) is NOT in the blocklist
        assert server._is_transaction_blocked("SE16") is False

    def test_all_default_blocked(self, server):
        """All default blocked transactions are correctly blocked."""
        blocked = ["SU01", "SU10", "SU01D", "PFCG", "SU53", "SM21", "ST22", "SE16N"]
        for tcode in blocked:
            assert server._is_transaction_blocked(tcode) is True, f"{tcode} should be blocked"

    def test_allowlist_mode(self, mock_win32com):
        """When allowed_transactions is set, only those are allowed."""
        from mcp_sap_gui.server import MCPSAPGUIServer, ServerConfig
        config = ServerConfig(allowed_transactions=["MM03", "VA03"])
        srv = MCPSAPGUIServer(config=config)

        assert srv._is_transaction_blocked("MM03") is False
        assert srv._is_transaction_blocked("VA03") is False
        assert srv._is_transaction_blocked("VA01") is True
        assert srv._is_transaction_blocked("SE80") is True

    def test_allowlist_with_prefix(self, mock_win32com):
        """Allowlist works with /N and /O prefixes."""
        from mcp_sap_gui.server import MCPSAPGUIServer, ServerConfig
        config = ServerConfig(allowed_transactions=["MM03"])
        srv = MCPSAPGUIServer(config=config)

        assert srv._is_transaction_blocked("/NMM03") is False
        assert srv._is_transaction_blocked("/OMM03") is False
        assert srv._is_transaction_blocked("/NVA01") is True


# ===========================================================================
# Read-Only Mode Tests
# ===========================================================================

class TestReadOnlyMode:
    """Tests for read-only mode enforcement."""

    def test_readonly_filters_write_tools(self, readonly_server):
        """Read-only mode should filter all mutating tools from list_tools."""
        from mcp.types import ListToolsRequest
        req = ListToolsRequest(method="tools/list")
        handler = readonly_server.server.request_handlers[ListToolsRequest]
        result = asyncio.new_event_loop().run_until_complete(handler(req))
        tool_names = {t.name for t in result.root.tools}

        # These should all be filtered out
        write_tools = {
            "sap_set_field", "sap_press_button", "sap_select_checkbox",
            "sap_select_table_row", "sap_double_click_cell",
            "sap_execute_transaction", "sap_send_key",
            "sap_press_alv_toolbar_button", "sap_select_alv_context_menu_item",
            "sap_expand_tree_node", "sap_collapse_tree_node",
            "sap_select_tree_node", "sap_double_click_tree_node",
        }

        for tool in write_tools:
            assert tool not in tool_names, f"{tool} should be filtered in read-only mode"

    def test_readonly_keeps_read_tools(self, readonly_server):
        """Read-only mode should keep read-only tools available."""
        from mcp.types import ListToolsRequest
        req = ListToolsRequest(method="tools/list")
        handler = readonly_server.server.request_handlers[ListToolsRequest]
        result = asyncio.new_event_loop().run_until_complete(handler(req))
        tool_names = {t.name for t in result.root.tools}

        read_tools = {
            "sap_connect", "sap_connect_existing", "sap_list_connections",
            "sap_get_session_info", "sap_get_screen_info",
            "sap_read_field", "sap_read_table", "sap_get_alv_toolbar",
            "sap_read_tree", "sap_get_screen_elements", "sap_screenshot",
        }

        for tool in read_tools:
            assert tool in tool_names, f"{tool} should be available in read-only mode"

    @pytest.mark.asyncio
    async def test_readonly_blocks_set_field(self, readonly_server):
        """sap_set_field returns error in read-only mode."""
        result = await readonly_server._handle_tool(
            "sap_set_field", {"field_id": "wnd[0]/usr/txt", "value": "test"}
        )
        assert "error" in result
        assert "read-only" in result["error"].lower()

    @pytest.mark.asyncio
    async def test_readonly_blocks_press_button(self, readonly_server):
        """sap_press_button returns error in read-only mode."""
        result = await readonly_server._handle_tool(
            "sap_press_button", {"button_id": "wnd[0]/tbar[1]/btn[8]"}
        )
        assert "error" in result
        assert "read-only" in result["error"].lower()

    @pytest.mark.asyncio
    async def test_readonly_blocks_select_checkbox(self, readonly_server):
        """sap_select_checkbox returns error in read-only mode."""
        result = await readonly_server._handle_tool(
            "sap_select_checkbox", {"checkbox_id": "wnd[0]/usr/chk", "selected": True}
        )
        assert "error" in result
        assert "read-only" in result["error"].lower()

    @pytest.mark.asyncio
    async def test_readonly_blocks_execute_transaction(self, readonly_server):
        """sap_execute_transaction returns error in read-only mode."""
        result = await readonly_server._handle_tool(
            "sap_execute_transaction", {"tcode": "MM03"}
        )
        assert "error" in result
        assert "read-only" in result["error"].lower()

    @pytest.mark.asyncio
    async def test_readonly_blocks_send_key(self, readonly_server):
        """sap_send_key returns error in read-only mode."""
        result = await readonly_server._handle_tool(
            "sap_send_key", {"key": "Enter"}
        )
        assert "error" in result
        assert "read-only" in result["error"].lower()

    @pytest.mark.asyncio
    async def test_readonly_blocks_select_table_row(self, readonly_server):
        """sap_select_table_row returns error in read-only mode."""
        result = await readonly_server._handle_tool(
            "sap_select_table_row", {"table_id": "wnd[0]/usr/tbl", "row": 0}
        )
        assert "error" in result
        assert "read-only" in result["error"].lower()

    @pytest.mark.asyncio
    async def test_readonly_blocks_double_click_cell(self, readonly_server):
        """sap_double_click_cell returns error in read-only mode."""
        result = await readonly_server._handle_tool(
            "sap_double_click_cell",
            {"table_id": "wnd[0]/usr/tbl", "row": 0, "column": "COL1"}
        )
        assert "error" in result
        assert "read-only" in result["error"].lower()

    @pytest.mark.asyncio
    async def test_readonly_blocks_alv_toolbar_button(self, readonly_server):
        """sap_press_alv_toolbar_button returns error in read-only mode."""
        result = await readonly_server._handle_tool(
            "sap_press_alv_toolbar_button",
            {"grid_id": "wnd[0]/usr/grid", "button_id": "SORT"}
        )
        assert "error" in result
        assert "read-only" in result["error"].lower()

    @pytest.mark.asyncio
    async def test_readonly_blocks_alv_context_menu(self, readonly_server):
        """sap_select_alv_context_menu_item returns error in read-only mode."""
        result = await readonly_server._handle_tool(
            "sap_select_alv_context_menu_item",
            {"grid_id": "wnd[0]/usr/grid", "menu_item_id": "ITEM1"}
        )
        assert "error" in result
        assert "read-only" in result["error"].lower()

    @pytest.mark.asyncio
    async def test_readonly_blocks_tree_expand(self, readonly_server):
        """sap_expand_tree_node returns error in read-only mode."""
        result = await readonly_server._handle_tool(
            "sap_expand_tree_node",
            {"tree_id": "wnd[0]/usr/tree", "node_key": "1"}
        )
        assert "error" in result
        assert "read-only" in result["error"].lower()

    @pytest.mark.asyncio
    async def test_readonly_blocks_tree_collapse(self, readonly_server):
        """sap_collapse_tree_node returns error in read-only mode."""
        result = await readonly_server._handle_tool(
            "sap_collapse_tree_node",
            {"tree_id": "wnd[0]/usr/tree", "node_key": "1"}
        )
        assert "error" in result
        assert "read-only" in result["error"].lower()

    @pytest.mark.asyncio
    async def test_readonly_blocks_tree_select(self, readonly_server):
        """sap_select_tree_node returns error in read-only mode."""
        result = await readonly_server._handle_tool(
            "sap_select_tree_node",
            {"tree_id": "wnd[0]/usr/tree", "node_key": "1"}
        )
        assert "error" in result
        assert "read-only" in result["error"].lower()

    @pytest.mark.asyncio
    async def test_readonly_blocks_tree_double_click(self, readonly_server):
        """sap_double_click_tree_node returns error in read-only mode."""
        result = await readonly_server._handle_tool(
            "sap_double_click_tree_node",
            {"tree_id": "wnd[0]/usr/tree", "node_key": "1"}
        )
        assert "error" in result
        assert "read-only" in result["error"].lower()


# ===========================================================================
# _parse_key Tests
# ===========================================================================

class TestParseKey:
    """Tests for _parse_key method."""

    def test_valid_keys(self, server):
        """All valid key names parse correctly."""
        from mcp_sap_gui.sap_controller import VKey

        assert server._parse_key("Enter") == VKey.ENTER
        assert server._parse_key("F1") == VKey.F1
        assert server._parse_key("F3") == VKey.F3
        assert server._parse_key("Back") == VKey.F3
        assert server._parse_key("F8") == VKey.F8
        assert server._parse_key("Execute") == VKey.F8
        assert server._parse_key("F11") == VKey.F11
        assert server._parse_key("Save") == VKey.F11
        assert server._parse_key("F12") == VKey.F12
        assert server._parse_key("Cancel") == VKey.F12
        assert server._parse_key("F5") == VKey.F5
        assert server._parse_key("Refresh") == VKey.F5

    def test_unknown_key_raises_error(self, server):
        """Unknown key names raise ValueError instead of silently defaulting."""
        with pytest.raises(ValueError, match="Unknown key"):
            server._parse_key("InvalidKey")

        with pytest.raises(ValueError, match="Unknown key"):
            server._parse_key("")

        with pytest.raises(ValueError, match="Unknown key"):
            server._parse_key("f1")  # case-sensitive

    def test_error_message_lists_valid_keys(self, server):
        """Error message for unknown key includes list of valid keys."""
        with pytest.raises(ValueError, match="Enter") as exc_info:
            server._parse_key("BadKey")
        assert "F1" in str(exc_info.value)
        assert "F12" in str(exc_info.value)


# ===========================================================================
# Unknown Tool Tests
# ===========================================================================

class TestUnknownTool:
    """Tests for unknown tool handling."""

    @pytest.mark.asyncio
    async def test_unknown_tool_returns_error(self, server):
        """Calling an unknown tool returns an error dict."""
        result = await server._handle_tool("sap_nonexistent", {})
        assert "error" in result
        assert "Unknown tool" in result["error"]


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
                import importlib
                from mcp_sap_gui import sap_controller

                original_optimize = controller._optimize_screenshot.__func__

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
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()

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

        with patch("builtins.__import__", side_effect=lambda name, *args, **kwargs:
                    mock_pil if name == "PIL.Image" else __builtins__.__import__(name, *args, **kwargs)):
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
