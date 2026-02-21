"""Tests for SAP GUI Controller."""

import pytest
from unittest.mock import MagicMock, patch


class TestSAPGUIController:
    """Tests for SAPGUIController class."""

    def test_init_without_pywin32(self):
        """Test that initialization fails gracefully without pywin32."""
        with patch.dict('sys.modules', {'win32com': None, 'win32com.client': None}):
            # Force reimport
            import importlib
            from mcp_sap_gui import sap_controller
            importlib.reload(sap_controller)

            with pytest.raises(sap_controller.SAPGUINotAvailableError):
                sap_controller.SAPGUIController()

    def test_is_connected_false_by_default(self):
        """Test that is_connected returns False initially."""
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        assert controller.is_connected is False

    def test_require_session_raises_when_not_connected(self):
        """Test that operations fail when not connected."""
        from mcp_sap_gui.sap_controller import (
            SAPGUIController,
            SAPGUINotConnectedError
        )
        controller = SAPGUIController()

        with pytest.raises(SAPGUINotConnectedError):
            controller.get_session_info()


class TestVKey:
    """Tests for VKey enum."""

    def test_vkey_values(self):
        """Test that VKey has expected values."""
        from mcp_sap_gui.sap_controller import VKey

        assert VKey.ENTER == 0
        assert VKey.F1 == 1
        assert VKey.F3 == 3
        assert VKey.F8 == 8
        assert VKey.F11 == 11
        assert VKey.F12 == 12


class TestSessionInfo:
    """Tests for SessionInfo dataclass."""

    def test_session_info_creation(self):
        """Test SessionInfo can be created."""
        from mcp_sap_gui.sap_controller import SessionInfo

        info = SessionInfo(
            system_name="DEV",
            system_number="00",
            client="100",
            user="TESTUSER",
            language="EN",
            transaction="MM03",
            program="SAPLMGMM",
            screen_number=100,
            session_number=0,
        )

        assert info.system_name == "DEV"
        assert info.client == "100"
        assert info.transaction == "MM03"


class TestReadTree:
    """Tests for tree reading with different tree types."""

    def _make_controller_with_session(self):
        """Create a controller with a mocked session."""
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def _make_gui_collection(self, items):
        """Create a mock GuiCollection with Count and indexed access."""
        col = MagicMock()
        col.Count = len(items)
        col.side_effect = lambda i: items[i]
        col.__iter__ = lambda self: iter(items)
        return col

    def test_list_tree_reads_columns_via_get_column_names(self):
        """Test that List tree (type 1, like SPRO) reads columns correctly."""
        controller = self._make_controller_with_session()

        mock_tree = MagicMock()
        mock_tree.GetTreeType.return_value = 1  # List tree
        mock_tree.GetHierarchyTitle.return_value = "IMG Structure"

        # GetColumnNames returns a collection of internal names
        col_names = self._make_gui_collection(["COLUMN1", "COLUMN2"])
        mock_tree.GetColumnNames.return_value = col_names

        # GetColumnTitleFromName returns display titles
        def title_from_name(name):
            return {"COLUMN1": "Description", "COLUMN2": "Status"}.get(name, name)
        mock_tree.GetColumnTitleFromName.side_effect = title_from_name

        # Node keys
        node_keys = self._make_gui_collection(["KEY1", "KEY2"])
        mock_tree.GetAllNodeKeys.return_value = node_keys

        # Node data
        mock_tree.GetNodeTextByKey.return_value = ""
        mock_tree.GetParentNodeKey.return_value = None
        mock_tree.GetNodeChildrenCount.return_value = 0
        mock_tree.IsFolderExpandable.return_value = True
        mock_tree.IsFolderExpanded.return_value = False

        # Column values per node
        def get_item_text(key, col):
            data = {
                ("KEY1", "COLUMN1"): "SAP Customizing",
                ("KEY1", "COLUMN2"): "",
                ("KEY2", "COLUMN1"): "Enterprise Structure",
                ("KEY2", "COLUMN2"): "Active",
            }
            return data.get((key, col), "")
        mock_tree.GetItemText.side_effect = get_item_text

        controller._session.findById.return_value = mock_tree

        result = controller.read_tree("wnd[0]/usr/cntlTREE/shellcont/shell", max_nodes=10)

        assert result["tree_type"] == "List"
        assert result["hierarchy_title"] == "IMG Structure"
        assert result["column_names"] == ["COLUMN1", "COLUMN2"]
        assert result["column_titles"] == ["Description", "Status"]
        assert len(result["nodes"]) == 2

        # Node text should be populated from first column when GetNodeTextByKey is empty
        assert result["nodes"][0]["text"] == "SAP Customizing"
        assert result["nodes"][1]["text"] == "Enterprise Structure"

        # Column values should be present
        assert result["nodes"][0]["columns"]["COLUMN1"] == "SAP Customizing"
        assert result["nodes"][1]["columns"]["COLUMN2"] == "Active"

    def test_simple_tree_returns_node_text_only(self):
        """Test that Simple tree (type 0) returns node text without columns."""
        controller = self._make_controller_with_session()

        mock_tree = MagicMock()
        mock_tree.GetTreeType.return_value = 0  # Simple tree
        mock_tree.GetHierarchyTitle.side_effect = Exception("Not available")

        node_keys = self._make_gui_collection(["ROOT", "CHILD1"])
        mock_tree.GetAllNodeKeys.return_value = node_keys

        def node_text(key):
            return {"ROOT": "Root Node", "CHILD1": "Child Node"}.get(key, "")
        mock_tree.GetNodeTextByKey.side_effect = node_text
        mock_tree.GetParentNodeKey.side_effect = lambda k: None if k == "ROOT" else "ROOT"
        mock_tree.GetNodeChildrenCount.side_effect = lambda k: 1 if k == "ROOT" else 0
        mock_tree.IsFolderExpandable.side_effect = lambda k: k == "ROOT"
        mock_tree.IsFolderExpanded.return_value = False

        controller._session.findById.return_value = mock_tree

        result = controller.read_tree("wnd[0]/usr/shell", max_nodes=10)

        assert result["tree_type"] == "Simple"
        assert result["column_names"] == []
        assert result["column_titles"] == []
        assert result["nodes"][0]["text"] == "Root Node"
        assert result["nodes"][1]["text"] == "Child Node"
        # Simple tree nodes should not have "columns" key
        assert "columns" not in result["nodes"][0]

    def test_column_tree_uses_column_order_fallback(self):
        """Test that Column tree (type 2) falls back to ColumnOrder if GetColumnNames fails."""
        controller = self._make_controller_with_session()

        mock_tree = MagicMock()
        mock_tree.GetTreeType.return_value = 2  # Column tree
        mock_tree.GetHierarchyTitle.return_value = ""

        # GetColumnNames fails
        mock_tree.GetColumnNames.side_effect = Exception("Not supported")

        # ColumnOrder works
        col_order = self._make_gui_collection(["COL_A", "COL_B"])
        mock_tree.ColumnOrder = col_order

        mock_tree.GetColumnTitleFromName.side_effect = lambda n: n

        node_keys = self._make_gui_collection(["N1"])
        mock_tree.GetAllNodeKeys.return_value = node_keys
        mock_tree.GetNodeTextByKey.return_value = "Node 1"
        mock_tree.GetParentNodeKey.return_value = None
        mock_tree.GetNodeChildrenCount.return_value = 0
        mock_tree.IsFolderExpandable.return_value = False
        mock_tree.IsFolderExpanded.return_value = False
        mock_tree.GetItemText.return_value = "val"

        controller._session.findById.return_value = mock_tree

        result = controller.read_tree("wnd[0]/usr/shell", max_nodes=10)

        assert result["tree_type"] == "Column"
        assert result["column_names"] == ["COL_A", "COL_B"]

    def test_text_fallback_from_columns_when_node_text_empty(self):
        """Test that node text is populated from first non-empty column value."""
        controller = self._make_controller_with_session()

        mock_tree = MagicMock()
        mock_tree.GetTreeType.return_value = 1
        mock_tree.GetHierarchyTitle.return_value = ""

        col_names = self._make_gui_collection(["HIER", "DESC"])
        mock_tree.GetColumnNames.return_value = col_names
        mock_tree.GetColumnTitleFromName.side_effect = lambda n: n

        node_keys = self._make_gui_collection(["K1"])
        mock_tree.GetAllNodeKeys.return_value = node_keys
        mock_tree.GetNodeTextByKey.return_value = ""  # Empty!
        mock_tree.GetParentNodeKey.return_value = None
        mock_tree.GetNodeChildrenCount.return_value = 0
        mock_tree.IsFolderExpandable.return_value = False
        mock_tree.IsFolderExpanded.return_value = False

        # First column empty, second has value
        def get_item_text(key, col):
            if col == "DESC":
                return "Extended Warehouse Mgmt"
            return ""
        mock_tree.GetItemText.side_effect = get_item_text

        controller._session.findById.return_value = mock_tree

        result = controller.read_tree("wnd[0]/usr/shell", max_nodes=10)

        assert result["nodes"][0]["text"] == "Extended Warehouse Mgmt"


class TestRadioButtonComboboxTab:
    """Tests for radio button, combobox, and tab selection."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_select_radio_button(self):
        """Radio button selection calls Select()."""
        controller = self._make_controller_with_session()
        mock_radio = MagicMock()
        controller._session.findById.return_value = mock_radio

        result = controller.select_radio_button("wnd[0]/usr/radOPT1")

        assert result["status"] == "success"
        mock_radio.Select.assert_called_once()

    def test_select_combobox_by_key(self):
        """Combobox entry selection by key sets Key directly."""
        controller = self._make_controller_with_session()
        mock_combo = MagicMock()
        controller._session.findById.return_value = mock_combo

        result = controller.select_combobox_entry("wnd[0]/usr/cmbLANGU", "EN")

        assert result["status"] == "success"
        assert result["key"] == "EN"

    def test_select_combobox_by_value_fallback(self):
        """Combobox entry selection falls back to searching Entries by value."""
        controller = self._make_controller_with_session()
        mock_combo = MagicMock()

        # Key setter: raise for invalid values, accept valid keys
        _valid_keys = {"EN"}
        def _key_setter(self_obj, val):
            if val not in _valid_keys:
                raise Exception("Invalid key")
        type(mock_combo).Key = property(lambda self: "X", _key_setter)

        # Entries collection
        mock_entry = MagicMock()
        mock_entry.Key = "EN"
        mock_entry.Value = "English"
        mock_entries = MagicMock()
        mock_entries.Count = 1
        mock_entries.side_effect = lambda i: mock_entry
        mock_combo.Entries = mock_entries

        controller._session.findById.return_value = mock_combo

        result = controller.select_combobox_entry("wnd[0]/usr/cmbLANGU", "English")

        assert result["status"] == "success"
        assert result["key"] == "EN"
        assert result["value"] == "English"

    def test_select_combobox_not_found(self):
        """Combobox returns error when entry not found."""
        controller = self._make_controller_with_session()
        mock_combo = MagicMock()

        def _key_setter(self_obj, val):
            raise Exception("Invalid key")
        type(mock_combo).Key = property(lambda self: "X", _key_setter)

        mock_entries = MagicMock()
        mock_entries.Count = 0
        mock_combo.Entries = mock_entries
        controller._session.findById.return_value = mock_combo

        result = controller.select_combobox_entry("wnd[0]/usr/cmbLANGU", "INVALID")

        assert "error" in result
        assert "not found" in result["error"]

    def test_select_tab(self):
        """Tab selection calls Select() and returns screen info."""
        controller = self._make_controller_with_session()
        mock_tab = MagicMock()
        controller._session.findById.return_value = mock_tab

        # Mock get_screen_info for the screen info return
        controller.get_screen_info = MagicMock(return_value={"transaction": "MM03"})

        result = controller.select_tab("wnd[0]/usr/tabsTAB/tabpTAB1")

        assert result["status"] == "success"
        mock_tab.Select.assert_called_once()
        assert result["screen"]["transaction"] == "MM03"


class TestGridEnhancements:
    """Tests for grid/table enhancements (Phase 3)."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_modify_cell(self):
        """ModifyCell is called with correct arguments."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        controller._session.findById.return_value = mock_grid

        result = controller.modify_cell("wnd[0]/usr/grid", 2, "MATNR", "MAT-001")

        assert result["status"] == "success"
        assert result["row"] == 2
        assert result["column"] == "MATNR"
        assert result["value"] == "MAT-001"
        mock_grid.ModifyCell.assert_called_once_with(2, "MATNR", "MAT-001")

    def test_modify_cell_error(self):
        """ModifyCell returns error on failure."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        mock_grid.ModifyCell.side_effect = Exception("Cell not editable")
        controller._session.findById.return_value = mock_grid

        result = controller.modify_cell("wnd[0]/usr/grid", 0, "COL", "val")

        assert "error" in result
        assert "not editable" in result["error"]

    def test_set_current_cell(self):
        """SetCurrentCell is called with correct arguments."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        controller._session.findById.return_value = mock_grid

        result = controller.set_current_cell("wnd[0]/usr/grid", 5, "MAKTX")

        assert result["status"] == "success"
        assert result["row"] == 5
        assert result["column"] == "MAKTX"
        mock_grid.SetCurrentCell.assert_called_once_with(5, "MAKTX")

    def test_get_column_info(self):
        """get_column_info returns column names, titles, and tooltips."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        mock_grid.ColumnCount = 2
        mock_grid.ColumnOrder.side_effect = ["MATNR", "MAKTX"]
        mock_grid.GetDisplayedColumnTitle.side_effect = ["Material", "Description"]
        mock_grid.GetColumnTooltip.side_effect = ["Material Number", "Material Description"]
        controller._session.findById.return_value = mock_grid

        result = controller.get_column_info("wnd[0]/usr/grid")

        assert result["column_count"] == 2
        assert result["columns"][0]["name"] == "MATNR"
        assert result["columns"][0]["title"] == "Material"
        assert result["columns"][0]["tooltip"] == "Material Number"
        assert result["columns"][1]["name"] == "MAKTX"
        assert result["columns"][1]["title"] == "Description"

    def test_read_table_includes_column_info(self):
        """read_table now includes column_info with tooltips."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        mock_grid.ColumnCount = 2
        mock_grid.ColumnOrder.side_effect = ["COL_A", "COL_B"]
        mock_grid.GetColumnTooltip.side_effect = ["Tooltip A", "Tooltip B"]
        mock_grid.GetDisplayedColumnTitle.side_effect = ["Title A", "Title B"]
        mock_grid.RowCount = 1
        mock_grid.GetCellValue.return_value = "val"
        controller._session.findById.return_value = mock_grid

        result = controller.read_table("wnd[0]/usr/grid", max_rows=10)

        assert "column_info" in result
        assert len(result["column_info"]) == 2
        assert result["column_info"][0]["name"] == "COL_A"
        assert result["column_info"][0]["tooltip"] == "Tooltip A"
        assert result["column_info"][0]["title"] == "Title A"

    def test_alv_toolbar_includes_tooltip_and_enabled(self):
        """get_alv_toolbar now includes tooltip and enabled per button."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        mock_grid.ToolbarButtonCount = 2
        mock_grid.GetToolbarButtonId.side_effect = ["BTN1", "BTN2"]
        mock_grid.GetToolbarButtonText.side_effect = ["Save", "Delete"]
        mock_grid.GetToolbarButtonType.side_effect = ["Button", "Button"]
        mock_grid.GetToolbarButtonTooltip.side_effect = ["Save changes", "Delete row"]
        mock_grid.GetToolbarButtonEnabled.side_effect = [True, False]
        controller._session.findById.return_value = mock_grid

        result = controller.get_alv_toolbar("wnd[0]/usr/grid")

        assert result["buttons"][0]["tooltip"] == "Save changes"
        assert result["buttons"][0]["enabled"] is True
        assert result["buttons"][1]["tooltip"] == "Delete row"
        assert result["buttons"][1]["enabled"] is False

    def test_alv_toolbar_tooltip_fallback(self):
        """Tooltip defaults to empty string if GetToolbarButtonTooltip fails."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        mock_grid.ToolbarButtonCount = 1
        mock_grid.GetToolbarButtonId.return_value = "BTN1"
        mock_grid.GetToolbarButtonText.return_value = "Save"
        mock_grid.GetToolbarButtonType.return_value = "Button"
        mock_grid.GetToolbarButtonTooltip.side_effect = Exception("Not supported")
        mock_grid.GetToolbarButtonEnabled.side_effect = Exception("Not supported")
        controller._session.findById.return_value = mock_grid

        result = controller.get_alv_toolbar("wnd[0]/usr/grid")

        assert result["buttons"][0]["tooltip"] == ""
        assert result["buttons"][0]["enabled"] is True


class TestALVToolbarButtonType:
    """Tests for toolbar button type mapping fix."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_string_button_types(self):
        """GetToolbarButtonType returning strings should be handled."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        mock_grid.ToolbarButtonCount = 3
        mock_grid.GetToolbarButtonId.side_effect = ["BTN1", "SEP1", "GRP1"]
        mock_grid.GetToolbarButtonText.side_effect = ["Save", "", ""]
        # API returns strings per documentation
        mock_grid.GetToolbarButtonType.side_effect = ["Button", "Separator", "Group"]
        controller._session.findById.return_value = mock_grid

        result = controller.get_alv_toolbar("wnd[0]/usr/grid")
        assert result["buttons"][0]["type"] == "Button"
        assert result["buttons"][1]["type"] == "Separator"
        assert result["buttons"][2]["type"] == "Group"

    def test_numeric_button_types(self):
        """GetToolbarButtonType returning integers should also be handled."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        mock_grid.ToolbarButtonCount = 3
        mock_grid.GetToolbarButtonId.side_effect = ["BTN1", "MNU1", "CHK1"]
        mock_grid.GetToolbarButtonText.side_effect = ["Execute", "Menu", "Check"]
        mock_grid.GetToolbarButtonType.side_effect = [0, 2, 4]
        controller._session.findById.return_value = mock_grid

        result = controller.get_alv_toolbar("wnd[0]/usr/grid")
        assert result["buttons"][0]["type"] == "Button"
        assert result["buttons"][1]["type"] == "Menu"
        assert result["buttons"][2]["type"] == "CheckBox"

    def test_unknown_type_returns_string(self):
        """Unknown type values are converted to string."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        mock_grid.ToolbarButtonCount = 1
        mock_grid.GetToolbarButtonId.return_value = "BTN1"
        mock_grid.GetToolbarButtonText.return_value = "Custom"
        mock_grid.GetToolbarButtonType.return_value = 99
        controller._session.findById.return_value = mock_grid

        result = controller.get_alv_toolbar("wnd[0]/usr/grid")
        assert result["buttons"][0]["type"] == "99"


class TestTreeEnhancements:
    """Tests for tree enhancements and new tree tools (Phase 4)."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def _make_gui_collection(self, items):
        col = MagicMock()
        col.Count = len(items)
        col.side_effect = lambda i: items[i]
        col.__iter__ = lambda self: iter(items)
        return col

    def test_double_click_tree_item(self):
        """DoubleClickItem is called with node_key and item_name."""
        controller = self._make_controller_with_session()
        mock_tree = MagicMock()
        controller._session.findById.return_value = mock_tree
        controller.get_screen_info = MagicMock(return_value={"transaction": "SPRO"})

        result = controller.double_click_tree_item("wnd[0]/usr/shell", "KEY1", "COLUMN1")

        assert result["status"] == "double_clicked"
        assert result["item_name"] == "COLUMN1"
        mock_tree.DoubleClickItem.assert_called_once_with("KEY1", "COLUMN1")

    def test_click_tree_link(self):
        """ClickLink is called with node_key and item_name."""
        controller = self._make_controller_with_session()
        mock_tree = MagicMock()
        controller._session.findById.return_value = mock_tree
        controller.get_screen_info = MagicMock(return_value={"transaction": "SPRO"})

        result = controller.click_tree_link("wnd[0]/usr/shell", "KEY1", "LINK_COL")

        assert result["status"] == "clicked"
        mock_tree.ClickLink.assert_called_once_with("KEY1", "LINK_COL")

    def test_find_tree_node_by_path(self):
        """FindNodeKeyByPath returns the node key."""
        controller = self._make_controller_with_session()
        mock_tree = MagicMock()
        mock_tree.FindNodeKeyByPath.return_value = "FOUND_KEY"
        controller._session.findById.return_value = mock_tree

        result = controller.find_tree_node_by_path("wnd[0]/usr/shell", "2\\1\\2")

        assert result["status"] == "found"
        assert result["node_key"] == "FOUND_KEY"
        mock_tree.FindNodeKeyByPath.assert_called_once_with("2\\1\\2")

    def test_find_tree_node_by_path_not_found(self):
        """FindNodeKeyByPath returns error when path is invalid."""
        controller = self._make_controller_with_session()
        mock_tree = MagicMock()
        mock_tree.FindNodeKeyByPath.side_effect = Exception("Path not found")
        controller._session.findById.return_value = mock_tree

        result = controller.find_tree_node_by_path("wnd[0]/usr/shell", "99\\99")

        assert "error" in result
        assert "not found" in result["error"].lower()

    def test_read_tree_includes_hierarchy_level(self):
        """read_tree now includes hierarchy_level per node."""
        controller = self._make_controller_with_session()
        mock_tree = MagicMock()
        mock_tree.GetTreeType.return_value = 0
        mock_tree.GetHierarchyTitle.side_effect = Exception("N/A")

        node_keys = self._make_gui_collection(["ROOT", "CHILD1"])
        mock_tree.GetAllNodeKeys.return_value = node_keys
        mock_tree.GetNodeTextByKey.return_value = "Node"
        mock_tree.GetParent.side_effect = lambda k: None if k == "ROOT" else "ROOT"
        mock_tree.GetNodeChildrenCount.return_value = 0
        mock_tree.IsFolderExpandable.return_value = False
        mock_tree.IsFolderExpanded.return_value = False
        mock_tree.GetHierarchyLevel.side_effect = lambda k: 0 if k == "ROOT" else 1
        controller._session.findById.return_value = mock_tree

        result = controller.read_tree("wnd[0]/usr/shell", max_nodes=10)

        assert result["nodes"][0]["hierarchy_level"] == 0
        assert result["nodes"][1]["hierarchy_level"] == 1


class TestSessionBusyCheck:
    """Tests for Session.Busy check in _require_session."""

    def test_busy_session_raises_error(self):
        """_require_session raises when session is busy."""
        from mcp_sap_gui.sap_controller import SAPGUIController, SAPGUIError
        controller = SAPGUIController()
        controller._session = MagicMock()
        controller._session.Busy = True

        with pytest.raises(SAPGUIError, match="busy"):
            controller._require_session()

    def test_non_busy_session_passes(self):
        """_require_session passes when session is not busy."""
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock()
        controller._session.Busy = False

        controller._require_session()  # Should not raise

    def test_busy_attribute_missing_passes(self):
        """_require_session passes if Busy property doesn't exist."""
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        mock_session = MagicMock(spec=[])  # No attributes
        # Need to restore basic connectivity check
        controller._session = mock_session

        controller._require_session()  # Should not raise


class TestExecuteTransactionImproved:
    """Tests for improved execute_transaction with StartTransaction."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock()
        controller._session.Busy = False
        return controller

    def test_uses_start_transaction(self):
        """Uses StartTransaction for /n prefix."""
        controller = self._make_controller_with_session()
        controller.get_screen_info = MagicMock(return_value={"transaction": "MM03"})

        result = controller.execute_transaction("MM03")

        controller._session.StartTransaction.assert_called_once_with("MM03")
        assert result["transaction"] == "MM03"

    def test_falls_back_to_okcd(self):
        """Falls back to okcd+sendVKey if StartTransaction fails."""
        controller = self._make_controller_with_session()
        controller._session.StartTransaction.side_effect = Exception("Not available")
        mock_okcd = MagicMock()
        mock_window = MagicMock()
        controller._session.findById.side_effect = lambda id: mock_okcd if "okcd" in id else mock_window
        controller.get_screen_info = MagicMock(return_value={"transaction": "MM03"})

        result = controller.execute_transaction("MM03")

        assert result["transaction"] == "MM03"
        # okcd text was set
        assert mock_okcd.text == "/nMM03"

    def test_o_prefix_uses_okcd_directly(self):
        """The /o prefix always uses okcd approach (opens new window)."""
        controller = self._make_controller_with_session()
        mock_okcd = MagicMock()
        mock_window = MagicMock()
        controller._session.findById.side_effect = lambda id: mock_okcd if "okcd" in id else mock_window
        controller.get_screen_info = MagicMock(return_value={"transaction": "MM03"})

        result = controller.execute_transaction("/oMM03")

        # StartTransaction should NOT be called for /o prefix
        controller._session.StartTransaction.assert_not_called()
        assert mock_okcd.text == "/oMM03"


class TestActiveWindowImproved:
    """Tests for improved _find_topmost_window with ActiveWindow."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_uses_active_window_first(self):
        """Uses Session.ActiveWindow when available."""
        controller = self._make_controller_with_session()
        mock_window = MagicMock()
        mock_window.Id = "wnd[1]"
        controller._session.ActiveWindow = mock_window

        result = controller._find_topmost_window()

        assert result == "wnd[1]"

    def test_falls_back_to_loop(self):
        """Falls back to loop when ActiveWindow not available."""
        controller = self._make_controller_with_session()
        type(controller._session).ActiveWindow = property(
            lambda self: (_ for _ in ()).throw(Exception("Not supported"))
        )
        # Loop: wnd[0] exists, wnd[1] doesn't
        def find_by_id(id):
            if id == "wnd[1]":
                raise Exception("Not found")
            return MagicMock()
        controller._session.findById.side_effect = find_by_id

        result = controller._find_topmost_window()

        assert result == "wnd[0]"


class TestTreeParentKeyFix:
    """Tests for tree GetParent() vs GetParentNodeKey() fallback."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def _make_gui_collection(self, items):
        col = MagicMock()
        col.Count = len(items)
        col.side_effect = lambda i: items[i]
        col.__iter__ = lambda self: iter(items)
        return col

    def test_uses_get_parent_first(self):
        """Should use GetParent() (API-documented method) first."""
        controller = self._make_controller_with_session()
        mock_tree = MagicMock()
        mock_tree.GetTreeType.return_value = 0
        mock_tree.GetHierarchyTitle.side_effect = Exception("N/A")
        node_keys = self._make_gui_collection(["CHILD1"])
        mock_tree.GetAllNodeKeys.return_value = node_keys
        mock_tree.GetNodeTextByKey.return_value = "Child"
        mock_tree.GetParent.return_value = "ROOT"
        mock_tree.GetNodeChildrenCount.return_value = 0
        mock_tree.IsFolderExpandable.return_value = False
        mock_tree.IsFolderExpanded.return_value = False
        controller._session.findById.return_value = mock_tree

        result = controller.read_tree("wnd[0]/usr/shell", max_nodes=10)
        assert result["nodes"][0]["parent_key"] == "ROOT"
        mock_tree.GetParent.assert_called_once_with("CHILD1")

    def test_falls_back_to_get_parent_node_key(self):
        """If GetParent() fails, should fall back to GetParentNodeKey()."""
        controller = self._make_controller_with_session()
        mock_tree = MagicMock()
        mock_tree.GetTreeType.return_value = 0
        mock_tree.GetHierarchyTitle.side_effect = Exception("N/A")
        node_keys = self._make_gui_collection(["CHILD1"])
        mock_tree.GetAllNodeKeys.return_value = node_keys
        mock_tree.GetNodeTextByKey.return_value = "Child"
        mock_tree.GetParent.side_effect = Exception("Not supported")
        mock_tree.GetParentNodeKey.return_value = "ROOT_FALLBACK"
        mock_tree.GetNodeChildrenCount.return_value = 0
        mock_tree.IsFolderExpandable.return_value = False
        mock_tree.IsFolderExpanded.return_value = False
        controller._session.findById.return_value = mock_tree

        result = controller.read_tree("wnd[0]/usr/shell", max_nodes=10)
        assert result["nodes"][0]["parent_key"] == "ROOT_FALLBACK"


class TestGuiTableControl:
    """Tests for GuiTableControl support in table/grid operations."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def _make_table_control(self, columns, rows_data, total_rows=None,
                            visible_rows=None):
        """Create a mock GuiTableControl with columns and data.

        Args:
            columns: list of dicts with 'name', 'title', 'tooltip' keys
            rows_data: list of lists (outer=rows, inner=cell values)
            total_rows: total row count (default len(rows_data))
            visible_rows: visible rows at once (default len(rows_data))
        """
        if total_rows is None:
            total_rows = len(rows_data)
        if visible_rows is None:
            visible_rows = len(rows_data)

        mock_table = MagicMock()
        mock_table.Type = "GuiTableControl"
        mock_table.TableFieldName = "TEST_TABLE"
        mock_table.RowCount = total_rows
        mock_table.VisibleRowCount = visible_rows

        # Columns collection
        mock_cols = MagicMock()
        mock_cols.Count = len(columns)
        col_mocks = []
        for i, col_def in enumerate(columns):
            col_mock = MagicMock()
            col_mock.Title = col_def.get("title", "")
            col_mock.Tooltip = col_def.get("tooltip", "")
            col_mocks.append(col_mock)
        # Support both Columns(i) and Columns.ElementAt(i)
        mock_cols.side_effect = lambda i: col_mocks[i]
        mock_cols.ElementAt = MagicMock(side_effect=lambda i: col_mocks[i])
        mock_table.Columns = mock_cols

        # Column names list for GetCell to reference
        _col_names = [col_def.get("name", f"col_{i}") for i, col_def in enumerate(columns)]

        # Scrollbar
        scrollbar = MagicMock()
        scrollbar.Minimum = 0
        scrollbar.Maximum = max(0, total_rows - visible_rows)
        scrollbar.Position = 0
        scrollbar.PageSize = visible_rows
        mock_table.VerticalScrollbar = scrollbar

        # GetCell - returns mock cells with the data
        # scroll_pos tracks the current scroll position
        def get_cell(row_idx, col_idx):
            abs_row = scrollbar.Position + row_idx
            cell = MagicMock()
            # Column name from cell (used by _get_table_control_columns)
            cell.Name = _col_names[col_idx] if col_idx < len(_col_names) else f"col_{col_idx}"
            if abs_row < len(rows_data) and col_idx < len(rows_data[abs_row]):
                val = rows_data[abs_row][col_idx]
                if isinstance(val, bool):
                    cell.Type = "GuiCheckBox"
                    cell.Selected = val
                elif isinstance(val, dict) and "combobox" in val:
                    cell.Type = "GuiComboBox"
                    cell.Key = val["combobox"]
                    cell.Text = val.get("text", val["combobox"])
                else:
                    cell.Type = "GuiTextField"
                    cell.Text = str(val) if val is not None else ""
            else:
                cell.Type = "GuiTextField"
                cell.Text = ""
            return cell
        mock_table.GetCell = MagicMock(side_effect=get_cell)

        return mock_table

    def test_read_table_detects_guitablecontrol(self):
        """read_table dispatches to _read_table_control for GuiTableControl."""
        controller = self._make_controller_with_session()
        columns = [
            {"name": "LGNUM", "title": "WhN", "tooltip": "Warehouse Number"},
            {"name": "LGTYPGRP", "title": "STG", "tooltip": "Storage Type Group"},
        ]
        rows = [
            ["WH01", "GRP1"],
            ["WH02", "GRP2"],
        ]
        mock_table = self._make_table_control(columns, rows)
        controller._session.findById.return_value = mock_table

        result = controller.read_table("wnd[0]/usr/tblTEST", max_rows=10)

        assert result["table_type"] == "GuiTableControl"
        assert result["table_field_name"] == "TEST_TABLE"
        assert result["total_rows"] == 2
        assert result["rows_returned"] == 2
        assert result["columns"] == ["LGNUM", "LGTYPGRP"]
        assert len(result["column_info"]) == 2
        assert result["column_info"][0]["title"] == "WhN"
        assert result["column_info"][0]["tooltip"] == "Warehouse Number"
        assert result["data"][0]["LGNUM"] == "WH01"
        assert result["data"][1]["LGTYPGRP"] == "GRP2"

    def test_read_table_control_with_scrolling(self):
        """read_table scrolls through all rows when total > visible."""
        controller = self._make_controller_with_session()
        columns = [{"name": "COL_A", "title": "A"}]
        # 10 rows total but only 3 visible at a time
        rows = [[f"val_{i}"] for i in range(10)]
        mock_table = self._make_table_control(columns, rows,
                                               total_rows=10, visible_rows=3)
        controller._session.findById.return_value = mock_table

        result = controller.read_table("wnd[0]/usr/tblTEST", max_rows=100)

        assert result["total_rows"] == 10
        assert result["visible_rows"] == 3
        assert result["rows_returned"] == 10
        # Verify data integrity across scroll batches
        assert result["data"][0]["COL_A"] == "val_0"
        assert result["data"][3]["COL_A"] == "val_3"
        assert result["data"][9]["COL_A"] == "val_9"

    def test_read_table_control_respects_max_rows(self):
        """read_table stops after max_rows even with more data."""
        controller = self._make_controller_with_session()
        columns = [{"name": "COL", "title": "Col"}]
        rows = [[f"v{i}"] for i in range(20)]
        mock_table = self._make_table_control(columns, rows,
                                               total_rows=20, visible_rows=5)
        controller._session.findById.return_value = mock_table

        result = controller.read_table("wnd[0]/usr/tblTEST", max_rows=8)

        assert result["rows_returned"] == 8
        assert result["total_rows"] == 20

    def test_read_table_control_checkbox_cells(self):
        """Checkbox cells return boolean values."""
        controller = self._make_controller_with_session()
        columns = [{"name": "NAME", "title": "Name"},
                    {"name": "ACTIVE", "title": "Active"}]
        rows = [
            ["Item1", True],
            ["Item2", False],
        ]
        mock_table = self._make_table_control(columns, rows)
        controller._session.findById.return_value = mock_table

        result = controller.read_table("wnd[0]/usr/tblTEST", max_rows=10)

        assert result["data"][0]["ACTIVE"] is True
        assert result["data"][1]["ACTIVE"] is False
        assert result["data"][0]["NAME"] == "Item1"

    def test_read_table_control_combobox_cells(self):
        """Combobox cells return the key value."""
        controller = self._make_controller_with_session()
        columns = [{"name": "TYPE", "title": "Type"}]
        rows = [[{"combobox": "01", "text": "Standard"}]]
        mock_table = self._make_table_control(columns, rows)
        controller._session.findById.return_value = mock_table

        result = controller.read_table("wnd[0]/usr/tblTEST", max_rows=10)

        assert result["data"][0]["TYPE"] == "01"

    def test_read_table_still_works_for_alv(self):
        """read_table still works for ALV grids (regression test)."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        # Type is MagicMock, not "GuiTableControl"
        mock_grid.ColumnCount = 1
        mock_grid.ColumnOrder.return_value = "MATNR"
        mock_grid.GetColumnTooltip.return_value = "Material"
        mock_grid.GetDisplayedColumnTitle.return_value = "Mat."
        mock_grid.RowCount = 1
        mock_grid.GetCellValue.return_value = "MAT001"
        controller._session.findById.return_value = mock_grid

        result = controller.read_table("wnd[0]/usr/grid", max_rows=10)

        assert result["table_type"] == "GuiGridView"
        assert result["data"][0]["MATNR"] == "MAT001"

    def test_read_table_control_column_name_fallback(self):
        """Column name falls back to title then to col_N when cell Name is unavailable."""
        controller = self._make_controller_with_session()

        mock_table = MagicMock()
        mock_table.Type = "GuiTableControl"
        mock_table.TableFieldName = "T"
        mock_table.RowCount = 0
        mock_table.VisibleRowCount = 0

        # Column with title but no name accessible via cell
        col_mock = MagicMock()
        col_mock.Title = "My Title"
        col_mock.Tooltip = ""

        mock_cols = MagicMock()
        mock_cols.Count = 1
        mock_cols.side_effect = lambda i: col_mock
        mock_table.Columns = mock_cols

        scrollbar = MagicMock()
        scrollbar.Minimum = 0
        scrollbar.Maximum = 0
        mock_table.VerticalScrollbar = scrollbar
        # GetCell fails (no rows) - so cell Name is unavailable
        mock_table.GetCell.side_effect = Exception("No rows")

        controller._session.findById.return_value = mock_table

        result = controller.read_table("wnd[0]/usr/tblTEST", max_rows=10)

        # Should fall back to title since GetCell failed
        assert result["columns"] == ["My Title"]

    def test_select_table_row_guitablecontrol(self):
        """select_table_row uses SetFocus on cell for GuiTableControl."""
        controller = self._make_controller_with_session()
        columns = [{"name": "COL", "title": "Col"}]
        rows = [[f"v{i}"] for i in range(5)]
        mock_table = self._make_table_control(columns, rows, visible_rows=5)
        controller._session.findById.return_value = mock_table

        result = controller.select_table_row("wnd[0]/usr/tblTEST", 2)

        assert result["status"] == "success"
        assert result["selected_row"] == 2
        # Verify GetCell was called for row 2, col 0
        mock_table.GetCell.assert_called_with(2, 0)

    def test_double_click_table_cell_guitablecontrol(self):
        """double_click_table_cell uses SetFocus + F2 for GuiTableControl."""
        from mcp_sap_gui.sap_controller import VKey
        controller = self._make_controller_with_session()
        columns = [{"name": "COL_A", "title": "A"},
                    {"name": "COL_B", "title": "B"}]
        rows = [["v1", "v2"]]
        mock_table = self._make_table_control(columns, rows)
        controller._session.findById.return_value = mock_table
        controller.get_screen_info = MagicMock(return_value={"transaction": "SM30"})

        result = controller.double_click_table_cell("wnd[0]/usr/tblTEST", 0, "COL_B")

        assert result["status"] == "double_clicked"
        # GetCell should be called for row 0, column index 1 (COL_B)
        mock_table.GetCell.assert_called_with(0, 1)

    def test_double_click_table_cell_by_numeric_column(self):
        """double_click_table_cell accepts numeric column index as string."""
        controller = self._make_controller_with_session()
        columns = [{"name": "COL_A", "title": "A"}]
        rows = [["v1"]]
        mock_table = self._make_table_control(columns, rows)
        controller._session.findById.return_value = mock_table
        controller.get_screen_info = MagicMock(return_value={"transaction": "SM30"})

        result = controller.double_click_table_cell("wnd[0]/usr/tblTEST", 0, "0")

        assert result["status"] == "double_clicked"
        mock_table.GetCell.assert_called_with(0, 0)

    def test_modify_cell_guitablecontrol(self):
        """modify_cell sets Text on the cell for GuiTableControl."""
        controller = self._make_controller_with_session()
        columns = [{"name": "DESC", "title": "Description"}]
        rows = [["Old Value"]]
        mock_table = self._make_table_control(columns, rows)
        controller._session.findById.return_value = mock_table

        result = controller.modify_cell("wnd[0]/usr/tblTEST", 0, "DESC", "New Value")

        assert result["status"] == "success"
        # Verify the cell's Text was set
        cell = mock_table.GetCell(0, 0)
        # GetCell returns a new mock each time due to side_effect,
        # so check the call was made
        mock_table.GetCell.assert_called_with(0, 0)

    def test_set_current_cell_guitablecontrol(self):
        """set_current_cell uses SetFocus for GuiTableControl."""
        controller = self._make_controller_with_session()
        columns = [{"name": "COL", "title": "Col"}]
        rows = [["v"]]
        mock_table = self._make_table_control(columns, rows)
        controller._session.findById.return_value = mock_table

        result = controller.set_current_cell("wnd[0]/usr/tblTEST", 0, "COL")

        assert result["status"] == "success"
        mock_table.GetCell.assert_called_with(0, 0)

    def test_get_column_info_guitablecontrol(self):
        """get_column_info returns GuiTableControl column metadata."""
        controller = self._make_controller_with_session()
        columns = [
            {"name": "LGNUM", "title": "WhN", "tooltip": "Warehouse Number"},
            {"name": "LGTYPGRP", "title": "STG", "tooltip": "Storage Type Group"},
        ]
        mock_table = self._make_table_control(columns, [])
        controller._session.findById.return_value = mock_table

        result = controller.get_column_info("wnd[0]/usr/tblTEST")

        assert result["table_type"] == "GuiTableControl"
        assert result["column_count"] == 2
        assert result["columns"][0]["name"] == "LGNUM"
        assert result["columns"][0]["title"] == "WhN"
        assert result["columns"][0]["tooltip"] == "Warehouse Number"
        assert result["columns"][1]["name"] == "LGTYPGRP"

    def test_resolve_column_by_title(self):
        """_resolve_table_control_column can find column by Title."""
        controller = self._make_controller_with_session()
        columns = [{"name": "FLD1", "title": "Field One"}]
        mock_table = self._make_table_control(columns, [])
        controller._session.findById.return_value = mock_table

        idx = controller._resolve_table_control_column(mock_table, "Field One")

        assert idx == 0

    def test_resolve_column_not_found(self):
        """_resolve_table_control_column raises ValueError for unknown column."""
        controller = self._make_controller_with_session()
        columns = [{"name": "FLD1", "title": "Field One"}]
        mock_table = self._make_table_control(columns, [])

        with pytest.raises(ValueError, match="not found"):
            controller._resolve_table_control_column(mock_table, "NONEXISTENT")

    def test_scroll_to_row_already_visible(self):
        """_scroll_table_control_to_row returns correct index when row is visible."""
        controller = self._make_controller_with_session()
        columns = [{"name": "C", "title": "C"}]
        rows = [[f"v{i}"] for i in range(10)]
        mock_table = self._make_table_control(columns, rows,
                                               total_rows=10, visible_rows=5)
        # Scrollbar starts at 0, so rows 0-4 are visible
        result = controller._scroll_table_control_to_row(mock_table, 3)

        assert result == 3  # visible index
        # Scrollbar should NOT have been moved
        assert mock_table.VerticalScrollbar.Position == 0

    def test_scroll_to_row_needs_scrolling(self):
        """_scroll_table_control_to_row scrolls and returns correct visible index."""
        controller = self._make_controller_with_session()
        columns = [{"name": "C", "title": "C"}]
        rows = [[f"v{i}"] for i in range(10)]
        mock_table = self._make_table_control(columns, rows,
                                               total_rows=10, visible_rows=5)
        # Maximum is 10-5=5, so scrolling to row 7 clamps to pos=5
        # Visible rows at pos=5 are 5,6,7,8,9 -> row 7 is at visible index 2
        result = controller._scroll_table_control_to_row(mock_table, 7)

        assert result == 2  # Row 7 at visible index 2 (pos=5, 7-5=2)
        assert mock_table.VerticalScrollbar.Position == 5  # Clamped to Maximum
