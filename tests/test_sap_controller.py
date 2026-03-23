"""Tests for SAP GUI Controller."""

from unittest.mock import MagicMock, patch

import pytest


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
        from mcp_sap_gui.sap_controller import SAPGUIController, SAPGUINotConnectedError
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


class TestSearchTreeNodes:
    """Tests for tree node search functionality."""

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

    def _build_tree(self, node_defs):
        """Build a mock tree from a list of (key, text, parent_key) tuples.

        Returns a MagicMock tree configured for search_tree_nodes.
        """
        mock_tree = MagicMock()
        mock_tree.GetTreeType.return_value = 0  # Simple tree

        keys = [n[0] for n in node_defs]
        mock_tree.GetAllNodeKeys.return_value = self._make_gui_collection(keys)

        text_map = {n[0]: n[1] for n in node_defs}
        parent_map = {n[0]: n[2] for n in node_defs}

        mock_tree.GetNodeTextByKey.side_effect = lambda k: text_map.get(k, "")
        mock_tree.GetParent.side_effect = lambda k: parent_map.get(k)
        mock_tree.IsFolderExpandable.return_value = False
        mock_tree.GetHierarchyLevel.side_effect = lambda k: 0

        return mock_tree

    def test_basic_substring_match(self):
        """Search finds nodes whose text contains the search string."""
        controller = self._make_controller_with_session()
        mock_tree = self._build_tree([
            ("R", "Root", None),
            ("A", "Documents", "R"),
            ("B", "Warehouse Tasks", "A"),
            ("C", "Other Stuff", "R"),
        ])
        controller._session.findById.return_value = mock_tree

        result = controller.search_tree_nodes("tree1", "Warehouse")

        assert result["total_matches"] == 1
        assert result["matches"][0]["key"] == "B"
        assert result["matches"][0]["text"] == "Warehouse Tasks"

    def test_case_insensitive(self):
        """Search is case-insensitive."""
        controller = self._make_controller_with_session()
        mock_tree = self._build_tree([
            ("A", "Warehouse Tasks", None),
        ])
        controller._session.findById.return_value = mock_tree

        result = controller.search_tree_nodes("tree1", "warehouse tasks")

        assert result["total_matches"] == 1
        assert result["matches"][0]["key"] == "A"

    def test_ancestor_path_construction(self):
        """Full ancestor path is built by walking up the parent chain."""
        controller = self._make_controller_with_session()
        mock_tree = self._build_tree([
            ("R", "Root", None),
            ("OB", "Outbound", "R"),
            ("DOC", "Documents", "OB"),
            ("WT", "Warehouse Tasks", "DOC"),
        ])
        controller._session.findById.return_value = mock_tree

        result = controller.search_tree_nodes("tree1", "Warehouse Tasks")

        assert result["total_matches"] == 1
        assert result["matches"][0]["path"] == "Root > Outbound > Documents > Warehouse Tasks"

    def test_multiple_matches_disambiguation(self):
        """Same label in different branches returns both with distinct paths."""
        controller = self._make_controller_with_session()
        mock_tree = self._build_tree([
            ("R", "Root", None),
            ("OB", "Outbound", "R"),
            ("OB_D", "Documents", "OB"),
            ("OB_WT", "Warehouse Tasks", "OB_D"),
            ("IB", "Inbound", "R"),
            ("IB_DL", "Delivery", "IB"),
            ("IB_D", "Documents", "IB_DL"),
            ("IB_WT", "Warehouse Tasks", "IB_D"),
        ])
        controller._session.findById.return_value = mock_tree

        result = controller.search_tree_nodes("tree1", "Warehouse Tasks")

        assert result["total_matches"] == 2
        paths = [m["path"] for m in result["matches"]]
        assert "Root > Outbound > Documents > Warehouse Tasks" in paths
        assert "Root > Inbound > Delivery > Documents > Warehouse Tasks" in paths

    def test_column_search(self):
        """When column is specified, only that column is searched."""
        controller = self._make_controller_with_session()

        mock_tree = MagicMock()
        mock_tree.GetTreeType.return_value = 1  # List tree
        col_names = self._make_gui_collection(["HIER", "STATUS"])
        mock_tree.GetColumnNames.return_value = col_names
        mock_tree.GetColumnTitleFromName.side_effect = lambda n: n

        keys = self._make_gui_collection(["K1", "K2"])
        mock_tree.GetAllNodeKeys.return_value = keys
        mock_tree.GetNodeTextByKey.return_value = ""

        def get_item_text(key, col):
            data = {
                ("K1", "HIER"): "Node A",
                ("K1", "STATUS"): "Active",
                ("K2", "HIER"): "Node B",
                ("K2", "STATUS"): "Inactive",
            }
            return data.get((key, col), "")
        mock_tree.GetItemText.side_effect = get_item_text
        mock_tree.GetParent.return_value = None
        mock_tree.IsFolderExpandable.return_value = False
        mock_tree.GetHierarchyLevel.return_value = 0

        controller._session.findById.return_value = mock_tree

        result = controller.search_tree_nodes("tree1", "Active", column="STATUS")

        # "Active" matches K1, "Inactive" also contains "Active"
        assert result["total_matches"] == 2

    def test_no_matches(self):
        """Returns empty matches list when nothing matches."""
        controller = self._make_controller_with_session()
        mock_tree = self._build_tree([
            ("A", "Something Else", None),
        ])
        controller._session.findById.return_value = mock_tree

        result = controller.search_tree_nodes("tree1", "Nonexistent")

        assert result["total_matches"] == 0
        assert result["matches"] == []

    def test_max_results_cap(self):
        """Results are capped at max_results."""
        controller = self._make_controller_with_session()
        mock_tree = self._build_tree([
            ("K1", "Task 1", None),
            ("K2", "Task 2", None),
            ("K3", "Task 3", None),
        ])
        controller._session.findById.return_value = mock_tree

        result = controller.search_tree_nodes("tree1", "Task", max_results=2)

        assert result["total_matches"] == 2


class TestGetTreeNodeChildren:
    """Tests for get_tree_node_children."""

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

    def _build_tree(self, node_defs):
        """Build a mock tree from (key, text, parent_key) tuples."""
        mock_tree = MagicMock()
        mock_tree.GetTreeType.return_value = 0

        keys = [n[0] for n in node_defs]
        mock_tree.GetAllNodeKeys.return_value = self._make_gui_collection(keys)

        text_map = {n[0]: n[1] for n in node_defs}
        parent_map = {n[0]: n[2] for n in node_defs}

        mock_tree.GetNodeTextByKey.side_effect = lambda k: text_map.get(k, "")
        mock_tree.GetParent.side_effect = lambda k: parent_map.get(k)
        mock_tree.IsFolderExpandable.return_value = True
        mock_tree.IsFolderExpanded.return_value = False
        mock_tree.GetNodeChildrenCount.side_effect = lambda k: sum(
            1 for n in node_defs if n[2] == k
        )

        return mock_tree

    def test_root_children(self):
        """Empty node_key returns root-level nodes only."""
        controller = self._make_controller_with_session()
        mock_tree = self._build_tree([
            ("R1", "Root A", None),
            ("R2", "Root B", None),
            ("C1", "Child of A", "R1"),
        ])
        controller._session.findById.return_value = mock_tree

        result = controller.get_tree_node_children("tree1")

        assert result["children_count"] == 2
        keys = [c["key"] for c in result["children"]]
        assert "R1" in keys
        assert "R2" in keys
        assert "C1" not in keys

    def test_specific_node_children(self):
        """Returns only direct children of the specified node."""
        controller = self._make_controller_with_session()
        mock_tree = self._build_tree([
            ("R", "Root", None),
            ("A", "Child A", "R"),
            ("B", "Child B", "R"),
            ("A1", "Grandchild", "A"),
        ])
        controller._session.findById.return_value = mock_tree

        result = controller.get_tree_node_children("tree1", node_key="R")

        assert result["children_count"] == 2
        keys = [c["key"] for c in result["children"]]
        assert "A" in keys
        assert "B" in keys
        assert "A1" not in keys

    def test_expand_flag_calls_expand_node(self):
        """When expand=True, ExpandNode is called on the parent."""
        controller = self._make_controller_with_session()
        mock_tree = self._build_tree([
            ("R", "Root", None),
            ("C", "Child", "R"),
        ])
        controller._session.findById.return_value = mock_tree
        controller.get_screen_info = MagicMock(return_value={"transaction": "SPRO"})

        result = controller.get_tree_node_children("tree1", node_key="R", expand=True)

        mock_tree.ExpandNode.assert_called_once_with("R")
        assert "screen" in result

    def test_no_expand_no_screen_info(self):
        """When expand=False, no screen info is returned."""
        controller = self._make_controller_with_session()
        mock_tree = self._build_tree([
            ("R", "Root", None),
        ])
        controller._session.findById.return_value = mock_tree

        result = controller.get_tree_node_children("tree1")

        assert "screen" not in result

    def test_includes_parent_path(self):
        """When node_key given, result includes parent text and ancestor path."""
        controller = self._make_controller_with_session()
        mock_tree = self._build_tree([
            ("R", "Root", None),
            ("A", "EWM", "R"),
            ("B", "Master Data", "A"),
        ])
        controller._session.findById.return_value = mock_tree

        result = controller.get_tree_node_children("tree1", node_key="A")

        assert result["node_text"] == "EWM"
        assert result["path"] == "Root > EWM"
        assert result["children_count"] == 1
        assert result["children"][0]["key"] == "B"

    def test_leaf_node_returns_empty(self):
        """A leaf node with no children returns empty list."""
        controller = self._make_controller_with_session()
        mock_tree = self._build_tree([
            ("R", "Root", None),
            ("L", "Leaf", "R"),
        ])
        controller._session.findById.return_value = mock_tree

        result = controller.get_tree_node_children("tree1", node_key="L")

        assert result["children_count"] == 0
        assert result["children"] == []

    def test_expand_failure_surfaced(self):
        """When expand fails, the error is included in the response."""
        controller = self._make_controller_with_session()
        mock_tree = self._build_tree([
            ("R", "Root", None),
        ])
        mock_tree.ExpandNode.side_effect = Exception("Node not expandable")
        controller._session.findById.return_value = mock_tree
        controller.get_screen_info = MagicMock(return_value={"transaction": "SPRO"})

        result = controller.get_tree_node_children("tree1", node_key="R", expand=True)

        assert "expand_error" in result
        assert "not expandable" in result["expand_error"]


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

    def test_read_table_alv_includes_absolute_row_index(self):
        """ALV read_table rows include _absolute_row_index matching the loop index."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        mock_grid.ColumnCount = 1
        mock_grid.ColumnOrder.return_value = "MATNR"
        mock_grid.GetColumnTooltip.return_value = ""
        mock_grid.GetDisplayedColumnTitle.return_value = "Material"
        mock_grid.RowCount = 3
        mock_grid.GetCellValue.side_effect = lambda r, c: f"MAT{r:03d}"
        controller._session.findById.return_value = mock_grid

        result = controller.read_table("wnd[0]/usr/grid", max_rows=10)

        assert len(result["data"]) == 3
        for i, row in enumerate(result["data"]):
            assert row["_absolute_row_index"] == i

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

    def test_select_alv_context_menu_item_select_by_text(self):
        """select_by='text' uses SelectContextMenuItemByText."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        controller._session.findById.return_value = mock_grid
        controller.get_screen_info = MagicMock(return_value={"transaction": "/SCWM/MON"})

        result = controller.select_alv_context_menu_item(
            "wnd[0]/usr/grid",
            "Confirm WT in Foreground",
            select_by="text",
        )

        assert result["status"] == "selected"
        mock_grid.SelectContextMenuItemByText.assert_called_once_with("Confirm WT in Foreground")

    def test_select_alv_context_menu_item_select_by_position(self):
        """select_by='position' uses SelectContextMenuItemByPosition."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        controller._session.findById.return_value = mock_grid
        controller.get_screen_info = MagicMock(return_value={"transaction": "/SCWM/MON"})

        result = controller.select_alv_context_menu_item(
            "wnd[0]/usr/grid",
            "2\\1",
            select_by="position",
        )

        assert result["status"] == "selected"
        mock_grid.SelectContextMenuItemByPosition.assert_called_once_with("2\\1")

    def test_select_alv_context_menu_item_select_by_id_toolbar_menu(self):
        """select_by='id' uses SelectToolbarMenuItem for toolbar menus."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        controller._session.findById.return_value = mock_grid
        controller.get_screen_info = MagicMock(return_value={"transaction": "/SCWM/MON"})

        result = controller.select_alv_context_menu_item(
            "wnd[0]/usr/grid",
            "@M00006",
            toolbar_button_id="&MB_ACTIONS",
            select_by="id",
        )

        assert result["status"] == "selected"
        mock_grid.PressToolbarContextButton.assert_called_once_with("&MB_ACTIONS")
        mock_grid.SelectToolbarMenuItem.assert_called_once_with("@M00006")

    def test_select_alv_context_menu_item_select_by_auto_space_heuristic(self):
        """select_by='auto' keeps space/text and function-code fallback behavior."""
        controller = self._make_controller_with_session()
        controller.get_screen_info = MagicMock(return_value={"transaction": "/SCWM/MON"})

        # Space in menu_item_id -> treat as visible text directly.
        text_grid = MagicMock()
        controller._session.findById.return_value = text_grid
        text_result = controller.select_alv_context_menu_item(
            "wnd[0]/usr/grid",
            "Confirm WT in Foreground",
            select_by="auto",
        )

        assert text_result["status"] == "selected"
        text_grid.SelectContextMenuItemByText.assert_called_once_with("Confirm WT in Foreground")

        # No spaces -> try function-code APIs first, then fallback to text.
        id_grid = MagicMock()
        id_grid.SelectContextMenuItem.side_effect = Exception("Not a context-menu ID")
        id_grid.SelectToolbarMenuItem.side_effect = Exception("Not a toolbar-menu ID")
        controller._session.findById.return_value = id_grid
        id_result = controller.select_alv_context_menu_item(
            "wnd[0]/usr/grid",
            "@M00006",
            select_by="auto",
        )

        assert id_result["status"] == "selected"
        id_grid.SelectContextMenuItem.assert_called_once_with("@M00006")
        id_grid.SelectToolbarMenuItem.assert_called_once_with("@M00006")
        id_grid.SelectContextMenuItemByText.assert_called_once_with("@M00006")

    @pytest.mark.parametrize(
        ("select_by", "menu_item_id", "expected_method"),
        [
            ("auto", "@M00006", "SelectToolbarMenuItem"),
            ("id", "@M00006", "SelectToolbarMenuItem"),
            ("text", "Confirm WT in Foreground", "SelectContextMenuItemByText"),
            ("position", "2\\1", "SelectContextMenuItemByPosition"),
        ],
    )
    def test_select_alv_context_menu_item_toolbar_button_all_modes(
        self,
        select_by,
        menu_item_id,
        expected_method,
    ):
        """toolbar_button_id should be honored for every select_by mode."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        controller._session.findById.return_value = mock_grid
        controller.get_screen_info = MagicMock(return_value={"transaction": "/SCWM/MON"})

        result = controller.select_alv_context_menu_item(
            "wnd[0]/usr/grid",
            menu_item_id,
            toolbar_button_id="&MB_ACTIONS",
            select_by=select_by,
        )

        assert result["status"] == "selected"
        mock_grid.PressToolbarContextButton.assert_called_once_with("&MB_ACTIONS")
        getattr(mock_grid, expected_method).assert_called_once_with(menu_item_id)


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
        controller._session.findById.side_effect = (
            lambda id: mock_okcd if "okcd" in id else mock_window
        )
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
        controller._session.findById.side_effect = (
            lambda id: mock_okcd if "okcd" in id else mock_window
        )
        controller.get_screen_info = MagicMock(return_value={"transaction": "MM03"})

        controller.execute_transaction("/oMM03")

        # StartTransaction should NOT be called for /o prefix
        controller._session.StartTransaction.assert_not_called()
        assert mock_okcd.text == "/oMM03"

    def test_o_prefix_rebinds_to_new_active_session(self):
        """After /o, the controller should target the newly opened session."""
        controller = self._make_controller_with_session()
        old_session = controller._session
        old_session.Id = "ses[0]"

        new_session = MagicMock()
        new_session.Id = "ses[1]"

        mock_connection = MagicMock()
        mock_connection.Children.Count = 2
        mock_connection.Children.side_effect = lambda i: [old_session, new_session][i]
        controller._connection = mock_connection

        mock_app = MagicMock()
        mock_app.ActiveSession = new_session
        mock_app.Children.Count = 1
        mock_app.Children.side_effect = lambda i: [mock_connection][i]
        controller._application = mock_app

        mock_okcd = MagicMock()
        mock_window = MagicMock()
        old_session.findById.side_effect = (
            lambda id: mock_okcd if "okcd" in id else mock_window
        )
        controller.get_screen_info = MagicMock(return_value={"transaction": "MM03"})

        controller.execute_transaction("/oMM03")

        assert controller._session is new_session
        assert controller._connection is mock_connection

    def test_o_prefix_falls_back_to_latest_connection_child(self):
        """If ActiveSession is unavailable, /o should fall back to the new child session."""
        controller = self._make_controller_with_session()
        old_session = controller._session
        old_session.Id = "ses[0]"

        new_session = MagicMock()
        new_session.Id = "ses[1]"

        mock_connection = MagicMock()
        child_count = {"value": 1}

        def connection_child(index):
            return [old_session, new_session][index]

        type(mock_connection.Children).Count = property(lambda self: child_count["value"])
        mock_connection.Children.side_effect = connection_child
        controller._connection = mock_connection

        mock_app = MagicMock()
        mock_app.Children.Count = 1
        mock_app.Children.side_effect = lambda i: [mock_connection][i]
        type(mock_app).ActiveSession = property(
            lambda self: (_ for _ in ()).throw(Exception("Not supported"))
        )
        controller._application = mock_app

        mock_okcd = MagicMock()
        mock_window = MagicMock()

        def find_by_id(element_id):
            if "okcd" in element_id:
                return mock_okcd
            child_count["value"] = 2
            return mock_window

        old_session.findById.side_effect = find_by_id
        controller.get_screen_info = MagicMock(return_value={"transaction": "MM03"})

        controller.execute_transaction("/oMM03")

        assert controller._session is new_session

    def test_uppercase_n_prefix_stripped(self):
        """Uppercase /N prefix is stripped from the returned transaction."""
        controller = self._make_controller_with_session()
        controller.get_screen_info = MagicMock(return_value={"transaction": "MM03"})

        result = controller.execute_transaction("/NMM03")
        assert result["transaction"] == "MM03"

    def test_uppercase_o_prefix_stripped(self):
        """Uppercase /O prefix is stripped from the returned transaction."""
        controller = self._make_controller_with_session()
        mock_okcd = MagicMock()
        mock_window = MagicMock()
        controller._session.findById.side_effect = (
            lambda id: mock_okcd if "okcd" in id else mock_window
        )
        controller.get_screen_info = MagicMock(return_value={"transaction": "MM03"})

        result = controller.execute_transaction("/OMM03")
        assert result["transaction"] == "MM03"

    def test_star_prefix_stripped(self):
        """/* prefix is stripped from the returned transaction."""
        controller = self._make_controller_with_session()
        controller.get_screen_info = MagicMock(
            return_value={"transaction": "MM03"}
        )

        result = controller.execute_transaction("/*MM03")
        assert result["transaction"] == "MM03"


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

    def test_normalizes_full_active_window_id(self):
        """Full ActiveWindow IDs should be normalized to wnd[N]."""
        controller = self._make_controller_with_session()
        mock_window = MagicMock()
        mock_window.Id = "/app/con[0]/ses[0]/wnd[2]"
        controller._session.ActiveWindow = mock_window

        result = controller._find_topmost_window()

        assert result == "wnd[2]"

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

    def test_take_screenshot_uses_normalized_active_window_id(self):
        """take_screenshot should use the normalized wnd[N] form with findById()."""
        controller = self._make_controller_with_session()
        mock_window = MagicMock()
        mock_window.Id = "/app/con[0]/ses[0]/wnd[2]"
        controller._session.ActiveWindow = mock_window

        hardcopy_target = MagicMock()
        controller._session.findById.return_value = hardcopy_target
        controller._optimize_screenshot = MagicMock()

        result = controller.take_screenshot("test.png")

        controller._session.findById.assert_called_once_with("wnd[2]")
        hardcopy_target.HardCopy.assert_called_once_with("test.png", "PNG")
        assert result["window"] == "wnd[2]"


class TestDisconnectOwnership:
    """Tests for session ownership during disconnect."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        controller._connection = MagicMock()
        return controller

    def test_disconnect_closes_owned_session(self):
        """Sessions opened by the controller should be closed on disconnect."""
        controller = self._make_controller_with_session()
        controller._session.Id = "ses[0]"
        controller._owns_session = True
        connection = controller._connection

        controller.disconnect()

        connection.CloseSession.assert_called_once_with("ses[0]")
        assert controller._session is None
        assert controller._connection is None
        assert controller._owns_session is False

    def test_disconnect_does_not_close_attached_session(self):
        """Sessions attached via connect_existing should be left open on disconnect."""
        controller = self._make_controller_with_session()
        controller._session.Id = "ses[0]"
        controller._owns_session = False
        connection = controller._connection

        controller.disconnect()

        connection.CloseSession.assert_not_called()
        assert controller._session is None
        assert controller._connection is None
        assert controller._owns_session is False

    def test_connect_existing_marks_session_unowned(self):
        """connect_to_existing_session should mark the session as attached, not owned."""
        from mcp_sap_gui.sap_controller import SAPGUIController

        controller = SAPGUIController()
        mock_session = MagicMock()
        mock_connection = MagicMock()
        mock_connection.Children.Count = 1
        mock_connection.Children.side_effect = lambda i: [mock_session][i]

        mock_app = MagicMock()
        mock_app.Children.Count = 1
        mock_app.Children.side_effect = lambda i: [mock_connection][i]

        controller._application = mock_app
        controller.get_session_info = MagicMock(return_value={"system_name": "DEV"})

        controller.connect_to_existing_session()

        assert controller._owns_session is False


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
                            visible_rows=None, initial_scroll_position=0):
        """Create a mock GuiTableControl with columns and data.

        Args:
            columns: list of dicts with 'name', 'title', 'tooltip' keys
            rows_data: list of lists (outer=rows, inner=cell values)
            total_rows: total row count (default len(rows_data))
            visible_rows: visible rows at once (default len(rows_data))
            initial_scroll_position: starting scroll position (default 0)
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
        scrollbar.Position = initial_scroll_position
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
        assert result["first_visible_row"] == 0
        assert result["rows_returned"] == 2
        assert result["columns"] == ["LGNUM", "LGTYPGRP"]
        assert len(result["column_info"]) == 2
        assert result["column_info"][0]["title"] == "WhN"
        assert result["column_info"][0]["tooltip"] == "Warehouse Number"
        assert result["data"][0]["LGNUM"] == "WH01"
        assert result["data"][1]["LGTYPGRP"] == "GRP2"

    def test_read_table_control_capped_at_visible_rows(self):
        """read_table returns at most VisibleRowCount rows (no scrolling)."""
        controller = self._make_controller_with_session()
        columns = [{"name": "COL_A", "title": "A"}]
        # 10 rows total but only 3 visible at a time
        rows = [[f"val_{i}"] for i in range(10)]
        mock_table = self._make_table_control(columns, rows,
                                               total_rows=10, visible_rows=3)
        controller._session.findById.return_value = mock_table

        result = controller.read_table("wnd[0]/usr/tblTEST", max_rows=100)

        assert result["total_rows"] == 10
        assert result["first_visible_row"] == 0
        assert result["visible_rows"] == 3
        # Capped at visible rows - no scrolling
        assert result["rows_returned"] == 3
        assert result["data"][0]["COL_A"] == "val_0"
        assert result["data"][2]["COL_A"] == "val_2"

    def test_read_table_control_respects_max_rows(self):
        """read_table stops after max_rows even with more visible data."""
        controller = self._make_controller_with_session()
        columns = [{"name": "COL", "title": "Col"}]
        rows = [[f"v{i}"] for i in range(20)]
        mock_table = self._make_table_control(columns, rows,
                                               total_rows=20, visible_rows=10)
        controller._session.findById.return_value = mock_table

        result = controller.read_table("wnd[0]/usr/tblTEST", max_rows=5)

        assert result["rows_returned"] == 5
        assert result["first_visible_row"] == 0
        assert result["total_rows"] == 20

    def test_read_table_control_from_scrolled_position(self):
        """Reading from a scrolled position returns data from that position onward."""
        controller = self._make_controller_with_session()
        columns = [{"name": "LGNUM", "title": "WhN"}]
        # 1705 rows total, 20 visible, user navigated to position 500
        rows = [[f"WH{i:04d}"] for i in range(1705)]
        mock_table = self._make_table_control(
            columns, rows, total_rows=1705, visible_rows=20,
            initial_scroll_position=500,
        )
        controller._session.findById.return_value = mock_table

        result = controller.read_table("wnd[0]/usr/tblTEST", max_rows=100)

        assert result["first_visible_row"] == 500
        # Capped at 20 visible rows (no scrolling)
        assert result["rows_returned"] == 20
        assert result["total_rows"] == 1705
        # Data starts from row 500, not row 0
        assert result["data"][0]["LGNUM"] == "WH0500"
        assert result["data"][19]["LGNUM"] == "WH0519"

    def test_read_table_control_absolute_row_index(self):
        """Each row includes _absolute_row_index = start_position + vis_idx."""
        controller = self._make_controller_with_session()
        columns = [{"name": "COL", "title": "Col"}]
        rows = [[f"v{i}"] for i in range(50)]
        mock_table = self._make_table_control(
            columns, rows, total_rows=50, visible_rows=5,
            initial_scroll_position=10,
        )
        controller._session.findById.return_value = mock_table

        result = controller.read_table("wnd[0]/usr/tblTEST", max_rows=100)

        assert result["rows_returned"] == 5
        for i, row in enumerate(result["data"]):
            assert row["_absolute_row_index"] == 10 + i

    def test_read_table_control_near_end_with_padding(self):
        """Reading near end excludes padding rows from the visible window.

        RowCount includes padding rows (empty rows filling the visible area).
        At scroll position 95 in a 100-entry table with 10 visible rows,
        only 5 rows have real data — the other 5 visible rows are padding.
        """
        controller = self._make_controller_with_session()
        columns = [{"name": "COL", "title": "Col"}]
        real_entries = 100
        visible = 10
        padded_total = real_entries + visible - 1  # 109
        rows = [[f"v{i}"] for i in range(real_entries)]
        mock_table = self._make_table_control(
            columns, rows, total_rows=padded_total, visible_rows=visible,
            initial_scroll_position=95,
        )
        controller._session.findById.return_value = mock_table

        result = controller.read_table("wnd[0]/usr/tblTEST", max_rows=100)

        assert result["first_visible_row"] == 95
        assert result["rows_returned"] == 5  # only rows 95-99 are real
        assert result["data"][0]["COL"] == "v95"
        assert result["data"][4]["COL"] == "v99"

    def test_read_table_control_padding_detection(self):
        """Padding rows (empty rows beyond real data) are excluded from results.

        This simulates a table with 57 real entries but RowCount=73 (16
        padding rows).  After Position... to entry 52 (position 51), 17 rows
        are visible but only 6 contain data — the rest are padding.
        """
        controller = self._make_controller_with_session()
        columns = [{"name": "LGNUM", "title": "WhN"},
                    {"name": "LNUMT", "title": "Desc"}]
        real_entries = 57
        visible = 17
        padded_total = real_entries + visible - 1  # 73
        rows = [[f"WH{i:03d}", f"Warehouse {i}"] for i in range(real_entries)]
        mock_table = self._make_table_control(
            columns, rows, total_rows=padded_total, visible_rows=visible,
            initial_scroll_position=51,
        )
        controller._session.findById.return_value = mock_table

        result = controller.read_table("wnd[0]/usr/tblTEST", max_rows=100)

        assert result["first_visible_row"] == 51
        # Only 6 real rows visible (51-56), padding rows excluded
        assert result["rows_returned"] == 6
        assert result["data"][0]["LGNUM"] == "WH051"
        assert result["data"][5]["LGNUM"] == "WH056"

    def test_read_table_control_max_rows_from_scrolled_position(self):
        """max_rows cap still applies when reading from a scrolled position."""
        controller = self._make_controller_with_session()
        columns = [{"name": "COL", "title": "Col"}]
        rows = [[f"v{i}"] for i in range(500)]
        mock_table = self._make_table_control(
            columns, rows, total_rows=500, visible_rows=20,
            initial_scroll_position=200,
        )
        controller._session.findById.return_value = mock_table

        result = controller.read_table("wnd[0]/usr/tblTEST", max_rows=10)

        assert result["first_visible_row"] == 200
        assert result["rows_returned"] == 10
        assert result["data"][0]["COL"] == "v200"
        assert result["data"][9]["COL"] == "v209"

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
        """select_table_row uses GetAbsoluteRow().Selected for GuiTableControl."""
        controller = self._make_controller_with_session()
        columns = [{"name": "COL", "title": "Col"}]
        rows = [[f"v{i}"] for i in range(5)]
        mock_table = self._make_table_control(columns, rows, visible_rows=5)
        controller._session.findById.return_value = mock_table

        result = controller.select_table_row("wnd[0]/usr/tblTEST", 2)

        assert result["status"] == "success"
        assert result["selected_row"] == 2
        # Verify GetAbsoluteRow was called with absolute row index
        mock_table.GetAbsoluteRow.assert_called_with(2)
        assert mock_table.GetAbsoluteRow(2).Selected is True

    def test_double_click_table_cell_guitablecontrol(self):
        """double_click_table_cell uses SetFocus + F2 for GuiTableControl."""
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

    # --- scroll + interaction tests ---

    def test_scroll_to_row_already_visible(self):
        """_scroll_table_control_to_row returns visible offset when row is already visible."""
        controller = self._make_controller_with_session()
        columns = [{"name": "COL", "title": "Col"}]
        rows = [[f"v{i}"] for i in range(50)]
        mock_table = self._make_table_control(
            columns, rows, total_rows=50, visible_rows=10,
            initial_scroll_position=20,
        )

        vis_offset = controller._scroll_table_control_to_row(mock_table, 25)

        assert vis_offset == 5  # 25 - 20
        # Scrollbar should NOT have been changed
        assert mock_table.VerticalScrollbar.Position == 20

    def test_scroll_to_row_needs_scrolling(self):
        """_scroll_table_control_to_row scrolls and returns visible offset."""
        controller = self._make_controller_with_session()
        columns = [{"name": "COL", "title": "Col"}]
        rows = [[f"v{i}"] for i in range(100)]
        mock_table = self._make_table_control(
            columns, rows, total_rows=100, visible_rows=10,
            initial_scroll_position=0,
        )

        vis_offset = controller._scroll_table_control_to_row(mock_table, 50)

        assert vis_offset == 0  # 50 - 50 (scrollbar set to 50)
        assert mock_table.VerticalScrollbar.Position == 50

    def test_select_table_row_from_scrolled_position(self):
        """select_table_row scrolls and selects the correct row."""
        controller = self._make_controller_with_session()
        columns = [{"name": "FIELD", "title": "Field"}]
        rows = [[f"row_{i}"] for i in range(100)]
        mock_table = self._make_table_control(
            columns, rows, total_rows=100, visible_rows=10,
            initial_scroll_position=60,
        )
        controller._session.findById.return_value = mock_table

        result = controller.select_table_row("wnd[0]/usr/tblTEST", 75)

        assert result["status"] == "success"
        assert result["selected_row"] == 75
        # Row 75 is at scroll position 60, visible offset = 75 - 60 = 15,
        # but 15 >= 10 (visible), so it scrolls to 75, offset = 0
        assert mock_table.VerticalScrollbar.Position == 75
        # GetAbsoluteRow uses absolute row index
        mock_table.GetAbsoluteRow.assert_called_with(75)
        assert mock_table.GetAbsoluteRow(75).Selected is True

    def test_select_table_row_succeeds_when_scroll_fails(self):
        """select_table_row uses GetAbsoluteRow even if scroll throws COM error."""
        controller = self._make_controller_with_session()
        columns = [{"name": "FIELD", "title": "Field"}]
        rows = [[f"row_{i}"] for i in range(84)]
        mock_table = self._make_table_control(
            columns, rows, total_rows=84, visible_rows=9,
            initial_scroll_position=0,
        )
        # Make scrollbar.Position setter throw (simulates popup table COM error)
        type(mock_table.VerticalScrollbar).Position = property(
            lambda self: 0,
            lambda self, v: (_ for _ in ()).throw(
                Exception("(-2147417851, 'The server threw an exception.')")
            ),
        )
        controller._session.findById.return_value = mock_table

        result = controller.select_table_row("wnd[1]/usr/tblSEL_FLDS", 75)

        assert result["status"] == "success"
        assert result["selected_row"] == 75
        mock_table.GetAbsoluteRow.assert_called_with(75)
        assert mock_table.GetAbsoluteRow(75).Selected is True

    def test_double_click_cell_from_scrolled_position(self):
        """double_click_table_cell scrolls and focuses the correct cell."""
        controller = self._make_controller_with_session()
        columns = [{"name": "COL_A", "title": "A"},
                    {"name": "COL_B", "title": "B"}]
        rows = [[f"a{i}", f"b{i}"] for i in range(80)]
        mock_table = self._make_table_control(
            columns, rows, total_rows=80, visible_rows=10,
            initial_scroll_position=0,
        )
        controller._session.findById.return_value = mock_table
        controller.get_screen_info = MagicMock(return_value={"transaction": "SM30"})

        result = controller.double_click_table_cell("wnd[0]/usr/tblTEST", 40, "COL_B")

        assert result["status"] == "double_clicked"
        # Scrolled to position 40, visible offset = 0, column COL_B = index 1
        assert mock_table.VerticalScrollbar.Position == 40
        mock_table.GetCell.assert_called_with(0, 1)


class TestEnrichedStatusBar:
    """Tests for enriched status bar in get_screen_info."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_status_bar_returns_message_type(self):
        """get_screen_info includes message_type from status bar."""
        controller = self._make_controller_with_session()
        mock_window = MagicMock()
        mock_window.Text = "Display Material"
        mock_info = MagicMock()
        mock_info.Transaction = "MM03"
        mock_info.Program = "SAPLMGMM"
        mock_info.ScreenNumber = 100
        controller._session.Info = mock_info

        mock_sbar = MagicMock()
        mock_sbar.Text = "Material 100 displayed"
        mock_sbar.MessageType = "S"
        mock_sbar.MessageId = "MM"
        mock_sbar.MessageNumber = "050"

        def find_by_id(id):
            if id == "wnd[0]/sbar":
                return mock_sbar
            return mock_window
        controller._session.findById.side_effect = find_by_id

        result = controller.get_screen_info()

        assert result["message"] == "Material 100 displayed"
        assert result["message_type"] == "S"
        assert result["message_id"] == "MM"
        assert result["message_number"] == "050"

    def test_status_bar_info_with_parameters(self):
        """_get_status_bar_info includes message parameters."""
        controller = self._make_controller_with_session()
        mock_sbar = MagicMock()
        mock_sbar.Text = "Error in field"
        mock_sbar.MessageType = "E"
        mock_sbar.MessageId = "00"
        mock_sbar.MessageNumber = "001"
        mock_sbar.MessageParameter = "MATNR"
        mock_sbar.MessageParameter1 = "100"
        controller._session.findById.return_value = mock_sbar

        result = controller._get_status_bar_info()

        assert result["text"] == "Error in field"
        assert result["message_type"] == "E"
        assert "message_parameters" in result
        assert "MATNR" in result["message_parameters"]


class TestReadFieldMetadata:
    """Tests for enriched read_field metadata."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_read_field_includes_required(self):
        """read_field returns required flag for text fields."""
        controller = self._make_controller_with_session()
        mock_field = MagicMock()
        mock_field.Text = "100"
        mock_field.Type = "GuiTextField"
        mock_field.Name = "MATNR"
        mock_field.Changeable = True
        mock_field.Required = True
        mock_field.MaxLength = 18
        mock_field.Numerical = False
        mock_field.Highlighted = False
        mock_label = MagicMock()
        mock_label.Text = "Material"
        mock_field.LeftLabel = mock_label
        controller._session.findById.return_value = mock_field

        result = controller.read_field("wnd[0]/usr/txtMATNR")

        assert result["value"] == "100"
        assert result["required"] is True
        assert result["max_length"] == 18
        assert result["numerical"] is False
        assert result["left_label"] == "Material"

    def test_read_field_without_metadata(self):
        """read_field works for elements without extended metadata."""
        controller = self._make_controller_with_session()
        mock_field = MagicMock(spec=['Text', 'Type', 'Name', 'Changeable'])
        mock_field.Text = "Hello"
        mock_field.Type = "GuiLabel"
        mock_field.Name = "LBL1"
        mock_field.Changeable = False
        controller._session.findById.return_value = mock_field

        result = controller.read_field("wnd[0]/usr/lblLBL1")

        assert result["value"] == "Hello"
        assert result["type"] == "GuiLabel"
        # No required/max_length for labels
        assert "required" not in result
        assert "max_length" not in result


class TestGetComboboxEntries:
    """Tests for get_combobox_entries."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_returns_all_entries(self):
        """get_combobox_entries returns all key-value pairs."""
        controller = self._make_controller_with_session()
        mock_combo = MagicMock()
        mock_combo.Key = "EN"

        entry1 = MagicMock()
        entry1.Key = "EN"
        entry1.Value = "English"
        entry2 = MagicMock()
        entry2.Key = "DE"
        entry2.Value = "German"

        mock_entries = MagicMock()
        mock_entries.Count = 2
        mock_entries.side_effect = lambda i: [entry1, entry2][i]
        mock_combo.Entries = mock_entries

        controller._session.findById.return_value = mock_combo

        result = controller.get_combobox_entries("wnd[0]/usr/cmbLANGU")

        assert result["entry_count"] == 2
        assert result["current_key"] == "EN"
        assert result["entries"][0]["key"] == "EN"
        assert result["entries"][0]["value"] == "English"
        assert result["entries"][1]["key"] == "DE"
        assert result["entries"][1]["value"] == "German"


class TestSetBatchFields:
    """Tests for set_batch_fields."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_sets_multiple_fields(self):
        """set_batch_fields sets all provided fields."""
        controller = self._make_controller_with_session()
        mock_field1 = MagicMock()
        mock_field2 = MagicMock()

        def find_by_id(id):
            return {"wnd[0]/usr/txtF1": mock_field1, "wnd[0]/usr/txtF2": mock_field2}[id]
        controller._session.findById.side_effect = find_by_id

        result = controller.set_batch_fields({
            "wnd[0]/usr/txtF1": "val1",
            "wnd[0]/usr/txtF2": "val2",
        })

        assert result["total"] == 2
        assert result["succeeded"] == 2
        assert result["failed"] == 0

    def test_partial_failure(self):
        """set_batch_fields reports partial failures."""
        controller = self._make_controller_with_session()
        mock_field1 = MagicMock()

        def find_by_id(id):
            if id == "wnd[0]/usr/txtBAD":
                raise Exception("Not found")
            return mock_field1
        controller._session.findById.side_effect = find_by_id

        result = controller.set_batch_fields({
            "wnd[0]/usr/txtF1": "val1",
            "wnd[0]/usr/txtBAD": "val2",
        })

        assert result["succeeded"] == 1
        assert result["failed"] == 1
        assert result["results"]["wnd[0]/usr/txtF1"] == "success"
        assert "error" in result["results"]["wnd[0]/usr/txtBAD"]


class TestTextedit:
    """Tests for read_textedit and set_textedit."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_read_textedit(self):
        """read_textedit returns text and line_count (no lines key)."""
        controller = self._make_controller_with_session()
        mock_te = MagicMock()
        mock_te.LineCount = 3
        mock_te.GetLineText.side_effect = ["Line 1", "Line 2", "Line 3"]
        mock_te.Changeable = True
        controller._session.findById.return_value = mock_te

        result = controller.read_textedit("wnd[0]/usr/txtEdit")

        assert result["line_count"] == 3
        assert result["text"] == "Line 1\nLine 2\nLine 3"
        assert "lines" not in result
        assert result["changeable"] is True

    def test_read_textedit_max_lines_truncation(self):
        """max_lines caps the number of lines returned and adds truncated flag."""
        controller = self._make_controller_with_session()
        mock_te = MagicMock()
        mock_te.LineCount = 100
        mock_te.GetLineText.side_effect = [f"Line {i}" for i in range(100)]
        mock_te.Changeable = True
        controller._session.findById.return_value = mock_te

        result = controller.read_textedit("wnd[0]/usr/txtEdit", max_lines=5)

        assert result["line_count"] == 5
        assert result["text"] == "Line 0\nLine 1\nLine 2\nLine 3\nLine 4"
        assert result["truncated"] is True
        assert result["total_lines"] == 100

    def test_read_textedit_max_lines_zero_returns_all(self):
        """max_lines=0 (default) returns all lines."""
        controller = self._make_controller_with_session()
        mock_te = MagicMock()
        mock_te.LineCount = 3
        mock_te.GetLineText.side_effect = ["A", "B", "C"]
        mock_te.Changeable = False
        controller._session.findById.return_value = mock_te

        result = controller.read_textedit("wnd[0]/usr/txtEdit", max_lines=0)

        assert result["line_count"] == 3
        assert result["text"] == "A\nB\nC"
        assert "truncated" not in result
        assert "total_lines" not in result

    def test_set_textedit(self):
        """set_textedit sets the Text property."""
        controller = self._make_controller_with_session()
        mock_te = MagicMock()
        controller._session.findById.return_value = mock_te

        result = controller.set_textedit("wnd[0]/usr/txtEdit", "New text content")

        assert result["status"] == "success"
        assert mock_te.Text == "New text content"

    def test_set_textedit_fallback_to_unprotected(self):
        """set_textedit falls back to SetUnprotectedTextPart on Text failure."""
        controller = self._make_controller_with_session()
        mock_te = MagicMock()
        def _raise(s, v):
            raise Exception("Protected")

        type(mock_te).Text = property(lambda s: "", _raise)
        controller._session.findById.return_value = mock_te

        result = controller.set_textedit("wnd[0]/usr/txtEdit", "Fallback text")

        assert result["status"] == "success"
        mock_te.SetUnprotectedTextPart.assert_called_once_with("Fallback text")


class TestSetFocus:
    """Tests for set_focus."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_set_focus(self):
        """set_focus calls SetFocus on element."""
        controller = self._make_controller_with_session()
        mock_elem = MagicMock()
        controller._session.findById.return_value = mock_elem

        result = controller.set_focus("wnd[0]/usr/txtMATNR")

        assert result["status"] == "success"
        mock_elem.SetFocus.assert_called_once()


class TestScrollTableControl:
    """Tests for scroll_table_control."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_scroll_to_position(self):
        """scroll_table_control sets scrollbar position."""
        controller = self._make_controller_with_session()
        mock_table = MagicMock()
        mock_table.Type = "GuiTableControl"
        mock_table.VisibleRowCount = 10
        mock_table.RowCount = 100
        scrollbar = MagicMock()
        scrollbar.Minimum = 0
        scrollbar.Maximum = 90
        scrollbar.Position = 0
        mock_table.VerticalScrollbar = scrollbar
        controller._session.findById.return_value = mock_table

        result = controller.scroll_table_control("wnd[0]/usr/tblTEST", 50)

        assert result["status"] == "success"
        assert result["position"] == 50
        assert scrollbar.Position == 50

    def test_scroll_clamps_to_max(self):
        """scroll_table_control clamps position to scrollbar maximum."""
        controller = self._make_controller_with_session()
        mock_table = MagicMock()
        mock_table.Type = "GuiTableControl"
        mock_table.VisibleRowCount = 10
        mock_table.RowCount = 50
        scrollbar = MagicMock()
        scrollbar.Minimum = 0
        scrollbar.Maximum = 40
        scrollbar.Position = 0
        mock_table.VerticalScrollbar = scrollbar
        controller._session.findById.return_value = mock_table

        result = controller.scroll_table_control("wnd[0]/usr/tblTEST", 999)

        assert result["position"] == 40  # Clamped to max

    def test_scroll_rejects_non_tablecontrol(self):
        """scroll_table_control rejects non-GuiTableControl elements."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        mock_grid.Type = "GuiGridView"
        controller._session.findById.return_value = mock_grid

        result = controller.scroll_table_control("wnd[0]/usr/grid", 10)

        assert "error" in result
        assert "Not a GuiTableControl" in result["error"]

    def test_scroll_returns_diagnostic_on_com_error(self):
        """scroll_table_control returns diagnostic info when Position setter throws."""
        controller = self._make_controller_with_session()
        mock_table = MagicMock()
        mock_table.Type = "GuiTableControl"
        mock_table.RowCount = 84
        scrollbar = MagicMock()
        scrollbar.Minimum = 0
        scrollbar.Maximum = 75
        # Make Position setter throw (simulates popup table COM error)
        type(scrollbar).Position = property(
            lambda self: 0,
            lambda self, v: (_ for _ in ()).throw(
                Exception("(-2147417851, 'The server threw an exception.')")
            ),
        )
        mock_table.VerticalScrollbar = scrollbar
        controller._session.findById.return_value = mock_table

        result = controller.scroll_table_control("wnd[1]/usr/tblSEL_FLDS", 70)

        assert "error" in result
        assert "hint" in result
        assert result["clamped_position"] == 70
        assert result["scroll_max"] == 75
        assert result["total_rows"] == 84


class TestGetTableControlRowInfo:
    """Tests for get_table_control_row_info."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_returns_row_metadata(self):
        """get_table_control_row_info returns selectable/selected for each row."""
        controller = self._make_controller_with_session()
        mock_table = MagicMock()
        mock_table.Type = "GuiTableControl"
        mock_table.VisibleRowCount = 5
        scrollbar = MagicMock()
        scrollbar.Position = 0
        mock_table.VerticalScrollbar = scrollbar

        def get_abs_row(r):
            row_mock = MagicMock()
            row_mock.Selectable = True
            row_mock.Selected = (r == 2)
            return row_mock
        mock_table.GetAbsoluteRow.side_effect = get_abs_row
        controller._session.findById.return_value = mock_table

        result = controller.get_table_control_row_info("wnd[0]/usr/tblTEST", rows=[0, 1, 2])

        assert result["row_count"] == 3
        assert result["rows"][0]["selected"] is False
        assert result["rows"][2]["selected"] is True
        assert result["rows"][0]["selectable"] is True

    def test_defaults_to_visible_rows(self):
        """get_table_control_row_info queries visible rows when rows=None."""
        controller = self._make_controller_with_session()
        mock_table = MagicMock()
        mock_table.Type = "GuiTableControl"
        mock_table.VisibleRowCount = 3
        scrollbar = MagicMock()
        scrollbar.Position = 10
        mock_table.VerticalScrollbar = scrollbar

        row_mock = MagicMock()
        row_mock.Selectable = True
        row_mock.Selected = False
        mock_table.GetAbsoluteRow.return_value = row_mock
        controller._session.findById.return_value = mock_table

        result = controller.get_table_control_row_info("wnd[0]/usr/tblTEST")

        assert result["row_count"] == 3
        assert result["rows"][0]["row"] == 10
        assert result["rows"][2]["row"] == 12


class TestSelectAllTableControlColumns:
    """Tests for select_all_table_control_columns."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_select_all(self):
        """select_all_table_control_columns calls SelectAllColumns."""
        controller = self._make_controller_with_session()
        mock_table = MagicMock()
        mock_table.Type = "GuiTableControl"
        controller._session.findById.return_value = mock_table

        result = controller.select_all_table_control_columns("wnd[0]/usr/tblTEST", True)

        assert result["status"] == "all_selected"
        mock_table.SelectAllColumns.assert_called_once()

    def test_deselect_all(self):
        """select_all_table_control_columns calls DeselectAllColumns."""
        controller = self._make_controller_with_session()
        mock_table = MagicMock()
        mock_table.Type = "GuiTableControl"
        controller._session.findById.return_value = mock_table

        result = controller.select_all_table_control_columns("wnd[0]/usr/tblTEST", False)

        assert result["status"] == "all_deselected"
        mock_table.DeselectAllColumns.assert_called_once()


class TestGetCellInfo:
    """Tests for get_cell_info (ALV)."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_returns_cell_metadata(self):
        """get_cell_info returns value and metadata."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        mock_grid.GetCellValue.return_value = "MAT-001"
        mock_grid.GetCellChangeable.return_value = True
        mock_grid.GetCellColor.return_value = 0
        mock_grid.GetCellTooltip.return_value = "Material number"
        mock_grid.GetCellStyle.return_value = 0
        mock_grid.GetCellMaxLength.return_value = 18
        controller._session.findById.return_value = mock_grid

        result = controller.get_cell_info("wnd[0]/usr/grid", 0, "MATNR")

        assert result["value"] == "MAT-001"
        assert result["changeable"] is True
        assert result["tooltip"] == "Material number"
        assert result["max_length"] == 18

    def test_handles_missing_methods(self):
        """get_cell_info handles grids where some methods fail."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        mock_grid.GetCellValue.return_value = "val"
        mock_grid.GetCellChangeable.side_effect = Exception("Not supported")
        mock_grid.GetCellColor.side_effect = Exception("Not supported")
        mock_grid.GetCellTooltip.side_effect = Exception("Not supported")
        mock_grid.GetCellStyle.side_effect = Exception("Not supported")
        mock_grid.GetCellMaxLength.side_effect = Exception("Not supported")
        controller._session.findById.return_value = mock_grid

        result = controller.get_cell_info("wnd[0]/usr/grid", 0, "COL")

        assert result["value"] == "val"
        assert "changeable" not in result
        assert "error" not in result


class TestPressColumnHeader:
    """Tests for press_column_header (ALV)."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_press_column_header(self):
        """press_column_header calls PressColumnHeader."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        controller._session.findById.return_value = mock_grid
        controller.get_screen_info = MagicMock(return_value={"transaction": "MM03"})

        result = controller.press_column_header("wnd[0]/usr/grid", "MATNR")

        assert result["status"] == "pressed"
        mock_grid.PressColumnHeader.assert_called_once_with("MATNR")


class TestSelectAllRows:
    """Tests for select_all_rows (ALV)."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_select_all_rows(self):
        """select_all_rows calls SelectAll on grid."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        controller._session.findById.return_value = mock_grid

        result = controller.select_all_rows("wnd[0]/usr/grid")

        assert result["status"] == "all_selected"
        mock_grid.SelectAll.assert_called_once()


class TestGetCurrentCell:
    """Tests for get_current_cell (both table types)."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_alv_grid(self):
        """get_current_cell returns ALV CurrentCellRow/Column."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        mock_grid.Type = "GuiGridView"
        mock_grid.CurrentCellRow = 5
        mock_grid.CurrentCellColumn = "MATNR"
        controller._session.findById.return_value = mock_grid

        result = controller.get_current_cell("wnd[0]/usr/grid")

        assert result["table_type"] == "GuiGridView"
        assert result["current_row"] == 5
        assert result["current_column"] == "MATNR"

    def test_table_control(self):
        """get_current_cell returns TableControl CurrentRow/Col."""
        controller = self._make_controller_with_session()
        mock_table = MagicMock()
        mock_table.Type = "GuiTableControl"
        mock_table.CurrentRow = 3
        mock_table.CurrentCol = 2
        controller._session.findById.return_value = mock_table

        result = controller.get_current_cell("wnd[0]/usr/tblTEST")

        assert result["table_type"] == "GuiTableControl"
        assert result["current_row"] == 3
        assert result["current_col"] == 2


# ===========================================================================
# Popup Window Detection Tests
# ===========================================================================

class TestGetPopupWindow:
    """Tests for get_popup_window()."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_no_popup(self):
        """Returns popup_exists=False when no popup window."""
        controller = self._make_controller_with_session()
        # wnd[1] not found
        controller._session.findById.side_effect = Exception("not found")

        result = controller.get_popup_window()
        assert result["popup_exists"] is False

    def test_popup_found(self):
        """Returns popup info when wnd[1] exists."""
        controller = self._make_controller_with_session()

        mock_popup = MagicMock()
        mock_popup.Text = "Confirm Action"

        mock_usr = MagicMock()
        mock_usr.Children.Count = 1
        mock_label = MagicMock()
        mock_label.Type = "GuiLabel"
        mock_label.Text = "Do you want to continue?"
        mock_label.Children = MagicMock()
        mock_label.Children.Count = 0
        mock_usr.Children.side_effect = lambda i: mock_label

        mock_tbar = MagicMock()
        mock_btn = MagicMock()
        mock_btn.Type = "GuiButton"
        mock_btn.Id = "wnd[1]/tbar[0]/btn[0]"
        mock_btn.Text = "Yes"
        mock_btn.Tooltip = "Confirm"
        mock_tbar.Children.Count = 1
        mock_tbar.Children.side_effect = lambda i: mock_btn

        def find_by_id(element_id):
            if element_id == "wnd[1]":
                return mock_popup
            if element_id == "wnd[2]":
                raise Exception("not found")
            if element_id == "wnd[1]/sbar":
                raise Exception("no sbar")
            if element_id == "wnd[1]/usr":
                return mock_usr
            if element_id == "wnd[1]/tbar[0]":
                return mock_tbar
            if element_id == "wnd[1]/tbar[1]":
                raise Exception("no tbar[1]")
            raise Exception("not found")

        controller._session.findById.side_effect = find_by_id

        result = controller.get_popup_window()
        assert result["popup_exists"] is True
        assert result["title"] == "Confirm Action"
        assert "texts" in result
        assert "Do you want to continue?" in result["texts"]
        assert len(result["buttons"]) >= 1
        assert result["buttons"][0]["text"] == "Yes"


# ===========================================================================
# Toolbar Button Discovery Tests
# ===========================================================================

class TestGetToolbarButtons:
    """Tests for get_toolbar_buttons()."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_reads_toolbar_buttons(self):
        """Enumerates buttons from tbar[0] and tbar[1]."""
        controller = self._make_controller_with_session()

        # tbar[0]: one system button
        mock_sys_btn = MagicMock()
        mock_sys_btn.Type = "GuiButton"
        mock_sys_btn.Id = "wnd[0]/tbar[0]/btn[0]"
        mock_sys_btn.Text = ""
        mock_sys_btn.Tooltip = "Save (Ctrl+S)"
        mock_sys_btn.Changeable = True
        mock_sys_tbar = MagicMock()
        mock_sys_tbar.Children.Count = 1
        mock_sys_tbar.Children.side_effect = lambda i: mock_sys_btn

        # tbar[1]: one app button
        mock_app_btn = MagicMock()
        mock_app_btn.Type = "GuiButton"
        mock_app_btn.Id = "wnd[0]/tbar[1]/btn[8]"
        mock_app_btn.Text = "Execute"
        mock_app_btn.Tooltip = "Execute (F8)"
        mock_app_btn.Changeable = True
        mock_app_tbar = MagicMock()
        mock_app_tbar.Children.Count = 1
        mock_app_tbar.Children.side_effect = lambda i: mock_app_btn

        def find_by_id(element_id):
            if element_id == "wnd[0]/tbar[0]":
                return mock_sys_tbar
            if element_id == "wnd[0]/tbar[1]":
                return mock_app_tbar
            raise Exception("not found")

        controller._session.findById.side_effect = find_by_id

        result = controller.get_toolbar_buttons()
        assert "system_toolbar" in result["toolbars"]
        assert "application_toolbar" in result["toolbars"]
        assert result["toolbars"]["system_toolbar"][0]["tooltip"] == "Save (Ctrl+S)"
        assert result["toolbars"]["application_toolbar"][0]["text"] == "Execute"

    def test_empty_toolbars(self):
        """Handles missing toolbars gracefully."""
        controller = self._make_controller_with_session()
        controller._session.findById.side_effect = Exception("not found")

        result = controller.get_toolbar_buttons()
        assert result["toolbars"] == {}


# ===========================================================================
# Multi-Row Selection Tests
# ===========================================================================

class TestSelectMultipleRows:
    """Tests for select_multiple_rows()."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_alv_grid(self):
        """ALV grid gets comma-separated SelectedRows."""
        controller = self._make_controller_with_session()
        mock_grid = MagicMock()
        mock_grid.Type = "GuiGridView"
        controller._session.findById.return_value = mock_grid

        result = controller.select_multiple_rows("wnd[0]/usr/grid", [0, 3, 7])

        assert mock_grid.SelectedRows == "0,3,7"
        assert result["status"] == "success"
        assert result["count"] == 3

    def test_table_control(self):
        """TableControl iterates rows and sets Selected."""
        controller = self._make_controller_with_session()
        mock_table = MagicMock()
        mock_table.Type = "GuiTableControl"
        mock_table.VerticalScrollbar = MagicMock()
        mock_table.VerticalScrollbar.Position = 0
        mock_table.VerticalScrollbar.Maximum = 100
        mock_rows = {}
        for r in [1, 4]:
            mock_rows[r] = MagicMock()
        mock_table.GetAbsoluteRow.side_effect = lambda r: mock_rows.get(r, MagicMock())
        # Stub VisibleRowCount for _scroll_table_control_to_row
        mock_table.VisibleRowCount = 20
        controller._session.findById.return_value = mock_table

        result = controller.select_multiple_rows("wnd[0]/usr/tblTEST", [1, 4])
        assert result["status"] == "success"
        assert result["count"] == 2


# ===========================================================================
# Shell Content Reading Tests
# ===========================================================================

class TestReadShellContent:
    """Tests for read_shell_content()."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def test_html_viewer(self):
        """Reads InnerHTML and URL from HTMLViewer shell."""
        controller = self._make_controller_with_session()
        mock_shell = MagicMock()
        mock_shell.Type = "GuiShell"
        mock_shell.SubType = "HTMLViewer"
        mock_shell.InnerHTML = "<h1>Report</h1>"
        mock_shell.CurrentUrl = "about:blank"
        mock_shell.Text = "Report"
        controller._session.findById.return_value = mock_shell

        result = controller.read_shell_content("wnd[0]/usr/shell")

        assert result["sub_type"] == "HTMLViewer"
        assert result["inner_html"] == "<h1>Report</h1>"
        assert result["url"] == "about:blank"

    def test_generic_shell(self):
        """Falls back to Text property for unknown shell types."""
        controller = self._make_controller_with_session()
        mock_shell = MagicMock()
        mock_shell.Type = "GuiShell"
        mock_shell.SubType = "Unknown"
        mock_shell.Text = "Some content"
        # No InnerHTML / CurrentUrl
        del mock_shell.InnerHTML
        del mock_shell.CurrentUrl
        controller._session.findById.return_value = mock_shell

        result = controller.read_shell_content("wnd[0]/usr/shell")
        assert result["text"] == "Some content"

    def test_error_handling(self):
        """Returns error dict when element not found."""
        controller = self._make_controller_with_session()
        controller._session.findById.side_effect = Exception("not found")

        result = controller.read_shell_content("wnd[0]/usr/shell")
        assert "error" in result


# ===========================================================================
# Screen Info Popup Detection Tests
# ===========================================================================

class TestScreenInfoActiveWindow:
    """Tests for active window detection in get_screen_info()."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def _setup_screen(self, controller, popup_wnd_id=None, popup_title=None):
        """Set up mocks. If popup_wnd_id is given, simulates a popup."""
        mock_main = MagicMock()
        mock_main.Text = "SAP Easy Access"

        mock_popup = MagicMock()
        mock_popup.Text = popup_title or ""

        mock_info = MagicMock()
        mock_info.Transaction = "SESSION_MANAGER"
        mock_info.Program = "SAPMSYST"
        mock_info.ScreenNumber = 100

        mock_sbar = MagicMock()
        mock_sbar.Text = ""
        mock_sbar.MessageType = ""
        mock_sbar.MessageId = ""
        mock_sbar.MessageNumber = ""

        controller._session.Info = mock_info

        # ActiveWindow
        if popup_wnd_id:
            mock_active = MagicMock()
            mock_active.Id = f"/app/con[0]/ses[0]/{popup_wnd_id}"
            mock_active.Text = popup_title or ""
            controller._session.ActiveWindow = mock_active
        else:
            mock_active = MagicMock()
            mock_active.Id = "/app/con[0]/ses[0]/wnd[0]"
            controller._session.ActiveWindow = mock_active

        def find_by_id(element_id):
            if element_id == "wnd[0]":
                return mock_main
            if element_id == "wnd[0]/sbar":
                return mock_sbar
            if popup_wnd_id and element_id == popup_wnd_id:
                return mock_popup
            if popup_wnd_id and element_id == f"{popup_wnd_id}/sbar":
                raise Exception("no sbar on popup")
            raise Exception("not found")

        controller._session.findById.side_effect = find_by_id

    def test_no_popup_active_window_is_wnd0(self):
        """When no popup, active_window is 'wnd[0]'."""
        controller = self._make_controller_with_session()
        self._setup_screen(controller)

        result = controller.get_screen_info()
        assert result["active_window"] == "wnd[0]"
        assert result["title"] == "SAP Easy Access"
        assert result["transaction"] == "SESSION_MANAGER"

    def test_popup_active_window_is_wnd1(self):
        """When popup opens, active_window is 'wnd[1]' and title is popup's."""
        controller = self._make_controller_with_session()
        self._setup_screen(controller, popup_wnd_id="wnd[1]",
                           popup_title="Enter Values")

        result = controller.get_screen_info()
        assert result["active_window"] == "wnd[1]"
        assert result["title"] == "Enter Values"
        assert result["transaction"] == "SESSION_MANAGER"

    def test_nested_popup_wnd2(self):
        """Detects wnd[2] (popup on top of popup)."""
        controller = self._make_controller_with_session()
        self._setup_screen(controller, popup_wnd_id="wnd[2]",
                           popup_title="Confirmation")

        result = controller.get_screen_info()
        assert result["active_window"] == "wnd[2]"
        assert result["title"] == "Confirmation"

    def test_active_window_exception_falls_back_to_wnd0(self):
        """When ActiveWindow throws, falls back to wnd[0]."""
        controller = self._make_controller_with_session()
        self._setup_screen(controller)

        type(controller._session).ActiveWindow = property(
            lambda self: (_ for _ in ()).throw(Exception("not supported"))
        )

        result = controller.get_screen_info()
        assert result["active_window"] == "wnd[0]"
        assert result["title"] == "SAP Easy Access"

    def test_active_window_none_falls_back_to_wnd0(self):
        """When ActiveWindow returns None, falls back to wnd[0]."""
        controller = self._make_controller_with_session()
        self._setup_screen(controller)
        controller._session.ActiveWindow = None

        result = controller.get_screen_info()
        assert result["active_window"] == "wnd[0]"

    def test_action_tool_reports_popup(self):
        """press_button result includes active_window from screen info."""
        controller = self._make_controller_with_session()
        self._setup_screen(controller, popup_wnd_id="wnd[1]",
                           popup_title="Selection Screen")

        mock_button = MagicMock()
        original_find = controller._session.findById.side_effect

        def find_by_id(element_id):
            if element_id == "wnd[0]/tbar[1]/btn[8]":
                return mock_button
            return original_find(element_id)

        controller._session.findById.side_effect = find_by_id

        result = controller.press_button("wnd[0]/tbar[1]/btn[8]")
        assert result["status"] == "pressed"
        assert result["screen"]["active_window"] == "wnd[1]"
        assert result["screen"]["title"] == "Selection Screen"

    def test_status_bar_falls_back_to_wnd0(self):
        """When popup has no sbar, status bar is read from wnd[0]."""
        controller = self._make_controller_with_session()
        self._setup_screen(controller, popup_wnd_id="wnd[1]",
                           popup_title="Popup Without Sbar")

        # The _setup_screen mock already raises for wnd[1]/sbar
        # and returns a valid sbar for wnd[0]/sbar
        result = controller.get_screen_info()
        assert result["active_window"] == "wnd[1]"
        # Status bar should still be populated from wnd[0]
        assert result["message"] is not None  # got "" from mock, not None


# ===========================================================================
# Screen Element Filtering Tests (P1)
# ===========================================================================

class TestGetScreenElementsFiltering:
    """Tests for type_filter and changeable_only in get_screen_elements."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def _make_container(self, elements):
        """Build a mock container with child elements.

        elements: list of dicts with type, name, text, changeable, children (optional).
        """
        container = MagicMock()
        children = []
        for elem in elements:
            child = MagicMock()
            child.Id = f"wnd[0]/usr/{elem.get('name', 'x')}"
            child.Type = elem["type"]
            child.Name = elem.get("name", "")
            child.Text = elem.get("text", "")
            child.Changeable = elem.get("changeable", False)
            child.Visible = True

            sub_children = elem.get("children", [])
            if sub_children:
                sub_mocks = []
                for sc in sub_children:
                    sc_mock = MagicMock()
                    sc_mock.Id = f"wnd[0]/usr/{elem.get('name', 'x')}/{sc.get('name', 'y')}"
                    sc_mock.Type = sc["type"]
                    sc_mock.Name = sc.get("name", "")
                    sc_mock.Text = sc.get("text", "")
                    sc_mock.Changeable = sc.get("changeable", False)
                    sc_mock.Visible = True
                    sc_mock.Children.Count = 0
                    sub_mocks.append(sc_mock)
                child.Children.Count = len(sub_mocks)
                child.Children.side_effect = lambda i, _mocks=sub_mocks: _mocks[i]
            else:
                child.Children.Count = 0

            children.append(child)

        container.Children.Count = len(children)
        container.Children.side_effect = lambda i: children[i]
        return container

    def test_no_filters_returns_all(self):
        """Without filters, all elements are returned (backward compat)."""
        controller = self._make_controller_with_session()
        container = self._make_container([
            {"type": "GuiTextField", "name": "txtA", "changeable": True},
            {"type": "GuiLabel", "name": "lblB", "changeable": False},
            {"type": "GuiButton", "name": "btnC", "changeable": True},
        ])
        controller._session.findById.return_value = container

        result = controller.get_screen_elements("wnd[0]/usr")
        assert len(result) == 3

    def test_type_filter_returns_only_matching_types(self):
        """type_filter CSV filters to specified types only."""
        controller = self._make_controller_with_session()
        container = self._make_container([
            {"type": "GuiTextField", "name": "txtA", "changeable": True},
            {"type": "GuiLabel", "name": "lblB", "changeable": False},
            {"type": "GuiButton", "name": "btnC", "changeable": True},
        ])
        controller._session.findById.return_value = container

        result = controller.get_screen_elements(
            "wnd[0]/usr", type_filter="GuiTextField,GuiButton",
        )
        types = {e.type for e in result}
        assert types == {"GuiTextField", "GuiButton"}
        assert len(result) == 2

    def test_changeable_only_filters_non_changeable(self):
        """changeable_only=True excludes non-changeable elements."""
        controller = self._make_controller_with_session()
        container = self._make_container([
            {"type": "GuiTextField", "name": "txtA", "changeable": True},
            {"type": "GuiLabel", "name": "lblB", "changeable": False},
            {"type": "GuiCTextField", "name": "ctxtC", "changeable": True},
        ])
        controller._session.findById.return_value = container

        result = controller.get_screen_elements(
            "wnd[0]/usr", changeable_only=True,
        )
        assert len(result) == 2
        assert all(e.changeable for e in result)

    def test_combined_filters(self):
        """type_filter + changeable_only work together."""
        controller = self._make_controller_with_session()
        container = self._make_container([
            {"type": "GuiTextField", "name": "txtA", "changeable": True},
            {"type": "GuiTextField", "name": "txtB", "changeable": False},
            {"type": "GuiButton", "name": "btnC", "changeable": True},
        ])
        controller._session.findById.return_value = container

        result = controller.get_screen_elements(
            "wnd[0]/usr", type_filter="GuiTextField", changeable_only=True,
        )
        assert len(result) == 1
        assert result[0].name == "txtA"

    def test_filter_recurses_into_non_matching_containers(self):
        """Filtering doesn't prevent recursion into container elements."""
        controller = self._make_controller_with_session()
        # A container (GuiSimpleContainer) holds a GuiTextField child.
        # Even though the container doesn't match type_filter, its child should.
        container = self._make_container([
            {"type": "GuiSimpleContainer", "name": "cont", "changeable": False,
             "children": [
                 {"type": "GuiTextField", "name": "txtInner", "changeable": True},
             ]},
        ])
        controller._session.findById.return_value = container

        result = controller.get_screen_elements(
            "wnd[0]/usr", type_filter="GuiTextField",
        )
        assert len(result) == 1
        assert result[0].type == "GuiTextField"

    def test_max_depth_limits_recursion(self):
        """max_depth=1 returns only direct children, not grandchildren."""
        controller = self._make_controller_with_session()
        # Container with a nested child (depth 2)
        container = self._make_container([
            {"type": "GuiSimpleContainer", "name": "cont", "changeable": False,
             "children": [
                 {"type": "GuiTextField", "name": "txtDeep", "changeable": True},
             ]},
            {"type": "GuiButton", "name": "btnTop", "changeable": True},
        ])
        controller._session.findById.return_value = container

        # depth=1: only direct children (container + button), no grandchildren
        result_shallow = controller.get_screen_elements(
            "wnd[0]/usr", max_depth=1,
        )
        types_shallow = [e.type for e in result_shallow]
        assert "GuiSimpleContainer" in types_shallow
        assert "GuiButton" in types_shallow
        assert "GuiTextField" not in types_shallow

        # depth=2: recurse into container, so grandchild is included
        result_deep = controller.get_screen_elements(
            "wnd[0]/usr", max_depth=2,
        )
        types_deep = [e.type for e in result_deep]
        assert "GuiTextField" in types_deep
        assert len(result_deep) == 3  # container + button + textfield

    def test_invalid_container_raises_error(self):
        """Invalid container IDs should raise instead of looking like empty screens."""
        from mcp_sap_gui.sap_controller import SAPGUIError

        controller = self._make_controller_with_session()
        controller._session.findById.side_effect = Exception("not found")

        with pytest.raises(SAPGUIError, match="wnd\\[0\\]/bad"):
            controller.get_screen_elements("wnd[0]/bad")


# ===========================================================================
# Table Read Filtering Tests (P2)
# ===========================================================================

class TestReadTableFiltering:
    """Tests for columns, columns_only, start_row in read_table."""

    def _make_controller_with_session(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        return controller

    def _make_alv_grid(self, columns, rows_data):
        """Create a mock ALV grid (GuiGridView)."""
        grid = MagicMock()
        grid.ColumnCount = len(columns)
        grid.ColumnOrder.side_effect = lambda i: columns[i]
        grid.GetColumnTooltip.side_effect = lambda c: f"Tip_{c}"
        grid.GetDisplayedColumnTitle.side_effect = lambda c: f"Title_{c}"
        grid.RowCount = len(rows_data)

        def get_cell_value(row, col):
            col_idx = columns.index(col)
            return rows_data[row][col_idx]
        grid.GetCellValue.side_effect = get_cell_value
        return grid

    # ---- ALV tests ----

    def test_read_table_columns_only_alv(self):
        """columns_only returns metadata with empty data for ALV."""
        controller = self._make_controller_with_session()
        grid = self._make_alv_grid(["COL_A", "COL_B"], [["a1", "b1"], ["a2", "b2"]])
        controller._session.findById.return_value = grid

        result = controller.read_table("grid1", columns_only=True)

        assert result["columns_only"] is True
        assert result["data"] == []
        assert result["rows_returned"] == 0
        assert result["total_rows"] == 2
        assert result["columns"] == ["COL_A", "COL_B"]
        assert len(result["column_info"]) == 2

    def test_read_table_column_filter_alv(self):
        """columns CSV filters to selected columns in ALV data."""
        controller = self._make_controller_with_session()
        grid = self._make_alv_grid(
            ["COL_A", "COL_B", "COL_C"],
            [["a1", "b1", "c1"], ["a2", "b2", "c2"]],
        )
        controller._session.findById.return_value = grid

        result = controller.read_table("grid1", columns="COL_A,COL_C")

        assert result["columns"] == ["COL_A", "COL_C"]
        assert len(result["column_info"]) == 2
        assert set(result["data"][0].keys()) == {"COL_A", "COL_C", "_absolute_row_index"}
        assert result["data"][0]["COL_A"] == "a1"
        assert result["data"][0]["COL_C"] == "c1"

    def test_read_table_start_row_alv(self):
        """start_row skips first N rows for ALV."""
        controller = self._make_controller_with_session()
        grid = self._make_alv_grid(
            ["COL"],
            [["v0"], ["v1"], ["v2"], ["v3"], ["v4"]],
        )
        controller._session.findById.return_value = grid

        result = controller.read_table("grid1", max_rows=2, start_row=2)

        assert result["start_row"] == 2
        assert result["rows_returned"] == 2
        assert result["data"][0]["COL"] == "v2"
        assert result["data"][0]["_absolute_row_index"] == 2
        assert result["data"][1]["COL"] == "v3"

    def test_read_table_start_row_past_end_alv(self):
        """start_row beyond row count returns empty data for ALV."""
        controller = self._make_controller_with_session()
        grid = self._make_alv_grid(["COL"], [["v0"], ["v1"]])
        controller._session.findById.return_value = grid

        result = controller.read_table("grid1", start_row=10)

        assert result["rows_returned"] == 0
        assert result["data"] == []

    # ---- TableControl tests ----

    def _make_table_control(self, columns, rows_data, total_rows=None,
                            visible_rows=None, initial_scroll_position=0):
        """Create a mock GuiTableControl (reusable helper)."""
        if total_rows is None:
            total_rows = len(rows_data)
        if visible_rows is None:
            visible_rows = len(rows_data)

        mock_table = MagicMock()
        mock_table.Type = "GuiTableControl"
        mock_table.TableFieldName = "TEST"
        mock_table.RowCount = total_rows
        mock_table.VisibleRowCount = visible_rows

        mock_cols = MagicMock()
        mock_cols.Count = len(columns)
        col_mocks = []
        for col_def in columns:
            col_mock = MagicMock()
            col_mock.Title = col_def.get("title", "")
            col_mock.Tooltip = col_def.get("tooltip", "")
            col_mocks.append(col_mock)
        mock_cols.side_effect = lambda i: col_mocks[i]
        mock_table.Columns = mock_cols

        _col_names = [c.get("name", f"col_{i}") for i, c in enumerate(columns)]

        scrollbar = MagicMock()
        scrollbar.Minimum = 0
        scrollbar.Maximum = max(0, total_rows - visible_rows)
        scrollbar.Position = initial_scroll_position
        mock_table.VerticalScrollbar = scrollbar

        def get_cell(row_idx, col_idx):
            abs_row = scrollbar.Position + row_idx
            cell = MagicMock()
            cell.Name = _col_names[col_idx] if col_idx < len(_col_names) else f"col_{col_idx}"
            if abs_row < len(rows_data) and col_idx < len(rows_data[abs_row]):
                cell.Type = "GuiTextField"
                cell.Text = str(rows_data[abs_row][col_idx])
            else:
                cell.Type = "GuiTextField"
                cell.Text = ""
            return cell
        mock_table.GetCell = MagicMock(side_effect=get_cell)
        return mock_table

    def test_read_table_columns_only_table_control(self):
        """columns_only returns metadata with empty data for TableControl."""
        controller = self._make_controller_with_session()
        cols = [{"name": "A", "title": "ColA"}, {"name": "B", "title": "ColB"}]
        mock_table = self._make_table_control(cols, [["a1", "b1"]])
        controller._session.findById.return_value = mock_table

        result = controller.read_table("tbl1", columns_only=True)

        assert result["columns_only"] is True
        assert result["data"] == []
        assert result["rows_returned"] == 0
        assert result["columns"] == ["A", "B"]
        assert len(result["column_info"]) == 2

    def test_read_table_column_filter_table_control(self):
        """columns CSV filters to selected columns in TableControl data."""
        controller = self._make_controller_with_session()
        cols = [
            {"name": "COL_A", "title": "A"},
            {"name": "COL_B", "title": "B"},
            {"name": "COL_C", "title": "C"},
        ]
        mock_table = self._make_table_control(
            cols, [["a1", "b1", "c1"], ["a2", "b2", "c2"]],
        )
        controller._session.findById.return_value = mock_table

        result = controller.read_table("tbl1", columns="COL_A,COL_C")

        assert result["columns"] == ["COL_A", "COL_C"]
        assert len(result["column_info"]) == 2
        assert set(result["data"][0].keys()) == {"COL_A", "COL_C", "_absolute_row_index"}
        assert result["data"][0]["COL_A"] == "a1"
        assert result["data"][0]["COL_C"] == "c1"

    def test_read_table_start_row_table_control(self):
        """start_row scrolls TableControl to the requested position."""
        controller = self._make_controller_with_session()
        cols = [{"name": "COL", "title": "Col"}]
        rows = [[f"v{i}"] for i in range(50)]
        mock_table = self._make_table_control(
            cols, rows, total_rows=50, visible_rows=5,
        )
        controller._session.findById.return_value = mock_table

        result = controller.read_table("tbl1", max_rows=3, start_row=10)

        # Verify scroll was requested
        assert mock_table.VerticalScrollbar.Position == 10
        assert result["first_visible_row"] == 10
        assert result["rows_returned"] == 3
        assert result["data"][0]["COL"] == "v10"

    def test_read_table_negative_start_row_clamped_alv(self):
        """Negative start_row is clamped to 0 for ALV."""
        controller = self._make_controller_with_session()
        grid = self._make_alv_grid(["COL"], [["v0"], ["v1"]])
        controller._session.findById.return_value = grid

        result = controller.read_table("grid1", start_row=-5)

        assert result["start_row"] == 0
        assert result["rows_returned"] == 2
        assert result["data"][0]["_absolute_row_index"] == 0

    def test_read_table_column_filter_empty_col_with_data_in_other_cols(self):
        """TableControl: rows with data in non-filtered columns are not treated as padding."""
        controller = self._make_controller_with_session()
        cols = [
            {"name": "KEY", "title": "Key"},
            {"name": "OPT", "title": "Optional"},
        ]
        # OPT is empty for all rows, but KEY has data
        rows = [["k1", ""], ["k2", ""], ["k3", ""]]
        mock_table = self._make_table_control(cols, rows)
        controller._session.findById.return_value = mock_table

        # Filter to only OPT — should still return 3 rows because
        # padding detection checks all columns
        result = controller.read_table("tbl1", columns="OPT")

        assert result["rows_returned"] == 3

