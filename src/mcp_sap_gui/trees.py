"""
Trees Mixin - tree control reading and interaction.

Provides tree reading, expand/collapse, and node interaction for
SAP GUI tree controls (SimpleTree, ListTree, ColumnTree).
"""

import logging
from typing import Any, Dict

logger = logging.getLogger(__name__)


class TreesMixin:
    """Mixin for tree control operations on SAP GUI screens."""

    # =========================================================================
    # Tree Control Operations
    # =========================================================================

    def _get_tree_column_info(self, tree) -> tuple:
        """
        Get column names and titles from a tree control.

        Uses the official SAP GUI Scripting API methods:
        - GetColumnNames() -> GuiCollection of internal column names
        - GetColumnHeaders() -> GuiCollection of display titles
        - GetColumnTitleFromName(name) -> title for a specific column
        - ColumnOrder property -> column sequence (Column Tree only)

        Tree types (GetTreeType): 0=Simple, 1=List, 2=Column

        Returns:
            Tuple of (column_names: list, column_titles: list)
        """
        column_names = []
        column_titles = []

        # Detect tree type
        tree_type_num = -1
        try:
            tree_type_num = tree.GetTreeType()
            logger.debug("Tree type number: %s", tree_type_num)
        except Exception as e:
            logger.debug("GetTreeType failed: %s", e)

        # Simple trees (type 0) have no columns
        if tree_type_num == 0:
            return column_names, column_titles

        # Strategy 1: GetColumnNames() -- returns a GuiCollection (works for List & Column trees)
        try:
            names_col = tree.GetColumnNames()
            if hasattr(names_col, 'Count'):
                for i in range(names_col.Count):
                    column_names.append(str(names_col(i)))
            elif hasattr(names_col, 'Length'):
                for i in range(names_col.Length):
                    column_names.append(str(names_col(i)))
            elif hasattr(names_col, '__iter__'):
                column_names = [str(n) for n in names_col]
            logger.debug("Got column names via GetColumnNames: %s", column_names)
        except Exception as e:
            logger.debug("GetColumnNames failed: %s", e)

        # Strategy 2: ColumnOrder property (Column Tree type 2 only)
        if not column_names and tree_type_num == 2:
            try:
                col_order = tree.ColumnOrder
                if hasattr(col_order, 'Count'):
                    for i in range(col_order.Count):
                        column_names.append(str(col_order(i)))
                elif hasattr(col_order, '__iter__'):
                    column_names = [str(n) for n in col_order]
                logger.debug("Got column names via ColumnOrder: %s", column_names)
            except Exception as e:
                logger.debug("ColumnOrder failed: %s", e)

        # Get column titles
        if column_names:
            # Try GetColumnTitleFromName for each column
            for name in column_names:
                try:
                    column_titles.append(tree.GetColumnTitleFromName(name))
                except Exception:
                    column_titles.append(name)

            # If all titles are empty, try GetColumnHeaders as fallback
            if all(not t for t in column_titles):
                try:
                    headers_col = tree.GetColumnHeaders()
                    fallback_titles = []
                    if hasattr(headers_col, 'Count'):
                        for i in range(headers_col.Count):
                            fallback_titles.append(str(headers_col(i)))
                    if fallback_titles:
                        column_titles = fallback_titles
                        logger.debug("Got titles via GetColumnHeaders: %s", column_titles)
                except Exception as e:
                    logger.debug("GetColumnHeaders fallback failed: %s", e)

        return column_names, column_titles

    def read_tree(self, tree_id: str, max_nodes: int = 200) -> Dict[str, Any]:
        """
        Read data from a tree control (SAP.TableTreeControl, SAP.ColumnTreeControl, etc.).

        Args:
            tree_id: SAP GUI tree ID (e.g., "wnd[0]/usr/shell/shellcont[0]/shell")
            max_nodes: Maximum number of nodes to read (default 200)

        Returns:
            Dict with tree structure, columns, and node data
        """
        self._require_session()

        try:
            tree = self._session.findById(tree_id)

            # Detect tree type (0=Simple, 1=List, 2=Column)
            tree_type = ""
            tree_type_num = -1
            try:
                tree_type_num = tree.GetTreeType()
                tree_type_names = {0: "Simple", 1: "List", 2: "Column"}
                tree_type = tree_type_names.get(tree_type_num, f"Unknown({tree_type_num})")
            except Exception:
                # Fallback to SubType / Text for display
                try:
                    tree_type = tree.SubType if hasattr(tree, 'SubType') else ""
                except Exception:
                    pass
                if not tree_type:
                    try:
                        tree_type = tree.Text
                    except Exception:
                        pass
            logger.debug("Tree type: %s (num=%s)", tree_type, tree_type_num)

            # Get hierarchy title if available (List/Column trees)
            hierarchy_title = ""
            try:
                hierarchy_title = tree.GetHierarchyTitle()
            except Exception:
                pass

            # Get column info using fallback chain
            column_names, column_titles = self._get_tree_column_info(tree)

            # Get all visible node keys
            node_keys = []
            try:
                keys = tree.GetAllNodeKeys()
                if hasattr(keys, 'Count'):
                    for i in range(min(keys.Count, max_nodes)):
                        node_keys.append(keys(i))
                elif hasattr(keys, '__iter__'):
                    for i, key in enumerate(keys):
                        if i >= max_nodes:
                            break
                        node_keys.append(key)
            except Exception as e:
                return {"tree_id": tree_id, "error": f"Cannot read node keys: {e}"}

            # Read each node
            nodes = []
            for key in node_keys:
                node = {"key": key}

                # Node text (works for SimpleTree; empty for TableTree/ColumnTree)
                try:
                    node["text"] = tree.GetNodeTextByKey(key)
                except Exception:
                    node["text"] = ""

                # Parent key - API documents GetParent(), fall back to GetParentNodeKey()
                try:
                    node["parent_key"] = tree.GetParent(key)
                except Exception:
                    try:
                        node["parent_key"] = tree.GetParentNodeKey(key)
                    except Exception:
                        node["parent_key"] = None

                # Children count
                try:
                    node["children_count"] = tree.GetNodeChildrenCount(key)
                except Exception:
                    node["children_count"] = 0

                # Folder state
                try:
                    node["is_folder"] = tree.IsFolderExpandable(key)
                except Exception:
                    node["is_folder"] = False

                try:
                    node["is_expanded"] = tree.IsFolderExpanded(key)
                except Exception:
                    node["is_expanded"] = False

                # Hierarchy level
                try:
                    node["hierarchy_level"] = tree.GetHierarchyLevel(key)
                except Exception:
                    node["hierarchy_level"] = None

                # Column values (for TableTree / ColumnTree)
                if column_names:
                    col_values = {}
                    for col_name in column_names:
                        try:
                            col_values[col_name] = tree.GetItemText(key, col_name)
                        except Exception:
                            col_values[col_name] = None
                    node["columns"] = col_values

                    # If node text is empty, use first non-empty column value
                    if not node["text"]:
                        for col_name in column_names:
                            val = col_values.get(col_name)
                            if val:
                                node["text"] = val
                                break

                nodes.append(node)

            return {
                "tree_id": tree_id,
                "tree_type": tree_type,
                "hierarchy_title": hierarchy_title,
                "total_nodes": len(node_keys),
                "column_titles": column_titles,
                "column_names": column_names,
                "nodes": nodes,
            }

        except Exception as e:
            return {"tree_id": tree_id, "error": str(e)}

    def expand_tree_node(self, tree_id: str, node_key: str) -> Dict[str, Any]:
        """
        Expand a folder node in a tree control.

        Args:
            tree_id: SAP GUI tree ID
            node_key: The key of the node to expand

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            tree = self._session.findById(tree_id)
            tree.ExpandNode(node_key)

            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "status": "expanded",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {"tree_id": tree_id, "node_key": node_key, "error": str(e)}

    def collapse_tree_node(self, tree_id: str, node_key: str) -> Dict[str, Any]:
        """
        Collapse a folder node in a tree control.

        Args:
            tree_id: SAP GUI tree ID
            node_key: The key of the node to collapse

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            tree = self._session.findById(tree_id)
            tree.CollapseNode(node_key)

            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "status": "collapsed",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {"tree_id": tree_id, "node_key": node_key, "error": str(e)}

    def select_tree_node(self, tree_id: str, node_key: str) -> Dict[str, Any]:
        """
        Select a node in a tree control.

        Args:
            tree_id: SAP GUI tree ID
            node_key: The key of the node to select

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            tree = self._session.findById(tree_id)
            tree.SelectNode(node_key)

            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "status": "selected",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {"tree_id": tree_id, "node_key": node_key, "error": str(e)}

    def double_click_tree_node(self, tree_id: str, node_key: str) -> Dict[str, Any]:
        """
        Double-click a node in a tree control (often opens details).

        Args:
            tree_id: SAP GUI tree ID
            node_key: The key of the node to double-click

        Returns:
            Dict with result status and screen info
        """
        self._require_session()

        try:
            tree = self._session.findById(tree_id)
            tree.DoubleClickNode(node_key)

            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "status": "double_clicked",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {"tree_id": tree_id, "node_key": node_key, "error": str(e)}

    def double_click_tree_item(self, tree_id: str, node_key: str,
                               item_name: str) -> Dict[str, Any]:
        """
        Double-click a specific item (column cell) in a tree node.

        Unlike DoubleClickNode which clicks the node itself, this clicks
        on a specific column cell within the node row.

        Args:
            tree_id: SAP GUI tree ID
            node_key: The key of the node
            item_name: Column name / item name to double-click

        Returns:
            Dict with result status and screen info
        """
        self._require_session()

        try:
            tree = self._session.findById(tree_id)
            tree.DoubleClickItem(node_key, item_name)

            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "item_name": item_name,
                "status": "double_clicked",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {
                "tree_id": tree_id, "node_key": node_key,
                "item_name": item_name, "error": str(e),
            }

    def click_tree_link(self, tree_id: str, node_key: str,
                        item_name: str) -> Dict[str, Any]:
        """
        Click a hyperlink in a tree node item.

        Args:
            tree_id: SAP GUI tree ID
            node_key: The key of the node
            item_name: Column name / item name containing the link

        Returns:
            Dict with result status and screen info
        """
        self._require_session()

        try:
            tree = self._session.findById(tree_id)
            tree.ClickLink(node_key, item_name)

            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "item_name": item_name,
                "status": "clicked",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return {
                "tree_id": tree_id, "node_key": node_key,
                "item_name": item_name, "error": str(e),
            }

    def find_tree_node_by_path(self, tree_id: str, path: str) -> Dict[str, Any]:
        """
        Find a node key by its path in the tree hierarchy.

        The path is a backslash-separated string of child indices,
        e.g. "2\\1\\2" means: 2nd child of root, then 1st child, then 2nd child.

        Args:
            tree_id: SAP GUI tree ID
            path: Path string (e.g., "2\\1\\2")

        Returns:
            Dict with the found node key
        """
        self._require_session()

        try:
            tree = self._session.findById(tree_id)
            node_key = tree.FindNodeKeyByPath(path)

            return {
                "tree_id": tree_id,
                "path": path,
                "node_key": node_key,
                "status": "found",
            }
        except Exception as e:
            return {"tree_id": tree_id, "path": path, "error": str(e)}
