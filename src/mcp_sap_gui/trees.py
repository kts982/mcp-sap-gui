"""
Trees Mixin - tree control reading and interaction.

Provides tree reading, expand/collapse, and node interaction for
SAP GUI tree controls (SimpleTree, ListTree, ColumnTree).
"""

import logging
from typing import Any, Dict, List, Optional, Set

logger = logging.getLogger(__name__)

# GuiShell.OcxEvents lists the events the application registered on a tree:
# the only events that reach it. The ids are the tree control's own, the
# ABAP constants CL_TREE_CONTROL_BASE / CL_ITEM_TREE_CONTROL=>EVENTID_*.
_TREE_EVENT_NAMES = {
    18: "expand_no_children",
    21: "selection_changed",
    22: "item_double_click",
    25: "node_double_click",
    26: "item_context_menu",
    28: "header_click",
    29: "button_click",
    33: "checkbox_change",
    35: "link_click",
    36: "node_context_menu",
    39: "item_keypress",
    40: "node_keypress",
    41: "header_context_menu",
    42: "default_context_menu",
    45: "drop_external_files",
}
_SELECTION_CHANGED = 21
_ITEM_DOUBLE_CLICK = 22
_NODE_DOUBLE_CLICK = 25

# GetSelectionMode 1 (multiple nodes) and 3 (multiple items): SelectNode
# adds to the selection there instead of replacing it.
_MULTIPLE_SELECTION_MODES = (1, 3)

_STALE_NODE_KEY = (
    "Node key {key!r} is not in the tree (any more). Some trees number "
    "their nodes anew after every round trip (SE19 does): read the tree "
    "again (sap_read_tree or sap_get_tree_node_children) and use the "
    "current key."
)


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

    def _tree_events(self, tree) -> Optional[Set[int]]:
        """Event ids the application registered on *tree*; None if unknown."""
        try:
            events = tree.OcxEvents
            return {int(events(i)) for i in range(events.Count)}
        except Exception as e:
            logger.debug("OcxEvents not readable: %s", e)
            return None

    @staticmethod
    def _event_names(events: Set[int]) -> List[str]:
        return [_TREE_EVENT_NAMES.get(e, f"event_{e}") for e in sorted(events)]

    def _node_key_gone(self, tree_id: str, node_key: str) -> bool:
        """True when the tree is still there but *node_key* is not.

        GetHierarchyLevel raises for an unknown key; GetNodeTextByKey
        returns '' and cannot tell.
        """
        try:
            tree = self._find_element(tree_id)
        except Exception:
            return False
        try:
            tree.GetHierarchyLevel(node_key)
            return False
        except Exception:
            return True

    def _tree_error(self, context: Dict[str, Any], exc: Exception,
                    fallback: str) -> Dict[str, Any]:
        """Error result that says so when the node key went stale."""
        node_key = context.get("node_key")
        if node_key and self._node_key_gone(context["tree_id"], node_key):
            logger.warning("%s: %s (node key no longer in the tree)", fallback, exc)
            result = dict(context)
            result["error"] = _STALE_NODE_KEY.format(key=node_key)
            return result
        return self._error_result(context, exc, fallback)

    def _text_item(self, tree, node_key: str) -> str:
        """The item that shows the node's text: where a mouse double-click lands.

        Not simply the first column: SPRO's first columns are icon links
        ('1' opens the documentation, '2' runs the activity), its text is
        in 'TEXT', and GetColumnNames leaves out its '&Hierarchy' column.
        """
        column_names, _ = self._get_tree_column_info(tree)
        for name in [*column_names, "&Hierarchy"]:
            try:
                if tree.GetItemText(node_key, name):
                    return name
            except Exception:
                continue
        return ""

    def _select_before_double_click(self, tree, node_key: str,
                                    item_name: str = "") -> None:
        """Select the node first, as a mouse double-click does.

        Some applications act on the selected node, not on the one the
        event names (SE19's BAdI implementation tree). A failed selection
        is logged only: the double-click is still sent.
        """
        if item_name:
            try:
                tree.SelectItem(node_key, item_name)  # replaces the selection
                return
            except Exception as e:
                logger.debug("SelectItem(%r, %r) failed: %s", node_key, item_name, e)
        try:
            multiple = tree.GetSelectionMode() in _MULTIPLE_SELECTION_MODES
        except Exception:
            multiple = False
        try:
            if multiple:
                tree.UnselectAll()
            tree.SelectNode(node_key)
        except Exception as e:
            logger.debug("SelectNode(%r) failed: %s", node_key, e)

    def _double_click(self, tree_id: str, node_key: str,
                      item_name: str = "") -> Dict[str, Any]:
        """Double-click a node (no item_name) or an item, as a mouse does.

        Sends the double-click the application registered: a tree that
        handles only item double-clicks (SM34 view clusters, SPRO) ignores
        a node double-click, and the other way round.
        """
        tree = self._find_element(tree_id)
        events = self._tree_events(tree)
        send_item = bool(item_name)
        result: Dict[str, Any] = {}
        if events is not None:
            has_node = _NODE_DOUBLE_CLICK in events
            has_item = _ITEM_DOUBLE_CLICK in events
            if send_item and not has_item and has_node:
                send_item = False
                result["sent_as"] = "node_double_click"
            elif not send_item and not has_node and has_item:
                item_name = self._text_item(tree, node_key)
                if item_name:
                    send_item = True
                    result["sent_as"] = "item_double_click"
                    result["item_name"] = item_name
            if not (has_node or has_item):
                result["warning"] = (
                    "This tree reacts to no double-click, so it was probably "
                    "ignored. Registered events: "
                    + (", ".join(self._event_names(events)) or "none")
                    + ". For link_click use sap_click_tree_link."
                )
            elif not send_item and not has_node:
                result["warning"] = (
                    "This tree reacts only to item double-clicks: use "
                    "sap_double_click_tree_item with a column from sap_read_tree."
                )

        self._select_before_double_click(tree, node_key, item_name if send_item else "")
        if events and _SELECTION_CHANGED in events:
            # The selection was a round trip of its own: find the tree again.
            tree = self._find_element(tree_id)
        if send_item:
            tree.DoubleClickItem(node_key, item_name)
        else:
            tree.DoubleClickNode(node_key)

        if self._node_key_gone(tree_id, node_key):
            result["node_keys_changed"] = True
        return result

    def read_tree(self, tree_id: str, max_nodes: int = 200) -> Dict[str, Any]:
        """
        Read data from a tree control (SAP.TableTreeControl, SAP.ColumnTreeControl, etc.).

        Args:
            tree_id: SAP GUI tree ID (e.g., "wnd[0]/usr/shell/shellcont[0]/shell")
            max_nodes: Maximum number of nodes to read (default 200)

        Returns:
            Dict with tree structure, columns, and node data.
            ``truncated`` is True when the tree has more loaded nodes
            (``available_nodes``) than were read (``total_nodes``).
        """
        self._require_session()

        try:
            tree = self._find_element(tree_id)

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
            available_nodes = 0
            try:
                keys = tree.GetAllNodeKeys()
                if hasattr(keys, 'Count'):
                    available_nodes = keys.Count
                    for i in range(min(keys.Count, max_nodes)):
                        node_keys.append(keys(i))
                elif hasattr(keys, '__iter__'):
                    # Count every key (cheap: no node is read past the cap).
                    for i, key in enumerate(keys):
                        available_nodes = i + 1
                        if i < max_nodes:
                            node_keys.append(key)
            except Exception as e:
                return self._error_result(
                    {"tree_id": tree_id},
                    e,
                    "Cannot read tree node keys",
                )

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

            result = {
                "tree_id": tree_id,
                "tree_type": tree_type,
                "hierarchy_title": hierarchy_title,
                # total_nodes is the number READ, kept for compatibility;
                # available_nodes is how many keys the tree has loaded.
                "total_nodes": len(node_keys),
                "nodes_returned": len(nodes),
                "available_nodes": available_nodes,
                "truncated": available_nodes > len(node_keys),
                "column_titles": column_titles,
                "column_names": column_names,
                "nodes": nodes,
            }
            # What the application reacts to, e.g. link_click or only
            # item_double_click: left out when SAP GUI does not say.
            events = self._tree_events(tree)
            if events is not None:
                result["events"] = self._event_names(events)
            return result

        except Exception as e:
            return self._error_result(
                {"tree_id": tree_id},
                e,
                "Could not read tree",
            )

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
            tree = self._find_element(tree_id)
            tree.ExpandNode(node_key)

            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "status": "expanded",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return self._tree_error(
                {"tree_id": tree_id, "node_key": node_key},
                e,
                "Could not expand tree node",
            )

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
            tree = self._find_element(tree_id)
            tree.CollapseNode(node_key)

            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "status": "collapsed",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return self._tree_error(
                {"tree_id": tree_id, "node_key": node_key},
                e,
                "Could not collapse tree node",
            )

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
            tree = self._find_element(tree_id)
            tree.SelectNode(node_key)

            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "status": "selected",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return self._tree_error(
                {"tree_id": tree_id, "node_key": node_key},
                e,
                "Could not select tree node",
            )

    def double_click_tree_node(self, tree_id: str, node_key: str) -> Dict[str, Any]:
        """
        Double-click a node in a tree control (often opens details).

        Args:
            tree_id: SAP GUI tree ID
            node_key: The key of the node to double-click

        Returns:
            Dict with result status and screen info; ``sent_as`` when the
            tree registered only the item double-click, ``node_keys_changed``
            when the tree renumbered its nodes.
        """
        self._require_session()

        try:
            sent = self._double_click(tree_id, node_key)
            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "status": "double_clicked",
                **sent,
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return self._tree_error(
                {"tree_id": tree_id, "node_key": node_key},
                e,
                "Could not double-click tree node",
            )

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
            Dict with result status and screen info; ``sent_as`` when the
            tree registered only the node double-click, ``node_keys_changed``
            when the tree renumbered its nodes.
        """
        self._require_session()

        try:
            sent = self._double_click(tree_id, node_key, item_name)
            return {
                "tree_id": tree_id,
                "node_key": node_key,
                "item_name": item_name,
                "status": "double_clicked",
                **sent,
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return self._tree_error(
                {
                    "tree_id": tree_id,
                    "node_key": node_key,
                    "item_name": item_name,
                },
                e,
                "Could not double-click tree item",
            )

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
            tree = self._find_element(tree_id)
            tree.ClickLink(node_key, item_name)

            result = {
                "tree_id": tree_id,
                "node_key": node_key,
                "item_name": item_name,
                "status": "clicked",
            }
            if self._node_key_gone(tree_id, node_key):
                result["node_keys_changed"] = True
            result["screen"] = self.get_screen_info()
            return result
        except Exception as e:
            return self._tree_error(
                {
                    "tree_id": tree_id,
                    "node_key": node_key,
                    "item_name": item_name,
                },
                e,
                "Could not click tree link",
            )

    def _get_node_text(self, tree, key, column_names):
        """Get searchable text for a node, with column fallback."""
        text = ""
        try:
            text = tree.GetNodeTextByKey(key)
        except Exception:
            pass
        if not text and column_names:
            for col_name in column_names:
                try:
                    val = tree.GetItemText(key, col_name)
                    if val:
                        text = val
                        break
                except Exception:
                    pass
        return text

    def _build_ancestor_path(self, tree, node_key, column_names):
        """Walk up the parent chain and build a ' > ' separated path string."""
        parts = []
        current = node_key
        seen = set()
        while current and current not in seen:
            seen.add(current)
            text = self._get_node_text(tree, current, column_names)
            parts.append(text or current)
            # Walk up
            parent = None
            try:
                parent = tree.GetParent(current)
            except Exception:
                try:
                    parent = tree.GetParentNodeKey(current)
                except Exception:
                    pass
            current = parent
        parts.reverse()
        return " > ".join(parts)

    def search_tree_nodes(self, tree_id: str, search_text: str,
                          column: str = "",
                          max_results: int = 20) -> Dict[str, Any]:
        """
        Search tree nodes by text and return matches with full ancestor paths.

        Case-insensitive substring match. For each match, walks up the parent
        chain to build the full path, making it easy to disambiguate nodes
        with the same label in different branches.

        Args:
            tree_id: SAP GUI tree ID
            search_text: Text to search for (case-insensitive substring)
            column: Optional column name to restrict search to
            max_results: Maximum number of matches to return

        Returns:
            Dict with matching nodes and their full ancestor paths.
            ``stopped_early`` is True when the search reached max_results
            before the last loaded node, so more matches are possible.
        """
        self._require_session()

        try:
            tree = self._find_element(tree_id)

            # Get column info for text fallback
            column_names, _ = self._get_tree_column_info(tree)

            # Get all node keys
            node_keys = []
            try:
                keys = tree.GetAllNodeKeys()
                if hasattr(keys, 'Count'):
                    for i in range(keys.Count):
                        node_keys.append(keys(i))
                elif hasattr(keys, '__iter__'):
                    node_keys = list(keys)
            except Exception as e:
                return self._error_result(
                    {"tree_id": tree_id},
                    e,
                    "Cannot read tree node keys",
                )

            needle = search_text.lower()
            matches = []
            searched_nodes = 0

            for key in node_keys:
                if len(matches) >= max_results:
                    break
                searched_nodes += 1

                # Determine text to search against
                if column:
                    try:
                        haystack = tree.GetItemText(key, column) or ""
                    except Exception:
                        continue
                else:
                    haystack = self._get_node_text(tree, key, column_names)

                if needle not in haystack.lower():
                    continue

                # Build match entry
                match = {"key": key, "text": haystack}

                match["path"] = self._build_ancestor_path(
                    tree, key, column_names
                )

                try:
                    match["is_folder"] = tree.IsFolderExpandable(key)
                except Exception:
                    match["is_folder"] = False

                try:
                    match["hierarchy_level"] = tree.GetHierarchyLevel(key)
                except Exception:
                    match["hierarchy_level"] = None

                matches.append(match)

            return {
                "tree_id": tree_id,
                "search_text": search_text,
                # total_matches counts the matches RETURNED. stopped_early:
                # the search hit max_results before the last node, so more
                # matches are possible (not certain).
                "total_matches": len(matches),
                "total_nodes": len(node_keys),
                "searched_nodes": searched_nodes,
                "stopped_early": searched_nodes < len(node_keys),
                "matches": matches,
            }

        except Exception as e:
            return self._error_result(
                {"tree_id": tree_id, "search_text": search_text},
                e,
                "Could not search tree nodes",
            )

    def get_tree_node_children(self, tree_id: str, node_key: str = "",
                               expand: bool = False) -> Dict[str, Any]:
        """
        Get the direct children of a specific tree node.

        Much more efficient than read_tree for navigating deep trees step by
        step — returns only the immediate children of the given node instead
        of the entire tree.

        Args:
            tree_id: SAP GUI tree ID
            node_key: Key of the parent node. Empty string = root nodes.
            expand: If True, expand the node first to reveal children.

        Returns:
            Dict with the parent info and a list of child nodes.
        """
        self._require_session()

        try:
            tree = self._find_element(tree_id)
            column_names, _ = self._get_tree_column_info(tree)

            # Optionally expand first
            expand_error = None
            if expand and node_key:
                try:
                    tree.ExpandNode(node_key)
                except Exception as e:
                    expand_error = self._sanitize_error_message(
                        e,
                        "Could not expand tree node",
                    )
                    logger.debug("ExpandNode failed for %s: %s", node_key, e)

            # Get all node keys and filter to direct children
            all_keys = []
            try:
                keys = tree.GetAllNodeKeys()
                if hasattr(keys, 'Count'):
                    for i in range(keys.Count):
                        all_keys.append(keys(i))
                elif hasattr(keys, '__iter__'):
                    all_keys = list(keys)
            except Exception as e:
                return self._error_result(
                    {"tree_id": tree_id},
                    e,
                    "Cannot read tree node keys",
                )

            children = []
            for key in all_keys:
                # Determine parent of this node
                parent = None
                try:
                    parent = tree.GetParent(key)
                except Exception:
                    try:
                        parent = tree.GetParentNodeKey(key)
                    except Exception:
                        pass

                # Match: root nodes (parent is None/empty) when node_key is empty,
                # or nodes whose parent matches node_key
                is_child = False
                if not node_key:
                    is_child = not parent
                else:
                    is_child = (parent == node_key)

                if not is_child:
                    continue

                child = {"key": key}
                child["text"] = self._get_node_text(tree, key, column_names)

                try:
                    child["is_folder"] = tree.IsFolderExpandable(key)
                except Exception:
                    child["is_folder"] = False

                try:
                    child["is_expanded"] = tree.IsFolderExpanded(key)
                except Exception:
                    child["is_expanded"] = False

                try:
                    child["children_count"] = tree.GetNodeChildrenCount(key)
                except Exception:
                    child["children_count"] = 0

                children.append(child)

            result = {
                "tree_id": tree_id,
                "node_key": node_key or "(root)",
                "children_count": len(children),
                "children": children,
            }

            # Include parent text for context
            if node_key:
                result["node_text"] = self._get_node_text(tree, node_key, column_names)
                result["path"] = self._build_ancestor_path(tree, node_key, column_names)
            else:
                # First look at the tree: say what the application reacts to.
                events = self._tree_events(tree)
                if events is not None:
                    result["events"] = self._event_names(events)

            if expand and node_key:
                result["screen"] = self.get_screen_info()

            if expand_error:
                result["expand_error"] = expand_error

            return result

        except Exception as e:
            return self._error_result(
                {"tree_id": tree_id, "node_key": node_key},
                e,
                "Could not read tree node children",
            )

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
            tree = self._find_element(tree_id)
            node_key = tree.FindNodeKeyByPath(path)

            return {
                "tree_id": tree_id,
                "path": path,
                "node_key": node_key,
                "status": "found",
            }
        except Exception as e:
            return self._error_result(
                {"tree_id": tree_id, "path": path},
                e,
                "Could not find tree node by path",
            )
