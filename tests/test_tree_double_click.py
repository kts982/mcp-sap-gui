"""Tree double-clicks reach the application the way a mouse double-click does.

Modelled on trees seen live on SAP GUI 8.10 (issue #20):
- SE19's BAdI implementation tree registers node and item double-click but
  opens the SELECTED node, and numbers its nodes anew after every round trip;
- an SM34 view-cluster tree registers only the item double-click;
- the events a tree reacts to are GuiShell.OcxEvents.
"""

from typing import Dict, List, Optional

import pytest

from tests.fakes import controller_for

TREE_ID = "wnd[0]/shellcont/shell"

SE19_EVENTS = [18, 41, 28, 39, 22, 25]
SM34_EVENTS = [22, 26, 39]


class Collection:
    """GuiCollection-shaped: Count plus call-by-index."""

    def __init__(self, items):
        self._items = list(items)
        self.Count = len(self._items)

    def __call__(self, i):
        return self._items[i]


class FakeTree:
    """A column tree whose behaviour follows its registered events."""

    def __init__(self, texts: List[str], events: Optional[List[int]],
                 mode: int = 2, acts_on: str = "event", renumber: bool = False,
                 columns: List[str] = ("&Hierarchy",)) -> None:
        self.nodes: Dict[str, str] = {f"{i + 1:>11}": t for i, t in enumerate(texts)}
        self.events = events
        self.mode = mode
        self.acts_on = acts_on
        self.renumber = renumber
        self.columns = list(columns)
        self.selected: List[str] = []
        self.opened: Optional[str] = None
        self.sent: List[tuple] = []

    @property
    def OcxEvents(self):  # noqa: N802 - COM name
        if self.events is None:
            raise AttributeError("OcxEvents")  # an SAP GUI that does not say
        return Collection(self.events)

    def _check(self, key):
        if key not in self.nodes:
            raise Exception("Exception occurred.")

    def key_of(self, text: str) -> str:
        return next(k for k, t in self.nodes.items() if t == text)

    # --- reads -------------------------------------------------------------
    def GetTreeType(self):  # noqa: N802
        return 2

    def GetColumnNames(self):  # noqa: N802
        return Collection(self.columns)

    def GetColumnTitleFromName(self, name):  # noqa: N802
        return name

    def GetSelectionMode(self):  # noqa: N802
        return self.mode

    def GetAllNodeKeys(self):  # noqa: N802
        return Collection(self.nodes)

    def GetNodeTextByKey(self, key):  # noqa: N802
        return self.nodes.get(key, "")

    def GetParent(self, key):  # noqa: N802
        return ""

    def GetHierarchyLevel(self, key):  # noqa: N802
        self._check(key)
        return 0

    def GetItemText(self, key, item):  # noqa: N802
        self._check(key)
        if item != self.columns[-1] and len(self.columns) > 1:
            return ""  # icon columns before the text, as in SPRO
        return self.nodes[key]

    # --- actions -----------------------------------------------------------
    def SelectNode(self, key):  # noqa: N802
        self._check(key)
        if self.mode in (1, 3):
            self.selected.append(key)
        else:
            self.selected = [key]

    def SelectItem(self, key, item):  # noqa: N802
        self._check(key)
        self.selected = [key]

    def UnselectAll(self):  # noqa: N802
        self.selected = []

    def DoubleClickNode(self, key):  # noqa: N802
        self._check(key)
        self.sent.append(("node", key))
        if 25 in (self.events or [25]):
            self._round_trip(key)

    def DoubleClickItem(self, key, item):  # noqa: N802
        self._check(key)
        self.sent.append(("item", key, item))
        if 22 in (self.events or [22]):
            self._round_trip(key)

    def _round_trip(self, key):
        target = self.selected[-1] if self.acts_on == "selection" else key
        self.opened = self.nodes[target]
        if self.renumber:
            shift = len(self.nodes)
            self.nodes = {f"{int(k) + shift:>11}": t for k, t in self.nodes.items()}
            self.selected = []


def control(tree: FakeTree):
    controller = controller_for(tree)
    controller.get_screen_info = lambda include_popup=True: {"active_window": "wnd[0]"}
    return controller


def se19():
    # The implementation itself is selected, as SE19 opens it.
    tree = FakeTree(["ATTRIBUTE_PROPERTIES", "Implementing Class", "Filter Values"],
                    SE19_EVENTS, acts_on="selection", renumber=True)
    tree.selected = [tree.key_of("ATTRIBUTE_PROPERTIES")]
    return tree


@pytest.mark.parametrize("method", ["node", "item"])
def test_se19_opens_the_clicked_node_not_the_selected_one(method):
    tree = se19()
    key = tree.key_of("Filter Values")

    if method == "node":
        result = control(tree).double_click_tree_node(TREE_ID, key)
    else:
        result = control(tree).double_click_tree_item(TREE_ID, key, "&Hierarchy")

    assert result["status"] == "double_clicked"
    assert tree.opened == "Filter Values"  # without the select: the implementation
    assert "sent_as" not in result  # both double-clicks are registered


def test_renumbered_tree_is_reported_and_the_old_key_explained():
    tree = se19()
    old_key = tree.key_of("Implementing Class")
    controller = control(tree)

    first = controller.double_click_tree_node(TREE_ID, old_key)
    assert first["node_keys_changed"] is True

    second = controller.double_click_tree_node(TREE_ID, old_key)
    assert "not in the tree (any more)" in second["error"]
    assert "read the tree" in second["error"]

    stale_select = controller.select_tree_node(TREE_ID, old_key)
    assert "not in the tree (any more)" in stale_select["error"]


def test_stable_keys_are_not_reported_as_changed():
    tree = FakeTree(["Object Type Codes", "Node Type Codes"], SM34_EVENTS,
                    columns=["Column1"])

    result = control(tree).double_click_tree_item(TREE_ID, tree.key_of("Node Type Codes"),
                                                  "Column1")

    assert "node_keys_changed" not in result
    assert tree.opened == "Node Type Codes"


def test_node_double_click_goes_as_item_double_click_where_only_that_is_registered():
    tree = FakeTree(["Object Type Codes", "Node Type Codes"], SM34_EVENTS,
                    columns=["Column1"])
    key = tree.key_of("Node Type Codes")

    result = control(tree).double_click_tree_node(TREE_ID, key)

    assert result["sent_as"] == "item_double_click"
    assert result["item_name"] == "Column1"
    assert tree.sent == [("item", key, "Column1")]
    assert tree.opened == "Node Type Codes"
    assert "warning" not in result


def test_routed_double_click_lands_on_the_text_not_on_an_icon_column():
    # SPRO: '1' opens the documentation and '2' runs the activity (links);
    # the text is in 'TEXT', where a mouse double-click would land.
    tree = FakeTree(["Activate Business Functions"], [18, 22, 26, 33, 35, 39, 40, 45],
                    columns=["1", "2", "TEXT"])
    key = tree.key_of("Activate Business Functions")

    result = control(tree).double_click_tree_node(TREE_ID, key)

    assert result["item_name"] == "TEXT"
    assert tree.sent == [("item", key, "TEXT")]


def test_item_double_click_goes_as_node_double_click_where_only_that_is_registered():
    tree = FakeTree(["A", "B"], [25])
    key = tree.key_of("B")

    result = control(tree).double_click_tree_item(TREE_ID, key, "&Hierarchy")

    assert result["sent_as"] == "node_double_click"
    assert tree.sent == [("node", key)]
    assert tree.opened == "B"


def test_tree_without_a_double_click_event_says_so():
    tree = FakeTree(["Activity"], [18, 35, 39])

    result = control(tree).double_click_tree_node(TREE_ID, tree.key_of("Activity"))

    assert result["status"] == "double_clicked"  # still sent, as before
    assert "link_click" in result["warning"]
    assert "sap_click_tree_link" in result["warning"]
    assert tree.opened is None


def test_unknown_events_send_what_was_asked_after_selecting():
    tree = FakeTree(["A", "B"], None, acts_on="selection")
    key = tree.key_of("B")

    result = control(tree).double_click_tree_node(TREE_ID, key)

    assert tree.sent == [("node", key)]
    assert tree.opened == "B"
    assert "sent_as" not in result and "warning" not in result


def test_multiple_selection_is_replaced_not_extended():
    tree = FakeTree(["A", "B"], [25], mode=3, acts_on="selection")
    tree.selected = [tree.key_of("A")]

    control(tree).double_click_tree_node(TREE_ID, tree.key_of("B"))

    assert tree.opened == "B"


def test_failed_selection_does_not_stop_the_double_click():
    tree = FakeTree(["A", "B"], [25])

    def refuse(*args):
        raise Exception("Exception occurred.")

    tree.SelectNode = refuse
    tree.SelectItem = refuse
    key = tree.key_of("B")

    result = control(tree).double_click_tree_node(TREE_ID, key)

    assert result["status"] == "double_clicked"
    assert tree.sent == [("node", key)]


def test_read_tree_and_root_children_name_the_registered_events():
    tree = FakeTree(["A", "B"], SE19_EVENTS)
    controller = control(tree)

    read = controller.read_tree(TREE_ID)
    root = controller.get_tree_node_children(TREE_ID)
    child = controller.get_tree_node_children(TREE_ID, tree.key_of("A"))

    expected = ["expand_no_children", "item_double_click", "node_double_click",
                "header_click", "item_keypress", "header_context_menu"]
    assert read["events"] == expected
    assert root["events"] == expected
    assert "events" not in child


def test_read_tree_leaves_events_out_when_sap_gui_does_not_say():
    tree = FakeTree(["A"], None)

    assert "events" not in control(tree).read_tree(TREE_ID)
