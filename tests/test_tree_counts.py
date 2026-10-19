"""read_tree / search_tree_nodes say when the count was capped."""

import pytest

from tests.fakes import controller_for

TREE_ID = "wnd[0]/shellcont/shell"


class KeyCollection:
    """GuiCollection-shaped: Count plus call-by-index."""

    def __init__(self, keys):
        self._keys = keys
        self.Count = len(keys)

    def __call__(self, i):
        return self._keys[i]


class FakeTree:
    def __init__(self, texts, collection=True):
        self.texts = texts
        self.collection = collection

    def GetTreeType(self):  # noqa: N802 - COM name
        return 0  # simple tree, no columns

    def GetAllNodeKeys(self):  # noqa: N802
        keys = list(self.texts)
        return KeyCollection(keys) if self.collection else iter(keys)

    def GetNodeTextByKey(self, key):  # noqa: N802
        return self.texts[key]

    def GetParent(self, key):  # noqa: N802
        return ""


def tree(n, collection=True, text=lambda i: f"Node {i}"):
    return FakeTree({f"K{i:03}": text(i) for i in range(n)}, collection)


@pytest.mark.parametrize("collection", [True, False])
def test_read_tree_flags_truncation(collection):
    result = controller_for(tree(5, collection)).read_tree(TREE_ID, max_nodes=3)
    assert result["total_nodes"] == 3  # unchanged meaning: nodes read
    assert result["nodes_returned"] == 3
    assert result["available_nodes"] == 5
    assert result["truncated"] is True


@pytest.mark.parametrize("collection", [True, False])
def test_read_tree_complete(collection):
    result = controller_for(tree(5, collection)).read_tree(TREE_ID, max_nodes=5)
    assert result["available_nodes"] == 5
    assert result["truncated"] is False


def test_search_stopped_early():
    result = controller_for(tree(5)).search_tree_nodes(
        TREE_ID, "node", max_results=2,
    )
    assert result["total_matches"] == 2
    assert result["searched_nodes"] == 2
    assert result["total_nodes"] == 5
    assert result["stopped_early"] is True


def test_search_exactly_max_results_at_the_end_is_not_early():
    texts = lambda i: "match" if i >= 3 else "other"  # noqa: E731
    result = controller_for(tree(5, text=texts)).search_tree_nodes(
        TREE_ID, "match", max_results=2,
    )
    assert result["total_matches"] == 2
    assert result["searched_nodes"] == 5
    assert result["stopped_early"] is False


def test_search_below_cap():
    result = controller_for(tree(5)).search_tree_nodes(
        TREE_ID, "node", max_results=20,
    )
    assert result["total_matches"] == 5
    assert result["stopped_early"] is False
