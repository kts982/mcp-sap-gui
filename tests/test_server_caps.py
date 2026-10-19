"""The server says when it lowered a caller's limit to max_table_rows."""

import pytest

server = pytest.importorskip("mcp_sap_gui.server")


def test_capped_limit_is_reported():
    assert server._note_cap({"nodes": []}, 1000, 500) == {
        "nodes": [], "capped_to": 500,
    }


def test_uncapped_limit_adds_nothing():
    assert server._note_cap({"nodes": []}, 200, 200) == {"nodes": []}


def test_non_dict_result_passes_through():
    assert server._note_cap(None, 1000, 500) is None
