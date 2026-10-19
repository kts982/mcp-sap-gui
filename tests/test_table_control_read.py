"""read_table on a GuiTableControl: dropdown padding, start_row notes, paging keys.

Checkbox padding and empty_rows_from itself are covered in
test_sap_controller.py; these tests add the dropdown's empty key and paging.
"""

import pytest

from tests.fakes import TABLE_ID, FakeTableControl, check, combo, controller_for, text

NAMES = ["KEY", "ACTIVE", "MODE"]


def data_row(i: int, active: bool = True) -> list:
    return [text(f"K{i:03}"), check(active), combo("A")]


def padding_row() -> list:
    # What a padding row's cells read as: an unticked box, the dropdown's
    # empty entry (key " "), an empty text.
    return [text(""), check(False), combo(" ")]


def table(n_data: int, row_count: int, visible: int, **kw) -> FakeTableControl:
    rows = [data_row(i) for i in range(n_data)]
    rows += [padding_row() for _ in range(row_count - n_data)]
    return FakeTableControl(NAMES, rows, visible, **kw)


@pytest.mark.parametrize("object_tree", [False, True])
def test_padding_rows_with_checkbox_and_dropdown_are_not_data(object_tree):
    tc = table(3, row_count=10, visible=10)
    result = controller_for(tc, object_tree).read_table(TABLE_ID)
    assert [r["KEY"] for r in result["data"]] == ["K000", "K001", "K002"]
    assert result["rows_returned"] == 3
    assert result["empty_rows_from"] == 3
    assert result["has_more"] is False
    assert result["next_start_row"] == 3
    # The values of real rows are unchanged.
    assert result["data"][0]["ACTIVE"] is True
    assert result["data"][0]["MODE"] == "A"


def test_full_page_without_padding_has_no_stop_marker():
    tc = table(30, row_count=30, visible=10)
    result = controller_for(tc).read_table(TABLE_ID)
    assert result["rows_returned"] == 10
    assert "empty_rows_from" not in result
    assert result["has_more"] is True
    assert result["next_start_row"] == 10
    assert "start_row" not in result  # none given: the current position
    assert "start_row_note" not in result


def test_paging_with_next_start_row_reads_every_row_once():
    tc = table(30, row_count=30, visible=10)  # scrollbar maximum 20
    controller = controller_for(tc)
    keys, start, pages = [], 0, 0
    while True:
        page = controller.read_table(TABLE_ID, start_row=start)
        keys += [r["KEY"] for r in page["data"]]
        pages += 1
        if not page["has_more"]:
            break
        start = page["next_start_row"]
    assert keys == [f"K{i:03}" for i in range(30)]
    assert pages == 3


def test_last_page_with_padding_ends_paging():
    tc = table(25, row_count=30, visible=10)  # maximum 20
    controller = controller_for(tc)
    page = controller.read_table(TABLE_ID, start_row=20)
    assert page["first_visible_row"] == 20
    assert page["rows_returned"] == 5
    assert page["has_more"] is False
    assert page["empty_rows_from"] == 25


def test_max_rows_below_visible_rows_has_more():
    tc = table(30, row_count=30, visible=10, maximum=20, position=20)
    result = controller_for(tc).read_table(TABLE_ID, max_rows=4, start_row=20)
    assert result["rows_returned"] == 4
    assert result["has_more"] is True
    assert result["next_start_row"] == 24


def test_no_start_row_reads_current_position_without_a_note():
    tc = table(30, row_count=30, visible=10, position=5)
    result = controller_for(tc).read_table(TABLE_ID)
    assert result["first_visible_row"] == 5
    assert result["data"][0]["_absolute_row_index"] == 5
    assert result["next_start_row"] == 15
    assert "start_row_note" not in result


def test_start_row_past_maximum_is_clamped_and_says_so():
    tc = table(30, row_count=30, visible=10)  # maximum 20
    result = controller_for(tc).read_table(TABLE_ID, start_row=25)
    assert result["start_row"] == 25
    assert result["first_visible_row"] == 20
    assert "scrollbar maximum 20" in result["start_row_note"]
    assert result["has_more"] is False


def test_failed_scroll_is_reported():
    tc = table(30, row_count=30, visible=10, refuse_scroll=True)
    result = controller_for(tc).read_table(TABLE_ID, start_row=10)
    assert result["first_visible_row"] == 0
    assert "scroll failed" in result["start_row_note"]
    assert result["data"][0]["KEY"] == "K000"


def test_reached_start_row_has_no_note():
    tc = table(30, row_count=30, visible=10)
    result = controller_for(tc).read_table(TABLE_ID, start_row=10)
    assert result["first_visible_row"] == 10
    assert "start_row_note" not in result


def test_columns_only_carries_no_paging_keys():
    tc = table(3, row_count=10, visible=10)
    result = controller_for(tc).read_table(TABLE_ID, columns_only=True)
    assert result["columns_only"] is True
    assert "has_more" not in result
    assert "next_start_row" not in result


@pytest.mark.parametrize("value, blank", [
    (None, True), ("", True), (" ", True), (False, True),
    ("0", False), ("A", False), (True, False), (0, False),
])
def test_blank_cell(value, blank):
    from mcp_sap_gui.tables import _is_blank_cell
    assert _is_blank_cell(value) is blank
