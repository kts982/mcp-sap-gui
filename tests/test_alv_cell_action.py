"""Actions on an ALV cell, and column selection.

Each check comes from a live run (SAP GUI 8.10, the ALV test programs
BCALV_TEST_GRID, BCALV_EDIT_03, BCALV_EDIT_05, BCALV_GRID_DEMO):

- Click on a cell that is no hotspot and PressButton on a cell that is no
  button raise a bare "Exception occurred".
- ModifyCheckBox on a cell that is no checkbox returns without an error and
  changes nothing.
- PressF4 on a cell without value help opens an information popup.
- Click or PressButton on a row that is not on screen only scrolls there;
  the event fires on the second call.
- PressEnter after ModifyCell with a wrong value puts the cell in state
  Error; PressTotalRow closes a subtotal's group and the grid loses rows.
"""

import pytest

from mcp_sap_gui import confirmation
from tests.fakes import controller_for

GRID_ID = "wnd[0]/usr/cntlGRID/shellcont/shell"
COLUMNS = ["CARRID", "CONNID", "FLAG", "PRICE"]


class Collection:
    def __init__(self, items):
        self._items = items

    @property
    def Count(self):  # noqa: N802 - COM name
        return len(self._items)

    def __call__(self, i):
        return self._items[i]


class MenuItem:
    def __init__(self, name, text):
        self.Name, self.Text = name, text
        self.Children = Collection([])


class FakeGrid:
    """A grid that records what is sent to it."""

    Type = "GuiShell"

    def __init__(self, cells=None, rows=5, levels=None):
        self.cells = cells or {}
        self.levels = levels or {}
        self.calls = []
        self.selected = []
        self.RowCount = rows
        self.ColumnCount = len(COLUMNS)
        self.Title = ""
        self.SelectedRows = ""
        self.SelectedCells = Collection([])
        self.CurrentContextMenu = type("Menu", (), {"Children": Collection([])})()
        self.close_on_total = 0
        self.total_columns = ["PRICE"]
        self.nothing_to_toggle = False

    @property
    def SelectedColumns(self):  # noqa: N802
        return Collection(list(self.selected))

    def ColumnOrder(self, i):  # noqa: N802
        return COLUMNS[i]

    def _cell(self, row, col):
        return self.cells.get((row, col), {})

    def IsCellHotspot(self, row, col):  # noqa: N802
        return self._cell(row, col).get("hotspot", False)

    def GetCellType(self, row, col):  # noqa: N802
        return self._cell(row, col).get("type", "Normal")

    def GetCellChangeable(self, row, col):  # noqa: N802
        return self._cell(row, col).get("changeable", False)

    def HasCellF4Help(self, row, col):  # noqa: N802
        return self._cell(row, col).get("f4", False)

    def GetCellCheckBoxChecked(self, row, col):  # noqa: N802
        return self._cell(row, col).get("checked", False)

    def GetCellValue(self, row, col):  # noqa: N802
        return self._cell(row, col).get("value", "")

    def GetCellState(self, row, col):  # noqa: N802
        return self._cell(row, col).get("state", "Normal")

    def GetRowTotalLevel(self, row):  # noqa: N802
        return self.levels.get(row, 0)

    def GetColumnTotalType(self, col):  # noqa: N802
        return "Total" if col in self.total_columns else "None"

    def SetCurrentCell(self, row, col):  # noqa: N802
        self.calls.append(("SetCurrentCell", row, col))

    def Click(self, row, col):  # noqa: N802
        self.calls.append(("Click", row, col))

    def PressButton(self, row, col):  # noqa: N802
        self.calls.append(("PressButton", row, col))

    def PressF4(self):  # noqa: N802
        self.calls.append(("PressF4",))

    def PressEnter(self):  # noqa: N802
        self.calls.append(("PressEnter",))

    def ContextMenu(self):  # noqa: N802
        self.calls.append(("ContextMenu",))
        self.CurrentContextMenu.Children = Collection([
            MenuItem("&DETAIL", "Details"), MenuItem("", ""),
            MenuItem("&FILTER", "Set Filter..."),
        ])

    def PressTotalRow(self, row, col):  # noqa: N802
        # Seen live: only a column with a total (or the line's own subtotal
        # column) is taken, and an overall total without subtotals raises.
        if col not in self.total_columns or self.nothing_to_toggle:
            raise RuntimeError("Exception occurred.")
        self.calls.append(("PressTotalRow", row, col))
        self.RowCount -= self.close_on_total

    def ModifyCheckBox(self, row, col, value):  # noqa: N802
        self.calls.append(("ModifyCheckBox", row, col, value))
        cell = self.cells.setdefault((row, col), {})
        if not cell.get("stuck"):
            cell["checked"] = value

    def ClearSelection(self):  # noqa: N802
        self.calls.append(("ClearSelection",))
        self.selected = []

    def SelectColumn(self, name):  # noqa: N802
        self.calls.append(("SelectColumn", name))
        self.selected.append(name)


def act(grid, action, **kw):
    controller = controller_for(grid)
    controller.get_screen_info = lambda: {"active_window": "wnd[0]"}
    return controller.alv_cell_action(GRID_ID, action, **kw)


# -- click ---------------------------------------------------------------------

def test_a_hotspot_is_made_current_and_clicked():
    grid = FakeGrid({(3, "CARRID"): {"hotspot": True}})

    result = act(grid, "click", row=3, column="CARRID")

    # Current cell first: a click on a row that is not on screen only scrolls.
    assert grid.calls == [("SetCurrentCell", 3, "CARRID"), ("Click", 3, "CARRID")]
    assert result["status"] == "clicked"
    assert result["screen"] == {"active_window": "wnd[0]"}


def test_a_click_on_a_cell_that_is_no_hotspot_is_refused():
    grid = FakeGrid()

    result = act(grid, "click", row=1, column="PRICE")

    assert grid.calls == []  # nothing was sent
    assert result["status"] == "rejected"
    assert "not a hotspot" in result["error"]
    assert result["cell_type"] == "Normal"
    assert "screen" not in result


def test_a_grid_that_cannot_say_is_not_held_back():
    class OldGrid(FakeGrid):
        def IsCellHotspot(self, row, col):  # noqa: N802
            raise RuntimeError("member not found")

    grid = OldGrid()

    assert act(grid, "click", row=1, column="PRICE")["status"] == "clicked"
    assert ("Click", 1, "PRICE") in grid.calls


# -- press_button --------------------------------------------------------------

def test_a_button_cell_is_pressed():
    grid = FakeGrid({(2, "CONNID"): {"type": "Button"}})

    result = act(grid, "press_button", row=2, column="CONNID")

    assert grid.calls == [("SetCurrentCell", 2, "CONNID"), ("PressButton", 2, "CONNID")]
    assert result["status"] == "pressed"


def test_a_cell_that_is_no_button_is_refused():
    grid = FakeGrid()

    result = act(grid, "press_button", row=2, column="PRICE")

    assert grid.calls == []
    assert result["status"] == "rejected" and "not a button" in result["error"]


# -- set_checkbox --------------------------------------------------------------

def test_a_checkbox_is_set_and_read_back():
    grid = FakeGrid({(0, "FLAG"): {"type": "CheckBox", "changeable": True}})

    result = act(grid, "set_checkbox", row=0, column="FLAG", checked=True)

    assert grid.calls == [("ModifyCheckBox", 0, "FLAG", True)]
    assert result["status"] == "success" and result["checked"] is True
    assert "enter" in result["note"]
    assert "screen" not in result  # nothing reached the application yet


def test_a_cell_that_is_no_checkbox_is_refused():
    # ModifyCheckBox would return without an error and change nothing.
    grid = FakeGrid()

    result = act(grid, "set_checkbox", row=0, column="CARRID", checked=True)

    assert grid.calls == []
    assert result["status"] == "rejected" and "not a checkbox" in result["error"]


def test_a_locked_checkbox_is_refused():
    grid = FakeGrid({(0, "FLAG"): {"type": "CheckBox", "changeable": False}})

    result = act(grid, "set_checkbox", row=0, column="FLAG", checked=True)

    assert grid.calls == []
    assert "not changeable" in result["error"]


def test_set_checkbox_needs_the_state():
    grid = FakeGrid({(0, "FLAG"): {"type": "CheckBox", "changeable": True}})

    result = act(grid, "set_checkbox", row=0, column="FLAG")

    assert grid.calls == []
    assert "needs checked" in result["error"]


def test_a_checkbox_that_did_not_take_the_value_is_not_reported_as_set():
    grid = FakeGrid({(0, "FLAG"): {"type": "CheckBox", "changeable": True,
                                   "checked": False, "stuck": True}})

    result = act(grid, "set_checkbox", row=0, column="FLAG", checked=True)

    assert result["status"] == "rejected"
    assert result["checked"] is False
    assert "did not take" in result["error"]


# -- f4, enter, context_menu, toggle_total ---------------------------------------

def test_f4_opens_the_value_help_of_the_cell():
    grid = FakeGrid({(1, "CARRID"): {"f4": True}})

    result = act(grid, "f4", row=1, column="CARRID")

    assert grid.calls == [("SetCurrentCell", 1, "CARRID"), ("PressF4",)]
    assert result["status"] == "opened"


def test_f4_on_a_cell_without_value_help_is_refused():
    grid = FakeGrid()

    result = act(grid, "f4", row=1, column="PRICE")

    assert grid.calls == []
    assert "no value help" in result["error"]


def test_enter_needs_no_cell():
    grid = FakeGrid()

    result = act(grid, "enter")

    assert grid.calls == [("PressEnter",)]
    assert result["status"] == "sent"
    assert "state" not in result


def test_enter_with_a_cell_reports_what_the_application_made_of_it():
    grid = FakeGrid({(2, "CARRID"): {"value": "ZZ", "state": "Error"}})

    result = act(grid, "enter", row=2, column="CARRID")

    assert grid.calls == [("PressEnter",)]
    assert result["value"] == "ZZ" and result["state"] == "Error"


def test_context_menu_lists_its_items():
    grid = FakeGrid()

    result = act(grid, "context_menu", row=2, column="PRICE")

    assert grid.calls == [("SetCurrentCell", 2, "PRICE"), ("ContextMenu",)]
    assert result["status"] == "menu_opened"
    assert result["menu_items"] == [
        {"text": "Details", "function_code": "&DETAIL"},
        {"text": "Set Filter...", "function_code": "&FILTER"},
    ]


def test_toggle_total_needs_only_the_row_and_says_that_the_rows_changed():
    grid = FakeGrid(rows=40, levels={9: 2})
    grid.close_on_total = 14

    result = act(grid, "toggle_total", row=9)

    # Pressed in the column that carries the total, which the tool finds.
    assert grid.calls == [("SetCurrentCell", 9, "PRICE"), ("PressTotalRow", 9, "PRICE")]
    assert result["status"] == "toggled"
    assert (result["rows_before"], result["rows_after"]) == (40, 26)
    assert "read the table again" in result["note"]


def test_toggle_total_does_not_press_in_a_column_without_a_total():
    grid = FakeGrid(rows=40, levels={9: 2})

    result = act(grid, "toggle_total", row=9, column="CARRID")

    assert ("PressTotalRow", 9, "PRICE") in grid.calls
    assert result["status"] == "toggled"


def test_toggle_total_on_a_record_is_refused():
    grid = FakeGrid(rows=40, levels={9: 2})

    result = act(grid, "toggle_total", row=3)

    assert grid.calls == []
    assert "not a total or subtotal" in result["error"]


def test_an_overall_total_without_subtotals_has_nothing_to_toggle():
    grid = FakeGrid(rows=31, levels={30: 1})
    grid.nothing_to_toggle = True

    result = act(grid, "toggle_total", row=30)

    assert result["status"] == "rejected"
    assert "no group to open or close" in result["error"]


def test_toggle_total_needs_the_row():
    result = act(FakeGrid(), "toggle_total")

    assert "needs the row" in result["error"]


# -- what every action checks ----------------------------------------------------

@pytest.mark.parametrize("kwargs, text", [
    ({"row": 5, "column": "PRICE"}, "rows 0 to 4"),
    ({"row": -1, "column": "PRICE"}, "rows 0 to 4"),
    ({"row": 1, "column": "NOPE"}, "is not in the grid"),
    ({"row": 1}, "needs row and column"),
    ({"column": "PRICE"}, "needs row and column"),
    ({}, "needs row and column"),
])
def test_a_cell_that_is_not_there_is_refused(kwargs, text):
    grid = FakeGrid({(1, "PRICE"): {"hotspot": True}})

    result = act(grid, "click", **kwargs)

    assert grid.calls == []
    assert result["status"] == "rejected" and text in result["error"]


def test_an_unknown_column_lists_the_grids_columns():
    result = act(FakeGrid(), "click", row=1, column="NOPE")

    assert result["available_columns"] == COLUMNS


def test_enter_with_half_a_cell_is_refused():
    grid = FakeGrid()

    assert "needs row and column" in act(grid, "enter", row=1)["error"]
    assert grid.calls == []


def test_an_unknown_action_is_refused():
    result = act(FakeGrid(), "triple_click", row=1, column="PRICE")

    assert result["status"] == "rejected" and "click, press_button" in result["error"]


def test_a_table_control_is_refused():
    class TableControl(FakeGrid):
        Type = "GuiTableControl"

    result = act(TableControl(), "click", row=1, column="PRICE")

    assert result["status"] == "rejected" and "table control" in result["error"]


# -- column selection ------------------------------------------------------------

def select(grid, columns, **kw):
    return controller_for(grid).select_alv_columns(GRID_ID, columns, **kw)


def test_columns_are_selected_in_place_of_what_was_selected():
    grid = FakeGrid()
    grid.selected = ["CARRID"]

    result = select(grid, ["PRICE", "CONNID"])

    assert grid.calls == [("ClearSelection",), ("SelectColumn", "PRICE"),
                          ("SelectColumn", "CONNID")]
    assert result == {"grid_id": GRID_ID, "status": "selected",
                      "selected_columns": ["PRICE", "CONNID"]}


def test_columns_can_be_added():
    grid = FakeGrid()
    grid.selected = ["CARRID"]

    result = select(grid, ["PRICE"], add=True)

    assert ("ClearSelection",) not in grid.calls
    assert result["selected_columns"] == ["CARRID", "PRICE"]


def test_an_empty_list_clears_the_selection():
    grid = FakeGrid()
    grid.selected = ["CARRID"]

    result = select(grid, [])

    assert grid.calls == [("ClearSelection",)]
    assert result["selected_columns"] == []


def test_an_unknown_column_selects_nothing():
    grid = FakeGrid()
    grid.selected = ["CARRID"]

    result = select(grid, ["PRICE", "NOPE"])

    assert grid.calls == []  # not even the clear
    assert result["status"] == "rejected"
    assert result["unknown_columns"] == ["NOPE"]
    assert result["available_columns"] == COLUMNS


# -- the confirmation point --------------------------------------------------------

def test_set_checkbox_falls_under_field_writes():
    args = {"grid_id": GRID_ID, "action": "set_checkbox", "row": 0,
            "column": "FLAG", "checked": True}

    assert confirmation.point_for_call(
        "sap_alv_cell_action", args, {"save", "field_writes"}, {"write"}) == "field_writes"


def test_the_other_actions_fall_under_all_writes_only():
    args = {"grid_id": GRID_ID, "action": "click", "row": 0, "column": "CARRID"}

    assert confirmation.point_for_call(
        "sap_alv_cell_action", args, {"save", "field_writes"}, {"write"}) is None
    assert confirmation.point_for_call(
        "sap_alv_cell_action", args, {"save", "all_writes"}, {"write"}) == "all_writes"
