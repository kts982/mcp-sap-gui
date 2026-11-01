"""A table read says what it could not read.

- Column names the table does not have are no longer dropped without a word:
  none known is an error that lists the table's names, some unknown are
  listed in unknown_columns.
- A table-control row whose cells all fail to read is not taken for the end
  of the data when rows further down the page have data (a padding row has
  no cells either, so only what follows tells the two apart).
- An ALV cell whose GetCellValue raises is listed in unreadable_cells.
"""

import pytest

from tests.fakes import TABLE_ID, FakeTableControl, controller_for, text

NAMES = ["KEY", "TEXT"]
GRID_ID = "wnd[0]/usr/cntlGRID/shellcont/shell"


def row(key: str, value: str = "x") -> list:
    return [text(key), text(value)]


class UnreadableRows(FakeTableControl):
    """GetCell raises for the absolute rows in *broken* (as for a padding row,
    which has no cells): over COM and in the GetObjectTree read alike."""

    def __init__(self, rows, visible, broken, **kw):
        super().__init__(NAMES, rows, visible, **kw)
        self.broken = set(broken)

    def GetCell(self, row_idx, col):  # noqa: N802 - COM name
        if self.VerticalScrollbar.Position + row_idx in self.broken:
            raise RuntimeError("cell not readable")
        return super().GetCell(row_idx, col)


# ---- table control: rows that could not be read --------------------------


@pytest.mark.parametrize("object_tree", [False, True])
def test_unreadable_row_between_data_rows_is_not_the_end(object_tree):
    # Codex's case: A, B unreadable, C. Before, only A came back, with
    # has_more false and empty_rows_from 1.
    tc = UnreadableRows([row("A"), row("B"), row("C")], visible=3, broken={1})
    result = controller_for(tc, object_tree).read_table(TABLE_ID)

    assert [r["KEY"] for r in result["data"]] == ["A", None, "C"]
    assert result["data"][1]["_absolute_row_index"] == 1
    assert result["unreadable_rows"] == [1]
    assert "No cell of row(s) 1 could be read" in result["warning"]
    assert "empty_rows_from" not in result
    assert result["has_more"] is False  # all three rows were read


@pytest.mark.parametrize("object_tree", [False, True])
def test_padding_after_the_data_still_ends_it(object_tree):
    # Rows 5..9 have no cells (past the data): the end, as before.
    tc = UnreadableRows([row(f"K{i}") for i in range(5)] + [row("")] * 5,
                        visible=10, broken=set(range(5, 10)))
    result = controller_for(tc, object_tree).read_table(TABLE_ID)

    assert result["rows_returned"] == 5
    assert result["empty_rows_from"] == 5
    assert "unreadable_rows" not in result
    assert "warning" not in result


def test_blank_line_between_entries_is_kept_as_a_row():
    # A line with cells but no values, followed by data: a blank line, not
    # the end, and not unreadable.
    tc = FakeTableControl(NAMES, [row("A"), row("", ""), row("C"), row("", "")],
                          visible=4)
    result = controller_for(tc).read_table(TABLE_ID)

    assert [r["KEY"] for r in result["data"]] == ["A", "", "C"]
    assert result["empty_rows_from"] == 3
    assert "unreadable_rows" not in result


def test_unreadable_row_is_reported_with_a_column_filter_too():
    tc = UnreadableRows([row("A"), row("B"), row("C")], visible=3, broken={1})
    result = controller_for(tc).read_table(TABLE_ID, columns="TEXT")

    assert result["columns"] == ["TEXT"]
    assert [r["TEXT"] for r in result["data"]] == ["x", None, "x"]
    assert result["unreadable_rows"] == [1]


# ---- table control: unknown column names ----------------------------------


def test_table_control_with_no_known_column_is_an_error():
    tc = FakeTableControl(NAMES, [row("A")], visible=3)
    result = controller_for(tc).read_table(TABLE_ID, columns="KEY_TYPO")

    assert result["unknown_columns"] == ["KEY_TYPO"]
    assert result["available_columns"] == NAMES
    assert "None of the requested columns" in result["error"]
    assert "KEY, TEXT" in result["error"]
    assert "data" not in result


def test_table_control_with_some_unknown_columns_reads_the_rest():
    tc = FakeTableControl(NAMES, [row("A")], visible=3)
    result = controller_for(tc).read_table(TABLE_ID, columns="KEY,NOPE")

    assert result["columns"] == ["KEY"]
    assert result["data"][0]["KEY"] == "A"
    assert result["unknown_columns"] == ["NOPE"]
    assert "left out: NOPE" in result["warning"]


# ---- ALV ------------------------------------------------------------------


class Grid:
    Type = "GuiGridView"
    Title = ""
    SelectedRows = ""

    def __init__(self, failing=()):
        self.failing = set(failing)  # (row, column) whose read raises
        self.ColumnCount = 2
        self.RowCount = 3
        self.VisibleRowCount = 10
        self.firstVisibleRow = 0

    def ColumnOrder(self, i):  # noqa: N802 - COM name
        return ["CARRID", "CONNID"][i]

    def GetCellValue(self, r, col):  # noqa: N802
        if (r, col) in self.failing:
            raise RuntimeError("cell not readable")
        return f"{col}{r}"

    def GetColumnTooltip(self, col):  # noqa: N802
        return col

    def GetDisplayedColumnTitle(self, col):  # noqa: N802
        return col


def test_alv_with_no_known_column_is_an_error():
    # Live on 2026-10-31 (SFLIGHT): columns=CARRID_REVIEW_TYPO returned two
    # rows holding only their index, and no word about the name.
    result = controller_for(Grid()).read_table(GRID_ID, columns="CARRID_REVIEW_TYPO")

    assert result["table_type"] == "GuiGridView"
    assert result["unknown_columns"] == ["CARRID_REVIEW_TYPO"]
    assert result["available_columns"] == ["CARRID", "CONNID"]
    assert "None of the requested columns" in result["error"]
    assert "data" not in result


def test_alv_with_some_unknown_columns_reads_the_rest():
    result = controller_for(Grid()).read_table(GRID_ID, columns="CONNID,CARRID_X")

    assert result["columns"] == ["CONNID"]
    assert result["data"][0] == {"CONNID": "CONNID0", "_absolute_row_index": 0}
    assert result["unknown_columns"] == ["CARRID_X"]
    assert "left out: CARRID_X" in result["warning"]


def test_alv_known_columns_have_no_unknown_key():
    result = controller_for(Grid()).read_table(GRID_ID, columns="CARRID")
    assert "unknown_columns" not in result
    assert "warning" not in result


def test_alv_cell_that_raises_is_listed():
    result = controller_for(Grid(failing={(1, "CONNID")})).read_table(GRID_ID)

    assert result["data"][1] == {"CARRID": "CARRID1", "CONNID": None,
                                 "_absolute_row_index": 1}
    assert result["unreadable_cells"] == [{"row": 1, "column": "CONNID"}]
    assert result["unreadable_cell_count"] == 1
    assert "1 cell(s) could not be read" in result["warning"]


def test_alv_unreadable_cells_are_capped_but_counted():
    failing = {(r, c) for r in range(15) for c in ("CARRID", "CONNID")}
    grid = Grid(failing=failing)
    grid.RowCount = 15
    result = controller_for(grid).read_table(GRID_ID)

    assert result["unreadable_cell_count"] == 30
    assert len(result["unreadable_cells"]) == 20
    assert result["unreadable_cells"][0] == {"row": 0, "column": "CARRID"}
