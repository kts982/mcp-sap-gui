"""An ALV read marks the totals lines of the layout.

A total or subtotal counts as a row but is a sum, not a record. Seen live
on BCALV_GRID_DEMO with a total on one column and subtotals on two: 40
rows for 30 records, levels 1 (innermost subtotal) to 3 (overall total);
with "totals lines above the entries" the overall total is row 0; a
closed group drops its records and inner subtotals from the grid. The
counter row of /SCWM/MON is the same kind of row (level 1, last).
"""

from tests.fakes import controller_for

GRID_ID = "wnd[0]/usr/cntlGRID/shellcont/shell"


class FakeGrid:
    """A grid whose rows are given as total levels: 0 is a record."""

    Type = "GuiGridView"

    def __init__(self, levels, collapsed=(), visible=50):
        self._levels = list(levels)
        self._collapsed = set(collapsed)
        self.level_calls = []
        self.ColumnCount = 2
        self.RowCount = len(self._levels)
        self.VisibleRowCount = visible
        self.firstVisibleRow = 0
        self.Title = ""
        self.SelectedRows = ""

    def ColumnOrder(self, i):  # noqa: N802 - COM name
        return ["CARRID", "SEATSOCC"][i]

    def GetCellValue(self, row, col):  # noqa: N802
        return f"{col}{row}"

    def GetColumnTooltip(self, col):  # noqa: N802
        return col

    def GetDisplayedColumnTitle(self, col):  # noqa: N802
        return col

    def GetRowTotalLevel(self, row):  # noqa: N802
        self.level_calls.append(row)
        return self._levels[row]

    def IsTotalRowExpanded(self, row):  # noqa: N802
        if self._levels[row] == 0:
            raise RuntimeError("not a totals line")  # as SAP GUI does
        return row not in self._collapsed


def marks(result):
    return {row["_absolute_row_index"]: row["_total_level"]
            for row in result["data"] if "_total_level" in row}


def test_a_grid_without_totals_costs_two_calls_and_adds_nothing():
    grid = FakeGrid([0] * 30)

    result = controller_for(grid).read_table(GRID_ID)

    assert grid.level_calls == [0, 29]
    assert "totals_lines" not in result
    assert all("_total_level" not in row for row in result["data"])
    assert "warning" not in result


def test_the_overall_total_at_the_end_is_marked():
    grid = FakeGrid([0, 0, 0, 1])

    result = controller_for(grid).read_table(GRID_ID)

    assert result["totals_lines"] == {"highest_level": 1, "above_entries": False}
    assert marks(result) == {3: 1}
    assert result["total_rows"] == 4  # still every row of the grid


def test_subtotals_get_their_level_and_the_overall_total_the_highest():
    # Two groups with one subtotal each, then the overall total.
    grid = FakeGrid([0, 0, 1, 0, 1, 2])

    result = controller_for(grid).read_table(GRID_ID)

    assert result["totals_lines"] == {"highest_level": 2, "above_entries": False}
    assert marks(result) == {2: 1, 4: 1, 5: 2}
    assert "_total_level" not in result["data"][0]


def test_totals_lines_above_the_entries():
    grid = FakeGrid([2, 1, 0, 0, 1, 0])

    result = controller_for(grid).read_table(GRID_ID)

    assert result["totals_lines"] == {"highest_level": 2, "above_entries": True}
    assert marks(result) == {0: 2, 1: 1, 4: 1}


def test_a_page_without_totals_lines_still_says_the_grid_has_them():
    grid = FakeGrid([0] * 9 + [1])

    result = controller_for(grid).read_table(GRID_ID, max_rows=3)

    assert result["totals_lines"] == {"highest_level": 1, "above_entries": False}
    assert marks(result) == {}
    # The two rows that decide, then only the rows of this page.
    assert grid.level_calls == [0, 9, 1, 2]


def test_a_closed_group_is_marked_and_warned_about():
    # Group one is closed: only its subtotal is left in the grid.
    grid = FakeGrid([1, 0, 0, 1, 2], collapsed=[0])

    result = controller_for(grid).read_table(GRID_ID)

    assert result["data"][0]["_collapsed"] is True
    assert "_collapsed" not in result["data"][3]
    assert "1 subtotal line(s) in this read are collapsed" in result["warning"]


def test_every_group_closed_with_totals_below():
    grid = FakeGrid([1, 1, 2], collapsed=[0, 1])

    result = controller_for(grid).read_table(GRID_ID)

    assert result["totals_lines"] == {"highest_level": 2, "above_entries": False}
    assert marks(result) == {0: 1, 1: 1, 2: 2}


def test_schema_read_says_the_grid_has_totals_lines():
    grid = FakeGrid([0, 0, 1])

    result = controller_for(grid).read_table(GRID_ID, columns_only=True)

    assert result["totals_lines"] == {"highest_level": 1, "above_entries": False}
    assert result["data"] == []


def test_a_grid_that_cannot_answer_reads_as_before():
    class OldGrid(FakeGrid):
        def GetRowTotalLevel(self, row):  # noqa: N802
            raise RuntimeError("not supported")

    result = controller_for(OldGrid([0, 0, 1])).read_table(GRID_ID)

    assert "totals_lines" not in result
    assert all("_total_level" not in row for row in result["data"])


def test_a_single_row_grid():
    assert "totals_lines" not in controller_for(FakeGrid([0])).read_table(GRID_ID)

    result = controller_for(FakeGrid([1])).read_table(GRID_ID)

    assert result["totals_lines"] == {"highest_level": 1, "above_entries": False}
    assert marks(result) == {0: 1}
