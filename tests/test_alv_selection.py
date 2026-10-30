"""An ALV read says which list it is and what is selected.

Seen live in /SCWM/MON: two grids that differ only by their title, and a
toolbar menu ("More Methods") whose functions act on the selected rows.
"""

from tests.fakes import controller_for

GRID_ID = "wnd[0]/usr/cntlGRID/shellcont/shell"


class Collection:
    """GuiCollection-shaped: Count plus call-by-index."""

    def __init__(self, items):
        self._items = list(items)
        self.Count = len(self._items)

    def __call__(self, i):
        return self._items[i]


class FakeGrid:
    Type = "GuiGridView"

    def __init__(self, title="", rows="", cells=(), columns=()):
        self.Title = title
        self.SelectedRows = rows
        self.SelectedCells = Collection(cells)
        self.SelectedColumns = Collection(columns)
        self.ColumnCount = 2
        self.RowCount = 3
        self.VisibleRowCount = 10
        self.firstVisibleRow = 0

    def ColumnOrder(self, i):  # noqa: N802 - COM name
        return ["WHO", "TANUM"][i]

    def GetCellValue(self, row, col):  # noqa: N802
        return f"{col}{row}"

    def GetColumnTooltip(self, col):  # noqa: N802
        return col

    def GetDisplayedColumnTitle(self, col):  # noqa: N802
        return col


def test_read_names_the_list_and_its_selection():
    grid = FakeGrid(title="Warehouse Task ", rows="1,2,4-8,10",
                    cells=["0,WHO", "2,TANUM"], columns=["TANUM"])

    result = controller_for(grid).read_table(GRID_ID)

    assert result["title"] == "Warehouse Task"
    assert result["selected_rows"] == "1,2,4-8,10"  # SAP's own form
    assert result["selected_row_count"] == 8
    assert result["selected_cells"] == [{"row": 0, "column": "WHO"},
                                        {"row": 2, "column": "TANUM"}]
    assert result["selected_columns"] == ["TANUM"]


def test_nothing_selected_adds_nothing():
    result = controller_for(FakeGrid()).read_table(GRID_ID)

    for key in ("title", "selected_rows", "selected_row_count",
                "selected_cells", "selected_cell_count", "selected_columns"):
        assert key not in result, key


def test_schema_read_carries_the_selection_too():
    result = controller_for(FakeGrid(title="Inbound Delivery", rows="0")).read_table(
        GRID_ID, columns_only=True)

    assert result["title"] == "Inbound Delivery"
    assert result["selected_rows"] == "0"
    assert result["selected_row_count"] == 1
    assert result["data"] == []


def test_a_large_cell_selection_is_capped_with_its_count():
    cells = [f"{r},WHO" for r in range(150)]

    result = controller_for(FakeGrid(cells=cells)).read_table(GRID_ID)

    assert len(result["selected_cells"]) == 100
    assert result["selected_cell_count"] == 150
