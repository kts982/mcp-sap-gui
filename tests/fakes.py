"""COM-free stand-ins for the SAP GUI objects these tests touch."""

import re
from typing import Any, List, Optional, Tuple

from mcp_sap_gui.sap_controller import SAPGUIController

TABLE_ID = "wnd[0]/usr/tblSAPLTEST"

# A cell is (Type, value): value is Text, Selected (bool) or Key by type.
Cell = Tuple[str, Any]


def text(value: str) -> Cell:
    return ("GuiCTextField", value)


def check(selected: bool) -> Cell:
    return ("GuiCheckBox", selected)


def combo(key: str) -> Cell:
    return ("GuiComboBox", key)


class FakeCell:
    def __init__(self, cell: Cell, name: str, cell_id: str) -> None:
        self.Type, value = cell
        self.Name = name
        self.Id = cell_id
        self.Text = value if isinstance(value, str) else ""
        self.Selected = bool(value) if self.Type == "GuiCheckBox" else False
        self.Key = value if self.Type == "GuiComboBox" else ""


class FakeColumns:
    def __init__(self, titles: List[str]) -> None:
        self._titles = titles
        self.Count = len(titles)

    def __call__(self, i: int):
        return type("Col", (), {"Title": self._titles[i], "Tooltip": ""})()


class FakeScrollbar:
    def __init__(self, maximum: int, position: int = 0,
                 refuse: bool = False) -> None:
        self.Minimum = 0
        self.Maximum = maximum
        self._position = position
        self._refuse = refuse

    @property
    def Position(self) -> int:  # noqa: N802 - COM name
        return self._position

    @Position.setter
    def Position(self, value: int) -> None:  # noqa: N802 - COM name
        if self._refuse:
            raise RuntimeError("The control could not be changed")
        self._position = value


class FakeTableControl:
    """A GuiTableControl over absolute *rows*; GetCell takes VISIBLE rows."""

    Type = "GuiTableControl"
    TableFieldName = "TEST"

    def __init__(self, names: List[str], rows: List[List[Cell]],
                 visible: int, maximum: Optional[int] = None,
                 position: int = 0, refuse_scroll: bool = False) -> None:
        self.names = names
        self.rows = rows
        self.Id = "/app/con[0]/ses[0]/" + TABLE_ID
        self.RowCount = len(rows)
        self.VisibleRowCount = visible
        self.Columns = FakeColumns(names)
        if maximum is None:
            maximum = max(len(rows) - visible, 0)
        self.VerticalScrollbar = FakeScrollbar(maximum, position, refuse_scroll)

    def GetCell(self, row: int, col: int) -> FakeCell:  # noqa: N802
        absolute = self.VerticalScrollbar.Position + row
        if absolute >= len(self.rows):
            raise RuntimeError("no cell")
        return FakeCell(
            self.rows[absolute][col], self.names[col],
            f"/app/con[0]/ses[0]/{TABLE_ID}/ctxtF-{self.names[col]}[{col},{row}]",
        )

    def object_tree(self) -> dict:
        """What GetObjectTree returns for the control (string properties).

        No Selected: the server never asks for it (it crashes SAP GUI 8.00)
        and reads a checkbox cell over COM by its Id instead.
        """
        children = []
        for row in range(self.VisibleRowCount):
            for col in range(len(self.names)):
                try:
                    cell = self.GetCell(row, col)
                except RuntimeError:
                    continue
                children.append({"properties": {
                    "Id": cell.Id, "Type": cell.Type, "Text": cell.Text,
                    "Key": cell.Key,
                }})
        return {"properties": {"Id": self.Id}, "children": children}


_CELL_INDEX = re.compile(r"\[(\d+),(\d+)\]$")


class FakeSession:
    def __init__(self, element: Any) -> None:
        self.element = element

    def findById(self, element_id: str) -> Any:  # noqa: N802 - COM name
        # A cell ID ends in [column,visible_row]; anything else is the element.
        match = _CELL_INDEX.search(element_id)
        if match and hasattr(self.element, "GetCell"):
            col, row = int(match.group(1)), int(match.group(2))
            return self.element.GetCell(row, col)
        return self.element


def controller_for(element: Any, object_tree: bool = False) -> SAPGUIController:
    """A controller bound to *element*, without pywin32 or SAP GUI."""
    controller = SAPGUIController.__new__(SAPGUIController)
    controller._session = FakeSession(element)
    controller._require_session = lambda: None
    if object_tree:
        controller._object_tree = lambda root, props, **kw: element.object_tree()
    else:
        controller._object_tree = lambda root, props, **kw: None
    return controller
