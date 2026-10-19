"""Fake SAP GUI objects for controller tests (no COM)."""

from __future__ import annotations

from typing import Any, Dict, Optional

from mcp_sap_gui.sap_controller import SAPGUIController


class FakeField:
    """A GuiTextField that refuses writes the way pywin32 reports them.

    pywin32 swallows SAP's com_error on a failed property put and raises
    AttributeError "Property '<unknown>.text' can not be set." instead.
    """

    Type = "GuiTextField"

    def __init__(self, text: str = "", max_length: Optional[int] = 10,
                 changeable: Optional[bool] = True,
                 unreadable: tuple = ()) -> None:
        self._text = text
        self._max_length = max_length
        self._changeable = changeable
        self._unreadable = set(unreadable)
        self.writes = 0

    @property
    def MaxLength(self):  # noqa: N802 (SAP GUI COM name)
        if "MaxLength" in self._unreadable:
            raise AttributeError("MaxLength")
        return self._max_length

    @property
    def Changeable(self):  # noqa: N802 (SAP GUI COM name)
        if "Changeable" in self._unreadable:
            raise AttributeError("Changeable")
        return self._changeable

    @property
    def text(self) -> str:
        return self._text

    @text.setter
    def text(self, value: str) -> None:
        self.writes += 1
        too_long = (isinstance(self._max_length, int)
                    and len(value) > self._max_length)
        if self._changeable is False or too_long:
            raise AttributeError("Property '<unknown>.text' can not be set.")
        self._text = value

    Text = text


class FakeSession:
    Busy = False

    def __init__(self, elements: Dict[str, Any]) -> None:
        self.elements = elements

    def findById(self, element_id: str):  # noqa: N802 (SAP GUI COM name)
        try:
            return self.elements[element_id]
        except KeyError:
            raise Exception(f"The control could not be found by id. {element_id}")


def make_controller(elements: Dict[str, Any]) -> SAPGUIController:
    """A controller bound to a fake session holding *elements*."""
    controller = SAPGUIController()
    controller._session = FakeSession(elements)
    return controller
