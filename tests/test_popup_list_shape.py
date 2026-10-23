"""get_popup_window on a list-shaped popup: an F4 hit list, a note or a
message's long text all come back as ``list`` with a hint to sap_read_list,
and never with their lines in ``texts``.

Live (SAP GUI 8.10 and 8.00): SCAL's "Transport link note", a message's long
text and an F4 hit list all offer Cancel (F12). Only their button words
differ, so the response does not try to tell them apart.

The popups are fakes: a GetObjectTree JSON node for SAP GUI 7.70 PL3+, and
COM-style objects with ``Children`` for the older walk. No SAP GUI is used.
"""

import json

from mcp_sap_gui.sap_controller import SAPGUIController

SES = "/app/con[0]/ses[0]"


class _ComError(Exception):
    """Stands in for pywintypes.com_error."""


class _Controller(SAPGUIController):
    def _check_dependencies(self):  # no pywin32 needed for these tests
        pass


# --- GetObjectTree fakes ----------------------------------------------------

def _node(element_id, type_, text="", children=(), **props):
    properties = {"Id": f"{SES}/{element_id}", "Type": type_, "Text": text}
    properties.update(props)
    return {"properties": properties, "children": list(children)}


def _list_cell(row, col, text, kind="lbl", type_="GuiLabel"):
    return _node(f"wnd[1]/usr/{kind}[{col},{row}]", type_, text)


def _window(title, usr_children, buttons=("Continue",)):
    """``buttons``: texts at btn[0], btn[1], ..., or (index, text, tooltip)."""
    nodes = []
    for i, button in enumerate(buttons):
        index, text, tooltip = (button if isinstance(button, tuple)
                                else (i, button, ""))
        nodes.append(_node(f"wnd[1]/tbar[0]/btn[{index}]", "GuiButton", text,
                           Tooltip=tooltip))
    tbar = _node("wnd[1]/tbar[0]", "GuiToolbar", children=nodes)
    usr = _node("wnd[1]/usr", "GuiUserArea", children=usr_children)
    return _node("wnd[1]", "GuiModalWindow", title, [usr, tbar])


class _TreeSession:
    """A session with one popup, read through GetObjectTree."""

    Busy = False

    def __init__(self, window):
        self._window = window

    def findById(self, element_id):  # noqa: N802 (SAP GUI COM name)
        if element_id in ("wnd[0]", "wnd[1]"):
            return object()
        raise _ComError(f"not found: {element_id}")

    def GetObjectTree(self, root_id, props):  # noqa: N802 (SAP GUI COM name)
        assert root_id == "wnd[1]"
        return json.dumps({"children": [self._window]})


def _popup(window):
    controller = _Controller()
    controller._session = _TreeSession(window)
    return controller


LIST_KEYS = {"cells", "rows", "hint"}

# SCAL's "Transport link note" (SAPMSSY0 120), buttons as seen live on SAP
# GUI 8.00: Continue, Print, Find, Cancel. The wording here is made up.
NOTE = _window("Transport link note", [
    _list_cell(1, 1, "Changes to calendars are"),
    _list_cell(1, 26, "not transported automatically."),
    _list_cell(2, 1, ""),
    _list_cell(3, 1, "Include them in a request"),
    _list_cell(3, 27, "yourself."),
], buttons=((0, "", "Continue   (Enter)"), (86, "", "Print   (Ctrl+P)"),
            (71, "", "Find   (Ctrl+F)"), (12, "", "Cancel   (F12)")))

HIT_LIST = _window("Airline Code 2 Entries", [
    _list_cell(1, 1, "AA"),
    _list_cell(1, 4, "American Airlines"),
    _list_cell(2, 1, "AB"),
    _list_cell(2, 4, "Air Berlin"),
], buttons=((0, "", "Apply   (Enter)"), (17, "", "Restrict Values   (Shift+F5)"),
            (12, "", "Cancel   (F12)")))


def test_a_note_shown_as_a_list_points_to_sap_read_list():
    popup = _popup(NOTE).get_popup_window()

    assert "texts" not in popup
    assert set(popup["list"]) == LIST_KEYS
    assert popup["list"]["cells"] == 5
    assert popup["list"]["rows"] == 3
    assert popup["list"]["hint"].startswith(
        "Classic list: its lines are not in texts. Read them with "
        "sap_read_list(window_id='wnd[1]')."
    )
    assert popup["classification"] == "list"
    assert "safe_auto_action" not in popup
    assert popup["has_inputs"] is False
    assert "interactive_elements" not in popup


def test_a_hit_list_gets_the_same_shape_and_how_to_pick_a_line():
    popup = _popup(HIT_LIST).get_popup_window()

    assert "texts" not in popup
    assert set(popup["list"]) == LIST_KEYS
    hint = popup["list"]["hint"]
    assert "sap_read_list(window_id='wnd[1]')" in hint
    assert "In an F4 hit list, pick a line" in hint
    assert "with_ids=true" in hint
    assert popup["classification"] == "list"


def test_list_words_do_not_reach_the_classifier():
    # "Error" and "Warning" in a list's lines must not make it an error or
    # warning popup: the lines are not part of its texts.
    window = _window("Message classes", [
        _list_cell(1, 1, "Error messages"),
        _list_cell(2, 1, "Warning messages"),
    ])

    popup = _popup(window).get_popup_window()

    assert popup["classification"] == "list"
    assert "safe_auto_action" not in popup
    assert "texts" not in popup


def test_other_texts_stay_and_list_lines_stay_out():
    window = _window("Note", [
        _node("wnd[1]/usr/lblHEADER", "GuiLabel", "Read this first"),
        _list_cell(1, 1, "A1"),
    ])

    popup = _popup(window).get_popup_window()

    assert popup["texts"] == ["Read this first"]
    assert popup["list"]["rows"] == 1


def test_checkbox_cells_are_counted_in_the_list():
    window = _window("Choose", [
        _list_cell(1, 1, "", kind="chk", type_="GuiCheckBox"),
        _list_cell(1, 3, "Option A"),
        _list_cell(2, 1, "", kind="chk", type_="GuiCheckBox"),
    ])

    popup = _popup(window).get_popup_window()

    assert popup["list"]["cells"] == 3
    assert popup["list"]["rows"] == 2
    assert "texts" not in popup
    assert "interactive_elements" not in popup


def test_message_popup_without_list_keeps_its_text():
    window = _window("Information", [
        _node("wnd[1]/usr/txtMESSTXT1", "GuiTextField", "Data was saved",
              Name="MESSTXT1", Changeable="false"),
    ], buttons=("Continue",))

    popup = _popup(window).get_popup_window()

    assert popup["texts"] == ["Data was saved"]
    assert "list" not in popup


def test_digest_carries_the_list_and_its_hint():
    digest = _popup(NOTE)._popup_digest()

    assert digest["classification"] == "list"
    assert "texts" not in digest
    assert digest["list"]["hint"].startswith("Classic list: its lines are")


# --- COM walk (SAP GUI before 7.70 PL3) -----------------------------------

class _Children:
    def __init__(self, items):
        self._items = list(items)

    @property
    def Count(self):  # noqa: N802 (SAP GUI COM name)
        return len(self._items)

    def __call__(self, i):
        return self._items[i]


class _Element:
    def __init__(self, element_id, type_, text="", children=()):
        self.Id = f"{SES}/{element_id}"
        self.Type = type_
        self.Text = text
        self.Tooltip = ""
        self.Children = _Children(children)


class _ComSession:
    Busy = False

    def __init__(self, elements):
        self._elements = elements

    def findById(self, element_id):  # noqa: N802 (SAP GUI COM name)
        if element_id in self._elements:
            return self._elements[element_id]
        raise _ComError(f"not found: {element_id}")

    def GetObjectTree(self, root_id, props):  # noqa: N802 (SAP GUI COM name)
        raise _ComError("GetObjectTree needs SAP GUI 7.70 PL3")


def test_com_walk_gives_the_same_shape():
    cells = [
        _Element("wnd[1]/usr/lbl[1,1]", "GuiLabel", "Changes to calendars are"),
        _Element("wnd[1]/usr/lbl[26,1]", "GuiLabel",
                 "not transported automatically."),
    ]
    usr = _Element("wnd[1]/usr", "GuiUserArea", children=cells)
    buttons = [
        _Element("wnd[1]/tbar[0]/btn[0]", "GuiButton", "Continue"),
        _Element("wnd[1]/tbar[0]/btn[12]", "GuiButton", "Cancel"),
    ]
    tbar = _Element("wnd[1]/tbar[0]", "GuiToolbar", children=buttons)
    window = _Element("wnd[1]", "GuiModalWindow", "Transport link note")
    controller = _Controller()
    controller._session = _ComSession({
        "wnd[0]": object(),
        "wnd[1]": window,
        "wnd[1]/usr": usr,
        "wnd[1]/tbar[0]": tbar,
    })

    popup = controller.get_popup_window()

    assert popup["title"] == "Transport link note"
    assert "texts" not in popup
    assert popup["list"]["rows"] == 1
    assert popup["list"]["hint"].startswith("Classic list: its lines are")
    assert popup["classification"] == "list"
