"""get_popup_window: a list-shaped popup also carries its lines in ``texts``.

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
    tbar = _node("wnd[1]/tbar[0]", "GuiToolbar", children=[
        _node(f"wnd[1]/tbar[0]/btn[{i}]", "GuiButton", text)
        for i, text in enumerate(buttons)
    ])
    usr = _node("wnd[1]/usr", "GuiUserArea", children=usr_children)
    return _node("wnd[1]", "GuiModalWindow", title, [usr, tbar])


class _TreeSession:
    """A session with one popup, read through GetObjectTree."""

    Busy = False

    def __init__(self, window):
        self._window = window

    def findById(self, element_id):
        if element_id in ("wnd[0]", "wnd[1]"):
            return object()
        raise _ComError(f"not found: {element_id}")

    def GetObjectTree(self, root_id, props):
        assert root_id == "wnd[1]"
        return json.dumps({"children": [self._window]})


def _popup(window):
    controller = _Controller()
    controller._session = _TreeSession(window)
    return controller


# SCAL's "Transport link note" (SAPMSSY0 120) came back in 0.5.0 as
# list {cells, rows} and no text. The wording here is made up.
NOTE = _window("Transport link note", [
    _list_cell(1, 1, "Changes to calendars are"),
    _list_cell(1, 26, "not transported automatically."),
    _list_cell(2, 1, ""),
    _list_cell(3, 1, "Include them in a request"),
    _list_cell(3, 27, "yourself."),
])


def test_list_popup_has_its_lines_in_texts():
    popup = _popup(NOTE).get_popup_window()

    assert popup["texts"] == [
        "Changes to calendars are not transported automatically.",
        "Include them in a request yourself.",
    ]
    assert popup["list"]["lines_in_texts"] == 2
    assert popup["list"]["truncated"] is False
    assert "Its first lines are in texts" in popup["list"]["hint"]
    assert "sap_read_list(window_id='wnd[1]')" in popup["list"]["hint"]


def test_list_popup_keeps_existing_keys_and_classification():
    popup = _popup(NOTE).get_popup_window()

    assert popup["list"]["cells"] == 5
    assert popup["list"]["rows"] == 3
    assert popup["classification"] == "list"
    assert "safe_auto_action" not in popup
    assert popup["has_inputs"] is False
    assert "interactive_elements" not in popup


def test_lines_follow_columns_as_read_list_does():
    window = _window("Hit list", [
        _list_cell(1, 10, "second"),
        _list_cell(1, 1, "first"),
    ])

    popup = _popup(window).get_popup_window()

    # "first" starts at column 1 and "second" at column 10, as on screen.
    assert popup["texts"] == ["first    second"]


def test_long_list_is_capped_and_marked_truncated():
    window = _window("Hit list", [
        _list_cell(row, 1, f"line {row}") for row in range(1, 26)
    ])

    popup = _popup(window).get_popup_window()

    assert popup["texts"] == [f"line {row}" for row in range(1, 21)]
    assert popup["list"]["lines_in_texts"] == 20
    assert popup["list"]["truncated"] is True
    assert popup["list"]["rows"] == 25


def test_list_words_do_not_reclassify_a_hit_list():
    # Before this change the lines were not in texts, so "error" and
    # "warning" in a hit list never reached the classifier. They still do not.
    window = _window("Message classes", [
        _list_cell(1, 1, "Error messages"),
        _list_cell(2, 1, "Warning messages"),
    ], buttons=("Copy", "Cancel"))

    popup = _popup(window).get_popup_window()

    assert popup["classification"] == "list"
    assert popup["texts"] == ["Error messages", "Warning messages"]


def test_other_texts_come_before_the_list_lines():
    window = _window("Hit list", [
        _node("wnd[1]/usr/lblHEADER", "GuiLabel", "Restrict values"),
        _list_cell(1, 1, "A1"),
    ])

    popup = _popup(window).get_popup_window()

    assert popup["texts"] == ["Restrict values", "A1"]


def test_checkbox_cells_are_counted_but_not_copied():
    window = _window("Choose", [
        _list_cell(1, 1, "", kind="chk", type_="GuiCheckBox"),
        _list_cell(1, 3, "Option A"),
        _list_cell(2, 1, "", kind="chk", type_="GuiCheckBox"),
    ])

    popup = _popup(window).get_popup_window()

    assert popup["list"]["cells"] == 3
    assert popup["texts"] == ["Option A"]


def test_list_without_text_cells_has_no_texts():
    window = _window("Choose", [
        _list_cell(1, 1, "", kind="chk", type_="GuiCheckBox"),
    ])

    popup = _popup(window).get_popup_window()

    assert "texts" not in popup
    assert popup["list"]["lines_in_texts"] == 0
    assert popup["list"]["truncated"] is False
    assert popup["list"]["hint"].startswith(
        "Classic list: an F4 hit list, or a message or log shown as a list. "
        "Read it with sap_read_list(window_id='wnd[1]')."
    )


def test_message_popup_without_list_is_unchanged():
    window = _window("Information", [
        _node("wnd[1]/usr/txtMESSTXT1", "GuiTextField", "Data was saved",
              Name="MESSTXT1", Changeable="false"),
    ], buttons=("Continue",))

    popup = _popup(window).get_popup_window()

    assert popup["texts"] == ["Data was saved"]
    assert "list" not in popup


def test_digest_shows_the_list_lines():
    digest = _popup(NOTE)._popup_digest()

    assert digest["classification"] == "list"
    assert digest["texts"][0] == (
        "Changes to calendars are not transported automatically."
    )
    assert digest["list"]["lines_in_texts"] == 2


# --- COM walk (SAP GUI before 7.70 PL3) -----------------------------------

class _Children:
    def __init__(self, items):
        self._items = list(items)

    @property
    def Count(self):
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

    def findById(self, element_id):
        if element_id in self._elements:
            return self._elements[element_id]
        raise _ComError(f"not found: {element_id}")

    def GetObjectTree(self, root_id, props):
        raise _ComError("GetObjectTree needs SAP GUI 7.70 PL3")


def test_com_walk_fills_texts_the_same_way():
    cells = [
        _Element("wnd[1]/usr/lbl[1,1]", "GuiLabel", "Changes to calendars are"),
        _Element("wnd[1]/usr/lbl[26,1]", "GuiLabel",
                 "not transported automatically."),
    ]
    usr = _Element("wnd[1]/usr", "GuiUserArea", children=cells)
    button = _Element("wnd[1]/tbar[0]/btn[0]", "GuiButton", "Continue")
    tbar = _Element("wnd[1]/tbar[0]", "GuiToolbar", children=[button])
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
    assert popup["texts"] == [
        "Changes to calendars are not transported automatically.",
    ]
    assert popup["list"]["rows"] == 1
    assert popup["classification"] == "list"

