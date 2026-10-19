"""get_popup_window: a list-shaped popup without Cancel (a note shown as a
list) also carries its lines in ``texts``; an F4 hit list does not.

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
    window = _window("Note", [
        _list_cell(1, 10, "second"),
        _list_cell(1, 1, "first"),
    ])

    popup = _popup(window).get_popup_window()

    # "first" starts at column 1 and "second" at column 10, as on screen.
    assert popup["texts"] == ["first    second"]


def test_long_list_is_capped_and_marked_truncated():
    window = _window("Log", [
        _list_cell(row, 1, f"line {row}") for row in range(1, 26)
    ])

    popup = _popup(window).get_popup_window()

    assert popup["texts"] == [f"line {row}" for row in range(1, 21)]
    assert popup["list"]["lines_in_texts"] == 20
    assert popup["list"]["truncated"] is True
    assert popup["list"]["rows"] == 25


def test_list_words_do_not_reclassify_a_list_popup():
    # Before this change the lines were not in texts, so "error" and
    # "warning" in a list never reached the classifier. They still do not.
    window = _window("Message classes", [
        _list_cell(1, 1, "Error messages"),
        _list_cell(2, 1, "Warning messages"),
    ])

    popup = _popup(window).get_popup_window()

    assert popup["classification"] == "list"
    assert "safe_auto_action" not in popup
    assert popup["texts"] == ["Error messages", "Warning messages"]


HIT_LIST_KEYS = {"cells", "rows", "hint"}


def test_hit_list_with_cancel_keeps_its_lines_out_of_texts():
    # An F4 hit list is a choice: its rows are data for sap_read_list, not
    # message text. Its response is the one 0.5.0 gives.
    window = _window("Airline Code 3 Entries", [
        _list_cell(1, 1, "AA"),
        _list_cell(1, 4, "American Airlines"),
        _list_cell(2, 1, "AB"),
        _list_cell(2, 4, "Air Berlin"),
    ], buttons=((0, "", "Apply   (Enter)"), (12, "", "Cancel   (F12)")))

    popup = _popup(window).get_popup_window()

    assert "texts" not in popup
    assert set(popup["list"]) == HIT_LIST_KEYS
    assert popup["list"]["hint"].startswith(
        "Classic list (e.g. an F4 hit list): read it with "
        "sap_read_list(window_id='wnd[1]')."
    )
    assert popup["classification"] == "list"


def test_cancel_is_recognised_by_f12_in_any_language():
    # German logon: the tooltip does not say "Cancel"; F12 is btn[12].
    window = _window("Fluggesellschaft 1 Eintrag", [
        _list_cell(1, 1, "LH"),
    ], buttons=((0, "", "Übernehmen   (Enter)"),
                (12, "", "Abbrechen   (F12)")))

    popup = _popup(window).get_popup_window()

    assert "texts" not in popup
    assert set(popup["list"]) == HIT_LIST_KEYS


def test_cancel_button_is_recognised_without_f12():
    # A Cancel-type word at another position (French logon).
    window = _window("Compagnie aérienne 1 entrée", [
        _list_cell(1, 1, "LH"),
    ], buttons=((0, "", "Valider   (Entrée)"), (5, "Annuler", "")))

    popup = _popup(window).get_popup_window()

    assert "texts" not in popup
    assert set(popup["list"]) == HIT_LIST_KEYS


def test_a_word_containing_no_or_back_is_not_cancel():
    # Substring matching read "Display Note" (no) and "Background" (back)
    # as Cancel, and the note's text was left out again.
    window = _window("Transport link note", [
        _list_cell(1, 1, "Changes are not transported."),
    ], buttons=((0, "", "Continue   (Enter)"), (5, "Display Note", ""),
                (6, "Background", "")))

    popup = _popup(window).get_popup_window()

    assert popup["texts"] == ["Changes are not transported."]
    assert popup["list"]["lines_in_texts"] == 1


def test_a_single_button_only_acknowledges():
    # One button, even one called Close or Cancel: nothing to choose.
    window = _window("Note", [
        _list_cell(1, 1, "Read me."),
    ], buttons=((12, "", "Close   (F12)"),))

    popup = _popup(window).get_popup_window()

    assert popup["texts"] == ["Read me."]


def test_other_texts_come_before_the_list_lines():
    window = _window("Note", [
        _node("wnd[1]/usr/lblHEADER", "GuiLabel", "Read this first"),
        _list_cell(1, 1, "A1"),
    ])

    popup = _popup(window).get_popup_window()

    assert popup["texts"] == ["Read this first", "A1"]


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
    assert popup["list"]["hint"] == (
        "Classic list in a popup without Cancel (a note or message shown as "
        "a list). Read it with sap_read_list(window_id='wnd[1]')."
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

