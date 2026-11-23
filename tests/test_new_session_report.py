"""An action that opens another SAP session says so, and the binding stays.

Seen live (SAP GUI 8.10, System > New GUI Window through the menu): the
call returns after about 20 ms with the screen unchanged, and the new
session shows on the connection 0.5 to 0.7 s later. A hotspot that opens a
document in a window of its own behaves the same. Only /o was noticed
before; every other action looked as if nothing had happened.
"""

from unittest.mock import MagicMock

import pytest

from mcp_sap_gui import controller as controller_module
from mcp_sap_gui.models import VKey
from mcp_sap_gui.sap_controller import SAPGUIController


class Collection:
    def __init__(self, items):
        self.items = items

    @property
    def Count(self):  # noqa: N802 - COM name
        return len(self.items)

    def __call__(self, index):
        return self.items[index]


class OtherSession:
    """A session on the connection that the controller is not bound to."""

    def __init__(self, number, transaction="VA03", title="Display Sales Order"):
        self.Id = f"/app/con[0]/ses[{number}]"
        self.Info = MagicMock(Transaction=transaction)
        self._title = title

    def findById(self, element_id):  # noqa: N802
        return MagicMock(Text=self._title)


class FakeConnection:
    Id = "/app/con[0]"

    def __init__(self, sessions):
        self.sessions = list(sessions)
        self.Children = Collection(self.sessions)


def bound_session():
    session = MagicMock(Busy=False)
    session.Id = "/app/con[0]/ses[0]"
    session.Info.Transaction = "VA05"
    session.Info.Program = "SAPMV75A"
    session.Info.ScreenNumber = 100
    session.ActiveWindow.Id = "/app/con[0]/ses[0]/wnd[0]"
    session.ActiveWindow.Text = "List of Sales Orders"
    return session


@pytest.fixture
def setup(monkeypatch):
    """A controller bound to ses[0] of a connection, with no real waiting."""
    session = bound_session()
    connection = FakeConnection([session])
    controller = SAPGUIController()
    controller._session = session
    controller._connection = connection
    controller._application = engine = object()
    controller._require_session = lambda: None
    controller._object_tree = lambda root, props, **kw: None
    controller._discover_connections = lambda: [
        ("saplogon", "SAPGUI", object(), FakeConnection([])),
        ("saplogon", "SAPGUI", engine, connection),
    ]
    sleeps = []
    monkeypatch.setattr(controller_module.time, "sleep", sleeps.append)
    return controller, connection, session, sleeps


def test_the_first_look_reports_nothing(setup):
    controller, connection, _session, _sleeps = setup
    connection.sessions.append(OtherSession(1))  # was open before we looked

    assert "new_sessions" not in controller.get_screen_info()


def test_a_session_that_appeared_is_reported_and_the_binding_stays(setup):
    controller, connection, session, _sleeps = setup
    controller.get_screen_info()
    connection.sessions.append(OtherSession(1))

    screen = controller.get_screen_info()

    assert screen["new_sessions"] == [{
        "session_id": "/app/con[0]/ses[1]",
        "connection_index": 1,
        "session_index": 1,
        "transaction": "VA03",
        "title": "Display Sales Order",
    }]
    assert "still belong to the session you were in" in screen["new_sessions_note"]
    assert screen["transaction"] == "VA05"  # this session's screen, as before
    assert controller._session is session


def test_it_is_reported_once(setup):
    controller, connection, _session, _sleeps = setup
    controller.get_screen_info()
    connection.sessions.append(OtherSession(1))
    assert "new_sessions" in controller.get_screen_info()

    again = controller.get_screen_info()

    assert "new_sessions" not in again
    assert "new_sessions_note" not in again


def test_session_index_is_the_position_not_the_number(setup):
    controller, connection, _session, _sleeps = setup
    controller.get_screen_info()
    # A new session takes the lowest free ses[n]: here ses[3] at position 1.
    connection.sessions.append(OtherSession(3))

    entry = controller.get_screen_info()["new_sessions"][0]

    assert entry["session_id"].endswith("ses[3]")
    assert entry["session_index"] == 1


def test_a_session_closed_before_the_report_is_left_out(setup):
    controller, connection, _session, _sleeps = setup
    controller.get_screen_info()
    connection.sessions.append(OtherSession(1))
    controller._check_sessions()  # noticed, not yet reported
    connection.sessions.pop()

    assert "new_sessions" not in controller.get_screen_info()


def test_an_unknown_connection_index_is_left_out(setup):
    controller, connection, _session, _sleeps = setup
    controller._discover_connections = lambda: []
    controller.get_screen_info()
    connection.sessions.append(OtherSession(1))

    entry = controller.get_screen_info()["new_sessions"][0]

    assert "connection_index" not in entry
    assert entry["session_index"] == 1


# -- the wait after an action that can navigate ------------------------------

def test_an_unchanged_screen_waits_until_the_session_shows(setup):
    controller, connection, _session, sleeps = setup
    mark = controller._navigation_mark()

    def arrive(_seconds):
        sleeps.append(_seconds)
        if len(sleeps) == 3:
            connection.sessions.append(OtherSession(1))
    controller_module.time.sleep = arrive

    controller._await_new_session(mark)

    assert len(sleeps) == 3  # stopped as soon as it was there
    assert controller.get_screen_info()["new_sessions"][0]["transaction"] == "VA03"


def test_a_changed_screen_does_not_wait(setup):
    controller, _connection, session, sleeps = setup
    mark = controller._navigation_mark()
    session.Info.Transaction = "VA03"  # the action navigated in this session

    controller._await_new_session(mark)

    assert sleeps == []


def test_a_popup_does_not_wait(setup):
    controller, _connection, session, sleeps = setup
    mark = controller._navigation_mark()
    session.ActiveWindow.Id = "/app/con[0]/ses[0]/wnd[1]"

    controller._await_new_session(mark)

    assert sleeps == []


def test_nothing_new_gives_up_after_the_wait(setup, monkeypatch):
    controller, _connection, _session, sleeps = setup
    clock = iter(range(1000))
    monkeypatch.setattr(controller_module.time, "monotonic", lambda: next(clock) * 0.5)
    mark = controller._navigation_mark()

    controller._await_new_session(mark)

    assert 1 <= len(sleeps) <= 4  # 1.5 s at most, not for ever
    assert "new_sessions" not in controller.get_screen_info()


def test_an_unchanged_screen_with_nothing_new_says_so_once(setup):
    controller, _connection, session, _sleeps = setup
    controller._ACTION_NEW_SESSION_WAIT_S = 0
    session.findById.return_value = MagicMock()

    result = controller.press_button("wnd[0]/usr/btnREFRESH")

    note = result["screen"]["unchanged_note"]
    assert "did not change" in note and "new_sessions" in note
    assert "unchanged_note" not in controller.get_screen_info()


def test_a_changed_screen_carries_no_note(setup):
    controller, _connection, session, _sleeps = setup
    button = MagicMock()

    def navigate():
        session.Info.Transaction = "VA03"
    button.press.side_effect = navigate
    session.findById.return_value = button

    result = controller.press_button("wnd[0]/usr/btnDISPLAY")

    assert "unchanged_note" not in result["screen"]


def test_a_session_that_showed_carries_no_unchanged_note(setup):
    controller, connection, session, _sleeps = setup
    button = MagicMock()
    button.press.side_effect = lambda: connection.sessions.append(OtherSession(1))
    session.findById.return_value = button

    screen = controller.press_button("wnd[0]/usr/btnDISPLAY")["screen"]

    assert "new_sessions" in screen and "unchanged_note" not in screen


def test_the_caller_can_wait_longer(setup, monkeypatch):
    controller, _connection, _session, sleeps = setup
    clock = iter(range(1000))
    monkeypatch.setattr(controller_module.time, "monotonic", lambda: float(next(clock)))
    mark = controller._navigation_mark()

    controller._await_new_session(mark, wait_s=8)

    assert 6 <= len(sleeps) <= 8  # one look a "second" here, eight seconds long


def test_a_session_that_cannot_answer_yet_is_found_at_the_next_look(setup):
    controller, connection, _session, _sleeps = setup
    controller.get_screen_info()
    connection.sessions.append(OtherSession(1))
    describe = controller._describe_new_sessions
    controller._describe_new_sessions = MagicMock(side_effect=RuntimeError("not yet available"))

    assert "new_sessions" not in controller.get_screen_info()

    controller._describe_new_sessions = describe
    assert controller.get_screen_info()["new_sessions"][0]["session_index"] == 1


def test_without_a_connection_there_is_nothing_to_wait_for(setup):
    controller, _connection, _session, sleeps = setup
    controller._connection = None

    assert controller._navigation_mark() is None
    controller._await_new_session(None)
    assert sleeps == []


# -- the actions --------------------------------------------------------------

def test_a_button_that_opens_a_session_reports_it_in_its_response(setup):
    controller, connection, session, sleeps = setup
    button = MagicMock()
    button.press.side_effect = lambda: connection.sessions.append(OtherSession(1))
    session.findById.return_value = button

    result = controller.press_button("wnd[0]/usr/btnDISPLAY")

    assert result["status"] == "pressed"
    assert result["screen"]["new_sessions"][0]["session_id"] == "/app/con[0]/ses[1]"
    assert sleeps == []  # already there: no waiting
    assert controller._session is session


def spy_marks(controller):
    calls = []
    controller._ACTION_NEW_SESSION_WAIT_S = 0  # these tests are not about the wait
    original = controller._navigation_mark

    def spy():
        calls.append(True)
        return original()
    controller._navigation_mark = spy
    return calls


def test_f2_is_watched_and_enter_is_not(setup):
    controller, _connection, session, _sleeps = setup
    controller._vkey_refusal = lambda target, window, vkey: None
    controller._find_window = lambda window: session
    session.Info.Transaction = "VA05"
    calls = spy_marks(controller)

    controller.send_vkey(VKey.ENTER)
    assert calls == []

    controller.send_vkey(VKey.F2)
    assert calls == [True]


def test_the_grids_own_functions_are_not_watched(setup):
    controller, _connection, _session, _sleeps = setup
    grid = MagicMock()
    controller._find_element = lambda element_id: grid
    controller._alv_toolbar_button_type = lambda g, button_id: "Button"
    calls = spy_marks(controller)

    controller.press_alv_toolbar_button("wnd[0]/usr/cntlG/shellcont/shell", "&SORT_ASC")
    assert calls == []

    controller.press_alv_toolbar_button("wnd[0]/usr/cntlG/shellcont/shell", "ZDISPLAY")
    assert calls == [True]


def test_a_session_bound_by_o_is_not_reported_as_new(setup):
    controller, connection, _session, _sleeps = setup
    controller.get_screen_info()
    before = controller._session_ids(connection)
    opened = OtherSession(1)
    connection.sessions.append(opened)

    assert controller._rebind_after_new_session(before) is True

    assert controller._session is opened
    assert controller._take_new_sessions() == []
