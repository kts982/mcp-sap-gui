"""A SAP session that opened beside the bound one is reported, and the binding stays.

Seen live (SAP GUI 8.10): System > New GUI Window through the menu on a
local system, and a document hotspot in /SCWM/MON on a remote one. The
call returns with the screen unchanged, and the new session shows on the
connection later: 0.5 to 0.7 s on the local system, 3 to 5 s on the remote
one. Only /o was noticed before; every other action looked as if nothing
had happened.

Nothing waits for the session: the first try did, for 1.5 s when the
screen "did not change", and a list that was only sorted in place then
took 1.9 s and was told that its screen had not changed.
"""

from unittest.mock import MagicMock

import pytest

from mcp_sap_gui import controller as controller_module
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
    """A controller bound to ses[0] of a connection; sleeping is recorded."""
    session = bound_session()
    connection = FakeConnection([session])
    controller = SAPGUIController()
    controller._session = session
    controller._connection = connection
    controller._application = object()
    controller._require_session = lambda: None
    controller._object_tree = lambda root, props, **kw: None
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

    # What it is, and no index to attach with: the report does not invite.
    assert screen["new_sessions"] == [{
        "session_id": "/app/con[0]/ses[1]",
        "transaction": "VA03",
        "title": "Display Sales Order",
    }]
    note = screen["new_sessions_note"]
    assert "still in the session you were in" in note
    assert "or by the user" in note and "only if they ask" in note
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


def test_a_session_closed_before_the_report_is_left_out(setup):
    controller, connection, _session, _sleeps = setup
    controller.get_screen_info()
    connection.sessions.append(OtherSession(1))
    controller._check_sessions()  # noticed, not yet reported
    connection.sessions.pop()

    assert "new_sessions" not in controller.get_screen_info()


def test_a_session_that_cannot_answer_yet_is_found_at_the_next_look(setup):
    # Seen once on a session that was still starting: reading its Id raised
    # "The data necessary to complete this operation is not yet available."
    controller, connection, _session, _sleeps = setup
    controller.get_screen_info()
    connection.sessions.append(OtherSession(1))
    describe = controller._describe_new_sessions
    controller._describe_new_sessions = MagicMock(side_effect=RuntimeError("not yet available"))

    assert "new_sessions" not in controller.get_screen_info()

    controller._describe_new_sessions = describe
    assert controller.get_screen_info()["new_sessions"][0]["session_id"].endswith("ses[1]")


def test_a_button_whose_session_is_there_already_reports_it_in_its_response(setup):
    controller, connection, session, sleeps = setup
    controller.get_screen_info()
    button = MagicMock()
    button.press.side_effect = lambda: connection.sessions.append(OtherSession(1))
    session.findById.return_value = button

    result = controller.press_button("wnd[0]/usr/btnDISPLAY")

    assert result["status"] == "pressed"
    assert result["screen"]["new_sessions"][0]["session_id"] == "/app/con[0]/ses[1]"
    assert controller._session is session
    assert sleeps == []


def test_an_action_does_not_wait_for_a_session(setup):
    # A sort, a page down or an RF step changes only what is inside the
    # screen: the response must come at once and claim nothing about it.
    controller, _connection, session, sleeps = setup
    controller.get_screen_info()
    session.findById.return_value = MagicMock()

    screen = controller.press_button("wnd[0]/tbar[1]/btn[28]")["screen"]

    assert sleeps == []
    assert "new_sessions" not in screen
    assert "unchanged_note" not in screen


def test_a_session_that_shows_later_is_in_the_next_response(setup):
    controller, connection, session, _sleeps = setup
    controller.get_screen_info()
    session.findById.return_value = MagicMock()
    assert "new_sessions" not in controller.press_button("wnd[0]/usr/btnDISPLAY")["screen"]

    connection.sessions.append(OtherSession(1))  # seconds later, on a remote system

    assert controller.get_screen_info()["new_sessions"][0]["transaction"] == "VA03"


def test_a_session_bound_by_o_is_not_reported_as_new(setup):
    controller, connection, _session, _sleeps = setup
    controller.get_screen_info()
    before = controller._session_ids(connection)
    opened = OtherSession(1)
    connection.sessions.append(opened)

    assert controller._rebind_after_new_session(before) is True

    assert controller._session is opened
    assert controller._take_new_sessions() == []
