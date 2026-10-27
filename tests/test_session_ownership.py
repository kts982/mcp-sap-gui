"""Session ownership: close only what this controller opened.

Fakes stand in for the SAP GUI scripting objects. A closed fake session
raises on ``Id``, like a GuiSession whose window was closed.
"""

from types import SimpleNamespace

import pytest

from mcp_sap_gui.controller import SAPGUIControllerBase
from mcp_sap_gui.session_manager import ManagedSession, SessionManager


class FakeSession:
    def __init__(self, session_id: str, connection: "FakeConnection") -> None:
        self._id = session_id
        self._connection = connection
        self.closed = False
        self.Busy = False
        self.Info = SimpleNamespace(
            SystemName="D01", SystemNumber=0, Client="100", User="U",
            Language="EN", Transaction="SESSION_MANAGER", Program="",
            ScreenNumber=0, SessionNumber=1, ScriptingModeReadOnly=False,
        )
        self.okcd = SimpleNamespace(text="")

    @property
    def Id(self) -> str:  # noqa: N802 - COM name
        if self.closed:
            raise RuntimeError("RPC_E_DISCONNECTED")
        return self._id

    def findById(self, element_id: str):  # noqa: N802 - COM name
        if element_id == "wnd[0]/tbar[0]/okcd":
            return self.okcd
        if element_id == "wnd[0]":
            return SimpleNamespace(sendVKey=self._send_vkey)
        raise AssertionError(element_id)

    def _send_vkey(self, _vkey: int) -> None:
        if self.okcd.text.upper().startswith("/O") and self._connection.opens_on_o:
            self._connection.open_session()


class FakeConnection:
    def __init__(self, index: int, slots: list[int]) -> None:
        self.index = index
        self.Id = f"/app/con[{index}]"
        self.sessions: list[FakeSession] = []
        self.closed_ids: list[str] = []
        self.opens_on_o = True
        for slot in slots:
            self.sessions.append(FakeSession(self._sid(slot), self))

    def _sid(self, slot: int) -> str:
        return f"/app/con[{self.index}]/ses[{slot}]"

    @property
    def Children(self) -> "_Collection":  # noqa: N802 - COM name
        return _Collection(self.sessions)

    def open_session(self) -> FakeSession:
        """A new session takes the lowest free ses[n], like SAP GUI."""
        used = {s._id for s in self.sessions}
        slot = 0
        while self._sid(slot) in used:
            slot += 1
        session = FakeSession(self._sid(slot), self)
        self.sessions.insert(slot, session)  # Children() in ses[n] order
        return session

    def close_by_user(self, session: FakeSession) -> None:
        session.closed = True
        self.sessions.remove(session)

    def CloseSession(self, session_id: str) -> None:  # noqa: N802 - COM name
        for session in self.sessions:
            if session._id == session_id:
                self.closed_ids.append(session_id)
                self.close_by_user(session)
                return
        raise RuntimeError(f"no session {session_id}")


class _Collection:
    def __init__(self, items: list) -> None:
        self._items = items

    @property
    def Count(self) -> int:  # noqa: N802 - COM name
        return len(self._items)

    def __call__(self, index: int):
        return self._items[index]


class FakeEngine:
    def __init__(self, connections: list[FakeConnection]) -> None:
        self.connections = connections
        self.ActiveSession = None
        self.to_open: FakeConnection | None = None

    @property
    def Children(self) -> "_Collection":  # noqa: N802 - COM name
        return _Collection(self.connections)

    def OpenConnection(self, _description: str, _sync: bool):  # noqa: N802
        self.connections.append(self.to_open)
        return self.to_open


@pytest.fixture
def make_controller(monkeypatch):
    monkeypatch.setattr(SAPGUIControllerBase, "_check_dependencies", lambda self: None)
    monkeypatch.setattr(SAPGUIControllerBase, "_NEW_SESSION_WAIT_S", 0.0)

    def make(engine: FakeEngine) -> SAPGUIControllerBase:
        controller = SAPGUIControllerBase()
        controller._rot_engines["SAPGUI"] = engine
        controller.get_screen_info = lambda include_popup=True: {}
        controller._discover_connections = lambda: [
            ("saplogon", "SAPGUI", engine, conn) for conn in engine.connections
        ]
        return controller

    return make


def _operator_connection() -> FakeConnection:
    return FakeConnection(0, [0, 1])  # the user's own logon, two windows


def _connect(make_controller) -> tuple:
    operator = _operator_connection()
    engine = FakeEngine([operator])
    engine.to_open = FakeConnection(1, [0])
    controller = make_controller(engine)
    controller.connect("D01")
    return controller, engine, operator, engine.to_open


def test_owned_session_is_closed(make_controller):
    controller, _engine, operator, owned = _connect(make_controller)
    assert controller._owns_session

    assert controller.disconnect() == ["/app/con[1]/ses[0]"]
    assert owned.closed_ids == ["/app/con[1]/ses[0]"]
    assert operator.closed_ids == []


def test_close_owned_false_leaves_owned_session_open(make_controller):
    controller, _engine, _operator, owned = _connect(make_controller)

    assert controller.disconnect(close_owned=False) == []
    assert owned.closed_ids == []
    assert not controller.is_connected


def test_attached_session_is_never_closed(make_controller):
    operator = _operator_connection()
    controller = make_controller(FakeEngine([operator]))
    controller.connect_to_existing_session(0, 1)
    assert not controller._owns_session

    assert controller.disconnect() == []
    assert operator.closed_ids == []


def test_attach_after_connect_forgets_ownership(make_controller):
    controller, _engine, operator, owned = _connect(make_controller)
    controller.connect_to_existing_session(0, 0)

    assert controller.disconnect() == []
    assert operator.closed_ids == [] and owned.closed_ids == []


def test_o_rebind_ignores_active_session_on_other_connection(make_controller):
    controller, engine, operator, owned = _connect(make_controller)
    # The user's focused window on another connection.
    engine.ActiveSession = operator.sessions[1]

    result = controller.execute_transaction("/oSE16")

    assert result["new_session_bound"] is True
    assert controller._session.Id == "/app/con[1]/ses[1]"
    assert controller._connection is owned
    # Both sessions this controller opened are closed, the user's are not.
    assert controller.disconnect() == ["/app/con[1]/ses[1]", "/app/con[1]/ses[0]"]
    assert operator.closed_ids == []


def test_o_rebind_takes_the_new_id_not_the_last_index(make_controller):
    controller, _engine, _operator, owned = _connect(make_controller)
    # A window the user opened on this connection by hand, at ses[2]
    # (ses[1] was free), sits last in Children.
    owned.sessions.append(FakeSession("/app/con[1]/ses[2]", owned))

    controller.execute_transaction("/oSE16")

    assert controller._session.Id == "/app/con[1]/ses[1]"
    controller.disconnect()
    assert sorted(owned.closed_ids) == ["/app/con[1]/ses[0]", "/app/con[1]/ses[1]"]
    assert [s.Id for s in owned.sessions] == ["/app/con[1]/ses[2]"]


def test_o_from_attached_session_is_not_owned(make_controller):
    operator = _operator_connection()
    controller = make_controller(FakeEngine([operator]))
    controller.connect_to_existing_session(0, 0)

    result = controller.execute_transaction("/oSE16")

    assert result["new_session_bound"] is True
    assert controller._session.Id == "/app/con[0]/ses[2]"
    assert not controller._owns_session
    assert controller.disconnect() == []
    assert operator.closed_ids == []


def test_o_without_new_session_keeps_binding(make_controller):
    controller, engine, operator, owned = _connect(make_controller)
    owned.opens_on_o = False  # e.g. the session limit was reached
    engine.ActiveSession = operator.sessions[0]

    result = controller.execute_transaction("/oSE16")

    assert result["new_session_bound"] is False
    assert controller._session.Id == "/app/con[1]/ses[0]"


def test_reused_id_of_closed_owned_session_is_not_closed(make_controller):
    controller, _engine, _operator, owned = _connect(make_controller)
    controller.execute_transaction("/oSE16")
    opened = controller._session
    owned.close_by_user(opened)  # the user closes the /o window...
    newcomer = owned.open_session()  # ...and opens one that reuses ses[1]
    assert newcomer.Id == "/app/con[1]/ses[1]"

    assert controller.disconnect() == ["/app/con[1]/ses[0]"]
    assert newcomer in owned.sessions


def test_not_a_new_session_command_has_no_rebind_key(make_controller):
    operator = _operator_connection()
    controller = make_controller(FakeEngine([operator]))
    controller.connect_to_existing_session(0, 0)
    controller._session.StartTransaction = lambda tcode: None

    assert "new_session_bound" not in controller.execute_transaction("SE16")


# ---------------------------------------------------------------------------
# SessionManager
# ---------------------------------------------------------------------------

class FakeController:
    def __init__(self, owned: int) -> None:
        self.is_connected = True
        self._owns_session = owned > 0
        self._owned_sessions = [object()] * owned
        self.calls: list[bool] = []

    def disconnect(self, close_owned: bool = True) -> list:
        self.calls.append(close_owned)
        return [f"ses[{i}]" for i in range(len(self._owned_sessions))] if close_owned else []


def _manager_with(*controllers: FakeController) -> SessionManager:
    manager = SessionManager()
    for key, controller in enumerate(controllers):
        manager._sessions[key] = ManagedSession(controller=controller)
    return manager


def test_release_reports_closed_sessions():
    manager = _manager_with(FakeController(owned=1))
    assert manager.release(0) == {
        "released": True, "was_connected": True, "action": "closed",
        "closed_sessions": ["ses[0]"],
    }


def test_release_without_close_reports_sessions_left_open():
    manager = _manager_with(FakeController(owned=2))
    assert manager.release(0, close_owned=False) == {
        "released": True, "was_connected": True, "action": "detached",
        "owned_sessions_left_open": 2,
    }


def test_release_attached_is_detached():
    manager = _manager_with(FakeController(owned=0))
    assert manager.release(0)["action"] == "detached"


def test_release_all_without_close_detaches_every_binding():
    first, second = FakeController(owned=1), FakeController(owned=1)
    manager = _manager_with(first, second)

    manager.release_all(close_owned=False)

    assert first.calls == [False] and second.calls == [False]


def test_release_all_default_closes():
    controller = FakeController(owned=1)
    manager = _manager_with(controller)
    manager.release_all()
    assert controller.calls == [True]
