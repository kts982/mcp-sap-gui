"""_com_with_retry: which COM HRESULTs are retried, and which mean SAP GUI is gone."""

import pytest

from mcp_sap_gui import server
from mcp_sap_gui.models import SAPGUINotConnectedError

CALL_REJECTED = -2147418111  # 0x80010001
SERVERCALL_RETRYLATER = -2147417846  # 0x8001010A
SERVERFAULT = -2147417851  # 0x80010105
DISCONNECTED = -2147417848  # 0x80010108
SERVER_UNAVAILABLE = -2147023174  # 0x800706BA


class FakeComError(Exception):
    """Shaped like pywintypes.com_error: args (hresult, strerror, excepinfo,
    argerror) plus an ``hresult`` attribute."""

    def __init__(self, hresult, strerror="", with_attr=True):
        super().__init__(hresult, strerror, None, None)
        if with_attr:
            self.hresult = hresult


class Flaky:
    """Raises the given errors in turn, then returns 'ok'."""

    def __init__(self, *errors):
        self.errors = list(errors)
        self.calls = 0

    def __call__(self):
        self.calls += 1
        if self.errors:
            raise self.errors.pop(0)
        return "ok"


@pytest.fixture
def sleeps(monkeypatch):
    calls = []
    monkeypatch.setattr(server.time, "sleep", calls.append)
    return calls


def _u32(value):
    return value & 0xFFFFFFFF


def test_constants_decode_to_the_named_codes():
    assert {_u32(h) for h in server._TRANSIENT_COM_HRESULTS} == {0x80010001, 0x8001010A}
    assert {_u32(h) for h in server._DEAD_SESSION_HRESULTS} == {0x80010108, 0x800706BA}
    assert not server._TRANSIENT_COM_HRESULTS & server._DEAD_SESSION_HRESULTS


@pytest.mark.parametrize("hresult", [CALL_REJECTED, SERVERCALL_RETRYLATER])
def test_busy_callee_is_retried(hresult, sleeps):
    fn = Flaky(FakeComError(hresult), FakeComError(hresult))
    assert server._com_with_retry(fn) == "ok"
    assert fn.calls == 3
    assert sleeps == [0.3, 0.6]


def test_hresult_in_args_only_is_recognised(sleeps):
    fn = Flaky(FakeComError(CALL_REJECTED, with_attr=False))
    assert server._com_with_retry(fn) == "ok"
    assert fn.calls == 2


def test_retries_are_bounded(sleeps):
    errors = [FakeComError(CALL_REJECTED) for _ in range(server._COM_MAX_RETRIES + 1)]
    fn = Flaky(*errors)
    with pytest.raises(FakeComError) as info:
        server._com_with_retry(fn)
    assert info.value is errors[-1]
    assert fn.calls == server._COM_MAX_RETRIES + 1
    assert len(sleeps) == server._COM_MAX_RETRIES


def test_server_fault_is_not_retried(sleeps):
    err = FakeComError(SERVERFAULT, "The server threw an exception.")
    fn = Flaky(err)
    with pytest.raises(FakeComError) as info:
        server._com_with_retry(fn)
    assert info.value is err
    assert fn.calls == 1
    assert sleeps == []


@pytest.mark.parametrize(
    ("hresult", "text", "name", "hex_code"),
    [
        (DISCONNECTED, "The object invoked has disconnected from its clients.",
         "RPC_E_DISCONNECTED", "0x80010108"),
        (SERVER_UNAVAILABLE, "The RPC server is unavailable.",
         "RPC_S_SERVER_UNAVAILABLE", "0x800706BA"),
    ],
)
@pytest.mark.parametrize("with_attr", [True, False])
def test_dead_sap_gui_surfaces_at_once_as_not_connected(
        hresult, text, name, hex_code, with_attr, sleeps):
    err = FakeComError(hresult, text, with_attr=with_attr)
    fn = Flaky(err)
    with pytest.raises(SAPGUINotConnectedError) as info:
        server._com_with_retry(fn)
    assert fn.calls == 1
    assert sleeps == []
    assert info.value.__cause__ is err
    message = str(info.value)
    assert name in message and hex_code in message
    assert "sap_connect_existing" in message


def test_non_com_errors_pass_through(sleeps):
    err = ValueError("Write operations disabled in read-only mode")
    fn = Flaky(err)
    with pytest.raises(ValueError) as info:
        server._com_with_retry(fn)
    assert info.value is err
    assert fn.calls == 1
    assert sleeps == []


def test_success_is_returned_unchanged(sleeps):
    assert server._com_with_retry(lambda: {"a": 1}) == {"a": 1}
    assert sleeps == []
