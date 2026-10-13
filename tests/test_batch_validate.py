"""set_batch_fields(validate=True): no Enter after a failed field, and a
refused Enter is not reported as a clean validation."""

from tests.fakes import controller_for

GOOD = "wnd[0]/usr/txtF-GOOD"
BAD = "wnd[0]/usr/txtF-BAD"

SCREEN = {"message": "Entry checked", "message_type": "S"}


class FakeField:
    Changeable = True
    Highlighted = False

    def __init__(self, fail: bool) -> None:
        self._fail = fail
        self._text = ""

    @property
    def text(self) -> str:
        return self._text

    @text.setter
    def text(self, value: str) -> None:
        if self._fail:
            raise RuntimeError("field is too long")
        self._text = value


def controller(enter_result=None):
    c = controller_for(None)
    fields = {GOOD: FakeField(False), BAD: FakeField(True)}
    c._find_element = lambda fid: fields[fid]
    c.enter_calls = 0

    def press_enter():
        c.enter_calls += 1
        return enter_result if enter_result is not None else {
            "vkey": 0, "screen": SCREEN,
        }

    c.press_enter = press_enter
    return c


def test_failed_field_skips_enter():
    c = controller()
    result = c.set_batch_fields({GOOD: "x", BAD: "y"}, validate=True)
    assert result["failed"] == 1
    assert c.enter_calls == 0
    assert result["validation"] == {
        "performed": False,
        "reason": "1 field(s) were not set; Enter not pressed",
    }


def test_validate_partial_presses_enter_anyway():
    c = controller()
    result = c.set_batch_fields(
        {GOOD: "x", BAD: "y"}, validate=True, validate_partial=True,
    )
    assert c.enter_calls == 1
    assert result["validation"]["performed"] is True
    assert result["validation"]["message_type"] == "S"


def test_all_fields_set_validates():
    c = controller()
    result = c.set_batch_fields({GOOD: "x"}, validate=True)
    assert c.enter_calls == 1
    assert result["validation"]["performed"] is True
    assert result["validation"]["message"] == "Entry checked"


def test_no_field_set_still_skips_enter():
    c = controller()
    result = c.set_batch_fields({BAD: "y"}, validate=True)
    assert c.enter_calls == 0
    assert result["validation"] == {
        "performed": False, "reason": "no fields were set",
    }


def test_refused_enter_is_not_performed():
    refused = {
        "vkey": 0, "window": "wnd[0]",
        "error": "Enter is not active on this screen",
        "available_keys": ["F3"],
        "screen": {"message": "old message", "message_type": "S"},
    }
    result = controller(refused).set_batch_fields({GOOD: "x"}, validate=True)
    validation = result["validation"]
    assert validation["performed"] is False
    assert validation["error"] == "Enter is not active on this screen"
    assert validation["available_keys"] == ["F3"]
    assert "message_type" not in validation


def test_com_failure_on_enter_is_not_performed():
    # send_vkey's _error_result shape: no screen at all.
    failure = {"vkey": 0, "window": "wnd[0]", "error": "Could not send SAP key"}
    result = controller(failure).set_batch_fields({GOOD: "x"}, validate=True)
    assert result["validation"] == {
        "performed": False, "error": "Could not send SAP key",
    }


def test_refusal_from_the_real_send_vkey():
    c = controller_for(None)
    field = FakeField(False)
    c._find_element = lambda fid: field

    class Window:
        def IsVKeyAllowed(self, vkey):  # noqa: N802 - COM name
            return vkey == 3

        def sendVKey(self, vkey):  # noqa: N802
            raise AssertionError("must not be sent")

    c._find_window = lambda window: Window()
    c.get_screen_info = lambda *a, **k: {"message": "", "message_type": ""}
    result = c.set_batch_fields({GOOD: "x"}, validate=True)
    assert result["validation"]["performed"] is False
    assert "Enter is not active" in result["validation"]["error"]
    assert result["validation"]["available_keys"] == ["F3"]
