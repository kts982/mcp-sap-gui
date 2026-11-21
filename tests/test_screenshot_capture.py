"""Screenshot capture by window handle: PrintWindow when possible, clear refusals."""

from pathlib import Path
from unittest.mock import MagicMock, patch

from mcp_sap_gui.models import SAPGUIError
from mcp_sap_gui.sap_controller import SAPGUIController


def _controller(handle=None):
    controller = SAPGUIController()
    controller._session = MagicMock(Busy=False)
    controller._session.ActiveWindow = MagicMock(Id="wnd[0]")
    window = MagicMock()
    if handle is not None:
        window.Handle = handle
    window.HardCopy.side_effect = lambda path, _fmt: Path(path).write_bytes(b"png")
    controller._session.findById.return_value = window
    return controller, window


class TestCaptureByHandle:
    def test_printwindow_is_used_when_the_handle_is_real(self, tmp_path):
        controller, window = _controller(handle=4242)
        target = tmp_path / "e.png"

        def fake(hwnd, path):
            Path(path).write_bytes(b"png")
            return (0, 100)

        with patch("mcp_sap_gui.discovery.print_window_png", side_effect=fake) as pw:
            result = controller.save_screenshot(str(target), inline=False)

        pw.assert_called_once_with(4242, str(target))
        assert result["capture"] == "printwindow"
        window.HardCopy.assert_not_called()

    def test_hidden_session_or_minimized_window_is_refused(self, tmp_path):
        controller, window = _controller(handle=4242)
        target = tmp_path / "e.png"

        with patch(
            "mcp_sap_gui.discovery.print_window_png",
            side_effect=SAPGUIError("This session is not the visible one"),
        ):
            result = controller.save_screenshot(str(target))

        assert "not the visible one" in result["error"]
        assert not target.exists()
        window.HardCopy.assert_not_called()

    def test_without_a_handle_it_falls_back_to_hardcopy(self, tmp_path):
        controller, window = _controller()  # a MagicMock handle is not an int

        result = controller.save_screenshot(str(tmp_path / "e.png"), inline=False)

        assert result["capture"] == "hardcopy"
        window.HardCopy.assert_called_once()

    def test_missing_libraries_fall_back_to_hardcopy(self, tmp_path):
        from mcp_sap_gui.capture import CaptureUnavailableError

        controller, window = _controller(handle=4242)

        with patch(
            "mcp_sap_gui.discovery.print_window_png",
            side_effect=CaptureUnavailableError("no pywin32"),
        ):
            result = controller.save_screenshot(str(tmp_path / "e.png"), inline=False)

        assert result["capture"] == "hardcopy"
        window.HardCopy.assert_called_once()

    def test_the_inline_screenshot_uses_the_same_capture(self):
        controller, window = _controller(handle=4242)

        with patch(
            "mcp_sap_gui.discovery.print_window_png",
            side_effect=SAPGUIError("The SAP window is minimized; restore it and retry."),
        ):
            result = controller.take_screenshot()

        assert "minimized" in result["error"]
