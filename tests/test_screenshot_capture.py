"""Screenshot capture by host: PrintWindow in Business Client, HardCopy on plain SAP GUI."""

import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import pytest

from mcp_sap_gui import capture
from mcp_sap_gui.models import SAPGUIError
from mcp_sap_gui.sap_controller import SAPGUIController


def _controller(handle=None, iconic=False):
    controller = SAPGUIController()
    controller._session = MagicMock(Busy=False)
    controller._session.ActiveWindow = MagicMock(Id="wnd[0]")
    window = MagicMock(Iconic=iconic)
    if handle is not None:
        window.Handle = handle
    window.HardCopy.side_effect = lambda path, _fmt: Path(path).write_bytes(b"png")
    controller._session.findById.return_value = window
    return controller, window


def _embedded(value=True):
    return patch("mcp_sap_gui.discovery.is_embedded", return_value=value)


class TestCaptureByHost:
    def test_business_client_uses_printwindow(self, tmp_path):
        controller, window = _controller(handle=4242)

        def fake(hwnd, path):
            Path(path).write_bytes(b"png")
            return (0, 100)

        with _embedded(), patch("mcp_sap_gui.discovery.print_window_png", side_effect=fake) as pw:
            result = controller.save_screenshot(str(tmp_path / "e.png"), inline=False)

        assert pw.call_args[0][0] == 4242
        assert result["capture"] == "printwindow"
        window.HardCopy.assert_not_called()

    def test_business_client_refusal_does_not_fall_back_to_hardcopy(self, tmp_path):
        controller, window = _controller(handle=4242)
        target = tmp_path / "e.png"

        with _embedded(), patch(
            "mcp_sap_gui.discovery.print_window_png",
            side_effect=SAPGUIError("This session is not the visible one"),
        ):
            result = controller.save_screenshot(str(target))

        assert "not the visible one" in result["error"]
        assert not target.exists()
        window.HardCopy.assert_not_called()

    def test_plain_sap_gui_keeps_hardcopy(self, tmp_path):
        controller, window = _controller(handle=4242)

        with _embedded(False), patch("mcp_sap_gui.discovery.print_window_png") as pw:
            result = controller.save_screenshot(str(tmp_path / "e.png"), inline=False)

        assert result["capture"] == "hardcopy"
        window.HardCopy.assert_called_once()
        pw.assert_not_called()

    def test_plain_sap_gui_refuses_a_minimized_window(self, tmp_path):
        controller, window = _controller(handle=4242, iconic=True)

        with _embedded(False):
            result = controller.take_screenshot()

        assert "minimized" in result["error"]
        window.HardCopy.assert_not_called()

    @staticmethod
    def _popup_in_business_client(popup_handle):
        """A session whose main window (4242) is embedded and whose popup is a dialog of its own."""
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        controller._session.ActiveWindow = MagicMock(Id="wnd[1]")
        main = MagicMock(Iconic=False, Handle=4242)
        popup = MagicMock(Iconic=False)  # SAP's Iconic stays False in Business Client
        if popup_handle is not None:
            popup.Handle = popup_handle
        popup.HardCopy.side_effect = lambda path, _fmt: Path(path).write_bytes(b"png")
        controller._session.findById.side_effect = {"wnd[0]": main, "wnd[1]": popup}.__getitem__
        return controller, popup

    def test_a_popup_in_business_client_uses_printwindow(self, tmp_path):
        controller, popup = self._popup_in_business_client(5000)

        def fake(hwnd, path):
            Path(path).write_bytes(b"png")
            return (21, 324)

        with (
            patch("mcp_sap_gui.discovery.is_embedded", side_effect=lambda hwnd: hwnd == 4242),
            patch("mcp_sap_gui.discovery.print_window_png", side_effect=fake) as pw,
        ):
            result = controller.save_screenshot(str(tmp_path / "e.png"), inline=False)

        assert pw.call_args[0][0] == 5000
        assert (result["window"], result["capture"]) == ("wnd[1]", "printwindow")
        popup.HardCopy.assert_not_called()

    def test_a_popup_in_business_client_without_a_handle_is_refused(self, tmp_path):
        controller, popup = self._popup_in_business_client(None)
        target = tmp_path / "e.png"

        with patch("mcp_sap_gui.discovery.is_embedded", side_effect=lambda hwnd: hwnd == 4242):
            result = controller.save_screenshot(str(target))

        assert "no handle" in result["error"]
        assert not target.exists()
        popup.HardCopy.assert_not_called()

    def test_without_a_handle_it_is_plain_hardcopy(self, tmp_path):
        controller, window = _controller()  # a MagicMock handle is not an int

        result = controller.save_screenshot(str(tmp_path / "e.png"), inline=False)

        assert result["capture"] == "hardcopy"


def _fake_win32(visible=True, root_iconic=False, classes=None, owner_iconic=False,
                rect=None):
    """win32gui stand-in: handle 1 is a Business Client tab, 2 its parent, 99 the root.

    77 is the window that owns the root: the Business Client main window, when
    the root is a popup's dialog.
    """
    classes = classes or {1: "SAP_FRONTEND_EMBEDDED", 2: "Shell Embedding"}
    parents = {1: 2, 2: 0}

    def measure(_hwnd):  # reached only when the checks pass
        if rect is None:
            raise AssertionError("must refuse before measuring")
        return rect

    shown = iter(visible) if isinstance(visible, list) else None

    return SimpleNamespace(
        GetClassName=classes.__getitem__,
        GetParent=parents.__getitem__,
        GetWindow=lambda _hwnd, _relation: 77,
        IsIconic=lambda hwnd: (root_iconic and hwnd == 99) or (owner_iconic and hwnd == 77),
        GetAncestor=lambda _hwnd, _flag: 99,
        IsWindowVisible=lambda _hwnd: next(shown) if shown else visible,
        GetWindowRect=measure,
        GetWindowDC=lambda _hwnd: 5,
        ReleaseDC=lambda _hwnd, _hdc: 1,
        DeleteObject=lambda _handle: 1,
    )


def _print(fake):
    modules = {"win32gui": fake, "win32ui": MagicMock(), "PIL": MagicMock()}
    user32 = MagicMock()
    user32.SetThreadDpiAwarenessContext.return_value = 0
    windll = patch.object(capture.ctypes, "windll", create=True)
    with patch.dict(sys.modules, modules), windll as dll:
        dll.user32 = user32
        capture.print_window_png(1, "unused.png")


class TestPrintWindowChecks:
    def test_hidden_business_client_tab_is_refused(self):
        with pytest.raises(SAPGUIError, match="not the visible one"):
            _print(_fake_win32(visible=False))

    def test_minimized_business_client_is_refused_by_its_root_window(self):
        with pytest.raises(SAPGUIError, match="minimized"):
            _print(_fake_win32(root_iconic=True))

    def test_a_popup_of_a_minimized_business_client_is_refused_by_its_owner(self):
        # The popup's dialog is its own root and not minimized; its owner is.
        with pytest.raises(SAPGUIError, match="minimized"):
            _print(_fake_win32(visible=False, owner_iconic=True))

    def test_a_tab_switch_during_the_render_is_refused_and_nothing_is_saved(self):
        # Visible at the first check, hidden once PrintWindow has returned.
        fake = _fake_win32(visible=[True, False], rect=(0, 0, 10, 10))
        with pytest.raises(SAPGUIError, match="not the visible one"):
            _print(fake)

    def test_missing_pillow_names_the_extra(self):
        modules = {"win32gui": MagicMock(), "win32ui": MagicMock(), "PIL": None}
        with patch.dict(sys.modules, modules):
            with pytest.raises(SAPGUIError, match=r"mcp-sap-gui\[screenshots\]"):
                capture.print_window_png(1, "unused.png")


class TestIsEmbedded:
    def test_the_class_on_the_handle_or_a_parent_marks_business_client(self):
        fake = _fake_win32(classes={1: "SAP_FRONTEND_SESSION", 2: "SAP_FRONTEND_EMBEDDED"})
        with patch.dict(sys.modules, {"win32gui": fake}):
            assert capture.is_embedded(1)

    def test_plain_sap_gui_is_not_embedded(self):
        fake = _fake_win32(classes={1: "SAP_FRONTEND_SESSION", 2: "SAP_FRONTEND_SESSION"})
        with patch.dict(sys.modules, {"win32gui": fake}):
            assert not capture.is_embedded(1)
