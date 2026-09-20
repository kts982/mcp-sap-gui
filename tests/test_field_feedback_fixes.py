"""Regression tests for the first real-use feedback round.

Sustained agent work on a live system surfaced defects that the probe-style
tests had missed. Each class pins one of them:

- docking containers (``wnd[0]/shellcont``) were rejected by the element-ID
  validator, so view clusters (SM34, IMG dialog structures) and the SE80 tree
  could not be navigated at all
- ``message_id`` came back padded to 20 characters
- a bare ``/n`` (leave the current transaction) was rejected as empty
- ``sap_screenshot`` could not save to a file, so screenshots never reached a
  document
"""

import base64
import json
import struct
from unittest.mock import MagicMock, PropertyMock, patch

import pytest

import mcp_sap_gui.server as srv_mod
from mcp_sap_gui.session_manager import SessionManager

# ===========================================================================
# Helpers
# ===========================================================================

def _make_mock_ctx():
    return MagicMock()


def _make_controller_with_session():
    from mcp_sap_gui.sap_controller import SAPGUIController
    controller = SAPGUIController()
    controller._session = MagicMock(Busy=False)
    return controller


def _png_bytes(width: int, height: int) -> bytes:
    """Smallest byte string that carries a valid PNG signature + IHDR size."""
    return (
        b"\x89PNG\r\n\x1a\n"
        + struct.pack(">I", 13) + b"IHDR"
        + struct.pack(">II", width, height)
        + b"\x08\x02\x00\x00\x00"
    )


@pytest.fixture
def srv():
    """Fresh server module with a clean SessionManager + default config."""
    srv_mod._session_mgr = SessionManager()
    srv_mod.config = srv_mod.ServerConfig()
    yield srv_mod
    srv_mod._session_mgr.shutdown()


# ===========================================================================
# Docking containers
# ===========================================================================

class TestDockingContainerIds:
    """wnd[n]/shellcont... is a valid element path; a bare window is not."""

    @pytest.mark.parametrize("element_id", [
        "wnd[0]/shellcont/shell",
        "wnd[0]/shellcont[1]/shell",
        "wnd[0]/shellcont/shell/shellcont[0]/shell",
        "wnd[1]/shellcont/shell",
        "wnd[0]/titl",
    ])
    def test_docking_container_paths_are_valid(self, element_id):
        controller = _make_controller_with_session()
        assert controller._validate_element_id(element_id) == element_id

    def test_full_session_path_normalizes_to_short_form(self):
        controller = _make_controller_with_session()
        assert controller._validate_element_id(
            "/app/con[0]/ses[0]/wnd[0]/shellcont/shell"
        ) == "wnd[0]/shellcont/shell"

    @pytest.mark.parametrize("element_id", [
        "wnd[0]",                  # a window is not an element
        "wnd[0]/shell",            # unknown top-level area
        "wnd[0]/shellcontX/shell",
        "wnd[0]/shellcont[a]/shell",
        "wnd[0]/bad",
    ])
    def test_other_top_level_paths_stay_rejected(self, element_id):
        controller = _make_controller_with_session()
        with pytest.raises(ValueError, match="Invalid SAP element ID"):
            controller._validate_element_id(element_id)

    def test_error_text_mentions_docking_containers(self):
        controller = _make_controller_with_session()
        with pytest.raises(ValueError, match=r"wnd\[0\]/shellcont"):
            controller._validate_element_id("wnd[0]/bad")

    def test_read_tree_reaches_a_docked_tree(self):
        """The view-cluster tree must get past validation to findById."""
        controller = _make_controller_with_session()

        controller.read_tree("wnd[0]/shellcont/shell", max_nodes=5)

        controller._session.findById.assert_called_with("wnd[0]/shellcont/shell")


class TestBareWindowDiscovery:
    """Only discovery accepts a bare window, to list its top-level areas."""

    def test_container_id_accepts_a_bare_window(self):
        controller = _make_controller_with_session()
        assert controller._validate_container_id("wnd[0]") == "wnd[0]"
        assert controller._validate_container_id(
            "/app/con[0]/ses[0]/wnd[1]"
        ) == "wnd[1]"

    def test_container_id_still_rejects_malformed_paths(self):
        controller = _make_controller_with_session()
        with pytest.raises(ValueError, match="Invalid SAP element ID"):
            controller._validate_container_id("wnd[0]/bad")

    def test_get_screen_elements_lists_window_children(self):
        controller = _make_controller_with_session()
        dock = MagicMock(
            Id="/app/con[0]/ses[0]/wnd[0]/shellcont", Type="GuiContainerShell",
            Text="", Changeable=False, Visible=True,
        )
        dock.Name = "shellcont"
        dock.Children.Count = 0
        window = MagicMock()
        window.Children.Count = 1
        window.Children.side_effect = lambda i: dock
        controller._session.findById.return_value = window

        elements = controller.get_screen_elements("wnd[0]", max_depth=1)

        controller._session.findById.assert_called_once_with("wnd[0]")
        assert [e.type for e in elements] == ["GuiContainerShell"]

    def test_a_child_with_a_failing_property_does_not_hide_its_siblings(self):
        """Seen live: GuiTitlebar.Changeable raises IndexError, which used to
        abort the loop and drop every later child (usr, sbar, shellcont)."""
        controller = _make_controller_with_session()

        titlebar = MagicMock(
            Id="/app/con[0]/ses[0]/wnd[0]/titl", Type="GuiTitlebar",
            Text="SAP Easy Access", Visible=True,
        )
        titlebar.Name = "titl"
        titlebar.Children.Count = 0
        type(titlebar).Changeable = PropertyMock(
            side_effect=IndexError("tuple index out of range"),
        )
        usr = MagicMock(
            Id="/app/con[0]/ses[0]/wnd[0]/usr", Type="GuiUserArea",
            Text="", Changeable=False, Visible=True,
        )
        usr.Name = "usr"
        usr.Children.Count = 0
        children = [titlebar, usr]
        window = MagicMock()
        window.Children.Count = len(children)
        window.Children.side_effect = lambda i: children[i]
        controller._session.findById.return_value = window

        elements = controller.get_screen_elements("wnd[0]", max_depth=1)

        assert [e.type for e in elements] == ["GuiTitlebar", "GuiUserArea"]
        assert elements[0].changeable is False

    def test_other_tools_keep_rejecting_a_bare_window(self):
        controller = _make_controller_with_session()

        result = controller.read_field("wnd[0]")

        assert "Invalid SAP element ID" in result["error"]
        controller._session.findById.assert_not_called()


class TestGetDockingContainers:
    def _window_with_children(self, controller, child_ids):
        children = [MagicMock(Id=child_id) for child_id in child_ids]
        window = MagicMock()
        window.Children.Count = len(children)
        window.Children.side_effect = lambda i: children[i]
        controller._session.findById.return_value = window

    def test_returns_short_ids_of_docking_containers_only(self):
        controller = _make_controller_with_session()
        self._window_with_children(controller, [
            "/app/con[0]/ses[0]/wnd[0]/titl",
            "/app/con[0]/ses[0]/wnd[0]/mbar",
            "/app/con[0]/ses[0]/wnd[0]/tbar[0]",
            "/app/con[0]/ses[0]/wnd[0]/usr",
            "/app/con[0]/ses[0]/wnd[0]/sbar",
            "/app/con[0]/ses[0]/wnd[0]/shellcont",
            "/app/con[0]/ses[0]/wnd[0]/shellcont[1]",
        ])

        assert controller.get_docking_containers("wnd[0]") == [
            "wnd[0]/shellcont", "wnd[0]/shellcont[1]",
        ]

    def test_plain_screen_has_none(self):
        controller = _make_controller_with_session()
        self._window_with_children(controller, [
            "/app/con[0]/ses[0]/wnd[0]/usr",
            "/app/con[0]/ses[0]/wnd[0]/sbar",
        ])

        assert controller.get_docking_containers("wnd[0]") == []

    def test_com_failure_degrades_to_empty(self):
        """A hint must never turn a working discovery call into an error."""
        controller = _make_controller_with_session()
        controller._session.findById.side_effect = Exception("COM gone")

        assert controller.get_docking_containers("wnd[0]") == []


class TestScreenElementsDockingHint:
    async def test_user_area_discovery_reports_docking_containers(self, srv):
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        mock_ctrl.get_screen_elements.return_value = []
        mock_ctrl.get_docking_containers.return_value = ["wnd[0]/shellcont"]
        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            result = await srv.sap_get_screen_elements(ctx)

        mock_ctrl.get_docking_containers.assert_called_once_with("wnd[0]")
        assert result["docking_containers"] == ["wnd[0]/shellcont"]

    async def test_full_session_path_resolves_the_window(self, srv):
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        mock_ctrl.get_screen_elements.return_value = []
        mock_ctrl.get_docking_containers.return_value = []
        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            await srv.sap_get_screen_elements(ctx, "/app/con[0]/ses[0]/wnd[1]/usr")

        mock_ctrl.get_docking_containers.assert_called_once_with("wnd[1]")

    async def test_key_is_omitted_on_plain_screens(self, srv):
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        mock_ctrl.get_screen_elements.return_value = []
        mock_ctrl.get_docking_containers.return_value = []
        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            result = await srv.sap_get_screen_elements(ctx)

        assert "docking_containers" not in result

    async def test_not_queried_below_the_user_area(self, srv):
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        mock_ctrl.get_screen_elements.return_value = []
        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            result = await srv.sap_get_screen_elements(ctx, "wnd[0]/usr/subSUB")

        mock_ctrl.get_docking_containers.assert_not_called()
        assert "docking_containers" not in result


# ===========================================================================
# Status bar
# ===========================================================================

class TestStatusBarTrimsPadding:
    def test_message_id_is_trimmed(self):
        controller = _make_controller_with_session()
        sbar = MagicMock()
        sbar.Text = "Data was saved"
        sbar.MessageType = "S"
        sbar.MessageId = "00                  "
        sbar.MessageNumber = "001"
        window = MagicMock(Text="Change View")
        controller._session.findById.side_effect = (
            lambda id: sbar if id == "wnd[0]/sbar" else window
        )

        info = controller.get_screen_info()

        assert info["message"] == "Data was saved"
        assert info["message_id"] == "00"
        assert info["message_type"] == "S"
        assert info["message_number"] == "001"


# ===========================================================================
# Bare /n
# ===========================================================================

class TestLeaveTransactionCommand:
    @pytest.mark.parametrize("tcode", ["/n", "/N", "  /n  "])
    def test_policy_lets_a_bare_n_through(self, srv, tcode):
        assert srv._enforce_transaction_policy(tcode) == "/N"

    def test_allowlist_does_not_block_leaving(self, srv):
        srv.config = srv.ServerConfig(allowed_transactions=["MM03"])
        assert srv._enforce_transaction_policy("/n") == "/N"

    @pytest.mark.parametrize("tcode", ["/nSU01", "/n su01", "/n/nSU01"])
    def test_a_prefixed_blocked_transaction_stays_blocked(self, srv, tcode):
        with pytest.raises(ValueError, match="blocked by security policy"):
            srv._enforce_transaction_policy(tcode)

    @pytest.mark.parametrize("tcode", ["", "   ", "/o", "/*"])
    def test_other_empty_commands_stay_rejected(self, srv, tcode):
        with pytest.raises(ValueError, match="cannot be empty"):
            srv._enforce_transaction_policy(tcode)

    async def test_tool_forwards_the_command(self, srv):
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        mock_ctrl.execute_transaction.return_value = {"transaction": ""}
        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            await srv.sap_execute_transaction("/n", ctx)

        mock_ctrl.execute_transaction.assert_called_once_with("/n")

    def test_controller_sends_it_through_the_command_field(self):
        controller = _make_controller_with_session()
        mock_okcd = MagicMock()
        mock_window = MagicMock()
        controller._session.findById.side_effect = (
            lambda id: mock_okcd if "okcd" in id else mock_window
        )
        controller.get_screen_info = MagicMock(return_value={"transaction": ""})

        result = controller.execute_transaction(" /n ")

        controller._session.StartTransaction.assert_not_called()
        assert mock_okcd.text == "/n"
        mock_window.sendVKey.assert_called_once()
        assert "error" not in result


# ===========================================================================
# Screenshot to file
# ===========================================================================

class TestSaveScreenshot:
    def _controller_writing(self, payload: bytes):
        controller = _make_controller_with_session()
        controller._session.ActiveWindow = MagicMock(Id="wnd[0]")
        window = MagicMock()
        window.HardCopy.side_effect = (
            lambda path, _fmt: open(path, "wb").write(payload)
        )
        controller._session.findById.return_value = window
        return controller, window

    def test_saves_full_resolution_file_and_reports_size(self, tmp_path):
        payload = _png_bytes(3176, 1768)
        controller, window = self._controller_writing(payload)
        controller._optimize_screenshot = MagicMock()
        target = tmp_path / "selection_screen.png"

        result = controller.save_screenshot(str(target))

        # The numeric constant: the string "PNG" makes SAP GUI write a BMP.
        window.HardCopy.assert_called_once_with(str(target), 2)
        assert result["filepath"] == str(target)
        assert (result["width"], result["height"]) == (3176, 1768)
        assert result["size_bytes"] == len(payload)
        assert result["window"] == "wnd[0]"
        # The saved file is never the one that gets downscaled.
        assert target.read_bytes() == payload
        optimized_path = controller._optimize_screenshot.call_args[0][0]
        assert optimized_path != str(target)
        assert result["data"] == base64.b64encode(payload).decode()

    def test_inline_false_skips_the_image(self, tmp_path):
        controller, _ = self._controller_writing(_png_bytes(800, 600))
        controller._optimize_screenshot = MagicMock()

        result = controller.save_screenshot(str(tmp_path / "a.png"), inline=False)

        assert "data" not in result
        controller._optimize_screenshot.assert_not_called()

    def test_never_overwrites_an_existing_file(self, tmp_path):
        controller, window = self._controller_writing(_png_bytes(1, 1))
        target = tmp_path / "keep.png"
        target.write_bytes(b"original")

        result = controller.save_screenshot(str(target))

        assert "already exists" in result["error"]
        assert target.read_bytes() == b"original"
        window.HardCopy.assert_not_called()

    @pytest.mark.parametrize("name", ["shot.jpg", "shot", "notes.txt", "shot.png.exe"])
    def test_rejects_non_png_targets(self, tmp_path, name):
        controller, window = self._controller_writing(_png_bytes(1, 1))

        result = controller.save_screenshot(str(tmp_path / name))

        assert ".png" in result["error"]
        window.HardCopy.assert_not_called()

    def test_rejects_a_missing_directory(self, tmp_path):
        controller, window = self._controller_writing(_png_bytes(1, 1))

        result = controller.save_screenshot(str(tmp_path / "nope" / "shot.png"))

        assert "Directory does not exist" in result["error"]
        window.HardCopy.assert_not_called()

    def test_reports_when_sap_gui_wrote_nothing(self, tmp_path):
        controller = _make_controller_with_session()
        controller._session.ActiveWindow = MagicMock(Id="wnd[0]")
        controller._session.findById.return_value = MagicMock()

        result = controller.save_screenshot(str(tmp_path / "shot.png"))

        assert "did not write" in result["error"]

    def test_temp_copy_is_removed(self, tmp_path):
        controller, _ = self._controller_writing(_png_bytes(10, 10))
        controller._optimize_screenshot = MagicMock()

        controller.save_screenshot(str(tmp_path / "shot.png"))

        import os
        temp_path = controller._optimize_screenshot.call_args[0][0]
        assert not os.path.exists(temp_path)


class TestScreenshotTool:
    async def test_without_save_path_returns_only_the_image(self, srv):
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        mock_ctrl.take_screenshot.return_value = {"data": "QUJD", "window": "wnd[0]"}
        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            result = await srv.sap_screenshot(ctx)

        mock_ctrl.save_screenshot.assert_not_called()
        assert [block.type for block in result.content] == ["image"]
        assert result.content[0].data == "QUJD"
        assert result.content[0].mimeType == "image/png"

    async def test_save_path_reports_the_file_and_keeps_the_image(self, srv):
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        mock_ctrl.save_screenshot.return_value = {
            "format": "png", "filepath": "C:\\docs\\shot.png", "window": "wnd[0]",
            "size_bytes": 10, "width": 3176, "height": 1768, "data": "QUJD",
        }
        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            result = await srv.sap_screenshot(ctx, save_path="shot.png")

        mock_ctrl.save_screenshot.assert_called_once_with("shot.png", inline=True)
        assert [block.type for block in result.content] == ["text", "image"]
        meta = json.loads(result.content[0].text)
        assert meta["filepath"] == "C:\\docs\\shot.png"
        assert (meta["width"], meta["height"]) == (3176, 1768)
        assert "data" not in meta  # the image travels once, as an image block

    async def test_inline_false_returns_text_only(self, srv):
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        mock_ctrl.save_screenshot.return_value = {
            "format": "png", "filepath": "C:\\docs\\shot.png", "window": "wnd[0]",
        }
        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            result = await srv.sap_screenshot(ctx, save_path="shot.png", inline=False)

        mock_ctrl.save_screenshot.assert_called_once_with("shot.png", inline=False)
        assert [block.type for block in result.content] == ["text"]

    async def test_save_errors_surface_to_the_agent(self, srv):
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        mock_ctrl.save_screenshot.return_value = {"error": "File already exists"}
        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            with pytest.raises(ValueError, match="already exists"):
                await srv.sap_screenshot(ctx, save_path="shot.png")
