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
- responses were heavy: absolute element IDs, one discovery element per table
  cell, every field echoed by a batch fill, a dead popup dumped after confirm
- a classic list (``WRITE`` output) could only be read from a screenshot
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


# ===========================================================================
# Response size
# ===========================================================================

def _child(element_id, type_, *, text="", changeable=False, children=()):
    child = MagicMock(Id=element_id, Type=type_, Text=text, Changeable=changeable)
    child.Name = element_id.rsplit("/", 1)[-1]
    kids = list(children)
    child.Children.Count = len(kids)
    child.Children.side_effect = lambda i: kids[i]
    return child


def _tree_value(value):
    """GetObjectTree's rendering: strings, "true"/"false", "" when unset."""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (str, int)):
        return str(value)
    return ""


def _object_tree_json(element, props=("Id", "Type", "Name", "Text", "Changeable")):
    """What GetObjectTree returns for a _child() screen."""
    def node(el):
        result = {"properties": {name: _tree_value(getattr(el, name)) for name in props}}
        try:
            kids = [el.Children(i) for i in range(el.Children.Count)]
        except Exception:
            kids = []
        if kids:
            result["children"] = [node(kid) for kid in kids]
        return result
    return json.dumps({"children": [node(element)]})


@pytest.fixture(params=["walk", "tree"])
def discovery_path(request):
    """Discovery reads the screen with GetObjectTree (SAP GUI 7.70 PL3+) or,
    on an older SAP GUI, by walking the elements: both must agree."""
    return request.param


def _serve(controller, screen, path):
    if path == "tree":
        controller._session.GetObjectTree.return_value = _object_tree_json(screen)
        controller._session.findById.side_effect = AssertionError("walked COM")
    else:
        controller._session.GetObjectTree.side_effect = AttributeError("GetObjectTree")
        controller._session.findById.return_value = screen


class TestSessionTraits:
    """SAP GUI release and the server's scripting mode, read once per bind."""

    def _bound(self, major=8100, patch=0, read_only=False):
        controller = _make_controller_with_session()
        controller._application = MagicMock(MajorVersion=major, Patchlevel=patch)
        controller._session.Info.ScriptingModeReadOnly = read_only
        controller._read_session_traits()
        return controller

    @pytest.mark.parametrize("major, patch, expected", [
        (8100, 0, "8.10 PL0"),   # live value on SAP GUI 8.10
        (7700, 3, "7.70 PL3"),
        (7600, 12, "7.60 PL12"),
        (8000, 7, "8.00 PL7"),
    ])
    def test_version_string(self, major, patch, expected):
        assert self._bound(major, patch).sap_gui_version == expected

    def test_session_info_reports_both(self):
        info = self._bound(read_only=1).get_session_info()  # COM gives a Byte

        assert info.sap_gui_version == "8.10 PL0"
        assert info.scripting_read_only is True

    def test_unreadable_traits_fall_back(self):
        controller = _make_controller_with_session()
        controller._application = MagicMock()
        type(controller._application).MajorVersion = PropertyMock(
            side_effect=Exception("Member not found"))
        type(controller._session.Info).ScriptingModeReadOnly = PropertyMock(
            side_effect=Exception("Member not found"))

        controller._read_session_traits()

        assert controller.sap_gui_version == ""
        assert controller.scripting_read_only is False

    def test_disconnect_forgets_them(self):
        controller = self._bound(read_only=True)

        controller.disconnect()

        assert (controller.sap_gui_version, controller.scripting_read_only) == ("", False)

    def test_attaching_reads_them(self):
        from mcp_sap_gui.sap_controller import SAPGUIController
        controller = SAPGUIController()
        session = MagicMock(Busy=False)
        session.Info.ScriptingModeReadOnly = True
        connection = MagicMock()
        connection.Children.Count = 1
        connection.Children.side_effect = lambda i: session
        engine = MagicMock(MajorVersion=7700, Patchlevel=3)

        with patch.object(controller, "_discover_connections",
                          return_value=[("saplogon", "SAPGUI", engine, connection)]), \
             patch.object(controller, "_is_scripting_disabled", return_value=False):
            info = controller.connect_to_existing_session(0, 0)

        assert (info.sap_gui_version, info.scripting_read_only) == ("7.70 PL3", True)


class TestShortIdsInResponses:
    """The session prefix is stripped on input anyway, so returning it only
    costs tokens and suggests a session addressing that does not exist."""

    def test_discovery_returns_short_ids(self):
        controller = _make_controller_with_session()
        usr = _child("/app/con[0]/ses[0]/wnd[0]/usr", "GuiUserArea", children=[
            _child("/app/con[0]/ses[0]/wnd[0]/usr/ctxtP_DEVID", "GuiCTextField"),
        ])
        controller._session.findById.return_value = usr

        elements = controller.get_screen_elements("wnd[0]/usr")

        assert [e.id for e in elements] == ["wnd[0]/usr/ctxtP_DEVID"]

    def test_popup_buttons_return_short_ids(self):
        controller = _make_controller_with_session()
        button = _child("/app/con[0]/ses[0]/wnd[1]/usr/btnBUTTON_1", "GuiButton", text="Yes")
        button.Tooltip = ""
        usr = _child("/app/con[0]/ses[0]/wnd[1]/usr", "GuiUserArea", children=[button])

        def find_by_id(element_id):
            if element_id == "wnd[1]":
                return MagicMock(Text="Confirm")
            if element_id == "wnd[1]/usr":
                return usr
            raise Exception("not found")
        controller._session.findById.side_effect = find_by_id

        popup = controller.get_popup_window()

        assert popup["buttons"][0]["id"] == "wnd[1]/usr/btnBUTTON_1"


class TestTableControlsStayOneElement:
    def _screen(self):
        cells = [
            _child(f"/app/con[0]/ses[0]/wnd[0]/usr/tblT/txtV-F[{c},{r}]",
                   "GuiTextField", changeable=True)
            for r in range(3) for c in range(2)
        ]
        table = _child("/app/con[0]/ses[0]/wnd[0]/usr/tblT", "GuiTableControl",
                       changeable=True, children=cells)
        button = _child("/app/con[0]/ses[0]/wnd[0]/usr/btnPOSI", "GuiButton")
        return _child("/app/con[0]/ses[0]/wnd[0]/usr", "GuiUserArea",
                      children=[table, button])

    def test_cells_are_not_listed_by_default(self, discovery_path):
        controller = _make_controller_with_session()
        _serve(controller, self._screen(), discovery_path)

        elements = controller.get_screen_elements("wnd[0]/usr", max_depth=3)

        assert [e.type for e in elements] == ["GuiTableControl", "GuiButton"]

    def test_expand_tables_lists_the_cells(self, discovery_path):
        controller = _make_controller_with_session()
        _serve(controller, self._screen(), discovery_path)

        elements = controller.get_screen_elements(
            "wnd[0]/usr", max_depth=3, expand_tables=True,
        )

        assert len(elements) == 2 + 6

    def test_changeable_filter_no_longer_floods_with_cells(self, discovery_path):
        """changeable_only on an SM30 screen used to return every input cell."""
        controller = _make_controller_with_session()
        _serve(controller, self._screen(), discovery_path)

        elements = controller.get_screen_elements(
            "wnd[0]/usr", max_depth=3, changeable_only=True,
        )

        assert [e.type for e in elements] == ["GuiTableControl"]

    async def test_tool_forwards_expand_tables(self, srv):
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        mock_ctrl.get_screen_elements.return_value = []
        mock_ctrl.get_docking_containers.return_value = []
        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            await srv.sap_get_screen_elements(ctx, expand_tables=True)

        assert mock_ctrl.get_screen_elements.call_args.kwargs["expand_tables"] is True


class TestClassicListsStayOutOfDiscovery:
    """One page of an SE16 standard list was 1,139 elements (124k characters),
    too large for the client to accept: every word is its own lbl[col,row]."""

    _PREFIX = "/app/con[0]/ses[0]/wnd[0]/usr"

    def _screen(self):
        cells = [_child(f"{self._PREFIX}/lbl[0,0]", "GuiLabel", text="Table:")]
        for row in (5, 6):
            cells.append(_child(f"{self._PREFIX}/chk[1,{row}]", "GuiCheckBox",
                                changeable=True))
            cells.append(_child(f"{self._PREFIX}/lbl[4,{row}]", "GuiLabel", text="001"))
        return _child(self._PREFIX, "GuiUserArea", children=cells)

    def test_cells_are_counted_not_listed(self, discovery_path):
        controller = _make_controller_with_session()
        _serve(controller, self._screen(), discovery_path)
        lists = []

        elements = controller.get_screen_elements("wnd[0]/usr", lists=lists)

        assert elements == []
        assert lists == [{"container": "wnd[0]/usr", "cells": 5, "rows": 3}]

    def test_cell_properties_are_not_read(self):
        """On the fallback walk a skipped cell costs one COM call (its Id),
        not five."""
        controller = _make_controller_with_session()
        usr = self._screen()
        text = PropertyMock(return_value="")
        for i in range(usr.Children.Count):
            type(usr.Children(i)).Text = text
        controller._session.findById.return_value = usr

        controller.get_screen_elements("wnd[0]/usr")

        assert text.call_count == 0

    def test_other_elements_are_kept(self, discovery_path):
        controller = _make_controller_with_session()
        usr = self._screen()
        button = _child(f"{self._PREFIX}/btnB", "GuiButton")
        kids = [button] + [usr.Children(i) for i in range(usr.Children.Count)]
        usr.Children.Count = len(kids)
        usr.Children.side_effect = lambda i: kids[i]
        _serve(controller, usr, discovery_path)

        elements = controller.get_screen_elements("wnd[0]/usr")

        assert [e.id for e in elements] == ["wnd[0]/usr/btnB"]

    def test_expand_tables_lists_the_cells(self, discovery_path):
        controller = _make_controller_with_session()
        _serve(controller, self._screen(), discovery_path)
        lists = []

        elements = controller.get_screen_elements(
            "wnd[0]/usr", expand_tables=True, lists=lists,
        )

        assert len(elements) == 5
        assert lists == []

    def test_table_control_cells_are_not_list_cells(self, discovery_path):
        """tblT/txtV-F[0,1] also ends in [col,row], but carries a field name."""
        cell = _child(f"{self._PREFIX}/tblT/txtV-F[0,1]", "GuiTextField")
        table = _child(f"{self._PREFIX}/tblT", "GuiTableControl", children=[cell])
        controller = _make_controller_with_session()
        _serve(controller, _child(self._PREFIX, "GuiUserArea", children=[table]),
               discovery_path)
        lists = []

        controller.get_screen_elements(
            "wnd[0]/usr", max_depth=3, expand_tables=True, lists=lists,
        )

        assert lists == []

    async def test_tool_reports_the_list_and_points_to_read_list(self, srv):
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        summary = {"container": "wnd[0]/usr", "cells": 1139, "rows": 58}

        def get_screen_elements(*args, lists, **kwargs):
            lists.append(summary)
            return []
        mock_ctrl.get_screen_elements.side_effect = get_screen_elements
        mock_ctrl.get_docking_containers.return_value = []
        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            result = await srv.sap_get_screen_elements(ctx)

        assert result["lists"] == [summary]
        assert "sap_read_list" in result["note"]

    async def test_tool_adds_nothing_without_a_list(self, srv):
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        mock_ctrl.get_screen_elements.return_value = []
        mock_ctrl.get_docking_containers.return_value = []
        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            result = await srv.sap_get_screen_elements(ctx)

        assert "lists" not in result and "note" not in result


class TestObjectTreeDiscovery:
    """GetObjectTree reads a whole screen in one call: 0.2 s against 4.7 s
    for walking a 1,139-cell list page element by element."""

    _PREFIX = "/app/con[0]/ses[0]/wnd[0]"

    def _nested(self):
        field = _child(f"{self._PREFIX}/usr/sub/ctxtP_F", "GuiCTextField",
                       changeable=True, text="X")
        sub = _child(f"{self._PREFIX}/usr/sub", "GuiSimpleContainer", children=[field])
        return _child(f"{self._PREFIX}/usr", "GuiUserArea", children=[sub])

    def test_one_call_reads_the_screen(self):
        controller = _make_controller_with_session()
        _serve(controller, self._nested(), "tree")

        elements = controller.get_screen_elements("wnd[0]/usr")

        controller._session.GetObjectTree.assert_called_once_with(
            "wnd[0]/usr", ["Id", "Type", "Name", "Text", "Changeable"],
        )
        assert [(e.id, e.text, e.changeable) for e in elements] == [
            ("wnd[0]/usr/sub", "", False),
            ("wnd[0]/usr/sub/ctxtP_F", "X", True),
        ]

    def test_max_depth_is_the_same_on_both_paths(self, discovery_path):
        controller = _make_controller_with_session()
        _serve(controller, self._nested(), discovery_path)

        elements = controller.get_screen_elements("wnd[0]/usr", max_depth=1)

        assert [e.id for e in elements] == ["wnd[0]/usr/sub"]

    def test_a_leaf_raising_on_children_keeps_its_siblings(self, discovery_path):
        """A status-bar pane raises a COM error (not AttributeError) on
        Children: the walk listed pane[0] and silently dropped the rest."""
        panes = [_child(f"{self._PREFIX}/sbar/pane[{i}]", "GuiStatusPane")
                 for i in range(3)]
        for pane in panes:
            type(pane).Children = PropertyMock(side_effect=IndexError("Children"))
        sbar = _child(f"{self._PREFIX}/sbar", "GuiStatusbar", children=panes)
        window = _child(self._PREFIX, "GuiMainWindow", children=[sbar])
        controller = _make_controller_with_session()
        _serve(controller, window, discovery_path)

        elements = controller.get_screen_elements("wnd[0]", max_depth=2)

        assert [e.id for e in elements] == [
            "wnd[0]/sbar", *(f"wnd[0]/sbar/pane[{i}]" for i in range(3)),
        ]

    def test_an_empty_tree_falls_back_to_walking(self):
        controller = _make_controller_with_session()
        controller._session.GetObjectTree.return_value = '{"children": []}'
        controller._session.findById.return_value = self._nested()

        elements = controller.get_screen_elements("wnd[0]/usr")

        assert len(elements) == 2

    def test_no_visible_flag(self, discovery_path):
        """SAP GUI Scripting has no Visible property: it was always True."""
        controller = _make_controller_with_session()
        _serve(controller, self._nested(), discovery_path)

        elements = controller.get_screen_elements("wnd[0]/usr")

        assert "visible" not in elements[0].__dict__


def _serve_window(controller, window, path):
    """Serve a _child() window to findById, and on the tree path also to
    GetObjectTree for any root inside it."""
    by_id = {}

    def index(el):
        by_id[el.Id.split("ses[0]/")[-1]] = el
        for i in range(el.Children.Count):
            index(el.Children(i))
    index(window)

    def find(element_id):
        if element_id in by_id:
            return by_id[element_id]
        raise Exception(f"not found: {element_id}")
    controller._session.findById.side_effect = find
    if path == "tree":
        controller._session.GetObjectTree.side_effect = (
            lambda root, props: _object_tree_json(by_id[root], props)
        )
    else:
        controller._session.GetObjectTree.side_effect = AttributeError("GetObjectTree")


class TestPopupAndListFromObjectTree:
    """The popup check (after every action that leaves a popup open) and
    read_list read through GetObjectTree too, with the COM walk as fallback."""

    _WND1 = "/app/con[0]/ses[0]/wnd[1]"

    def _popup(self):
        label = _child(f"{self._WND1}/usr/txtMSG", "GuiLabel", text=" Data will be lost ")
        field = _child(f"{self._WND1}/usr/subSUB/ctxtTRKORR", "GuiCTextField",
                       text="DEVK900001", changeable=True)
        yes = _child(f"{self._WND1}/usr/btnSPOP-OPTION1", "GuiButton", text="Yes")
        yes.Tooltip = "Save"
        sub = _child(f"{self._WND1}/usr/subSUB", "GuiSimpleContainer", children=[field])
        usr = _child(f"{self._WND1}/usr", "GuiUserArea", children=[label, sub, yes])
        enter = _child(f"{self._WND1}/tbar[0]/btn[0]", "GuiButton")
        enter.Tooltip = "Continue (Enter)"
        tbar = _child(f"{self._WND1}/tbar[0]", "GuiToolbar", children=[enter])
        sbar = _child(f"{self._WND1}/sbar", "GuiStatusbar", text="Check entry")
        sbar.MessageType = "W"
        return _child(self._WND1, "GuiModalWindow", text="Save changes?",
                      children=[usr, tbar, sbar])

    def test_popup_reads_the_same_on_both_paths(self, discovery_path):
        controller = _make_controller_with_session()
        _serve_window(controller, self._popup(), discovery_path)

        popup = controller.get_popup_window()

        assert popup["title"] == "Save changes?"
        assert (popup["message"], popup["message_type"]) == ("Check entry", "W")
        assert popup["texts"] == ["Data will be lost"]
        assert popup["buttons"] == [
            {"id": "wnd[1]/usr/btnSPOP-OPTION1", "text": "Yes", "tooltip": "Save"},
            {"id": "wnd[1]/tbar[0]/btn[0]", "text": "", "tooltip": "Continue (Enter)"},
        ]
        assert popup["interactive_elements"] == [{
            "id": "wnd[1]/usr/subSUB/ctxtTRKORR", "type": "GuiCTextField",
            "name": "ctxtTRKORR", "text": "DEVK900001", "changeable": True,
        }]
        assert popup["prefilled_inputs"][0]["value"] == "DEVK900001"

    def _hit_list(self):
        """F4 hit list as seen live: "Airline Code 18 Entries", Apply/Cancel."""
        cells = []
        for row, (code, name) in ((1, ("ID", "Airline")), (3, ("AA", "American Airlines")),
                                  (4, ("AB", "Air Berlin"))):
            cells.append(_child(f"{self._WND1}/usr/lbl[1,{row}]", "GuiLabel", text=code))
            cells.append(_child(f"{self._WND1}/usr/lbl[4,{row}]", "GuiLabel", text=name))
        usr = _child(f"{self._WND1}/usr", "GuiUserArea", children=cells)
        apply_ = _child(f"{self._WND1}/tbar[0]/btn[0]", "GuiButton")
        apply_.Tooltip = "Apply   (Enter)"
        cancel = _child(f"{self._WND1}/tbar[0]/btn[12]", "GuiButton")
        cancel.Tooltip = "Cancel   (F12)"
        tbar = _child(f"{self._WND1}/tbar[0]", "GuiToolbar", children=[apply_, cancel])
        return _child(self._WND1, "GuiModalWindow", text="Airline Code 18 Entries",
                      children=[usr, tbar])

    def test_hit_list_popup_is_a_list(self, discovery_path):
        """Its cells were 38 'texts', it was classified information, and
        auto-handling cancelled the value help (no button matched OK)."""
        controller = _make_controller_with_session()
        _serve_window(controller, self._hit_list(), discovery_path)

        popup = controller.get_popup_window()

        assert popup["classification"] == "list"
        assert popup["recommended_action"] == "read"
        assert "safe_auto_action" not in popup
        assert "texts" not in popup
        assert (popup["list"]["cells"], popup["list"]["rows"]) == (6, 3)
        assert "sap_read_list(window_id='wnd[1]')" in popup["list"]["hint"]

    def test_auto_only_reads_a_hit_list(self, discovery_path):
        controller = _make_controller_with_session()
        _serve_window(controller, self._hit_list(), discovery_path)

        result = controller.handle_popup("auto")

        assert result["action"] == "read"
        assert result["list"]["rows"] == 3

    def test_action_digest_carries_the_list(self):
        controller = _make_controller_with_session()
        _serve_window(controller, self._hit_list(), "tree")

        digest = controller._popup_digest()

        assert digest["classification"] == "list" and digest["list"]["cells"] == 6
        assert "texts" not in digest

    def _calendar(self):
        """F4 on a date field, as seen live: SAPLSCAC 100, Continue / Cancel."""
        cal = _child(f"{self._WND1}/usr/cntlCONTAINER/shellcont/shell", "GuiShell",
                     text="SAP.CalendarControl.1", changeable=True)
        cal.SubType = "Calendar"
        cal.focusDate = "20261005"
        shellcont = _child(f"{self._WND1}/usr/cntlCONTAINER/shellcont",
                           "GuiContainerShell", children=[cal])
        control = _child(f"{self._WND1}/usr/cntlCONTAINER", "GuiCustomControl",
                         children=[shellcont])
        usr = _child(f"{self._WND1}/usr", "GuiUserArea", children=[control])
        go = _child(f"{self._WND1}/tbar[0]/btn[0]", "GuiButton")
        go.Tooltip = "Continue   (Enter)"
        cancel = _child(f"{self._WND1}/tbar[0]/btn[12]", "GuiButton")
        cancel.Tooltip = "Cancel   (F12)"
        tbar = _child(f"{self._WND1}/tbar[0]", "GuiToolbar", children=[go, cancel])
        return _child(self._WND1, "GuiModalWindow", text="Calendar", children=[usr, tbar])

    def test_calendar_popup_is_a_date_picker(self, discovery_path):
        """It read as a confirmation; Continue writes the focused date."""
        controller = _make_controller_with_session()
        _serve_window(controller, self._calendar(), discovery_path)

        popup = controller.get_popup_window()

        assert popup["classification"] == "date_picker"
        assert "safe_auto_action" not in popup
        assert popup["calendar"]["id"] == "wnd[1]/usr/cntlCONTAINER/shellcont/shell"
        assert popup["calendar"]["focus_date"] == "20261005"
        assert "sap_set_field" in popup["calendar"]["hint"]
        assert controller.handle_popup("auto")["action"] == "read"

    def test_popup_tree_is_one_call(self):
        controller = _make_controller_with_session()
        _serve_window(controller, self._popup(), "tree")

        controller.get_popup_window()

        assert controller._session.GetObjectTree.call_count == 1

    def test_list_reads_the_same_on_both_paths(self, discovery_path):
        prefix = "/app/con[0]/ses[0]/wnd[0]/usr"
        heading = _child(f"{prefix}/lbl[0,0]", "GuiLabel", text="Table:")
        heading.ColorIndex = 1
        box = _child(f"{prefix}/chk[1,2]", "GuiCheckBox")
        box.Selected = True
        key = _child(f"{prefix}/lbl[4,2]", "GuiLabel", text="001")
        key.ColorIndex = 4
        usr = _child(prefix, "GuiUserArea", children=[heading, box, key])
        controller = _make_controller_with_session()
        _serve_window(controller, usr, discovery_path)

        page = controller.read_list(with_ids=True)

        assert page["lines"] == ["Table:", "", " [x]001"]
        assert page["colors"] == {"0": ["heading"], "2": ["key"]}
        assert page["line_ids"] == {"0": "wnd[0]/usr/lbl[0,0]",
                                    "2": "wnd[0]/usr/chk[1,2]"}

    def test_list_tree_is_one_call(self):
        prefix = "/app/con[0]/ses[0]/wnd[0]/usr"
        cells = [_child(f"{prefix}/lbl[{c},0]", "GuiLabel", text="x") for c in range(3)]
        controller = _make_controller_with_session()
        _serve_window(controller, _child(prefix, "GuiUserArea", children=cells), "tree")

        page = controller.read_list()

        assert page["lines"] == ["xxx"]
        controller._session.GetObjectTree.assert_called_once_with(
            "wnd[0]/usr", ["Id", "Text", "ColorIndex", "Selected"],
        )


class TestTableControlFromObjectTree:
    """read_table on a table control takes the visible cells from one
    GetObjectTree call instead of ~3 COM calls per cell."""

    _TID = "/app/con[0]/ses[0]/wnd[0]/usr/tblTEST"
    _NAMES = ["V-NAME", "V-ACTIVE", "V-TYPE"]
    _ROWS = [["A1", True, "01"], ["A2", False, "02"], ["A3", True, ""]]

    def _cell(self, row, col):
        """The cell in visible row *row* of the current scroll position."""
        abs_row = self._scroll.Position + row
        cell = MagicMock(Name=self._NAMES[col])
        cell.Children.Count = 0
        prefix, cell.Type = [("txt", "GuiTextField"), ("chk", "GuiCheckBox"),
                             ("cmb", "GuiComboBox")][col]
        cell.Id = f"{self._TID}/{prefix}{self._NAMES[col]}[{col},{row}]"
        value = self._ROWS[abs_row][col] if abs_row < len(self._ROWS) else ""
        cell.Text, cell.Selected, cell.Key = str(value), value is True, value
        if col == 1 and abs_row >= len(self._ROWS):
            cell.Selected = False
        return cell

    def _table(self, path):
        self._scroll = MagicMock(Minimum=0, Maximum=1, Position=0, PageSize=2)
        table = MagicMock(Id=self._TID, Type="GuiTableControl", TableFieldName="TEST",
                          RowCount=len(self._ROWS), VisibleRowCount=2,
                          VerticalScrollbar=self._scroll)
        table.Columns.Count = len(self._NAMES)
        table.Columns.side_effect = lambda i: MagicMock(Title=f"T{i}", Tooltip="")
        table.GetCell.side_effect = self._cell
        controller = _make_controller_with_session()
        controller._session.findById.return_value = table
        if path == "tree":
            def object_tree(root, props):
                cells = [self._cell(r, c) for r in range(2) for c in range(3)]
                return _object_tree_json(_child(self._TID, "GuiTableControl",
                                                children=cells), props)
            controller._session.GetObjectTree.side_effect = object_tree
        else:
            controller._session.GetObjectTree.side_effect = AttributeError("GetObjectTree")
        return controller, table

    def test_rows_are_the_same_on_both_paths(self, discovery_path):
        controller, _table = self._table(discovery_path)

        result = controller.read_table("wnd[0]/usr/tblTEST")

        assert result["data"] == [
            {"V-NAME": "A1", "V-ACTIVE": True, "V-TYPE": "01", "_absolute_row_index": 0},
            {"V-NAME": "A2", "V-ACTIVE": False, "V-TYPE": "02", "_absolute_row_index": 1},
        ]

    def test_scrolled_rows_are_the_same_on_both_paths(self, discovery_path):
        controller, _table = self._table(discovery_path)

        result = controller.read_table("wnd[0]/usr/tblTEST", start_row=1,
                                       columns="V-NAME")

        assert result["data"] == [
            {"V-NAME": "A2", "_absolute_row_index": 1},
            {"V-NAME": "A3", "_absolute_row_index": 2},
        ]

    def test_tree_path_reads_no_cell_values_over_com(self):
        controller, table = self._table("tree")

        controller.read_table("wnd[0]/usr/tblTEST")

        # Column names still come from row 0 (three cells); the six data
        # cells came from the one GetObjectTree call.
        assert controller._session.GetObjectTree.call_count == 1
        assert table.GetCell.call_count == 3


class TestTableControlColumnTemplates:
    def _table(self, cells):
        table = MagicMock()
        table.Columns.Count = len(cells)
        table.Columns.side_effect = lambda i: MagicMock(Title=f"Title {i}", Tooltip="")

        def get_cell(row, col):
            if cells[col] is None:
                raise Exception("no cell")
            return cells[col]
        table.GetCell.side_effect = get_cell
        return table

    def test_schema_carries_cell_type_and_id_template(self):
        controller = _make_controller_with_session()
        key = MagicMock(Id="/app/con[0]/ses[0]/wnd[0]/usr/tblT/ctxtV_T005-LAND1[0,0]",
                        Type="GuiCTextField")
        key.Name = "V_T005-LAND1"
        text = MagicMock(Id="/app/con[0]/ses[0]/wnd[0]/usr/tblT/txtV_T005-LANDX[1,0]",
                         Type="GuiTextField")
        text.Name = "V_T005-LANDX"

        columns = controller._get_table_control_columns(
            self._table([key, text]), with_cell_ids=True,
        )

        assert columns[0]["name"] == "V_T005-LAND1"
        assert columns[0]["cell_type"] == "GuiCTextField"
        assert columns[0]["cell_id"] == "wnd[0]/usr/tblT/ctxtV_T005-LAND1[0,{row}]"
        assert columns[1]["cell_id"] == "wnd[0]/usr/tblT/txtV_T005-LANDX[1,{row}]"
        assert "name_is_title" not in columns[0]

    def test_data_reads_stay_lean(self):
        controller = _make_controller_with_session()
        cell = MagicMock(Id="/app/con[0]/ses[0]/wnd[0]/usr/tblT/txtF[0,0]", Type="GuiTextField")
        cell.Name = "F"

        columns = controller._get_table_control_columns(self._table([cell]))

        assert "cell_id" not in columns[0] and "cell_type" not in columns[0]

    def test_empty_table_flags_that_names_are_titles(self):
        """Seen live: an empty view in display mode has no cells, so the
        technical names cannot be read and the titles were passed off as names."""
        controller = _make_controller_with_session()

        columns = controller._get_table_control_columns(
            self._table([None, None]), with_cell_ids=True,
        )

        assert [c["name"] for c in columns] == ["Title 0", "Title 1"]
        assert all(c["name_is_title"] for c in columns)
        assert "cell_id" not in columns[0]

    def test_read_table_explains_the_title_fallback(self):
        controller = _make_controller_with_session()
        table = self._table([None])
        table.RowCount = 0
        table.VisibleRowCount = 10

        result = controller._read_table_control(
            table, "wnd[0]/usr/tblT", max_rows=10, columns_only=True,
        )

        assert "TITLES" in result["note"]


class TestCompactBatchResult:
    def _controller(self):
        controller = _make_controller_with_session()

        def find_by_id(element_id):
            if element_id.endswith("BAD"):
                raise Exception("not found")
            return MagicMock()
        controller._session.findById.side_effect = find_by_id
        return controller

    def test_only_failures_are_listed(self):
        fields = {f"wnd[0]/usr/txtF{i}": "x" for i in range(70)}
        fields["wnd[0]/usr/txtBAD"] = "x"

        result = self._controller().set_batch_fields(fields)

        assert (result["total"], result["succeeded"], result["failed"]) == (71, 70, 1)
        assert list(result["results"]) == ["wnd[0]/usr/txtBAD"]

    def test_all_good_gives_an_empty_results_dict(self):
        result = self._controller().set_batch_fields({"wnd[0]/usr/txtF1": "x"})

        assert result["succeeded"] == 1
        assert result["results"] == {}

    def test_verbose_lists_every_field(self):
        result = self._controller().set_batch_fields(
            {"wnd[0]/usr/txtF1": "x", "wnd[0]/usr/txtBAD": "x"}, verbose=True,
        )

        assert result["results"]["wnd[0]/usr/txtF1"] == "success"
        assert result["results"]["wnd[0]/usr/txtBAD"].startswith("error")


class TestCompactPopupResult:
    def _controller(self, popup_after):
        controller = _make_controller_with_session()
        popup = {
            "popup_exists": True, "window_id": "wnd[1]", "title": "Prompt for Request",
            "texts": ["Request"], "has_inputs": True,
            "classification": "input_required", "recommended_action": "read",
            "buttons": [{"id": "wnd[1]/tbar[0]/btn[0]", "text": "Continue", "tooltip": ""}],
            "interactive_elements": [
                {"id": "wnd[1]/usr/ctxtKO008-TRKORR", "type": "GuiCTextField",
                 "name": "KO008-TRKORR", "text": "DEVK900123", "changeable": True},
                {"id": "wnd[1]/usr/txtEMPTY", "type": "GuiTextField",
                 "name": "EMPTY", "text": "", "changeable": True},
                # Seen live: selection popups carry read-only description
                # fields with text. They are labels, not accepted values.
                {"id": "wnd[1]/usr/txt%_P_LGNUM_%_APP_%-TEXT", "type": "GuiTextField",
                 "name": "%_P_LGNUM_%_APP_%-TEXT", "text": "Warehouse Number",
                 "changeable": False},
            ],
        }
        controller.get_popup_window = MagicMock(side_effect=[popup, popup_after])
        controller.get_screen_info = MagicMock(return_value={"active_window": "wnd[0]"})
        return controller

    def test_closed_popup_keeps_the_trail_not_the_dead_ids(self):
        controller = self._controller({"popup_exists": False})

        result = controller.handle_popup("confirm")

        assert result["action"] == "confirmed"
        assert result["popup_closed"] is True
        assert result["title"] == "Prompt for Request"
        assert result["inputs"] == {"KO008-TRKORR": "DEVK900123"}
        for dead in ("buttons", "interactive_elements", "popup_after"):
            assert dead not in result

    def test_a_follow_up_popup_stays_complete(self):
        follow_up = {"popup_exists": True, "window_id": "wnd[1]", "title": "Next",
                     "buttons": [{"id": "wnd[1]/tbar[0]/btn[0]", "text": "OK"}]}
        controller = self._controller(follow_up)

        result = controller.handle_popup("confirm")

        assert result["popup_closed"] is False
        assert result["popup_after"]["buttons"][0]["text"] == "OK"

    def test_read_is_still_complete(self):
        controller = self._controller({"popup_exists": False})

        result = controller.handle_popup("read")

        assert len(result["interactive_elements"]) == 3
        assert result["buttons"]


# ===========================================================================
# Popup notice
# ===========================================================================

class TestPrefilledPopupNotice:
    """Generic on purpose: no popup is recognised by name. Any dialog whose
    changeable inputs already hold a value gets the same notice."""

    def _popup_controller(self, inputs):
        controller = _make_controller_with_session()
        children = []
        for element_id, type_, text, changeable in inputs:
            child = _child(element_id, type_, text=text, changeable=changeable)
            children.append(child)
        usr = _child("/app/con[0]/ses[0]/wnd[1]/usr", "GuiUserArea", children=children)

        def find_by_id(element_id):
            if element_id == "wnd[1]":
                return MagicMock(Text="Prompt for Request")
            if element_id == "wnd[1]/usr":
                return usr
            raise Exception("not found")
        controller._session.findById.side_effect = find_by_id
        return controller

    def test_prefilled_changeable_input_gets_a_notice(self):
        controller = self._popup_controller([
            ("/app/con[0]/ses[0]/wnd[1]/usr/ctxtKO008-TRKORR", "GuiCTextField",
             "DEVK900123", True),
        ])

        popup = controller.get_popup_window()

        assert popup["prefilled_inputs"] == [{
            "id": "wnd[1]/usr/ctxtKO008-TRKORR",
            "name": "ctxtKO008-TRKORR",
            "value": "DEVK900123",
        }]
        assert "pre-filled" in popup["notice"]

    def test_no_notice_without_a_prefilled_changeable_input(self):
        controller = self._popup_controller([
            ("/app/con[0]/ses[0]/wnd[1]/usr/ctxtEMPTY", "GuiCTextField", "", True),
            ("/app/con[0]/ses[0]/wnd[1]/usr/txtSHOWN", "GuiTextField", "fixed", False),
            ("/app/con[0]/ses[0]/wnd[1]/usr/chkFLAG", "GuiCheckBox", "A label", True),
        ])

        popup = controller.get_popup_window()

        assert "prefilled_inputs" not in popup
        assert "notice" not in popup

    def test_secret_values_are_masked(self):
        controller = self._popup_controller([
            ("/app/con[0]/ses[0]/wnd[1]/usr/txtRSYST-BCODE", "GuiTextField", "s3cret", True),
        ])

        popup = controller.get_popup_window()

        assert popup["prefilled_inputs"][0]["value"] == "***"

    def _screen_info_controller(self, active_window):
        controller = _make_controller_with_session()
        controller._session.ActiveWindow = MagicMock(Id=active_window)
        controller._session.findById.return_value = MagicMock(Text="Title")
        controller.get_popup_window = MagicMock(return_value={
            "popup_exists": True, "title": "Prompt for Request",
            "classification": "input_required", "texts": ["Request"],
            "buttons": [{"id": "wnd[1]/tbar[0]/btn[0]", "text": "", "tooltip": "Continue"}],
            "prefilled_inputs": [{"id": "wnd[1]/usr/ctxtX", "name": "X", "value": "V"}],
            "notice": "This popup has pre-filled input values.",
        })
        return controller

    def test_action_responses_carry_the_digest_when_a_popup_opened(self):
        controller = self._screen_info_controller("/app/con[0]/ses[0]/wnd[1]")

        screen = controller.get_screen_info()

        assert screen["active_window"] == "wnd[1]"
        assert screen["popup"] == {
            "classification": "input_required",
            "texts": ["Request"],
            "buttons": ["Continue"],
            "prefilled_inputs": [{"id": "wnd[1]/usr/ctxtX", "name": "X", "value": "V"}],
            "notice": "This popup has pre-filled input values.",
        }

    def test_no_digest_on_the_main_window(self):
        controller = self._screen_info_controller("/app/con[0]/ses[0]/wnd[0]")

        screen = controller.get_screen_info()

        assert "popup" not in screen
        controller.get_popup_window.assert_not_called()

    def test_digest_can_be_switched_off(self):
        controller = self._screen_info_controller("/app/con[0]/ses[0]/wnd[1]")

        assert "popup" not in controller.get_screen_info(include_popup=False)


# ===========================================================================
# SM30 / SM34 guide
# ===========================================================================

class TestSm30Guide:
    @pytest.mark.parametrize("alias", [
        "SM30", "sm34", "table maintenance", "View Cluster", " view maintenance ",
    ])
    async def test_aliases_resolve_to_the_one_guide(self, srv, alias):
        result = await srv.sap_get_transaction_guide(alias)

        assert result["transaction"] == "SM30"
        assert result["mode"] == "read-first"

    async def test_guide_carries_the_lessons_learned_live(self, srv):
        guide = (await srv.sap_get_transaction_guide("SM34", task="add RF steps"))["guide"]

        assert "add RF steps" in guide
        for lesson in (
            "wnd[0]/shellcont",              # the tree is a docking container
            "sap_double_click_tree_item",    # ...and needs an item double-click
            "does NOT switch the view",
            "Select the parent row first",
            "cell_id",                       # cell IDs come from the schema
            "name_is_title",
            "prefilled_inputs",              # the customizing request prompt
            "Never confirm it blindly",
        ):
            assert lesson in guide, lesson


# ===========================================================================
# Classic lists
# ===========================================================================

class TestReadList:
    def _label(self, col, row, text, color=0):
        return MagicMock(Id=f"/app/con[0]/ses[0]/wnd[0]/usr/lbl[{col},{row}]",
                         Text=text, ColorIndex=color)

    def _controller(self, children, scroll_max=0):
        controller = _make_controller_with_session()
        usr = MagicMock()
        usr.Children.Count = len(children)
        usr.Children.side_effect = lambda i: children[i]
        usr.VerticalScrollbar = MagicMock(Maximum=scroll_max, Position=0, PageSize=40)
        usr.HorizontalScrollbar = MagicMock(Maximum=0, Position=0, PageSize=100)
        controller._session.findById.return_value = usr
        return controller, usr

    def test_labels_become_lines_at_their_columns(self):
        controller, _ = self._controller([
            self._label(10, 1, "4711"),          # out of order on purpose
            self._label(0, 1, "Device"),
            self._label(0, 0, "Log"),
            self._label(0, 3, "Done"),
        ])

        result = controller.read_list()

        assert result["is_list"] is True
        assert result["first_row"] == 0
        assert result["lines"] == ["Log", "Device    4711", "", "Done"]
        assert "colors" not in result and "scroll" not in result

    def test_semantic_colours_are_reported_per_row(self):
        controller, _ = self._controller([
            self._label(0, 0, "ok", color=2),
            self._label(0, 1, "Error: device unknown", color=6),
            self._label(0, 2, "Total", color=3),
        ])

        result = controller.read_list()

        assert result["colors"] == {"1": ["negative"], "2": ["total"]}

    def test_checkboxes_render_inline(self):
        box = MagicMock(Id="/app/con[0]/ses[0]/wnd[0]/usr/chk[0,0]", Selected=True)
        controller, _ = self._controller([box, self._label(3, 0, "Released")])

        assert controller.read_list()["lines"] == ["[x]Released"]

    def test_long_list_reports_scroll_state_and_pages(self):
        controller, usr = self._controller([self._label(0, 0, "row")], scroll_max=500)

        result = controller.read_list(scroll_to=40)

        assert usr.VerticalScrollbar.Position == 40
        assert result["scroll"]["vertical"]["maximum"] == 500
        assert result["scroll"]["vertical"]["page_size"] == 40
        # re-acquired after the scroll: the first reference is stale by then
        assert controller._session.findById.call_count == 2

    def test_max_lines_truncates(self):
        controller, _ = self._controller([self._label(0, r, f"l{r}") for r in range(5)])

        result = controller.read_list(max_lines=2)

        assert result["lines"] == ["l0", "l1"]
        assert result["truncated"] is True

    def test_with_ids_gives_a_focusable_label_per_line(self):
        controller, _ = self._controller([self._label(4, 2, "x")])

        result = controller.read_list(with_ids=True)

        assert result["line_ids"] == {"2": "wnd[0]/usr/lbl[4,2]"}

    def test_a_screen_without_labels_is_not_a_list(self):
        field = MagicMock(Id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtP_DEVID")
        controller, _ = self._controller([field])

        result = controller.read_list()

        assert result["is_list"] is False
        assert "read_table" in result["note"]

    def test_invalid_window_is_rejected(self):
        controller, _ = self._controller([])

        assert "Invalid SAP window ID" in controller.read_list("wnd[0]/usr")["error"]

    async def test_tool_forwards_its_arguments(self, srv):
        ctx = _make_mock_ctx()
        mock_ctrl = MagicMock()
        mock_ctrl.read_list.return_value = {"is_list": True, "lines": []}
        with patch.object(srv, "_ctrl", return_value=mock_ctrl):
            await srv.sap_read_list(ctx, window_id="wnd[1]", scroll_to=47, with_ids=True)

        mock_ctrl.read_list.assert_called_once_with(
            "wnd[1]", max_lines=200, scroll_to=47, with_ids=True,
        )
