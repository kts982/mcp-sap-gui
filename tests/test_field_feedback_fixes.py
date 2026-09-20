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
    child = MagicMock(Id=element_id, Type=type_, Text=text,
                      Changeable=changeable, Visible=True)
    child.Name = element_id.rsplit("/", 1)[-1]
    kids = list(children)
    child.Children.Count = len(kids)
    child.Children.side_effect = lambda i: kids[i]
    return child


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

    def test_cells_are_not_listed_by_default(self):
        controller = _make_controller_with_session()
        controller._session.findById.return_value = self._screen()

        elements = controller.get_screen_elements("wnd[0]/usr", max_depth=3)

        assert [e.type for e in elements] == ["GuiTableControl", "GuiButton"]

    def test_expand_tables_lists_the_cells(self):
        controller = _make_controller_with_session()
        controller._session.findById.return_value = self._screen()

        elements = controller.get_screen_elements(
            "wnd[0]/usr", max_depth=3, expand_tables=True,
        )

        assert len(elements) == 2 + 6

    def test_changeable_filter_no_longer_floods_with_cells(self):
        """changeable_only on an SM30 screen used to return every input cell."""
        controller = _make_controller_with_session()
        controller._session.findById.return_value = self._screen()

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
