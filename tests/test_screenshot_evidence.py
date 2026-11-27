"""Screenshot evidence: capture by handle guards, highlight, mask, caption."""

import sys
from contextlib import contextmanager
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import pytest

from mcp_sap_gui.models import SAPGUIError

Image = pytest.importorskip("PIL.Image")

from mcp_sap_gui.capture import annotate_png  # noqa: E402
from mcp_sap_gui.sap_controller import SAPGUIController  # noqa: E402


def _white_png(path: Path, size=(200, 100)) -> None:
    Image.new("RGB", size, (255, 255, 255)).save(path, "PNG")


class TestAnnotate:
    def test_mask_is_black_and_highlight_is_red(self, tmp_path):
        png = tmp_path / "a.png"
        _white_png(png)

        annotate_png(str(png), mask=[(10, 10, 40, 20)], highlight=[(100, 40, 50, 30)])

        with Image.open(png) as img:
            assert img.getpixel((20, 15)) == (0, 0, 0)
            assert img.getpixel((100, 40 - 2)) == (255, 0, 0)  # frame edge
            assert img.getpixel((180, 90)) == (255, 255, 255)

    def test_highlight_cannot_expose_masked_pixels(self, tmp_path):
        png = tmp_path / "a.png"
        _white_png(png)

        annotate_png(str(png), mask=[(50, 30, 60, 30)], highlight=[(50, 30, 60, 30)])

        with Image.open(png) as img:
            assert img.getpixel((80, 45)) == (0, 0, 0)

    @pytest.mark.parametrize("box", [(300, 10, 20, 20), (-50, 10, 20, 20), (10, 150, 20, 20)])
    def test_a_target_outside_the_image_is_refused(self, tmp_path, box):
        png = tmp_path / "a.png"
        _white_png(png)

        with pytest.raises(SAPGUIError, match="outside the captured window"):
            annotate_png(str(png), mask=[box])

    def test_caption_adds_a_strip_below(self, tmp_path):
        png = tmp_path / "a.png"
        _white_png(png)

        annotate_png(str(png), caption="Field under review", meta="DEV/100 user TESTUSER tcode SU3")

        with Image.open(png) as img:
            assert img.width == 200 and img.height > 100
            assert img.getpixel((0, 99)) == (255, 255, 255)  # original area intact


def _controller(tmp_path):
    controller = SAPGUIController()
    controller._session = MagicMock(Busy=False)
    controller._session.ActiveWindow = MagicMock(Id="wnd[0]")
    window = MagicMock(ScreenLeft=0, ScreenTop=100, Width=200, Height=100)
    window.HardCopy.side_effect = lambda path, _fmt: _white_png(Path(path))
    field = MagicMock(ScreenLeft=10, ScreenTop=110, Width=40, Height=20)
    usr = MagicMock(ScreenLeft=0, ScreenTop=100, Width=200, Height=100)
    label = MagicMock(ScreenLeft=60, ScreenTop=110, Width=40, Height=20)
    button = MagicMock(ScreenLeft=5, ScreenTop=102, Width=20, Height=20)
    known = {
        "wnd[0]/tbar[1]/btn[8]": button,
        "wnd[0]": window, "wnd[0]/usr": usr, "wnd[0]/usr/txtF": field,
        "wnd[0]/usr/lbl%#AUTOTEXT001": label,
    }

    def find(element_id):
        if element_id not in known:
            raise RuntimeError("nope")
        return known[element_id]

    controller._session.findById.side_effect = find
    controller._known = known
    return controller, window


class TestSaveScreenshotEvidence:
    def test_masks_element_relative_to_window_origin(self, tmp_path):
        controller, _ = _controller(tmp_path)
        target = tmp_path / "e.png"

        result = controller.save_screenshot(
            str(target), inline=False, mask=["wnd[0]/usr/txtF"], highlight=["wnd[0]/usr/txtF"],
        )

        assert "error" not in result, result
        assert result["masked"] == ["wnd[0]/usr/txtF"]
        with Image.open(target) as img:
            assert img.getpixel((20, 15)) == (0, 0, 0)  # (10,110) minus origin (0,100)

    def test_an_id_with_a_hash_is_an_element_id(self, tmp_path):
        controller, _ = _controller(tmp_path)

        result = controller.save_screenshot(
            str(tmp_path / "e.png"), inline=False, mask=["wnd[0]/usr/lbl%#AUTOTEXT001"],
        )

        assert "error" not in result, result

    def test_a_toolbar_element_is_checked_against_the_window(self, tmp_path):
        controller, _ = _controller(tmp_path)

        result = controller.save_screenshot(
            str(tmp_path / "e.png"), inline=False, mask=["wnd[0]/tbar[1]/btn[8]"],
        )

        assert "error" not in result, result

    def test_a_target_outside_its_container_saves_nothing(self, tmp_path):
        controller, _ = _controller(tmp_path)
        controller._known["wnd[0]/usr"].Height = 5  # a small window: field below the user area
        target = tmp_path / "e.png"

        result = controller.save_screenshot(str(target), mask=["wnd[0]/usr/txtF"])

        assert "outside its container" in result["error"]
        assert not target.exists()

    def test_a_target_without_size_saves_nothing(self, tmp_path):
        controller, _ = _controller(tmp_path)
        controller._known["wnd[0]/usr/txtF"].Width = 0

        result = controller.save_screenshot(str(tmp_path / "e.png"), highlight=["wnd[0]/usr/txtF"])

        assert "no size" in result["error"]

    def test_a_screen_that_moved_during_the_capture_saves_nothing(self, tmp_path):
        controller, window = _controller(tmp_path)
        field = controller._known["wnd[0]/usr/txtF"]
        target = tmp_path / "e.png"

        def scroll_then_write(path, _fmt):
            field.ScreenTop = 140  # the user scrolled meanwhile
            _white_png(Path(path))

        window.HardCopy.side_effect = scroll_then_write

        result = controller.save_screenshot(str(target), mask=["wnd[0]/usr/txtF"])

        assert "changed during the capture" in result["error"]
        assert not target.exists()

    def test_an_image_of_another_size_than_the_window_saves_nothing(self, tmp_path):
        controller, window = _controller(tmp_path)
        window.Width = 400  # SAP says 400 wide, the image is 200 (another monitor scaling)
        target = tmp_path / "e.png"

        result = controller.save_screenshot(str(target), mask=["wnd[0]/usr/txtF"])

        assert "SAP reports the window" in result["error"]
        assert not target.exists()

    def test_an_image_scaled_the_same_in_both_directions_is_drawn_on(self, tmp_path):
        # Seen live: a window on a 100% monitor, SAP started at 150%, is reported
        # 1.5 times the image on every length (2880x1430 for a 1920x953 image).
        controller, window = _controller(tmp_path)
        window.Width, window.Height = 400, 200  # the image is 200x100: k = 0.5
        field = controller._known["wnd[0]/usr/txtF"]
        field.ScreenLeft, field.ScreenTop, field.Width, field.Height = 20, 120, 80, 40
        controller._known["wnd[0]/usr"].Height = 200
        target = tmp_path / "e.png"

        result = controller.save_screenshot(str(target), inline=False, mask=["wnd[0]/usr/txtF"])

        assert "error" not in result, result
        with Image.open(target) as img:  # (20,120) minus window (0,100), halved: (10,10, 40x20)
            assert img.getpixel((15, 15)) == (0, 0, 0)
            assert img.getpixel((60, 15)) == (255, 255, 255)

    def test_a_one_pixel_difference_is_not_a_scaling(self, tmp_path):
        controller, window = _controller(tmp_path)
        window.Width = 199  # the image is 200 wide
        target = tmp_path / "e.png"

        result = controller.save_screenshot(str(target), mask=["wnd[0]/usr/txtF"])

        assert "SAP reports the window" in result["error"]
        assert not target.exists()

    def test_a_cell_is_drawn_where_it_is_on_a_scaled_window(self, tmp_path):
        # State 3 end to end: units 2x the image (k 0.5), cells in the same units.
        controller, window = _controller(tmp_path)
        window.Width, window.Height = 400, 200
        controller._known["wnd[0]/usr"].Height = 200
        grid = MagicMock(ScreenLeft=20, ScreenTop=140, Width=200, Height=100,
                         FirstVisibleRow=0, VisibleRowCount=5, ColumnCount=1)
        grid.ColumnOrder.side_effect = lambda i: "COL_A"
        grid.GetCellLeft.return_value, grid.GetCellTop.return_value = 10, 20
        grid.GetCellWidth.return_value, grid.GetCellHeight.return_value = 40, 20
        controller._known["wnd[0]/usr/shell"] = grid
        target = tmp_path / "e.png"

        real_unit = SAPGUIController._cell_unit  # the real function, with a handle
        with patch("mcp_sap_gui.discovery.cell_scale", return_value=0.5), patch(
            "mcp_sap_gui.discovery.physical_width", return_value=200
        ), patch.object(
            SAPGUIController, "_cell_unit", side_effect=lambda _hwnd, w: real_unit(7, w)
        ):
            result = controller.save_screenshot(
                str(target), inline=False,
                mask=[{"id": "wnd[0]/usr/shell", "column": "COL_A", "row": 0}],
            )

        assert "error" not in result, result
        with Image.open(target) as img:  # (30,60) 40x20 units -> (15,30) 20x10 pixels
            assert img.getpixel((20, 35)) == (0, 0, 0)
            assert img.getpixel((50, 35)) == (255, 255, 255)

    def test_a_mask_on_a_scaled_image_covers_the_pixel_its_edge_falls_in(self, tmp_path):
        # Units of 2/3 pixel: the field starts at 14.67 px, so pixel 14 shows part of it.
        # Rounded to the nearest pixel the mask began at 15.
        controller, window = _controller(tmp_path)
        window.Width, window.Height = 300, 150  # the image is 200x100
        field = controller._known["wnd[0]/usr/txtF"]
        field.ScreenLeft, field.ScreenTop, field.Width, field.Height = 22, 122, 80, 40
        controller._known["wnd[0]/usr"].Height = 150
        target = tmp_path / "e.png"

        result = controller.save_screenshot(str(target), inline=False, mask=["wnd[0]/usr/txtF"])

        assert "error" not in result, result
        with Image.open(target) as img:
            assert img.getpixel((14, 14)) == (0, 0, 0)
            assert img.getpixel((12, 12)) == (255, 255, 255)

    def test_a_target_in_another_window_than_the_capture_saves_nothing(self, tmp_path):
        controller, window = _controller(tmp_path)
        controller._session.ActiveWindow = MagicMock(Id="wnd[1]")  # a popup is on top
        controller._known["wnd[1]"] = window
        target = tmp_path / "e.png"

        result = controller.save_screenshot(str(target), mask=["wnd[0]/usr/txtF"])

        assert "the capture is wnd[1]" in result["error"]
        assert not target.exists()

    def test_the_image_is_saved_across_drives(self, tmp_path):
        controller, _ = _controller(tmp_path)
        target = tmp_path / "e.png"

        with patch("os.replace", side_effect=OSError(17, "not the same drive")):
            result = controller.save_screenshot(str(target), inline=False)

        assert "error" not in result, result
        assert target.is_file()

    def test_an_existing_file_is_never_overwritten(self, tmp_path):
        target = tmp_path / "e.png"
        target.write_bytes(b"mine")
        source = tmp_path / "new.png"
        source.write_bytes(b"new")

        with pytest.raises(FileExistsError):
            SAPGUIController._save_new(str(source), str(target))

        assert target.read_bytes() == b"mine"

    def test_unlocatable_mask_saves_nothing(self, tmp_path):
        controller, _ = _controller(tmp_path)
        target = tmp_path / "e.png"

        result = controller.save_screenshot(str(target), mask=["wnd[0]/usr/txtGONE"])

        assert "error" in result
        assert not target.exists()
        assert list(tmp_path.iterdir()) == []  # no unmasked temp left in the target dir


@contextmanager
def _plain_sap_gui(maximized: bool):
    """Plain SAP GUI (not embedded); yields the mock that is asked whether it is maximized."""
    with (
        patch("mcp_sap_gui.discovery.is_embedded", return_value=False),
        patch("mcp_sap_gui.discovery.is_maximized", return_value=maximized) as asked,
    ):
        yield asked


class TestMaximizedWindow:
    """SAP reports a maximized window with its invisible frame; HardCopy leaves it out."""

    @staticmethod
    def _maximized(tmp_path, frame=12):
        controller, window = _controller(tmp_path)
        # The 200x100 image is the window less the frame on every side.
        window.Handle = 4242
        window.ScreenLeft, window.ScreenTop = -frame, 100 - frame
        window.Width, window.Height = 200 + 2 * frame, 100 + 2 * frame
        return controller, window

    def test_the_image_starts_one_frame_inside_the_window(self, tmp_path):
        controller, _ = self._maximized(tmp_path)
        target = tmp_path / "e.png"

        with _plain_sap_gui(maximized=True):
            result = controller.save_screenshot(
                str(target), inline=False, mask=["wnd[0]/usr/txtF"],
            )

        assert "error" not in result, result
        with Image.open(target) as img:
            assert img.getpixel((15, 15)) == (0, 0, 0)  # (10,110) minus the image's (0,100)
            assert img.getpixel((55, 35)) == (255, 255, 255)  # not at SAP's origin (-12,88)

    def test_a_window_that_is_not_maximized_is_still_refused(self, tmp_path):
        controller, _ = self._maximized(tmp_path)
        target = tmp_path / "e.png"

        with _plain_sap_gui(maximized=False):
            result = controller.save_screenshot(str(target), mask=["wnd[0]/usr/txtF"])

        assert "SAP reports the window" in result["error"]
        assert not target.exists()

    def test_a_difference_that_is_not_a_frame_is_still_refused(self, tmp_path):
        controller, window = self._maximized(tmp_path)
        window.Height = 110  # 24 wider but 10 higher: not the same frame on every side
        target = tmp_path / "e.png"

        with _plain_sap_gui(maximized=True):
            result = controller.save_screenshot(str(target), mask=["wnd[0]/usr/txtF"])

        assert "SAP reports the window" in result["error"]
        assert not target.exists()

    def test_without_a_window_handle_it_is_refused(self, tmp_path):
        controller, window = self._maximized(tmp_path)
        window.Handle = MagicMock()  # Windows cannot be asked about this window
        target = tmp_path / "e.png"

        with _plain_sap_gui(maximized=True) as asked:
            result = controller.save_screenshot(str(target), mask=["wnd[0]/usr/txtF"])

        assert "SAP reports the window" in result["error"]
        assert not target.exists()
        asked.assert_not_called()

    def test_a_status_bar_reaching_into_the_frame_is_cut_at_the_image_edge(self, tmp_path):
        controller, _ = self._maximized(tmp_path)
        # Inside the window SAP reports (bottom 212), 3 px below the image (bottom 200).
        controller._known["wnd[0]/sbar"] = MagicMock(
            ScreenLeft=0, ScreenTop=180, Width=200, Height=23,
        )
        target = tmp_path / "e.png"

        with _plain_sap_gui(maximized=True):
            result = controller.save_screenshot(str(target), inline=False, mask=["wnd[0]/sbar"])

        assert "error" not in result, result
        with Image.open(target) as img:
            assert img.size == (200, 100)
            assert img.getpixel((100, 99)) == (0, 0, 0)  # down to the last row
            assert img.getpixel((100, 75)) == (255, 255, 255)

    @pytest.mark.parametrize("box, cut", [
        ((0, 80, 200, 23), (0, 80, 200, 20)),      # 3 px into the frame below
        ((-5, -5, 30, 30), (0, 0, 25, 25)),        # into the frame at the top left
        ((10, 10, 40, 20), (10, 10, 40, 20)),      # inside: untouched
        ((0, 80, 200, 40), (0, 80, 200, 40)),      # further than the frame: left for the refusal
        ((-12, 10, 12, 20), (-12, 10, 12, 20)),    # only in the frame: nothing would be left
    ])
    def test_a_box_is_cut_by_the_frame_and_no_further(self, box, cut):
        from mcp_sap_gui.capture import cut_at_frame

        assert cut_at_frame(box, (200, 100), 12) == cut

    @pytest.mark.parametrize("box, k, size, scaled", [
        # Seen live at 1.75: a cell 49.0 to 96.25 px wide, 362.25 to 388.5 px high.
        ((28, 207, 27, 15), 1.75, (3199, 1778), (49, 362, 48, 27)),
        ((20, 20, 80, 40), 0.5, (200, 100), (10, 10, 40, 20)),        # whole pixels: kept
        ((0, 0, 399, 200), 200 / 399, (200, 100), (0, 0, 200, 100)),  # 100.25 px high: the edge
        ((0, 0, 420, 200), 0.5, (200, 100), (0, 0, 210, 100)),        # outside: for the refusal
    ])
    def test_a_scaled_box_is_rounded_outward(self, box, k, size, scaled):
        from mcp_sap_gui.capture import scale_outward

        assert scale_outward(box, k, size) == scaled

    def test_windows_is_asked_for_the_show_state(self):
        from mcp_sap_gui import capture

        placement = {1: (0, 3, (-1, -1), (-1, -1), (0, 0, 800, 600)),  # maximized
                     2: (0, 1, (-1, -1), (-1, -1), (0, 0, 800, 600))}  # normal
        fake = SimpleNamespace(GetWindowPlacement=placement.__getitem__)
        modules = {"win32gui": fake, "win32con": SimpleNamespace(SW_SHOWMAXIMIZED=3)}

        with patch.dict(sys.modules, modules):
            assert capture.is_maximized(1)
            assert not capture.is_maximized(2)
            assert not capture.is_maximized(99)  # no such window: not maximized


class TestScreenshotToolArguments:
    async def test_evidence_options_need_save_path(self):
        import mcp_sap_gui.server as srv

        with patch.object(srv, "_ctrl", return_value=MagicMock()):
            with pytest.raises(ValueError, match="need save_path"):
                await srv.sap_screenshot(MagicMock(), caption="x")

    async def test_isolate_needs_save_path(self):
        import mcp_sap_gui.server as srv

        with patch.object(srv, "_ctrl", return_value=MagicMock()):
            with pytest.raises(ValueError, match="need save_path"):
                await srv.sap_screenshot(MagicMock(), isolate=True)


CELL = {"id": "wnd[0]/usr/shell", "row": 3, "column": "COL_B"}


class TestAlvCellTargets:
    @pytest.fixture(autouse=True)
    def _physical_cells(self):
        # Cells in the units of Screen*; the unit itself is tested in TestCellUnit.
        with patch.object(SAPGUIController, "_cell_unit", return_value=1.0):
            yield

    @staticmethod
    def _grid(first=2, count=10):
        grid = MagicMock(ScreenLeft=100, ScreenTop=200, Width=500, Height=300,
                         FirstVisibleRow=first, VisibleRowCount=count)
        grid.GetCellLeft.return_value = 50
        grid.GetCellTop.return_value = 40
        grid.ColumnCount = 2
        grid.ColumnOrder.side_effect = lambda i: ["COL_A", "COL_B"][i]
        grid.GetCellWidth.return_value = 80
        grid.GetCellHeight.return_value = 20
        return grid

    def _controller(self, grid):
        controller = SAPGUIController()
        controller._session = MagicMock(Busy=False)
        controller._session.findById.return_value = grid
        return controller

    def test_cell_box_is_grid_origin_plus_cell_offsets(self):
        grid = self._grid()
        box = self._controller(grid)._screen_box(CELL, "wnd[0]")

        assert box == (100 + 50, 200 + 40, 80, 20)
        grid.GetCellLeft.assert_called_with(3, "COL_B")

    def test_cell_offsets_are_scaled_by_the_cell_unit(self):
        # Seen live: a process that started at 96 DPI reports GetCell* in 96-DPI pixels
        # while the window is at 150% (ScreenLeft/Top are physical).
        grid = self._grid()
        with patch.object(SAPGUIController, "_cell_unit", return_value=1.5):
            box = self._controller(grid)._screen_box(CELL, "wnd[0]")

        assert box == (100 + 75, 200 + 60, 120, 30)

    def test_the_scaled_bottom_edge_never_overshoots_the_grid(self):
        # 323 + 17 = 340 -> 510 at 1.5; rounding top (485) and height (26) apart gave 511.
        grid = self._grid()
        grid.GetCellTop.return_value = 323
        grid.GetCellHeight.return_value = 17
        grid.Height = 510
        with patch.object(SAPGUIController, "_cell_unit", return_value=1.5):
            box = self._controller(grid)._screen_box(CELL, "wnd[0]", clip=False)

        assert box[1] + box[3] == 200 + 510

    def test_column_runs_from_first_visible_row_to_the_grid_bottom(self):
        grid = self._grid()
        target = {"id": "wnd[0]/usr/shell", "column": "COL_B"}
        box = self._controller(grid)._screen_box(target, "wnd[0]")

        assert box == (150, 240, 80, 200 + 300 - 240)
        grid.GetCellLeft.assert_called_with(2, "COL_B")

    @pytest.mark.parametrize("row", [1, 12, 99])
    def test_a_row_scrolled_out_of_view_is_an_error(self, row):
        controller = self._controller(self._grid(first=2, count=10))

        with pytest.raises(SAPGUIError, match="not visible"):
            controller._screen_box(dict(CELL, row=row), "wnd[0]")

    def test_a_column_beyond_the_grids_right_edge_is_an_error(self):
        grid = self._grid()
        grid.GetCellLeft.return_value = 900  # right of the 500 px grid

        with pytest.raises(SAPGUIError, match="outside the visible"):
            self._controller(grid)._screen_box(CELL, "wnd[0]")

    def test_a_cell_cut_by_the_grid_edge_is_clipped_for_a_mask_and_refused_for_a_highlight(self):
        grid = self._grid()
        grid.GetCellLeft.return_value = 450  # 80 px wide in a grid that ends 50 px on

        clipped = self._controller(grid)._screen_box(CELL, "wnd[0]", clip=True)
        assert clipped == (100 + 450, 200 + 40, 50, 20)
        with pytest.raises(SAPGUIError, match="only partly visible"):
            self._controller(grid)._screen_box(CELL, "wnd[0]", clip=False)

    def test_a_cell_reported_with_no_size_is_an_error(self):
        grid = self._grid()
        grid.GetCellWidth.return_value = 0  # what SAP reports for a scrolled-off row
        grid.GetCellHeight.return_value = 0

        with pytest.raises(SAPGUIError, match="not visible"):
            self._controller(grid)._screen_box(CELL, "wnd[0]")

    def test_an_unknown_column_is_an_error_not_a_bogus_box(self):
        grid = self._grid()

        with pytest.raises(SAPGUIError, match="Unknown column 'NOPE'.*COL_A, COL_B"):
            self._controller(grid)._screen_box(dict(CELL, column="NOPE"), "wnd[0]")
        grid.GetCellLeft.assert_not_called()

    @pytest.mark.parametrize("target", [
        {"id": "wnd[0]/usr/shell", "row": 3},                       # no column
        {"id": "wnd[0]/usr/shell", "column": "COL_A", "rows": 3},   # unknown key
        {"id": "wnd[0]/usr/shell", "column": "COL_A", "row": "3"},  # row not a number
    ])
    def test_a_malformed_cell_target_is_an_error(self, target):
        with pytest.raises(SAPGUIError):
            self._controller(self._grid())._screen_box(target, "wnd[0]")

    def test_the_tool_passes_cell_targets_on_as_plain_dicts(self):
        from mcp_sap_gui.server import CellTarget, _targets

        cell = CellTarget(id="wnd[0]/usr/shell", column="COL_A")

        assert _targets(["wnd[0]/usr/txtF", cell]) == [
            "wnd[0]/usr/txtF", {"id": "wnd[0]/usr/shell", "column": "COL_A"},
        ]


class TestIsolateAndContext:
    def test_isolate_keeps_only_the_highlighted_area(self, tmp_path):
        png = tmp_path / "a.png"
        Image.new("RGB", (200, 100), (255, 255, 255)).save(png, "PNG")

        annotate_png(str(png), highlight=[(50, 30, 40, 20)], isolate=True)

        with Image.open(png) as img:
            assert img.getpixel((70, 40)) == (255, 255, 255)  # inside the highlight
            assert img.getpixel((10, 10)) == (0, 0, 0)
            assert img.getpixel((190, 90)) == (0, 0, 0)

    def test_context_line_is_off_unless_requested(self, tmp_path):
        controller, _ = _controller(tmp_path)
        controller._evidence_meta = MagicMock(return_value="DEV/100 user TESTUSER")

        controller.save_screenshot(str(tmp_path / "a.png"), inline=False, caption="Note")
        controller._evidence_meta.assert_not_called()

        controller.save_screenshot(
            str(tmp_path / "b.png"), inline=False, caption="Note", include_context=True,
        )
        controller._evidence_meta.assert_called_once()

    def test_include_context_without_a_caption_saves_nothing(self, tmp_path):
        controller, _ = _controller(tmp_path)
        target = tmp_path / "a.png"

        result = controller.save_screenshot(str(target), include_context=True)

        assert "give a caption" in result["error"]
        assert not target.exists()

    def test_isolate_without_a_highlight_saves_nothing(self, tmp_path):
        controller, _ = _controller(tmp_path)
        target = tmp_path / "a.png"

        result = controller.save_screenshot(str(target), isolate=True)

        assert "isolate needs" in result["error"]
        assert not target.exists()


class TestCellScale:
    """cell_scale = effective DPI of the window's monitor / system DPI of the SAP process."""

    @staticmethod
    def _scale(monitor_dpi, process_dpi, proc=1, monitor=5, hresult=0):
        from mcp_sap_gui import capture

        def get_dpi_for_monitor(_monitor, _kind, by_x, _by_y):
            by_x._obj.value = monitor_dpi
            return hresult

        windll = MagicMock()
        windll.user32.MonitorFromWindow.return_value = monitor
        windll.shcore.GetDpiForMonitor.side_effect = get_dpi_for_monitor
        windll.user32.GetSystemDpiForProcess.return_value = process_dpi
        windll.kernel32.OpenProcess.return_value = proc
        with patch.object(capture.ctypes, "windll", windll, create=True), patch.object(
            capture, "_physical_pixels", MagicMock()
        ):
            return capture.cell_scale(7)

    def test_process_started_at_96_on_a_150_percent_monitor(self):
        assert self._scale(144, 96) == 1.5

    def test_process_signed_in_at_the_monitor_dpi_needs_no_scaling(self):
        assert self._scale(168, 168) == 1.0

    def test_a_window_on_a_monitor_below_the_process_dpi_is_scaled_down(self):
        # Seen live in Business Client and in plain SAP GUI (system-aware: its
        # GetDpiForWindow says 144 on the 96-DPI monitor, the monitor says 96).
        assert self._scale(96, 144) == pytest.approx(2 / 3)

    @pytest.mark.parametrize(
        "monitor_dpi,process_dpi,proc,monitor,hresult",
        [(0, 96, 1, 5, 0), (144, 0, 1, 5, 0), (144, 96, 0, 5, 0), (144, 96, 1, 0, 0),
         (144, 96, 1, 5, 1)],
    )
    def test_unknown_values_give_none(self, monitor_dpi, process_dpi, proc, monitor, hresult):
        assert self._scale(monitor_dpi, process_dpi, proc, monitor, hresult) is None


class TestCellUnit:
    """Cells are scaled to the units of Screen*/Width, which are physical only sometimes."""

    @staticmethod
    def _unit(cell, physical, reported):
        window = MagicMock(Width=reported)
        with patch("mcp_sap_gui.discovery.cell_scale", return_value=cell), patch(
            "mcp_sap_gui.discovery.physical_width", return_value=physical
        ):
            return SAPGUIController._cell_unit(7, window)

    def test_screen_units_are_physical(self):
        assert self._unit(1.5, 1920, 1920) == 1.5

    def test_screen_units_larger_than_the_pixels_are_followed(self):
        assert self._unit(2 / 3, 1920, 2880) == pytest.approx(1.0)

    @pytest.mark.parametrize("cell,physical", [(None, 1920), (1.5, None)])
    def test_an_unknown_value_is_refused(self, cell, physical):
        with pytest.raises(SAPGUIError, match="Cannot tell the unit"):
            self._unit(cell, physical, 1920)

    def test_without_a_handle_it_is_refused(self):
        with pytest.raises(SAPGUIError, match="Cannot tell the unit"):
            SAPGUIController._cell_unit(None, MagicMock(Width=1920))
