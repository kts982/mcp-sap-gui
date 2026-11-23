"""Window capture by handle for screenshots.

``GuiFrameWindow.HardCopy`` copies the window's rectangle *from the screen*, so
whatever covers SAP (another application, another Business Client tab) ends up
in the image, and a minimized window gives a black one. ``PrintWindow`` renders
the window itself. Every session of a Business Client tab has its own HWND and
only the visible tab renders, so the caller must check visibility first.
"""

import ctypes
import logging
from contextlib import contextmanager
from typing import Tuple

from .models import SAPGUIError

logger = logging.getLogger(__name__)

_PW_RENDERFULLCONTENT = 2
_GA_ROOT = 2
_GW_OWNER = 4
_PER_MONITOR_AWARE_V2 = -4  # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2


_EMBEDDED_CLASS = "SAP_FRONTEND_EMBEDDED"


def is_embedded(hwnd: int) -> bool:
    """True when the SAP window sits inside another host (Business Client).

    There ``HardCopy`` is a screen copy, so only ``PrintWindow`` is right. On
    plain SAP GUI ``HardCopy`` draws the window itself. Looks at the handle and
    its parents for the class SAP gives the embedded window.
    """
    try:
        import win32gui
    except ImportError:
        return False
    while hwnd:
        if win32gui.GetClassName(hwnd) == _EMBEDDED_CLASS:
            return True
        hwnd = win32gui.GetParent(hwnd)
    return False


@contextmanager
def _physical_pixels():
    """Make this thread DPI-aware so window sizes and PrintWindow use real pixels.

    The server process is not DPI-aware; at 150% display scaling Windows would
    hand it a 1280x673 view of a 1920x1010 window and the capture is cropped.
    The per-thread context is restored afterwards. Before Windows 10 1607 the
    call does not exist and nothing changes.
    """
    user32 = ctypes.windll.user32
    try:
        user32.SetThreadDpiAwarenessContext.restype = ctypes.c_void_p
        user32.SetThreadDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        old = user32.SetThreadDpiAwarenessContext(ctypes.c_void_p(_PER_MONITOR_AWARE_V2))
    except AttributeError:
        old = None
    try:
        yield
    finally:
        if old:
            user32.SetThreadDpiAwarenessContext(ctypes.c_void_p(old))


def print_window_png(hwnd: int, path: str) -> Tuple[int, int]:
    """Render the window with this handle into a PNG; return its screen (left, top).

    Runs DPI-aware, so the returned origin is in the same physical pixels SAP
    reports for ``ScreenLeft``/``ScreenTop``.

    Raises SAPGUIError when the window is minimized, not visible (a hidden
    Business Client tab), cannot be rendered, or Pillow is missing. There is no
    HardCopy fallback: in Business Client it would copy the screen.
    """
    with _physical_pixels():
        return _print_window_png(hwnd, path)


def _print_window_png(hwnd: int, path: str) -> Tuple[int, int]:
    try:
        import win32gui
        import win32ui
        from PIL import Image
    except ImportError as exc:
        raise SAPGUIError(
            "Capturing a Business Client window needs Pillow and pywin32: "
            "install mcp-sap-gui[screenshots]"
        ) from exc

    # A Business Client tab is a child window: only the NWBC main window is
    # minimized. A popup is a top-level dialog which that main window owns.
    root = ctypes.windll.user32.GetAncestor(hwnd, _GA_ROOT)
    owner = win32gui.GetWindow(root or hwnd, _GW_OWNER)
    if any(handle and win32gui.IsIconic(handle) for handle in (hwnd, root, owner)):
        raise SAPGUIError("The SAP window is minimized; restore it and retry.")
    if not win32gui.IsWindowVisible(hwnd):
        raise SAPGUIError(
            "This session is not the visible one (another Business Client tab "
            "or window is shown). Switch to it in SAP and retry; the capture "
            "would show the wrong session."
        )

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width, height = right - left, bottom - top
    if width <= 0 or height <= 0:
        raise SAPGUIError("The SAP window has no size; cannot capture it.")

    hdc = win32gui.GetWindowDC(hwnd)
    src = win32ui.CreateDCFromHandle(hdc)
    mem = src.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    try:
        bmp.CreateCompatibleBitmap(src, width, height)
        mem.SelectObject(bmp)
        if not ctypes.windll.user32.PrintWindow(
            hwnd, mem.GetSafeHdc(), _PW_RENDERFULLCONTENT
        ):
            raise SAPGUIError("Windows could not render the SAP window.")
        image = Image.frombuffer(
            "RGB", (width, height), bmp.GetBitmapBits(True), "raw", "BGRX", 0, 1
        )
    finally:
        win32gui.DeleteObject(bmp.GetHandle())
        mem.DeleteDC()
        src.DeleteDC()
        win32gui.ReleaseDC(hwnd, hdc)

    if image.getextrema() == ((0, 0), (0, 0), (0, 0)):
        raise SAPGUIError("The capture is completely black; not saving it.")
    image.save(path, "PNG")
    return left, top
