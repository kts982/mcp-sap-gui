"""Window capture by handle and evidence annotation for screenshots.

``GuiFrameWindow.HardCopy`` copies the window's rectangle *from the screen*, so
whatever covers SAP (another application, another Business Client tab) ends up
in the image, and a minimized window gives a black one. ``PrintWindow`` renders
the window itself. Every session of a Business Client tab has its own HWND and
only the visible tab renders, so the caller must check visibility first.
"""

import ctypes
import logging
from contextlib import contextmanager
from datetime import datetime
from typing import Iterable, List, Tuple

from .models import SAPGUIError

logger = logging.getLogger(__name__)

_PW_RENDERFULLCONTENT = 2
_GA_ROOT = 2
_GW_OWNER = 4
_PER_MONITOR_AWARE_V2 = -4  # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2

Box = Tuple[int, int, int, int]  # left, top, width, height in image pixels


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


def is_maximized(hwnd: int) -> bool:
    """True when Windows shows the window with this handle maximized.

    SAP reports a maximized window with its invisible frame, and ``HardCopy``
    leaves the frame out of the image; the caller needs to know which case it
    has. Anything that cannot be asked counts as not maximized.
    """
    try:
        import win32con
        import win32gui

        return win32gui.GetWindowPlacement(hwnd)[1] == win32con.SW_SHOWMAXIMIZED
    except Exception:
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


def _require_shown(win32gui, hwnd: int) -> None:
    """SAPGUIError unless this window is the one on screen and not minimized."""
    # A Business Client tab is a child window: only the NWBC main window is
    # minimized. A popup is a top-level dialog which that main window owns.
    root = win32gui.GetAncestor(hwnd, _GA_ROOT)
    owner = win32gui.GetWindow(root or hwnd, _GW_OWNER)
    if any(handle and win32gui.IsIconic(handle) for handle in (hwnd, root, owner)):
        raise SAPGUIError("The SAP window is minimized; restore it and retry.")
    if not win32gui.IsWindowVisible(hwnd):
        raise SAPGUIError(
            "This session is not the visible one (another Business Client tab "
            "or window is shown). Switch to it in SAP and retry; the capture "
            "would show the wrong session."
        )


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
    """Render the window with this handle into a PNG; return its screen (left, top).

    Raises SAPGUIError when the window is minimized, not visible (a hidden
    Business Client tab), cannot be rendered, or Pillow is missing.
    """
    try:
        import win32gui
        import win32ui
        from PIL import Image
    except ImportError as exc:
        raise SAPGUIError(
            "Capturing a Business Client window needs Pillow and pywin32: "
            "install mcp-sap-gui[screenshots]"
        ) from exc

    _require_shown(win32gui, hwnd)

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
        # A tab switch between the first check and the render makes PrintWindow
        # return the other tab's screen, so ask again before keeping the image.
        _require_shown(win32gui, hwnd)
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


def _pillow():
    """Import Pillow lazily; without it refuse rather than save an unmasked file."""
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError as exc:
        raise SAPGUIError(
            "highlight, mask and caption need Pillow: install "
            "mcp-sap-gui[screenshots]"
        ) from exc
    return Image, ImageDraw, ImageFont


def _wrap(draw, text: str, font, max_width: int) -> List[str]:
    lines: List[str] = []
    for paragraph in text.splitlines() or [""]:
        line = ""
        for word in paragraph.split():
            candidate = f"{line} {word}".strip()
            if line and draw.textlength(candidate, font=font) > max_width:
                lines.append(line)
                line = word
            else:
                line = candidate
        lines.append(line)
    return lines


def _require_inside(boxes: Iterable[Box], size: Tuple[int, int]) -> None:
    """Refuse a target that is not wholly inside the captured window.

    A target partly outside would be cut off, and for a mask that leaks the
    data it was meant to hide.
    """
    width, height = size
    for left, top, box_width, box_height in boxes:
        if left < 0 or top < 0 or left + box_width > width or top + box_height > height:
            raise SAPGUIError("A target lies outside the captured window; nothing saved")


def cut_at_frame(box: Box, size: Tuple[int, int], frame: int) -> Box:
    """*box* cut at the image's edge where it reaches into the window's frame.

    The image of a maximized window lacks the frame, *frame* pixels on every
    side, and an element beside it can reach into it: the status bar does, by
    3 px at the bottom. What is cut off is in no capture. A box that reaches
    further than the frame, or of which nothing would be left, comes back
    unchanged and is then refused as outside the captured window.
    """
    left, top, width, height = box
    right, bottom = left + width, top + height
    if left < -frame or top < -frame or right > size[0] + frame or bottom > size[1] + frame:
        return box
    left, top = max(left, 0), max(top, 0)
    right, bottom = min(right, size[0]), min(bottom, size[1])
    if right <= left or bottom <= top:
        return box
    return left, top, right - left, bottom - top


def _isolated(image, highlight: Iterable[Box]):
    """A black copy of *image* where only the highlight boxes keep their pixels."""
    pil_image, _, _ = _pillow()
    kept = pil_image.new("RGB", image.size, (0, 0, 0))
    for left, top, width, height in highlight:
        area = (max(left, 0), max(top, 0),
                min(left + width, image.width), min(top + height, image.height))
        kept.paste(image.crop(area), area[:2])
    return kept


def _with_caption(image, caption: str, meta: str):
    """*image* with a dark strip below: the caption, then the grey context line."""
    pil_image, pil_draw, pil_font = _pillow()
    font = pil_font.load_default(size=max(14, image.width // 110))
    pad = 8
    lines = _wrap(pil_draw.Draw(image), caption, font, image.width - 2 * pad) if caption else []
    if meta:
        lines.append(meta)
    line_h = int(font.size * 1.4)
    strip = pil_image.new("RGB", (image.width, pad * 2 + line_h * len(lines)), (32, 32, 32))
    sdraw = pil_draw.Draw(strip)
    for i, line in enumerate(lines):
        is_meta = bool(meta) and i == len(lines) - 1
        colour = (190, 190, 190) if is_meta else (255, 255, 255)
        sdraw.text((pad, pad + i * line_h), line, font=font, fill=colour)
    combined = pil_image.new("RGB", (image.width, image.height + strip.height))
    combined.paste(image, (0, 0))
    combined.paste(strip, (0, image.height))
    return combined


def annotate_png(
    path: str,
    *,
    mask: Iterable[Box] = (),
    highlight: Iterable[Box] = (),
    caption: str = "",
    meta: str = "",
    isolate: bool = False,
) -> None:
    """Black out *mask* boxes, frame *highlight* boxes, append a caption strip.

    With *isolate*, everything outside the highlight boxes is blacked out
    first, so only the evidence stays visible. Masks go on before the frames
    so a highlight can never expose masked pixels. Needs
    Pillow; raises SAPGUIError without it rather than saving an unmasked file.
    """
    pil_image, pil_draw, _ = _pillow()
    with pil_image.open(path) as opened:
        image = opened.convert("RGB")
    _require_inside((*mask, *highlight), image.size)
    if isolate:
        image = _isolated(image, highlight)
    draw = pil_draw.Draw(image)
    for left, top, width, height in mask:
        draw.rectangle([left, top, left + width, top + height], fill=(0, 0, 0))
    for left, top, width, height in highlight:
        draw.rectangle(
            [left - 2, top - 2, left + width + 2, top + height + 2],
            outline=(255, 0, 0), width=3,
        )
    if caption or meta:
        image = _with_caption(image, caption, meta)
    image.save(path, "PNG")


def evidence_line(system: str, client: str, user: str, tcode: str) -> str:
    return f"{system}/{client}  user {user}  tcode {tcode}  {datetime.now():%Y-%m-%d %H:%M:%S}"
